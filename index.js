/**
 * JTe Pauta -> MySQL + XLSX + Email SMTP
 *
 * deps:
 *   npm i playwright mysql2 nodemailer exceljs
 *
 * env (exemplo):
 *   DB_ENABLED=true
 *   DB_HOST=127.0.0.1
 *   DB_PORT=3306
 *   DB_USER=root
 *   DB_PASS=senha
 *   DB_NAME=jte
 *
 *   SMTP_HOST=smtp.office365.com
 *   SMTP_PORT=587
 *   SMTP_SECURE=false
 *   SMTP_USER=seu-email@dominio.com
 *   SMTP_PASS=sua-senha-ou-app-password
 *   MAIL_FROM="Robô JTe <seu-email@dominio.com>"
 *   MAIL_TO=destinatario@dominio.com;dest2@dominio.com
 */
require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const mysql = require('mysql2/promise');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');

/* =========================
   CSV HELPERS
========================= */

function csvEscape(value) {
  if (value === null || value === undefined) return '';
  const s = String(value);
  if (/[;"\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeCsv(filePath, headers, rows) {
  const bom = '\uFEFF';
  const lines = [];
  lines.push(headers.map(csvEscape).join(';'));
  for (const row of rows) {
    lines.push(headers.map((h) => csvEscape(row[h])).join(';'));
  }
  fs.writeFileSync(filePath, bom + lines.join('\n'), 'utf8');
}

/* =========================
   DATE HELPERS
========================= */

function parseBRDate(br) {
  const [dd, mm, yyyy] = br.split('/').map(Number);
  return new Date(yyyy, mm - 1, dd);
}

function isWeekdayBR(dataBR) {
  const d = parseBRDate(dataBR);
  const day = d.getDay();
  return day !== 0 && day !== 6;
}

function getTodayBR() {
  const hoje = new Date();
  const dd = String(hoje.getDate()).padStart(2, '0');
  const mm = String(hoje.getMonth() + 1).padStart(2, '0');
  const yyyy = hoje.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function gerarDatasProximosDoisMeses() {
  const hoje = new Date();
  const datas = [];

  let current = new Date(hoje);
  current.setDate(current.getDate() + 7);

  const fim = new Date(hoje);
  fim.setMonth(fim.getMonth() + 2);
  fim.setDate(fim.getDate() + 10);

  while (current <= fim) {
    const dd = String(current.getDate()).padStart(2, '0');
    const mm = String(current.getMonth() + 1).padStart(2, '0');
    const yyyy = current.getFullYear();
    const dataBR = `${dd}/${mm}/${yyyy}`;

    if (isWeekdayBR(dataBR)) datas.push(dataBR);
    current.setDate(current.getDate() + 1);
  }

  return datas;
}

/* =========================
   DB (MySQL)
========================= */

function getEnv(name, fallback = undefined) {
  const v = process.env[name];
  return (v === undefined || v === null || v === '') ? fallback : v;
}

function isTrue(v) {
  return String(v ?? '').trim().toLowerCase() === 'true' || String(v ?? '').trim() === '1';
}

async function openDb() {
  const host = getEnv('DB_HOST', '127.0.0.1'); // ✅ melhor que localhost no Windows
  const port = Number(getEnv('DB_PORT', '3306'));
  const user = getEnv('DB_USER', 'root');
  const password = getEnv('DB_PASS', '');
  const database = getEnv('DB_NAME', 'jte');

  const pool = mysql.createPool({
    host,
    port,
    user,
    password,
    database,
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    multipleStatements: false,
    connectTimeout: 8000, // ✅ pra não travar demais quando não tem DB
  });

  return pool;
}

async function ensureSchema(pool) {
  const sql = `
CREATE TABLE IF NOT EXISTS pauta_processos (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  geradoEm DATETIME(3) NOT NULL,
  vara VARCHAR(255) NOT NULL,
  dataBR VARCHAR(10) NOT NULL,
  dataISO DATE NOT NULL,
  numeroProcesso VARCHAR(64) NOT NULL,
  sessao VARCHAR(255) NULL,
  juiz VARCHAR(255) NULL,
  reclamante VARCHAR(255) NULL,
  reclamada VARCHAR(255) NULL,
  createdAt TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_pauta (vara, dataISO, numeroProcesso),
  KEY ix_data (dataISO),
  KEY ix_vara (vara)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
`;
  await pool.query(sql);
}

async function initDbIfEnabled() {
  const enabled = isTrue(getEnv('DB_ENABLED', 'false'));

  if (!enabled) {
    console.log('ℹ️ DB_ENABLED=false -> rodando sem MySQL (só CSV/XLSX/email).');
    return null;
  }

  let pool = null;

  try {
    pool = await openDb();

    // ✅ força conexão (pega ECONNREFUSED aqui e não no meio do código)
    await pool.query('SELECT 1');

    await ensureSchema(pool);

    console.log('✅ MySQL conectado e schema OK');
    return pool;
  } catch (err) {
    console.warn(
      `⚠️ MySQL indisponível (seguindo sem DB): ${err.code || ''} ${err.message || err}`
    );
    try { if (pool) await pool.end(); } catch { }
    return null;
  }
}

function brToIsoDateString(br) {
  const d = parseBRDate(br);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

async function insertRowsMySql(pool, rows, chunkSize = 800) {
  if (!pool) return { insertedOrUpdated: 0 };
  if (!rows || rows.length === 0) return { insertedOrUpdated: 0 };

  let total = 0;

  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);

    const values = [];
    const placeholders = chunk.map((r) => {
      const dataISO = brToIsoDateString(r.data);
      values.push(
        new Date(r.geradoEm),
        r.vara,
        r.data,
        dataISO,
        r.numeroProcesso,
        r.sessao ?? null,
        r.juiz ?? null,
        r.reclamante ?? null,
        r.reclamada ?? null
      );
      return '(?,?,?,?,?,?,?,?,?)';
    });

    const sql = `
INSERT INTO pauta_processos
(geradoEm, vara, dataBR, dataISO, numeroProcesso, sessao, juiz, reclamante, reclamada)
VALUES ${placeholders.join(',')}
ON DUPLICATE KEY UPDATE
  geradoEm = VALUES(geradoEm),
  sessao = VALUES(sessao),
  juiz = VALUES(juiz),
  reclamante = VALUES(reclamante),
  reclamada = VALUES(reclamada)
`;
    const [res] = await pool.query(sql, values);
    total += Number(res.affectedRows || 0);
  }

  return { insertedOrUpdated: total };
}

/* =========================
   XLSX
========================= */

async function writeXlsx(filePath, headers, rows) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'JTe Bot';
  wb.created = new Date();

  const ws = wb.addWorksheet('Pauta');

  ws.columns = headers.map((h) => ({
    header: h,
    key: h,
    width: Math.max(12, Math.min(40, h.length + 6)),
  }));

  for (const r of rows) ws.addRow(r);

  ws.getRow(1).font = { bold: true };
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: headers.length },
  };

  for (let c = 1; c <= headers.length; c++) {
    let maxLen = String(headers[c - 1]).length;
    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const v = row.getCell(c).value;
      const s = (v === null || v === undefined) ? '' : String(v);
      if (s.length > maxLen) maxLen = s.length;
    });
    ws.getColumn(c).width = Math.max(12, Math.min(60, maxLen + 2));
  }

  await wb.xlsx.writeFile(filePath);
}

/* =========================
   EMAIL (SMTP)
========================= */

function parseMailTo(value) {
  if (!value) return [];
  return String(value)
    .split(/[;,]/g)
    .map((s) => s.trim())
    .filter(Boolean);
}

async function sendEmailWithAttachment({ subject, text, attachmentPath }) {
  const host = getEnv('SMTP_HOST');
  const port = Number(getEnv('SMTP_PORT', '587'));
  const secure = String(getEnv('SMTP_SECURE', 'false')).toLowerCase() === 'true';
  const user = getEnv('SMTP_USER');
  const pass = getEnv('SMTP_PASS');
  const from = getEnv('MAIL_FROM', user);
  const to = parseMailTo(getEnv('MAIL_TO', ''));

  if (!host || !user || !pass || !from || !to.length) {
    throw new Error(
      'Config SMTP incompleta. Verifique SMTP_HOST/SMTP_USER/SMTP_PASS/MAIL_FROM/MAIL_TO no env.'
    );
  }

  const transporter = nodemailer.createTransport({
    host,
    port,
    secure,
    auth: { user, pass },
  });

  const info = await transporter.sendMail({
    from,
    to,
    subject,
    text,
    attachments: [
      {
        filename: path.basename(attachmentPath),
        path: attachmentPath,
      },
    ],
  });

  return info;
}

/* =========================
   OVERLAY & RETRY
========================= */

async function fecharOverlays(page) {
  try {
    const backdrop = page.locator('.cdk-overlay-backdrop');
    await page.keyboard.press('Escape').catch(() => { });
    await page.waitForTimeout(200);

    if (await backdrop.isVisible({ timeout: 400 }).catch(() => false)) {
      await backdrop.click({ position: { x: 5, y: 5 }, force: true }).catch(() => { });
    }
    await backdrop.waitFor({ state: 'detached', timeout: 1500 }).catch(() => { });
  } catch { }
}

async function retryOperation(page, operation, maxRetries = 5, delayMs = 1200) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await operation();
    } catch (err) {
      console.warn(`⚠️ Tentativa ${attempt}/${maxRetries} falhou: ${err.message}`);
      if (attempt === maxRetries) throw err;
      await fecharOverlays(page);
      await page.waitForTimeout(delayMs);
    }
  }
}

/* =========================
   NAVEGAÇÃO JTe
========================= */

async function abrirJTeSelecionarTRT2(page) {
  console.log('➡️ Acessando JTe...');
  await page.goto('https://jte.csjt.jus.br/start', { waitUntil: 'networkidle', timeout: 60000 });

  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(2500);

  const locator = page.getByText('TRT2 - São Paulo', { exact: true });
  await retryOperation(page, async () => {
    await locator.waitFor({ state: 'visible', timeout: 20000 });
    await locator.click({ force: true });
  });

  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(800);
  console.log('✅ TRT2 selecionado');
}

async function abrirModuloPauta(page) {
  console.log('➡️ Abrindo módulo Pauta...');
  const card = page.locator('ion-card-content.card-content-modulo:has-text("Pauta")').first();
  await retryOperation(page, async () => {
    await card.waitFor({ state: 'visible', timeout: 20000 });
    await card.click({ force: true });
  });

  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(800);
  console.log('✅ Módulo Pauta aberto');
}

/* =========================
   LISTAR VARAS
========================= */

async function listarVaras(page) {
  console.log('➡️ Listando varas...');
  const botaoUnidade = page.getByTestId('pautaButtonSelecaoUnidade');
  await botaoUnidade.waitFor({ state: 'visible', timeout: 20000 });
  await botaoUnidade.click({ force: true });

  await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', { timeout: 20000 });

  const selectTipo = page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select');
  await selectTipo.click();
  await page.locator('mat-option:has-text("Audiências 1º grau")').first().click();
  await page.waitForTimeout(200);

  const selectMunicipio = page.locator('mat-form-field[data-testid="municipio"] mat-select');
  await selectMunicipio.click();
  await page.locator('.mat-mdc-select-panel mat-option:has-text("São Paulo - Zonas Central, Norte e Oeste")').click();
  await page.waitForTimeout(200);

  await page.waitForSelector('mat-form-field[data-testid="orgao"] mat-select[aria-disabled="false"]', { timeout: 20000 });

  const selectOrgao = page.locator('mat-form-field[data-testid="orgao"] mat-select');
  await selectOrgao.click();

  const opcoes = page.locator('.mat-mdc-select-panel mat-option');
  const total = await opcoes.count();

  const varas = [];
  for (let i = 0; i < total; i++) {
    const label = await opcoes.nth(i).locator('.mdc-list-item__primary-text').textContent();
    if (label) varas.push(label.trim());
  }

  await page.keyboard.press('Escape').catch(() => { });
  await page.getByTestId('ButtonCancelar').click().catch(() => { });

  console.log(`✅ ${varas.length} varas encontradas`);
  return varas;
}

/* =========================
   SELECIONAR UNIDADE
========================= */
function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function openMatSelect(page, selectLocator) {
  await selectLocator.scrollIntoViewIfNeeded().catch(() => { });
  await selectLocator.click({ force: true });

  const panel = page.locator('.mat-mdc-select-panel');
  await panel.waitFor({ state: 'visible', timeout: 20000 });
  return panel;
}

async function clickMatOptionByText(page, text, opts) {
  const panel = page.locator('.mat-mdc-select-panel');
  await panel.waitFor({ state: 'visible', timeout: 20000 });

  const pattern = opts?.exact === false
    ? text
    : new RegExp(`^\\s*${escapeRegExp(text)}\\s*$`, 'i');

  const option = panel.locator('mat-option').filter({ hasText: pattern }).first();
  await option.waitFor({ state: 'visible', timeout: 20000 });
  await option.scrollIntoViewIfNeeded().catch(() => { });
  await option.click({ force: true });

  await panel.waitFor({ state: 'hidden', timeout: 20000 }).catch(() => { });
}

async function matSelectChoose(page, selectLocator, optionText, opts) {
  await openMatSelect(page, selectLocator);
  await clickMatOptionByText(page, optionText, opts);
  await page.waitForTimeout(150);
}

async function waitMatSelectEnabled(page, selector) {
  await page.waitForSelector(selector, { timeout: 20000 });
  const loc = page.locator(selector);
  await page.waitForFunction(
    (el) => el.getAttribute('aria-disabled') !== 'true',
    await loc.elementHandle(),
    { timeout: 20000 }
  );
}

async function selecionarUnidade(page, varaLabel) {
  console.log(`\n🏛️ Selecionando vara: ${varaLabel}`);

  await fecharOverlays(page);

  const botaoUnidade = page.getByTestId('pautaButtonSelecaoUnidade');
  await botaoUnidade.waitFor({ state: 'visible', timeout: 20000 });
  await botaoUnidade.click({ force: true });

  await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', { timeout: 20000 });

  const selectTipo = page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select');
  await matSelectChoose(page, selectTipo, 'Audiências 1º grau', { exact: true });

  const selectMunicipio = page.locator('mat-form-field[data-testid="municipio"] mat-select');
  await matSelectChoose(page, selectMunicipio, 'São Paulo - Zonas Central, Norte e Oeste', { exact: true });

  await waitMatSelectEnabled(page, 'mat-form-field[data-testid="orgao"] mat-select');

  const selectOrgao = page.locator('mat-form-field[data-testid="orgao"] mat-select');
  await matSelectChoose(page, selectOrgao, varaLabel, { exact: true });

  const confirmar = page.getByTestId('ButtonConfirmar');
  await confirmar.waitFor({ state: 'visible', timeout: 20000 });

  await page.waitForFunction(
    (el) => !el.hasAttribute('disabled') && el.getAttribute('aria-disabled') !== 'true',
    await confirmar.elementHandle(),
    { timeout: 20000 }
  ).catch(() => { });

  await confirmar.click({ delay: 80 }).catch(() => { });
  await page.waitForLoadState('networkidle').catch(() => { });
  await page.waitForTimeout(700);

  const todayBR = getTodayBR();
  console.log(`🔄 Tentando ajustar para hoje (sem calendário): ${todayBR}`);

  const ok = await selecionarDataComConfirmacao(page, todayBR, 1);
  if (ok) console.log(`✅ Ajustado para hoje`);
  else console.warn(`⚠️ Não conseguiu ajustar para hoje (sem calendário). Seguindo mesmo assim.`);
}

/* =========================
   SELEÇÃO DE DATA (SEM CALENDÁRIO)
========================= */

const BTN_NEXT_CSS = '#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(3) > ion-button';
const BTN_PREV_CSS = '#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(1) > ion-button';
const BTN_DATE_DISPLAY_XPATH = '//*[@id="main-content"]/ng-component[3]/ion-content/div/div/ion-grid/ion-row[2]/ion-col[2]/ion-button';

function extrairDataBR(texto) {
  const raw = String(texto ?? '').trim();
  const m = raw.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
  return m?.[0] ?? '';
}

async function clickIonButtonByCss(page, cssSel) {
  try {
    return await page.evaluate((sel) => {
      const host = document.querySelector(sel);
      if (!host) return false;
      const btn = host.shadowRoot?.querySelector('button') || host.querySelector?.('button');
      (btn || host).click();
      return true;
    }, cssSel);
  } catch {
    return false;
  }
}

async function readIonButtonTextByXpath(page, xpath) {
  try {
    return await page.evaluate((xp) => {
      const node = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      if (!node) return '';
      const btn = node.shadowRoot?.querySelector('button') || node.querySelector?.('button');
      const txt = (btn?.innerText || btn?.textContent || node.innerText || node.textContent || '').trim();
      return txt;
    }, xpath);
  } catch {
    return '';
  }
}

async function ionExistsByCss(page, cssSel) {
  try {
    return await page.evaluate((sel) => !!document.querySelector(sel), cssSel);
  } catch {
    return false;
  }
}

async function lerTextoDataExibida(page) {
  const raw = await readIonButtonTextByXpath(page, BTN_DATE_DISPLAY_XPATH);
  if (raw) return raw;

  try {
    const t = await page.getByTestId('pautaButtonData').innerText({ timeout: 800 }).catch(() => '');
    if (t && String(t).trim()) return String(t).trim();
  } catch { }

  return '';
}

async function esperarTextoMudar(page, anterior, timeoutMs = 2500) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const atual = await lerTextoDataExibida(page);
    if (atual && atual !== anterior) return atual;
    await page.waitForTimeout(80);
  }
  return await lerTextoDataExibida(page);
}

async function irAteDataPorBotoes(page, alvoBR, maxSteps = 180) {
  await fecharOverlays(page);

  const alvoDate = parseBRDate(alvoBR).getTime();

  const hasPrev = await ionExistsByCss(page, BTN_PREV_CSS);
  const hasNext = await ionExistsByCss(page, BTN_NEXT_CSS);
  if (!hasNext && !hasPrev) {
    console.warn('⚠️ Não achei botões prev/next para navegar datas.');
    return false;
  }

  let raw = await lerTextoDataExibida(page);
  if (raw && raw.includes(alvoBR)) return true;

  for (let step = 1; step <= maxSteps; step++) {
    raw = await lerTextoDataExibida(page);
    if (raw && raw.includes(alvoBR)) return true;

    const atualBR = extrairDataBR(raw);
    let direction = +1;

    if (atualBR) {
      const atualDate = parseBRDate(atualBR).getTime();
      if (hasPrev && atualDate > alvoDate) direction = -1;
      else direction = +1;
    } else {
      direction = +1;
    }

    const antes = raw || '';
    const clicked = direction === -1
      ? await clickIonButtonByCss(page, BTN_PREV_CSS)
      : await clickIonButtonByCss(page, BTN_NEXT_CSS);

    if (!clicked) {
      console.warn(`⚠️ Falhou clique no botão ${direction === -1 ? 'PREV' : 'NEXT'} (step ${step}).`);
      await page.waitForTimeout(150);
      continue;
    }

    const depois = await esperarTextoMudar(page, antes, 2500);
    if (depois && String(depois).includes(alvoBR)) return true;

    await page.waitForTimeout(120);
  }

  return false;
}

async function selecionarDataComConfirmacao(page, dataBR, maxTentativas = 3) {
  for (let t = 1; t <= maxTentativas; t++) {
    const rawAntes = await lerTextoDataExibida(page);

    const ok = await irAteDataPorBotoes(page, dataBR, 220);
    const rawDepois = await lerTextoDataExibida(page);

    console.log(`🧾 Data (tentativa ${t}/${maxTentativas}): antes="${rawAntes}" | depois="${rawDepois}" | alvo=${dataBR}`);

    if (ok) return true;

    console.warn(`⚠️ Não achou data via botões: ${dataBR}. Retentando...`);
    await fecharOverlays(page);
    await page.waitForTimeout(600);
  }

  return false;
}

/* =========================
   ESPERAR PAUTA ESTABILIZAR
========================= */

async function esperarPautaEstabilizar(page) {
  const spinners = [
    page.locator('ion-spinner').first(),
    page.locator('.mat-mdc-progress-spinner').first(),
    page.locator('.mat-mdc-progress-bar').first(),
  ];

  for (const sp of spinners) {
    if (await sp.isVisible({ timeout: 300 }).catch(() => false)) {
      await sp.waitFor({ state: 'hidden', timeout: 15000 }).catch(() => { });
    }
  }

  let last = -1;
  for (let i = 0; i < 20; i++) {
    const count = await page.locator('ion-list ion-item').count().catch(() => 0);

    if (count === last) {
      await page.waitForTimeout(600);
      const count2 = await page.locator('ion-list ion-item').count().catch(() => 0);
      if (count2 === count) return;
    }

    last = count;
    await page.waitForTimeout(250);
  }
}

/* =========================
   EXTRAÇÃO
========================= */

async function extrairProcessosDaPauta(page) {
  const count = await page.locator('ion-list ion-item').count().catch(() => 0);
  if (!count) return [];

  const processos = await page.evaluate(() => {
    const items = document.querySelectorAll('ion-list ion-item');
    return Array.from(items).map((item) => {
      const getText = (sel) => {
        const el = item.querySelector(sel);
        return el ? el.textContent.replace(/\u00a0/g, ' ').trim() : '';
      };

      const hora = getText('.sessao');
      const status = getText('.palavrasRight');
      const numeroProcesso = getText('.JT-item-texto-negrito');

      const partes = Array.from(item.querySelectorAll('.item-desc-small.item-text-wrap'))
        .map((e) => e.textContent.replace(/\u00a0/g, ' ').trim())
        .filter(Boolean);

      return {
        numeroProcesso,
        sessao: [hora, status].filter(Boolean).join(' - '),
        juiz: partes[0] || '',
        reclamante: partes[1] || '',
        reclamada: partes[2] || '',
      };
    });
  });

  return processos;
}

/* =========================
   MAIN
========================= */

async function main() {

  // ✅ DB pool (opcional)
  const pool = await initDbIfEnabled();
  const browser = await chromium.launch({ headless: false, slowMo: 100 });
  const page = await browser.newPage();

  const geradoEm = new Date().toISOString();
  const rowsCsv = [];
  const headers = ['geradoEm', 'vara', 'data', 'numeroProcesso', 'sessao', 'juiz', 'reclamante', 'reclamada'];

  try {
    await abrirJTeSelecionarTRT2(page);
    await abrirModuloPauta(page);

    const varas = await listarVaras(page);
    const varasAlvo = varas;

    const datas = gerarDatasProximosDoisMeses();
    console.log(`📅 ${datas.length} datas alvo`);

    for (const vara of varasAlvo) {
      await selecionarUnidade(page, vara);

      for (const dataBR of datas) {
        console.log(`📅 Procurando data (sem calendário): ${dataBR}`);

        const ok = await selecionarDataComConfirmacao(page, dataBR, 2);
        if (!ok) {
          console.warn(`⚠️ Pulando data (não achou no header): ${vara} | ${dataBR}`);
          continue;
        }

        await esperarPautaEstabilizar(page);

        const processos = await extrairProcessosDaPauta(page);
        console.log(`📌 ${vara} | ${dataBR} | ${processos.length} processos`);

        for (const p of processos) {
          rowsCsv.push({
            geradoEm,
            vara,
            data: dataBR,
            numeroProcesso: p.numeroProcesso,
            sessao: p.sessao,
            juiz: p.juiz,
            reclamante: p.reclamante,
            reclamada: p.reclamada,
          });
        }

        // ✅ grava em DB só se DB estiver OK
        if (pool && processos.length) {
          const rowsChunk = processos.map((p) => ({
            geradoEm,
            vara,
            data: dataBR,
            numeroProcesso: p.numeroProcesso,
            sessao: p.sessao,
            juiz: p.juiz,
            reclamante: p.reclamante,
            reclamada: p.reclamada,
          }));

          const r = await insertRowsMySql(pool, rowsChunk, 800);
          console.log(`💾 MySQL: affectedRows=${r.insertedOrUpdated} (insert/update)`);
        }
      }
    }

    const outDir = path.join(process.cwd(), 'output');
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const stamp = Date.now();
    const csvPath = path.join(outDir, `pauta_trt2_2meses_${stamp}.csv`);
    const xlsxPath = path.join(outDir, `pauta_trt2_2meses_${stamp}.xlsx`);

    writeCsv(csvPath, headers, rowsCsv);
    await writeXlsx(xlsxPath, headers, rowsCsv);

    console.log(`\n✅ Concluído! ${rowsCsv.length} linhas`);
    console.log(`📄 CSV:  ${csvPath}`);
    console.log(`📊 XLSX: ${xlsxPath}`);

    const subject = `Pauta TRT2 (2 meses) - ${new Date().toLocaleString('pt-BR')}`;
    const text =
      `Olá!\n\n` +
      `Segue em anexo o arquivo XLSX com a extração da pauta do TRT2 para os próximos ~2 meses.\n\n` +
      `Total de linhas: ${rowsCsv.length}\n` +
      `Gerado em: ${geradoEm}\n\n` +
      `Atenciosamente,\nRobô JTe`;

    const info = await sendEmailWithAttachment({
      subject,
      text,
      attachmentPath: xlsxPath,
    });

    console.log(`📧 Email enviado! messageId=${info.messageId || '(sem id)'}`);
  } catch (err) {
    console.error('❌ Erro:', err);
  } finally {
    try { if (pool) await pool.end(); } catch { }
    // await browser.close().catch(() => {});
  }
}

main();
