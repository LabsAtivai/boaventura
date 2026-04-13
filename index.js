/**
 * JTe Pauta TRT-2 — Scraper completo
 *
 * O que faz:
 *   1. Acessa o JTe público e coleta pautas das varas de SP (Zonas Central, Norte e Oeste)
 *   2. Para cada processo na pauta, abre o detalhe e verifica se o polo passivo está vazio
 *      OU se todos os polos passivos não têm advogado constituído
 *   3. Salva os processos encontrados no MySQL (upsert), gera XLSX e envia por e-mail
 *
 * Período: +1 dia até +7 dias úteis a partir de hoje
 *
 * Deps: npm i playwright mysql2 nodemailer exceljs dotenv
 *
 * Variáveis de ambiente (.env):
 *   DB_ENABLED, DB_HOST, DB_PORT, DB_USER, DB_PASS, DB_NAME
 *   SMTP_HOST, SMTP_PORT, SMTP_SECURE, SMTP_USER, SMTP_PASS
 *   MAIL_FROM, MAIL_TO
 *   DEBUG_HTML=true  (opcional: salva HTML dos primeiros processos para inspeção)
 */

require("dotenv").config();

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");
const mysql = require("mysql2/promise");
const nodemailer = require("nodemailer");
const ExcelJS = require("exceljs");

/* ─────────────────────────────────────────
   HELPERS GERAIS
───────────────────────────────────────── */

function getEnv(name, fallback = undefined) {
  const v = process.env[name];
  return v === undefined || v === null || v === "" ? fallback : v.trim();
}

function isTrue(v) {
  const s = String(v ?? "").trim().toLowerCase();
  return s === "true" || s === "1";
}

function csvEscape(value) {
  if (value === null || value === undefined) return "";
  const s = String(value);
  if (/[;"'\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeCsv(filePath, headers, rows) {
  const lines = [headers.map(csvEscape).join(";")];
  for (const row of rows)
    lines.push(headers.map((h) => csvEscape(row[h])).join(";"));
  fs.writeFileSync(filePath, "\uFEFF" + lines.join("\n"), "utf8");
  console.log(`💾 CSV salvo: ${filePath} (${rows.length} linhas)`);
}

/* ─────────────────────────────────────────
   DATE HELPERS
───────────────────────────────────────── */

function parseBRDate(br) {
  const [dd, mm, yyyy] = br.split("/").map(Number);
  return new Date(yyyy, mm - 1, dd);
}

function brToIso(br) {
  const d = parseBRDate(br);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function isWeekday(dataBR) {
  const day = parseBRDate(dataBR).getDay();
  return day !== 0 && day !== 6;
}

/** Gera dias úteis de +1 dia até +7 dias corridos */
function gerarDatasProximos7Dias() {
  const hoje = new Date();
  const datas = [];

  const inicio = new Date(hoje);
  inicio.setDate(inicio.getDate() + 1); // começa amanhã

  const fim = new Date(hoje);
  fim.setDate(fim.getDate() + 7); // 7 dias corridos à frente

  let cur = new Date(inicio);
  while (cur <= fim) {
    const dd = String(cur.getDate()).padStart(2, "0");
    const mm = String(cur.getMonth() + 1).padStart(2, "0");
    const yyyy = cur.getFullYear();
    const br = `${dd}/${mm}/${yyyy}`;
    if (isWeekday(br)) datas.push(br);
    cur.setDate(cur.getDate() + 1);
  }

  console.log(
    `📅 ${datas.length} dias úteis gerados (${datas[0]} → ${datas[datas.length - 1]})`,
  );
  return datas;
}

/* ─────────────────────────────────────────
   MYSQL
───────────────────────────────────────── */

async function inicializarBanco() {
  if (!isTrue(getEnv("DB_ENABLED", "false"))) {
    console.log("ℹ️  DB_ENABLED=false → rodando sem MySQL.");
    return null;
  }

  let pool = null;
  try {
    pool = mysql.createPool({
      host: getEnv("DB_HOST", "127.0.0.1"),
      port: Number(getEnv("DB_PORT", "3306")),
      user: getEnv("DB_USER", "root"),
      password: getEnv("DB_PASS", ""),
      database: getEnv("DB_NAME", "jte"),
      waitForConnections: true,
      connectionLimit: 10,
      connectTimeout: 8000,
    });

    await pool.query("SELECT 1");

    await pool.query(`
      CREATE TABLE IF NOT EXISTS processos_sem_polo_passivo (
        id               BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        geradoEm         DATETIME(3)     NOT NULL,
        vara             VARCHAR(255)    NOT NULL,
        dataBR           VARCHAR(10)     NOT NULL,
        dataISO          DATE            NOT NULL,
        numeroProcesso   VARCHAR(64)     NOT NULL,
        sessao           VARCHAR(255)    NULL,
        juiz             VARCHAR(255)    NULL,
        reclamante       VARCHAR(255)    NULL,
        createdAt        TIMESTAMP       NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_proc (vara, dataISO, numeroProcesso),
        KEY ix_data (dataISO),
        KEY ix_vara (vara(100))
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);

    console.log("✅ MySQL conectado e schema OK");
    return pool;
  } catch (err) {
    console.warn(
      `⚠️  MySQL indisponível (${err.code || err.message}) → seguindo sem DB.`,
    );
    try {
      if (pool) await pool.end();
    } catch {}
    return null;
  }
}

async function salvarNoBanco(pool, geradoEm, vara, dataBR, processo) {
  if (!pool) return;
  try {
    await pool.query(
      `INSERT INTO processos_sem_polo_passivo
         (geradoEm, vara, dataBR, dataISO, numeroProcesso, sessao, juiz, reclamante)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE
         geradoEm   = VALUES(geradoEm),
         sessao     = VALUES(sessao),
         juiz       = VALUES(juiz),
         reclamante = VALUES(reclamante)`,
      [
        new Date(geradoEm),
        vara,
        dataBR,
        brToIso(dataBR),
        processo.numeroProcesso,
        processo.sessao || null,
        processo.juiz || null,
        processo.reclamante || null,
      ],
    );
  } catch (err) {
    console.warn(
      `⚠️  Erro ao salvar no banco (${processo.numeroProcesso}): ${err.message}`,
    );
  }
}

/* ─────────────────────────────────────────
   XLSX
───────────────────────────────────────── */

async function gerarXLSX(filePath, rows) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "JTe Bot";
  wb.created = new Date();

  const ws = wb.addWorksheet("Sem Polo Passivo");

  const headers = [
    "vara",
    "data",
    "numeroProcesso",
    "sessao",
    "juiz",
    "reclamante",
  ];

  ws.columns = headers.map((h) => ({
    header: h,
    key: h,
    width: Math.max(14, Math.min(55, h.length + 6)),
  }));

  for (const r of rows) ws.addRow(r);

  ws.getRow(1).font = { bold: true };
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: headers.length },
  };

  for (let c = 1; c <= headers.length; c++) {
    let maxLen = String(headers[c - 1]).length;
    ws.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const v = row.getCell(c).value;
      const s = v == null ? "" : String(v);
      if (s.length > maxLen) maxLen = s.length;
    });
    ws.getColumn(c).width = Math.max(14, Math.min(60, maxLen + 2));
  }

  await wb.xlsx.writeFile(filePath);
  console.log(`📊 XLSX salvo: ${filePath} (${rows.length} linhas)`);
}

/* ─────────────────────────────────────────
   EMAIL
───────────────────────────────────────── */

function parseMailTo(value) {
  if (!value) return [];
  return String(value)
    .split(/[;,]/g)
    .map((s) => s.trim())
    .filter(Boolean);
}

async function enviarEmailComAnexo(xlsxPath, totalProcessos) {
  const host = getEnv("SMTP_HOST");
  const port = Number(getEnv("SMTP_PORT", "587"));
  const secure = isTrue(getEnv("SMTP_SECURE", "false"));
  const user = getEnv("SMTP_USER");
  const pass = getEnv("SMTP_PASS");
  const from = getEnv("MAIL_FROM", user);
  const to = parseMailTo(getEnv("MAIL_TO", ""));

  if (!host || !user || !pass || !to.length) {
    console.warn("⚠️  Config SMTP incompleta → e-mail não enviado.");
    return;
  }

  const transporter = nodemailer.createTransport({
    host,
    port,
    secure,
    auth: { user, pass },
  });

  const hoje = new Date().toLocaleDateString("pt-BR");

  await transporter.sendMail({
    from,
    to,
    subject: `[JTe TRT-2] Processos sem Polo Passivo — ${hoje}`,
    text: [
      `Relatório gerado em ${hoje}.`,
      ``,
      `Total de processos sem polo passivo (ou sem advogado constituído): ${totalProcessos}`,
      `Período: próximos 7 dias úteis`,
      `Regional: São Paulo — Zonas Central, Norte e Oeste`,
      ``,
      `Arquivo XLSX em anexo.`,
      ``,
      `— Robô JTe`,
    ].join("\n"),
    attachments: [{ filename: path.basename(xlsxPath), path: xlsxPath }],
  });

  console.log(`📧 E-mail enviado para: ${to.join(", ")}`);
}

/* ─────────────────────────────────────────
   BROWSER HELPERS
───────────────────────────────────────── */

async function fecharOverlays(page) {
  try {
    await page.keyboard.press("Escape").catch(() => {});
    await page.waitForTimeout(200);
    const backdrop = page.locator(".cdk-overlay-backdrop");
    if (await backdrop.isVisible({ timeout: 400 }).catch(() => false)) {
      await backdrop
        .click({ position: { x: 5, y: 5 }, force: true })
        .catch(() => {});
    }
    await backdrop
      .waitFor({ state: "detached", timeout: 1500 })
      .catch(() => {});
  } catch {}
}

async function retryOp(page, fn, maxRetries = 5, delayMs = 1200) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (err) {
      console.warn(`⚠️  Tentativa ${attempt}/${maxRetries}: ${err.message}`);
      if (attempt === maxRetries) throw err;
      await fecharOverlays(page);
      await page.waitForTimeout(delayMs);
    }
  }
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function matSelectChoose(page, selectLocator, optionText, exact = true) {
  await selectLocator.scrollIntoViewIfNeeded().catch(() => {});
  await selectLocator.click({ force: true });

  const panel = page.locator(".mat-mdc-select-panel");
  await panel.waitFor({ state: "visible", timeout: 20000 });

  const pattern = exact
    ? new RegExp(`^\\s*${escapeRegExp(optionText)}\\s*$`, "i")
    : optionText;

  const option = panel
    .locator("mat-option")
    .filter({ hasText: pattern })
    .first();
  await option.waitFor({ state: "visible", timeout: 20000 });
  await option.scrollIntoViewIfNeeded().catch(() => {});
  await option.click({ force: true });

  await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(150);
}

async function waitMatSelectEnabled(page, selector) {
  await page.waitForSelector(selector, { timeout: 20000 });
  const loc = page.locator(selector);
  await page.waitForFunction(
    (el) => el.getAttribute("aria-disabled") !== "true",
    await loc.elementHandle(),
    { timeout: 20000 },
  );
}

/* ─────────────────────────────────────────
   DEBUG: salva HTML e innerText da página
───────────────────────────────────────── */

let debugContador = 0;
const DEBUG_HABILITADO = isTrue(getEnv("DEBUG_HTML", "false"));

async function debugSalvarPagina(page, numeroProcesso, outDir) {
  if (!DEBUG_HABILITADO && debugContador >= 3) return;
  debugContador++;

  const slug = (numeroProcesso || "sem_numero").replace(/[^0-9]/g, "");
  const base = path.join(outDir, `debug_${String(debugContador).padStart(2, "0")}_${slug}`);

  // Salva HTML completo
  try {
    const html = await page.content();
    fs.writeFileSync(base + ".html", html, "utf8");
  } catch {}

  // Salva texto visível (innerText) — muito mais fácil de ler
  try {
    const texto = await page.evaluate(() => document.body.innerText || "");
    fs.writeFileSync(base + ".txt", texto, "utf8");
    console.log(`🔍 Debug salvo: ${base}.txt`);
  } catch {}
}

/* ─────────────────────────────────────────
   NAVEGAÇÃO JTe
───────────────────────────────────────── */

async function abrirJTe(page) {
  console.log("➡️  Acessando JTe...");
  await page.goto("https://jte.csjt.jus.br/start", {
    waitUntil: "networkidle",
    timeout: 60000,
  });
  await page.waitForTimeout(2500);

  // Tela de autenticação: clica em "Não" para acesso anônimo
  try {
    const seletores = [
      page.getByRole("button", { name: /^não$/i }),
      page.getByRole("button", { name: /não autenticar/i }),
      page.getByRole("button", { name: /continuar sem/i }),
      page.getByRole("button", { name: /acesso público/i }),
      page.getByRole("button", { name: /acesso anônimo/i }),
      page.locator("ion-button").filter({ hasText: /^não$/i }),
      page.locator("button").filter({ hasText: /^não$/i }),
    ];

    for (const btn of seletores) {
      if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
        console.log('🔒 Tela de autenticação detectada → clicando em "Não"');
        await btn.click({ force: true });
        await page.waitForLoadState("networkidle").catch(() => {});
        await page.waitForTimeout(1000);
        break;
      }
    }
  } catch {}

  await retryOp(page, async () => {
    const loc = page.getByText("TRT2 - São Paulo", { exact: true });
    await loc.waitFor({ state: "visible", timeout: 20000 });
    await loc.click({ force: true });
  });

  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(800);
  console.log("✅ TRT-2 selecionado");
}

async function abrirModuloPauta(page) {
  console.log("➡️  Abrindo módulo Pauta...");
  await retryOp(page, async () => {
    const card = page
      .locator('ion-card-content.card-content-modulo:has-text("Pauta")')
      .first();
    await card.waitFor({ state: "visible", timeout: 20000 });
    await card.click({ force: true });
  });
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(800);
  console.log("✅ Módulo Pauta aberto");
}

/* ─────────────────────────────────────────
   LISTAR VARAS
───────────────────────────────────────── */

async function listarVaras(page) {
  console.log("➡️  Listando varas...");

  const botao = page.getByTestId("pautaButtonSelecaoUnidade");
  await botao.waitFor({ state: "visible", timeout: 20000 });
  await botao.click({ force: true });

  await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', {
    timeout: 20000,
  });

  await matSelectChoose(
    page,
    page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select'),
    "Audiências 1º grau",
  );

  await matSelectChoose(
    page,
    page.locator('mat-form-field[data-testid="municipio"] mat-select'),
    "São Paulo - Zonas Central, Norte e Oeste",
  );

  await waitMatSelectEnabled(
    page,
    'mat-form-field[data-testid="orgao"] mat-select',
  );

  const selectOrgao = page.locator(
    'mat-form-field[data-testid="orgao"] mat-select',
  );
  await selectOrgao.click({ force: true });

  const panel = page.locator(".mat-mdc-select-panel");
  await panel.waitFor({ state: "visible", timeout: 20000 });
  await page.waitForSelector(".mat-mdc-select-panel mat-option", {
    timeout: 20000,
  });

  const opcoes = panel.locator("mat-option");
  const total = await opcoes.count();
  const varas = [];

  for (let i = 0; i < total; i++) {
    const label = await opcoes
      .nth(i)
      .locator(".mdc-list-item__primary-text")
      .textContent()
      .catch(() => "");
    if (label && label.trim()) varas.push(label.trim());
  }

  await page.keyboard.press("Escape").catch(() => {});
  await page.getByTestId("ButtonCancelar").click().catch(() => {});

  console.log(`✅ ${varas.length} varas encontradas`);
  return varas;
}

/* ─────────────────────────────────────────
   SELECIONAR UNIDADE (VARA)
───────────────────────────────────────── */

async function selecionarUnidade(page, varaLabel) {
  console.log(`\n🏛️  Selecionando vara: ${varaLabel}`);
  await fecharOverlays(page);

  const botao = page.getByTestId("pautaButtonSelecaoUnidade");
  await botao.waitFor({ state: "visible", timeout: 20000 });
  await botao.click({ force: true });

  await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', {
    timeout: 20000,
  });

  await matSelectChoose(
    page,
    page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select'),
    "Audiências 1º grau",
  );

  await matSelectChoose(
    page,
    page.locator('mat-form-field[data-testid="municipio"] mat-select'),
    "São Paulo - Zonas Central, Norte e Oeste",
  );

  await waitMatSelectEnabled(
    page,
    'mat-form-field[data-testid="orgao"] mat-select',
  );

  await matSelectChoose(
    page,
    page.locator('mat-form-field[data-testid="orgao"] mat-select'),
    varaLabel,
  );

  const confirmar = page.getByTestId("ButtonConfirmar");
  await confirmar.waitFor({ state: "visible", timeout: 20000 });
  await confirmar.click({ delay: 80 }).catch(() => {});

  await page.waitForLoadState("networkidle").catch(() => {});
  await page.waitForTimeout(700);
}

/* ─────────────────────────────────────────
   NAVEGAÇÃO DE DATAS (botões prev/next)
───────────────────────────────────────── */

const SEL_BTN_NEXT =
  "#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(3) > ion-button";
const SEL_BTN_PREV =
  "#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(1) > ion-button";
const XPATH_BTN_DATA =
  '//*[@id="main-content"]/ng-component[3]/ion-content/div/div/ion-grid/ion-row[2]/ion-col[2]/ion-button';

function extrairDataBR(texto) {
  const m = String(texto ?? "").match(/\b\d{2}\/\d{2}\/\d{4}\b/);
  return m?.[0] ?? "";
}

async function lerDataExibida(page) {
  try {
    const raw = await page.evaluate((xp) => {
      const node = document.evaluate(
        xp,
        document,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null,
      ).singleNodeValue;
      if (!node) return "";
      const btn =
        node.shadowRoot?.querySelector("button") ||
        node.querySelector?.("button");
      return (
        btn?.innerText ||
        btn?.textContent ||
        node.innerText ||
        node.textContent ||
        ""
      ).trim();
    }, XPATH_BTN_DATA);
    if (raw) return raw;
  } catch {}

  try {
    return (
      await page
        .getByTestId("pautaButtonData")
        .innerText({ timeout: 800 })
        .catch(() => "")
    ).trim();
  } catch {}

  return "";
}

async function clicarBotaoIon(page, sel) {
  return page
    .evaluate((s) => {
      const host = document.querySelector(s);
      if (!host) return false;
      const btn =
        host.shadowRoot?.querySelector("button") ||
        host.querySelector?.("button");
      (btn || host).click();
      return true;
    }, sel)
    .catch(() => false);
}

async function navegarParaData(page, alvoBR, maxSteps = 220) {
  const alvoMs = parseBRDate(alvoBR).getTime();

  for (let step = 0; step < maxSteps; step++) {
    const raw = await lerDataExibida(page);
    if (raw && raw.includes(alvoBR)) return true;

    const atualBR = extrairDataBR(raw);
    const atualMs = atualBR ? parseBRDate(atualBR).getTime() : 0;

    const irParaTras = atualMs && atualMs > alvoMs;
    const antes = raw;

    const clicou = irParaTras
      ? await clicarBotaoIon(page, SEL_BTN_PREV)
      : await clicarBotaoIon(page, SEL_BTN_NEXT);

    if (!clicou) {
      await page.waitForTimeout(200);
      continue;
    }

    const inicio = Date.now();
    while (Date.now() - inicio < 2500) {
      const depois = await lerDataExibida(page);
      if (depois && depois !== antes) {
        if (depois.includes(alvoBR)) return true;
        break;
      }
      await page.waitForTimeout(80);
    }
  }

  return (await lerDataExibida(page)).includes(alvoBR);
}

async function selecionarData(page, dataBR, maxTentativas = 3) {
  for (let t = 1; t <= maxTentativas; t++) {
    await fecharOverlays(page);
    const ok = await navegarParaData(page, dataBR, 220);
    if (ok) return true;
    console.warn(`⚠️  Data não aplicou (${t}/${maxTentativas}): ${dataBR}`);
    await page.waitForTimeout(600);
  }
  return false;
}

/* ─────────────────────────────────────────
   ESPERAR PAUTA CARREGAR
───────────────────────────────────────── */

async function esperarPauta(page) {
  for (const sel of [
    "ion-spinner",
    ".mat-mdc-progress-spinner",
    ".mat-mdc-progress-bar",
  ]) {
    const sp = page.locator(sel).first();
    if (await sp.isVisible({ timeout: 300 }).catch(() => false)) {
      await sp.waitFor({ state: "hidden", timeout: 15000 }).catch(() => {});
    }
  }

  let last = -1;
  for (let i = 0; i < 20; i++) {
    const count = await page
      .locator("ion-list ion-item")
      .count()
      .catch(() => 0);
    if (count === last) {
      await page.waitForTimeout(600);
      const count2 = await page
        .locator("ion-list ion-item")
        .count()
        .catch(() => 0);
      if (count2 === count) return count;
    }
    last = count;
    await page.waitForTimeout(250);
  }

  return last;
}

/* ─────────────────────────────────────────
   EXTRAIR LISTA DA PAUTA
───────────────────────────────────────── */

async function extrairListaPauta(page) {
  const total = await page
    .locator("ion-list ion-item")
    .count()
    .catch(() => 0);
  if (!total) return [];

  return page.evaluate(() => {
    return Array.from(document.querySelectorAll("ion-list ion-item")).map(
      (item, idx) => {
        const getText = (sel) => {
          const el = item.querySelector(sel);
          return el ? el.textContent.replace(/\u00a0/g, " ").trim() : "";
        };

        const partes = Array.from(
          item.querySelectorAll(".item-desc-small.item-text-wrap"),
        )
          .map((e) => e.textContent.replace(/\u00a0/g, " ").trim())
          .filter(Boolean);

        return {
          idx,
          numeroProcesso: getText(".JT-item-texto-negrito"),
          sessao: [getText(".sessao"), getText(".palavrasRight")]
            .filter(Boolean)
            .join(" - "),
          juiz: partes[0] || "",
          reclamante: partes[1] || "",
        };
      },
    );
  });
}

/* ─────────────────────────────────────────
   VERIFICAR POLO PASSIVO SEM ADVOGADO
   ─────────────────────────────────────────
   Estratégia: usa innerText linha a linha para ser imune ao shadow DOM
   dos Web Components Ionic (ion-*).

   Retorna TRUE quando:
   a) A página carregou E não existe nenhum "Polo passivo" listado, OU
   b) Existe(m) polo(s) passivo(s) mas NENHUM deles tem advogado constituído
      (campo "Advogado(s)" ausente, vazio, "-" ou "—")

   Retorna FALSE (conservador) quando:
   - A página não carregou corretamente
   - Qualquer polo passivo tem advogado com nome preenchido
   - Não foi possível localizar o campo Advogado(s) após o polo passivo
     (evita falsos positivos por falha de parsing)
───────────────────────────────────────── */

async function verificarPoloPassivoSemAdvogado(page, outDir, numeroProcesso) {
  try {
    // Aguarda algum indicador de que a página de detalhes carregou
    await page
      .waitForSelector(
        'h1, .processo-numero, [class*="numero"], .titulo-processo, ion-title',
        { timeout: 30000 },
      )
      .catch(() => {});

    // Tempo extra — o JTe carrega partes de forma assíncrona
    await page.waitForTimeout(3000);

    // Salva debug dos primeiros processos para inspeção manual
    await debugSalvarPagina(page, numeroProcesso, outDir);

    return await page.evaluate(() => {
      // ── Coleta o texto visível completo da página ──
      // innerText respeita visibilidade CSS e quebras de linha reais,
      // sendo muito mais confiável que textContent para este caso.
      const textoCompleto = (document.body.innerText || "").trim();

      if (!textoCompleto) {
        console.log("[JTe] Página vazia — ignorando");
        return false;
      }

      const linhas = textoCompleto
        .split("\n")
        .map((l) => l.trim())
        .filter((l) => l.length > 0);

      // ── Verifica se a página de detalhes realmente carregou ──
      const paginaCarregou = linhas.some((l) =>
        /polo\s+ativo|autuaç|classe|vara\s+do\s+trabalho|reclamante|número\s+do\s+processo/i.test(l),
      );

      if (!paginaCarregou) {
        console.log("[JTe] Página não identificada como detalhes — ignorando");
        return false;
      }

      // ── Verifica se existe polo passivo ──
      const indicesPoloPassivo = [];
      for (let i = 0; i < linhas.length; i++) {
        if (/polo\s+passivo/i.test(linhas[i])) {
          indicesPoloPassivo.push(i);
        }
      }

      // Sem polo passivo algum → processo sem réu constituído
      if (indicesPoloPassivo.length === 0) {
        console.log("[JTe] Nenhum polo passivo encontrado → incluir");
        return true;
      }

      // ── Para cada polo passivo, verifica o campo Advogado(s) ──
      for (const idxPolo of indicesPoloPassivo) {
        let encontrouCampoAdv = false;
        let temAdvogado = false;

        // Varre as próximas linhas até encontrar outro polo ou limite de 15 linhas
        for (let j = idxPolo + 1; j < Math.min(idxPolo + 15, linhas.length); j++) {
          const linha = linhas[j];

          // Parou em outro bloco de polo → encerra busca para este polo
          if (/^polo\s+(passivo|ativo)/i.test(linha)) break;

          if (/^advogado/i.test(linha)) {
            encontrouCampoAdv = true;

            // O valor do advogado pode estar:
            // (a) na mesma linha após "Advogado(s):" → "Advogado(s): João Silva"
            // (b) na linha seguinte → linha j+1
            const mesmaLinha = linha.replace(/^advogado\(s\)\s*:?\s*/i, "").trim();
            const proximaLinha = linhas[j + 1] ?? "";

            const valor = mesmaLinha || proximaLinha;

            const vazio =
              !valor ||
              valor === "-" ||
              valor === "—" ||
              valor === "–" ||
              /^advogado/i.test(valor); // próxima linha é outro campo

            console.log(
              `[JTe] Polo passivo[${idxPolo}] advogado="${valor}" vazio=${vazio}`,
            );

            if (!vazio) {
              temAdvogado = true;
            }
            break;
          }
        }

        // ── Decisão conservadora ──
        // Se achou o campo e tem advogado → descartar processo
        if (encontrouCampoAdv && temAdvogado) {
          console.log("[JTe] Polo passivo COM advogado → descartar");
          return false;
        }

        // Se NÃO encontrou o campo Advogado(s) após o polo passivo,
        // não assume vazio — pode ser erro de parsing ou carregamento parcial.
        // Retorna false para não gerar falso positivo.
        if (!encontrouCampoAdv) {
          console.log(
            `[JTe] Campo Advogado(s) não encontrado após polo passivo[${idxPolo}] → descartar por segurança`,
          );
          return false;
        }
      }

      // Chegou aqui: todos os polos passivos têm campo Advogado(s) vazio/"-"
      console.log("[JTe] Todos os polos passivos SEM advogado → incluir");
      return true;
    });
  } catch (err) {
    console.warn(`⚠️  Erro ao verificar polo passivo: ${err.message}`);
    return false;
  }
}

/* ─────────────────────────────────────────
   VOLTAR PARA A PAUTA
───────────────────────────────────────── */

async function voltarParaPauta(page) {
  try {
    const seletoresVoltar = [
      page.locator("ion-back-button").first(),
      page.locator('ion-toolbar ion-button[fill="clear"]').first(),
      page.locator("ion-header ion-button").first(),
    ];

    for (const btn of seletoresVoltar) {
      if (await btn.isVisible({ timeout: 1000 }).catch(() => false)) {
        await btn.click({ force: true });
        await page.waitForLoadState("networkidle").catch(() => {});
        await page.waitForTimeout(1500);
        return;
      }
    }
  } catch {}

  // Fallback: navegação nativa do browser
  try {
    await page.goBack({ waitUntil: "networkidle", timeout: 15000 });
    await page.waitForTimeout(1500);
  } catch {}
}

/* ─────────────────────────────────────────
   MAIN
───────────────────────────────────────── */

async function main() {
  const geradoEm = new Date().toISOString();

  // ── Saída
  const outDir = path.join(process.cwd(), "output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const ts = Date.now();
  const csvPath = path.join(outDir, `sem_polo_passivo_${ts}.csv`);
  const xlsxPath = path.join(outDir, `sem_polo_passivo_${ts}.xlsx`);

  // ── Banco
  const pool = await inicializarBanco();

  // ── Browser
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const page = await browser.newPage();
  page.setDefaultTimeout(60000);
  page.setDefaultNavigationTimeout(60000);

  // Captura logs do console do browser (útil para ver os [JTe] debug)
  page.on("console", (msg) => {
    if (msg.text().includes("[JTe]")) {
      console.log(`   🖥️  ${msg.text()}`);
    }
  });

  const resultados = [];

  try {
    await abrirJTe(page);
    await abrirModuloPauta(page);

    const todasVaras = await listarVaras(page);
    const varas = todasVaras.slice(0, 3); // ← limita a 3 varas para teste
    const datas = gerarDatasProximos7Dias();

    console.log(
      `\n🧪 MODO TESTE: usando ${varas.length}/${todasVaras.length} varas`,
    );
    console.log(`   Varas selecionadas: ${varas.join(" | ")}`);
    console.log(
      `\n🚀 Iniciando coleta: ${varas.length} varas × ${datas.length} dias úteis\n`,
    );

    if (DEBUG_HABILITADO) {
      console.log("🔍 DEBUG_HTML=true → salvando HTML/TXT dos primeiros processos em output/");
    } else {
      console.log("💡 Dica: defina DEBUG_HTML=true no .env para salvar HTML dos primeiros processos");
    }

    for (const vara of varas) {
      await selecionarUnidade(page, vara);

      for (const dataBR of datas) {
        console.log(`\n📅 ${vara} | ${dataBR}`);

        const okData = await selecionarData(page, dataBR, 3);
        if (!okData) {
          console.warn(`⚠️  Pulando: data não aplicou (${dataBR})`);
          continue;
        }

        await page.waitForTimeout(2000);
        const totalItens = await esperarPauta(page);

        if (!totalItens) {
          console.log(`   (sem audiências)`);
          continue;
        }

        const processos = await extrairListaPauta(page);
        console.log(`   📋 ${processos.length} processo(s) na pauta`);

        for (let i = 0; i < processos.length; i++) {
          const p = processos[i];
          if (!p.numeroProcesso) continue;

          try {
            // Re-busca o item pelo índice para evitar referências stale
            const items = page.locator("ion-list ion-item");
            const item = items.nth(p.idx);

            await item.scrollIntoViewIfNeeded().catch(() => {});

            // Tenta clicar no botão de 3 pontos (⋮)
            const menuBtn = item
              .locator(
                'ion-button[slot="end"], ion-button.more-button, ion-button:last-of-type',
              )
              .first();

            const menuBtnVisivel = await menuBtn
              .isVisible({ timeout: 1000 })
              .catch(() => false);

            if (menuBtnVisivel) {
              await menuBtn.click({ force: true });
            } else {
              await item.click({ force: true });
            }

            // Aguarda menu contextual e clica em "Detalhes do processo"
            const detalhesOpcao = page
              .locator("text=Detalhes do processo")
              .first();
            const menuAbriu = await detalhesOpcao
              .isVisible({ timeout: 4000 })
              .catch(() => false);

            if (menuAbriu) {
              await detalhesOpcao.click({ force: true });
            }

            await page.waitForTimeout(2000);
            await page.waitForLoadState("networkidle").catch(() => {});

            // Passa outDir e número do processo para permitir debug
            const semAdvogado = await verificarPoloPassivoSemAdvogado(
              page,
              outDir,
              p.numeroProcesso,
            );

            if (semAdvogado) {
              console.log(
                `   ✅ Polo passivo SEM advogado → ${p.numeroProcesso}`,
              );

              const registro = {
                vara,
                data: dataBR,
                numeroProcesso: p.numeroProcesso,
                sessao: p.sessao,
                juiz: p.juiz,
                reclamante: p.reclamante,
              };

              resultados.push(registro);
              await salvarNoBanco(pool, geradoEm, vara, dataBR, p);
            } else {
              console.log(
                `   — ${p.numeroProcesso} (polo passivo com advogado ou parsing inconclusivo)`,
              );
            }

            await voltarParaPauta(page);
            await page.waitForTimeout(1000 + Math.floor(Math.random() * 800));

          } catch (err) {
            console.warn(
              `   ⚠️  Erro no processo ${p.numeroProcesso}: ${err.message}`,
            );
            await voltarParaPauta(page);
            await page.waitForTimeout(2000);
          }
        }
      }
    }

    // ── Salvar resultados ──
    console.log(`\n📦 Total sem polo passivo: ${resultados.length}`);

    const headers = [
      "vara",
      "data",
      "numeroProcesso",
      "sessao",
      "juiz",
      "reclamante",
    ];
    writeCsv(csvPath, headers, resultados);
    await gerarXLSX(xlsxPath, resultados);
    await enviarEmailComAnexo(xlsxPath, resultados.length);

    console.log("\n✅ Execução concluída com sucesso!");

  } catch (err) {
    console.error("❌ Erro fatal:", err);

    if (resultados.length > 0) {
      const headers = [
        "vara",
        "data",
        "numeroProcesso",
        "sessao",
        "juiz",
        "reclamante",
      ];
      writeCsv(csvPath, headers, resultados);
      await gerarXLSX(xlsxPath, resultados).catch(() => {});
      console.log(`💾 Parcial salvo: ${resultados.length} registros`);
    }
  } finally {
    if (pool) await pool.end().catch(() => {});
    await browser.close().catch(() => {});
  }
}

main();