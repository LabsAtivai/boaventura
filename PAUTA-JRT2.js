const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

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
   OVERLAY & RETRY
========================= */

async function fecharOverlays(page) {
  try {
    const backdrop = page.locator('.cdk-overlay-backdrop');
    await page.keyboard.press('Escape').catch(() => {});
    await page.waitForTimeout(200);

    if (await backdrop.isVisible({ timeout: 400 }).catch(() => false)) {
      await backdrop.click({ position: { x: 5, y: 5 }, force: true }).catch(() => {});
    }
    await backdrop.waitFor({ state: 'detached', timeout: 1500 }).catch(() => {});
  } catch {}
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

  await page.keyboard.press('Escape').catch(() => {});
  await page.getByTestId('ButtonCancelar').click().catch(() => {});

  console.log(`✅ ${varas.length} varas encontradas`);
  return varas;
}

/* =========================
   SELECIONAR UNIDADE
========================= */

async function selecionarUnidade(page, varaLabel) {
  console.log(`\n🏛️ Selecionando vara: ${varaLabel}`);

  await fecharOverlays(page);

  const botaoUnidade = page.getByTestId('pautaButtonSelecaoUnidade');
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

  const opcao = page.locator('.mat-mdc-select-panel mat-option').filter({ hasText: varaLabel }).first();
  await opcao.click({ force: true });

  await page.getByTestId('ButtonConfirmar').click({ delay: 80 }).catch(() => {});
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(700);

  // Reset para hoje
  const todayBR = getTodayBR();
  console.log(`🔄 Resetando para data de hoje: ${todayBR}`);
  const ok = await selecionarDataComConfirmacao(page, todayBR, 3);
  if (ok) console.log(`✅ Reset para hoje bem-sucedido`);
  else console.warn(`⚠️ Não conseguiu resetar para hoje. Seguindo mesmo assim.`);
}

/* =========================
   SELEÇÃO DE DATA (JTe ROBUSTA)
   - header tolerante (mês abreviado)
   - aria-label pode estar no TD ou no BUTTON
========================= */

function sameMonthYear(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

async function getCalendarHeaderDateApprox(page) {
  const periodBtn = page.locator('mat-calendar button.mat-calendar-period-button').first();
  const header = ((await periodBtn.textContent().catch(() => '')) || '').trim().toLowerCase();

  const mYear = header.match(/(19|20)\d{2}/);
  if (!mYear) return null;
  const year = Number(mYear[0]);

  const monthMap = [
    { idx: 0, keys: ['jan', 'janeiro', 'january'] },
    { idx: 1, keys: ['fev', 'fevereiro', 'feb', 'february'] },
    { idx: 2, keys: ['mar', 'março', 'marco', 'march'] },
    { idx: 3, keys: ['abr', 'abril', 'apr', 'april'] },
    { idx: 4, keys: ['mai', 'maio', 'may'] },
    { idx: 5, keys: ['jun', 'junho', 'june'] },
    { idx: 6, keys: ['jul', 'julho', 'july'] },
    { idx: 7, keys: ['ago', 'agosto', 'aug', 'august'] },
    { idx: 8, keys: ['set', 'setembro', 'sep', 'september'] },
    { idx: 9, keys: ['out', 'outubro', 'oct', 'october'] },
    { idx: 10, keys: ['nov', 'novembro', 'november'] },
    { idx: 11, keys: ['dez', 'dezembro', 'dec', 'december'] },
  ];

  const found = monthMap.find(m => m.keys.some(k => header.includes(k)));
  if (!found) return null;

  return new Date(year, found.idx, 1);
}

async function navegarParaMes(page, targetDate, maxSteps = 36) {
  const nextBtn = page.locator('mat-calendar button.mat-calendar-next-button').first();
  const prevBtn = page.locator('mat-calendar button.mat-calendar-previous-button').first();

  for (let i = 0; i < maxSteps; i++) {
    const approx = await getCalendarHeaderDateApprox(page);

    if (!approx) {
      // se não conseguir ler header, tenta ir pra frente e segue
      await nextBtn.click({ force: true }).catch(() => {});
      await page.waitForTimeout(120);
      continue;
    }

    if (sameMonthYear(approx, targetDate)) return true;

    const curKey = approx.getFullYear() * 12 + approx.getMonth();
    const tgtKey = targetDate.getFullYear() * 12 + targetDate.getMonth();

    if (tgtKey > curKey) await nextBtn.click({ force: true });
    else await prevBtn.click({ force: true });

    await page.waitForTimeout(120);
  }

  return false;
}

async function clicarDiaNoCalendario(page, targetDate) {
  const day = targetDate.getDate();
  const year = targetDate.getFullYear();

  const ok = await page.evaluate(({ day, year }) => {
    const cal = document.querySelector('mat-calendar');
    if (!cal) return false;

    // 1) td com aria-label
    const tds = Array.from(cal.querySelectorAll('td.mat-calendar-body-cell[aria-label]'));
    for (const td of tds) {
      const aria = (td.getAttribute('aria-label') || '').trim();
      const txt = (td.textContent || '').trim();
      if (aria.includes(String(year)) && txt === String(day)) {
        const btn = td.querySelector('button');
        (btn || td).click();
        return true;
      }
    }

    // 2) button com aria-label
    const btns = Array.from(cal.querySelectorAll('td.mat-calendar-body-cell button[aria-label]'));
    for (const btn of btns) {
      const aria = (btn.getAttribute('aria-label') || '').trim();
      const txt = (btn.textContent || '').trim();
      if (aria.includes(String(year)) && txt === String(day)) {
        btn.click();
        return true;
      }
    }

    // 3) fallback por texto do dia (não disabled)
    const any = Array.from(cal.querySelectorAll('td.mat-calendar-body-cell:not(.mat-calendar-body-disabled)'));
    for (const td of any) {
      const txt = (td.textContent || '').trim();
      if (txt === String(day)) {
        const btn = td.querySelector('button');
        (btn || td).click();
        return true;
      }
    }

    return false;
  }, { day, year });

  return ok;
}

async function selecionarDataNoCalendario(page, dataBR) {
  await fecharOverlays(page);

  const targetDate = parseBRDate(dataBR);

  const botaoData = page.getByTestId('pautaButtonData');
  await botaoData.waitFor({ state: 'visible', timeout: 20000 });
  await botaoData.click({ force: true });

  await page.waitForSelector('mat-calendar', { timeout: 15000 });

  const okMes = await navegarParaMes(page, targetDate, 36);
  if (!okMes) console.warn('⚠️ Não consegui navegar até o mês alvo. Tentando clicar o dia mesmo assim...');

  const okDia = await clicarDiaNoCalendario(page, targetDate);

  await page.waitForTimeout(250);
  await fecharOverlays(page);

  return okDia;
}

async function selecionarDataComConfirmacao(page, dataBR, maxTentativas = 3) {
  for (let t = 1; t <= maxTentativas; t++) {
    const okClique = await selecionarDataNoCalendario(page, dataBR);

    if (!okClique) {
      console.warn(`⚠️ [${t}/${maxTentativas}] Clique no calendário falhou para ${dataBR}`);
      await page.waitForTimeout(500);
      continue;
    }

    const dataVisivel = (await page.getByTestId('pautaButtonData').innerText().catch(() => '')).trim();
    console.log(`🧾 Data visível no botão: ${dataVisivel} | alvo: ${dataBR}`);

    if (dataVisivel === dataBR) return true;

    console.warn(`⚠️ [${t}/${maxTentativas}] Data NÃO aplicou (visível=${dataVisivel}). Retentando...`);
    await page.waitForTimeout(700);
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
      await sp.waitFor({ state: 'hidden', timeout: 15000 }).catch(() => {});
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
  const browser = await chromium.launch({ headless: false, slowMo: 100 });
  const page = await browser.newPage();

  const geradoEm = new Date().toISOString();
  const rowsCsv = [];
  const headers = ['geradoEm', 'vara', 'data', 'numeroProcesso', 'sessao', 'juiz', 'reclamante', 'reclamada'];

  try {
    await abrirJTeSelecionarTRT2(page);
    await abrirModuloPauta(page);

    const varas = await listarVaras(page);
    const varasAlvo = varas; // para teste: varas.slice(0, 5)

    const datas = gerarDatasProximosDoisMeses();
    console.log(`📅 ${datas.length} datas alvo`);

    for (const vara of varasAlvo) {
      await selecionarUnidade(page, vara);

      for (const dataBR of datas) {
        console.log(`📅 Selecionando data: ${dataBR}`);

        const ok = await selecionarDataComConfirmacao(page, dataBR, 3);
        if (!ok) {
          console.warn(`⚠️ Pulando data (não aplicou): ${vara} | ${dataBR}`);
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
      }
    }

    const outDir = path.join(process.cwd(), 'output');
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const filePath = path.join(outDir, `pauta_trt2_2meses_${Date.now()}.csv`);
    writeCsv(filePath, headers, rowsCsv);

    console.log(`\n✅ Concluído! ${rowsCsv.length} linhas em ${filePath}`);
  } catch (err) {
    console.error('❌ Erro:', err);
  } finally {
    // await browser.close().catch(() => {});
  }
}

main();
