require("dotenv").config();

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

/* =========================
   LOGGING
========================= */

const LOG_DIR = path.join(process.cwd(), "output");
if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });

const logStream = fs.createWriteStream(
  path.join(LOG_DIR, `log_${Date.now()}.txt`),
  { flags: "a" }
);

function log(msg) {
  const line = `[${new Date().toISOString()}] ${msg}`;
  console.log(line);
  logStream.write(line + "\n");
}

function warn(msg) {
  const line = `[${new Date().toISOString()}] WARN: ${msg}`;
  console.warn(line);
  logStream.write(line + "\n");
}

/* =========================
   CSV HELPERS
========================= */

function csvEscape(value) {
  if (value === null || value === undefined) return "";
  const s = String(value);
  if (/[;"\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

const CSV_HEADERS = [
  "geradoEm",
  "vara",
  "data",
  "numeroProcesso",
  "hora",
  "status",
  "juiz",
  "reclamante",
  "reclamada",
  "polosPassivosSemAdvogado",
];

function appendCsvRows(filePath, rows) {
  const fileExists = fs.existsSync(filePath);
  const lines = [];
  if (!fileExists) {
    lines.push("﻿" + CSV_HEADERS.map(csvEscape).join(";"));
  }
  for (const row of rows) {
    lines.push(CSV_HEADERS.map((h) => csvEscape(row[h])).join(";"));
  }
  fs.appendFileSync(filePath, (fileExists ? "\n" : "") + lines.join("\n"), "utf8");
}

/* =========================
   DATE HELPERS
========================= */

function parseBRDate(br) {
  const [dd, mm, yyyy] = br.split("/").map(Number);
  return new Date(yyyy, mm - 1, dd);
}

function isWeekdayBR(dataBR) {
  const d = parseBRDate(dataBR);
  const day = d.getDay();
  return day !== 0 && day !== 6;
}

function dataJaPassou(dataBR) {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return parseBRDate(dataBR) < hoje;
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
    const dd = String(current.getDate()).padStart(2, "0");
    const mm = String(current.getMonth() + 1).padStart(2, "0");
    const dataBR = `${dd}/${mm}/${current.getFullYear()}`;
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
    const backdrop = page.locator(".cdk-overlay-backdrop");
    await page.keyboard.press("Escape").catch(() => {});
    await backdrop.waitFor({ state: "detached", timeout: 1500 }).catch(() => {});
    if (await backdrop.isVisible({ timeout: 400 }).catch(() => false)) {
      await backdrop.click({ position: { x: 5, y: 5 }, force: true }).catch(() => {});
    }
    await backdrop.waitFor({ state: "detached", timeout: 1500 }).catch(() => {});
  } catch {}
}

async function retryOperation(page, operation, maxRetries = 5, delayMs = 1200) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await operation();
    } catch (err) {
      warn(`Tentativa ${attempt}/${maxRetries} falhou: ${err.message}`);
      if (attempt === maxRetries) throw err;
      await fecharOverlays(page);
      await page.waitForTimeout(delayMs);
    }
  }
}

/* =========================
   SELEÇÃO DE FILTROS
========================= */

function escRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function matSel(page, loc, txt) {
  await loc.scrollIntoViewIfNeeded().catch(() => {});
  await loc.click({ force: true });
  const panel = page.locator(".mat-mdc-select-panel");
  await panel.waitFor({ state: "visible", timeout: 20000 });
  const opt = panel
    .locator("mat-option")
    .filter({ hasText: new RegExp(`^\\s*${escRe(txt)}\\s*$`, "i") })
    .first();
  await opt.waitFor({ state: "visible", timeout: 20000 });
  await opt.scrollIntoViewIfNeeded().catch(() => {});
  await opt.click({ force: true });
  await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(150);
}

async function abrirPainelEFiltrar(page) {
  const botaoUnidade = page.getByTestId("pautaButtonSelecaoUnidade");
  await botaoUnidade.waitFor({ state: "visible", timeout: 20000 });
  await botaoUnidade.click({ force: true });
  await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', { timeout: 20000 });
  await matSel(page, page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select'), "Audiências 1º grau");
  await matSel(page, page.locator('mat-form-field[data-testid="municipio"] mat-select'), "São Paulo - Zonas Central, Norte e Oeste");
  await page.waitForSelector('mat-form-field[data-testid="orgao"] mat-select[aria-disabled="false"]', { timeout: 20000 });
  return page.locator('mat-form-field[data-testid="orgao"] mat-select');
}

/* =========================
   NAVEGAÇÃO JTe
========================= */

async function abrirJTeSelecionarTRT2(page) {
  log("Acessando JTe...");
  await page.goto("https://jte.csjt.jus.br/start", { waitUntil: "networkidle", timeout: 60000 });
  await page.waitForLoadState("domcontentloaded");
  try {
    const seletores = [
      page.getByRole("button", { name: /^não$/i }),
      page.getByRole("button", { name: /não autenticar/i }),
      page.getByRole("button", { name: /continuar sem/i }),
      page.locator("ion-button").filter({ hasText: /^não$/i }),
      page.locator("button").filter({ hasText: /^não$/i }),
    ];
    for (const btn of seletores) {
      if (await btn.isVisible({ timeout: 2000 }).catch(() => false)) {
        await btn.click({ force: true });
        await page.waitForLoadState("networkidle").catch(() => {});
        await page.waitForTimeout(800);
        break;
      }
    }
  } catch {}
  const locator = page.getByText("TRT2 - São Paulo", { exact: true });
  await retryOperation(page, async () => {
    await locator.waitFor({ state: "visible", timeout: 20000 });
    await locator.click({ force: true });
  });
  await page.waitForLoadState("networkidle");
  log("TRT2 selecionado");
}

async function abrirModuloPauta(page) {
  log("Abrindo módulo Pauta...");
  const card = page.locator('ion-card-content.card-content-modulo:has-text("Pauta")').first();
  await retryOperation(page, async () => {
    await card.waitFor({ state: "visible", timeout: 20000 });
    await card.click({ force: true });
  });
  await page.waitForLoadState("networkidle");
  log("Módulo Pauta aberto");
}

/* =========================
   LISTAR VARAS
========================= */

async function listarVaras(page) {
  log("Listando varas...");
  const selectOrgao = await abrirPainelEFiltrar(page);
  await selectOrgao.click();
  const opcoes = page.locator(".mat-mdc-select-panel mat-option");
  const total = await opcoes.count();
  const varas = [];
  for (let i = 0; i < total; i++) {
    const label = await opcoes.nth(i).locator(".mdc-list-item__primary-text").textContent();
    if (label) varas.push(label.trim());
  }
  await page.keyboard.press("Escape").catch(() => {});
  await page.getByTestId("ButtonCancelar").click().catch(() => {});
  log(`${varas.length} varas encontradas`);
  return varas;
}

/* =========================
   SELECIONAR UNIDADE
========================= */

async function selecionarUnidade(page, varaLabel) {
  log(`Selecionando vara: ${varaLabel}`);
  await fecharOverlays(page);
  const selectOrgao = await abrirPainelEFiltrar(page);
  await selectOrgao.click();
  const opcao = page.locator(".mat-mdc-select-panel mat-option").filter({ hasText: varaLabel }).first();
  await opcao.click({ force: true });
  await page.getByTestId("ButtonConfirmar").click({ delay: 80 }).catch(() => {});
  await page.waitForLoadState("networkidle");
}

/* =========================
   NAVEGAÇÃO DE DATA (setas prev/next)
========================= */

const XDATA = '//*[@id="main-content"]/ng-component[3]/ion-content/div/div/ion-grid/ion-row[2]/ion-col[2]/ion-button';
const SNEXT = "#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(3) > ion-button";
const SPREV = "#main-content > ng-component:nth-child(3) > ion-content > div > div > ion-grid > ion-row:nth-child(2) > ion-col:nth-child(1) > ion-button";

async function lerData(page) {
  try {
    const r = await page.evaluate((xp) => {
      const n = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      if (!n) return "";
      const b = n.shadowRoot?.querySelector("button") || n.querySelector?.("button");
      return (b?.innerText || b?.textContent || n.innerText || "").trim();
    }, XDATA);
    if (r) return r;
  } catch {}
  return "";
}

async function ionClick(page, sel) {
  return page.evaluate((s) => {
    const h = document.querySelector(s);
    if (!h) return false;
    (h.shadowRoot?.querySelector("button") || h.querySelector?.("button") || h).click();
    return true;
  }, sel).catch(() => false);
}

async function navData(page, alvo) {
  const am = parseBRDate(alvo).getTime();
  for (let i = 0; i < 220; i++) {
    const r = await lerData(page);
    if (r?.includes(alvo)) return true;
    const m = String(r ?? "").match(/\b\d{2}\/\d{2}\/\d{4}\b/);
    const cm = m ? parseBRDate(m[0]).getTime() : 0;
    const ant = r;
    const ok = cm && cm > am ? await ionClick(page, SPREV) : await ionClick(page, SNEXT);
    if (!ok) { await page.waitForTimeout(200); continue; }
    const t0 = Date.now();
    while (Date.now() - t0 < 2500) {
      const dep = await lerData(page);
      if (dep && dep !== ant) { if (dep.includes(alvo)) return true; break; }
      await page.waitForTimeout(80);
    }
  }
  return (await lerData(page)).includes(alvo);
}

/* =========================
   ESPERAR PAUTA ESTABILIZAR
========================= */

const PROC = "ion-item:has(.JT-item-texto-negrito)";

async function esperarPautaEstabilizar(page) {
  for (const s of ["ion-spinner", ".mat-mdc-progress-spinner", ".mat-mdc-progress-bar"]) {
    const sp = page.locator(s).first();
    if (await sp.isVisible({ timeout: 300 }).catch(() => false))
      await sp.waitFor({ state: "hidden", timeout: 15000 }).catch(() => {});
  }
  let last = -1;
  for (let i = 0; i < 20; i++) {
    const count = await page.locator(PROC).count().catch(() => 0);
    if (count === last) {
      await page.waitForTimeout(600);
      if ((await page.locator(PROC).count().catch(() => 0)) === count) return;
    }
    last = count;
    await page.waitForTimeout(250);
  }
}

/* =========================
   EXTRAÇÃO DA LISTA
========================= */

async function extrairProcessosDaPauta(page) {
  return page.evaluate(() => {
    return Array.from(document.querySelectorAll("ion-item"))
      .map((item) => {
        const numEl = item.querySelector(".JT-item-texto-negrito");
        if (!numEl) return null;
        const getText = (sel) => {
          const el = item.querySelector(sel);
          return el ? el.textContent.replace(/ /g, " ").trim() : "";
        };
        const partes = Array.from(item.querySelectorAll(".item-desc-small.item-text-wrap"))
          .map((e) => e.textContent.replace(/ /g, " ").trim())
          .filter(Boolean);
        return {
          numeroProcesso: numEl.textContent.replace(/ /g, " ").trim(),
          hora: getText(".sessao"),
          status: getText(".palavrasRight"),
          juiz: partes[0] || "",
          reclamante: (partes[1] || "").replace(/ X$/i, "").trim(),
          reclamada: (partes[2] || "").replace(/\s+/g, " ").trim(),
        };
      })
      .filter(Boolean);
  });
}

/* =========================
   FILTROS DE STATUS
========================= */

const STATUS_EXCLUIR = ["realizada", "cancelada", "adiada", "suspensa", "arquivada", "em andamento"];

function filtrarProcessosAtivos(processos) {
  return processos.filter((p) => {
    if (!p.numeroProcesso) return false;
    const st = p.status.toLowerCase();
    return !STATUS_EXCLUIR.some((excl) => st.includes(excl));
  });
}

/* =========================
   VERIFICAÇÃO DE POLO PASSIVO VIA 3 PONTOS → DETALHES
========================= */

function analisarPoloPassivo(texto) {
  const polos = [];
  const regex = /Polo passivo:\s*(.+)/gi;
  let match;
  while ((match = regex.exec(texto)) !== null) {
    const nome = match[1].trim();
    const restante = texto.slice(match.index + match[0].length, match.index + match[0].length + 200);
    const temAdvogado = /Advogado\(s\)\s*\n\s*\d+/i.test(restante);
    polos.push({ nome, temAdvogado });
  }
  return polos;
}

async function verificarDetalhes(page, procItem) {
  try {
    const tresPontos = procItem.locator('ion-icon[name="ellipsis-vertical-outline"]').first();
    await tresPontos.scrollIntoViewIfNeeded().catch(() => {});
    await tresPontos.click({ force: true });
    await page.waitForTimeout(800);

    const popover = page.locator("ion-popover").first();
    await popover.waitFor({ state: "visible", timeout: 5000 });

    const btnDetalhes = popover.locator("ion-item, button").filter({ hasText: /detalh/i }).first();
    await btnDetalhes.click({ force: true });
    await page.waitForTimeout(1500);
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.waitForTimeout(500);

    const texto = await page.evaluate(() => document.body.innerText || "");
    const polos = analisarPoloPassivo(texto);

    // Voltar via seta ←
    const setaVoltar = page.locator("ion-back-button, ion-button").filter({
      has: page.locator('ion-icon[name="arrow-back"]'),
    }).first();

    if (await setaVoltar.isVisible({ timeout: 2000 }).catch(() => false)) {
      await setaVoltar.click({ force: true });
    } else {
      await page.goBack().catch(() => {});
    }
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.waitForTimeout(1500);

    return polos;
  } catch (e) {
    warn(`Erro ao verificar detalhes: ${e.message}`);
    // Tentar voltar mesmo com erro
    try {
      const setaVoltar = page.locator("ion-back-button, ion-button").filter({
        has: page.locator('ion-icon[name="arrow-back"]'),
      }).first();
      if (await setaVoltar.isVisible({ timeout: 1000 }).catch(() => false)) {
        await setaVoltar.click({ force: true });
        await page.waitForLoadState("networkidle").catch(() => {});
        await page.waitForTimeout(1000);
      }
    } catch {}
    return null;
  }
}

function temPoloPassivoSemAdvogado(polos) {
  if (!polos || polos.length === 0) return false;
  return polos.some((p) => !p.temAdvogado);
}

function nomesPolosSemAdvogado(polos) {
  if (!polos) return "";
  return polos
    .filter((p) => !p.temAdvogado)
    .map((p) => p.nome)
    .join(" | ");
}

/* =========================
   MAIN
========================= */

const HEADLESS = process.env.HEADLESS !== "false";

async function main() {
  const browser = await chromium.launch({ headless: HEADLESS });
  const page = await browser.newPage();
  page.setDefaultTimeout(45000);
  page.setDefaultNavigationTimeout(60000);

  const outDir = path.join(process.cwd(), "output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const csvPath = path.join(outDir, `pauta_trt2_2meses_${Date.now()}.csv`);

  let totalRows = 0;
  let totalVerificados = 0;
  let totalLeads = 0;

  try {
    await abrirJTeSelecionarTRT2(page);
    await abrirModuloPauta(page);

    const varas = await listarVaras(page);
    const datas = gerarDatasProximosDoisMeses();
    log(`${varas.length} varas | ${datas.length} datas alvo`);

    for (let vi = 0; vi < varas.length; vi++) {
      const vara = varas[vi];
      await selecionarUnidade(page, vara);

      const varaRows = [];

      for (const dataBR of datas) {
        if (dataJaPassou(dataBR)) continue;

        log(`[${vi + 1}/${varas.length}] ${vara} | ${dataBR}`);

        const ok = await navData(page, dataBR);
        if (!ok) {
          warn(`Pulando data (navegação falhou): ${vara} | ${dataBR}`);
          continue;
        }

        await esperarPautaEstabilizar(page);

        const processosRaw = await extrairProcessosDaPauta(page);
        const processos = filtrarProcessosAtivos(processosRaw);

        if (!processos.length) {
          log(`${vara} | ${dataBR} | 0 processos ativos (${processosRaw.length} brutos)`);
          continue;
        }

        log(`${vara} | ${dataBR} | ${processos.length} ativos — verificando polo passivo...`);

        for (let pi = 0; pi < processos.length; pi++) {
          const p = processos[pi];
          totalVerificados++;

          const procItem = page.locator(PROC).nth(pi);
          const polos = await verificarDetalhes(page, procItem);

          if (temPoloPassivoSemAdvogado(polos)) {
            totalLeads++;
            const nomes = nomesPolosSemAdvogado(polos);
            log(`  ✅ LEAD: ${p.numeroProcesso} → ${nomes}`);

            varaRows.push({
              geradoEm: new Date().toISOString(),
              vara,
              data: dataBR,
              numeroProcesso: p.numeroProcesso,
              hora: p.hora,
              status: p.status,
              juiz: p.juiz,
              reclamante: p.reclamante,
              reclamada: p.reclamada,
              polosPassivosSemAdvogado: nomes,
            });
          }
        }

        log(`${vara} | ${dataBR} | ${processos.length} verificados, leads até agora: ${totalLeads}`);
      }

      if (varaRows.length > 0) {
        appendCsvRows(csvPath, varaRows);
        totalRows += varaRows.length;
        log(`Vara "${vara}" salva (${varaRows.length} leads, total: ${totalRows})`);
      }
    }

    log(`Concluído! ${totalVerificados} processos verificados | ${totalLeads} leads em ${csvPath}`);
  } catch (err) {
    log(`ERRO FATAL: ${err.message}`);
    console.error(err);
    process.exitCode = 1;
  } finally {
    await browser.close().catch(() => {});
    logStream.end();
  }
}

process.on("SIGINT", () => {
  log("Interrompido pelo usuário (SIGINT)");
  process.exit(130);
});

main().catch((e) => {
  console.error("ERRO:", e);
  process.exitCode = 1;
});
