require("dotenv").config();

const { Queue, Worker } = require("bullmq");
const { chromium } = require("playwright");
const mysql = require("mysql2/promise");

const CONCURRENCY = Number(process.env.CONCURRENCY || 4);
const MESES = Number(process.env.MESES || 6);
const VARAS_LIMIT = Number(process.env.VARAS || 0);
const ONLY_QUEUE = process.argv.includes("--only-queue");
const ONLY_WORK = process.argv.includes("--only-work");
const QUEUE_NAME = "jte";
const REDIS = {
  host: process.env.REDIS_HOST || "127.0.0.1",
  port: Number(process.env.REDIS_PORT || 6379),
  password: process.env.REDIS_PASS || undefined,
};

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";

/* ── Helpers de data ───────────────────────────────────────── */
function isWeekday(d) { const w = d.getDay(); return w !== 0 && w !== 6; }

function brToIso(br) {
  const [dd, mm, yyyy] = br.split("/").map(Number);
  return `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
}

function parseBR(br) {
  const [dd, mm, yyyy] = br.split("/").map(Number);
  return new Date(yyyy, mm - 1, dd);
}

function dataJaPassou(dataBR) {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return parseBR(dataBR) < hoje;
}

function gerarDatas() {
  const i = new Date(); i.setMonth(i.getMonth() + 1);
  const f = new Date(i); f.setMonth(f.getMonth() + MESES);
  const datas = [];
  for (let d = new Date(i); d <= f; d.setDate(d.getDate() + 1)) {
    if (!isWeekday(d)) continue;
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    datas.push(`${dd}/${mm}/${d.getFullYear()}`);
  }
  return datas;
}

/* ── DB pool ───────────────────────────────────────────────── */
async function dbPool() {
  if (process.env.DB_ENABLED === "false") return null;
  const p = mysql.createPool({
    host: process.env.DB_HOST || "127.0.0.1",
    port: Number(process.env.DB_PORT || 3306),
    user: process.env.DB_USER || "jte",
    password: process.env.DB_PASS || "",
    database: process.env.DB_NAME || "jte",
    waitForConnections: true,
    connectionLimit: CONCURRENCY * 3,
  });
  await p.query("SELECT 1");
  return p;
}

/* ── Browser helpers ───────────────────────────────────────── */
async function novaPage(browser) {
  const ctx = await browser.newContext({ userAgent: UA });
  const page = await ctx.newPage();
  page.setDefaultTimeout(45000);
  page.setDefaultNavigationTimeout(60000);
  return page;
}

function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

async function fecharOverlays(page) {
  try {
    await page.keyboard.press("Escape").catch(() => {});
    await page.waitForTimeout(200);
    const b = page.locator(".cdk-overlay-backdrop");
    if (await b.isVisible({ timeout: 400 }).catch(() => false))
      await b.click({ position: { x: 5, y: 5 }, force: true }).catch(() => {});
    await b.waitFor({ state: "detached", timeout: 1500 }).catch(() => {});
  } catch {}
}

async function matSel(page, loc, txt) {
  await loc.scrollIntoViewIfNeeded().catch(() => {});
  await loc.click({ force: true });
  const panel = page.locator(".mat-mdc-select-panel");
  await panel.waitFor({ state: "visible", timeout: 20000 });
  const opt = panel.locator("mat-option")
    .filter({ hasText: new RegExp(`^\\s*${escRe(txt)}\\s*$`, "i") }).first();
  await opt.waitFor({ state: "visible", timeout: 20000 });
  await opt.scrollIntoViewIfNeeded().catch(() => {});
  await opt.click({ force: true });
  await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(150);
}

async function waitEnabled(page, sel) {
  const loc = page.locator(sel);
  await loc.waitFor({ timeout: 20000 });
  await page.waitForFunction(
    (el) => el.getAttribute("aria-disabled") !== "true",
    await loc.elementHandle(), { timeout: 20000 }
  );
}

/* ── Navegação inicial ─────────────────────────────────────── */
async function abrirJTe(page) {
  await page.goto("https://jte.csjt.jus.br/start", { waitUntil: "networkidle", timeout: 60000 });
  await page.waitForTimeout(2000);
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
  const trt2 = page.getByText("TRT2 - São Paulo", { exact: true });
  await trt2.waitFor({ state: "visible", timeout: 20000 });
  await trt2.click({ force: true });
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(800);
}

async function abrirModuloPauta(page) {
  const card = page.locator('ion-card-content.card-content-modulo:has-text("Pauta")').first();
  await card.waitFor({ state: "visible", timeout: 20000 });
  await card.click({ force: true });
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(800);
}

/* ── Listar varas ──────────────────────────────────────────── */
async function listarVaras(browser) {
  const page = await novaPage(browser);
  try {
    await abrirJTe(page);
    await abrirModuloPauta(page);
    const botao = page.getByTestId("pautaButtonSelecaoUnidade");
    await botao.waitFor({ state: "visible", timeout: 20000 });
    await botao.click({ force: true });
    await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', { timeout: 20000 });
    await matSel(page, page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select'), "Audiências 1º grau");
    await matSel(page, page.locator('mat-form-field[data-testid="municipio"] mat-select'), "São Paulo - Zonas Central, Norte e Oeste");
    await waitEnabled(page, 'mat-form-field[data-testid="orgao"] mat-select');
    const sel = page.locator('mat-form-field[data-testid="orgao"] mat-select');
    await sel.click({ force: true });
    const panel = page.locator(".mat-mdc-select-panel");
    await panel.waitFor({ state: "visible", timeout: 20000 });
    await page.waitForSelector(".mat-mdc-select-panel mat-option", { timeout: 20000 });
    const opts = panel.locator("mat-option");
    const n = await opts.count();
    const varas = [];
    for (let i = 0; i < n; i++) {
      const l = await opts.nth(i).locator(".mdc-list-item__primary-text").textContent().catch(() => "");
      if (l?.trim()) varas.push(l.trim());
    }
    await page.keyboard.press("Escape").catch(() => {});
    await page.getByTestId("ButtonCancelar").click().catch(() => {});
    return varas;
  } finally {
    await page.context().close().catch(() => {});
  }
}

/* ── Navegação de datas ────────────────────────────────────── */
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
  const am = parseBR(alvo).getTime();
  for (let i = 0; i < 220; i++) {
    const r = await lerData(page);
    if (r?.includes(alvo)) return true;
    const m = String(r ?? "").match(/\b\d{2}\/\d{2}\/\d{4}\b/);
    const cm = m ? parseBR(m[0]).getTime() : 0;
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

/* ── Esperar + Extrair ─────────────────────────────────────── */
const PROC = "ion-item:has(.JT-item-texto-negrito)";

async function esperar(page) {
  for (const s of ["ion-spinner", ".mat-mdc-progress-spinner"]) {
    const sp = page.locator(s).first();
    if (await sp.isVisible({ timeout: 300 }).catch(() => false))
      await sp.waitFor({ state: "hidden", timeout: 15000 }).catch(() => {});
  }
  let last = -1;
  for (let i = 0; i < 20; i++) {
    const c = await page.locator(PROC).count().catch(() => 0);
    if (c === last) {
      await page.waitForTimeout(500);
      if ((await page.locator(PROC).count().catch(() => 0)) === c) return c;
    }
    last = c;
    await page.waitForTimeout(250);
  }
  return last;
}

async function extrair(page) {
  return page.evaluate(() => {
    return Array.from(document.querySelectorAll("ion-item"))
      .map((item) => {
        const numEl = item.querySelector(".JT-item-texto-negrito");
        if (!numEl) return null;
        const get = (sel) => {
          const el = item.querySelector(sel);
          return el ? el.textContent.replace(/ /g, " ").trim() : "";
        };
        const pts = Array.from(item.querySelectorAll(".item-desc-small.item-text-wrap"))
          .map((e) => e.textContent.replace(/ /g, " ").trim()).filter(Boolean);
        return {
          numeroProcesso: numEl.textContent.replace(/ /g, " ").trim(),
          sessao: [get(".sessao"), get(".palavrasRight")].filter(Boolean).join(" - "),
          juiz: pts[0] || "",
          reclamante: (pts[1] || "").replace(/ X$/i, "").trim(),
          reclamada: (pts[2] || "").replace(/\s+/g, " ").trim(),
        };
      })
      .filter(Boolean);
  });
}

const STATUS_EXCLUIR = ["realizada", "cancelada", "adiada", "suspensa", "arquivada", "em andamento"];

function filtrarAtivos(processos) {
  return processos.filter((p) => {
    if (!p.numeroProcesso) return false;
    const st = (p.sessao || "").toLowerCase();
    return !STATUS_EXCLUIR.some((excl) => st.includes(excl));
  });
}

/* ── Verificação polo passivo via 3 pontos → detalhes ──────── */
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
    console.warn(`⚠️  Erro detalhes: ${e.message}`);
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

/* ── Processar 1 job (vara + data) ─────────────────────────── */
async function processarJob(browser, vara, dataBR, pool, geradoEm) {
  if (dataJaPassou(dataBR)) return { n: 0 };

  const page = await novaPage(browser);
  try {
    await abrirJTe(page);
    await abrirModuloPauta(page);

    await fecharOverlays(page);
    const botao = page.getByTestId("pautaButtonSelecaoUnidade");
    await botao.waitFor({ state: "visible", timeout: 20000 });
    await botao.click({ force: true });
    await page.waitForSelector('h1.tituloSelecaoTribunal:has-text("Órgão")', { timeout: 20000 });
    await matSel(page, page.locator('mat-form-field[data-testid="selecaoTribunal"] mat-select'), "Audiências 1º grau");
    await matSel(page, page.locator('mat-form-field[data-testid="municipio"] mat-select'), "São Paulo - Zonas Central, Norte e Oeste");
    await waitEnabled(page, 'mat-form-field[data-testid="orgao"] mat-select');
    await matSel(page, page.locator('mat-form-field[data-testid="orgao"] mat-select'), vara);

    const confirmar = page.getByTestId("ButtonConfirmar");
    await confirmar.waitFor({ state: "visible", timeout: 20000 });
    await confirmar.click({ delay: 80 });
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.waitForTimeout(700);

    if (!(await navData(page, dataBR))) return { n: 0 };
    await page.waitForTimeout(1000);
    if (!(await esperar(page))) return { n: 0 };

    const ps = filtrarAtivos(await extrair(page));
    if (!ps.length) return { n: 0 };

    const leads = [];
    for (let i = 0; i < ps.length; i++) {
      const procItem = page.locator(PROC).nth(i);
      const polos = await verificarDetalhes(page, procItem);

      if (polos && polos.length > 0 && polos.some((p) => !p.temAdvogado)) {
        leads.push(ps[i]);
      }
    }

    if (!leads.length) return { n: 0 };

    if (pool) {
      const vals = leads.map(() => "(?,?,?,?,?,?,?,?,?)").join(",");
      const args = leads.flatMap((p) => [
        new Date(geradoEm), vara, dataBR, brToIso(dataBR),
        p.numeroProcesso, p.sessao || null, p.juiz || null,
        p.reclamante || null, p.reclamada || null,
      ]);
      await pool.query(
        `INSERT INTO processos_sem_polo_passivo
           (geradoEm,vara,dataBR,dataISO,numeroProcesso,sessao,juiz,reclamante,reclamada)
         VALUES ${vals}
         ON DUPLICATE KEY UPDATE
           geradoEm=VALUES(geradoEm), sessao=VALUES(sessao), juiz=VALUES(juiz),
           reclamante=VALUES(reclamante), reclamada=VALUES(reclamada)`,
        args
      ).catch((e) => console.warn(`⚠️  DB: ${e.message}`));
    }

    return { n: leads.length };
  } finally {
    await page.context().close().catch(() => {});
  }
}

/* ── MAIN ──────────────────────────────────────────────────── */
async function main() {
  const geradoEm = new Date().toISOString();
  console.log(
    `\n🚀 Worker JTe | concorrência: ${CONCURRENCY} | meses: ${MESES}${VARAS_LIMIT ? ` | TESTE: ${VARAS_LIMIT} varas` : ""}`
  );

  const pool = await dbPool();
  console.log(pool ? "✅ MySQL conectado" : "⚠️  MySQL desabilitado");

  if (!ONLY_WORK) {
    console.log("\n📦 Listando varas...");
    const bTemp = await chromium.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
    });
    let varas = await listarVaras(bTemp);
    await bTemp.close();

    if (VARAS_LIMIT > 0) {
      varas = varas.slice(0, VARAS_LIMIT);
      console.log(`\n🧪 MODO TESTE: ${varas.length} varas`);
    }

    const datas = gerarDatas();
    console.log(`   ${varas.length} varas × ${datas.length} dias = ${varas.length * datas.length} jobs`);
    if (datas.length) console.log(`   ${datas[0]} → ${datas[datas.length - 1]}\n`);

    const fila = new Queue(QUEUE_NAME, {
      connection: REDIS,
      defaultJobOptions: {
        attempts: 3,
        backoff: { type: "exponential", delay: 5000 },
        removeOnComplete: { count: 50 },
        removeOnFail: { count: 200 },
      },
    });
    await fila.obliterate({ force: true }).catch(() => {});

    let n = 0;
    for (const vara of varas) {
      const jobs = datas.map((data) => ({
        name: "s",
        data: { vara, data, geradoEm },
        opts: { jobId: `${vara.replace(/\s+/g, "_")}_${data.replace(/\//g, "-")}` },
      }));
      for (let i = 0; i < jobs.length; i += 500) {
        await fila.addBulk(jobs.slice(i, i + 500));
        n += Math.min(500, jobs.length - i);
        process.stdout.write(`\r   ${n}/${varas.length * datas.length}`);
      }
    }
    console.log(`\n✅ ${n} jobs enfileirados\n`);
    await fila.close();
    if (ONLY_QUEUE) { if (pool) await pool.end(); return; }
  }

  console.log(`⚙️  Processando com ${CONCURRENCY} browsers...\n`);
  const browsers = await Promise.all(
    Array.from({ length: CONCURRENCY }, () =>
      chromium.launch({
        headless: true,
        args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
      })
    )
  );

  let done = 0, erros = 0, procs = 0;
  const t0 = Date.now();

  const filaInfo = new Queue(QUEUE_NAME, { connection: REDIS });
  const counts = await filaInfo.getJobCounts();
  const total = counts.waiting + counts.active + counts.delayed;
  await filaInfo.close();

  const worker = new Worker(QUEUE_NAME, async (job) => {
    const idx = done % CONCURRENCY;
    const browser = browsers[idx];
    try {
      const r = await processarJob(browser, job.data.vara, job.data.data, pool, job.data.geradoEm);
      done++; procs += r.n || 0;
      const pct = ((done / total) * 100).toFixed(1);
      const eta = done > 0 ? Math.round(((Date.now() - t0) / done) * (total - done) / 1000) : "?";
      process.stdout.write(`\r  [${pct}%] ${done}/${total} | ${procs} leads | ETA: ${eta}s   `);
    } catch (e) { erros++; done++; throw e; }
  }, { connection: REDIS, concurrency: CONCURRENCY });

  await new Promise((resolve) => {
    let checking = false;
    const ck = setInterval(async () => {
      if (checking) return;
      checking = true;
      try {
        const q = new Queue(QUEUE_NAME, { connection: REDIS });
        const c = await q.getJobCounts(); await q.close();
        if (c.waiting === 0 && c.active === 0 && c.delayed === 0) { clearInterval(ck); resolve(); }
      } catch {} finally { checking = false; }
    }, 5000);
    worker.on("error", (e) => console.error("\n❌", e.message));
  });

  console.log(`\n\n✅ ${((Date.now() - t0) / 1000).toFixed(0)}s | ${procs} leads qualificados | ${erros} erros`);
  await worker.close();
  await Promise.all(browsers.map((b) => b.close().catch(() => {})));
  if (pool) await pool.end();
}

process.on("SIGINT", () => {
  console.log("\n⚠️  Interrompido (SIGINT)");
  process.exit(130);
});

main().catch((e) => {
  console.error("\n❌", e);
  process.exit(1);
});
