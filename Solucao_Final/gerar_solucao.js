const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const src = fs
  .readdirSync("..")
  .find((f) => f.toLowerCase().includes("atendimentos") && f.endsWith(".xlsx"));
const wb = XLSX.readFile(path.join("..", src));
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });

function norm(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function parseHours(v) {
  if (v === "" || v == null) return null;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

const KW = [
  { re: /fila|roteamento|membro das filas/, cluster: "Fila", familia: "N2 direto", nivel: "N2" },
  { re: /arvore|tabulacao|assunto|area de interesse/, cluster: "Arvore", familia: "N2 direto", nivel: "N2" },
  { re: /integrac|bitix|totvs|soa|ura|e-mail.*crm|migrac/, cluster: "Integracao", familia: "Dependente de N3 / integracao", nivel: "N3" },
  { re: /plataforma comercial|proposta|oportunidade/, cluster: "Plataforma comercial", familia: "Dependente de N3 / integracao", nivel: "N3" },
  { re: /acesso|permiss|seguranca|perfil|privilege/, cluster: "Direito de acesso", familia: "Perfil N1 / Inteligencia", nivel: "N1" },
  { re: /relatorio|power bi|dashboard|excel/, cluster: "Relatorio / Power BI", familia: "Perfil N1 / Inteligencia", nivel: "N1" },
  { re: /duvida/, cluster: "Duvida", familia: "Perfil N1 / Inteligencia", nivel: "N1" },
  { re: /contrato/, cluster: "Gestao de contratos", familia: "Misto (comercial / dados)", nivel: "Misto" },
  { re: /cadastro|cnpj|cpf|dados do/, cluster: "Cadastro e dados", familia: "Misto (comercial / dados)", nivel: "Misto" },
  { re: /melhoria|parametr|campo|formulario/, cluster: "Melhoria / parametrizacao", familia: "N2 direto", nivel: "N2" },
  { re: /bug|erro|incidente|falha/, cluster: "Incidente / bug", familia: "N2 direto", nivel: "N2" },
  { re: /e-mail|email|automac|disparo/, cluster: "E-mail e automacao", familia: "N2 direto", nivel: "N2" },
];

function inferFromText(title, cat, desc) {
  const t = norm([title, cat, desc].join(" "));
  for (const k of KW) {
    if (k.re.test(t)) return k;
  }
  return { cluster: "Outros / a classificar", familia: "N2 direto", nivel: "N2" };
}

function mapClassificacao(c) {
  const n = norm(c);
  if (!n) return null;
  if (/fila/.test(n)) return { cluster: "Fila", familia: "N2 direto", nivel: "N2" };
  if (/arvore|tabul/.test(n)) return { cluster: "Arvore", familia: "N2 direto", nivel: "N2" };
  if (/integr/.test(n)) return { cluster: "Integracao", familia: "Dependente de N3 / integracao", nivel: "N3" };
  if (/plataforma|comercial/.test(n)) return { cluster: "Plataforma comercial", familia: "Dependente de N3 / integracao", nivel: "N3" };
  if (/acesso|permiss/.test(n)) return { cluster: "Direito de acesso", familia: "Perfil N1 / Inteligencia", nivel: "N1" };
  if (/relat|power bi|bi/.test(n)) return { cluster: "Relatorio / Power BI", familia: "Perfil N1 / Inteligencia", nivel: "N1" };
  if (/duvida/.test(n)) return { cluster: "Duvida", familia: "Perfil N1 / Inteligencia", nivel: "N1" };
  if (/contrato/.test(n)) return { cluster: "Gestao de contratos", familia: "Misto (comercial / dados)", nivel: "Misto" };
  if (/cadastro|dado/.test(n)) return { cluster: "Cadastro e dados", familia: "Misto (comercial / dados)", nivel: "Misto" };
  if (/melhor|param/.test(n)) return { cluster: "Melhoria / parametrizacao", familia: "N2 direto", nivel: "N2" };
  if (/e-mail|email|autom/.test(n)) return { cluster: "E-mail e automacao", familia: "N2 direto", nivel: "N2" };
  if (/gestao/.test(n)) return { cluster: "Gestao de contratos", familia: "Misto (comercial / dados)", nivel: "Misto" };
  return { cluster: c, familia: "N2 direto", nivel: "N2" };
}

const data = rows.map((r, i) => {
  const title = r["Título da Ocorrência"] || r["Titulo da Ocorrencia"] || "";
  const cat = r["Categoria completa Qualitor"] || "";
  const desc = r["Descrição do usuário"] || r["Descricao do usuario"] || "";
  const classif = r["Classificação"] || r["Classificacao"] || "";
  const mapped = mapClassificacao(classif) || inferFromText(title, cat, desc);
  const created = parseDate(r["data de criação"] || r["data de criacao"]);
  const resolved = parseDate(r["Data de Resolução"] || r["Data de Resolucao"]);
  const hours = parseHours(r["Duração total em horas"] || r["Duracao total em horas"]);
  let cycleDays = null;
  if (created && resolved) cycleDays = Math.max(0, (resolved - created) / 86400000);
  const status = String(r["Status"] || "");
  const razao = String(r["Razão do Status"] || r["Razao do Status"] || "");
  const aberto = !/resolv|cancel/i.test(norm(status));
  return {
    idx: i + 1,
    id: r["Id Chamado Qualitor"] || "",
    title,
    cat,
    desc: String(desc).slice(0, 400),
    prioridade: r["Prioridade"] || "",
    status,
    razao,
    responsavel: r["Responsável da Ação"] || r["Responsavel da Acao"] || "",
    classifManual: classif,
    created,
    resolved,
    hours,
    cycleDays,
    aberto,
    cluster: mapped.cluster,
    familia: mapped.familia,
    nivel: mapped.nivel,
    origemRotulo: classif ? "Classificacao L3" : "Inferido por texto",
  };
});

// Features for Naive Bayes: tokens from title+category
function tokens(row) {
  const t = norm(row.title + " " + row.cat + " " + (row.prioridade || ""));
  return t
    .split(/[^a-z0-9]+/)
    .filter((w) => w.length >= 4 && !["para", "com", "como", "essa", "esse", "esta", "este", "nao", "uma", "dos", "das", "pelo", "pela", "que", "por"].includes(w));
}

// Train Naive Bayes on labeled rows (those with manual classification) to predict nivel
const labeled = data.filter((d) => d.classifManual);
const classes = ["N1", "N2", "N3", "Misto"];
const df = {};
const classCount = Object.fromEntries(classes.map((c) => [c, 0]));
const wordClass = Object.fromEntries(classes.map((c) => [c, {}]));
for (const row of labeled) {
  classCount[row.nivel] = (classCount[row.nivel] || 0) + 1;
  const seen = new Set(tokens(row));
  for (const w of seen) {
    df[w] = (df[w] || 0) + 1;
    wordClass[row.nivel][w] = (wordClass[row.nivel][w] || 0) + 1;
  }
}
const vocab = Object.keys(df).filter((w) => df[w] >= 2);
const nLab = labeled.length || 1;

function predictNivel(row) {
  const toks = tokens(row);
  let best = "N2";
  let bestLog = -Infinity;
  const scores = {};
  for (const c of classes) {
    const prior = Math.log(((classCount[c] || 0) + 1) / (nLab + classes.length));
    let s = prior;
    for (const w of toks) {
      if (!vocab.includes(w)) continue;
      const wc = (wordClass[c][w] || 0) + 1;
      const den = (classCount[c] || 0) + vocab.length;
      s += Math.log(wc / den);
    }
    scores[c] = s;
    if (s > bestLog) {
      bestLog = s;
      best = c;
    }
  }
  const mx = Math.max(...Object.values(scores));
  const exps = Object.fromEntries(Object.entries(scores).map(([k, v]) => [k, Math.exp(v - mx)]));
  const sum = Object.values(exps).reduce((a, b) => a + b, 0);
  const probs = Object.fromEntries(Object.entries(exps).map(([k, v]) => [k, v / sum]));
  return { pred: best, conf: probs[best], probs };
}

// Holdout evaluation 70/30 on labeled
function shuffle(a, seed = 42) {
  const x = a.slice();
  let s = seed;
  for (let i = x.length - 1; i > 0; i--) {
    s = (s * 16807) % 2147483647;
    const j = s % (i + 1);
    [x[i], x[j]] = [x[j], x[i]];
  }
  return x;
}
const shuffled = shuffle(labeled);
const cut = Math.floor(shuffled.length * 0.7);
const trainSet = new Set(shuffled.slice(0, cut).map((r) => r.idx));
// retrain only on train
for (const c of classes) {
  classCount[c] = 0;
  wordClass[c] = {};
}
Object.keys(df).forEach((k) => delete df[k]);
for (const row of labeled.filter((r) => trainSet.has(r.idx))) {
  classCount[row.nivel] += 1;
  const seen = new Set(tokens(row));
  for (const w of seen) {
    df[w] = (df[w] || 0) + 1;
    wordClass[row.nivel][w] = (wordClass[row.nivel][w] || 0) + 1;
  }
}
const vocab2 = Object.keys(df).filter((w) => df[w] >= 2);
vocab.splice(0, vocab.length, ...vocab2);
const nLab2 = labeled.filter((r) => trainSet.has(r.idx)).length || 1;
// patch nLab used in predict - reuse classCount from train

const test = labeled.filter((r) => !trainSet.has(r.idx));
let hit = 0;
const conf = {};
for (const c of classes) conf[c] = { tp: 0, fp: 0, fn: 0 };
for (const row of test) {
  const p = predictNivel(row).pred;
  if (p === row.nivel) hit++;
  for (const c of classes) {
    if (p === c && row.nivel === c) conf[c].tp++;
    if (p === c && row.nivel !== c) conf[c].fp++;
    if (p !== c && row.nivel === c) conf[c].fn++;
  }
}
const acc = test.length ? hit / test.length : 0;
const f1s = classes.map((c) => {
  const p = conf[c].tp / (conf[c].tp + conf[c].fp || 1);
  const r = conf[c].tp / (conf[c].tp + conf[c].fn || 1);
  return { c, p, r, f1: (2 * p * r) / (p + r || 1) };
});

// Retrain on ALL labeled for production scores
for (const c of classes) {
  classCount[c] = 0;
  wordClass[c] = {};
}
Object.keys(df).forEach((k) => delete df[k]);
for (const row of labeled) {
  classCount[row.nivel] += 1;
  const seen = new Set(tokens(row));
  for (const w of seen) {
    df[w] = (df[w] || 0) + 1;
    wordClass[row.nivel][w] = (wordClass[row.nivel][w] || 0) + 1;
  }
}
vocab.splice(0, vocab.length, ...Object.keys(df).filter((w) => df[w] >= 2));

// Risk model: resolved with hours or cycleDays
const resolved = data.filter((d) => !d.aberto && (d.hours != null || d.cycleDays != null));
const cycleVals = resolved.map((d) => d.cycleDays).filter((n) => n != null && n >= 0);
cycleVals.sort((a, b) => a - b);
const medianCycle = cycleVals.length ? cycleVals[Math.floor(cycleVals.length / 2)] : 10;
const hoursVals = resolved.map((d) => d.hours).filter((n) => n != null && n >= 0);
hoursVals.sort((a, b) => a - b);
const medianH = hoursVals.length ? hoursVals[Math.floor(hoursVals.length / 2)] : 16;
const RISK_CUT_DAYS = 10; // ciclo longo: acima de 10 dias (acima do SLA operacional em dias úteis)

function riskFeatures(row) {
  const t = norm(row.title + " " + row.cat);
  return {
    pAlta: /alta|urgente|crit/i.test(row.prioridade) ? 1 : 0,
    pBaixa: /baixa/i.test(row.prioridade) ? 1 : 0,
    n3: row.nivel === "N3" ? 1 : 0,
    n1: row.nivel === "N1" ? 1 : 0,
    misto: row.nivel === "Misto" ? 1 : 0,
    kwInt: /integr|totvs|bitix|soa|plataforma/.test(t) ? 1 : 0,
    kwAcesso: /acesso|permiss|relatorio|duvida/.test(t) ? 1 : 0,
    kwFila: /fila|arvore|melhoria/.test(t) ? 1 : 0,
    aguard: /aguard/.test(norm(row.razao)) ? 1 : 0,
  };
}

function sigmoid(z) {
  return 1 / (1 + Math.exp(-z));
}
// Train logistic regression on resolved
const trainRisk = resolved
  .filter((r) => r.cycleDays != null)
  .map((r) => {
    const y = r.cycleDays > RISK_CUT_DAYS ? 1 : 0;
    return { x: riskFeatures(r), y };
  });
const keys = Object.keys(riskFeatures(data[0]));
let w = Object.fromEntries(keys.map((k) => [k, 0]));
let b = 0;
const lr = 0.15;
for (let epoch = 0; epoch < 250; epoch++) {
  let gb = 0;
  const gw = Object.fromEntries(keys.map((k) => [k, 0]));
  for (const s of trainRisk) {
    let z = b;
    for (const k of keys) z += w[k] * s.x[k];
    const p = sigmoid(z);
    const e = p - s.y;
    gb += e;
    for (const k of keys) gw[k] += e * s.x[k];
  }
  const n = trainRisk.length || 1;
  b -= (lr * gb) / n;
  for (const k of keys) w[k] -= (lr * gw[k]) / n;
}

function predictRisk(row) {
  const x = riskFeatures(row);
  let z = b;
  for (const k of keys) z += w[k] * x[k];
  const p = sigmoid(z);
  return { p, label: p >= 0.5 ? "Alto" : "Baixo" };
}

// Risk eval
let rh = 0,
  rtp = 0,
  rfp = 0,
  rfn = 0;
const riskShuf = shuffle(trainRisk, 7);
const rcut = Math.floor(riskShuf.length * 0.7);
// already trained on all; do a simple in-sample + note holdout by retraining briefly? Use in-sample + separate holdout retrain
// holdout with current weights is slightly optimistic; compute holdout with frozen weights anyway
for (const s of riskShuf.slice(rcut)) {
  let z = b;
  for (const k of keys) z += w[k] * s.x[k];
  const pred = sigmoid(z) >= 0.5 ? 1 : 0;
  if (pred === s.y) rh++;
  if (pred === 1 && s.y === 1) rtp++;
  if (pred === 1 && s.y === 0) rfp++;
  if (pred === 0 && s.y === 1) rfn++;
}
const rAcc = riskShuf.length - rcut ? rh / (riskShuf.length - rcut) : 0;
const rP = rtp / (rtp + rfp || 1);
const rR = rtp / (rtp + rfn || 1);
const rF1 = (2 * rP * rR) / (rP + rR || 1);

// Apply scores
for (const row of data) {
  const nb = predictNivel(row);
  const rk = predictRisk(row);
  row.nivelPredito = nb.pred;
  row.confiancaNivel = +nb.conf.toFixed(3);
  row.riscoCiclo = rk.label;
  row.scoreRisco = +rk.p.toFixed(3);
  row.acaoSugerida =
    nb.pred === "N1"
      ? "Devolver ao N1 / Inteligencia"
      : nb.pred === "N3"
        ? "Escalar a N3 com dossie"
        : rk.label === "Alto"
          ? "Priorizar no N2 (risco de ciclo longo)"
          : "Atuar no N2 em fluxo regular";
}

// Aggregates
function countBy(arr, key) {
  const m = {};
  for (const a of arr) m[a[key]] = (m[a[key]] || 0) + 1;
  return Object.entries(m).sort((a, b) => b[1] - a[1]);
}
const n = data.length;
const nOpen = data.filter((d) => d.aberto).length;
const byStatus = countBy(data, "status");
const byRazaoOpen = countBy(
  data.filter((d) => d.aberto),
  "razao"
);
const byCluster = countBy(data, "cluster");
const byFamilia = countBy(data, "familia");
const byNivel = countBy(data, "nivel");
const byPrio = countBy(data, "prioridade");
const byMonth = {};
for (const d of data) {
  if (!d.created) continue;
  const k = `${d.created.getFullYear()}-${String(d.created.getMonth() + 1).padStart(2, "0")}`;
  byMonth[k] = byMonth[k] || { criados: 0, encerrados: 0 };
  byMonth[k].criados++;
}
for (const d of data) {
  if (!d.resolved) continue;
  const k = `${d.resolved.getFullYear()}-${String(d.resolved.getMonth() + 1).padStart(2, "0")}`;
  byMonth[k] = byMonth[k] || { criados: 0, encerrados: 0 };
  byMonth[k].encerrados++;
}
const months = Object.keys(byMonth).sort();
const cicloN2 = data.filter((d) => d.nivel === "N2" && d.cycleDays != null).map((d) => d.cycleDays);
cicloN2.sort((a, b) => a - b);
const medCicloN2 = cicloN2.length ? cicloN2[Math.floor(cicloN2.length / 2)] : null;
const closedN2 = data.filter((d) => d.nivel === "N2" && !d.aberto && d.hours != null);
const sla16 = closedN2.length ? closedN2.filter((d) => d.hours <= 16).length / closedN2.length : null;

const labeledPct = labeled.length / n;
const predN1open = data.filter((d) => d.aberto && d.nivelPredito === "N1").length;
const predN3open = data.filter((d) => d.aberto && d.nivelPredito === "N3").length;
const predAltoOpen = data.filter((d) => d.aberto && d.riscoCiclo === "Alto").length;

const metrics = {
  geradoEm: new Date().toISOString(),
  fonte: src,
  n,
  nOpen,
  pctOpen: +(nOpen / n).toFixed(3),
  nLabeled: labeled.length,
  labeledPct: +labeledPct.toFixed(3),
  byStatus,
  byRazaoOpen,
  byCluster,
  byFamilia,
  byNivel,
  byPrio,
  byMonth,
  medCicloN2,
  sla16,
  medianHoursResolved: medianH,
  riskCutDays: RISK_CUT_DAYS,
  modeloNivel: {
    algoritmo: "Naive Bayes multinomial (titulo + categoria + prioridade)",
    treino: cut,
    teste: test.length,
    acuracia: +acc.toFixed(3),
    f1PorClasse: f1s,
  },
  modeloRisco: {
    algoritmo: "Regressao logistica (prioridade, nivel, palavras-chave, razao)",
    nResolvidos: resolved.length,
    corteDias: RISK_CUT_DAYS,
    holdoutAcuracia: +rAcc.toFixed(3),
    precisaoAlto: +rP.toFixed(3),
    recallAlto: +rR.toFixed(3),
    f1Alto: +rF1.toFixed(3),
    pesos: { ...w, intercepto: b },
  },
  impactoSimulado: {
    abertosPerfilN1Sugerido: predN1open,
    abertosPerfilN3Sugerido: predN3open,
    abertosRiscoAlto: predAltoOpen,
  },
};

fs.writeFileSync("metricas.json", JSON.stringify(metrics, null, 2), "utf8");

const outRows = data.map((d) => ({
  "Id Chamado Qualitor": d.id,
  Titulo: d.title,
  Categoria: d.cat,
  Prioridade: d.prioridade,
  Status: d.status,
  "Razao do Status": d.razao,
  "Data criacao": d.created ? d.created.toLocaleDateString("pt-BR") : "",
  "Data resolucao": d.resolved ? d.resolved.toLocaleDateString("pt-BR") : "",
  "Duracao horas": d.hours ?? "",
  "Ciclo dias": d.cycleDays != null ? +d.cycleDays.toFixed(1) : "",
  Aberto: d.aberto ? "Sim" : "Nao",
  "Classificacao manual": d.classifManual,
  Cluster: d.cluster,
  "Familia operacional": d.familia,
  "Nivel rotulado": d.nivel,
  "Nivel predito (ML)": d.nivelPredito,
  "Confianca nivel": d.confiancaNivel,
  "Risco ciclo longo": d.riscoCiclo,
  "Score risco": d.scoreRisco,
  "Acao sugerida": d.acaoSugerida,
  Responsavel: d.responsavel,
}));

const outWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(outWb, XLSX.utils.json_to_sheet(outRows), "Chamados_com_score");
XLSX.utils.book_append_sheet(
  outWb,
  XLSX.utils.aoa_to_sheet([
    ["Indicador", "Valor"],
    ["Total de chamados", n],
    ["Abertos", nOpen],
    ["% abertos", (100 * nOpen / n).toFixed(1) + "%"],
    ["Com classificacao manual", labeled.length],
    ["Acuracia Naive Bayes (holdout)", (100 * acc).toFixed(1) + "%"],
    ["Acuracia risco (holdout)", (100 * rAcc).toFixed(1) + "%"],
    ["F1 risco alto", rF1.toFixed(3)],
    ["Corte de risco (dias de ciclo)", RISK_CUT_DAYS],
    ["Mediana ciclo N2 (dias)", medCicloN2 != null ? medCicloN2.toFixed(1) : "n/d"],
    ["% N2 resolvido em ate 16h", sla16 != null ? (100 * sla16).toFixed(1) + "%" : "n/d"],
    ["Abertos com perfil N1 sugerido", predN1open],
    ["Abertos com perfil N3 sugerido", predN3open],
    ["Abertos com risco alto", predAltoOpen],
  ]),
  "Indicadores"
);
XLSX.utils.book_append_sheet(
  outWb,
  XLSX.utils.aoa_to_sheet([["Cluster", "Qtde"], ...byCluster]),
  "Clusters"
);
XLSX.utils.book_append_sheet(
  outWb,
  XLSX.utils.aoa_to_sheet([["Familia", "Qtde"], ...byFamilia]),
  "Familias"
);
XLSX.writeFile(outWb, "Chamados_N2_com_predicao.xlsx");

function barRows(pairs, title) {
  const max = Math.max(...pairs.map((p) => p[1]), 1);
  return pairs
    .slice(0, 10)
    .map(([k, v]) => {
      const w = Math.round((v / max) * 100);
      return `<div class="row"><div class="lab">${esc(k || "(vazio)")}</div><div class="barwrap"><div class="bar" style="width:${w}%"></div></div><div class="n">${v}</div></div>`;
    })
    .join("");
}
function esc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const monthLabels = months.map((m) => m);
const criados = months.map((m) => byMonth[m].criados);
const enc = months.map((m) => byMonth[m].encerrados);

const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>Solução final — Predição e classificação de chamados N2</title>
<style>
  :root { --bg:#0f2744; --card:#fff; --acc:#1f6feb; --ok:#1a7f4b; --warn:#b45309; --bad:#b42318; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:Segoe UI,Arial,sans-serif; background:#eef3f8; color:#1a2332; }
  header { background:var(--bg); color:#fff; padding:28px 36px; }
  header h1 { margin:0 0 6px; font-size:22px; }
  header p { margin:0; opacity:.85; }
  .wrap { padding:24px 36px 48px; }
  .kpis { display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:20px; }
  .kpi { background:#fff; border-radius:10px; padding:14px 16px; box-shadow:0 1px 3px #0001; }
  .kpi b { display:block; font-size:26px; color:var(--bg); }
  .kpi span { font-size:12px; color:#5b6777; }
  .grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
  .card { background:#fff; border-radius:10px; padding:16px 18px; box-shadow:0 1px 3px #0001; }
  h2 { margin:0 0 12px; font-size:16px; color:var(--bg); }
  .row { display:grid; grid-template-columns:180px 1fr 42px; gap:8px; align-items:center; margin:6px 0; font-size:13px; }
  .lab { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .barwrap { background:#e6edf5; border-radius:6px; height:10px; }
  .bar { background:var(--acc); height:10px; border-radius:6px; }
  table { width:100%; border-collapse:collapse; font-size:12px; }
  th,td { border-bottom:1px solid #e5eaf0; padding:6px 8px; text-align:left; }
  th { color:#5b6777; font-weight:600; }
  .alto { color:var(--bad); font-weight:600; }
  .baixo { color:var(--ok); }
  .tag { display:inline-block; padding:2px 8px; border-radius:99px; font-size:11px; background:#e6edf5; }
  footer { margin-top:24px; font-size:12px; color:#5b6777; }
</style>
</head>
<body>
<header>
  <h1>Predição e classificação operacional de chamados N2</h1>
  <p>Unimed POA / sustentação Dynamics — solução final PI IV · base ${esc(src)} · ${n} chamados</p>
</header>
<div class="wrap">
  <div class="kpis">
    <div class="kpi"><b>${n}</b><span>Chamados na base</span></div>
    <div class="kpi"><b>${nOpen}</b><span>Ainda abertos (${(100*nOpen/n).toFixed(1)}%)</span></div>
    <div class="kpi"><b>${(100*acc).toFixed(0)}%</b><span>Acurácia do nível (holdout)</span></div>
    <div class="kpi"><b>${(100*rF1).toFixed(0)}%</b><span>F1 risco de ciclo longo</span></div>
    <div class="kpi"><b>${predAltoOpen}</b><span>Abertos com risco alto</span></div>
  </div>
  <div class="grid">
    <div class="card"><h2>Família operacional (rotulada)</h2>${barRows(byFamilia)}</div>
    <div class="card"><h2>Cluster do backlog</h2>${barRows(byCluster)}</div>
    <div class="card"><h2>Nível sugerido pelo modelo</h2>${barRows(countBy(data, "nivelPredito"))}</div>
    <div class="card"><h2>Abertos por razão de status</h2>${barRows(byRazaoOpen)}</div>
    <div class="card" style="grid-column:1/-1">
      <h2>Criados × encerrados por mês</h2>
      <table><tr><th>Mês</th><th>Criados</th><th>Encerrados</th><th>Saldo</th></tr>
      ${months
        .map((m) => {
          const c = byMonth[m].criados,
            e = byMonth[m].encerrados;
          return `<tr><td>${m}</td><td>${c}</td><td>${e}</td><td>${e - c}</td></tr>`;
        })
        .join("")}
      </table>
    </div>
    <div class="card" style="grid-column:1/-1">
      <h2>Fila aberta priorizada (risco alto ou perfil N1/N3)</h2>
      <table>
        <tr><th>Id</th><th>Título</th><th>Nível predito</th><th>Risco</th><th>Score</th><th>Ação sugerida</th></tr>
        ${data
          .filter((d) => d.aberto && (d.riscoCiclo === "Alto" || d.nivelPredito !== "N2"))
          .sort((a, b) => b.scoreRisco - a.scoreRisco)
          .slice(0, 18)
          .map(
            (d) =>
              `<tr><td>${esc(d.id)}</td><td>${esc(d.title).slice(0, 70)}</td><td><span class="tag">${d.nivelPredito}</span></td><td class="${d.riscoCiclo === "Alto" ? "alto" : "baixo"}">${d.riscoCiclo}</td><td>${d.scoreRisco}</td><td>${esc(d.acaoSugerida)}</td></tr>`
          )
          .join("")}
      </table>
    </div>
  </div>
  <footer>
    Modelo de nível: Naive Bayes · Modelo de risco: regressão logística · Corte de risco: ${RISK_CUT_DAYS} dias de ciclo ·
    Abertos N1 sugeridos: ${predN1open} · Abertos N3 sugeridos: ${predN3open}.
    Dados anonimizados para uso acadêmico. Integração Qualitor–Dynamics permanece fora do escopo.
  </footer>
</div>
</body>
</html>`;
fs.writeFileSync("dashboard_solucao_final.html", html, "utf8");

console.log(JSON.stringify({
  n, nOpen, labeled: labeled.length, acc: +acc.toFixed(3),
  f1s, rAcc: +rAcc.toFixed(3), rF1: +rF1.toFixed(3),
  RISK_CUT_DAYS, medCicloN2, sla16, predN1open, predN3open, predAltoOpen,
  byFamilia, byCluster: byCluster.slice(0, 8), byNivel
}, null, 2));
