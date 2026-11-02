const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const wb = XLSX.readFile(path.join(__dirname, "Chamados_N2_com_predicao.xlsx"));
const rows = XLSX.utils.sheet_to_json(wb.Sheets["Chamados_com_score"], { defval: "" });

function esc(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
function countBy(arr, key) {
  const m = {};
  for (const a of arr) {
    const k = a[key] || "(vazio)";
    m[k] = (m[k] || 0) + 1;
  }
  return Object.entries(m).sort((a, b) => b[1] - a[1]);
}
function bars(pairs, color) {
  const max = Math.max(...pairs.map((p) => p[1]), 1);
  return pairs
    .slice(0, 8)
    .map(([k, v]) => {
      const w = Math.round((v / max) * 100);
      return `<div class="barrow"><span class="blab" title="${esc(k)}">${esc(k)}</span><div class="btrack"><div class="bfill" style="width:${w}%;background:${color}"></div></div><span class="bn">${v}</span></div>`;
    })
    .join("");
}

const n = rows.length;
const open = rows.filter((r) => r.Aberto === "Sim");
const nOpen = open.length;
const nAlto = open.filter((r) => r["Risco ciclo longo"] === "Alto").length;
const nN1 = open.filter((r) => r["Nivel predito (ML)"] === "N1").length;
const nN3 = open.filter((r) => r["Nivel predito (ML)"] === "N3").length;
const byFam = countBy(rows, "Familia operacional");
const byClu = countBy(rows, "Cluster");
const byNivel = countBy(rows, "Nivel predito (ML)");
const byRisco = countBy(rows, "Risco ciclo longo");
const byRazao = countBy(open, "Razao do Status");
const byPrio = countBy(rows, "Prioridade");

const months = {};
for (const r of rows) {
  const c = String(r["Data criacao"] || "");
  const e = String(r["Data resolucao"] || "");
  const mc = c.match(/(\d{2})\/(\d{4})/);
  const me = e.match(/(\d{2})\/(\d{4})/);
  if (mc) {
    const k = `${mc[2]}-${mc[1]}`;
    months[k] = months[k] || { c: 0, e: 0 };
    months[k].c++;
  }
  if (me) {
    const k = `${me[2]}-${me[1]}`;
    months[k] = months[k] || { c: 0, e: 0 };
    months[k].e++;
  }
}
const mk = Object.keys(months).filter((k) => k.startsWith("2026")).sort();
const maxM = Math.max(...mk.map((k) => Math.max(months[k].c, months[k].e)), 1);

const monthSvg = (() => {
  const W = 860, H = 210, pad = 36;
  const bw = (W - pad * 2) / mk.length / 2.4;
  let g = "";
  mk.forEach((k, i) => {
    const x = pad + i * ((W - pad * 2) / mk.length);
    const hc = ((months[k].c / maxM) * (H - 50));
    const he = ((months[k].e / maxM) * (H - 50));
    g += `<rect x="${x}" y="${H - 24 - hc}" width="${bw}" height="${hc}" fill="#00ADEF"/>`;
    g += `<rect x="${x + bw + 3}" y="${H - 24 - he}" width="${bw}" height="${he}" fill="#00995D"/>`;
    g += `<text x="${x + bw}" y="${H - 8}" font-size="10" fill="#263339" text-anchor="middle">${k.slice(5)}</text>`;
  });
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" height="210">${g}
    <text x="36" y="14" font-size="11" fill="#5b6777">2026 · azul = criados · verde = encerrados</text></svg>`;
})();

const fila = open
  .slice()
  .sort((a, b) => Number(b["Score risco"] || 0) - Number(a["Score risco"] || 0))
  .slice(0, 12);

const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>Relatório Power BI — Chamados N2</title>
<style>
  :root {
    --az:#00ADEF; --vd:#00995D; --vd2:#005C40; --am:#FEBE00;
    --tx:#263339; --bg:#F4F7F6; --card:#fff; --bad:#CC3366;
  }
  * { box-sizing:border-box; }
  html,body { margin:0; height:100%; font-family:"Segoe UI",Arial,sans-serif; background:var(--bg); color:var(--tx); }
  .pbi-top {
    height:40px; background:#F3F2F1; border-bottom:1px solid #E1DFDD;
    display:flex; align-items:center; gap:16px; padding:0 16px; font-size:12px; color:#605E5C;
  }
  .pbi-top b { color:var(--vd2); font-size:13px; }
  .tabs { display:flex; gap:0; background:#fff; border-bottom:1px solid #E1DFDD; padding:0 16px; }
  .tab { padding:10px 18px; cursor:pointer; font-size:13px; border-bottom:3px solid transparent; color:#605E5C; }
  .tab.on { color:var(--vd2); border-bottom-color:var(--vd); font-weight:600; }
  .page { display:none; padding:14px 18px 18px; }
  .page.on { display:block; }
  .kpis { display:grid; grid-template-columns:repeat(5,1fr); gap:10px; margin-bottom:12px; }
  .kpi { background:#fff; border-radius:4px; padding:10px 12px; box-shadow:0 1px 3px #0001; border-top:3px solid var(--vd); }
  .kpi.alt { border-top-color:var(--az); }
  .kpi.warn { border-top-color:var(--bad); }
  .kpi span { font-size:11px; color:#605E5C; }
  .kpi b { display:block; font-size:26px; line-height:1.15; color:var(--vd2); }
  .grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
  .grid3 { display:grid; grid-template-columns:1.2fr 1fr 1fr; gap:10px; }
  .card { background:#fff; border-radius:4px; padding:12px 14px; box-shadow:0 1px 3px #0001; min-height:220px; }
  .card h3 { margin:0 0 10px; font-size:12px; color:var(--vd2); font-weight:600; }
  .barrow { display:grid; grid-template-columns:150px 1fr 36px; gap:8px; align-items:center; margin:5px 0; font-size:12px; }
  .blab { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .btrack { background:#E6EEE9; height:8px; border-radius:6px; }
  .bfill { height:8px; border-radius:6px; }
  .bn { text-align:right; color:#605E5C; }
  table { width:100%; border-collapse:collapse; font-size:11px; }
  th { text-align:left; color:#605E5C; font-weight:600; padding:5px 6px; border-bottom:1px solid #E1DFDD; }
  td { padding:5px 6px; border-bottom:1px solid #F3F2F1; }
  .alto { color:var(--bad); font-weight:600; }
  .baixo { color:var(--vd); }
  .pill { display:inline-block; background:#E6F4EC; color:var(--vd2); border-radius:99px; padding:1px 8px; font-size:10px; }
  .filters { display:flex; gap:8px; margin-bottom:12px; font-size:12px; color:#605E5C; }
  .chip { background:#fff; border:1px solid #E1DFDD; border-radius:12px; padding:4px 10px; }
  footer.note { margin-top:10px; font-size:11px; color:#605E5C; }
</style>
</head>
<body>
<div class="pbi-top">
  <b>Power BI</b>
  <span>Relatorio de Chamados.pbix</span>
  <span>·</span>
  <span>Instituição analisada / N2 Dynamics</span>
  <span style="margin-left:auto">Filtros do relatório · Base 26/10/2026 · ${n} chamados</span>
</div>
<div class="tabs">
  <div class="tab on" data-p="op">Visão Operacional</div>
  <div class="tab" data-p="pred">Predição e classificação</div>
</div>

<section id="op" class="page on">
  <div class="filters">
    <span class="chip">Status: todos</span>
    <span class="chip">Ano: 2026</span>
    <span class="chip">Origem: Qualitor + score ML</span>
  </div>
  <div class="kpis">
    <div class="kpi"><span>Total de chamados</span><b>${n}</b></div>
    <div class="kpi alt"><span>Ainda abertos</span><b>${nOpen}</b></div>
    <div class="kpi"><span>% abertos</span><b>${(100*nOpen/n).toFixed(1)}%</b></div>
    <div class="kpi alt"><span>Encerrados</span><b>${n-nOpen}</b></div>
    <div class="kpi warn"><span>Abertos com risco alto</span><b>${nAlto}</b></div>
  </div>
  <div class="grid3">
    <div class="card" style="grid-column:1/-1">
      <h3>Criados × encerrados (2026)</h3>
      ${monthSvg}
    </div>
    <div class="card"><h3>Família operacional</h3>${bars(byFam, "#00995D")}</div>
    <div class="card"><h3>Cluster do backlog</h3>${bars(byClu, "#00ADEF")}</div>
    <div class="card"><h3>Abertos por razão de status</h3>${bars(byRazao, "#FEBE00")}</div>
  </div>
  <p class="note">Página equivalente à Visão Operacional do PBIX, agora com a base classificada e o recorte de risco. Use a aba ao lado para a camada preditiva.</p>
</section>

<section id="pred" class="page">
  <div class="filters">
    <span class="chip">Modelo: Naive Bayes + regressão logística</span>
    <span class="chip">Corte de risco: ciclo &gt; 10 dias</span>
    <span class="chip">Holdout nível 78,4% · F1 risco 0,81</span>
  </div>
  <div class="kpis">
    <div class="kpi"><span>Acurácia do nível</span><b>78,4%</b></div>
    <div class="kpi alt"><span>F1 risco alto</span><b>0,81</b></div>
    <div class="kpi warn"><span>Abertos risco alto</span><b>${nAlto}</b></div>
    <div class="kpi"><span>Abertos perfil N1</span><b>${nN1}</b></div>
    <div class="kpi alt"><span>Abertos perfil N3</span><b>${nN3}</b></div>
  </div>
  <div class="grid">
    <div class="card"><h3>Nível sugerido pelo modelo</h3>${bars(byNivel, "#00ADEF")}</div>
    <div class="card"><h3>Risco de ciclo longo</h3>${bars(byRisco, "#CC3366")}</div>
    <div class="card" style="grid-column:1/-1; min-height:auto">
      <h3>Fila aberta priorizada</h3>
      <table>
        <tr><th>Id</th><th>Título</th><th>Nível</th><th>Risco</th><th>Score</th><th>Ação sugerida</th></tr>
        ${fila
          .map(
            (d) => `<tr>
              <td>${esc(d["Id Chamado Qualitor"])}</td>
              <td>${esc(String(d.Titulo).slice(0, 64))}</td>
              <td><span class="pill">${esc(d["Nivel predito (ML)"])}</span></td>
              <td class="${d["Risco ciclo longo"] === "Alto" ? "alto" : "baixo"}">${esc(d["Risco ciclo longo"])}</td>
              <td>${esc(d["Score risco"])}</td>
              <td>${esc(d["Acao sugerida"])}</td>
            </tr>`
          )
          .join("")}
      </table>
    </div>
  </div>
  <p class="note">Fonte: Chamados_N2_com_predicao.xlsx · Para publicar no Power BI Desktop: Obter dados → Excel → esta planilha → criar as visuais desta página.</p>
</section>
<script>
document.querySelectorAll('.tab').forEach(t => t.onclick = () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.remove('on'));
  document.querySelectorAll('.page').forEach(x => x.classList.remove('on'));
  t.classList.add('on');
  document.getElementById(t.dataset.p).classList.add('on');
});
</script>
</body>
</html>`;

const dest = path.join(__dirname, "Pagina_Relatorio_PowerBI.html");
fs.writeFileSync(dest, html, "utf8");
console.log("WROTE", dest, html.length);
