let chartPort = null, chartNova = null, chartPortGeral = null, chartTimeline = null;
let chartChip = null, chartPgto = null, chartRecorrencia = null, chartOperadoras = null, chartLogistica = null;

function renderDash(d) {
  preencherCards(d);
  preencherKPI(d);
  renderizarPlanos(d);
  renderizarGraficos(d);
  renderizarClientes(d);
  renderizarCampanhas(d);
}

function preencherCards(d) {
  set("total", fmt(d.total));
  set("ativos", fmt(d.ativos));         set("ativos-pct",    pct(d.ativos, d.total)    + " do total");
  set("cancelados", fmt(d.cancelados)); set("cancelados-pct", pct(d.cancelados, d.total) + " do total");
  set("portabilidades", fmt(d.portabilidades.total)); set("port-pct", pct(d.portabilidades.total, d.total) + " do total");
  set("novas", fmt(d.novasLinhas.total)); set("novas-pct", pct(d.novasLinhas.total, d.total) + " do total");
  set("bonificados", fmt(d.bonificados.total)); set("bonif-pct", pct(d.bonificados.total, d.total) + " do total");
  set("logistica", fmt(d.logistica.total)); set("logistica-pct", pct(d.logistica.total, d.total) + " do total");
  set("tempo-vida", d.tempoMedioVida + " dias");

  // Portabilidades
  const [pPortAtiv, pPortCanc] = pctGroup([d.portabilidades.ativos, d.portabilidades.cancelados], d.portabilidades.total);
  set("port-total",          fmt(d.portabilidades.total));
  set("port-ativos",         fmt(d.portabilidades.ativos));     set("port-ativos-pct",     pPortAtiv);
  set("port-cancelados",     fmt(d.portabilidades.cancelados)); set("port-cancelados-pct", pPortCanc);
  set("port-status-total",   fmt(d.portabilidades.totalAtivos));
  const [pAprov, pNeg, pAnd] = pctGroup([d.portabilidades.aprovadas, d.portabilidades.negadas, d.portabilidades.andamento], d.portabilidades.totalAtivos);
  set("port-aprov",      fmt(d.portabilidades.aprovadas));  set("port-aprov-pct",    pAprov);
  set("port-neg",        fmt(d.portabilidades.negadas));    set("port-neg-pct",      pNeg);
  set("port-andamento",  fmt(d.portabilidades.andamento));  set("port-andamento-pct",pAnd);

  // Novas linhas
  const [pNovasAtiv, pNovasCanc] = pctGroup([d.novasLinhas.ativos, d.novasLinhas.cancelados], d.novasLinhas.total);
  set("novas-total",          fmt(d.novasLinhas.total));
  set("novas-ativos",         fmt(d.novasLinhas.ativos));     set("novas-ativos-pct",    pNovasAtiv);
  set("novas-cancelados",     fmt(d.novasLinhas.cancelados)); set("novas-cancelados-pct", pNovasCanc);

  // Chip
  const [pEsim, pFisico] = pctGroup([d.chip.esim, d.chip.fisico], d.chip.total);
  set("chip-total",      fmt(d.chip.total));
  set("chip-esim",       fmt(d.chip.esim));    set("chip-esim-pct",   pEsim);
  set("chip-fisico",     fmt(d.chip.fisico));  set("chip-fisico-pct", pFisico);

  // Pagamento
  set("pgto-total", fmt(d.pagamento.total));
  const tabelaPgto = document.getElementById("pgto-tabela");
  if (tabelaPgto) {
    const pgtoEntries = Object.entries(d.pagamento.formas).sort((a,b)=>b[1]-a[1]);
    const pgtoQtds    = pgtoEntries.map(([,q])=>q);
    const pgtoPcts    = pctGroup(pgtoQtds, d.pagamento.total);
    tabelaPgto.innerHTML = pgtoEntries.map(([nome,qtd],i)=>`<tr><td>${nome}</td><td>${fmt(qtd)}</td><td>${pgtoPcts[i]}</td></tr>`).join("");
  }

  // Bonificados
  set("bonif-total",      fmt(d.bonificados.total));
  set("bonif-normal",     fmt(d.bonificados.normal));
  set("pagamento-normal", fmt(d.bonificados.normal));

  // Recorrência
  const totalRec = d.recorrencia.com + d.recorrencia.sem;
  const [pRecCom, pRecSem] = pctGroup([d.recorrencia.com, d.recorrencia.sem], totalRec);
  set("rec-total",   fmt(totalRec));
  set("rec-com",     fmt(d.recorrencia.com)); set("rec-com-pct", pRecCom);
  set("rec-sem",     fmt(d.recorrencia.sem)); set("rec-sem-pct", pRecSem);

  // Logística
  const [pLogAtiv, pLogNao, pLogCanc] = pctGroup([d.logistica.ativa, d.logistica.naoAtivada, d.logistica.cancelada], d.logistica.total);
  set("log-total",       fmt(d.logistica.total));
  set("log-ativa",       fmt(d.logistica.ativa));       set("log-ativa-pct",  pLogAtiv);
  set("log-nao-ativada", fmt(d.logistica.naoAtivada));  set("log-nao-pct",    pLogNao);
  set("log-cancelada",   fmt(d.logistica.cancelada));   set("log-cancel-pct", pLogCanc);
}

function preencherKPI(d) {
  const k = d.kpi;
  const varAtiv  = variacao(k.ativacoes.atual, k.ativacoes.anterior);
  const varCanc  = variacaoInversa(k.cancelamentos.atual, k.cancelamentos.anterior);
  set("kpi-mes-ativ",   k.mesAtual);
  set("kpi-mes-canc",   k.mesAtual);
  set("kpi-ativ-atual", fmt(k.ativacoes.atual));      set("kpi-ativ-ant", fmt(k.ativacoes.anterior));
  const elVarAtiv = document.getElementById("kpi-ativ-var");
  if (elVarAtiv) { elVarAtiv.textContent = varAtiv.texto; elVarAtiv.className = "kpi-var " + varAtiv.classe; }
  set("kpi-canc-atual", fmt(k.cancelamentos.atual));  set("kpi-canc-ant", fmt(k.cancelamentos.anterior));
  const elVarCanc = document.getElementById("kpi-canc-var");
  if (elVarCanc) { elVarCanc.textContent = varCanc.texto; elVarCanc.className = "kpi-var " + varCanc.classe; }
  // Churn
  set("kpi-mes-churn",   k.mesAtual);
  set("kpi-churn-atual", k.churn.atual + "%");
  set("kpi-churn-ant",   k.churn.anterior + "%");
  const varChurn = variacaoInversa(k.churn.atual, k.churn.anterior);
  const elVarChurn = document.getElementById("kpi-churn-var");
  if (elVarChurn) { elVarChurn.textContent = varChurn.texto; elVarChurn.className = "kpi-var " + varChurn.classe; }
}

function renderizarPlanos(d) {
  const container = document.getElementById("planosChart");
  if (!container) return;
  const cores = ["#10B981","#06B6D4","#F59E0B","#8B5CF6","#EC4899","#F97316","#14B8A6","#6366F1"];
  const maximo = Math.max(...d.planos.map(p => p.total));
  container.innerHTML = d.planos.sort((a,b)=>b.total-a.total).map((p,i)=>{
    const largura = Math.round((p.total/maximo)*100);
    return `<div class="bar-row"><div class="bar-label" title="${p.nome}">${p.nome}</div><div class="bar-track"><div class="bar-fill" style="width:${largura}%;background:${cores[i%cores.length]};">${fmt(p.total)}</div></div><div class="bar-count">${fmt(p.total)}</div></div>`;
  }).join("");
}

function renderizarGraficos(d) {
  const gridColor = "#ffffff0f", tickColor = "#94A3B8";
  [chartPort,chartNova,chartPortGeral,chartChip,chartPgto,chartTimeline,chartRecorrencia,chartOperadoras,chartLogistica].forEach(c=>{if(c)c.destroy();});
  chartPort=chartNova=chartPortGeral=chartChip=chartPgto=chartTimeline=chartRecorrencia=chartOperadoras=chartLogistica=null;

  const optD = { responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } } };
  chartPort = new Chart(document.getElementById("chartPort"),{type:"doughnut",data:{labels:["Ativos","Cancelados"],datasets:[{data:[d.portabilidades.ativos,d.portabilidades.cancelados],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  chartNova = new Chart(document.getElementById("chartNova"),{type:"doughnut",data:{labels:["Ativos","Cancelados"],datasets:[{data:[d.novasLinhas.ativos,d.novasLinhas.cancelados],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  const elPortGeral = document.getElementById("chartPortGeral");
  if (elPortGeral) chartPortGeral = new Chart(elPortGeral,{type:"doughnut",data:{labels:["Aprovadas","Negadas","Em andamento"],datasets:[{data:[d.portabilidades.aprovadas,d.portabilidades.negadas,d.portabilidades.andamento],backgroundColor:["#EC4899","#8B5CF6","#10B981"],borderWidth:0}]},options:optD});
  chartChip = new Chart(document.getElementById("chartChip"),{type:"doughnut",data:{labels:["eSIM","Físico"],datasets:[{data:[d.chip.esim,d.chip.fisico],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  chartRecorrencia = new Chart(document.getElementById("chartRecorrencia"),{type:"doughnut",data:{labels:["Com recorrência","Sem recorrência"],datasets:[{data:[d.recorrencia.com,d.recorrencia.sem],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  const elLog = document.getElementById("chartLogistica");
  if (elLog) chartLogistica = new Chart(elLog,{type:"doughnut",data:{labels:["Ativa","Não ativada","Cancelada"],datasets:[{data:[d.logistica.ativa,d.logistica.naoAtivada,d.logistica.cancelada],backgroundColor:["#8B5CF6","#EC4899","#10B981"],borderWidth:0}]},options:optD});
  const pgtoEntries = Object.entries(d.pagamento.formas).sort((a,b)=>b[1]-a[1]);
  chartPgto = new Chart(document.getElementById("chartPgto"),{type:"doughnut",data:{labels:pgtoEntries.map(([n])=>n),datasets:[{data:pgtoEntries.map(([,q])=>q),backgroundColor:["#EC4899","#8B5CF6","#10B981","#06B6D4","#F59E0B","#F97316","#14B8A6"],borderWidth:0}]},options:optD});
  const topOp = Object.entries(d.operadoras).sort((a,b)=>b[1]-a[1]).slice(0,6);
  chartOperadoras = new Chart(document.getElementById("chartOperadoras"),{type:"bar",data:{labels:topOp.map(([n])=>n),datasets:[{data:topOp.map(([,q])=>q),backgroundColor:["#10B981","#06B6D4","#F59E0B","#8B5CF6","#EC4899","#F97316"],borderRadius:4,borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:tickColor,font:{size:11}},grid:{display:false}},y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true}}}});

  // Timeline — usa dias se mês específico selecionado
  const mesSel = dadosGlobais?._mesSelecionado;
  let tlLabels, tlAtiv, tlCanc;
  if (mesSel && mesSel !== "todos" && d.timeline.ativacoesPorDia[mesSel]) {
    const [ano, mes] = mesSel.split("-").map(Number);
    const diasNoMes  = new Date(ano, mes, 0).getDate();
    const diasAtiv   = d.timeline.ativacoesPorDia[mesSel] || {};
    const diasCanc   = d.timeline.cancelamentosPorDia[mesSel] || {};
    tlLabels = Array.from({length:diasNoMes}, (_,i) => String(i+1));
    tlAtiv   = Array.from({length:diasNoMes}, (_,i) => diasAtiv[i+1] || 0);
    tlCanc   = Array.from({length:diasNoMes}, (_,i) => diasCanc[i+1] || 0);
  } else {
    tlLabels = d.timeline.labels;
    tlAtiv   = d.timeline.ativacoes;
    tlCanc   = d.timeline.cancelamentos;
  }
  set("legend-atv-total",  fmt(tlAtiv.reduce((a,b)=>a+b,0)));
  set("legend-canc-total", fmt(tlCanc.reduce((a,b)=>a+b,0)));
  chartTimeline = new Chart(document.getElementById("chartTimeline"),{type:"line",data:{labels:tlLabels,datasets:[
    {label:"Ativações",data:tlAtiv,borderColor:"#10B981",backgroundColor:"#10b9811a",fill:true,tension:0.3,pointRadius:3},
    {label:"Cancelamentos",data:tlCanc,borderColor:"#EF4444",backgroundColor:"#ef44441a",fill:true,tension:0.3,pointRadius:3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{
    x:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor}},
    y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true},
  }}});
}

// ── Mapa ──────────────────────────────────────────────────────────────────────
function renderizarMapa(d) {
  const estados = d.estados;
  const cidadesPorEstado = d.cidadesPorEstado || {};
  const maxVal  = Math.max(...Object.values(estados), 1);
  const colorScale = d3.scaleLinear().domain([0, maxVal]).range(["#334155", "#10B981"]);
  const ranking = Object.entries(estados).sort((a,b)=>b[1]-a[1]);
  const totalEstados = Object.values(estados).reduce((a,b)=>a+b,0);
  const tabelaEl = document.getElementById("mapaTabela");
  if (tabelaEl) tabelaEl.innerHTML = `<table class="split-table"><thead><tr><th>Estado</th><th>Clientes</th><th>%</th></tr></thead><tbody>${ranking.map(([uf,val])=>`<tr style="cursor:pointer;" onclick="abrirModalCidades('${uf}',dadosGlobais)"><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${colorScale(val)};margin-right:8px;"></span>${nomesEstados[uf]||uf} (${uf})</td><td>${fmt(val)}</td><td style="color:var(--text-muted);font-size:11px;">${pct(val,totalEstados)}</td></tr>`).join("")}</tbody></table>`;
  const container = document.getElementById("mapaBrasil");
  container.innerHTML = "";
  fetch("https://raw.githubusercontent.com/codeforamerica/click_that_hood/master/public/data/brazil-states.geojson")
    .then(r=>r.json())
    .then(geojson=>{
      const w=container.offsetWidth||400, h=220;
      const projection=d3.geoMercator().fitSize([w,h],geojson);
      const path=d3.geoPath().projection(projection);
      const svg=d3.select(container).append("svg").attr("width",w).attr("height",h).attr("viewBox",`0 0 ${w} ${h}`);
      const tip=d3.select("body").selectAll("#mapaTooltipEl").data([1]).join("div").attr("id","mapaTooltipEl")
        .style("position","fixed").style("display","none").style("background","var(--bg-card)").style("border","1px solid var(--border)").style("border-radius","8px").style("padding","8px 12px").style("font-size","12px").style("color","var(--text)").style("pointer-events","none").style("z-index","999");
      svg.selectAll("path").data(geojson.features).enter().append("path").attr("d",path)
        .attr("fill",f=>colorScale(estados[f.properties.sigla||f.id]||0))
        .attr("stroke","var(--bg)").attr("stroke-width",0.5).style("cursor","pointer")
        .on("mousemove",(ev,f)=>{
          const uf=f.properties.sigla||f.id;
          const cidades=cidadesPorEstado[uf]||{};
          const numCidades=Object.keys(cidades).length;
          tip.style("display","block").style("left",(ev.clientX+14)+"px").style("top",(ev.clientY-10)+"px")
            .html(`<strong>${nomesEstados[uf]||uf}</strong><br>${fmt(estados[uf]||0)} clientes${numCidades>0?`<br><span style="color:var(--text-muted);font-size:11px;"> ${numCidades} cidade${numCidades>1?"s":""} · clique para detalhes</span>`:""}`);
        })
        .on("mouseleave",()=>tip.style("display","none"))
        .on("click",(ev,f)=>{ tip.style("display","none"); abrirModalCidades(f.properties.sigla||f.id, d); });
      svg.selectAll("text").data(geojson.features).enter().append("text")
        .attr("x",f=>path.centroid(f)[0]).attr("y",f=>path.centroid(f)[1])
        .attr("text-anchor","middle").attr("dominant-baseline","middle").attr("font-size","7").attr("font-weight","600").attr("fill","white").style("pointer-events","none")
        .text(f=>f.properties.sigla||f.id||"");
    })
    .catch(()=>{container.innerHTML='<div style="color:var(--text-muted);padding:1rem;">❌ Não foi possível carregar o mapa.</div>';});
}

function abrirModalCidades(uf, d) {
  const cidades = (d.cidadesPorEstado || {})[uf] || {};
  const totalUF = d.estados[uf] || 0;
  const nomeEstado = nomesEstados[uf] || uf;
  const ranking = Object.entries(cidades).sort((a,b)=>b[1]-a[1]);
  const maxCidade = ranking.length ? ranking[0][1] : 1;
  const semCidadeCount = totalUF - Object.values(cidades).reduce((a,b)=>a+b,0);
  const linhas = ranking.map(([cidade, qtd]) => {
    const largura = Math.round((qtd / maxCidade) * 100);
    return `<div data-cidade="${cidade.replace(/"/g,'&quot;')}" style="margin-bottom:10px;">
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px;">
        <span style="color:var(--text);">${cidade}</span>
        <span style="color:var(--text-muted);">${fmt(qtd)} <span style="font-size:11px;">(${pct(qtd,totalUF)})</span></span>
      </div>
      <div style="height:14px;background:var(--border);border-radius:3px;overflow:hidden;">
        <div style="height:100%;width:${largura}%;background:#10B981;border-radius:3px;transition:width .4s ease;"></div>
      </div>
    </div>`;
  }).join("");
  const semCidadeHtml = semCidadeCount > 0 ? `<p style="font-size:12px;color:var(--text-muted);margin-top:8px;">⚠️ ${fmt(semCidadeCount)} registro${semCidadeCount>1?"s":""} sem cidade informada</p>` : "";
  const semDadosHtml  = ranking.length === 0 ? `<div style="text-align:center;padding:2rem;color:var(--text-muted);">Nenhuma cidade encontrada.</div>` : "";
  document.getElementById("modalCidadesHeader").innerHTML = `<span>🏙️ ${nomeEstado} (${uf})</span><span style="font-size:13px;font-weight:400;color:var(--text-muted);">${fmt(totalUF)} clientes · ${ranking.length} cidades</span>`;
  document.getElementById("modalCidadesBody").innerHTML = semDadosHtml + linhas + semCidadeHtml;
  const modal = document.getElementById("modalCidades");
  const busca = document.getElementById("modalCidadesBusca");
  if (busca) busca.value = "";
  modal.style.display = "flex";
  requestAnimationFrame(() => modal.classList.add("modal-visible"));
}
function fecharModalCidades() {
  const modal = document.getElementById("modalCidades");
  modal.classList.remove("modal-visible");
  setTimeout(() => { modal.style.display = "none"; }, 200);
}
document.addEventListener("keydown", (e) => { if (e.key === "Escape") fecharModalCidades(); });
function filtrarModalCidades(termo) {
  const t = termo.toLowerCase();
  const body = document.getElementById("modalCidadesBody");
  if (!body) return;
  body.querySelectorAll("[data-cidade]").forEach(el => { el.style.display = el.dataset.cidade.toLowerCase().includes(t) ? "" : "none"; });
}
