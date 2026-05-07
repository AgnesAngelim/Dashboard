// ── dash-clientes.JS — Dashboard de Clientes ─────────────────────────────────

let chartPortStatus = null, chartOperadoraSucesso = null, chartOperadoraNegada = null;

function renderizarClientes(d) {
  const c = d.clientes;
  document.getElementById("cli-total-docs").textContent        = fmt(c.docList.length);
  document.getElementById("cli-com-rep").textContent           = fmt(c.docList.filter(x=>x.repetidas>0).length);
  document.getElementById("cli-tels-rep").textContent          = fmt(c.telsDetalhes.length);
  document.getElementById("cli-alerta-rapido").textContent     = fmt(c.alertaRapido.length);
  document.getElementById("cli-multi-clientes").textContent    = fmt(c.telPorMultiplosClientes.length);
  document.getElementById("cli-alto-cancelamento").textContent = fmt(c.altoCancelamento.length);

  const tabDoc = document.getElementById("cli-tab-doc");
  if (tabDoc) tabDoc.innerHTML = c.docList.slice(0,100).map(x=>`<tr><td>${x.nome}</td><td style="color:var(--text-muted);font-size:11px;">${x.doc}</td><td><b>${fmt(x.total)}</b></td><td>${x.repetidas>0?`<span style="color:#F59E0B;font-weight:600;">${fmt(x.repetidas)}</span>`:`<span style="color:var(--text-muted);">0</span>`}</td></tr>`).join("");

  const tabRapido = document.getElementById("cli-tab-rapido");
  if (tabRapido) tabRapido.innerHTML = c.alertaRapido.slice(0,50).map(x=>`<tr><td>${x.nome}</td><td style="color:var(--text-muted);font-size:11px;">${x.doc}</td><td><b>${fmt(x.total)}</b></td><td><span class="risco-badge">${x.diasEntre3} dias</span></td></tr>`).join("");

  const tabMulti = document.getElementById("cli-tab-multi");
  if (tabMulti) tabMulti.innerHTML = c.telPorMultiplosClientes.slice(0,50).map(x=>`<tr><td style="font-family:monospace;font-size:12px;">${x.tel}</td><td><b>${fmt(x.totalLinhas)}</b></td><td><span class="risco-badge">${fmt(x.totalClientes)} clientes</span></td><td style="font-size:11px;color:var(--text-muted);">${x.docs.slice(0,2).join(", ")}${x.docs.length>2?` +${x.docs.length-2}`:""}</td></tr>`).join("");

  const tabCancel = document.getElementById("cli-tab-cancel");
  if (tabCancel) tabCancel.innerHTML = c.altoCancelamento.slice(0,50).map(x=>`<tr><td>${x.nome}</td><td style="color:var(--text-muted);font-size:11px;">${x.doc}</td><td><b>${fmt(x.total)}</b></td><td>${fmt(x.canceladas)}</td><td><span class="risco-badge risco-red">${x.pctCancel}%</span></td></tr>`).join("");

  if (chartPortStatus) { chartPortStatus.destroy(); chartPortStatus = null; }
  const statusEntries = Object.entries(c.statusPortMap).sort((a,b)=>b[1]-a[1]);
  const statusCores = {"Sucesso":"#10B981","Portabilidade negada":"#8B5CF6","Portabilidade em andamento":"#EC4899","Aguardando confirmação":"#06B6D4"};
  const elStatusChart = document.getElementById("chartPortStatus");
  if (elStatusChart) chartPortStatus = new Chart(elStatusChart,{type:"doughnut",data:{labels:statusEntries.map(([s])=>s),datasets:[{data:statusEntries.map(([,v])=>v),backgroundColor:statusEntries.map(([s])=>statusCores[s]||"#8B5CF6"),borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});

  const totalPort = Object.values(c.statusPortMap).reduce((a,b)=>a+b,0);
  const tabStatus = document.getElementById("cli-tab-status");
  if (tabStatus) tabStatus.innerHTML = statusEntries.map(([s,v])=>`<tr><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${statusCores[s]||"#8B5CF6"};margin-right:8px;"></span>${s}</td><td><b>${fmt(v)}</b></td><td style="color:var(--text-muted);font-size:11px;">${pct(v,totalPort)}</td></tr>`).join("");

  const tabTels = document.getElementById("cli-tab-tels");
  if (tabTels) tabTels.innerHTML = c.telsDetalhes.map(t=>`<tr><td style="font-family:monospace;font-size:12px;">${t.tel}</td><td><b>${fmt(t.count)}</b></td><td style="font-size:11px;color:var(--text-muted);">${t.docs.slice(0,3).join(", ")}${t.docs.length>3?` +${t.docs.length-3}`:""}</td></tr>`).join("");

  const tickColor="#94A3B8", gridColor="#ffffff0f";
  const top10 = c.operadoraList.slice(0,10);
  if (chartOperadoraSucesso) { chartOperadoraSucesso.destroy(); chartOperadoraSucesso=null; }
  const elOpSucesso = document.getElementById("chartOpSucesso");
  if (elOpSucesso&&top10.length) chartOperadoraSucesso=new Chart(elOpSucesso,{type:"bar",data:{labels:top10.map(o=>o.op),datasets:[{label:"Sucesso",data:top10.map(o=>o.sucesso),backgroundColor:"#10B981",borderRadius:3},{label:"Negada",data:top10.map(o=>o.negada),backgroundColor:"#EF4444",borderRadius:3},{label:"Andamento",data:top10.map(o=>o.andamento),backgroundColor:"#F59E0B",borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,labels:{color:tickColor,font:{size:11}}}},scales:{x:{ticks:{color:tickColor,font:{size:10}},grid:{display:false},stacked:true},y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true,stacked:true}}}});

  if (chartOperadoraNegada) { chartOperadoraNegada.destroy(); chartOperadoraNegada=null; }
  const elOpNegada = document.getElementById("chartOpNegada");
  const topNegadas = [...c.operadoraList].sort((a,b)=>b.negada-a.negada).slice(0,10);
  if (elOpNegada&&topNegadas.length) chartOperadoraNegada=new Chart(elOpNegada,{type:"bar",data:{labels:topNegadas.map(o=>o.op),datasets:[{label:"Negadas",data:topNegadas.map(o=>o.negada),backgroundColor:"#EF4444",borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:tickColor,font:{size:10}},grid:{display:false}},y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true}}}});

  const tabOp = document.getElementById("cli-tab-op");
  if (tabOp) tabOp.innerHTML = c.operadoraList.slice(0,20).map(o=>`<tr><td>${o.op}</td><td>${fmt(o.total)}</td><td style="color:#10B981;font-weight:600;">${fmt(o.sucesso)} <small style="color:var(--text-muted);font-weight:400;">(${o.taxaSucesso}%)</small></td><td style="color:#EF4444;font-weight:600;">${fmt(o.negada)} <small style="color:var(--text-muted);font-weight:400;">(${o.taxaNegada}%)</small></td></tr>`).join("");
}

function filtrarTabelaDoc(termo) {
  const rows = document.querySelectorAll("#cli-tab-doc tr");
  const t = termo.toLowerCase();
  rows.forEach(row => { row.style.display = row.textContent.toLowerCase().includes(t) ? "" : "none"; });
}