// ── dash-clientes.js — Dashboard de Clientes ─────────────────────────────────

let chartPortStatus = null;

// Função genérica de busca para qualquer tabela
function filtrarTabela(inputId, tbodyId) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("input", () => {
    const t = input.value.toLowerCase();
    document.querySelectorAll(`#${tbodyId} tr`).forEach(row => {
      row.style.display = row.textContent.toLowerCase().includes(t) ? "" : "none";
    });
  });
}

function renderizarClientes(d) {
  const c = d.clientes;

  // KPI cards
  document.getElementById("cli-total-docs").textContent     = fmt(c.docList.length);
  document.getElementById("cli-com-rep").textContent        = fmt(c.docList.filter(x=>x.repetidas>0).length);
  document.getElementById("cli-tels-rep").textContent       = fmt(c.telsDetalhes.length);
  document.getElementById("cli-multi-clientes").textContent = fmt(c.telPorMultiplosClientes.length);

  // Tabela: linhas ativadas e repetidas — só com repetidas, sem limite
  const tabDoc = document.getElementById("cli-tab-doc");
  if (tabDoc) {
    tabDoc.innerHTML = c.docList
      .filter(x => x.repetidas > 0)
      .sort((a, b) => b.repetidas - a.repetidas)
      .map(x => `
        <tr>
          <td>${x.nome}</td>
          <td style="color:var(--text-muted);font-size:11px;">${x.doc}</td>
          <td style="color:var(--text-muted);font-size:11px;">${x.idLicenciado || "—"}</td>
          <td><b>${fmt(x.total)}</b></td>
          <td><span style="color:#F59E0B;font-weight:600;">${fmt(x.repetidas)}</span></td>
        </tr>`).join("");
  }
  filtrarTabela("cli-search-doc", "cli-tab-doc");

  // Tabela: número em múltiplos clientes com ID licenciado
  const tabMulti = document.getElementById("cli-tab-multi");
  if (tabMulti) tabMulti.innerHTML = c.telPorMultiplosClientes.slice(0, 50).map(x => `
    <tr>
      <td style="font-family:monospace;font-size:12px;">${x.tel}</td>
      <td><b>${fmt(x.totalLinhas)}</b></td>
      <td><span class="risco-badge">${fmt(x.totalClientes)} clientes</span></td>
      <td style="font-size:11px;color:var(--text-muted);">${(x.idsLicenciado||[]).slice(0,2).join(", ")}${(x.idsLicenciado||[]).length>2?` +${(x.idsLicenciado||[]).length-2}`:""}</td>
      <td style="font-size:11px;color:var(--text-muted);">${x.docs.slice(0,2).join(", ")}${x.docs.length>2?` +${x.docs.length-2}`:""}</td>
    </tr>`).join("");
  filtrarTabela("cli-search-multi", "cli-tab-multi");

  // Tabela: índice de cancelamento com ID licenciado
  const tabCancel = document.getElementById("cli-tab-cancel");
  if (tabCancel) tabCancel.innerHTML = c.altoCancelamento.slice(0, 50).map(x => `
    <tr>
      <td>${x.nome}</td>
      <td style="color:var(--text-muted);font-size:11px;">${x.doc}</td>
      <td style="color:var(--text-muted);font-size:11px;">${x.idLicenciado || "—"}</td>
      <td><b>${fmt(x.total)}</b></td>
      <td>${fmt(x.canceladas)}</td>
      <td><span class="risco-badge risco-red">${x.pctCancel}%</span></td>
    </tr>`).join("");
  filtrarTabela("cli-search-cancel", "cli-tab-cancel");

  // Tabela: clientes com mais de uma recarga
  window._dadosRecarga = c.comMaisDeUmaRecarga;
  const tabRecarga = document.getElementById("cli-tab-recarga");
  if (tabRecarga) tabRecarga.innerHTML = c.comMaisDeUmaRecarga.map((x, i) => `
    <tr>
      <td><span style="cursor:pointer;font-weight:500;" onclick="abrirModalLinhas(${i})">${x.nome}</span></td>
      <td style="color:var(--text-muted);font-size:11px;">${x.doc}</td>
      <td style="color:var(--text-muted);font-size:11px;">${x.idLicenciado || "—"}</td>
      <td><b>${fmt(x.total)}</b></td>
      <td><span style="color:#10B981;font-weight:600;">${fmt(x.recargas)}</span></td>
    </tr>`).join("");
  filtrarTabela("cli-search-recarga", "cli-tab-recarga");

  // Gráfico status portabilidades com filtro de mês
  const statusCores = {"Sucesso":"#10B981","Portabilidade negada":"#8B5CF6","Portabilidade em andamento":"#EC4899","Aguardando confirmação":"#06B6D4"};

  const filtroPortMes = document.getElementById("filtro-port-mes");
  if (filtroPortMes) {
    const meses = Object.keys(c.statusPortMapPorMes || {}).filter(m => m !== "sem-data").sort();
    filtroPortMes.innerHTML =
      `<button class="filtro-btn active" data-mes="todos">Todos</button>` +
      meses.map(m => `<button class="filtro-btn" data-mes="${m}">${formatarMesLabel(m)}</button>`).join("");
    filtroPortMes.querySelectorAll(".filtro-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        filtroPortMes.querySelectorAll(".filtro-btn").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        const mes = btn.dataset.mes;
        const map = mes === "todos" ? c.statusPortMap : (c.statusPortMapPorMes[mes] || {});
        atualizarGraficoPortStatus(map, statusCores);
      });
    });
  }
  atualizarGraficoPortStatus(c.statusPortMap, statusCores);

  // Tabela: números repetidos na operadora doadora com ID licenciado
  const tabTels = document.getElementById("cli-tab-tels");
  if (tabTels) tabTels.innerHTML = c.telsDetalhes.map(t=>`
    <tr>
      <td style="font-family:monospace;font-size:12px;">${t.tel}</td>
      <td><b>${fmt(t.count)}</b></td>
      <td style="font-size:11px;color:var(--text-muted);">${(t.idsLicenciado||[]).slice(0,2).join(", ")}${(t.idsLicenciado||[]).length>2?` +${(t.idsLicenciado||[]).length-2}`:""}</td>
      <td style="font-size:11px;color:var(--text-muted);">${t.docs.slice(0,3).join(", ")}${t.docs.length>3?` +${t.docs.length-3}`:""}</td>
    </tr>`).join("");
  filtrarTabela("cli-search-tels", "cli-tab-tels");

  // Detalhamento por operadora
  const tabOp = document.getElementById("cli-tab-op");
  if (tabOp) tabOp.innerHTML = c.operadoraList.slice(0,20).map(o=>`<tr><td>${o.op}</td><td>${fmt(o.total)}</td><td style="color:#10B981;font-weight:600;">${fmt(o.sucesso)} <small style="color:var(--text-muted);font-weight:400;">(${o.taxaSucesso}%)</small></td><td style="color:#EF4444;font-weight:600;">${fmt(o.negada)} <small style="color:var(--text-muted);font-weight:400;">(${o.taxaNegada}%)</small></td></tr>`).join("");
}

function filtrarTabelaDoc(termo) {
  const t = termo.toLowerCase();
  document.querySelectorAll("#cli-tab-doc tr").forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(t) ? "" : "none";
  });
}

function atualizarGraficoPortStatus(map, cores) {
  const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);
  if (chartPortStatus) { chartPortStatus.destroy(); chartPortStatus = null; }
  const elStatusChart = document.getElementById("chartPortStatus");
  if (elStatusChart) chartPortStatus = new Chart(elStatusChart, {
    type: "doughnut",
    data: {
      labels: entries.map(([s]) => s),
      datasets: [{ data: entries.map(([,v]) => v), backgroundColor: entries.map(([s]) => cores[s] || "#8B5CF6"), borderWidth: 0 }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
  });
  const total = Object.values(map).reduce((a,b) => a+b, 0);
  const tabStatus = document.getElementById("cli-tab-status");
  if (tabStatus) tabStatus.innerHTML = entries.map(([s,v]) =>
    `<tr><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${cores[s]||"#8B5CF6"};margin-right:8px;"></span>${s}</td><td><b>${fmt(v)}</b></td><td style="color:var(--text-muted);font-size:11px;">${pct(v,total)}</td></tr>`
  ).join("");
}

function abrirModalLinhas(idx) {
  const x = (window._dadosRecarga || [])[idx];
  if (!x) return;
  const modal = document.getElementById("modalLinhasCli");
  const header = document.getElementById("modalLinhasCliHeader");
  const body   = document.getElementById("modalLinhasCliBody");
  if (!modal || !header || !body) return;

  header.innerHTML = `
    <span>📋 ${x.nome}</span>
    <span style="font-size:13px;font-weight:400;color:var(--text-muted);">${x.doc} · ${fmt(x.total)} linha${x.total>1?"s":""} · ${fmt(x.recargas)} recarga${x.recargas>1?"s":""}</span>`;

  const linhas = (x.linhas || []).sort((a,b) => b.recargas - a.recargas);
  body.innerHTML = linhas.length === 0
    ? `<div style="text-align:center;padding:2rem;color:var(--text-muted);">Nenhuma linha encontrada.</div>`
    : `<table class="split-table">
        <thead><tr><th>Linha</th><th>Plano</th><th>Status</th><th>Ativação</th><th>Recargas</th></tr></thead>
        <tbody>
          ${linhas.map(l => `
            <tr>
              <td style="font-family:monospace;font-size:12px;">${l.linha}</td>
              <td>${l.plano}</td>
              <td><span style="color:${l.status.toLowerCase().includes("ativ") ? "#10B981" : l.status.toLowerCase().includes("cancel") ? "#EF4444" : "#94A3B8"};font-size:12px;">${l.status}</span></td>
              <td style="color:var(--text-muted);font-size:12px;">${l.dataAtiv}</td>
              <td><span style="color:#10B981;font-weight:600;">${fmt(l.recargas)}</span></td>
            </tr>`).join("")}
        </tbody>
      </table>`;

  modal.style.display = "flex";
  requestAnimationFrame(() => modal.classList.add("modal-visible"));
}

function fecharModalLinhas() {
  const modal = document.getElementById("modalLinhasCli");
  if (!modal) return;
  modal.classList.remove("modal-visible");
  setTimeout(() => { modal.style.display = "none"; }, 200);
}

document.addEventListener("keydown", (e) => { if (e.key === "Escape") fecharModalLinhas(); });
