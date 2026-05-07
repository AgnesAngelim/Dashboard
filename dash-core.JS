// ── dash-core.JS — utilitários, cálculo, upload, nav, tema, filtro global ──────

const pct = (parte, total) => total ? ((parte / total) * 100).toFixed(1) + "%" : "0%";
const pctGroup = (partes, total) => {
  if (!total) return partes.map(() => "0%");
  const vals = partes.map(p => parseFloat(((p / total) * 100).toFixed(1)));
  const soma = vals.slice(0, -1).reduce((a, b) => a + b, 0);
  vals[vals.length - 1] = parseFloat((100 - soma).toFixed(1));
  return vals.map(v => v + "%");
};
const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
const fmt = (n) => Number(n).toLocaleString("pt-BR");

let dadosGlobais = null;
let arquivoPrincipal  = null;
let arquivoBackoffice = null;

const nomesEstados = {
  AC:"Acre", AL:"Alagoas", AM:"Amazonas", AP:"Amapá", BA:"Bahia",
  CE:"Ceará", DF:"Distrito Federal", ES:"Espírito Santo", GO:"Goiás",
  MA:"Maranhão", MG:"Minas Gerais", MS:"Mato Grosso do Sul", MT:"Mato Grosso",
  PA:"Pará", PB:"Paraíba", PE:"Pernambuco", PI:"Piauí", PR:"Paraná",
  RJ:"Rio de Janeiro", RN:"Rio Grande do Norte", RO:"Rondônia", RR:"Roraima",
  RS:"Rio Grande do Sul", SC:"Santa Catarina", SE:"Sergipe", SP:"São Paulo",
  TO:"Tocantins",
};

// ── Tema ──────────────────────────────────────────────────────────────────────
function setupTema() {
  const btn = document.getElementById("btnTema");
  const saved = localStorage.getItem("tema") || "dark";
  aplicarTema(saved);
  btn.addEventListener("click", () => {
    const atual = document.body.dataset.tema || "dark";
    const novo = atual === "dark" ? "light" : "dark";
    aplicarTema(novo);
    localStorage.setItem("tema", novo);
  });
}
function aplicarTema(tema) {
  document.body.dataset.tema = tema;
  const btn = document.getElementById("btnTema");
  btn.textContent = tema === "dark" ? "☀️ Modo claro" : "🌙 Modo escuro";
}

// ── Navegação ─────────────────────────────────────────────────────────────────
function setupNav() {
  document.querySelectorAll(".nav-item").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".nav-item").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const page = btn.dataset.page;
      document.querySelectorAll(".page").forEach(p => p.style.display = "none");
      document.getElementById("page-" + page).style.display = "block";
    });
  });
}

// ── Upload ────────────────────────────────────────────────────────────────────
function setupDropZone() {
  configurarZone("dropZone1", "fileInput1", "status1", processarPrincipal);
  configurarZone("dropZone2", "fileInput2", "status2", processarBackoffice);
}
function configurarZone(zoneId, inputId, statusId, handler) {
  const zone  = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone || !input) return;
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", (e) => { if (e.target.files[0]) handler(e.target.files[0], statusId, zone); });
  zone.addEventListener("dragover",  (e) => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault(); zone.classList.remove("dragover");
    if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0], statusId, zone);
  });
}
function lerExcel(file, callback) {
  const reader = new FileReader();
  reader.onload = (e) => {
    setTimeout(() => {
      try {
        callback(XLSX.read(e.target.result, { type: "array", dense: true, cellDates: false, raw: true }));
      } catch (err) {
        const el = document.getElementById("uploadStatus");
        if (el) el.textContent = "❌ Erro: " + err.message;
      }
    }, 10);
  };
  reader.readAsArrayBuffer(file);
}
function processarPrincipal(file, statusId, zone) {
  const badge = document.getElementById(statusId);
  badge.textContent = "⏳ Lendo..."; badge.className = "upload-badge loading";
  lerExcel(file, (wb) => {
    const aba = wb.SheetNames[0];
    const dados = XLSX.utils.sheet_to_json(wb.Sheets[aba], { defval: null });
    if (!dados.length) { badge.textContent = "❌ Planilha vazia"; badge.className = "upload-badge erro"; return; }
    arquivoPrincipal = dados;
    badge.textContent = `✅ ${fmt(dados.length)} registros`; badge.className = "upload-badge ok";
    zone.style.borderColor = "#10B981";
    tentarProcessar();
  });
}
function processarBackoffice(file, statusId, zone) {
  const badge = document.getElementById(statusId);
  badge.textContent = "⏳ Lendo..."; badge.className = "upload-badge loading";
  lerExcel(file, (wb) => {
    const abaBO = wb.SheetNames.find(n => n.toLowerCase().replace(/\s/g,"").includes("backoffice")) || wb.SheetNames[0];
    const bo = XLSX.utils.sheet_to_json(wb.Sheets[abaBO], { raw: false, defval: null });
    arquivoBackoffice = new Set(
      bo.map(r => String(r["Chip"] || "").trim().split(".")[0]).filter(v => v !== "")
    );
    badge.textContent = `✅ ${fmt(arquivoBackoffice.size)} chips`; badge.className = "upload-badge ok";
    zone.style.borderColor = "#10B981";
    tentarProcessar();
  });
}
function tentarProcessar() {
  const status = document.getElementById("uploadStatus");
  if (!arquivoPrincipal)  { status.textContent = "⏳ Aguardando planilha principal..."; return; }
  if (!arquivoBackoffice) { status.textContent = "⏳ Aguardando planilha BackOffice..."; return; }
  status.textContent = "⏳ Processando dados...";
  const dados = calcularDados(arquivoPrincipal, arquivoBackoffice);
  dadosGlobais = dados;
  renderDash(dados);
  construirFiltroGlobal(dados);
  document.getElementById("uploadArea").style.display = "none";
  document.getElementById("page-geral").style.display = "block";
  setTimeout(() => renderizarMapa(dados), 100);
  status.textContent = `✅ ${fmt(arquivoPrincipal.length)} registros carregados`;
}

// ── Helpers de data ───────────────────────────────────────────────────────────
function parseDateBR(val) {
  if (val === null || val === undefined || val === "") return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()));
  }
  if (typeof val === "number") {
    if (val <= 0) return null;
    const serial = val > 59 ? val - 1 : val;
    const msDesde1900 = (serial - 1) * 86400000;
    const base = Date.UTC(1900, 0, 1);
    const d = new Date(base + msDesde1900);
    if (isNaN(d.getTime())) return null;
    return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return null;
    const partes = s.split("/");
    if (partes.length >= 3) {
      const dia = parseInt(partes[0]), mes = parseInt(partes[1]), ano = parseInt(partes[2].split(" ")[0]);
      if (!isNaN(dia) && !isNaN(mes) && !isNaN(ano))
        return new Date(Date.UTC(ano, mes - 1, dia));
    }
    const iso = s.split(" ")[0].split("-");
    if (iso.length === 3) {
      const ano = parseInt(iso[0]), mes = parseInt(iso[1]), dia = parseInt(iso[2]);
      if (!isNaN(ano) && !isNaN(mes) && !isNaN(dia))
        return new Date(Date.UTC(ano, mes - 1, dia));
    }
  }
  return null;
}
function mesLabel(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}
const MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
function formatarMesLabel(yyyymm) { if (!yyyymm) return ""; const [y, m] = yyyymm.split("-"); return `${MESES_PT[parseInt(m) - 1]}/${y.slice(2)}`; }
function variacao(atual, anterior) { if (anterior === 0) return { texto:"—", classe:"" }; const diff = atual - anterior; return { texto:`${diff>=0?"▲":"▼"} ${Math.abs(((diff/anterior)*100)).toFixed(1)}%`, classe: diff>=0?"var-up":"var-down" }; }
function variacaoInversa(atual, anterior) { if (anterior === 0) return { texto:"—", classe:"" }; const diff = atual - anterior; return { texto:`${diff>=0?"▲":"▼"} ${Math.abs(((diff/anterior)*100)).toFixed(1)}%`, classe: diff>=0?"var-down":"var-up" }; }

// ── Cálculo dos dados ─────────────────────────────────────────────────────────
function calcularDados(relatorio, backoffice) {
  const isAtivo     = x => { const s = String(x["Status"] || "").toLowerCase().trim(); return s === "ativos" || s === "ativo" || s === "ativa"; };
  const isCancelado = x => { const s = String(x["Status"] || "").toLowerCase().trim(); return s === "cancelados" || s === "cancelado" || s === "cancelada"; };
  const total      = relatorio.length;
  const ativos     = relatorio.filter(isAtivo).length;
  const cancelados = relatorio.filter(isCancelado).length;

  const portRows       = relatorio.filter(r => ["True","Verdadeiro","true","verdadeiro"].includes(String(r["Portabilidades"] || "").trim()));
  const novaRows       = relatorio.filter(r => !["True","Verdadeiro","true","verdadeiro"].includes(String(r["Portabilidades"] || "").trim()));
  const portAtivos     = portRows.filter(isAtivo).length;
  const portCancel     = portRows.filter(isCancelado).length;
  const portAtivosRows = portRows.filter(isAtivo);
  const portAprovadas  = portAtivosRows.filter(r => String(r["Portabilidade"] || "").toUpperCase().includes("SUCESSO")).length;
  const portNegadas    = portAtivosRows.filter(r => String(r["Portabilidade"] || "").toLowerCase().includes("negada")).length;
  const portAndamento  = portAtivosRows.length - portAprovadas - portNegadas;
  const novaAtivos     = novaRows.filter(isAtivo).length;
  const novaCancel     = novaRows.filter(isCancelado).length;

  const normalizar = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();

  const esimRows   = relatorio.filter(r => isAtivo(r) && normalizar(r["Tipo de chip"]) === "esim");
  const fisicoRows = relatorio.filter(r => isAtivo(r) && normalizar(r["Tipo de chip"]) === "fisico");

  const bonifTotal  = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "BONIFICADA").length;
  const bonifNormal = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "NORMAL").length;

  const logisticaRows = relatorio.filter(r => {
    const v = String(r["Logistica"] || r["Logística"] || "").trim();
    return v !== "" && v !== "null" && v !== "-" && v !== "N/D" && v !== "0";
  });
  const logisticaTotal      = logisticaRows.length;
  const logisticaAtiva      = logisticaRows.filter(isAtivo).length;
  const logisticaCancelada  = logisticaRows.filter(isCancelado).length;
  const logisticaNaoAtivada = logisticaTotal - logisticaAtiva - logisticaCancelada;

  const excluirPgtoNorm = ["baixa manual", "pagamento com saldo", "-"];
  const contagemPgto = {};
  relatorio.filter(isAtivo).forEach(r => {
    const pgtoRaw = r["Forma de pagamento"];
    if (!pgtoRaw || String(pgtoRaw).trim() === "") return;
    const pgtoNorm = normalizar(pgtoRaw);
    if (excluirPgtoNorm.includes(pgtoNorm)) return;
    const pgtoF = String(pgtoRaw).trim().charAt(0).toUpperCase() + String(pgtoRaw).trim().slice(1).toLowerCase();
    contagemPgto[pgtoF] = (contagemPgto[pgtoF] || 0) + 1;
  });
  const totalPgto = Object.values(contagemPgto).reduce((a, b) => a + b, 0);

  const ativosTotal = relatorio.filter(isAtivo);
  const comRecorrencia = ativosTotal.filter(r => {
    const v = String(r["Recorrência"] || "");
    return v === "Cartão de crédito | recorrência cartão" || v === "Criptomoeda | recorrência cartão";
  }).length;
  const semRecorrencia = ativosTotal.length - comRecorrencia;

  const contagemOp = {};
  relatorio.forEach(r => { const op = r["Operadora"]; if (op) contagemOp[op] = (contagemOp[op] || 0) + 1; });

  const contagemPlanos = {};
  relatorio.filter(isAtivo).forEach(r => { const p = r["Plano"]; if (p) contagemPlanos[p] = (contagemPlanos[p] || 0) + 1; });
  const planos = Object.entries(contagemPlanos).map(([nome, qtd]) => ({ nome, total: qtd }));

  const ativacoesPorMes = {}, cancelamentosPorMes = {}, ativacoesPorDia = {}, cancelamentosPorDia = {};
  relatorio.forEach(r => {
    const dataAtiv = parseDateBR(r["Data de ativação"]);
    if (dataAtiv) {
      const mes = mesLabel(dataAtiv), dia = dataAtiv.getUTCDate();
      ativacoesPorMes[mes] = (ativacoesPorMes[mes] || 0) + 1;
      if (!ativacoesPorDia[mes]) ativacoesPorDia[mes] = {};
      ativacoesPorDia[mes][dia] = (ativacoesPorDia[mes][dia] || 0) + 1;
    }
    if (isCancelado(r)) {
      const dataCanc = parseDateBR(r["Data cancelado"]);
      if (dataCanc) {
        const mes = mesLabel(dataCanc), dia = dataCanc.getUTCDate();
        cancelamentosPorMes[mes] = (cancelamentosPorMes[mes] || 0) + 1;
        if (!cancelamentosPorDia[mes]) cancelamentosPorDia[mes] = {};
        cancelamentosPorDia[mes][dia] = (cancelamentosPorDia[mes][dia] || 0) + 1;
      }
    }
  });

  const todosMeses = Array.from(
    new Set([...Object.keys(ativacoesPorMes), ...Object.keys(cancelamentosPorMes)])
  ).sort((a, b) => {
    const [ay, am] = a.split("-").map(Number);
    const [by, bm] = b.split("-").map(Number);
    return ay !== by ? ay - by : am - bm;
  });
  const labelsTimeline   = todosMeses.map(m => formatarMesLabel(m));
  const ativacoesArr     = todosMeses.map(m => ativacoesPorMes[m] || 0);
  const cancelamentosArr = todosMeses.map(m => cancelamentosPorMes[m] || 0);
  const churnArr = todosMeses.map(m => {
    const canc = cancelamentosPorMes[m] || 0;
    return ativos > 0 ? parseFloat(((canc / ativos) * 100).toFixed(1)) : 0;
  });

  let somaVida = 0, countVida = 0;
  relatorio.forEach(r => {
    if (isCancelado(r)) {
      const dataAtiv = parseDateBR(r["Data de ativação"]);
      const dataCanc = parseDateBR(r["Data cancelado"]);
      if (dataAtiv && dataCanc) {
        const dias = Math.round((dataCanc.getTime() - dataAtiv.getTime()) / (1000 * 60 * 60 * 24));
        if (dias >= 0) { somaVida += dias; countVida++; }
      }
    }
  });
  const tempoMedioVida = countVida > 0 ? Math.round(somaVida / countVida) : 0;

  const mesesComAtivacoes     = todosMeses.filter(m => (ativacoesPorMes[m]    || 0) > 0);
  const mesesComCancelamentos = todosMeses.filter(m => (cancelamentosPorMes[m] || 0) > 0);
  const ultMesAtiv = mesesComAtivacoes[mesesComAtivacoes.length - 1]         || null;
  const penMesAtiv = mesesComAtivacoes[mesesComAtivacoes.length - 2]         || null;
  const ultMesCanc = mesesComCancelamentos[mesesComCancelamentos.length - 1] || null;
  const penMesCanc = mesesComCancelamentos[mesesComCancelamentos.length - 2] || null;

  const kpi = {
    ativacoes:     { atual: ativacoesPorMes[ultMesAtiv]    || 0, anterior: ativacoesPorMes[penMesAtiv]    || 0 },
    cancelamentos: { atual: cancelamentosPorMes[ultMesCanc] || 0, anterior: cancelamentosPorMes[penMesCanc] || 0 },
    churn: {
      atual:    ativos > 0 ? parseFloat(((cancelamentosPorMes[ultMesCanc] || 0) / ativos * 100).toFixed(2)) : 0,
      anterior: ativos > 0 ? parseFloat(((cancelamentosPorMes[penMesCanc] || 0) / ativos * 100).toFixed(2)) : 0,
    },
    mesAtual:    formatarMesLabel(ultMesAtiv || ""),
    mesAnterior: formatarMesLabel(penMesAtiv || ""),
  };

  const contagemEstados = {};
  const cidadesPorEstado = {};
  relatorio.forEach(r => {
    const uf = String(r["Estado"] || "").trim().toUpperCase();
    if (uf) {
      contagemEstados[uf] = (contagemEstados[uf] || 0) + 1;
      const cidade = String(r["Cidade"] || r["Municipio"] || r["Município"] || "").trim();
      if (cidade) {
        if (!cidadesPorEstado[uf]) cidadesPorEstado[uf] = {};
        cidadesPorEstado[uf][cidade] = (cidadesPorEstado[uf][cidade] || 0) + 1;
      }
    }
  });

  const ativacoesPorDoc = {}, numTelCount = {}, telDocMap = {}, dadosPorDoc = {}, statusPortMap = {}, operadoraStats = {};
  const relatorioClientes = backoffice && backoffice.size > 0
    ? relatorio.filter(r => {
        const chip = String(r["Chip"] || r["Numero"] || "").trim().split(".")[0];
        return !chip || backoffice.has(chip);
      })
    : relatorio;

  relatorioClientes.forEach(r => {
    const doc       = String(r["CPF"] || r["Clientes"] || "—").trim();
    const nome      = String(r["Clientes"] || "—");
    const tel       = String(r["Numero De Origem"] || r["Numero de origem"] || "").trim();
    const cancelada = isCancelado(r);
    const dataAtiv  = parseDateBR(r["Data de ativação"]);
    if (!ativacoesPorDoc[doc]) ativacoesPorDoc[doc] = { doc, nome, total: 0 };
    ativacoesPorDoc[doc].total++;
    if (!dadosPorDoc[doc]) dadosPorDoc[doc] = { datas: [], canceladas: 0 };
    if (dataAtiv) dadosPorDoc[doc].datas.push(dataAtiv);
    if (cancelada) dadosPorDoc[doc].canceladas++;
    if (tel) {
      numTelCount[tel] = (numTelCount[tel] || 0) + 1;
      if (!telDocMap[tel]) telDocMap[tel] = new Set();
      telDocMap[tel].add(doc);
    }
    const raw = r["Portabilidade"];
    if (raw) {
      let s = String(raw).trim();
      if (s.startsWith("Portabilidade negada")) s = "Portabilidade negada";
      if (s.toUpperCase() === "SUCESSO" || s.trimEnd() === "Sucesso") s = "Sucesso";
      statusPortMap[s] = (statusPortMap[s] || 0) + 1;
    }
    const op = r["Operadora"];
    if (op) {
      if (!operadoraStats[op]) operadoraStats[op] = { sucesso: 0, negada: 0, andamento: 0, total: 0 };
      const s = String(r["Portabilidade"] || "").trim();
      operadoraStats[op].total++;
      if (s.toUpperCase() === "SUCESSO" || s.trimEnd() === "Sucesso") operadoraStats[op].sucesso++;
      else if (s.startsWith("Portabilidade negada")) operadoraStats[op].negada++;
      else operadoraStats[op].andamento++;
    }
  });

  const telsRepetidos = new Set(Object.keys(numTelCount).filter(t => numTelCount[t] > 1));
  const repetidasPorDoc = {};
  telsRepetidos.forEach(tel => {
    if (telDocMap[tel]) telDocMap[tel].forEach(doc => { repetidasPorDoc[doc] = (repetidasPorDoc[doc] || 0) + numTelCount[tel]; });
  });
  const docList = Object.values(ativacoesPorDoc).map(d => ({ ...d, repetidas: repetidasPorDoc[d.doc] || 0 })).sort((a, b) => b.total - a.total);
  const telsDetalhes = Object.entries(numTelCount).filter(([, c]) => c > 1).sort((a, b) => b[1] - a[1]).slice(0, 50).map(([tel, count]) => ({ tel, count, docs: [...(telDocMap[tel] || [])] }));

  const alertaRapido = [];
  Object.entries(dadosPorDoc).forEach(([doc, info]) => {
    if (info.datas.length < 3) return;
    const sorted = info.datas.slice().sort((a, b) => a - b);
    for (let i = 0; i <= sorted.length - 3; i++) {
      const dias = Math.round((sorted[i + 2] - sorted[i]) / 86400000);
      if (dias <= 30) { const cli = ativacoesPorDoc[doc]; alertaRapido.push({ doc, nome: cli?.nome || "—", total: cli?.total || 0, diasEntre3: dias }); break; }
    }
  });
  alertaRapido.sort((a, b) => a.diasEntre3 - b.diasEntre3);

  const telPorMultiplosClientes = Object.entries(telDocMap).filter(([, docs]) => docs.size > 1).map(([tel, docs]) => ({ tel, totalLinhas: numTelCount[tel], totalClientes: docs.size, docs: [...docs] })).sort((a, b) => b.totalClientes - a.totalClientes);

  const altoCancelamento = [];
  Object.entries(dadosPorDoc).forEach(([doc, info]) => {
    const cli = ativacoesPorDoc[doc];
    if (!cli || cli.total < 2) return;
    const ratio = info.canceladas / cli.total;
    if (ratio > 0.5) altoCancelamento.push({ doc, nome: cli.nome, total: cli.total, canceladas: info.canceladas, pctCancel: (ratio * 100).toFixed(0) });
  });
  altoCancelamento.sort((a, b) => b.pctCancel - a.pctCancel);

  const operadoraList = Object.entries(operadoraStats).map(([op, s]) => ({ op, ...s, taxaSucesso: s.total > 0 ? ((s.sucesso / s.total) * 100).toFixed(1) : 0, taxaNegada: s.total > 0 ? ((s.negada / s.total) * 100).toFixed(1) : 0 })).filter(o => o.total >= 5).sort((a, b) => b.total - a.total);

  const VALOR_BONIF      = 54.90;
  const bonifRows        = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "BONIFICADA");
  const totalBonificados = bonifRows.length;
  let campComRecarga = 0, campSemRecarga = 0, campViaBackoffice = 0;
  bonifRows.forEach(r => {
    const temRecarga = r["Data de ultima recarga"] && String(r["Data de ultima recarga"]).trim() !== "";
    const faturas    = parseInt(r["Qº de recargas"] || 0);
    if (temRecarga || faturas > 0) { campComRecarga++; if (faturas > 0) campViaBackoffice++; }
    else campSemRecarga++;
  });

  const campanhas = {
    totalBonificados, comRecarga: campComRecarga, semRecarga: campSemRecarga, viaBackoffice: campViaBackoffice,
    ganho: campComRecarga * VALOR_BONIF, perda: campSemRecarga * VALOR_BONIF,
    liquido: (campComRecarga - campSemRecarga) * VALOR_BONIF, valorBonif: VALOR_BONIF,
    pctConversao: totalBonificados > 0 ? ((campComRecarga / totalBonificados) * 100).toFixed(1) : 0,
  };

  return {
    total, ativos, cancelados,
    portabilidades: { total: portRows.length, ativos: portAtivos, cancelados: portCancel, aprovadas: portAprovadas, negadas: portNegadas, andamento: portAndamento, totalAtivos: portAtivosRows.length },
    novasLinhas:    { total: novaRows.length, ativos: novaAtivos, cancelados: novaCancel },
    chip:           { total: esimRows.length + fisicoRows.length, esim: esimRows.length, fisico: fisicoRows.length },
    pagamento:      { total: totalPgto, formas: contagemPgto },
    bonificados:    { total: bonifTotal, normal: bonifNormal },
    logistica:      { total: logisticaTotal, ativa: logisticaAtiva, naoAtivada: logisticaNaoAtivada, cancelada: logisticaCancelada },
    recorrencia:    { total: ativosTotal.length, com: comRecorrencia, sem: semRecorrencia },
    operadoras: contagemOp, estados: contagemEstados, cidadesPorEstado, bugs: { total: 0 },
    tempoMedioVida, kpi, planos,
    timeline: { todosMeses, labels: labelsTimeline, ativacoes: ativacoesArr, cancelamentos: cancelamentosArr, churn: churnArr, ativacoesPorMes, cancelamentosPorMes, ativacoesPorDia, cancelamentosPorDia },
    clientes: { docList, statusPortMap, telsDetalhes, alertaRapido, telPorMultiplosClientes, altoCancelamento, operadoraList },
    campanhas,
  };
}

// ── Filtro global ─────────────────────────────────────────────────────────────
function construirFiltroGlobal(d) {
  const container = document.getElementById("filtroGlobal");
  if (!container) return;
  container.innerHTML = "";

  const periodoWrap = document.createElement("div");
  periodoWrap.style.cssText = "display:flex;flex-direction:column;gap:6px;margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--border);width:100%;";
  periodoWrap.innerHTML = `
    <span style="font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;">Período personalizado</span>
    <div style="display:flex;gap:6px;align-items:center;">
      <input type="date" id="filtroDataInicio" style="flex:1;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:5px 8px;font-size:12px;color:var(--text);outline:none;">
      <span style="color:var(--text-muted);font-size:12px;">até</span>
      <input type="date" id="filtroDataFim" style="flex:1;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:5px 8px;font-size:12px;color:var(--text);outline:none;">
      <button class="filtro-btn" id="btnAplicarPeriodo" style="white-space:nowrap;">✓ Aplicar</button>
    </div>
  `;
  container.appendChild(periodoWrap);

  document.getElementById("btnAplicarPeriodo").addEventListener("click", () => {
    const inicio = document.getElementById("filtroDataInicio").value;
    const fim    = document.getElementById("filtroDataFim").value;
    if (!inicio || !fim) return;
    const dtInicio = new Date(inicio + "T00:00:00Z");
    const dtFim    = new Date(fim    + "T23:59:59Z");
    const relFiltrado = arquivoPrincipal.filter(r => {
      const dataAtiv = parseDateBR(r["Data de ativação"]);
      const dataCanc = parseDateBR(r["Data cancelado"]);
      return (dataAtiv && dataAtiv >= dtInicio && dataAtiv <= dtFim) || (dataCanc && dataCanc >= dtInicio && dataCanc <= dtFim);
    });
    document.querySelectorAll("#filtroGlobal .filtro-btn").forEach(b => b.classList.remove("active"));
    document.getElementById("btnAplicarPeriodo").classList.add("active");
    const dadosFiltrados = calcularDados(relFiltrado, arquivoBackoffice);
    dadosFiltrados._mesSelecionado = "todos";
    const tl = dadosFiltrados.timeline;
    const mesFim = mesLabel(new Date(fim + "T00:00:00Z"));
    const dtFimObj   = new Date(fim + "T00:00:00Z");
    const ultimoDia  = new Date(Date.UTC(dtFimObj.getUTCFullYear(), dtFimObj.getUTCMonth() + 1, 0)).getUTCDate();
    const mesFimCompleto = dtFimObj.getUTCDate() >= ultimoDia - 2;
    const mesesAtiv = mesFimCompleto ? tl.todosMeses.filter(m => (tl.ativacoesPorMes[m]    || 0) > 0) : tl.todosMeses.filter(m => (tl.ativacoesPorMes[m]    || 0) > 0 && m < mesFim);
    const mesesCanc = mesFimCompleto ? tl.todosMeses.filter(m => (tl.cancelamentosPorMes[m] || 0) > 0) : tl.todosMeses.filter(m => (tl.cancelamentosPorMes[m] || 0) > 0 && m < mesFim);
    const ultAtiv = mesesAtiv[mesesAtiv.length - 1] || null;
    const penAtiv = mesesAtiv[mesesAtiv.length - 2] || null;
    const ultCanc = mesesCanc[mesesCanc.length - 1] || null;
    const penCanc = mesesCanc[mesesCanc.length - 2] || null;
    dadosFiltrados.kpi.ativacoes     = { atual: tl.ativacoesPorMes[ultAtiv]    || 0, anterior: tl.ativacoesPorMes[penAtiv]    || 0 };
    dadosFiltrados.kpi.cancelamentos = { atual: tl.cancelamentosPorMes[ultCanc] || 0, anterior: tl.cancelamentosPorMes[penCanc] || 0 };
    dadosFiltrados.kpi.mesAtual      = formatarMesLabel(ultAtiv || "");
    dadosFiltrados.kpi.mesAnterior   = formatarMesLabel(penAtiv || "");
    const atv = dadosFiltrados.ativos;
    dadosFiltrados.kpi.churn = {
      atual:    atv > 0 ? parseFloat((((tl.cancelamentosPorMes[ultCanc] || 0) / atv) * 100).toFixed(2)) : 0,
      anterior: atv > 0 ? parseFloat((((tl.cancelamentosPorMes[penCanc] || 0) / atv) * 100).toFixed(2)) : 0,
    };
    dadosGlobais = dadosFiltrados;
    renderDash(dadosFiltrados);
    setTimeout(() => renderizarMapa(dadosFiltrados), 100);
    document.getElementById("filtroGlobal").style.display = "none";
  });

  const labelMeses = document.createElement("span");
  labelMeses.style.cssText = "font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;display:block;";
  labelMeses.textContent = "Por mês";
  container.appendChild(labelMeses);

  const btnTodos = document.createElement("button");
  btnTodos.textContent = "Todos os meses"; btnTodos.className = "filtro-btn active"; btnTodos.dataset.mes = "todos";
  btnTodos.addEventListener("click", () => { aplicarFiltroGlobal("todos", d); document.getElementById("filtroGlobal").style.display = "none"; });
  container.appendChild(btnTodos);

  d.timeline.todosMeses.forEach((mes, i) => {
    const btn = document.createElement("button");
    btn.textContent = d.timeline.labels[i]; btn.className = "filtro-btn"; btn.dataset.mes = mes;
    btn.addEventListener("click", () => { aplicarFiltroGlobal(mes, d); document.getElementById("filtroGlobal").style.display = "none"; });
    container.appendChild(btn);
  });
}

function aplicarFiltroGlobal(mesSelecionado, d) {
  document.querySelectorAll("#filtroGlobal .filtro-btn").forEach(b => b.classList.toggle("active", b.dataset.mes === mesSelecionado));
  const relFiltrado = mesSelecionado === "todos"
    ? arquivoPrincipal
    : arquivoPrincipal.filter(r => {
        const dataAtiv = parseDateBR(r["Data de ativação"]);
        const dataCanc = parseDateBR(r["Data cancelado"]);
        const mesAtiv  = dataAtiv ? mesLabel(dataAtiv) : null;
        const mesCanc  = dataCanc ? mesLabel(dataCanc) : null;
        return mesAtiv === mesSelecionado || mesCanc === mesSelecionado;
      });
  const dadosFiltrados = calcularDados(relFiltrado, arquivoBackoffice);
  dadosFiltrados._mesSelecionado = mesSelecionado;
  if (mesSelecionado !== "todos") {
    const ativAtual  = d.timeline.ativacoesPorMes[mesSelecionado]    || 0;
    const cancAtual  = d.timeline.cancelamentosPorMes[mesSelecionado] || 0;
    const mesesComAtiv = d.timeline.todosMeses.filter(m => m < mesSelecionado && (d.timeline.ativacoesPorMes[m]    || 0) > 0);
    const mesesComCanc = d.timeline.todosMeses.filter(m => m < mesSelecionado && (d.timeline.cancelamentosPorMes[m] || 0) > 0);
    const mesAntAtiv   = mesesComAtiv[mesesComAtiv.length - 1] || null;
    const mesAntCanc   = mesesComCanc[mesesComCanc.length - 1] || null;
    const ativAnterior = mesAntAtiv ? (d.timeline.ativacoesPorMes[mesAntAtiv]    || 0) : 0;
    const cancAnterior = mesAntCanc ? (d.timeline.cancelamentosPorMes[mesAntCanc] || 0) : 0;
    dadosFiltrados.kpi.ativacoes     = { atual: ativAtual,  anterior: ativAnterior };
    dadosFiltrados.kpi.cancelamentos = { atual: cancAtual,  anterior: cancAnterior };
    dadosFiltrados.kpi.mesAtual      = formatarMesLabel(mesSelecionado);
    dadosFiltrados.kpi.mesAnterior   = formatarMesLabel(mesAntAtiv || "");
    const atv = dadosFiltrados.ativos;
    dadosFiltrados.kpi.churn = {
      atual:    atv > 0 ? parseFloat(((cancAtual    / atv) * 100).toFixed(2)) : 0,
      anterior: atv > 0 ? parseFloat(((cancAnterior / atv) * 100).toFixed(2)) : 0,
    };
  }
  dadosGlobais = dadosFiltrados;
  renderDash(dadosFiltrados);
  setTimeout(() => renderizarMapa(dadosFiltrados), 100);
}

function toggleFiltroGlobal() {
  const el = document.getElementById("filtroGlobal");
  el.style.display = el.style.display === "none" ? "flex" : "none";
}
document.addEventListener("click", (e) => {
  const btn = document.getElementById("btnFiltroGlobal");
  const dropdown = document.getElementById("filtroGlobal");
  if (!btn || !dropdown) return;
  if (!btn.contains(e.target) && !dropdown.contains(e.target)) dropdown.style.display = "none";
});

// ── Init ──────────────────────────────────────────────────────────────────────
setupTema();
setupDropZone();
setupNav();
