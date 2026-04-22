// DASH.JS — lê o Excel automaticamente via SheetJS
// ─────────────────────────────────────────────

const pct = (parte, total) => total ? ((parte / total) * 100).toFixed(1) + "%" : "0%";
const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
const fmt = (n) => Number(n).toLocaleString("pt-BR");

let chartPort = null, chartNova = null, chartTimeline = null;
let chartChip = null, chartPgto = null, chartRecorrencia = null, chartOperadoras = null;
let dadosGlobais = null;
let labels, ativacoes, cancelamentos, churn;

const nomesEstados = {
  AC:"Acre", AL:"Alagoas", AM:"Amazonas", AP:"Amapá", BA:"Bahia",
  CE:"Ceará", DF:"Distrito Federal", ES:"Espírito Santo", GO:"Goiás",
  MA:"Maranhão", MG:"Minas Gerais", MS:"Mato Grosso do Sul", MT:"Mato Grosso",
  PA:"Pará", PB:"Paraíba", PE:"Pernambuco", PI:"Piauí", PR:"Paraná",
  RJ:"Rio de Janeiro", RN:"Rio Grande do Norte", RO:"Rondônia", RR:"Roraima",
  RS:"Rio Grande do Sul", SC:"Santa Catarina", SE:"Sergipe", SP:"São Paulo",
  TO:"Tocantins",
};

// ── Tema claro/escuro ─────────────────────────────────────────────────────────
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

// ── Navegação lateral ─────────────────────────────────────────────────────────
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

// ── Upload de dois arquivos ───────────────────────────────────────────────────
let arquivoPrincipal  = null;
let arquivoBackoffice = null;

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

// ── CORREÇÃO PRINCIPAL: lerExcel agora passa cellDates:false para receber
//    números seriais brutos, que parseDateBR converte corretamente em UTC ──────
function lerExcel(file, callback) {
  const reader = new FileReader();
  reader.onload = (e) => {
    setTimeout(() => {
      try {
        // cellDates:false → datas chegam como número serial do Excel
        // raw:true        → garante que os valores não sejam formatados como string
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
    // defval:null garante que colunas ausentes virem null e não undefined
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
    // raw:false no backoffice para que os chips venham como string formatada
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
  construirFiltroMeses(dados);

  document.getElementById("uploadArea").style.display = "none";
  document.getElementById("page-geral").style.display = "block";
  setTimeout(() => renderizarMapa(dados), 100);

  status.textContent = `✅ ${fmt(arquivoPrincipal.length)} registros carregados`;
}

// ── Cálculo dos dados ─────────────────────────────────────────────────────────
function calcularDados(relatorio, backoffice) {

  const isAtivo     = x => { const s = String(x["Status"] || "").toLowerCase().trim(); return s === "ativos" || s === "ativo" || s === "ativa"; };
  const isCancelado = x => { const s = String(x["Status"] || "").toLowerCase().trim(); return s === "cancelados" || s === "cancelado" || s === "cancelada"; };
  const total      = relatorio.length;
  const ativos     = relatorio.filter(isAtivo).length;
  const cancelados = relatorio.filter(isCancelado).length;

  const portRows   = relatorio.filter(r => ["True","Verdadeiro","true","verdadeiro"].includes(String(r["Portabilidades"] || "").trim()));
  const novaRows   = relatorio.filter(r => !["True","Verdadeiro","true","verdadeiro"].includes(String(r["Portabilidades"] || "").trim()));
  const portAtivos = portRows.filter(isAtivo).length;
  const portCancel = portRows.filter(isCancelado).length;
  const novaAtivos = novaRows.filter(isAtivo).length;
  const novaCancel = novaRows.filter(isCancelado).length;

  // ── CORREÇÃO: normalize NFD para remover acentos antes de comparar ──────────
  const normalizar = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();

  const esimRows   = relatorio.filter(r => normalizar(r["Tipo de chip"]) === "esim");
  const fisicoRows = relatorio.filter(r => normalizar(r["Tipo de chip"]) === "fisico");

  const bonifTotal  = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "BONIFICADA").length;
  const bonifNormal = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "NORMAL").length;

  // Logística: registros onde a coluna "Logistica" tem valor preenchido (não vazio/null)
 const logisticaTotal = relatorio.filter(r => {
  const v = String(r["Logistica"] || r["Logística"] || "").trim();
  return v !== "" && v !== "null" && v !== "-" && v !== "N/D" && v !== "0";
}).length;

  // ── CORREÇÃO: normalizar forma de pagamento antes de excluir ────────────────
  const excluirPgtoNorm = ["baixa manual", "pagamento com saldo", "-", "bonificacao"];
  const contagemPgto = {};
  relatorio.forEach(r => {
    const pgtoRaw = r["Forma de pagamento"];
    if (!pgtoRaw || String(pgtoRaw).trim() === "") return;
    const pgtoNorm = normalizar(pgtoRaw);
    if (excluirPgtoNorm.includes(pgtoNorm)) return;
    const pgtoF = String(pgtoRaw).trim().charAt(0).toUpperCase() + String(pgtoRaw).trim().slice(1).toLowerCase();
    contagemPgto[pgtoF] = (contagemPgto[pgtoF] || 0) + 1;
  });
  const totalPgto = Object.values(contagemPgto).reduce((a, b) => a + b, 0);

  // ── CORREÇÃO: cartão com ou sem acento ──────────────────────────────────────
  const comRecorrencia = relatorio.filter(r => normalizar(r["Forma de pagamento"]).includes("cartao")).length;
  const semRecorrencia = total - comRecorrencia;

  const contagemOp = {};
  relatorio.forEach(r => { const op = r["Operadora"]; if (op) contagemOp[op] = (contagemOp[op] || 0) + 1; });

  const contagemPlanos = {};
  relatorio.forEach(r => { const p = r["Plano"]; if (p) contagemPlanos[p] = (contagemPlanos[p] || 0) + 1; });
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

  // ── CORREÇÃO: ordenação de meses agora é cronológica garantida ───────────────
  const todosMeses     = Array.from(new Set([...Object.keys(ativacoesPorMes), ...Object.keys(cancelamentosPorMes)])).sort();
  const labelsTimeline = todosMeses.map(m => formatarMesLabel(m));
  const ativacoesArr   = todosMeses.map(m => ativacoesPorMes[m] || 0);
  const cancelamentosArr = todosMeses.map(m => cancelamentosPorMes[m] || 0);
  const churnArr = todosMeses.map(m => {
    const atv = ativacoesPorMes[m] || 0, canc = cancelamentosPorMes[m] || 0;
    return atv > 0 ? parseFloat(((canc / atv) * 100).toFixed(1)) : 0;
  });

  // ── Tempo médio de vida ───────────────────────────────────────────────────────
  let somaVida = 0, countVida = 0;
  relatorio.forEach(r => {
    if (isCancelado(r)) {
      const dataAtiv = parseDateBR(r["Data de ativação"]);
      const dataCanc = parseDateBR(r["Data cancelado"]);
      if (dataAtiv && dataCanc) {
        // usar getTime() para ambos; ambos são em UTC por parseDateBR
        const dias = Math.round((dataCanc.getTime() - dataAtiv.getTime()) / (1000 * 60 * 60 * 24));
        if (dias >= 0) { somaVida += dias; countVida++; }
      }
    }
  });
  const tempoMedioVida = countVida > 0 ? Math.round(somaVida / countVida) : 0;

  const ultMes = todosMeses[todosMeses.length - 1];
  const penMes = todosMeses[todosMeses.length - 2];
  const kpi = {
    ativacoes:     { atual: ativacoesPorMes[ultMes] || 0,     anterior: ativacoesPorMes[penMes] || 0 },
    cancelamentos: { atual: cancelamentosPorMes[ultMes] || 0, anterior: cancelamentosPorMes[penMes] || 0 },
    churn:         { atual: churnArr[todosMeses.length - 1] || 0, anterior: churnArr[todosMeses.length - 2] || 0 },
    mesAtual: formatarMesLabel(ultMes || ""), mesAnterior: formatarMesLabel(penMes || ""),
  };

  const contagemEstados = {};
  relatorio.forEach(r => {
    const uf = String(r["Estado"] || "").trim().toUpperCase();
    if (uf) contagemEstados[uf] = (contagemEstados[uf] || 0) + 1;
  });

  // ── Dashboard de Clientes (usa backoffice para filtrar duplicatas) ────────────
  const ativacoesPorDoc = {}, numTelCount = {}, telDocMap = {}, dadosPorDoc = {}, statusPortMap = {}, operadoraStats = {};

  // ── CORREÇÃO: filtrar pela planilha de backoffice (chips únicos) ─────────────
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

  // ── Dashboard de Campanhas ────────────────────────────────────────────────────
  const VALOR_BONIF      = 54.90;
  const bonifRows        = relatorio.filter(r => String(r["Tipo"] || "").trim().toUpperCase() === "BONIFICADA");
  const totalBonificados = bonifRows.length;
  let campComRecarga = 0, campSemRecarga = 0, campViaBackoffice = 0;

  bonifRows.forEach(r => {
    const temRecarga = r["Data de ultima recarga"] && String(r["Data de ultima recarga"]).trim() !== "";
    const faturas    = parseInt(r["Qº de recargas"] || 0);
    if (temRecarga || faturas > 0) {
      campComRecarga++;
      if (faturas > 0) campViaBackoffice++;
    } else {
      campSemRecarga++;
    }
  });

  const campanhas = {
    totalBonificados,
    comRecarga:    campComRecarga,
    semRecarga:    campSemRecarga,
    viaBackoffice: campViaBackoffice,
    ganho:         campComRecarga * VALOR_BONIF,
    perda:         campSemRecarga * VALOR_BONIF,
    liquido:       (campComRecarga - campSemRecarga) * VALOR_BONIF,
    valorBonif:    VALOR_BONIF,
    pctConversao:  totalBonificados > 0 ? ((campComRecarga / totalBonificados) * 100).toFixed(1) : 0,
  };

  return {
    total, ativos, cancelados,
    portabilidades: { total: portRows.length, ativos: portAtivos, cancelados: portCancel },
    novasLinhas:    { total: novaRows.length, ativos: novaAtivos, cancelados: novaCancel },
    chip:           { total: esimRows.length + fisicoRows.length, esim: esimRows.length, fisico: fisicoRows.length },
    pagamento:      { total: totalPgto, formas: contagemPgto },
    bonificados:    { total: bonifTotal, normal: bonifNormal },
    logistica:      { total: logisticaTotal },
    recorrencia:    { com: comRecorrencia, sem: semRecorrencia },
    operadoras: contagemOp, estados: contagemEstados, bugs: { total: 0 },
    tempoMedioVida, kpi, planos,
    timeline: { todosMeses, labels: labelsTimeline, ativacoes: ativacoesArr, cancelamentos: cancelamentosArr, churn: churnArr, ativacoesPorMes, cancelamentosPorMes, ativacoesPorDia, cancelamentosPorDia },
    clientes: { docList, statusPortMap, telsDetalhes, alertaRapido, telPorMultiplosClientes, altoCancelamento, operadoraList },
    campanhas,
  };
}

// ── Helpers de data ───────────────────────────────────────────────────────────
// CORREÇÃO PRINCIPAL: parseDateBR sempre retorna Date em UTC para evitar
// deslocamento de fuso horário. Número serial do Excel → UTC correto.
function parseDateBR(val) {
  if (val === null || val === undefined || val === "") return null;

  // Já é um objeto Date (quando cellDates: true no SheetJS)
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    // Recriar como UTC para evitar deslocamento
    return new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()));
  }

  // Número serial do Excel
  // CORREÇÃO: usar Math.floor (não Math.round) e recriar como UTC
  if (typeof val === "number") {
    if (val <= 0) return null;
    // Excel serial: dias desde 1900-01-00 (com bug do 1900-02-29)
    const serial = val > 59 ? val - 1 : val; // corrige bug do Lotus 1-2-3
    const msDesde1900 = (serial - 1) * 86400000;
    const base = Date.UTC(1900, 0, 1); // 1900-01-01 UTC
    const d = new Date(base + msDesde1900);
    if (isNaN(d.getTime())) return null;
    // Recriar como UTC puro sem hora
    return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }

  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return null;
    // dd/mm/yyyy ou dd/mm/yyyy hh:mm
    const partes = s.split("/");
    if (partes.length >= 3) {
      const dia = parseInt(partes[0]), mes = parseInt(partes[1]), ano = parseInt(partes[2].split(" ")[0]);
      if (!isNaN(dia) && !isNaN(mes) && !isNaN(ano))
        return new Date(Date.UTC(ano, mes - 1, dia));
    }
    // yyyy-mm-dd ou yyyy-mm-dd hh:mm:ss
    const iso = s.split(" ")[0].split("-");
    if (iso.length === 3) {
      const ano = parseInt(iso[0]), mes = parseInt(iso[1]), dia = parseInt(iso[2]);
      if (!isNaN(ano) && !isNaN(mes) && !isNaN(dia))
        return new Date(Date.UTC(ano, mes - 1, dia));
    }
  }
  return null;
}

// ── CORREÇÃO: mesLabel usa getUTCFullYear/getUTCMonth para datas UTC ──────────
function mesLabel(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

const MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
function formatarMesLabel(yyyymm) { if (!yyyymm) return ""; const [y, m] = yyyymm.split("-"); return `${MESES_PT[parseInt(m) - 1]}/${y.slice(2)}`; }
function variacao(atual, anterior) { if (anterior === 0) return { texto:"—", classe:"" }; const diff = atual - anterior; return { texto:`${diff>=0?"▲":"▼"} ${Math.abs(((diff/anterior)*100)).toFixed(1)}%`, classe: diff>=0?"var-up":"var-down" }; }
function variacaoInversa(atual, anterior) { if (anterior === 0) return { texto:"—", classe:"" }; const diff = atual - anterior; return { texto:`${diff>=0?"▲":"▼"} ${Math.abs(((diff/anterior)*100)).toFixed(1)}%`, classe: diff>=0?"var-down":"var-up" }; }

// ── Renderização ──────────────────────────────────────────────────────────────
function renderDash(d) {
  preencherCards(d);
  preencherKPI(d);
  renderizarPlanos(d);
  renderizarGraficos(d);
  renderizarClientes(d);
  renderizarCampanhas(d);
}

function preencherCards(d) {
  set("total", fmt(d.total)); set("ativos", fmt(d.ativos)); set("ativos-pct", pct(d.ativos, d.total) + " do total");
  set("cancelados", fmt(d.cancelados)); set("cancelados-pct", pct(d.cancelados, d.total) + " do total");
  set("portabilidades", fmt(d.portabilidades.total)); set("port-pct", pct(d.portabilidades.total, d.total) + " do total");
  set("novas", fmt(d.novasLinhas.total)); set("novas-pct", pct(d.novasLinhas.total, d.total) + " do total");
  set("bonificados", fmt(d.bonificados.total)); set("bonif-pct", pct(d.bonificados.total, d.total) + " do total");
  set("logistica", fmt(d.logistica.total)); set("logistica-pct", pct(d.logistica.total, d.total) + " do total");
  set("tempo-vida", d.tempoMedioVida + " dias");
  set("port-total", fmt(d.portabilidades.total)); set("port-ativos", fmt(d.portabilidades.ativos)); set("port-ativos-pct", pct(d.portabilidades.ativos, d.portabilidades.total));
  set("port-cancelados", fmt(d.portabilidades.cancelados)); set("port-cancelados-pct", pct(d.portabilidades.cancelados, d.portabilidades.total));
  set("novas-total", fmt(d.novasLinhas.total)); set("novas-ativos", fmt(d.novasLinhas.ativos)); set("novas-ativos-pct", pct(d.novasLinhas.ativos, d.novasLinhas.total));
  set("novas-cancelados", fmt(d.novasLinhas.cancelados)); set("novas-cancelados-pct", pct(d.novasLinhas.cancelados, d.novasLinhas.total));
  set("chip-total", fmt(d.chip.total)); set("chip-esim", fmt(d.chip.esim)); set("chip-esim-pct", pct(d.chip.esim, d.chip.total));
  set("chip-fisico", fmt(d.chip.fisico)); set("chip-fisico-pct", pct(d.chip.fisico, d.chip.total));
  set("pgto-total", fmt(d.pagamento.total));
  const tabelaPgto = document.getElementById("pgto-tabela");
  if (tabelaPgto) tabelaPgto.innerHTML = Object.entries(d.pagamento.formas).sort((a,b)=>b[1]-a[1]).map(([nome,qtd])=>`<tr><td>${nome}</td><td>${fmt(qtd)}</td><td>${pct(qtd,d.pagamento.total)}</td></tr>`).join("");
  set("bonif-total", fmt(d.bonificados.total)); set("bonif-normal", fmt(d.bonificados.normal)); set("pagamento-normal", fmt(d.bonificados.normal));
  const totalRec = d.recorrencia.com + d.recorrencia.sem;
  set("rec-com", fmt(d.recorrencia.com)); set("rec-com-pct", pct(d.recorrencia.com, totalRec));
  set("rec-sem", fmt(d.recorrencia.sem)); set("rec-sem-pct", pct(d.recorrencia.sem, totalRec));
}

function preencherKPI(d) {
  const k = d.kpi;
  const varAtiv  = variacao(k.ativacoes.atual, k.ativacoes.anterior);
  const varCanc  = variacaoInversa(k.cancelamentos.atual, k.cancelamentos.anterior);
  const varChurn = variacaoInversa(k.churn.atual, k.churn.anterior);
  set("kpi-mes", `${k.mesAnterior} → ${k.mesAtual}`);
  set("kpi-ativ-atual", fmt(k.ativacoes.atual)); set("kpi-ativ-ant", fmt(k.ativacoes.anterior));
  const elVarAtiv = document.getElementById("kpi-ativ-var");
  if (elVarAtiv) { elVarAtiv.textContent = varAtiv.texto; elVarAtiv.className = "kpi-var " + varAtiv.classe; }
  set("kpi-canc-atual", fmt(k.cancelamentos.atual)); set("kpi-canc-ant", fmt(k.cancelamentos.anterior));
  const elVarCanc = document.getElementById("kpi-canc-var");
  if (elVarCanc) { elVarCanc.textContent = varCanc.texto; elVarCanc.className = "kpi-var " + varCanc.classe; }
  set("kpi-churn-atual", k.churn.atual + "%"); set("kpi-churn-ant", k.churn.anterior + "%");
  const elVarChurn = document.getElementById("kpi-churn-var");
  if (elVarChurn) { elVarChurn.textContent = varChurn.texto; elVarChurn.className = "kpi-var " + varChurn.classe; }
}

function renderizarPlanos(d) {
  const container = document.getElementById("planosChart");
  const cores = ["#10B981","#06B6D4","#F59E0B","#8B5CF6","#EC4899","#F97316","#14B8A6","#6366F1"];
  const maximo = Math.max(...d.planos.map(p => p.total));
  container.innerHTML = d.planos.sort((a,b)=>b.total-a.total).map((p,i)=>{
    const largura = Math.round((p.total/maximo)*100);
    return `<div class="bar-row"><div class="bar-label" title="${p.nome}">${p.nome}</div><div class="bar-track"><div class="bar-fill" style="width:${largura}%;background:${cores[i%cores.length]};">${fmt(p.total)}</div></div><div class="bar-count">${fmt(p.total)}</div></div>`;
  }).join("");
}

function renderizarGraficos(d) {
  const gridColor = "#ffffff0f", tickColor = "#94A3B8";
  [chartPort,chartNova,chartChip,chartPgto,chartTimeline,chartRecorrencia,chartOperadoras].forEach(c=>{if(c)c.destroy();});
  chartPort=chartNova=chartChip=chartPgto=chartTimeline=chartRecorrencia=chartOperadoras=null;

  const optD = { responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } } };
  chartPort = new Chart(document.getElementById("chartPort"),{type:"doughnut",data:{labels:["Ativos","Cancelados"],datasets:[{data:[d.portabilidades.ativos,d.portabilidades.cancelados],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  chartNova = new Chart(document.getElementById("chartNova"),{type:"doughnut",data:{labels:["Ativos","Cancelados"],datasets:[{data:[d.novasLinhas.ativos,d.novasLinhas.cancelados],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});
  chartChip = new Chart(document.getElementById("chartChip"),{type:"doughnut",data:{labels:["eSIM","Físico"],datasets:[{data:[d.chip.esim,d.chip.fisico],backgroundColor:["#8B5CF6","#EC4899"],borderWidth:0}]},options:optD});

  const pgtoEntries = Object.entries(d.pagamento.formas).sort((a,b)=>b[1]-a[1]);
  chartPgto = new Chart(document.getElementById("chartPgto"),{type:"doughnut",data:{labels:pgtoEntries.map(([n])=>n),datasets:[{data:pgtoEntries.map(([,q])=>q),backgroundColor:["#EC4899","#8B5CF6","#10B981","#06B6D4","#F59E0B","#F97316","#14B8A6"],borderWidth:0}]},options:optD});
  chartRecorrencia = new Chart(document.getElementById("chartRecorrencia"),{type:"doughnut",data:{labels:["Com recorrência","Sem recorrência"],datasets:[{data:[d.recorrencia.com,d.recorrencia.sem],backgroundColor:["#10B981","#EC4899"],borderWidth:0}]},options:optD});

  const topOp = Object.entries(d.operadoras).sort((a,b)=>b[1]-a[1]).slice(0,6);
  chartOperadoras = new Chart(document.getElementById("chartOperadoras"),{type:"bar",data:{labels:topOp.map(([n])=>n),datasets:[{data:topOp.map(([,q])=>q),backgroundColor:["#10B981","#06B6D4","#F59E0B","#8B5CF6","#EC4899","#F97316"],borderRadius:4,borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:tickColor,font:{size:11}},grid:{display:false}},y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true}}}});

  chartTimeline = new Chart(document.getElementById("chartTimeline"),{type:"line",data:{labels:d.timeline.labels,datasets:[
    {label:"Ativações",data:d.timeline.ativacoes,borderColor:"#10B981",backgroundColor:"#10b9811a",fill:true,tension:0.3,pointRadius:3,yAxisID:"y"},
    {label:"Cancelamentos",data:d.timeline.cancelamentos,borderColor:"#EF4444",backgroundColor:"#ef44441a",fill:true,tension:0.3,pointRadius:3,yAxisID:"y"},
    {label:"Churn %",data:d.timeline.churn,borderColor:"#F59E0B",backgroundColor:"#f59e0b0d",fill:false,tension:0.3,pointRadius:3,borderDash:[5,3],yAxisID:"y2"},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{
    x:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor}},
    y:{ticks:{color:tickColor,font:{size:11}},grid:{color:gridColor},beginAtZero:true,position:"left"},
    y2:{ticks:{color:"#CBD5E1",font:{size:11},callback:v=>v+"%"},grid:{display:false},beginAtZero:true,position:"right"},
  }}});
}

// ── Filtro de meses ───────────────────────────────────────────────────────────
function construirFiltroMeses(d) {
  const container = document.getElementById("filtroMeses");
  container.innerHTML = "";
  const btnTodos = document.createElement("button");
  btnTodos.textContent = "Todos"; btnTodos.className = "filtro-btn active"; btnTodos.dataset.mes = "todos";
  btnTodos.addEventListener("click", () => aplicarFiltro("todos"));
  container.appendChild(btnTodos);
  d.timeline.todosMeses.forEach((mes, i) => {
    const btn = document.createElement("button");
    btn.textContent = d.timeline.labels[i]; btn.className = "filtro-btn"; btn.dataset.mes = mes;
    btn.addEventListener("click", () => aplicarFiltro(mes));
    container.appendChild(btn);
     aplicarFiltro("todos");
  });
}

function aplicarFiltro(mesSelecionado) {
  document.querySelectorAll(".filtro-btn").forEach(b => b.classList.toggle("active", b.dataset.mes === mesSelecionado));
  const d = dadosGlobais.timeline;
  let labels, ativacoes, cancelamentos;
  if (mesSelecionado === "todos") {
    labels = d.labels; ativacoes = d.ativacoes; cancelamentos = d.cancelamentos; churn = d.churn || [];
  } else {
    const [ano, mes] = mesSelecionado.split("-").map(Number);
    const diasNoMes = new Date(ano, mes, 0).getDate();
    const diasAtiv  = d.ativacoesPorDia[mesSelecionado] || {};
    const diasCanc  = d.cancelamentosPorDia[mesSelecionado] || {};
    labels        = Array.from({length:diasNoMes},(_,i)=>String(i+1));
    ativacoes     = Array.from({length:diasNoMes},(_,i)=>diasAtiv[i+1]||0);
    cancelamentos = Array.from({length:diasNoMes},(_,i)=>diasCanc[i+1]||0);
    churn         = ativacoes.map((atv,i)=>atv>0?parseFloat(((cancelamentos[i]/atv)*100).toFixed(1)):0);
  }
  chartTimeline.data.labels=labels; chartTimeline.data.datasets[0].data=ativacoes;
  chartTimeline.data.datasets[1].data=cancelamentos; chartTimeline.data.datasets[2].data=churn;
  chartTimeline.update();
    const totalAtv  = ativacoes.reduce((a,b)=>a+b,0);
  const totalCanc = cancelamentos.reduce((a,b)=>a+b,0);
  const churnMed  = totalAtv > 0 ? ((totalCanc/totalAtv)*100).toFixed(1) : 0;
  set("legend-atv-total",  fmt(totalAtv));
  set("legend-canc-total", fmt(totalCanc));
  set("legend-churn-total", churnMed + "%");
}

// ── Mapa de calor do Brasil ───────────────────────────────────────────────────
function renderizarMapa(d) {
  const estados = d.estados;
  const maxVal  = Math.max(...Object.values(estados), 1);
  const colorScale = d3.scaleLinear().domain([0, maxVal]).range(["#334155", "#10B981"]);
  const ranking = Object.entries(estados).sort((a,b)=>b[1]-a[1]);
  const totalEstados = Object.values(estados).reduce((a,b)=>a+b,0);
  const tabelaEl = document.getElementById("mapaTabela");
  if (tabelaEl) tabelaEl.innerHTML = `<table class="split-table"><thead><tr><th>Estado</th><th>Clientes</th><th>%</th></tr></thead><tbody>${ranking.map(([uf,val])=>`<tr><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${colorScale(val)};margin-right:8px;"></span>${nomesEstados[uf]||uf} (${uf})</td><td>${fmt(val)}</td><td style="color:var(--text-muted);font-size:11px;">${pct(val,totalEstados)}</td></tr>`).join("")}</tbody></table>`;

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
        .on("mousemove",(ev,f)=>{const uf=f.properties.sigla||f.id;tip.style("display","block").style("left",(ev.clientX+14)+"px").style("top",(ev.clientY-10)+"px").html(`<strong>${nomesEstados[uf]||uf}</strong><br>${fmt(estados[uf]||0)} clientes`);})
        .on("mouseleave",()=>tip.style("display","none"));
      svg.selectAll("text").data(geojson.features).enter().append("text")
        .attr("x",f=>path.centroid(f)[0]).attr("y",f=>path.centroid(f)[1])
        .attr("text-anchor","middle").attr("dominant-baseline","middle").attr("font-size","7").attr("font-weight","600").attr("fill","white").style("pointer-events","none")
        .text(f=>f.properties.sigla||f.id||"");
    })
    .catch(()=>{container.innerHTML='<div style="color:var(--text-muted);padding:1rem;">❌ Não foi possível carregar o mapa.</div>';});
}

// ── Dashboard de Clientes ─────────────────────────────────────────────────────
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
  const statusCores = {"Sucesso":"#10B981","Portabilidade negada":"#EF4444","Portabilidade em andamento":"#F59E0B","Aguardando confirmação":"#06B6D4"};
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

// ── Dashboard de Campanhas ────────────────────────────────────────────────────
let chartCampanha = null;

function renderizarCampanhas(d) {
  const c = d.campanhas;
  const fmtBRL = (v) => "R$ " + Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  set("camp-total-bonif",   fmt(c.totalBonificados));
  set("camp-com-recarga",   fmt(c.comRecarga));
  set("camp-sem-recarga",   fmt(c.semRecarga));
  set("camp-pct-conversao", c.pctConversao + "%");
  set("camp-ganho",         fmtBRL(c.ganho));
  set("camp-perda",         fmtBRL(c.perda));
  set("camp-liquido",       fmtBRL(c.liquido));
  set("camp-valor-bonif",   fmtBRL(c.valorBonif));

  set("camp-tab-com",    fmt(c.comRecarga));
  set("camp-tab-ganho",  fmtBRL(c.ganho));
  set("camp-tab-sem",    fmt(c.semRecarga));
  set("camp-tab-perda",  fmtBRL(c.perda));
  set("camp-tab-total",  fmt(c.totalBonificados));
  const elLiq = document.getElementById("camp-tab-liquido");
  if (elLiq) { elLiq.textContent = fmtBRL(c.liquido); elLiq.style.color = c.liquido >= 0 ? "#10B981" : "#EF4444"; }

  if (chartCampanha) { chartCampanha.destroy(); chartCampanha = null; }
  const elChart = document.getElementById("chartCampanha");
  if (elChart) {
    chartCampanha = new Chart(elChart, {
      type: "doughnut",
      data: {
        labels: ["Recarregaram", "Não recarregaram"],
        datasets: [{ data: [c.comRecarga, c.semRecarga], backgroundColor: ["#10B981", "#EF4444"], borderWidth: 0 }],
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } },
    });
  }
}

// ── Init ──────────────────────────────────────────────────────────────────────
setupTema();
setupDropZone();
setupNav();