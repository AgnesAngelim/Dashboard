// ── dash-documentacao.js — Página de Documentação editável ───────────────────

const DOC_STORAGE_KEY = "igreen_documentacao_v1";

const DOC_DEFAULTS = [
  {
    id: "doc-geral-visao",
    titulo: "Dashboard Geral — Visão geral",
    conteudo: `Os cards principais mostram um resumo da base inteira.

Total de clientes — conta todas as linhas da planilha.
Ativos — registros com Status igual a "Ativo", "Ativos" ou "Ativa".
Cancelados — registros com Status igual a "Cancelado", "Cancelados" ou "Cancelada".
Portabilidades — linhas onde a coluna Portabilidades contém "True" ou "Verdadeiro".
Novas linhas — todas as linhas que não são portabilidade.
Bonificados — registros com Tipo igual a "BONIFICADA".
Logística — registros com a coluna Logistica preenchida (exclui vazio, "-", "N/D" e "0").
Tempo médio de vida — média de dias entre Data de ativação e Data cancelado, calculada apenas para cancelados.`
  },
  {
    id: "doc-geral-kpi",
    titulo: "Dashboard Geral — Variação mensal (KPI)",
    conteudo: `Os cards de KPI mostram o desempenho do mês mais recente comparado ao anterior.

Ativações — quantidade de registros com Data de ativação no mês. A seta mostra se subiu ou desceu em relação ao mês anterior.
Cancelamentos — quantidade de registros cancelados com Data cancelado no mês. A seta é invertida: subir é ruim.
Churn — cancelamentos do mês ÷ total de ativos da base × 100. Indica a taxa de saída de clientes.`
  },
  {
    id: "doc-geral-port",
    titulo: "Dashboard Geral — Portabilidade vs Nova linha",
    conteudo: `Divide a base entre linhas que vieram de portabilidade e linhas novas, mostrando quantas estão ativas ou canceladas em cada grupo.

Portabilidades — linhas onde Portabilidades = "True/Verdadeiro". Mostra ativos e cancelados.
Novas linhas — todas as demais linhas. Mostra ativos e cancelados.
Status de portabilidades (só ativos) — entre os portados ativos, mostra quantos tiveram portabilidade Aprovada (coluna Portabilidade contém "SUCESSO"), Negada (contém "negada") ou Em andamento (demais valores).`
  },
  {
    id: "doc-geral-chip",
    titulo: "Dashboard Geral — Tipo de chip e Forma de pagamento",
    conteudo: `Ambos os gráficos consideram apenas registros ativos.

eSIM vs Físico — lê a coluna Tipo de chip e agrupa em "esim" ou "fisico".
Forma de pagamento — lê a coluna Forma de pagamento. Exclui automaticamente: "Baixa manual", "Pagamento com saldo" e "-".`
  },
  {
    id: "doc-geral-recorrencia",
    titulo: "Dashboard Geral — Recorrência",
    conteudo: `Considera apenas registros ativos.

Com recorrência — registros onde a coluna Recorrência contém exatamente "Cartão de crédito | recorrência cartão" ou "Criptomoeda | recorrência cartão".
Sem recorrência — todos os demais ativos.`
  },
  {
    id: "doc-geral-operadoras",
    titulo: "Dashboard Geral — Top operadoras e Planos",
    conteudo: `Top operadoras doadoras — lê a coluna Operadora de todos os registros e exibe as 6 com mais linhas.
Planos mais usados — lê a coluna Plano apenas dos registros ativos, ordenados do mais usado para o menos.`
  },
  {
    id: "doc-geral-timeline",
    titulo: "Dashboard Geral — Ativações e cancelamentos ao longo do tempo",
    conteudo: `Gráfico de linha que mostra a evolução mês a mês (ou dia a dia ao filtrar um mês específico).

Ativações — conta todos os registros pela Data de ativação, independente do status atual.
Cancelamentos — conta apenas os registros com status cancelado, pela Data cancelado.

O filtro global (botão 🗓️ Filtros) afeta todo o Dashboard Geral. Ao selecionar um mês específico, o gráfico passa a mostrar os dias daquele mês.`
  },
  {
    id: "doc-geral-mapa",
    titulo: "Dashboard Geral — Clientes por estado",
    conteudo: `Mapa de calor do Brasil colorido pela concentração de registros por estado.

Estado — lê a coluna Estado de todos os registros.
Cidade — lê a coluna Cidade, Municipio ou Município. Ao clicar em um estado no mapa ou na tabela, abre um detalhamento por cidade.`
  },
  {
    id: "doc-clientes-kpi",
    titulo: "Dashboard de Clientes — Visão geral",
    conteudo: `Os cards do Dashboard de Clientes usam apenas registros filtrados pelo arquivo BackOffice — só entram chips presentes no backoffice.

Clientes únicos — quantidade de CPFs distintos na base filtrada.
Números repetidos — números da coluna Numero De Origem que aparecem em mais de uma linha.
Clientes com nº repetido — CPFs que possuem pelo menos um desses números repetidos associado.
Nº em múltiplos clientes — números que aparecem em CPFs diferentes. Indica possível fraude ou compartilhamento indevido.`
  },
  {
    id: "doc-clientes-tabelas",
    titulo: "Dashboard de Clientes — Tabelas principais",
    conteudo: `Linhas ativadas e repetidas — mostra apenas clientes com números repetidos, ordenados do maior para o menor número de repetições. Inclui nome, CPF, ID licenciado, total de ativações e quantidade de repetições.

Número em múltiplos clientes — números de origem que aparecem em CPFs diferentes, com total de linhas, clientes distintos e ID licenciado.

Índice de cancelamento — clientes com mais de 50% das linhas canceladas (mínimo 2 linhas), com percentual calculado por CPF.

Clientes com mais de uma recarga — soma a coluna Qº de recargas por CPF. Ao clicar no nome do cliente, abre um modal com todas as linhas dele mostrando número (Numero De Origem ou Numero gerado formatado como (00) 00000-0000), plano, status, data de ativação e recargas de cada linha.`
  },
  {
    id: "doc-clientes-port",
    titulo: "Dashboard de Clientes — Portabilidades e operadoras",
    conteudo: `Distribuição por status — gráfico de rosca com os status da coluna Portabilidade. Possui filtro de mês que atualiza apenas esse gráfico, baseado na Data de ativação.

Os status reconhecidos são: Sucesso, Portabilidade negada, Portabilidade em andamento, Aguardando confirmação.

Números repetidos na operadora doadora — top 50 números da coluna Numero De Origem com mais ocorrências, com ID licenciado e CPFs associados.

Detalhamento por operadora — agrupa a coluna Operadora e conta sucesso, negadas e andamento pela coluna Portabilidade.`
  },
  {
    id: "doc-planilhas",
    titulo: "Planilhas e filtros necessários",
    conteudo: `Planilha de Relatório — planilha principal com todos os registros. É obrigatória para carregar qualquer dashboard.

Planilha BackOffice — usada para filtrar registros no Dashboard de Clientes. Só entram clientes cujo número de chip (coluna Chip) está presente no backoffice. Se não houver backoffice, todos os registros entram.

Filtro global (🗓️ Filtros) — aparece no Dashboard Geral. Permite filtrar por mês de ativação/cancelamento ou por período personalizado com data de início e fim.

Filtro de mês — portabilidades — aparece no Dashboard de Clientes, dentro do gráfico de Distribuição por status. Afeta apenas esse gráfico.`
  },
];

function carregarDocumentacao() {
  try {
    const salvo = localStorage.getItem(DOC_STORAGE_KEY);
    if (salvo) return JSON.parse(salvo);
  } catch {}
  return DOC_DEFAULTS;
}

function salvarDocumentacao() {
  const secoes = document.querySelectorAll(".doc-secao");
  const dados = [];
  secoes.forEach(s => {
    dados.push({
      id:       s.dataset.id,
      titulo:   s.querySelector(".doc-titulo").innerText.trim(),
      conteudo: s.querySelector(".doc-corpo").innerText.trim(),
    });
  });
  localStorage.setItem(DOC_STORAGE_KEY, JSON.stringify(dados));
}

function resetarDocumentacao() {
  if (!confirm("Resetar toda a documentação para o conteúdo padrão?")) return;
  localStorage.removeItem(DOC_STORAGE_KEY);
  renderizarDocumentacao();
}

function adicionarSecao() {
  const id = "doc-custom-" + Date.now();
  const container = document.getElementById("doc-container");
  const div = document.createElement("div");
  div.className = "doc-secao chart-box";
  div.dataset.id = id;
  div.style.marginBottom = "16px";
  div.style.position = "relative";
  div.innerHTML = `
    <button onclick="removerSecao('${id}')" title="Remover seção"
      style="position:absolute;top:10px;right:10px;background:transparent;border:none;color:var(--text-muted);font-size:16px;cursor:pointer;line-height:1;padding:2px 6px;border-radius:4px;transition:color .15s;"
      onmouseover="this.style.color='#EF4444'" onmouseout="this.style.color='var(--text-muted)'">✕</button>
    <div
      class="doc-titulo"
      contenteditable="true"
      spellcheck="false"
      style="font-size:14px;font-weight:600;color:var(--white);margin-bottom:10px;outline:none;border-bottom:1px solid transparent;padding-bottom:4px;padding-right:24px;transition:border-color .2s;"
      onfocus="this.style.borderColor='#10B981'"
      onblur="this.style.borderColor='transparent';salvarDocumentacao()"
    >Novo campo</div>
    <div
      class="doc-corpo"
      contenteditable="true"
      spellcheck="false"
      style="font-size:13px;color:var(--text);line-height:1.8;outline:none;white-space:pre-wrap;min-height:40px;"
      onblur="salvarDocumentacao()"
    >Clique para editar o conteúdo...</div>
  `;
  container.appendChild(div);
  salvarDocumentacao();
  div.querySelector(".doc-titulo").focus();
}

function removerSecao(id) {
  if (!confirm("Remover esta seção?")) return;
  const el = document.querySelector(`.doc-secao[data-id="${id}"]`);
  if (el) el.remove();
  salvarDocumentacao();
}

function criarSecaoHTML(s, removivel) {
  return `
    <div class="doc-secao chart-box" data-id="${s.id}" style="margin-bottom:16px;position:relative;">
      ${removivel ? `<button onclick="removerSecao('${s.id}')" title="Remover seção"
        style="position:absolute;top:10px;right:10px;background:transparent;border:none;color:var(--text-muted);font-size:16px;cursor:pointer;line-height:1;padding:2px 6px;border-radius:4px;transition:color .15s;"
        onmouseover="this.style.color='#EF4444'" onmouseout="this.style.color='var(--text-muted)'">✕</button>` : ""}
      <div
        class="doc-titulo"
        contenteditable="true"
        spellcheck="false"
        style="font-size:14px;font-weight:600;color:var(--white);margin-bottom:10px;outline:none;border-bottom:1px solid transparent;padding-bottom:4px;padding-right:${removivel ? "24px" : "0"};transition:border-color .2s;"
        onfocus="this.style.borderColor='#10B981'"
        onblur="this.style.borderColor='transparent';salvarDocumentacao()"
      >${s.titulo}</div>
      <div
        class="doc-corpo"
        contenteditable="true"
        spellcheck="false"
        style="font-size:13px;color:var(--text);line-height:1.8;outline:none;white-space:pre-wrap;min-height:40px;"
        onblur="salvarDocumentacao()"
      >${s.conteudo}</div>
    </div>
  `;
}

function renderizarDocumentacao() {
  const container = document.getElementById("doc-container");
  if (!container) return;
  const secoes = carregarDocumentacao();
  const defaultIds = DOC_DEFAULTS.map(d => d.id);
  container.innerHTML = secoes.map(s => criarSecaoHTML(s, !defaultIds.includes(s.id))).join("");
}

document.addEventListener("DOMContentLoaded", () => {
  renderizarDocumentacao();
});