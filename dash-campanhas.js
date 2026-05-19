// ── dash-campanhas.JS — Dashboard de Campanhas ───────────────────────────────

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

  set("camp-tab-com",   fmt(c.comRecarga));
  set("camp-tab-ganho", fmtBRL(c.ganho));
  set("camp-tab-sem",   fmt(c.semRecarga));
  set("camp-tab-perda", fmtBRL(c.perda));
  set("camp-tab-total", fmt(c.totalBonificados));

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
