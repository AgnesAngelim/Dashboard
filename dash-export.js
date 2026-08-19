// ── dash-export.js — Exportar snapshot estático em HTML único ────────────────
// Gera um arquivo .html autossuficiente (CSS e JS embutidos) com os dados já
// carregados no dashboard, para compartilhar sem precisar enviar as planilhas.

document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("btnExportarHTML");
  if (btn) btn.addEventListener("click", exportarHTML);
});

function escaparParaScript(str) {
  // Evita que "</script" dentro dos dados feche a tag <script> prematuramente.
  return str.replace(/<\/script/gi, "<\\/script");
}

async function exportarHTML() {
  if (!dadosGlobais || !arquivoPrincipal) {
    alert("Carregue uma planilha antes de exportar o HTML.");
    return;
  }

  const btn = document.getElementById("btnExportarHTML");
  const textoOriginal = btn ? btn.textContent : "";
  if (btn) { btn.textContent = "⏳ Gerando..."; btn.disabled = true; }

  try {
    const base = location.href;
    const [htmlSrc, cssSrc, coreSrc, geralSrc, clientesSrc, docSrc] = await Promise.all([
      fetch(base, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("index.html"); return r.text(); }),
      fetch(new URL("style.css", base).href, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("style.css"); return r.text(); }),
      fetch(new URL("dash-core.js", base).href, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("dash-core.js"); return r.text(); }),
      fetch(new URL("dash-geral.js", base).href, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("dash-geral.js"); return r.text(); }),
      fetch(new URL("dash-clientes.js", base).href, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("dash-clientes.js"); return r.text(); }),
      fetch(new URL("dash-documentacao.js", base).href, { cache: "no-store" }).then(r => { if (!r.ok) throw new Error("dash-documentacao.js"); return r.text(); }),
    ]);

    // Tenta embutir o logo como imagem inline (não é crítico — se falhar, segue sem ele)
    let logoDataUrl = null;
    try {
      const logoResp = await fetch(new URL("logo.png", base).href, { cache: "no-store" });
      if (logoResp.ok) {
        const logoBlob = await logoResp.blob();
        logoDataUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.onerror = reject;
          reader.readAsDataURL(logoBlob);
        });
      }
    } catch (e) {
      console.warn("dash-export: não foi possível embutir o logo, seguindo sem ele.", e);
    }

    // Snapshot do conteúdo atual da página de Documentação (já editado por você)
    const docSecoes = [];
    document.querySelectorAll("#doc-container .doc-secao").forEach(s => {
      const tituloEl = s.querySelector(".doc-titulo");
      const corpoEl  = s.querySelector(".doc-corpo");
      if (!tituloEl || !corpoEl) return;
      docSecoes.push({ id: s.dataset.id, titulo: tituloEl.innerText.trim(), conteudo: corpoEl.innerText.trim() });
    });

    const paginaAtiva = document.querySelector(".nav-item.active")?.dataset.page || "geral";
    const temaAtual   = document.body.dataset.tema || "dark";

    // Captura o filtro global (mês / período personalizado / "todos") que está ativo agora
    const filtroBtnAtivo = document.querySelector("#filtroGlobal .filtro-btn.active");
    let filtroModo = "todos", filtroInicio = null, filtroFim = null;
    if (filtroBtnAtivo) {
      if (filtroBtnAtivo.id === "btnAplicarPeriodo") {
        filtroInicio = document.getElementById("filtroDataInicio")?.value || null;
        filtroFim    = document.getElementById("filtroDataFim")?.value || null;
        filtroModo   = (filtroInicio && filtroFim) ? "periodo" : "todos";
      } else if (filtroBtnAtivo.dataset.mes) {
        filtroModo = filtroBtnAtivo.dataset.mes;
      }
    }

    const agora = new Date();
    const dataStr = `${String(agora.getDate()).padStart(2,"0")}/${String(agora.getMonth()+1).padStart(2,"0")}/${agora.getFullYear()} ${String(agora.getHours()).padStart(2,"0")}:${String(agora.getMinutes()).padStart(2,"0")}`;

    let htmlOut = htmlSrc;

    // 1) Título da aba
    htmlOut = htmlOut.replace(/<title>[\s\S]*?<\/title>/, `<title>Dashboard iGreen — Exportado ${dataStr}</title>`);

    // 2) CSS embutido (substitui o <link>) — arquivo final vira 100% autossuficiente
    htmlOut = htmlOut.replace(/<link[^>]*href=["']style\.css["'][^>]*>/, `<style>\n${cssSrc}\n</style>`);

    // 2b) Logo embutido como imagem inline, se foi possível buscar
    if (logoDataUrl) {
      htmlOut = htmlOut.replace(/src=["']logo\.png["']/, `src="${logoDataUrl}"`);
    }

    // 3) Remove a biblioteca XLSX e o próprio script de exportação (não fazem sentido no snapshot)
    htmlOut = htmlOut.replace(/\s*<script[^>]*src=["'][^"']*xlsx[^"']*["'][^>]*><\/script>/i, "");
    htmlOut = htmlOut.replace(/\s*<script[^>]*src=["']dash-export\.js["'][^>]*><\/script>/, "");

    // 4) Substitui os <script src="dash-*.js"> pelo código embutido
    htmlOut = htmlOut.replace(/<script[^>]*src=["']dash-core\.js["'][^>]*><\/script>/, `<script>\n${coreSrc}\n</script>`);
    htmlOut = htmlOut.replace(/<script[^>]*src=["']dash-geral\.js["'][^>]*><\/script>/, `<script>\n${geralSrc}\n</script>`);
    htmlOut = htmlOut.replace(/<script[^>]*src=["']dash-clientes\.js["'][^>]*><\/script>/, `<script>\n${clientesSrc}\n</script>`);

    // 5) Script de inicialização — restaura os dados da planilha (para o Filtro Global
    //    continuar funcionando) e reconstrói a tela exatamente como estava. Vai logo
    //    após o dash-documentacao.js.
    const docSecoesJSON = escaparParaScript(JSON.stringify(docSecoes));
    const arquivoPrincipalJSON  = escaparParaScript(JSON.stringify(arquivoPrincipal || []));
    const arquivoBackofficeJSON = escaparParaScript(JSON.stringify(arquivoBackoffice ? [...arquivoBackoffice] : []));
    const scriptInit = `
  <script>
  (function() {
    // Pré-carrega a documentação com o conteúdo atual (antes do renderizarDocumentacao rodar)
    try {
      var docSnapshot = ${docSecoesJSON};
      if (docSnapshot.length && !localStorage.getItem(DOC_STORAGE_KEY)) {
        localStorage.setItem(DOC_STORAGE_KEY, JSON.stringify(docSnapshot));
      }
    } catch (e) {}

    // Tema salvo no momento da exportação
    try { localStorage.setItem("tema", ${JSON.stringify(temaAtual)}); } catch (e) {}
    aplicarTema(${JSON.stringify(temaAtual)});

    // Esconde elementos que só fazem sentido durante o upload (não durante a visualização)
    var uploadArea = document.getElementById("uploadArea");
    if (uploadArea) uploadArea.remove();
    document.querySelectorAll("button").forEach(function(b){
      if (b.textContent && b.textContent.indexOf("Nova planilha") !== -1) b.style.display = "none";
    });
    var btnExport = document.getElementById("btnExportarHTML");
    if (btnExport) btnExport.style.display = "none";

    // Aviso de que este é um snapshot exportado
    var main = document.querySelector(".main-content");
    if (main) {
      var aviso = document.createElement("div");
      aviso.style.cssText = "background:rgba(245,158,11,0.12);color:#F59E0B;border:1px solid rgba(245,158,11,0.3);border-radius:8px;padding:8px 14px;font-size:12px;margin-bottom:14px;";
      aviso.textContent = "📄 Snapshot exportado em ${dataStr} — os dados da planilha foram embutidos para o Filtro Global continuar funcionando.";
      main.insertBefore(aviso, main.firstChild);
    }

    // Restaura os dados brutos (para o Filtro Global recalcular por mês/período normalmente)
    arquivoPrincipal  = ${arquivoPrincipalJSON};
    arquivoBackoffice = new Set(${arquivoBackofficeJSON});
    var dadosCompletos = calcularDados(arquivoPrincipal, arquivoBackoffice);
    construirFiltroGlobal(dadosCompletos);

    // Mostra a aba que estava ativa no momento da exportação
    var paginaInicial = ${JSON.stringify(paginaAtiva)};
    document.querySelectorAll(".page").forEach(function(p){ p.style.display = "none"; });
    var pageEl = document.getElementById("page-" + paginaInicial);
    if (pageEl) pageEl.style.display = "block";
    document.querySelectorAll(".nav-item").forEach(function(b){ b.classList.toggle("active", b.dataset.page === paginaInicial); });

    // Reaplica o filtro (mês / período / todos) que estava ativo no momento da exportação
    var filtroModo   = ${JSON.stringify(filtroModo)};
    var filtroInicio = ${JSON.stringify(filtroInicio)};
    var filtroFim    = ${JSON.stringify(filtroFim)};
    if (filtroModo === "periodo" && filtroInicio && filtroFim) {
      document.getElementById("filtroDataInicio").value = filtroInicio;
      document.getElementById("filtroDataFim").value = filtroFim;
      document.getElementById("btnAplicarPeriodo").click();
    } else if (filtroModo && filtroModo !== "todos") {
      var btnMes = document.querySelector('#filtroGlobal .filtro-btn[data-mes="' + filtroModo + '"]');
      if (btnMes) btnMes.click();
      else { dadosGlobais = dadosCompletos; renderDash(dadosGlobais); setTimeout(function(){ renderizarMapa(dadosGlobais); }, 100); }
    } else {
      dadosGlobais = dadosCompletos;
      renderDash(dadosGlobais);
      setTimeout(function(){ renderizarMapa(dadosGlobais); }, 100);
    }
  })();
  </script>`;
    htmlOut = htmlOut.replace(/<script[^>]*src=["']dash-documentacao\.js["'][^>]*><\/script>/, `<script>\n${docSrc}\n</script>\n${scriptInit}`);

    // 6) Gera o arquivo e dispara o download
    const blob = new Blob([htmlOut], { type: "text/html;charset=utf-8" });
    const nomeArquivo = `dashboard-igreen-${agora.getFullYear()}${String(agora.getMonth()+1).padStart(2,"0")}${String(agora.getDate()).padStart(2,"0")}-${String(agora.getHours()).padStart(2,"0")}${String(agora.getMinutes()).padStart(2,"0")}.html`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nomeArquivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);

  } catch (err) {
    console.error("Erro ao exportar HTML:", err);
    alert("Não foi possível exportar o HTML. Isso costuma acontecer quando o dashboard é aberto direto do arquivo (file://) em vez de por um servidor/site. Erro: " + err.message);
  } finally {
    if (btn) { btn.textContent = textoOriginal; btn.disabled = false; }
  }
}
