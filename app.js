const LS = window.localStorage;
const LS_PANE = "mr_pane_visible";
const LS_VERSION = "mr_version";
const VERSION_URL = "version.json"; // relativo al sitio
let CURRENT_VERSION = null;

Office.onReady(async () => {
  await ensureVersionAndBust();
  markPaneVisible();        // dejamos marca para que el botón de cinta sepa el estado
  wireUI();
  // Aquí puedes llamar a tu renderizado real de meses/hojas
  // await renderWorkbook();
});

function markPaneVisible() {
  LS.setItem(LS_PANE, "1");
}

function wireUI() {
  const btnRefresh = document.getElementById("btnRefresh");
  const btnNew = document.getElementById("btnNew");
  const btnFind = document.getElementById("btnFind");

  if (btnRefresh) btnRefresh.addEventListener("click", async () => {
    await refreshPanel();
  });

  if (btnNew) btnNew.addEventListener("click", async () => {
    // TODO: tu lógica real de creación (plantillas Nacional/Internacional)
    alert("‘Nuevo tour’: aquí insertas tu flujo real de creación de hoja.");
  });

  if (btnFind) btnFind.addEventListener("click", async () => {
    const q = (document.getElementById("search")?.value || "").trim();
    if (!q) { alert("Escribe algo para buscar."); return; }
    // TODO: tu búsqueda real de hojas (por nombre/mes/estado)
    alert(`Buscar: ${q} (engancha aquí tu motor de búsqueda).`);
  });
}

async function refreshPanel() {
  // Aquí puedes re-leer el libro: recorrer hojas, estados por color, etc.
  // En este ejemplo sólo recargamos la UI actual.
  try {
    await ensureVersionAndBust(); // por si subiste algo a GitHub
    // Re-renderizar tu lista:
    // await renderWorkbook();
  } catch (e) {
    console.error(e);
  }
}

/** Lee version.json, guarda en localStorage y aplica busting ?v= */
async function ensureVersionAndBust() {
  try {
    const ver = await fetch(`${VERSION_URL}?t=${Date.now()}`).then(r=>r.json()).then(j=>String(j.version || "0"));
    CURRENT_VERSION = ver;
    const prev = LS.getItem(LS_VERSION);
    if (prev !== ver) {
      LS.setItem(LS_VERSION, ver);
      applyCacheBusting(ver);
    } else {
      // ya en la misma versión; igualmente aplicamos para asegurar
      applyCacheBusting(ver);
    }
  } catch (e) {
    console.warn("No se pudo leer version.json (continuamos):", e);
  }
}

/** Aplica ?v=VERSION a todos los <link rel="stylesheet"> y scripts locales */
function applyCacheBusting(ver) {
  // estilos
  document.querySelectorAll('link[rel="stylesheet"]').forEach(link => {
    if (!link.href) return;
    const url = new URL(link.href, location.origin);
    if (url.origin === location.origin) { // solo locales
      url.searchParams.set("v", ver);
      link.href = url.toString();
    }
  });
  // scripts locales (excepto office.js)
  document.querySelectorAll('script[src]').forEach(scr => {
    if (!scr.src) return;
    if (scr.src.includes("appsforoffice.microsoft.com")) return;
    const url = new URL(scr.src, location.origin);
    if (url.origin === location.origin) {
      url.searchParams.set("v", ver);
      // Si ya está cargado, lo reemplazamos forzando recarga
      const clone = document.createElement("script");
      clone.src = url.toString();
      clone.defer = scr.defer;
      scr.replaceWith(clone);
    }
  });
}
