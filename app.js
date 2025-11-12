// Inicialización segura para Excel Online
if (window.Office && Office.onReady) {
  Office.onReady(() => {
    const status = document.getElementById("status");
    const btn = document.getElementById("btnRefresh");

    if (status) status.textContent = "Listo para usar.";
    if (btn) {
      btn.addEventListener("click", () => {
        if (status) status.textContent = "Actualizando...";
        // Cache-busting para forzar a cargar la última versión del panel
        const url = new URL(window.location.href);
        url.searchParams.set("_r", Date.now().toString());
        window.location.replace(url.toString());
      });
    }
  });
} else {
  // Fallback si Office.js aún no está listo (útil al probar fuera de Excel)
  const status = document.getElementById("status");
  if (status) status.textContent = "Listo (sin Office.js).";
}
