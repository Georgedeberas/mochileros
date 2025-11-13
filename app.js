Office.onReady(() => {
    console.log("Mochileros RD Add-in iniciado correctamente.");
});

// -----------------------------
//   TOGGLE PANEL
// -----------------------------
async function togglePanel() {
    try {
        if (window._panelAbierto) {
            await Office.addin.hide();
            window._panelAbierto = false;
        } else {
            await Office.addin.showAsTaskpane();
            window._panelAbierto = true;

            // Forzar actualización solo al abrir
            reloadAssets();
        }
    } catch (err) {
        console.error("Error en togglePanel:", err);
    }
}

Office.actions.associate("togglePanel", togglePanel);

// -----------------------------
//   RECARGA DE ARCHIVOS
// -----------------------------
function reloadAssets() {
    const version = Date.now();

    // Recargar CSS (forzar sin cache)
    document.querySelectorAll('link[rel="stylesheet"]').forEach(link => {
        const href = link.href.split("?")[0];
        link.href = href + "?v=" + version;
    });

    // Recargar JS
    document.querySelectorAll('script').forEach(script => {
        if (script.src) {
            const src = script.src.split("?")[0];
            const nuevo = document.createElement("script");
            nuevo.src = src + "?v=" + version;
            script.replaceWith(nuevo);
        }
    });

    console.log("Archivos recargados:", version);
}
