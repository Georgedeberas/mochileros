Office.onReady(() => {
    console.log("Mochileros RD Add-in listo.");
});

// ----------- TOGGLE PANEL -----------
async function togglePanel() {
    try {
        if (window._panelOpen) {
            await Office.addin.hide();
            window._panelOpen = false;
        } else {
            await Office.addin.showAsTaskpane();
            window._panelOpen = true;

            // Recarga completa al abrir (para cargar cambios desde GitHub)
            reloadPanelFiles();
        }
    } catch (e) {
        console.error("Error al alternar panel:", e);
    }
}
Office.actions.associate("togglePanel", togglePanel);


// ----------- RECARGA AUTOMÁTICA CUANDO EL PANEL SE ABRE -----------

function reloadPanelFiles() {
    // Esta función fuerza a que los archivos se recarguen desde GitHub Pages.

    // Agregamos un query-string con timestamp para evitar caché
    const time = Date.now();

    // Recargar CSS
    document.querySelectorAll('link[rel="stylesheet"]').forEach(link => {
        const href = link.getAttribute('href').split('?')[0];
        link.setAttribute('href', `${href}?v=${time}`);
    });

    // Recargar JS
    document.querySelectorAll('script').forEach(script => {
        if (script.src) {
            const src = script.src.split('?')[0];
            const nuevo = document.createElement('script');
            nuevo.src = `${src}?v=${time}`;
            script.replaceWith(nuevo);
        }
    });

    console.log("Archivos recargados desde GitHub:", time);
}
