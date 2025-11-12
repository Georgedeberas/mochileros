// ===== Utilidades =====
const LS = window.localStorage;
const LS_PANE = "mr_pane_visible";      // "1" ó "0"
const LS_VERSION = "mr_version";        // última versión conocida (string)
const VERSION_URL = "https://georgedeberas.github.io/mochileros/version.json";

function setPaneVisible(flag) {
  LS.setItem(LS_PANE, flag ? "1" : "0");
}

function isPaneVisible() {
  return LS.getItem(LS_PANE) === "1";
}

async function fetchRemoteVersion() {
  const res = await fetch(`${VERSION_URL}?t=${Date.now()}`);
  if (!res.ok) throw new Error(`No se pudo leer version.json (${res.status})`);
  const json = await res.json();
  return (json && json.version) ? String(json.version) : null;
}

function setRibbonLabel(label) {
  if (!Office.ribbon || !Office.ribbon.requestUpdate) return;
  Office.ribbon.requestUpdate({
    tabs: [{
      id: "Tab.Mochileros",
      groups: [{
        id: "Group.Main",
        controls: [{
          id: "Btn.Toggle",
          label
        }]
      }]
    }]
  });
}

// ===== Botones de Cinta =====

// Toggle panel (mostrar/ocultar) y actualizar etiqueta
async function togglePanel(event) {
  try {
    const visible = isPaneVisible();
    if (visible) {
      if (Office.addin && Office.addin.hide) await Office.addin.hide();
      setPaneVisible(false);
      setRibbonLabel("Activar panel");
    } else {
      if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane(true);
      setPaneVisible(true);
      setRibbonLabel("Panel activado");
    }
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

// Actualizar Sistema: lee version.json y marca en localStorage para que el panel recargue con cache-busting
async function updateSystem(event) {
  try {
    const remote = await fetchRemoteVersion();
    if (remote) {
      const current = LS.getItem(LS_VERSION) || "0";
      if (remote !== current) {
        LS.setItem(LS_VERSION, remote);
        // Si el panel está visible, lo volvemos a mostrar (lo que hace que se refresque),
        // y el JS del panel recargará recursos con ?v={version}
        if (Office.addin && Office.addin.showAsTaskpane) {
          await Office.addin.showAsTaskpane(true);
        }
        // Feedback breve en el segundo botón (opcional)
        if (Office.ribbon && Office.ribbon.requestUpdate) {
          Office.ribbon.requestUpdate({
            tabs: [{
              id: "Tab.Mochileros",
              groups: [{
                id: "Group.Main",
                controls: [{
                  id: "Btn.Update",
                  label: "Actualizado ✔"
                }]
              }]
            }]
          });
          // volver al texto original en 2s
          setTimeout(() => {
            Office.ribbon.requestUpdate({
              tabs: [{
                id: "Tab.Mochileros",
                groups: [{
                  id: "Group.Main",
                  controls: [{ id: "Btn.Update", label: "Actualizar sistema" }]
                }]
              }]
            });
          }, 2000);
        }
      } else {
        // Ya está en la última versión
        console.log("Sistema ya está en la última versión:", remote);
      }
    }
  } catch (e) {
    console.error("updateSystem error:", e);
  } finally {
    event.completed();
  }
}

// ===== Inicialización =====
Office.onReady(() => {
  // Aseguramos etiqueta inicial del botón
  setRibbonLabel(isPaneVisible() ? "Panel activado" : "Activar panel");
});

// Export global para la cinta
window.togglePanel = togglePanel;
window.updateSystem = updateSystem;
