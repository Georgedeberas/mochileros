Office.onReady(() => {
  console.log("✅ Commands runtime listo.");
});

// Botones
function accionBasica(event) {
  notify("Botón simple presionado.");
  event.completed();
}

function accionIcono(event) {
  notify("Botón con ícono ejecutado.");
  event.completed();
}

// Menú
function menuAccionA(event) {
  notify("Menú → Opción A.");
  event.completed();
}
function menuAccionB(event) {
  notify("Menú → Opción B.");
  event.completed();
}

// Split button
function accionPrincipal(event) {
  notify("Menú dividido → Acción principal.");
  event.completed();
}
function accionSub1(event) {
  notify("Menú dividido → Subacción 1.");
  event.completed();
}
function accionSub2(event) {
  notify("Menú dividido → Subacción 2.");
  event.completed();
}

// Utilidad visual
function notify(msg) {
  if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
    console.log(msg);
  } else {
    console.log(msg);
  }
}

// Export global (Excel Online requiere global)
window.accionBasica = accionBasica;
window.accionIcono = accionIcono;
window.menuAccionA = menuAccionA;
window.menuAccionB = menuAccionB;
window.accionPrincipal = accionPrincipal;
window.accionSub1 = accionSub1;
window.accionSub2 = accionSub2;
