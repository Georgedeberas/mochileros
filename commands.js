Office.onReady(() => {
  console.log("✅ Mochileros RD Add-in - Comandos cargados correctamente.");
});

function accionBasica(event) {
  alert("🔘 Botón simple presionado.");
  event.completed();
}

function accionIcono(event) {
  alert("🟠 Botón con ícono ejecutado.");
  event.completed();
}

function menuAccionA(event) {
  alert("Seleccionaste: Opción A del menú desplegable.");
  event.completed();
}

function menuAccionB(event) {
  alert("Seleccionaste: Opción B del menú desplegable.");
  event.completed();
}

function accionPrincipal(event) {
  alert("Menú dividido → Acción principal ejecutada.");
  event.completed();
}

function accionSub1(event) {
  alert("Subacción 1 ejecutada.");
  event.completed();
}

function accionSub2(event) {
  alert("Subacción 2 ejecutada.");
  event.completed();
}

function colorRojo(event) {
  alert("🎨 Color Rojo seleccionado.");
  event.completed();
}

function colorVerde(event) {
  alert("🎨 Color Verde seleccionado.");
  event.completed();
}

function colorAzul(event) {
  alert("🎨 Color Azul seleccionado.");
  event.completed();
}

function accionFinal(event) {
  alert("✅ Acción final completada correctamente.");
  event.completed();
}
