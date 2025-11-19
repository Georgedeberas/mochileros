/* global Office, console */

Office.onReady((info) => {
    // Verificamos si estamos dentro de Excel
    if (info.host === Office.HostType.Excel) {
        console.log("Mochileros RD: Entorno de Office cargado correctamente.");
        document.getElementById("status-message").innerText = "Sistema listo.";
    } else {
        console.error("Mochileros RD: Error, no se detectó el host de Excel.");
    }
});
