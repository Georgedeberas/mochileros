/* global Office */

/**
 * Inicialización del contexto de Office.
 * No es obligatorio hacer nada aquí, pero es buen lugar para logs futuros.
 */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    console.log("[Mochileros RD] Commands.js listo en contexto Excel.");
  } else {
    console.log("[Mochileros RD] Commands.js cargado en host:", info.host);
  }
});

/**
 * Función asociada al botón "Actualizar sistema" en la cinta.
 *
 * IMPORTANTE (limitaciones actuales):
 * - No puede reinstalar el add-in ni cambiar el manifiesto.
 * - No puede cambiar el GUID ni hacer auto-deploy.
 *
 * Qué sí hace:
 * - Registra en consola que el usuario pidió actualizar.
 * - Escribe un mensaje en la celda seleccionada como confirmación visual.
 * - Deja el gancho preparado para que en el futuro (shared runtime, etc.)
 *   podamos integrar lógica más avanzada (versión, recargas, diálogos, etc.).
 *
 * @param {Office.AddinCommands.Event} event
 */
function mochilerosUpdateSystem(event) {
  try {
    console.log("[Mochileros RD] Botón 'Actualizar sistema' ejecutado.");

    var mensaje =
      "Mochileros RD - Actualizar sistema:\n" +
      "- Se ha solicitado actualizar el complemento.\n" +
      "- Si publicaste cambios en GitHub (index.html, app.js, etc.), " +
      "ciérralo y vuelve a abrir el panel para usar la última versión.\n" +
      "- Para cambios en la cinta (botones/menús), recuerda que debes instalar un nuevo manifiesto.";

    // Escribimos el mensaje en la celda seleccionada como feedback visual
    Office.context.document.setSelectedDataAsync(
      mensaje,
      { coercionType: Office.CoercionType.Text },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(
            "[Mochileros RD] Error al escribir mensaje de actualización:",
            asyncResult.error
          );
        } else {
          console.log(
            "[Mochileros RD] Mensaje de actualización escrito correctamente en la hoja."
          );
        }

        // Muy importante: liberar el botón de la cinta
        if (event && typeof event.completed === "function") {
          event.completed();
        }
      }
    );
  } catch (error) {
    console.error("[Mochileros RD] Error en mochilerosUpdateSystem:", error);

    // Asegurarse de no dejar el botón bloqueado aún si falla algo
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}

/**
 * Asociación de la función al nombre usado en el manifiesto.
 * Debe coincidir con FunctionName en manifest_mochileros_rd.xml:
 *   <FunctionName>mochilerosUpdateSystem</FunctionName>
 */
if (Office && Office.actions && Office.actions.associate) {
  Office.actions.associate("mochilerosUpdateSystem", mochilerosUpdateSystem);
}
