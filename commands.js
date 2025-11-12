/* global Office */

Office.onReady(function () {
  // No necesitamos inicializar nada especial aquí por ahora.
});

/**
 * Botón "Actualizar sistema" en la cinta.
 *
 * IMPORTANTE:
 * Sin shared runtime, desde este contexto NO podemos forzar un recargado
 * profundo del panel. Lo que sí puedes hacer es:
 *  - Publicar cambios en GitHub Pages.
 *  - Cerrar y volver a abrir el panel (o el libro).
 *
 * Este gancho queda preparado para futuras mejoras (diálogos, etc.).
 */
function mochilerosUpdateSystem(event) {
  try {
    console.log("[Mochileros RD] Solicitud de actualización del sistema.");
  } catch (e) {
    console.error(e);
  } finally {
    // SIEMPRE llamar a completed() para liberar el botón.
    event.completed();
  }
}

// Asociación por nombre (por compatibilidad con el modelo actions).
if (Office && Office.actions && Office.actions.associate) {
  Office.actions.associate("mochilerosUpdateSystem", mochilerosUpdateSystem);
}
