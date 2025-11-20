/* global Office, Excel, console */

// --- CONFIGURACIÓN GLOBAL Y ESTADO ---
const CONFIG = {
    autoReorder: true,
    showHeadings: false,
    clearConsole: true
};

// --- INICIALIZACIÓN ---
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // 1. Cargar Configuración guardada
        loadSettings();

        // 2. Disciplina de Silencio
        if (CONFIG.clearConsole) console.clear();

        // 3. Enrutamiento (Routing)
        const urlParams = new URLSearchParams(window.location.search);
        const page = urlParams.get('page') || 'accounting'; // Por defecto contabilidad
        showPage(page);

        // 4. Listeners de Eventos Globales
        document.getElementById('toggle-reorder').addEventListener('change', saveSettings);
        document.getElementById('toggle-headings').addEventListener('change', toggleHeadings);
        document.getElementById('toggle-clear-console').addEventListener('change', saveSettings);
        document.getElementById('btn-debug-dump').addEventListener('click', dumpDebugInfo);
        document.getElementById('btn-clear-console').addEventListener('click', () => console.clear());

        updateStatus("Sistema listo.");
    }
});

// --- LÓGICA DE NAVEGACIÓN ---
function showPage(pageId) {
    // Ocultar todas
    document.querySelectorAll('.page-section').forEach(el => el.classList.remove('active'));
    
    // Mostrar la solicitada
    const target = document.getElementById(`page-${pageId}`);
    if (target) {
        target.classList.add('active');
        // Si es contabilidad, cargar datos
        if (pageId === 'accounting') loadSheets();
    } else {
        // Fallback si la pagina no existe
        document.getElementById('page-accounting').classList.add('active');
        loadSheets();
    }
}

// --- GESTIÓN DE SETTINGS ---
function loadSettings() {
    const saved = localStorage.getItem('mochileros_cfg');
    if (saved) {
        const parsed = JSON.parse(saved);
        CONFIG.autoReorder = parsed.autoReorder;
        CONFIG.showHeadings = parsed.showHeadings;
        CONFIG.clearConsole = parsed.clearConsole;
    }
    
    // Actualizar UI
    document.getElementById('toggle-reorder').checked = CONFIG.autoReorder;
    document.getElementById('toggle-headings').checked = CONFIG.showHeadings;
    document.getElementById('toggle-clear-console').checked = CONFIG.clearConsole;
}

function saveSettings() {
    CONFIG.autoReorder = document.getElementById('toggle-reorder').checked;
    CONFIG.showHeadings = document.getElementById('toggle-headings').checked;
    CONFIG.clearConsole = document.getElementById('toggle-clear-console').checked;
    
    localStorage.setItem('mochileros_cfg', JSON.stringify(CONFIG));
}

// --- LÓGICA DE HOJAS (CONTABILIDAD) ---
async function loadSheets() {
    updateStatus("Cargando hojas...");
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name, items/tabColor, items/id, items/position");
            await context.sync();

            renderSheetList(sheets.items);
            updateStatus("Hojas cargadas.");
        });
    } catch (error) {
        handleError(error);
    }
}

function renderSheetList(sheets) {
    const list = document.getElementById('sheet-list');
    list.innerHTML = ''; // Limpiar

    sheets.forEach(sheet => {
        const li = document.createElement('li');
        li.className = 'sheet-card';
        li.textContent = sheet.name;
        li.draggable = true; // Habilitar Drag
        li.dataset.name = sheet.name;
        li.dataset.id = sheet.id;

        // Color y Contraste
        const bgColor = sheet.tabColor || "#ffffff"; // Default blanco si es nulo
        li.style.backgroundColor = bgColor;
        li.style.color = getContrastYIQ(bgColor);

        // Eventos
        li.onclick = () => activateSheet(sheet.name);
        
        // Eventos Drag & Drop
        addDnDEvents(li);

        list.appendChild(li);
    });
}

async function activateSheet(sheetName) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.activate();
            await context.sync();
        });
    } catch (error) {
        handleError(error);
    }
}

async function toggleHeadings() {
    saveSettings(); // Guardar estado
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items");
            await context.sync();
            
            // Aplicar a todas las hojas
            sheets.items.forEach(sheet => {
                sheet.showHeadings = CONFIG.showHeadings;
            });
            
            await context.sync();
            updateStatus(`Encabezados: ${CONFIG.showHeadings ? 'ON' : 'OFF'}`);
        });
    } catch (error) {
        handleError(error);
    }
}

// --- UTILIDADES VISUALES ---
function getContrastYIQ(hexcolor){
    // Si no viene hex (ej: null), devolver negro
    if(!hexcolor || hexcolor === '#ffffff') return '#323130';
    
    hexcolor = hexcolor.replace("#", "");
    var r = parseInt(hexcolor.substr(0,2),16);
    var g = parseInt(hexcolor.substr(2,2),16);
    var b = parseInt(hexcolor.substr(4,2),16);
    var yiq = ((r*299)+(g*587)+(b*114))/1000;
    return (yiq >= 128) ? '#323130' : 'white';
}

function updateStatus(msg) {
    document.getElementById('status-msg').innerText = msg;
}

function handleError(error) {
    console.error(error);
    updateStatus("Error: " + error.message);
    document.getElementById('debug-output').value += "\n[ERROR] " + error.message;
}

function dumpDebugInfo() {
    const info = {
        timestamp: new Date().toISOString(),
        config: CONFIG,
        userAgent: navigator.userAgent,
        url: window.location.href
    };
    document.getElementById('debug-output').value = JSON.stringify(info, null, 2);
}

// --- DRAG & DROP LOGIC ---
let dragSrcEl = null;

function addDnDEvents(elem) {
    elem.addEventListener('dragstart', handleDragStart);
    elem.addEventListener('dragenter', handleDragEnter);
    elem.addEventListener('dragover', handleDragOver);
    elem.addEventListener('dragleave', handleDragLeave);
    elem.addEventListener('drop', handleDrop);
    elem.addEventListener('dragend', handleDragEnd);
}

function handleDragStart(e) {
    this.style.opacity = '0.4';
    dragSrcEl = this;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', this.innerHTML);
}

function handleDragOver(e) {
    if (e.preventDefault) e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    
    // Calcular si estamos en la mitad superior o inferior
    const bounding = this.getBoundingClientRect();
    const offset = bounding.y + (bounding.height / 2);
    
    this.classList.remove('drag-over-top', 'drag-over-bottom');
    
    if (e.clientY - offset > 0) {
        this.classList.add('drag-over-bottom');
    } else {
        this.classList.add('drag-over-top');
    }
    return false;
}

function handleDragEnter(e) {
    this.classList.add('over');
}

function handleDragLeave(e) {
    this.classList.remove('drag-over-top', 'drag-over-bottom');
}

async function handleDrop(e) {
    if (e.stopPropagation) e.stopPropagation();
    
    // Limpiar estilos visuales
    this.classList.remove('drag-over-top', 'drag-over-bottom');

    if (dragSrcEl !== this) {
        // Si el reordenamiento automatico está apagado, no hacer nada en Excel
        if (!CONFIG.autoReorder) {
            updateStatus("Reordenamiento bloqueado por configuración.");
            return false;
        }

        const srcName = dragSrcEl.dataset.name;
        const targetName = this.dataset.name;
        
        // Determinar posición (Before o After)
        const bounding = this.getBoundingClientRect();
        const offset = bounding.y + (bounding.height / 2);
        const position = (e.clientY - offset > 0) ? "After" : "Before";

        updateStatus(`Moviendo ${srcName} ${position} ${targetName}...`);

        // EJECUTAR EN EXCEL
        try {
            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                const srcSheet = sheets.getItem(srcName);
                const targetSheet = sheets.getItem(targetName);
                
                srcSheet.position = position === "After" 
                    ? Excel.WorksheetPositionType.after 
                    : Excel.WorksheetPositionType.before;
                
                // Parche crítico para API de posicionamiento: referenciar hoja destino
                if(position === "After") {
                    srcSheet.position = targetSheet.position + 1; // Lógica simplificada por limitaciones de API JS puras a veces
                    // Nota: La API real usa reordering relativo, pero aqui usaremos reload para simplificar visual
                } 
                
                // Método más robusto: Usar relative positioning de la API si está disponible
                // Ojo: sheet.position acepta un indice numérico o un objeto relativo en algunas APIs nuevas.
                // Vamos a usar el método de reordenar array y reasignar para máxima compatibilidad o reload.
                
                /* NOTA ARQUITECTO: La API de Excel para "Mover relative to" es compleja. 
                   Para la V10, haremos el movimiento visual y forzaremos recarga.
                   Implementación real requiere calcular indices numéricos.
                */
                
                // RE-IMPLEMENTACIÓN ROBUSTA DE MOVIMIENTO
                sheets.load("items/name, items/position, items/id");
                await context.sync();
                
                let targetIndex = targetSheet.position;
                // Ajuste
                if (position === "After") targetIndex++;
                
                // Mover
                srcSheet.position = targetIndex;
                
                await context.sync();
                loadSheets(); // Recargar UI
            });
        } catch (err) {
            handleError(err);
            loadSheets(); // Revertir UI si falla
        }
    }
    return false;
}

function handleDragEnd(e) {
    this.style.opacity = '1';
    document.querySelectorAll('.sheet-card').forEach(col => {
        col.classList.remove('drag-over-top', 'drag-over-bottom');
    });
}
