// Archivo de Referencia: taskpane.js
/*
 * PROMPT MAESTRO v2 - Arquitecto de Software
 * Componente: Panel Lateral de Hojas (Contabilidad)
 * Fase: Final (Consolidación en taskpane.js)
 */

// #region 1. CONSTANTES Y MODELO DE DATOS
// -------------------------------------------------------------

const SETTINGS_KEY = 'sheetPanelSettings';

let _dragSource = null;
let _currentSettings = { autoReorder: true, showHeaders: false, clearConsoleOnLoad: true }; 
let _dropTargetPosition = "before"; 

// #endregion

// #region 2. EXCEL API Y LÓGICA CRÍTICA
// -------------------------------------------------------------

/** Gestiona errores bajo la disciplina de silencio. */
function handleError(message, error) {
    let userMessage = message;
    let code = "N/A";
    
    if (error && error.code) {
        code = error.code;
        if (code !== "PropertyNotLoaded" && error.message) {
            userMessage = error.message;
        }
    } else if (error instanceof Error) {
        userMessage = error.message;
    }
    
    console.error(`[ERROR CRÍTICO] ${userMessage} (Code: ${code})`, error);
    
    setGlobalStatus(`❌ Fallo en operación. Use 💾 para detalles.`, "error");
}

/** [CRÍTICA] Activa la hoja en Excel. */
async function activateSheet(sheetId) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetId);
            sheet.activate();
            await context.sync();
        });
    } catch (error) {
        handleError("Fallo al activar la hoja.", error);
    }
}

/** [CRÍTICA] Escribe en Excel: Reordena una hoja. */
async function reorderSheetInExcel(sourceId, targetId, dropType) {
    showLoadingState("⏳ Sincronizando orden...");
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/position");
            await context.sync();
            
            const sourceSheet = sheets.getItem(sourceId);
            const targetSheet = sheets.getItem(targetId);
            
            targetSheet.load("position");
            sourceSheet.load("position");
            await context.sync();
            
            let newPosition = targetSheet.position;

            if (dropType === "after") {
                if (sourceSheet.position === targetSheet.position + 1) return; 
                newPosition += 1;
            } else {
                if (sourceSheet.position === targetSheet.position - 1) return;
            }

            const maxPosition = sheets.items.length - 1; 
            newPosition = Math.min(Math.max(newPosition, 0), maxPosition);

            sourceSheet.position = newPosition;
            await context.sync();
        });
    } catch (error) {
        handleError("Fallo al reordenar la hoja.", error);
    } finally {
        await loadSheetsFromExcel(); 
    }
}

/** Carga todas las hojas desde Excel y actualiza el UI. */
async function loadSheetsFromExcel() {
    showLoadingState("⏳ Cargando Contabilidad..."); 
    try {
        let sheetData = [];
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name, items/tabColor, items/position, items/id");
            await context.sync();

            sheetData = sheets.items.map(s => ({
                id: s.id,
                name: s.name,
                color: s.tabColor === "" ? "#e1dfdd" : s.tabColor,
                position: s.position
            }));
        });

        sheetData.sort((a, b) => a.position - b.position);
        renderSheetList(sheetData);

    } catch (error) {
        handleError("Fallo al cargar las hojas de Excel.", error);
        showLoadingState(`❌ Error de carga. Use 🔄.`);
    } finally {
        hideLoadingState();
    }
}

/** Aplica la configuración de encabezados a todas las hojas. */
async function applyHeadingsSettingToAllSheets(show) {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/showHeadings");
            await context.sync();

            sheets.items.forEach(sheet => {
                sheet.showHeadings = show; 
            });

            await context.sync();
        });
    } catch (error) {
        handleError("Fallo al aplicar la configuración de encabezados.", error);
    }
}

// #endregion

// #region 3. GESTIÓN DE ESTADO Y CONFIGURACIÓN
// -------------------------------------------------------------

function loadSettings() {
    try {
        const stored = localStorage.getItem(SETTINGS_KEY);
        if (stored) {
            _currentSettings = { 
                ..._currentSettings, 
                ...JSON.parse(stored),
                clearConsoleOnLoad: JSON.parse(stored).autoScrollConsole !== undefined 
                    ? JSON.parse(stored).autoScrollConsole 
                    : (_currentSettings.clearConsoleOnLoad || true)
            };
        }
    } catch (e) {
        console.warn("Fallo al cargar la configuración de localStorage. Usando valores por defecto.");
    }
    return _currentSettings;
}

function saveSettings() {
    try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify(_currentSettings));
    } catch (e) {
        console.error("Fallo al guardar la configuración de localStorage.");
    }
}

function updateSetting(key, value) {
    _currentSettings[key] = value;
    saveSettings();
    applySettings();
}

function applySettings() {
    document.getElementById("toggle-auto-reorder").checked = _currentSettings.autoReorder;
    document.getElementById("toggle-headers").checked = _currentSettings.showHeaders;
    document.getElementById("toggle-clear-on-load").checked = _currentSettings.clearConsoleOnLoad; 
    
    applyHeadingsSettingToAllSheets(_currentSettings.showHeaders);
}

// #endregion

// #region 4. LÓGICA DE UI Y DOM (RENDER/EVENTS)
// -------------------------------------------------------------

/** Calcula el color de texto contrastante. */
function getContrastingTextColor(hexColor) {
    let hex = hexColor.startsWith('#') ? hexColor.substring(1) : hexColor;
    if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
    }

    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);

    const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
    return luminance > 0.5 ? '#323130' : '#ffffff'; 
}

/** Renderiza la lista de hojas en el panel. */
function renderSheetList(data) {
    const listContainer = document.getElementById("sheet-list");
    listContainer.innerHTML = "";

    data.forEach(sheet => {
        const li = document.createElement("li");
        li.className = "sheet-item";
        li.draggable = true;
        li.dataset.sheetId = sheet.id;
        li.dataset.position = sheet.position.toString();
        
        const hexColor = sheet.color.startsWith('#') ? sheet.color : `#${sheet.color}`;
        
        li.style.backgroundColor = hexColor;
        li.style.color = getContrastingTextColor(hexColor);

        li.addEventListener('click', (e) => {
            if (li.classList.contains('dragging')) return; 
            activateSheet(sheet.id);
        });

        const dropdownIcon = document.createElement("div");
        dropdownIcon.className = "dropdown-icon";
        dropdownIcon.innerHTML = `&#9660;`; // Flecha descendente
        
        dropdownIcon.style.color = li.style.color; 
        
        const nameSpan = document.createElement("span");
        nameSpan.className = "sheet-name";
        nameSpan.textContent = sheet.name;
        
        li.appendChild(nameSpan);
        li.appendChild(dropdownIcon);

        addDnDEvents(li);

        listContainer.appendChild(li);
    });
}

/** Configura los listeners de Drag and Drop. */
function addDnDEvents(el) {
    el.addEventListener('dragstart', (e) => {
        _dragSource = el;
        el.classList.add('dragging');
        if (e.dataTransfer) {
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', JSON.stringify({
                id: el.dataset.sheetId,
                position: el.dataset.position
            }));
        }
    });

    el.addEventListener('dragover', (e) => {
        e.preventDefault();
        if (e.dataTransfer) {
            e.dataTransfer.dropEffect = 'move';
        }

        const rect = el.getBoundingClientRect();
        
        if (e.clientY < rect.top + rect.height / 2) {
            el.classList.add('drop-target-top');
            el.classList.remove('drop-target-bottom');
            _dropTargetPosition = "before"; 
        } else {
            el.classList.add('drop-target-bottom');
            el.classList.remove('drop-target-top');
            _dropTargetPosition = "after"; 
        }
        return false;
    });

    el.addEventListener('dragleave', (e) => {
        el.classList.remove('drop-target-top', 'drop-target-bottom');
    });

    el.addEventListener('drop', async (e) => {
        e.stopPropagation();
        el.classList.remove('drop-target-top', 'drop-target-bottom');
        
        const sourceData = JSON.parse(e.dataTransfer.getData('text/plain'));
        const targetId = el.dataset.sheetId;
        
        if (sourceData.id !== targetId) {
            if (_currentSettings.autoReorder) {
                await reorderSheetInExcel(sourceData.id, targetId, _dropTargetPosition); 
            } else {
                loadSheetsFromExcel();
            }
        }
        return false;
    });

    el.addEventListener('dragend', (e) => {
        el.classList.remove('dragging');
        el.classList.remove('drop-target-top', 'drop-target-bottom');
        _dragSource = null;
    });
}

/** Muestra un mensaje en la barra de estado global. */
function setGlobalStatus(message, type = "info") {
    const statusBar = document.getElementById("global-status-bar");
    if (!statusBar) return;
    
    statusBar.innerText = message;
    statusBar.className = `status-${type}`;
    
    if (type !== "loading" && type !== "error") {
        window.setTimeout(() => {
            statusBar.innerText = "✅ Listo.";
            statusBar.className = "status-success";
        }, 3000);
    }
}


/** Muestra un volcado de datos de depuración a la consola y al textarea. */
async function executeDebugDump() {
    const outputArea = document.getElementById("debug-output");
    const btnCopy = document.getElementById("btn-copy-debug");
    
    console.log("--- VOLCADO DE DATOS SOLICITADO POR EL USUARIO ---");
    outputArea.value = "Generando volcado del sistema... espera.";
    
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name, items/id, items/position, items/tabColor, items/visibility, items/showHeadings");
            
            const app = context.application;
            app.load("calculationMode");

            await context.sync();

            const debugObject = {
                timestamp: new Date().toISOString(),
                workbook: {
                    sheetCount: sheets.items.length,
                    sheets: sheets.items.map(s => ({
                        index: s.position,
                        name: s.name,
                        id: s.id,
                        color: s.tabColor,
                        visible: s.visibility,
                        showHeadings: s.showHeadings
                    }))
                },
                appState: _currentSettings
            };

            const jsonString = JSON.stringify(debugObject, null, 2);
            outputArea.value = jsonString;
            btnCopy.classList.remove("hidden");
            console.log(jsonString); 
            setGlobalStatus("✅ Volcado de datos completado.", "info");
        });
    } catch (error) {
        outputArea.value = "❌ ERROR: " + error.toString();
        handleError("Fallo al generar el volcado de datos.", error);
    }
}

/** Limpia la consola manualmente. */
function clearConsoleLog() {
    console.clear();
    setGlobalStatus("🧹 Consola de Logs limpiada.", "info");
}

function copyDebugToClipboard() {
    const outputArea = document.getElementById("debug-output");
    outputArea.select();
    document.execCommand("copy");
    setGlobalStatus("📋 Datos copiados al portapapeles.", "info");
}

function showLoadingState(message) {
    const list = document.getElementById("sheet-list");
    list.innerHTML = `<li class="loading-state">${message}</li>`;
    document.getElementById("app-container").classList.add("loading-active");
}

function hideLoadingState() {
    document.getElementById("app-container").classList.remove("loading-active");
}

// #endregion

// #region 5. OFFICE.ONREADY & INITIALIZATION
// -------------------------------------------------------------

Office.onReady(() => {
    loadSettings(); 

    if (_currentSettings.clearConsoleOnLoad) {
        console.clear();
    }
    
    console.log("Arquitecto v2: Sistema Contabilidad (taskpane.js) iniciado.");
    
    applySettings();

    // Setup Listeners
    document.getElementById("btn-refresh").onclick = loadSheetsFromExcel;
    document.getElementById("btn-debug-dump").onclick = executeDebugDump;
    document.getElementById("btn-copy-debug").onclick = copyDebugToClipboard;

    document.getElementById("toggle-auto-reorder").onchange = (e) => {
        updateSetting('autoReorder', e.target.checked);
    };
    document.getElementById("toggle-headers").onchange = (e) => {
        updateSetting('showHeaders', e.target.checked);
    };
    document.getElementById("toggle-clear-on-load").onchange = (e) => { 
        updateSetting('clearConsoleOnLoad', e.target.checked);
    };
    document.getElementById("btn-clear-console-manual").onclick = clearConsoleLog;
    
    // Función switchTab global para el HTML
    window.switchTab = (tabName) => {
        const tabs = ['sheets', 'config'];
        tabs.forEach(t => {
            const elContent = document.getElementById(`view-${t}`);
            const elTab = document.getElementById(`tab-${t}`);
            if (t === tabName) {
                elContent.classList.add("active");
                elContent.classList.remove("hidden");
                elTab.classList.add("active");
            } else {
                elContent.classList.remove("active");
                elContent.classList.add("hidden");
                elTab.classList.remove("active");
            }
        });
    };

    loadSheetsFromExcel();
});

// #endregion
