/* global Office, Excel */

(function () {
  const COLOR_OPEN = "#1f7a3f";
  const COLOR_CANCELLED = "#0b4f8c";
  const COLOR_CLOSED = "#000000";
  const COLOR_SYSTEM = "#fcff3c";

  const STATE_FROM_COLOR = {};
  STATE_FROM_COLOR[COLOR_OPEN.toLowerCase()] = "open";
  STATE_FROM_COLOR[COLOR_CANCELLED.toLowerCase()] = "cancelled";
  STATE_FROM_COLOR[COLOR_CLOSED.toLowerCase()] = "closed";
  STATE_FROM_COLOR[COLOR_SYSTEM.toLowerCase()] = "system";

  const COLOR_FROM_STATE = {
    open: COLOR_OPEN,
    cancelled: COLOR_CANCELLED,
    closed: COLOR_CLOSED,
    system: COLOR_SYSTEM,
    unclassified: null
  };

  let allSheets = [];
  let filterText = "";
  let filterState = "all";

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      wireUi();
      refreshFromWorkbook();
    }
  });

  function wireUi() {
    const btnRefresh = document.getElementById("btnRefresh");
    const chkAuto = document.getElementById("chkAuto");
    const newToggle = document.getElementById("newToggle");
    const newArrow = document.getElementById("newArrow");
    const btnCreate = document.getElementById("btnCreate");
    const txtSearch = document.getElementById("txtSearch");
    const selState = document.getElementById("selState");

    if (btnRefresh) {
      btnRefresh.addEventListener("click", function () {
        refreshFromWorkbook();
      });
    }

    if (chkAuto) {
      chkAuto.addEventListener("change", function () {
        if (chkAuto.checked) {
          attachSheetEvents();
        } else {
          detachSheetEvents();
        }
      });
    }

    if (newToggle) {
      newToggle.addEventListener("click", toggleNewCard);
    }
    if (newArrow) {
      newArrow.addEventListener("click", function (ev) {
        ev.stopPropagation();
        toggleNewCard();
      });
    }

    if (btnCreate) {
      btnCreate.addEventListener("click", createTourFromForm);
    }

    if (txtSearch) {
      txtSearch.addEventListener("input", function () {
        filterText = txtSearch.value || "";
        renderMenu();
      });
    }

    if (selState) {
      selState.addEventListener("change", function () {
        filterState = selState.value || "all";
        renderMenu();
      });
    }
  }

  function setStatus(msg) {
    const el = document.getElementById("status");
    if (el) {
      el.textContent = msg;
    }
  }

  function toggleNewCard() {
    const body = document.getElementById("newBody");
    const arrow = document.getElementById("newArrow");
    if (!body) return;
    const hidden = body.style.display === "none";
    body.style.display = hidden ? "block" : "none";
    if (arrow) {
      arrow.textContent = hidden ? "▾" : "▸";
    }
  }

  async function refreshFromWorkbook() {
    setStatus("Leyendo hojas...");
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name,items/tabColor,items/position");
        await context.sync();

        const now = new Date();
        const curYear = now.getFullYear();
        const curMonth = now.getMonth() + 1;

        allSheets = sheets.items.map((ws) => {
          const info = parseName(ws.name, curYear, curMonth);
          const state = getStateFromColor(ws.tabColor);
          return {
            name: ws.name,
            tabColor: ws.tabColor || null,
            state: state,
            day: info.day,
            month: info.month,
            year: info.year
          };
        });
      });

      renderMenu();
      setStatus("Listo");
    } catch (err) {
      console.error(err);
      setStatus("Error leyendo el libro");
    }
  }

  function parseName(name, curYear, curMonth) {
    // Sufijo tipo " 30.12" al final
    const m = /(\d{1,2})\.(\d{1,2})\s*$/.exec(name);
    if (!m) {
      return { day: 0, month: 0, year: 0 };
    }
    const day = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    let year = 0;
    if (month >= 1 && month <= 12) {
      // Regla B: si mes < mes actual => asumimos año siguiente
      year = month < curMonth ? curYear + 1 : curYear;
    }
    return { day: day, month: month, year: year };
  }

  function getStateFromColor(color) {
    if (!color) return "unclassified";
    const key = String(color).toLowerCase();
    return STATE_FROM_COLOR[key] || "unclassified";
  }

  function getMonthLabel(m, y) {
    const meses = [
      "",
      "Enero",
      "Febrero",
      "Marzo",
      "Abril",
      "Mayo",
      "Junio",
      "Julio",
      "Agosto",
      "Septiembre",
      "Octubre",
      "Noviembre",
      "Diciembre"
    ];
    if (!m || !y) return "Sin fecha";
    return meses[m] + " " + y;
  }

  function normalize(text) {
    if (!text) return "";
    return text
      .toString()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function applyFilters(list) {
    let result = list.slice();

    // Filtro por estado
    if (filterState && filterState !== "all") {
      result = result.filter((s) => s.state === filterState);
    }

    // Filtro por texto: AND de palabras, insensitive y sin acentos
    const q = normalize(filterText).trim();
    if (q) {
      const tokens = q.split(/\s+/).filter(Boolean);
      result = result.filter((s) => {
        const haystack = normalize(
          s.name + " " + getMonthLabel(s.month, s.year) + " " + s.state
        );
        return tokens.every((t) => haystack.indexOf(t) !== -1);
      });
    }

    return result;
  }

  function groupByMonth(list) {
    const groups = {};

    list.forEach((s) => {
      let key;
      if (s.state === "system") {
        key = "Sistema";
      } else if (!s.month || !s.year) {
        key = "Sin fecha";
      } else {
        key = getMonthLabel(s.month, s.year);
      }
      if (!groups[key]) {
        groups[key] = {
          key: key,
          items: []
        };
      }
      groups[key].items.push(s);
    });

    // Orden de grupos: Sistema, meses por fecha, Sin fecha
    const monthGroups = [];
    let sistemaGroup = null;
    let sinFechaGroup = null;

    Object.keys(groups).forEach((k) => {
      if (k === "Sistema") {
        sistemaGroup = groups[k];
      } else if (k === "Sin fecha") {
        sinFechaGroup = groups[k];
      } else {
        monthGroups.push(groups[k]);
      }
    });

    // Parsear "Mes año" a (año, mes) para ordenar
    monthGroups.sort((a, b) => {
      const [mesA, yearA] = a.key.split(" ");
      const [mesB, yearB] = b.key.split(" ");
      const ma = monthNameToNumber(mesA);
      const mb = monthNameToNumber(mesB);
      const ya = parseInt(yearA, 10) || 0;
      const yb = parseInt(yearB, 10) || 0;
      if (ya !== yb) return ya - yb;
      return ma - mb;
    });

    const ordered = [];
    if (sistemaGroup) ordered.push(sistemaGroup);
    ordered.push.apply(ordered, monthGroups);
    if (sinFechaGroup) ordered.push(sinFechaGroup);
    return ordered;
  }

  function monthNameToNumber(name) {
    const map = {
      enero: 1,
      febrero: 2,
      marzo: 3,
      abril: 4,
      mayo: 5,
      junio: 6,
      julio: 7,
      agosto: 8,
      septiembre: 9,
      setiembre: 9,
      octubre: 10,
      noviembre: 11,
      diciembre: 12
    };
    const n = normalize(name);
    return map[n] || 0;
  }

  function renderMenu() {
    const container = document.getElementById("menu");
    if (!container) return;
    container.innerHTML = "";

    if (!allSheets || allSheets.length === 0) {
      container.textContent = "No se encontraron hojas.";
      return;
    }

    const filtered = applyFilters(allSheets);
    const groups = groupByMonth(filtered);

    if (!groups.length) {
      container.textContent = "No hay resultados para el filtro actual.";
      return;
    }

    groups.forEach((group) => {
      const groupEl = document.createElement("div");
      groupEl.className = "mochi-month";

      const header = document.createElement("div");
      header.className = "mochi-month-header";

      const title = document.createElement("div");
      title.className = "mochi-month-title";
      title.textContent = group.key;

      const arrow = document.createElement("div");
      arrow.className = "mochi-month-arrow";
      arrow.textContent = "▾";

      header.appendChild(title);
      header.appendChild(arrow);

      const body = document.createElement("div");
      body.className = "mochi-month-body";

      // Ordenar dentro del grupo por fecha (día) y nombre
      group.items.sort((a, b) => {
        if (a.year !== b.year) return a.year - b.year;
        if (a.month !== b.month) return a.month - b.month;
        if (a.day !== b.day) return a.day - b.day;
        return a.name.localeCompare(b.name, "es");
      });

      group.items.forEach((sheetInfo) => {
        const card = buildSheetCard(sheetInfo);
        body.appendChild(card);
      });

      // Colapsar / expandir
      header.addEventListener("click", () => {
        const hidden = body.style.display === "none";
        body.style.display = hidden ? "block" : "none";
        arrow.textContent = hidden ? "▾" : "▸";
      });

      groupEl.appendChild(header);
      groupEl.appendChild(body);
      container.appendChild(groupEl);
    });
  }

  function buildSheetCard(sheet) {
    const el = document.createElement("div");
    el.className = "mochi-sheet";

    const main = document.createElement("div");
    main.className = "mochi-sheet-main";

    const dot = document.createElement("span");
    dot.className = "mochi-dot " + classForState(sheet.state);

    const names = document.createElement("div");
    names.className = "mochi-sheet-names";

    const title = document.createElement("div");
    title.className = "mochi-sheet-name";
    title.textContent = sheet.name;

    const meta = document.createElement("div");
    meta.className = "mochi-sheet-meta";
    const label = getMonthLabel(sheet.month, sheet.year);
    meta.textContent = label === "Sin fecha" ? sheet.state : label;

    names.appendChild(title);
    names.appendChild(meta);

    main.appendChild(dot);
    main.appendChild(names);

    const controls = document.createElement("div");
    controls.className = "mochi-sheet-controls";

    const sel = document.createElement("select");
    sel.className = "mochi-state-select";
    [
      { value: "open", label: "Abierto" },
      { value: "cancelled", label: "Cancelado" },
      { value: "closed", label: "Cerrado" },
      { value: "system", label: "Sistema" },
      { value: "unclassified", label: "Sin clasificar" }
    ].forEach((opt) => {
      const o = document.createElement("option");
      o.value = opt.value;
      o.textContent = opt.label;
      if (opt.value === sheet.state) {
        o.selected = true;
      }
      sel.appendChild(o);
    });

    sel.addEventListener("change", function (ev) {
      ev.stopPropagation();
      const newState = sel.value;
      updateSheetState(sheet.name, newState);
    });

    controls.appendChild(sel);

    el.appendChild(main);
    el.appendChild(controls);

    el.addEventListener("click", function () {
      activateSheet(sheet.name);
    });

    return el;
  }

  function classForState(state) {
    switch (state) {
      case "open":
        return "mochi-dot-open";
      case "cancelled":
        return "mochi-dot-cancelled";
      case "closed":
        return "mochi-dot-closed";
      case "system":
        return "mochi-dot-system";
      default:
        return "mochi-dot-unclassified";
    }
  }

  async function activateSheet(name) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(name);
        sheet.activate();
        await context.sync();
      });
      setStatus('Hoja "' + name + '" activada');
    } catch (err) {
      console.error(err);
      setStatus("No se pudo activar la hoja");
    }
  }

  async function updateSheetState(name, newState) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(name);
        const color = COLOR_FROM_STATE[newState] || null;
        sheet.tabColor = color;
        await context.sync();
      });

      const item = allSheets.find((s) => s.name === name);
      if (item) {
        item.state = newState;
        item.tabColor = COLOR_FROM_STATE[newState] || null;
      }
      renderMenu();
      setStatus("Estado actualizado");
    } catch (err) {
      console.error(err);
      setStatus("Error actualizando estado");
    }
  }

  async function createTourFromForm() {
    const nameInput = document.getElementById("newName");
    const templateSelect = document.getElementById("newTemplate");
    const dateInput = document.getElementById("newDate");

    const baseName = ((nameInput && nameInput.value) || "").trim();
    const template = templateSelect && templateSelect.value;
    const dateStr = dateInput && dateInput.value;

    if (!baseName) {
      setStatus("Pon un nombre para el tour.");
      return;
    }
    if (!template) {
      setStatus("Selecciona una plantilla.");
      return;
    }
    if (!dateStr) {
      setStatus("Selecciona una fecha.");
      return;
    }

    const date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      setStatus("Fecha no válida.");
      return;
    }

    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const newSheetName = baseName + " " + day + "." + month;

    setStatus("Creando tour...");

    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let templateSheet = null;

        try {
          templateSheet = sheets.getItem(template);
          templateSheet.load("name");
        } catch (e) {
          templateSheet = null;
        }

        let newSheet;
        if (templateSheet) {
          // Copiar la plantilla al final
          newSheet = templateSheet.copy("End");
          newSheet.load("name");
        } else {
          // Hoja en blanco
          newSheet = sheets.add(newSheetName);
          newSheet.load("name");
        }

        await context.sync();

        if (newSheet.name !== newSheetName) {
          newSheet.name = newSheetName;
        }

        // Abierto (verde)
        newSheet.tabColor = COLOR_OPEN;

        await context.sync();
      });

      if (nameInput) nameInput.value = "";
      if (dateInput) dateInput.value = "";

      await refreshFromWorkbook();
      setStatus("Tour creado");
    } catch (err) {
      console.error(err);
      setStatus("No se pudo crear el tour");
    }
  }

  // Eventos de hojas (Auto)
  let sheetAddedHandler = null;
  let sheetDeletedHandler = null;

  async function attachSheetEvents() {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheetAddedHandler = sheets.onAdded.add(onSheetsChanged);
        sheetDeletedHandler = sheets.onDeleted.add(onSheetsChanged);
        await context.sync();
      });
      setStatus("Auto: activado");
    } catch (err) {
      console.warn("Eventos de hojas no disponibles en esta versión.", err);
      setStatus("Auto no disponible en este Excel");
    }
  }

  async function detachSheetEvents() {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        if (sheetAddedHandler) {
          sheets.onAdded.remove(sheetAddedHandler);
          sheetAddedHandler = null;
        }
        if (sheetDeletedHandler) {
          sheets.onDeleted.remove(sheetDeletedHandler);
          sheetDeletedHandler = null;
        }
        await context.sync();
      });
      setStatus("Auto: desactivado");
    } catch (err) {
      console.error(err);
    }
  }

  async function onSheetsChanged() {
    await refreshFromWorkbook();
  }
})();
