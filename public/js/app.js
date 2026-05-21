// ─── Variables Globales ────────────────────────────────────────────────────────
let modifiedData = {};
let masterData = [];

// ─── Inicialización ───────────────────────────────────────────────────────────
// initApp() ya NO se llama aquí automáticamente.
// Lo invoca el script inline de index.html DESPUÉS de que requireAuth() confirma sesión.
// Así se garantiza que nunca se cargan datos ni se pinta la UI si el usuario no está autenticado.

const days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

// Días tal como los espera el servidor (sin tilde en Miércoles)
const DIAS_API = {
  Lunes: "Lunes",
  Martes: "Martes",
  Miércoles: "Miercoles",
  Jueves: "Jueves",
  Viernes: "Viernes",
};

const Fases = ["UTE", "ULA 1", "ULA 2", "TSF", "Otra"];
const Programas = [
  "Hotelería",
  "Cocina",
  "Belleza",
  "Auxiliar Administrativo",
  "Otro",
];

const diseaseOptions = [
  { value: "SANOS", text: "SANOS (ingresados a CIDI)" },
  { value: "IRA", text: "IRA (gripes, cuadros virales)" },
  { value: "ALERGIAS", text: "ALERGIAS (respiratoria, piel, medicamentos)" },
  { value: "BROTES", text: "BROTES (escabiosis o contagiosos)" },
  { value: "EDA", text: "EDA (enfermedad diarreica aguda)" },
  { value: "VOMITOS", text: "VÓMITOS" },
  { value: "FIEBRE", text: "FIEBRE" },
  { value: "ACCIDENTE CASERO", text: "ACCIDENTE CASERO" },
  { value: "SITUACION PERSONAL", text: "SITUACIÓN PERSONAL" },
  { value: "ASISTE A FAMI", text: "ASISTE A FAMI" },
  { value: "CITA MEDICA / VACUNAS", text: "CITA MÉDICA / VACUNAS" },
  { value: "HOSPITALIZACION", text: "HOSPITALIZACIÓN" },
  { value: "OTROS", text: "OTROS (Transportes, mamá enferma)" },
];

// Columnas para exportar al Excel
const columnOrder = [
  "NombreBebe",
  "NombreMadre",
  "Fase",
  "ProgramaMadre",
  "Edad",
  "Asistencia",
  "Ubicacion",
  "Reporte",
  "SituacionEspecifica",
  "Nota",
  "Extras",
  "NoCidi",
];
const columnHeaders = {
  NombreBebe: "Nombre Bebé",
  NombreMadre: "Nombre Madre",
  Fase: "Fase",
  ProgramaMadre: "Programa",
  Edad: "Edad (meses)",
  Asistencia: "Asistencia",
  Ubicacion: "Ubicación",
  Reporte: "Reporte",
  SituacionEspecifica: "Situación Específica",
  Nota: "Nota",
  Extras: "Extras",
  NoCidi: "No CIDI",
};

// ─── Mapa de badges por programa ─────────────────────────────────────────────
const PROGRAMA_BADGE = {
  Cocina: "badge-cocina",
  Hotelería: "badge-hoteleria",
  Belleza: "badge-belleza",
  "Auxiliar Administrativo": "badge-auxiliar",
  Otro: "badge-otro",
  Hoteleria: "badge-hoteleria", // alias sin tilde por si viene así de BD
};

function badgePrograma(programa) {
  const clase = PROGRAMA_BADGE[programa] || "badge-vacio";
  const texto = programa || "—";
  return `<span class="badge-programa ${clase}">${texto}</span>`;
}

// ─── Elementos del DOM ────────────────────────────────────────────────────────
const daysTabs = document.getElementById("days-tabs");
const tabsContent = document.getElementById("tabs-content");
const exportBtn = document.getElementById("export-btn");
const searchInput = document.getElementById("search-input");

async function initApp() {
  loadFromLocalStorage();
  _lockTable(true);
  await loadMasterDataFromServer();
  _lockTable(false);
  openCurrentDayTab();
  setupEventListeners();
  addAddBabyButton();
}

/**
 * _lockTable(on)
 * Deshabilita o habilita la interacción con la tabla de asistencia.
 * Evita que la profesora toque una fila mientras los datos del servidor
 * aún no terminaron de renderizarse (race condition con realIndex).
 */
function _lockTable(on) {
  const content = document.getElementById("tabs-content");
  if (!content) return;
  content.style.pointerEvents = on ? "none" : "";
  content.style.opacity = on ? "0.55" : "";
}

function setupEventListeners() {
  exportBtn.addEventListener("click", () => {
    // Mostrar modal de confirmación antes de exportar
    const activeTab = document.querySelector(".tab.active");
    if (!activeTab) {
      alert("Por favor, selecciona un día primero.");
      return;
    }
    const dayToExport = activeTab.dataset.day;
    const dataToExport = modifiedData[dayToExport] || [];
    if (dataToExport.length === 0) {
      alert(`No hay datos para exportar en el día ${dayToExport}`);
      return;
    }

    const total = dataToExport.length;
    const presentes = dataToExport.filter(
      (r) => r.Asistencia === "Sí" || r.Asistencia === "Si",
    ).length;
    const ausentes = total - presentes;

    document.getElementById("exportConfirmMsg").innerHTML =
      `Vas a exportar el listado del <strong>${dayToExport}</strong> con:<br>
       <strong>${total}</strong> bebés registrados &nbsp;·&nbsp;
       <strong style="color:#2e7d32">${presentes}</strong> presentes &nbsp;·&nbsp;
       <strong style="color:#c62828">${ausentes}</strong> ausentes`;

    const modal = document.getElementById("exportConfirmModal");
    modal.style.display = "flex";

    document.getElementById("exportConfirmSi").onclick = () => {
      modal.style.display = "none";
      exportToExcel();
    };
    document.getElementById("exportConfirmNo").onclick = () => {
      modal.style.display = "none";
    };
  });
  searchInput.addEventListener("input", filterData);
}

// ─── Carga desde el servidor (Supabase vía API) ───────────────────────────────
async function loadMasterDataFromServer() {
  updateSyncStatus("loading", "Cargando listado desde base de datos...");
  try {
    // ── 1. Cargar masterData completo desde /api/bebes (todos los bebés con datos completos)
    const resMaster = await authFetch("/api/bebes");
    if (resMaster.ok) {
      const json = await resMaster.json();
      masterData = (json.bebes || []).map((b) => ({
        NombreBebe: b.NombreBebe || "",
        NombreMadre: b.NombreMadre || "",
        Fase: b.Fase || b.InstitucionMadre || "",
        ProgramaMadre: b.ProgramaMadre || "",
        Edad: normalizarEdad(b.Edad || ""),
      }));
    }

    // ── 2. Cargar listado por día para armar las tabs de asistencia
    const results = await Promise.all(
      Object.entries(DIAS_API).map(async ([dayName, apiDia]) => {
        const res = await authFetch(`/api/sheet/${apiDia}`);
        if (!res.ok) throw new Error(`Error cargando ${dayName}`);
        const text = await res.text();
        return { dayName, rows: parseCsv(text) };
      }),
    );

    // ── Si ya hay datos guardados hoy en localStorage, conservarlos ──────────
    const hayDatosGuardados = Object.keys(modifiedData).length > 0;

    if (hayDatosGuardados) {
      results.forEach(({ dayName, rows }) => {
        if (!rows || rows.length === 0) return;
        if (!modifiedData[dayName]) {
          modifiedData[dayName] = rows.map((row) => ({
            NombreBebe: (row["Nombre Bebe"] || row.NombreBebe || "").trim(),
            NombreMadre: (row["Nombre Madre"] || row.NombreMadre || "").trim(),
            Fase: (
              row.Fase ||
              row.Institucion ||
              row.InstitucionMadre ||
              ""
            ).trim(),
            ProgramaMadre: (row.Programa || row.ProgramaMadre || "").trim(),
            Edad: normalizarEdad(row.Edad || row["Edad (meses)"] || ""),
            Asistencia: "",
            Ubicacion: "",
            Reporte: "No",
            SituacionEspecifica: "",
            Nota: "",
            Extras: "",
            NoCidi: "",
          }));
        } else {
          const existentes = new Set(
            modifiedData[dayName].map((r) => r.NombreBebe.trim().toLowerCase()),
          );
          rows.forEach((row) => {
            const nombre = (row["Nombre Bebe"] || row.NombreBebe || "").trim();
            if (!nombre || existentes.has(nombre.toLowerCase())) return;
            modifiedData[dayName].push({
              NombreBebe: nombre,
              NombreMadre: (
                row["Nombre Madre"] ||
                row.NombreMadre ||
                ""
              ).trim(),
              Fase: (
                row.Fase ||
                row.Institucion ||
                row.InstitucionMadre ||
                ""
              ).trim(),
              ProgramaMadre: (row.Programa || row.ProgramaMadre || "").trim(),
              Edad: normalizarEdad(row.Edad || row["Edad (meses)"] || ""),
              Asistencia: "",
              Ubicacion: "",
              Reporte: "No",
              SituacionEspecifica: "",
              Nota: "",
              Extras: "",
              NoCidi: "",
            });
          });
        }
      });
      renderAllSavedTabs();
    } else {
      processFromServer(results);
    }

    saveToLocalStorage();
    exportBtn.disabled = false;
    updateSyncStatus(
      "ok",
      `BD conectada — ${masterData.length} bebés cargados`,
    );
  } catch (err) {
    updateSyncStatus("error", "Sin conexión a la BD — usando datos guardados");
    console.warn("No se pudo cargar desde la BD:", err.message);
  }
}

// ─── Parseo CSV del endpoint /api/sheet/:dia ──────────────────────────────────
function parseCsv(csvText) {
  const lines = csvText.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0]
    .split(",")
    .map((h) => h.trim().replace(/^"|"$/g, ""));
  return lines
    .slice(1)
    .map((line) => {
      const vals = [];
      let cur = "",
        inQ = false;
      for (const c of line) {
        if (c === '"') {
          inQ = !inQ;
        } else if (c === "," && !inQ) {
          vals.push(cur.trim());
          cur = "";
        } else {
          cur += c;
        }
      }
      vals.push(cur.trim());
      const row = {};
      headers.forEach((h, i) => {
        row[h] = (vals[i] || "").replace(/^"|"$/g, "");
      });
      return row;
    })
    .filter((r) => Object.values(r).some((v) => v !== ""));
}

// ─── Procesar resultados del servidor → modifiedData ─────────────────────────
function processFromServer(results) {
  modifiedData = {};
  daysTabs.innerHTML = "";
  tabsContent.innerHTML = "";

  results.forEach(({ dayName, rows }, index) => {
    if (!rows || rows.length === 0) return;

    modifiedData[dayName] = rows.map((row) => ({
      NombreBebe: (row["Nombre Bebe"] || row.NombreBebe || "").trim(),
      NombreMadre: (row["Nombre Madre"] || row.NombreMadre || "").trim(),
      Fase: (row.Fase || row.Institucion || row.InstitucionMadre || "").trim(),
      ProgramaMadre: (row.Programa || row.ProgramaMadre || "").trim(),
      Edad: normalizarEdad(row.Edad || row["Edad (meses)"] || ""),
      Asistencia: "",
      Ubicacion: "",
      Reporte: "No",
      SituacionEspecifica: "",
      Nota: "",
      Extras: "",
      NoCidi: "",
    }));

    renderTabHeader(dayName, index === 0);
    renderTabContainer(dayName, index === 0);
    renderTable(dayName, modifiedData[dayName]);
  });

  saveToLocalStorage();
  addAddBabyButton();
}

// ─── Utilidades ───────────────────────────────────────────────────────────────
function normalizarEdad(edad) {
  if (!edad) return "";
  const num = parseInt(String(edad).replace(/\D/g, ""), 10);
  if (isNaN(num)) return "";
  if (num >= 6 && num <= 15) return "6-15";
  if (num >= 16 && num <= 30) return "16-30";
  return String(edad).trim();
}

function normalizeText(text = "") {
  return text
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function formatDate(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${d}-${m}-${date.getFullYear()}`;
}

function updateSyncStatus(_state, _text) {
  // Estado de BD no se muestra en la UI — solo log interno
  console.log(`[BD] ${_text}`);
}

// ─── Renderizado de tabs ──────────────────────────────────────────────────────
function renderTabHeader(day, isActive) {
  const tab = document.createElement("div");
  tab.className = `tab ${isActive ? "active" : ""}`;
  tab.textContent = day;
  tab.dataset.day = day;
  tab.onclick = switchTab;
  daysTabs.appendChild(tab);
}

function renderTabContainer(day, isActive) {
  const content = document.createElement("div");
  content.className = `tab-content ${isActive ? "active" : ""}`;
  content.id = `content-${day}`;
  tabsContent.appendChild(content);
}

function switchTab(e) {
  const day = e.target.dataset.day;
  document
    .querySelectorAll(".tab")
    .forEach((t) => t.classList.remove("active"));
  document
    .querySelectorAll(".tab-content")
    .forEach((c) => c.classList.remove("active"));
  e.target.classList.add("active");
  document.getElementById(`content-${day}`).classList.add("active");
  updateCounter();
}

function openCurrentDayTab() {
  const daysMap = {
    1: "Lunes",
    2: "Martes",
    3: "Miércoles",
    4: "Jueves",
    5: "Viernes",
  };
  const todayName = daysMap[new Date().getDay()];
  const todayTab =
    todayName && document.querySelector(`.tab[data-day="${todayName}"]`);
  if (todayTab) {
    todayTab.click();
    updateCounter();
  }
}

function getTodayName() {
  return [
    "Domingo",
    "Lunes",
    "Martes",
    "Miércoles",
    "Jueves",
    "Viernes",
    "Sábado",
  ][new Date().getDay()];
}

// ─── Tabla principal ──────────────────────────────────────────────────────────
function renderTable(day, data, searchTerm = "") {
  const container = document.getElementById(`content-${day}`);
  if (!container) return;
  container.innerHTML = "";

  if (!data || data.length === 0) {
    container.innerHTML = "<p>No hay datos para mostrar.</p>";
    return;
  }

  const term = normalizeText(searchTerm);
  const filtered = data
    .map((row, realIndex) => ({ row, realIndex }))
    .filter(
      ({ row }) =>
        normalizeText(row.NombreBebe).includes(term) ||
        normalizeText(row.NombreMadre).includes(term),
    );

  filtered.sort((a, b) => {
    const p = (r) => (r.NoCidi === "Sí" ? 2 : r.Extras === "Sí" ? 1 : 0);
    return p(b.row) - p(a.row);
  });

  if (filtered.length === 0) {
    container.innerHTML = "<p>No se encontraron resultados.</p>";
    return;
  }

  const table = document.createElement("table");
  table.innerHTML = `
    <thead>
      <tr>
        <th>#</th><th>Nombre Bebé</th><th>Nombre Madre</th><th>Fase</th>
        <th>Programa</th><th>Edad</th><th>Asistencia</th>
      </tr>
    </thead>
  `;

  const tbody = document.createElement("tbody");
  filtered.forEach(({ row, realIndex }, displayIndex) => {
    const tr = document.createElement("tr");
    tr.className =
      row.NoCidi === "Sí"
        ? "nocidi-row"
        : row.Extras === "Sí"
          ? "visitor-row"
          : row.Asistencia === "No"
            ? "absent-row"
            : row.Asistencia === "Sí"
              ? "present-row"
              : "";
    renderRow(tr, row, day, realIndex, displayIndex + 1);
    tbody.appendChild(tr);
    if (tr._accordionTr) tbody.appendChild(tr._accordionTr);
  });

  table.appendChild(tbody);
  container.appendChild(table);

  // Pasar el subset visible al contador solo si hay búsqueda activa
  if (term) {
    updateCounter(filtered.map((f) => f.row));
  } else {
    updateCounter();
  }
}

function renderRow(tr, row, day, index, rowNum) {
  // Columna # — número de posición visible en la tabla
  const tdNum = document.createElement("td");
  tdNum.textContent = rowNum;
  tdNum.style.cssText =
    "width:32px;text-align:center;font-size:11px;color:#888;font-weight:500;";
  tr.appendChild(tdNum);

  tr.innerHTML += `<td>${row.NombreBebe}</td><td>${row.NombreMadre}</td>`;

  // Fase — solo texto estático, viene de BD
  const tdFase = document.createElement("td");
  tdFase.textContent = row.Fase || "—";
  tdFase.style.cssText = "font-size:13px; color:#2e7d32; font-weight:500;";
  tr.appendChild(tdFase);

  // Programa — solo badge estático, viene de BD
  const esTSFinicial = row.Fase === "TSF";
  const tdPrograma = document.createElement("td");
  tdPrograma.style.cssText = "min-width:140px; width:140px;";
  tdPrograma.innerHTML = badgePrograma(esTSFinicial ? "" : row.ProgramaMadre);
  tr.appendChild(tdPrograma);

  // Edad — botón toggle pill
  const tdEdad = document.createElement("td");
  const btnEdad = document.createElement("button");
  btnEdad.type = "button";
  btnEdad.className = "btn-edad";
  const setEdadStyle = (val) => {
    btnEdad.textContent = (val || "6-15") + " ⇄";
    btnEdad.dataset.edad = val || "6-15";
  };
  setEdadStyle(row.Edad || "6-15");
  btnEdad.onclick = () => {
    const next =
      (modifiedData[day][index].Edad || "6-15") === "6-15" ? "16-30" : "6-15";
    updateField(day, index, "Edad", next);
    setEdadStyle(next);
    updateCounter();
  };
  tdEdad.appendChild(btnEdad);
  tr.appendChild(tdEdad);

  // ── Asistencia: Sí / No ──────────────────────────────────────────────────
  const tdAsis = document.createElement("td");
  tdAsis.className = "td-asistencia";
  const btnPair = document.createElement("div");
  btnPair.className = "btn-pair";

  const btnSi = document.createElement("button");
  btnSi.type = "button";
  btnSi.textContent = "Sí";
  btnSi.className =
    "btn-asis btn-si" + (row.Asistencia === "Sí" ? " active" : "");

  const btnNo = document.createElement("button");
  btnNo.type = "button";
  btnNo.textContent = "No";
  btnNo.className =
    "btn-asis btn-no" + (row.Asistencia === "No" ? " active" : "");

  const btnVer = document.createElement("button");
  btnVer.type = "button";
  btnVer.className =
    "btn-asis btn-ver-reporte" + (row.Reporte === "Sí" ? "" : " acc-hidden");
  btnVer.textContent = "Ver";

  const btnEditar = document.createElement("button");
  btnEditar.type = "button";
  btnEditar.className =
    "btn-asis btn-editar-reporte" + (row.Reporte === "Sí" ? "" : " acc-hidden");
  btnEditar.textContent = "Editar";

  btnPair.append(btnSi, btnNo, btnVer, btnEditar);

  // ── Select tipo: Normal / No CIDI / Extras ────────────────────────────────
  // Permite cambiar el tipo de un bebé ya añadido sin tener que eliminarlo
  const selTipo = document.createElement("select");
  selTipo.className = "sel-tipo-bebe";
  selTipo.title = "Cambiar tipo de registro";
  const tipoActual =
    row.NoCidi === "Sí" ? "nocidi" : row.Extras === "Sí" ? "extras" : "normal";
  [
    { value: "normal", label: "Normal" },
    { value: "nocidi", label: "No CIDI" },
    { value: "extras", label: "Extras" },
  ].forEach(({ value, label }) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    if (value === tipoActual) opt.selected = true;
    selTipo.appendChild(opt);
  });
  selTipo.onchange = () => {
    const val = selTipo.value;
    updateField(day, index, "NoCidi", val === "nocidi" ? "Sí" : "");
    updateField(day, index, "Extras", val === "extras" ? "Sí" : "");
    tr.className = getRowClass();
    updateCounter();
    saveToLocalStorage();
  };
  btnPair.appendChild(selTipo);
  tdAsis.appendChild(btnPair);
  tr.appendChild(tdAsis);

  // ── Acordeón de reporte ───────────────────────────────────────────────────
  const accordionTr = document.createElement("tr");
  accordionTr.className = "accordion-tr acc-hidden";
  const accordionTd = document.createElement("td");
  // Calcular colSpan dinámicamente desde la fila de encabezado de la tabla
  const parentTable = tr.closest("table");
  accordionTd.colSpan = parentTable ? parentTable.rows[0].cells.length : 6;
  accordionTd.className = "accordion-td";

  const yaReportado = row.Reporte === "Sí";
  const formDiv = document.createElement("div");
  formDiv.className = "reporte-form" + (yaReportado ? " acc-hidden" : "");

  const grpUbic = document.createElement("div");
  grpUbic.className = "rpf-field";
  grpUbic.innerHTML = "<label>Ubicación</label>";
  grpUbic.appendChild(
    createSelect(
      ["Juanfe", "Casa", "Otro"],
      row.Ubicacion,
      (val) => updateField(day, index, "Ubicacion", val),
      "Seleccionar",
    ),
  );

  const grpRep = document.createElement("div");
  grpRep.className = "rpf-field";
  grpRep.innerHTML = "<label>Reporte</label>";
  const selRep = createSelect(["Sí", "No"], row.Reporte, (val) => {
    updateField(day, index, "Reporte", val);
    grpSitu.style.display = val === "Sí" ? "" : "none";
    grpNota.style.display = val === "Sí" ? "" : "none";
    if (val === "No") {
      // Auto-limpiar reporte — limpiar campos y ocultar botones sin necesidad de guardar
      updateField(day, index, "SituacionEspecifica", "");
      updateField(day, index, "Nota", "");
      updateField(day, index, "Ubicacion", "");
      btnVer.classList.add("acc-hidden");
      btnEditar.classList.add("acc-hidden");
      summaryDiv.classList.add("acc-hidden");
      accordionTr.classList.add("acc-hidden");
      saveToLocalStorage();
    }
    updateCounter();
  });
  grpRep.appendChild(selRep);

  const grpSitu = document.createElement("div");
  grpSitu.className = "rpf-field";
  grpSitu.style.display = row.Reporte === "Sí" ? "" : "none";
  grpSitu.innerHTML = "<label>Situación</label>";
  grpSitu.appendChild(
    createSelect(
      diseaseOptions,
      row.SituacionEspecifica,
      (val) => updateField(day, index, "SituacionEspecifica", val),
      "Seleccionar",
    ),
  );

  const grpNota = document.createElement("div");
  grpNota.className = "rpf-field";
  grpNota.style.display = row.Reporte === "Sí" ? "" : "none";
  grpNota.innerHTML = "<label>Nota</label>";
  const inputNota = document.createElement("input");
  inputNota.type = "text";
  inputNota.value = row.Nota || "";
  inputNota.placeholder = "Observación...";
  inputNota.oninput = (e) => updateField(day, index, "Nota", e.target.value);
  grpNota.appendChild(inputNota);

  const btnGuardar = document.createElement("button");
  btnGuardar.type = "button";
  btnGuardar.className = "btn-guardar-reporte";
  btnGuardar.textContent = "Guardar reporte";
  btnGuardar.onclick = () => {
    saveToLocalStorage();
    const r = modifiedData[day][index];
    formDiv.classList.add("acc-hidden");
    summaryDiv.classList.remove("acc-hidden");
    btnVer.classList.remove("acc-hidden");
    btnEditar.classList.remove("acc-hidden");
    // Actualizar texto del summary con ubicación + situación + nota
    const ubicTxt = r.Ubicacion ? `📍 ${r.Ubicacion}` : "";
    const situTxt = r.SituacionEspecifica
      ? `⚠️ ${r.SituacionEspecifica}`
      : "Sin situación";
    const notaTxt = r.Nota ? `💬 ${r.Nota}` : "";
    summaryText.innerHTML = [ubicTxt, situTxt, notaTxt]
      .filter(Boolean)
      .join(" &nbsp;·&nbsp; ");
    accordionTr.classList.add("acc-hidden");
    updateCounter();
  };
  formDiv.append(grpUbic, grpRep, grpSitu, grpNota, btnGuardar);

  const summaryDiv = document.createElement("div");
  summaryDiv.className = "reporte-summary" + (yaReportado ? "" : " acc-hidden");
  const summaryText = document.createElement("span");
  summaryText.className = "reporte-summary-text";
  if (yaReportado) {
    const ubicTxt = row.Ubicacion ? `📍 ${row.Ubicacion}` : "";
    const situTxt = row.SituacionEspecifica
      ? `⚠️ ${row.SituacionEspecifica}`
      : "Sin situación";
    const notaTxt = row.Nota ? `💬 ${row.Nota}` : "";
    summaryText.innerHTML = [ubicTxt, situTxt, notaTxt]
      .filter(Boolean)
      .join(" &nbsp;·&nbsp; ");
  }
  summaryDiv.appendChild(summaryText);

  accordionTd.append(formDiv, summaryDiv);
  accordionTr.appendChild(accordionTd);

  // ── Lógica Sí / No con toggle ─────────────────────────────────────────────
  const getRowClass = () => {
    const r = modifiedData[day][index];
    // NoCidi y Extras mantienen su color base SIEMPRE
    // más la clase de asistencia como segundo modificador
    if (r.NoCidi === "Sí") return "nocidi-row";
    if (r.Extras === "Sí") return "visitor-row";
    if (r.Asistencia === "Sí") return "present-row";
    if (r.Asistencia === "No") return "absent-row";
    return "";
  };

  btnSi.onclick = () => {
    if (row.Asistencia === "Sí") {
      updateField(day, index, "Asistencia", "");
      btnSi.classList.remove("active");
      tr.className = getRowClass();
    } else {
      updateField(day, index, "Asistencia", "Sí");
      // Si no tiene edad asignada, asignar 6-15 por defecto
      if (!modifiedData[day][index].Edad) {
        updateField(day, index, "Edad", "6-15");
        setEdadStyle("6-15");
      }
      btnSi.classList.add("active");
      btnNo.classList.remove("active");
      tr.className = getRowClass() || "present-row";
      accordionTr.classList.add("acc-hidden");
    }
    updateCounter();
  };

  btnNo.onclick = () => {
    if (row.Asistencia === "No") {
      updateField(day, index, "Asistencia", "");
      btnNo.classList.remove("active");
      tr.className = getRowClass();
      accordionTr.classList.add("acc-hidden");
    } else {
      updateField(day, index, "Asistencia", "No");
      btnNo.classList.add("active");
      btnSi.classList.remove("active");
      tr.className = getRowClass() || "absent-row";
      accordionTr.classList.remove("acc-hidden");
    }
    updateCounter();
  };

  btnVer.onclick = () => {
    accordionTr.classList.toggle("acc-hidden");
    if (!accordionTr.classList.contains("acc-hidden")) {
      summaryDiv.classList.remove("acc-hidden");
      formDiv.classList.add("acc-hidden");
    }
  };

  btnEditar.onclick = () => {
    accordionTr.classList.remove("acc-hidden");
    formDiv.classList.remove("acc-hidden");
    summaryDiv.classList.add("acc-hidden");
  };

  tr._accordionTr = accordionTr;
}

// ─── Helpers UI ───────────────────────────────────────────────────────────────
function createSelect(options, currentVal, onChange, placeholder = null) {
  const sel = document.createElement("select");
  if (placeholder) {
    const def = document.createElement("option");
    def.value = "";
    def.textContent = placeholder;
    def.disabled = true;
    if (!currentVal) def.selected = true;
    sel.appendChild(def);
  }
  options.forEach((opt) => {
    const val = typeof opt === "string" ? opt : opt.value;
    const text = typeof opt === "string" ? opt : opt.text;
    const o = document.createElement("option");
    o.value = val;
    o.textContent = text;
    if (val === currentVal) o.selected = true;
    sel.appendChild(o);
  });
  sel.onchange = (e) => onChange(e.target.value);
  return sel;
}

function wrapInTd(el) {
  const td = document.createElement("td");
  td.appendChild(el);
  return td;
}

// ─── Persistencia local ───────────────────────────────────────────────────────
function updateField(day, index, field, value) {
  if (modifiedData[day]?.[index] !== undefined) {
    modifiedData[day][index][field] = value;
    saveToLocalStorage();
  }
}

function saveToLocalStorage() {
  // Guardamos junto a los datos la fecha ISO del día actual.
  // Esto permite detectar datos de días anteriores al cargar.
  const hoy = new Date().toISOString().split("T")[0];
  localStorage.setItem(
    "datos_asistencia",
    JSON.stringify({ fecha: hoy, data: modifiedData }),
  );
}

function loadFromLocalStorage() {
  const raw = localStorage.getItem("datos_asistencia");
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    const hoy = new Date().toISOString().split("T")[0];

    // Formato viejo (sin fecha) o datos de otro día → descartar
    if (!parsed.fecha || !parsed.data) {
      localStorage.removeItem("datos_asistencia");
      return;
    }
    if (parsed.fecha !== hoy) {
      console.log(`⚠️ Datos del ${parsed.fecha} descartados — hoy es ${hoy}`);
      localStorage.removeItem("datos_asistencia");
      return;
    }

    modifiedData = parsed.data;
    renderAllSavedTabs();
    exportBtn.disabled = false;
  } catch (e) {
    console.warn("Error parseando localStorage:", e.message);
    localStorage.removeItem("datos_asistencia");
  }
  // masterData NO se restaura de localStorage — viene de /api/bebes
}

function renderAllSavedTabs() {
  daysTabs.innerHTML = "";
  tabsContent.innerHTML = "";
  days.forEach((day, index) => {
    if (!modifiedData[day]) return;
    renderTabHeader(day, index === 0);
    renderTabContainer(day, index === 0);
    renderTable(day, modifiedData[day]);
  });
}

// ─── Contadores ───────────────────────────────────────────────────────────────
// Si se pasa `subset`, los contadores reflejan solo esos registros (búsqueda activa).
// Si no se pasa, usa todos los datos del día.
function updateCounter(subset) {
  const activeTab = document.querySelector(".tab.active");
  const counterBar = document.getElementById("counter-bar");
  if (!activeTab || !counterBar) return;

  const day = activeTab.dataset.day;
  const data = subset !== undefined ? subset : modifiedData[day] || [];
  const totalDia = (modifiedData[day] || []).length; // total real del día siempre

  const present = data.filter((r) => r.Asistencia === "Sí").length;
  const absent = data.length - present;
  const reported = data.filter((r) => r.Reporte === "Sí").length;
  const extras = data.filter((r) => r.Extras === "Sí").length;
  const noCidi = data.filter((r) => r.NoCidi === "Sí").length;
  const edad1 = data.filter(
    (r) => r.Asistencia === "Sí" && r.Edad === "6-15",
  ).length;
  const edad2 = data.filter(
    (r) => r.Asistencia === "Sí" && r.Edad === "16-30",
  ).length;
  const rep1 = data.filter(
    (r) => r.Reporte === "Sí" && r.Edad === "6-15",
  ).length;
  const rep2 = data.filter(
    (r) => r.Reporte === "Sí" && r.Edad === "16-30",
  ).length;

  // "Total" siempre muestra el total del día, no el filtrado
  document.getElementById("count-total").textContent = totalDia;
  document.getElementById("count-present").textContent = present;
  document.getElementById("count-absent").textContent = absent;
  document.getElementById("count-reported").textContent = reported;
  document.getElementById("count-extras").textContent = extras;
  document.getElementById("count-nocidi").textContent = noCidi;
  document.getElementById("count-edad1").textContent = edad1;
  document.getElementById("count-edad2").textContent = edad2;
  document.getElementById("count-rep1").textContent = rep1;
  document.getElementById("count-rep2").textContent = rep2;
  counterBar.style.display = totalDia > 0 ? "flex" : "none";
}

function filterData() {
  const term = normalizeText(searchInput.value);
  const activeTab = document.querySelector(".tab.active");
  if (!activeTab) return;
  const day = activeTab.dataset.day;
  if (!modifiedData[day]) return;

  if (term) {
    // Con búsqueda activa: contadores reflejan solo los resultados visibles
    const subset = modifiedData[day].filter(
      (r) =>
        normalizeText(r.NombreBebe).includes(term) ||
        normalizeText(r.NombreMadre).includes(term),
    );
    renderTable(day, modifiedData[day], term);
    updateCounter(subset);
  } else {
    // Sin búsqueda: contadores muestran el día completo
    renderTable(day, modifiedData[day]);
    updateCounter();
  }
}

// ─── Exportación a Excel + guardar en Supabase ────────────────────────────────
function exportToExcel() {
  const activeTab = document.querySelector(".tab.active");
  if (!activeTab) {
    alert("Por favor, selecciona un día primero.");
    return;
  }

  const dayToExport = activeTab.dataset.day;
  const apiDia = DIAS_API[dayToExport];

  // Validar que sea un día hábil — en fin de semana apiDia sería undefined
  if (!apiDia) {
    alert(
      `"${dayToExport}" no es un día hábil del CIDI.\n` +
        `Solo se puede exportar de Lunes a Viernes.`,
    );
    return;
  }

  const dataToExport = modifiedData[dayToExport] || [];

  if (dataToExport.length === 0) {
    alert(`No hay datos para exportar en el día ${dayToExport}`);
    return;
  }

  const fechaISO = new Date().toISOString().split("T")[0];

  // 1. Preparar y descargar Excel
  const dataOrdered = dataToExport.map((row) => {
    const newRow = { Fecha: fechaISO, Dia: dayToExport };
    columnOrder.forEach((key) => {
      newRow[columnHeaders[key]] =
        key === "Asistencia" ? row[key] || "No" : (row[key] ?? "");
    });
    return newRow;
  });

  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    newWb,
    XLSX.utils.json_to_sheet(dataOrdered),
    apiDia,
  );
  const nombreArchivo = `asistencia-${dayToExport}-${formatDate(new Date())}.xlsx`;
  XLSX.writeFile(newWb, nombreArchivo);

  // 2. Guardar en Supabase
  // El respaldo en disco fue eliminado — Render tiene filesystem efímero
  // (se borra en cada deploy). Supabase es el único respaldo necesario.
  const registros = dataToExport.map((row) => ({
    NombreBebe: row.NombreBebe || "",
    NombreMadre: row.NombreMadre || "",
    Fase: row.Fase || "",
    ProgramaMadre: row.ProgramaMadre || "",
    Edad: row.Edad || "",
    Asistencia: row.Asistencia || "No",
    Ubicacion: row.Ubicacion || "",
    Reporte: row.Reporte || "No",
    SituacionEspecifica: row.SituacionEspecifica || "",
    Nota: row.Nota || "",
    Extras: row.Extras || "",
    NoCidi: row.NoCidi || "",
  }));

  authFetch("/api/asistencia/guardar", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fecha: fechaISO, dia: apiDia, registros }),
  })
    .then((r) => r.json())
    .then((r) => {
      if (r.ok)
        console.log(
          `✅ Supabase: ${r.guardados} registros — ${dayToExport} ${fechaISO}`,
        );
      else console.warn("⚠️ Supabase error:", r.error);
    })
    .catch((e) =>
      console.warn("⚠️ No se pudo guardar en Supabase:", e.message),
    );

  // 3. Limpiar tabla para el día siguiente
  // Se eliminan Extras y NoCidi (son temporales del día).
  // Edad NO se borra — es dato del perfil del bebé, no de la sesión.
  modifiedData[dayToExport] = modifiedData[dayToExport].filter(
    (row) => row.Extras !== "Sí" && row.NoCidi !== "Sí",
  );
  modifiedData[dayToExport].forEach((row) => {
    row.Asistencia = "";
    row.Ubicacion = "";
    row.Reporte = "No";
    row.SituacionEspecifica = "";
    row.Nota = "";
  });

  saveToLocalStorage();
  searchInput.value = "";
  renderTable(dayToExport, modifiedData[dayToExport]);
}

// ─── Modal: Añadir Bebé ───────────────────────────────────────────────────────
function addAddBabyButton() {
  const addBabyBtn = document.getElementById("add-baby-btn");
  if (!addBabyBtn || addBabyBtn.dataset.listenerRegistered) return;
  addBabyBtn.dataset.listenerRegistered = "true";

  const modal = document.getElementById("addBabyModal");
  const closeBtn = document.querySelector(".close");
  const babyForm = document.getElementById("babyForm");
  const searchBabyInput = document.getElementById("searchBaby");
  const searchResults = document.getElementById("searchResults");
  const babyNameInput = document.getElementById("babyName");
  const motherNameInput = document.getElementById("motherName");
  const faseSelect = document.getElementById("motherFase");
  const progSelect = document.getElementById("motherPrograma");
  const edadSelect = document.getElementById("babyEdad");
  const visitanteCheck = document.getElementById("esVisitante");
  const noCidiCheck = document.getElementById("esNoCidi");

  visitanteCheck.addEventListener("change", () => {
    noCidiCheck.disabled = visitanteCheck.checked;
    if (visitanteCheck.checked) noCidiCheck.checked = false;
  });
  noCidiCheck.addEventListener("change", () => {
    visitanteCheck.disabled = noCidiCheck.checked;
    if (noCidiCheck.checked) visitanteCheck.checked = false;
  });

  // TSF no tiene programa — deshabilitar select de programa en modal
  faseSelect.addEventListener("change", () => {
    const esTSF = faseSelect.value === "TSF";
    progSelect.disabled = esTSF;
    progSelect.style.opacity = esTSF ? "0.45" : "1";
    if (esTSF) progSelect.value = "";
  });

  function resetModal() {
    babyForm.reset();
    searchBabyInput.value = "";
    searchResults.innerHTML = "";
    searchResults.style.display = "none";
    visitanteCheck.disabled = false;
    noCidiCheck.disabled = false;
    progSelect.disabled = false;
    progSelect.style.opacity = "1";
  }

  function fillForm(baby) {
    searchBabyInput.value = baby.NombreBebe;
    babyNameInput.value = baby.NombreBebe;
    motherNameInput.value = baby.NombreMadre;
    faseSelect.value = baby.Fase || "";
    edadSelect.value = baby.Edad || "";
    const esTSF = baby.Fase === "TSF";
    progSelect.disabled = esTSF;
    progSelect.style.opacity = esTSF ? "0.45" : "1";
    progSelect.value = esTSF ? "" : baby.ProgramaMadre || "";
  }

  addBabyBtn.addEventListener("click", () => {
    resetModal();
    modal.style.display = "block";
    searchBabyInput.focus();
  });
  closeBtn.addEventListener("click", () => {
    modal.style.display = "none";
  });
  window.addEventListener("click", (e) => {
    if (e.target === modal) modal.style.display = "none";
  });

  searchBabyInput.addEventListener("input", () => {
    const term = normalizeText(searchBabyInput.value.trim());
    searchResults.innerHTML = "";
    if (term.length < 2) {
      searchResults.style.display = "none";
      return;
    }

    const matches = masterData
      .filter(
        (b) =>
          normalizeText(b.NombreBebe).includes(term) ||
          normalizeText(b.NombreMadre).includes(term),
      )
      .slice(0, 8);

    if (matches.length === 0) {
      searchResults.style.display = "none";
      return;
    }

    matches.forEach((baby) => {
      const item = document.createElement("div");
      item.className = "search-result-item";
      item.innerHTML = `<strong>${baby.NombreBebe}</strong><span>${baby.NombreMadre} · ${baby.Fase} ${baby.ProgramaMadre}</span>`;
      item.addEventListener("click", () => {
        fillForm(baby);
        searchResults.style.display = "none";
      });
      searchResults.appendChild(item);
    });
    searchResults.style.display = "block";
  });

  babyForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const activeTab = document.querySelector(".tab.active");
    const currentDay = activeTab ? activeTab.dataset.day : getTodayName();
    const nombreNuevo = babyNameInput.value.trim();

    if (
      modifiedData[currentDay]?.some(
        (b) => normalizeText(b.NombreBebe) === normalizeText(nombreNuevo),
      )
    ) {
      showSmartAlert(
        `El bebé "${nombreNuevo}" ya está en la lista de hoy (${currentDay}).`,
      );
      return;
    }

    const newBaby = {
      NombreBebe: nombreNuevo,
      NombreMadre: motherNameInput.value.trim(),
      Fase: faseSelect.value,
      ProgramaMadre: progSelect.value,
      Edad: edadSelect.value,
      Asistencia: "Sí",
      Ubicacion: "",
      Reporte: "No",
      SituacionEspecifica: "",
      Nota: "",
      Extras: visitanteCheck.checked ? "Sí" : "",
      NoCidi: noCidiCheck.checked ? "Sí" : "",
    };

    if (!newBaby.NombreBebe || !newBaby.NombreMadre) {
      alert("Por favor complete al menos el nombre del bebé y la madre.");
      return;
    }

    // Mostrar modal de confirmación con resumen antes de agregar
    const tipo =
      newBaby.NoCidi === "Sí"
        ? "No CIDI"
        : newBaby.Extras === "Sí"
          ? "Extras"
          : "Normal";
    const resumen = [
      newBaby.Fase ? `Fase: ${newBaby.Fase}` : null,
      newBaby.ProgramaMadre ? `Programa: ${newBaby.ProgramaMadre}` : null,
      newBaby.Edad ? `Edad: ${newBaby.Edad} meses` : null,
      `Tipo: ${tipo}`,
    ]
      .filter(Boolean)
      .join("  ·  ");

    modal.style.display = "none";
    showAddBabyConfirm({
      nombre: newBaby.NombreBebe,
      madre: newBaby.NombreMadre,
      resumen,
      onConfirm: () => {
        if (!modifiedData[currentDay]) modifiedData[currentDay] = [];
        if (newBaby.Extras === "Sí" || newBaby.NoCidi === "Sí") {
          modifiedData[currentDay].unshift(newBaby);
        } else {
          modifiedData[currentDay].push(newBaby);
        }
        saveToLocalStorage();
        renderTable(currentDay, modifiedData[currentDay]);
        updateCounter();
        resetModal();
        exportBtn.disabled = false;
        if (searchInput.value) filterData();
      },
      onCancel: () => {
        modal.style.display = "block"; // volver al modal de añadir
      },
    });
  });
}

function showSmartAlert(message) {
  const modal = document.getElementById("smartAlertModal");
  const msgP = document.getElementById("smartAlertMessage");
  const closeBtn = document.getElementById("closeSmartAlert");
  msgP.textContent = message;
  modal.style.display = "block";

  function cerrar() {
    modal.style.display = "none";
    closeBtn.removeEventListener("click", cerrar);
    window.removeEventListener("click", outsideClick);
  }
  function outsideClick(e) {
    if (e.target === modal) cerrar();
  }

  closeBtn.addEventListener("click", cerrar);
  window.addEventListener("click", outsideClick);
}

// ── Modal resumen confirmar añadir bebé ───────────────────────────────────────
function showAddBabyConfirm({ nombre, madre, resumen, onConfirm, onCancel }) {
  const modal = document.getElementById("addBabyConfirmModal");
  document.getElementById("abcNombre").textContent = nombre;
  document.getElementById("abcMadre").textContent = madre;
  document.getElementById("abcResumen").textContent = resumen;
  modal.style.display = "flex";

  const btnSi = document.getElementById("abcBtnSi");
  const btnNo = document.getElementById("abcBtnNo");

  const close = (cb) => {
    modal.style.display = "none";
    btnSi.onclick = null;
    btnNo.onclick = null;
    if (cb) cb();
  };

  btnSi.onclick = () => close(onConfirm);
  btnNo.onclick = () => close(onCancel);
  modal.onclick = (e) => {
    if (e.target === modal) close(onCancel);
  };
}
