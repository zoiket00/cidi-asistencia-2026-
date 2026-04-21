"use strict";
/* =============================================================================
   dashboard.js — Fundación Juanfe · Dashboard de Asistencia  v4
   =============================================================================

   ARQUITECTURA
   ─────────────────────────────────────────────────────────────────────────────
   1. UnifiedModel  — Modelo de datos: ingesta, deduplicación, catálogo de bebés
   2. Estado global — allData, headers, charts, filtros activos
   3. Almacenamiento — IndexedDB para archivos grandes (sin límite de 5MB)
   4. UI / Upload   — Carga de archivos, chips, clear session
   5. Vistas KPI    — Conteos fijos sobre el total (no cambian con filtros)
   6. Dashboard     — Gráficas por filtro activo
   7. Tabla         — Paginación, ordenamiento, exportación CSV
   8. ETL Avanzado  — Tendencia, Ranking, Comparación, Alertas de riesgo

   REGLAS DE NEGOCIO
   ─────────────────────────────────────────────────────────────────────────────
   • Identidad de un bebé = NombreBebe + NombreMadre  (distingue homónimos)
   • Un registro único   = bebé + fecha + día          (no duplicar el mismo día)
   • Extras y NoCidi cuentan en Total, Presentes y Ausentes (son personas reales)
   • Los KPIs nunca cambian al cambiar el filtro activo
   • La tabla muestra TODOS los registros paginados (50/página)
   • El CSV exporta el conjunto completo, no solo la página visible

   CORRECCIONES v4 vs versión anterior
   ─────────────────────────────────────────────────────────────────────────────
   ✓ _makeKeys incluye nombre+madre en la clave → homónimos correctamente separados
   ✓ clearCharts limpia el objeto charts completamente (sin keys fantasmas)
   ✓ updateKPIs usa solo _diasPresente (sin doble condición que sobrecontaba)
   ✓ buildDashboard eliminado (era redundante con buildDashboardForFilter)
   ✓ buildETLSection sin shadowing de la variable global allData
   ✓ buildETLRanking agrupa por nombre+madre (no solo nombre)
   ✓ buildChartEdad maneja tanto valores numéricos como rangos "6-15" / "16-30"
   ✓ buildViewHeader calcula pct correctamente para todos los filtros
   ✓ Comentarios coherentes con el comportamiento real del código
   ============================================================================= */

// =============================================================================
//  1. MODELO DE DATOS UNIFICADO
// =============================================================================

const UnifiedModel = (() => {
  // ---------------------------------------------------------------------------
  //  Mapeo de variantes de columna → nombre canónico
  //  Cubre todos los formatos de Excel usados en el sistema Juanfe
  // ---------------------------------------------------------------------------
  const COL_CANON = {
    // Bebé
    nombrebebe: "NombreBebe",
    "nombre bebe": "NombreBebe",
    "nombre bebé": "NombreBebe",
    bebe: "NombreBebe",
    bebé: "NombreBebe",
    // Madre
    nombremadre: "NombreMadre",
    "nombre madre": "NombreMadre",
    madre: "NombreMadre",
    // Institución / Fase (ambos nombres según versión del Excel)
    institucionmadre: "InstitucionMadre",
    institucion: "InstitucionMadre",
    institución: "InstitucionMadre",
    "institucion madre": "InstitucionMadre",
    fase: "InstitucionMadre", // nombre nuevo desde app.js v2
    // Programa
    programamadre: "ProgramaMadre",
    programa: "ProgramaMadre",
    "programa madre": "ProgramaMadre",
    // Edad
    edad: "Edad",
    "edad (meses)": "Edad",
    // Asistencia
    asistencia: "Asistencia",
    presente: "Asistencia",
    // Tiempo
    dia: "Dia",
    día: "Dia",
    fecha: "Fecha",
    // Seguimiento
    reporte: "Reporte",
    situacionespecifica: "SituacionEspecifica",
    "situacion especifica": "SituacionEspecifica",
    "situación específica": "SituacionEspecifica",
    situacion: "SituacionEspecifica",
    nota: "Nota",
    ubicacion: "Ubicacion",
    ubicación: "Ubicacion",
    // Categorías especiales — múltiples variantes según versión del Excel
    visitante: "Visitante",
    visitantes: "Visitante",
    extras: "Visitante",
    extra: "Visitante",
    nocidi: "NoCidi",
    "no cidi": "NoCidi",
    no_cidi: "NoCidi",
  };

  /** Normaliza texto: minúsculas + sin tildes + sin espacios extras */
  function _norm(str = "") {
    return String(str)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
  }

  /** Convierte una columna raw → nombre canónico, o null si no se reconoce */
  function _canonKey(rawKey) {
    return COL_CANON[_norm(rawKey)] || null;
  }

  /**
   * Normaliza una fila cruda del Excel → objeto con claves canónicas.
   * También resuelve el campo legado "CursoMadre" (ej: "UTE Hotelería").
   */
  function _normalizeRow(rawRow) {
    const out = {};
    for (const k of Object.keys(rawRow)) {
      const canon = _canonKey(k);
      if (canon) out[canon] = String(rawRow[k] ?? "").trim();
    }

    // Resolver CursoMadre legado: "UTE Hotelería" → InstitucionMadre + ProgramaMadre
    if (rawRow.CursoMadre && !out.InstitucionMadre) {
      const s = String(rawRow.CursoMadre).trim();
      const INSTS = ["ULA 2", "ULA 1", "UTE", "TSF"];
      let inst = "";
      for (const i of INSTS) {
        if (s.startsWith(i)) {
          inst = i;
          break;
        }
      }
      if (!out.InstitucionMadre) out.InstitucionMadre = inst;
      if (!out.ProgramaMadre) out.ProgramaMadre = s.replace(inst, "").trim();
    }

    // Normalizar Edad: acepta "6-15", "16-30", "6 a 15 meses", "16 a 30 meses", número suelto
    if (out.Edad) {
      const rawE = String(out.Edad)
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
      if (rawE === "6-15") out.Edad = "6-15";
      else if (rawE === "16-30") out.Edad = "16-30";
      else if (/6.{1,5}15/.test(rawE)) out.Edad = "6-15";
      else if (/16.{1,5}30/.test(rawE)) out.Edad = "16-30";
      else {
        const n = parseInt(rawE.replace(/\D/g, ""), 10);
        if (!isNaN(n)) {
          if (n >= 6 && n <= 15) out.Edad = "6-15";
          else if (n >= 16 && n <= 30) out.Edad = "16-30";
        }
      }
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  //  Clave de identidad
  //  REGLA: Un registro único = bebé (nombre+madre) + fecha + día
  //  Incluir madre evita que homónimos (Liam David Pertuz ≠ Liam David Hernandez)
  //  sean tratados como la misma persona.
  // ---------------------------------------------------------------------------
  function _makeRegistroKey(nombre, madre, fecha, dia) {
    const n = _norm(nombre) + "|" + _norm(madre || "");
    const f = _norm(fecha || "");
    const d = _norm(dia || "");
    if (!_norm(nombre)) return null; // fila sin nombre → ignorar
    if (f && d) return `${n}|${f}|${d}`;
    if (f) return `${n}|${f}`;
    if (d) return `${n}||${d}`;
    return null; // sin fecha ni día → solo va al catálogo
  }

  function _makeBebeKey(nombre, madre) {
    return _norm(nombre) + "|" + _norm(madre || "");
  }

  // ---------------------------------------------------------------------------
  //  Almacenamiento interno
  // ---------------------------------------------------------------------------
  /** Registros de asistencia: registroKey → entry */
  const _store = new Map();

  /** Catálogo de bebés: bebeKey → perfil */
  const _bebes = new Map();

  /** Historial de fuentes cargadas */
  const _fuentes = [];

  // ---------------------------------------------------------------------------
  //  API pública
  // ---------------------------------------------------------------------------

  /**
   * ingest(rows, nombreFuente)
   * Ingesta un batch de filas crudas desde un Excel.
   * Estrategia de duplicados: el archivo cargado más recientemente gana
   * (mismo bebé en el mismo día → se queda con el registro más nuevo).
   */
  function ingest(rows, nombreFuente = "archivo") {
    const ts = Date.now();
    let nuevos = 0,
      reemplazados = 0,
      soloBebes = 0;

    for (const rawRow of rows) {
      const row = _normalizeRow(rawRow);
      const nombre = row.NombreBebe;
      if (!nombre) continue;

      const madre = row.NombreMadre || "";

      // 1. Actualizar catálogo de bebés (siempre, aunque no haya fecha/día)
      const bKey = _makeBebeKey(nombre, madre);
      if (_bebes.has(bKey)) {
        // Fusión: rellenar solo campos vacíos (no sobreescribir datos existentes)
        const existing = _bebes.get(bKey);
        for (const f of [
          "NombreMadre",
          "InstitucionMadre",
          "ProgramaMadre",
          "Edad",
        ]) {
          if (!existing[f] && row[f]) existing[f] = row[f];
        }
      } else {
        _bebes.set(bKey, {
          NombreBebe: nombre,
          NombreMadre: madre,
          InstitucionMadre: row.InstitucionMadre || "",
          ProgramaMadre: row.ProgramaMadre || "",
          Edad: row.Edad || "",
        });
      }

      // 2. Registrar asistencia (solo si hay fecha o día)
      const rKey = _makeRegistroKey(nombre, madre, row.Fecha, row.Dia);
      if (!rKey) {
        soloBebes++;
        continue;
      }

      const entry = {
        NombreBebe: nombre,
        NombreMadre: madre,
        InstitucionMadre: row.InstitucionMadre || "",
        ProgramaMadre: row.ProgramaMadre || "",
        Edad: row.Edad || "",
        Fecha: row.Fecha || "",
        Dia: row.Dia || "",
        Asistencia: row.Asistencia || "No",
        Ubicacion: row.Ubicacion || "",
        Reporte: row.Reporte || "No",
        SituacionEspecifica: row.SituacionEspecifica || "",
        Nota: row.Nota || "",
        Visitante: row.Visitante || "",
        NoCidi: row.NoCidi || "",
        _fuente: nombreFuente,
        _ts: ts,
      };

      if (_store.has(rKey)) {
        // Gana el más reciente
        if (ts >= _store.get(rKey)._ts) {
          _store.set(rKey, entry);
          reemplazados++;
        }
      } else {
        _store.set(rKey, entry);
        nuevos++;
      }
    }

    _fuentes.push({
      nombre: nombreFuente,
      ts,
      filas: rows.length,
      nuevos,
      reemplazados,
    });
    console.log(
      `[UnifiedModel] "${nombreFuente}" → +${nuevos} nuevos, ${reemplazados} reemplazados, ${soloBebes} sin fecha`,
    );
    return { nuevos, reemplazados, soloBebes };
  }

  /**
   * getAll()
   * Devuelve todos los registros de asistencia únicos (un objeto por clave).
   * El store ya garantiza unicidad; aquí simplemente convertimos a array.
   */
  function getAll() {
    return Array.from(_store.values());
  }

  /** Lista de nombres de bebés para el autocomplete, ordenada alfabéticamente */
  function getBabyNames() {
    return Array.from(_bebes.values())
      .map((b) => b.NombreBebe)
      .filter(Boolean)
      .sort();
  }

  /** Estadísticas de la sesión actual */
  function getMeta() {
    const registros = _store.size;
    const bebesUnicos = _bebes.size;
    return { fuentes: _fuentes.length, registros, bebesUnicos };
  }

  /** Limpia todo el modelo */
  function clear() {
    _store.clear();
    _bebes.clear();
    _fuentes.length = 0;
  }

  return { ingest, getAll, getBabyNames, getMeta, clear };
})();

// =============================================================================
//  2. ESTADO GLOBAL
// =============================================================================

/** allData: array de registros canónicos. SIEMPRE se obtiene de UnifiedModel.getAll() */
let allData = [];
/** headers: lista de columnas canónicas activas */
let headers = [];
/** charts: mapa de instancias Chart.js activas */
let charts = {};
/** Filtro activo en el menú de vistas */
let activeFilter = "todos";
/** Nombre del bebé seleccionado para historial individual (null = vista general) */
let selectedBaby = null;
/** Metadatos de archivos cargados en sesión */
let loadedFiles = [];

// Columnas canónicas que el modelo siempre produce
const CANONICAL_HEADERS = [
  "NombreBebe",
  "NombreMadre",
  "InstitucionMadre",
  "ProgramaMadre",
  "Edad",
  "Fecha",
  "Dia",
  "Asistencia",
  "Ubicacion",
  "Reporte",
  "SituacionEspecifica",
  "Nota",
  "Visitante",
  "NoCidi",
];

// Mapa tipo → columna canónica (para las funciones val() y colKey())
const COL = {
  bebe: ["NombreBebe"],
  madre: ["NombreMadre"],
  institucion: ["InstitucionMadre"],
  programa: ["ProgramaMadre"],
  edad: ["Edad"],
  fecha: ["Fecha"],
  dia: ["Dia"],
  asistencia: ["Asistencia"],
  ubicacion: ["Ubicacion"],
  reporte: ["Reporte"],
  situacion: ["SituacionEspecifica"],
  nota: ["Nota"],
  visitante: ["Visitante"],
  nocidi: ["NoCidi"],
};

// Paleta de colores del dashboard
const VERDE = "#85da1a";
const ROJO = "#ef4444";
const NARANJA = "#f97316";
const AZUL = "#3b82f6";
const MORADO = "#8b5cf6";
const TEAL = "#14b8a6";
const PALETTE = [
  VERDE,
  ROJO,
  NARANJA,
  AZUL,
  MORADO,
  TEAL,
  "#ec4899",
  "#f59e0b",
  "#06b6d4",
  "#84cc16",
];

// =============================================================================
//  3. ALMACENAMIENTO — IndexedDB
//  Permite guardar 20+ archivos Excel sin el límite de ~5MB de localStorage.
//  localStorage solo guarda la lista de nombres (metadatos).
// =============================================================================

const LS_KEY = "juanfe_dash_v4_meta";
const IDB_NAME = "juanfe_dash_v4";
const IDB_STORE = "archivos";
let _idb = null;

/** Abre (o reutiliza) la conexión a IndexedDB */
function openIDB() {
  if (_idb) return Promise.resolve(_idb);
  return new Promise((res, rej) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = (e) =>
      e.target.result.createObjectStore(IDB_STORE, { keyPath: "name" });
    req.onsuccess = (e) => {
      _idb = e.target.result;
      res(_idb);
    };
    req.onerror = () => rej(req.error);
  });
}

/** Guarda los bytes de un archivo en IDB y su nombre en localStorage */
async function saveFileIDB(name, bytes) {
  try {
    const db = await openIDB();
    db.transaction(IDB_STORE, "readwrite")
      .objectStore(IDB_STORE)
      .put({ name, bytes });
    const meta = getFileMeta();
    if (!meta.find((m) => m.name === name)) {
      meta.push({ name });
      localStorage.setItem(LS_KEY, JSON.stringify(meta));
    }
  } catch (e) {
    console.warn("[IDB] saveFileIDB:", e);
  }
}

/** Elimina un archivo de IDB y de los metadatos */
async function removeFileIDB(name) {
  try {
    const db = await openIDB();
    db.transaction(IDB_STORE, "readwrite").objectStore(IDB_STORE).delete(name);
    localStorage.setItem(
      LS_KEY,
      JSON.stringify(getFileMeta().filter((m) => m.name !== name)),
    );
  } catch (e) {}
}

/** Lee los bytes de un archivo desde IDB */
function readFileIDB(db, name) {
  return new Promise((res, rej) => {
    const req = db
      .transaction(IDB_STORE, "readonly")
      .objectStore(IDB_STORE)
      .get(name);
    req.onsuccess = () => res(req.result?.bytes || null);
    req.onerror = () => rej(req.error);
  });
}

/** Devuelve los metadatos guardados en localStorage */
function getFileMeta() {
  try {
    return JSON.parse(localStorage.getItem(LS_KEY) || "[]");
  } catch {
    return [];
  }
}

/** Limpia IDB y localStorage completamente */
async function clearAllStorage() {
  localStorage.removeItem(LS_KEY);
  try {
    const db = await openIDB();
    db.transaction(IDB_STORE, "readwrite").objectStore(IDB_STORE).clear();
  } catch {}
}

// =============================================================================
//  4. UPLOAD Y GESTIÓN DE ARCHIVOS
// =============================================================================

// Referencias DOM
const fileInput = document.getElementById("fileInput");
const btnUpload = document.getElementById("btnUpload"); // puede ser null en modo Supabase
const toastEl = document.getElementById("toast");
const loadingEl = document.getElementById("loading");
const emptyState = document.getElementById("emptyState");
const dashGrid = document.getElementById("dashGrid");
const kpiStrip = document.getElementById("kpiStrip");
const filtersRow = document.getElementById("filtersRow");
const sheetRow = document.getElementById("sheetRow");
const babySection = document.getElementById("babySection");
const uploadZone = document.getElementById("supabaseZone"); // zona principal renombrada

const LS_RANGO_KEY = "juanfe_dash_ultimo_rango";

document.addEventListener("DOMContentLoaded", () => {
  fileInput.setAttribute("multiple", "true");
  setupUpload();
  setupFilters();
  setupBabySearch();
  setupSupabaseLoader(); // configura botones de carga desde BD
  // Restaurar el último rango de fechas si existe
  restaurarUltimoRango();
});

/** Configura los botones y fechas del panel de carga desde Supabase */
function setupSupabaseLoader() {
  const btnCargar = document.getElementById("btnCargarBD");
  const btnHoy = document.getElementById("btnHoy");
  const btnCambiar = document.getElementById("btnCambiarFecha");
  const inputDesde = document.getElementById("fechaDesde");
  const inputHasta = document.getElementById("fechaHasta");

  if (!btnCargar) return;

  // Poner fecha de hoy por defecto en los inputs
  const hoy = new Date().toISOString().split("T")[0];
  inputDesde.value = hoy;
  inputHasta.value = hoy;

  // Botón "Cargar datos"
  btnCargar.addEventListener("click", (e) => {
    e.stopPropagation();
    const desde = inputDesde.value;
    const hasta = inputHasta.value;
    if (!desde && !hasta) {
      showToast("Selecciona al menos una fecha", true);
      return;
    }
    cargarDesdeSupabase(desde, hasta);
  });

  // Botón "Hoy"
  if (btnHoy)
    btnHoy.addEventListener("click", (e) => {
      e.stopPropagation();
      inputDesde.value = hoy;
      inputHasta.value = hoy;
      cargarDesdeSupabase(hoy, hoy);
    });

  // Botón "Cambiar rango" — muestra el prompt de nuevo
  if (btnCambiar) {
    btnCambiar.addEventListener("click", () => {
      document.getElementById("fileLoaded").classList.remove("show");
      document.getElementById("supabasePrompt").style.display = "";
      const uploadZone = document.getElementById("supabaseZone");
      if (uploadZone) uploadZone.classList.remove("has-file");
    });
  }
}

/**
 * cargarDesdeSupabase(desde, hasta)
 * Consulta GET /api/asistencia con rango de fechas y carga los registros
 * directamente al modelo de datos — sin necesidad de Excels.
 */
async function cargarDesdeSupabase(desde, hasta) {
  try {
    showLoading(true);
    updateLoadingText("Consultando base de datos...");

    // Construir URL con filtros de fecha
    let url = "/api/asistencia";
    const params = [];
    if (desde) params.push(`desde=${desde}`);
    if (hasta) params.push(`hasta=${hasta}`);
    if (params.length) url += "?" + params.join("&");

    const res = await authFetch(url);
    if (!res.ok) throw new Error(`Error ${res.status}`);
    const { registros, total } = await res.json();

    if (!total || !registros?.length) {
      showLoading(false);
      showToast("No hay registros en ese rango de fechas", true);
      return;
    }

    // Limpiar modelo anterior y cargar los nuevos registros
    UnifiedModel.clear();
    allData = [];
    loadedFiles = [];

    // Agrupar por fecha+dia para que cada día sea una "fuente" independiente
    const grupos = {};
    for (const r of registros) {
      const key = `${r.Fecha}|${r.Dia}`;
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(r);
    }

    // Ingestar cada grupo como si fuera un archivo distinto
    for (const [key, rows] of Object.entries(grupos)) {
      const [fecha, dia] = key.split("|");
      const nombreFuente = `${dia}-${fecha}`;
      UnifiedModel.ingest(rows, nombreFuente);
      loadedFiles.push({
        name: nombreFuente,
        rowCount: rows.length,
        sheetCount: 1,
      });
    }

    allData = UnifiedModel.getAll();
    headers = CANONICAL_HEADERS;

    // Guardar el rango para restaurar al refrescar
    localStorage.setItem(LS_RANGO_KEY, JSON.stringify({ desde, hasta }));

    showLoading(false);
    buildFileChips();
    showUI();
    rebuildAll();

    const meta = UnifiedModel.getMeta();
    const rangoLabel =
      desde && hasta ? `${desde} → ${hasta}` : desde || hasta || "todos";
    showToast(
      `BD · ${meta.bebesUnicos} bebés · ${meta.registros} registros · ${rangoLabel}`,
    );
  } catch (err) {
    showLoading(false);
    showToast("No se pudo conectar a la base de datos: " + err.message, true);
    console.error("[Supabase] Error cargando datos:", err);
  }
}

/** Carga automática del día de hoy al abrir el dashboard */
async function cargarHistoricoServidor() {
  // Intentar cargar desde Supabase automáticamente con la fecha de hoy
  try {
    const hoy = new Date().toISOString().split("T")[0];
    const res = await authFetch(`/api/asistencia?fecha=${hoy}`);
    if (!res.ok) return;
    const { total } = await res.json();
    if (total > 0) {
      // Hay datos de hoy → cargar automáticamente
      await cargarDesdeSupabase(hoy, hoy);
    }
    // Si no hay datos de hoy, el dashboard queda vacío esperando que el usuario elija rango
  } catch {
    // Sin servidor — modo standalone con Excels manuales
  }
}

/**
 * restaurarUltimoRango()
 * Al refrescar la página, recupera el último rango de fechas cargado
 * y vuelve a consultar Supabase automáticamente.
 */
async function restaurarUltimoRango() {
  try {
    const raw = localStorage.getItem(LS_RANGO_KEY);
    if (!raw) return;
    const { desde, hasta } = JSON.parse(raw);
    if (!desde && !hasta) return;

    // Actualizar los inputs de fecha para que reflejen el rango restaurado
    const inputDesde = document.getElementById("fechaDesde");
    const inputHasta = document.getElementById("fechaHasta");
    if (inputDesde && desde) inputDesde.value = desde;
    if (inputHasta && hasta) inputHasta.value = hasta;

    // Cargar los datos automáticamente
    await cargarDesdeSupabase(desde, hasta);
  } catch (e) {
    console.warn("[Restaurar] No se pudo restaurar el rango:", e);
  }
}

function setupUpload() {
  // btnUpload puede no existir si se usa el modo Supabase
  if (btnUpload) {
    btnUpload.addEventListener("click", (e) => {
      e.stopPropagation();
      fileInput.click();
    });
  }

  // Click en la zona de carga (evitar disparar si se hizo clic en el estado "cargado")
  if (uploadZone)
    uploadZone.addEventListener("click", (e) => {
      if (e.target.closest("#fileLoaded")) return;
      if (e.target.closest("#supabasePrompt")) return; // no abrir file picker al clickear el prompt
      fileInput.click();
    });

  fileInput.addEventListener("change", (e) => {
    const files = Array.from(e.target.files);
    if (files.length) processFiles(files);
    fileInput.value = "";
  });

  // Drag & drop
  if (uploadZone)
    uploadZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      uploadZone.classList.add("dragover");
    });
  if (uploadZone)
    uploadZone.addEventListener("dragleave", () =>
      uploadZone.classList.remove("dragover"),
    );
  if (uploadZone)
    uploadZone.addEventListener("drop", (e) => {
      e.preventDefault();
      uploadZone.classList.remove("dragover");
      const files = Array.from(e.dataTransfer.files).filter((f) =>
        /\.(xlsx|xls|csv)$/i.test(f.name),
      );
      if (files.length) processFiles(files);
    });
}

/**
 * Procesa un array de archivos File de forma secuencial y asíncrona.
 * Actualiza el spinner con progreso para no bloquear la UI con 20+ archivos.
 */
async function processFiles(files) {
  showLoading(true);
  const resultsBatch = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    updateLoadingText(`Procesando ${i + 1}/${files.length}: ${file.name}`);
    await yieldToUI();

    try {
      const bytes = await readFileAsUint8(file);
      const wb = XLSX.read(bytes, { type: "array" });
      let rowCount = 0;

      for (const shName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[shName], {
          defval: "",
        });
        if (rows.length > 0) {
          rowCount += rows.length;
          resultsBatch.push({ rows, fileName: file.name });
        }
      }

      // Guardar solo si no estaba ya en sesión
      if (!loadedFiles.find((lf) => lf.name === file.name)) {
        loadedFiles.push({
          name: file.name,
          sheetCount: wb.SheetNames.length,
          rowCount,
        });
      }
      saveFileIDB(file.name, bytes); // asíncrono, no bloqueante
    } catch {
      showToast(`No se pudo leer "${file.name}"`, true);
    }
  }

  finishLoad(resultsBatch);
}

/** Lee un File → Uint8Array (promisificado) */
function readFileAsUint8(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(new Uint8Array(e.target.result));
    r.onerror = () => rej(r.error);
    r.readAsArrayBuffer(file);
  });
}

/** Cede el control al navegador por un frame (evita UI congelada) */
function yieldToUI() {
  return new Promise((r) => setTimeout(r, 0));
}

/** Actualiza el texto del spinner */
function updateLoadingText(text) {
  const el = document.querySelector(".loading-text");
  if (el) el.textContent = text;
}

/**
 * finishLoad: punto de entrada único al modelo de datos.
 * Todos los archivos pasan por UnifiedModel.ingest() antes de renderizar.
 */
function finishLoad(resultsBatch) {
  for (const { rows, fileName } of resultsBatch) {
    UnifiedModel.ingest(rows, fileName);
  }

  // allData se obtiene SIEMPRE desde el modelo — nunca se manipula directamente
  allData = UnifiedModel.getAll();
  headers = CANONICAL_HEADERS;

  showLoading(false);
  buildFileChips();
  showUI();
  rebuildAll();

  const meta = UnifiedModel.getMeta();
  showToast(
    `${loadedFiles.length} archivo(s) · ${meta.bebesUnicos} bebés · ${meta.registros} registros`,
  );
}

/** Construye los chips de archivos cargados */
function buildFileChips() {
  const container = document.getElementById("sheetTabs");
  container.innerHTML = "";
  for (const f of loadedFiles) {
    const chip = document.createElement("div");
    chip.className = "file-chip";
    chip.innerHTML = `
      <span class="file-chip-icon">📄</span>
      <span class="file-chip-name">${f.name}</span>
      <span class="file-chip-rows">${f.rowCount} filas</span>
      <button class="file-chip-remove" data-name="${f.name}" title="Quitar">✕</button>
    `;
    chip.querySelector(".file-chip-remove").addEventListener("click", (e) => {
      e.stopPropagation();
      removeFile(f.name);
    });
    container.appendChild(chip);
  }
  sheetRow.classList.add("show");
}

/** Quita un archivo de la sesión y re-ingesta los restantes */
function removeFile(name) {
  loadedFiles = loadedFiles.filter((f) => f.name !== name);
  removeFileIDB(name);
  UnifiedModel.clear();
  allData = [];
  headers = [];

  if (loadedFiles.length === 0) {
    clearSession();
    return;
  }

  // Re-ingestar desde IDB los archivos que quedan
  openIDB().then(async (db) => {
    const resultsBatch = [];
    for (const lf of loadedFiles) {
      try {
        const bytes = await readFileIDB(db, lf.name);
        if (!bytes) continue;
        const wb = XLSX.read(bytes, { type: "array" });
        for (const shName of wb.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[shName], {
            defval: "",
          });
          if (rows.length > 0) resultsBatch.push({ rows, fileName: lf.name });
        }
      } catch {}
    }
    finishLoad(resultsBatch);
  });
}

/** Restaura la sesión desde IDB al recargar la página */
async function restoreSession() {
  const meta = getFileMeta();
  if (!meta.length) return;
  showLoading(true);
  updateLoadingText("Restaurando sesión...");
  try {
    const db = await openIDB();
    const resultsBatch = [];
    for (let i = 0; i < meta.length; i++) {
      const { name } = meta[i];
      updateLoadingText(`Restaurando ${i + 1}/${meta.length}...`);
      await yieldToUI();
      try {
        const bytes = await readFileIDB(db, name);
        if (!bytes) {
          await removeFileIDB(name);
          continue;
        }
        const wb = XLSX.read(bytes, { type: "array" });
        let rowCount = 0;
        for (const shName of wb.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[shName], {
            defval: "",
          });
          if (rows.length > 0) {
            rowCount += rows.length;
            resultsBatch.push({ rows, fileName: name });
          }
        }
        if (!loadedFiles.find((lf) => lf.name === name))
          loadedFiles.push({
            name,
            sheetCount: wb.SheetNames.length,
            rowCount,
          });
      } catch {
        await removeFileIDB(name);
      }
    }
    if (resultsBatch.length) finishLoad(resultsBatch);
  } catch (e) {
    console.warn("[IDB] restoreSession:", e);
  }
  showLoading(false);
}

/** Limpia toda la sesión */
window.clearSession = function () {
  localStorage.removeItem(LS_RANGO_KEY);
  clearAllStorage();
  UnifiedModel.clear();
  allData = [];
  headers = [];
  loadedFiles = [];
  selectedBaby = null;
  activeFilter = "todos";
  clearCharts();
  dashGrid.innerHTML = "";

  // Ocultar secciones
  [dashGrid, filtersRow, sheetRow, babySection].forEach((el) =>
    el.classList.remove("show"),
  );
  const ksm = document.getElementById("kpiStripMulti");
  const ksm2 = document.getElementById("kpiStripMulti2");
  const krd = document.getElementById("kpiResumenDias");
  if (kpiStrip) kpiStrip.style.display = "none";
  if (ksm) ksm.style.display = "none";
  if (ksm2) ksm2.style.display = "none";
  if (krd) krd.style.display = "none";

  // Mostrar selector de fechas (no el empty state)
  emptyState.style.display = "none";
  const _up = document.getElementById("uploadPrompt");
  if (_up) _up.style.display = "none";

  // Volver a mostrar el prompt de Supabase
  const sp = document.getElementById("supabasePrompt");
  const fl = document.getElementById("fileLoaded");
  if (sp) {
    sp.style.display = "";
    sp.style.visibility = "visible";
  }
  if (fl) fl.classList.remove("show");
  if (uploadZone) uploadZone.classList.remove("has-file");

  document.getElementById("sheetTabs").innerHTML = "";
  document.getElementById("babySearch").value = "";
  document.getElementById("selectedBabyTag").style.display = "none";
  document
    .querySelectorAll(".filter-btn")
    .forEach((b) => b.classList.toggle("active", b.dataset.filter === "todos"));
  fileInput.value = "";
};

// =============================================================================
//  5. NORMALIZACIÓN Y UTILIDADES DE DATOS
// =============================================================================

/** Normaliza texto: minúsculas + sin tildes */
function norm(str) {
  return String(str)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

/** Resuelve el nombre de columna canónico para un tipo de dato */
function colKey(tipo) {
  return (COL[tipo] || []).find((c) => headers.includes(c)) || null;
}

/** Lee el valor de una columna en un registro, devuelve string vacío si no existe */
function val(row, tipo) {
  const c = colKey(tipo);
  return c ? String(row[c] || "").trim() : "";
}

/** Evalúa si un valor representa "Sí" en cualquiera de sus variantes */
function esSi(v) {
  const n = norm(v);
  return n === "si" || n === "sí" || n === "1" || n === "true";
}

/**
 * asistio(r) — fuente de verdad ÚNICA para evaluar asistencia.
 * • Resumen (getResumenBebes): usa _diasPresente acumulado → EXACTO
 * • Registro individual (getAllRegistros): usa campo Asistencia → CORRECTO
 * NUNCA mezclar: en el resumen, .Asistencia solo refleja el 1er registro.
 */
function asistio(r) {
  if (r._diasPresente !== undefined) return r._diasPresente > 0;
  return esSi(val(r, "asistencia"));
}

/**
 * reporto(r) — fuente de verdad ÚNICA para evaluar reportes.
 * Mismo principio que asistio(): _reportes para resumen, campo Reporte para registros.
 */
function reporto(r) {
  if (r._reportes !== undefined) return r._reportes > 0;
  return esSi(val(r, "reporte"));
}

// =============================================================================
//  6. DOS VISTAS DE LOS DATOS
//
//  getResumenBebes(data)
//    → 1 entrada por bebé (nombre+madre), con contadores acumulados de todos sus días.
//    → Incluye Extras y NoCidi (son personas reales con asistencia real).
//    → Usada exclusivamente para los KPIs del header.
//
//  getAllRegistros(data)
//    → Todos los registros individuales (bebé × día), sin duplicar el mismo bebé en el mismo día.
//    → Usada para la tabla, gráficas temporales, ETL y exportación CSV.
// =============================================================================

/**
 * getResumenBebes(data) → Array
 * Colapsa múltiples registros del mismo bebé en una sola entrada con acumulados.
 * La clave de identidad es NombreBebe + NombreMadre para distinguir homónimos.
 */
function getResumenBebes(data) {
  // REGLA CRÍTICA: _diasPresente y _reportes son la ÚNICA fuente de verdad.
  // Nunca se usa .Asistencia ni .Reporte para contar en el resumen,
  // porque esos campos reflejan solo el primer registro ingresado.
  const map = new Map();
  for (const r of data) {
    const nombreNorm = norm(val(r, "bebe"));
    if (!nombreNorm) continue;
    const bKey = nombreNorm + "|" + norm(val(r, "madre") || "");

    if (!map.has(bKey)) {
      map.set(bKey, {
        ...r,
        _diasTotal: 1,
        _diasPresente: esSi(val(r, "asistencia")) ? 1 : 0,
        _reportes: esSi(val(r, "reporte")) ? 1 : 0,
        // BUG FIX: _esVisitante y _esNoCidi se acumulan igual que _diasPresente.
        // Sin esto, el resumen solo lee el campo del 1er registro ingresado y
        // no detecta bebés que son Extras en un día pero regulares en otro.
        _esVisitante: esSi(val(r, "visitante")) ? 1 : 0,
        _esNoCidi: esSi(val(r, "nocidi")) ? 1 : 0,
      });
    } else {
      const prev = map.get(bKey);
      prev._diasTotal++;
      if (esSi(val(r, "asistencia"))) prev._diasPresente++;
      if (esSi(val(r, "reporte"))) prev._reportes++;
      if (esSi(val(r, "visitante"))) prev._esVisitante++; // ← nuevo
      if (esSi(val(r, "nocidi"))) prev._esNoCidi++; // ← nuevo
      // Enriquecer perfil con campos vacíos de días posteriores
      if (!val(prev, "institucion") && val(r, "institucion"))
        prev.InstitucionMadre = val(r, "institucion");
      if (!val(prev, "programa") && val(r, "programa"))
        prev.ProgramaMadre = val(r, "programa");
      if (!val(prev, "edad") && val(r, "edad")) prev.Edad = val(r, "edad");
      if (!val(prev, "madre") && val(r, "madre"))
        prev.NombreMadre = val(r, "madre");
    }
  }
  return Array.from(map.values());
}

/**
 * getAllRegistros(data) → Array
 * Deduplication: mismo bebé (nombre+madre) no puede aparecer dos veces en el mismo día+fecha.
 * Mantiene todos los días distintos de cada bebé.
 */
function getAllRegistros(data) {
  const seen = new Set();
  return data.filter((r) => {
    const bebe = norm(val(r, "bebe"));
    if (!bebe) return false;
    const key = `${bebe}|${norm(val(r, "madre") || "")}|${norm(val(r, "fecha") || "")}|${norm(val(r, "dia") || "")}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// =============================================================================
//  7. UI — KPIs, FILTROS, BÚSQUEDA DE BEBÉ
// =============================================================================

function showUI() {
  emptyState.style.display = "none";
  [filtersRow, babySection, dashGrid].forEach((el) => el.classList.add("show"));
  // Ocultar el prompt de Supabase y mostrar estado "cargado"
  try {
    const supabasePrompt = document.getElementById("supabasePrompt");
    if (supabasePrompt) supabasePrompt.style.display = "none";
    const supabaseZone = document.getElementById("supabaseZone");
    if (supabaseZone) supabaseZone.classList.add("has-file");
  } catch (e) {
    /* ignorar si los elementos no existen */
  }
  // KPI strips se muestran según modo en updateKPIs
  const _upS = document.getElementById("uploadPrompt");
  if (_upS) _upS.style.display = "none";
  const _flS = document.getElementById("fileLoaded");
  if (_flS) _flS.classList.add("show");
  if (uploadZone) uploadZone.classList.add("has-file");

  // Mostrar conteos reales en el área de carga
  const registrosUnicos = getAllRegistros(allData).length;
  const bebesUnicos = new Set(
    allData
      .map((r) => {
        const n = norm(val(r, "bebe"));
        return n ? n + "|" + norm(val(r, "madre") || "") : null;
      })
      .filter(Boolean),
  ).size;
  document.getElementById("fileSheetInfo").textContent =
    `${bebesUnicos} bebés · ${registrosUnicos} registros · ${loadedFiles.length} archivo(s)`;

  // Botón de exportación consolidada (ETL) — se inyecta una sola vez
  if (!document.getElementById("btnConsolidar")) {
    const btn = document.createElement("button");
    btn.id = "btnConsolidar";
    btn.title =
      "Exporta una tabla única limpia y sin duplicados, lista para dashboards";
    btn.style.cssText = `
      display:inline-flex; align-items:center; gap:8px;
      padding:7px 16px; border-radius:8px; cursor:pointer;
      border:1.5px solid var(--green); background:var(--green);
      color:white; font-family:'JetBrains Mono',monospace;
      font-size:11px; font-weight:700; letter-spacing:.5px;
      transition:opacity .15s; white-space:nowrap; margin-left:8px;
    `;
  }
}

function setupFilters() {
  document.querySelectorAll(".filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document
        .querySelectorAll(".filter-btn")
        .forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      activeFilter = btn.dataset.filter;
      rebuildAll();
    });
  });
}

function setupBabySearch() {
  const input = document.getElementById("babySearch");
  const dropdown = document.getElementById("babyDropdown");

  input.addEventListener("input", () => {
    const term = norm(input.value.trim());
    dropdown.innerHTML = "";
    if (!term) {
      dropdown.style.display = "none";
      return;
    }

    const allNames = UnifiedModel.getBabyNames();
    const matches = allNames.filter((n) => norm(n).includes(term)).slice(0, 10);
    if (!matches.length) {
      dropdown.style.display = "none";
      return;
    }

    matches.forEach((name, idx) => {
      const item = document.createElement("div");
      item.className = "baby-dropdown-item";
      const color = PALETTE[idx % PALETTE.length];
      item.innerHTML = `<span style="width:8px;height:8px;border-radius:50%;background:${color};display:inline-block;margin-right:10px;flex-shrink:0"></span>${name}`;
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        input.value = name;
        dropdown.style.display = "none";
        selectBaby(name);
      });
      dropdown.appendChild(item);
    });
    dropdown.style.display = "block";
  });

  input.addEventListener("blur", () =>
    setTimeout(() => {
      dropdown.style.display = "none";
    }, 150),
  );
  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      dropdown.style.display = "none";
      input.blur();
    }
    if (e.key === "Enter") {
      const first = dropdown.querySelector(".baby-dropdown-item");
      if (first) first.dispatchEvent(new MouseEvent("mousedown"));
    }
  });

  document.getElementById("btnClearBaby").addEventListener("click", () => {
    selectedBaby = null;
    input.value = "";
    dropdown.style.display = "none";
    document.getElementById("selectedBabyTag").style.display = "none";
    rebuildAll();
  });
}

function selectBaby(name) {
  selectedBaby = name;
  document.getElementById("selectedBabyName").textContent = name;
  document.getElementById("selectedBabyTag").style.display = "flex";
  rebuildAll();
}

// =============================================================================
//  8. KPIs
//  Los contadores del header son FIJOS: se calculan siempre sobre allData
//  completo y nunca cambian al cambiar el filtro activo.
// =============================================================================

/**
 * rebuildAll()
 * Punto de entrada de cada interacción (cambio de filtro, búsqueda, carga).
 * Siempre actualiza KPIs con el total completo, luego renderiza la vista activa.
 */
function rebuildAll() {
  const resumenCompleto = getResumenBebes(allData);
  updateKPIs(resumenCompleto, allData);

  if (selectedBaby) {
    // Vista de historial individual
    const registrosDelBebe = getAllRegistros(
      allData.filter((r) => val(r, "bebe") === selectedBaby),
    );
    buildBabyHistorial(registrosDelBebe);
  } else {
    buildDashboardForFilter(activeFilter);
  }
}

function updateKPIs(resumen, todosLosRegistros) {
  // ── Detectar modo: 1 día vs varios días ─────────────────────────────────
  const diasUnicos = [
    ...new Set(
      todosLosRegistros
        .map((r) => val(r, "dia") || val(r, "fecha"))
        .filter(Boolean),
    ),
  ];
  const fechasUnicas = [
    ...new Set(todosLosRegistros.map((r) => val(r, "fecha")).filter(Boolean)),
  ];
  const esModoUnDia =
    loadedFiles.length === 1 ||
    diasUnicos.length === 1 ||
    fechasUnicas.length === 1;

  const kpiStrip1 = document.getElementById("kpiStrip");
  const kpiStripM = document.getElementById("kpiStripMulti");
  const kpiResumen = document.getElementById("kpiResumenDias");

  if (esModoUnDia) {
    // ── MODO 1 DÍA ────────────────────────────────────────────────────────
    kpiStrip1.style.display = "";
    kpiStripM.style.display = "none";
    kpiResumen.style.display = "none";

    const total = resumen.length;
    const presentes = resumen.filter((r) => (r._diasPresente || 0) > 0).length;
    const ausentes = total - presentes;
    const reportes = resumen.filter((r) => (r._reportes || 0) > 0).length;
    const nExtras = resumen.filter((r) => (r._esVisitante || 0) > 0).length;
    const nNoCidi = resumen.filter((r) => (r._esNoCidi || 0) > 0).length;
    const diaLabel =
      diasUnicos[0] || fechasUnicas[0] || loadedFiles[0]?.name || "—";

    document.getElementById("kpiTotal").textContent = total;
    document.getElementById("kpiDia").textContent = diaLabel;
    document.getElementById("kpiPresentes").textContent = presentes;
    document.getElementById("kpiPresentePct").textContent =
      `${total ? Math.round((presentes / total) * 100) : 0}% de inscritos`;
    document.getElementById("kpiAusentes").textContent = ausentes;
    document.getElementById("kpiAusentePct").textContent =
      `${total ? Math.round((ausentes / total) * 100) : 0}% de inscritos`;
    document.getElementById("kpiReportes").textContent = reportes;
    const elE = document.getElementById("kpiExtras");
    const elN = document.getElementById("kpiNocidi");
    if (elE) elE.textContent = nExtras;
    if (elN) elN.textContent = nNoCidi;
  } else {
    // ── MODO VARIOS DÍAS ──────────────────────────────────────────────────
    kpiStrip1.style.display = "none";
    kpiStripM.style.display = "";
    kpiResumen.style.display = "";

    // Estadísticas por día/archivo — base real para tasa
    const statsPorDia = {};
    for (const r of todosLosRegistros) {
      const diaKey = val(r, "fecha") || val(r, "dia") || "?";
      const label = val(r, "dia") || val(r, "fecha") || "?";
      if (!statsPorDia[diaKey])
        statsPorDia[diaKey] = {
          label,
          inscritos: 0,
          presentes: 0,
          ausentes: 0,
          reportes: 0,
          extras: 0,
        };
      const s = statsPorDia[diaKey];
      s.inscritos++;
      if (esSi(val(r, "asistencia"))) s.presentes++;
      else s.ausentes++;
      if (esSi(val(r, "reporte"))) s.reportes++;
      if (esSi(val(r, "visitante"))) s.extras++;
    }

    const dias = Object.values(statsPorDia);
    dias.forEach((d) => {
      d.tasa = d.inscritos ? Math.round((d.presentes / d.inscritos) * 100) : 0;
    });

    // KPIs globales
    const totalBebesUnicos = resumen.length;
    const totalPresencias = dias.reduce((s, d) => s + d.presentes, 0);
    const totalInscritos = dias.reduce((s, d) => s + d.inscritos, 0);
    const tasaPromedio = totalInscritos
      ? Math.round((totalPresencias / totalInscritos) * 100)
      : 0;
    const nuncaVinieron = resumen.filter(
      (r) => (r._diasPresente || 0) === 0,
    ).length;
    const totalReportes = resumen.filter((r) => (r._reportes || 0) > 0).length;
    const mejorDia = dias.reduce((a, b) => (b.tasa > a.tasa ? b : a), dias[0]);
    const peorDia = dias.reduce((a, b) => (b.tasa < a.tasa ? b : a), dias[0]);

    document.getElementById("kpiTotalMulti").textContent = totalBebesUnicos;
    document.getElementById("kpiDiasMulti").textContent =
      `bebés distintos · ${dias.length} sesión(es)`;
    document.getElementById("kpiTasaMulti").textContent = `${tasaPromedio}%`;
    document.getElementById("kpiMejorDia").textContent = `${mejorDia?.tasa}%`;
    document.getElementById("kpiMejorDiaNombre").textContent =
      mejorDia?.label || "—";
    document.getElementById("kpiPeorDia").textContent = `${peorDia?.tasa}%`;
    document.getElementById("kpiPeorDiaNombre").textContent =
      peorDia?.label || "—";
    document.getElementById("kpiNuncaMulti").textContent = nuncaVinieron;
    document.getElementById("kpiReportesMulti").textContent = totalReportes;

    // ── Tabla resumen por día con semáforo ───────────────────────────────
    const DIAS_ORD = [
      "Lunes",
      "Martes",
      "Miércoles",
      "Jueves",
      "Viernes",
      "Sábado",
      "Domingo",
    ];
    dias.sort((a, b) => {
      const ai = DIAS_ORD.indexOf(a.label);
      const bi = DIAS_ORD.indexOf(b.label);
      if (ai !== -1 && bi !== -1) return ai - bi;
      return a.label.localeCompare(b.label);
    });

    const colorSemaforo = (tasa) =>
      tasa >= 70
        ? {
            bg: "#f0fbe6",
            border: "#85da1a",
            text: "#3a6b00",
            badge: "#85da1a",
          }
        : tasa >= 50
          ? {
              bg: "#fff7ed",
              border: "#f97316",
              text: "#7c3500",
              badge: "#f97316",
            }
          : {
              bg: "#fef2f2",
              border: "#ef4444",
              text: "#7f1d1d",
              badge: "#ef4444",
            };

    const filas = dias
      .map((d) => {
        const c = colorSemaforo(d.tasa);
        const emoji = d.tasa >= 70 ? "🟢" : d.tasa >= 50 ? "🟡" : "🔴";
        return `
        <tr style="background:${c.bg}; border-left: 3px solid ${c.border};">
          <td style="padding:10px 16px; font-weight:700; color:${c.text}; font-size:13px">${emoji} ${d.label}</td>
          <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px; color:#2d3748">${d.inscritos}</td>
          <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px; color:#16a34a; font-weight:700">${d.presentes}</td>
          <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px; color:#dc2626; font-weight:700">${d.ausentes}</td>
          <td style="padding:10px 16px; text-align:center;">
            <span style="background:${c.badge}; color:white; padding:3px 10px; border-radius:99px; font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:700">${d.tasa}%</span>
          </td>
          <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:12px; color:#f97316">${d.reportes > 0 ? `⚠️ ${d.reportes}` : "—"}</td>
          <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:12px; color:#3b82f6">${d.extras > 0 ? `🔵 ${d.extras}` : "—"}</td>
        </tr>`;
      })
      .join("");

    // Fila totales
    const c0 = colorSemaforo(tasaPromedio);
    const filaTotal = `
      <tr style="background:#f8fafc; border-top: 2px solid #e0eccc; font-weight:700">
        <td style="padding:10px 16px; font-family:'JetBrains Mono',monospace; font-size:12px; letter-spacing:1px; color:#8ba869; text-transform:uppercase">TOTAL / PROMEDIO</td>
        <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px">${totalInscritos}</td>
        <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px; color:#16a34a">${totalPresencias}</td>
        <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:13px; color:#dc2626">${totalInscritos - totalPresencias}</td>
        <td style="padding:10px 16px; text-align:center;">
          <span style="background:${c0.badge}; color:white; padding:3px 10px; border-radius:99px; font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:700">${tasaPromedio}%</span>
        </td>
        <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:12px; color:#f97316">${dias.reduce((s, d) => s + d.reportes, 0)}</td>
        <td style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:12px; color:#3b82f6">${dias.reduce((s, d) => s + d.extras, 0)}</td>
      </tr>`;

    kpiResumen.innerHTML = `
      <div style="background:white; border:1px solid #e0eccc; border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(133,218,26,.08);">
        <div style="padding:14px 20px; background:#f8fef0; border-bottom:1px solid #e0eccc; display:flex; align-items:center; gap:10px;">
          <span style="font-size:16px">📅</span>
          <span style="font-family:'JetBrains Mono',monospace; font-size:11px; letter-spacing:2px; font-weight:700; color:#8ba869; text-transform:uppercase">Resumen por día</span>
          <span style="margin-left:auto; font-size:11px; color:#a3b899">🟢 ≥70% &nbsp; 🟡 50-69% &nbsp; 🔴 &lt;50%</span>
        </div>
        <div style="padding:8px 20px 6px; background:#fffdf5; border-bottom:1px solid #f0e8cc; font-family:'JetBrains Mono',monospace; font-size:10px; color:#b08840; letter-spacing:.5px;">
          📌 Cada fila es un día de clase. Cada bebé aparece en la fila de cada día que asiste — por eso el total aquí puede ser mayor que el número de bebés de arriba.
        </div>
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#f0fbe6;">
                <th style="padding:10px 16px; text-align:left; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#8ba869; text-transform:uppercase; font-weight:700">Día</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#8ba869; text-transform:uppercase; font-weight:700">Inscripciones</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#16a34a; text-transform:uppercase; font-weight:700">Presentes</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#dc2626; text-transform:uppercase; font-weight:700">Ausentes</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#8ba869; text-transform:uppercase; font-weight:700">Tasa</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#f97316; text-transform:uppercase; font-weight:700">Reportes</th>
                <th style="padding:10px 16px; text-align:center; font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:1.5px; color:#3b82f6; text-transform:uppercase; font-weight:700">Extras</th>
              </tr>
            </thead>
            <tbody>${filas}</tbody>
            <tfoot>${filaTotal}</tfoot>
          </table>
        </div>
      </div>`;
  }
}

// =============================================================================
//  9. DASHBOARD POR FILTRO
// =============================================================================

/**
 * buildDashboardForFilter(filter)
 * Renderiza la vista de gráficas y tabla según el filtro activo.
 * Cada filtro usa sus propios datos exactos — no mezcla subconjuntos.
 */
function buildDashboardForFilter(filter) {
  clearCharts();
  dashGrid.innerHTML = "";

  // Datos base (sin duplicar bebé+día)
  const registros = getAllRegistros(allData);
  // Resumen (1 por bebé, con acumulados)
  const resumen = getResumenBebes(allData);

  // Subconjuntos para cada filtro
  // NOTA: Extras y NoCidi se clasifican según su campo Asistencia real
  const presentes = getAllRegistros(
    allData.filter((r) => esSi(val(r, "asistencia"))),
  );
  const ausentes = getAllRegistros(
    allData.filter((r) => !esSi(val(r, "asistencia"))),
  );
  const reportados = getAllRegistros(
    allData.filter((r) => esSi(val(r, "reporte"))),
  );
  const extras = getAllRegistros(
    allData.filter((r) => esSi(val(r, "visitante"))),
  );
  const nocidi = getAllRegistros(allData.filter((r) => esSi(val(r, "nocidi"))));

  // Conteos exactos para los encabezados de vista (mismo criterio que KPIs)
  const nTotal = resumen.length;
  const nPresentes = resumen.filter((r) => (r._diasPresente || 0) > 0).length;
  const nAusentes = nTotal - nPresentes;
  const nReportes = resumen.filter((r) => (r._reportes || 0) > 0).length;
  // FIX: usar _esVisitante/_esNoCidi (acumulados) igual que en updateKPIs
  const nExtras = resumen.filter((r) => (r._esVisitante || 0) > 0).length;
  const nNoCidi = resumen.filter((r) => (r._esNoCidi || 0) > 0).length;

  const hasDia = !!colKey("dia");
  const hasInst = !!colKey("institucion");
  const hasProg = !!colKey("programa");
  const hasEdad = !!colKey("edad");
  const hasRep = !!colKey("reporte");
  const hasSit = !!colKey("situacion");
  const hasUbic = !!colKey("ubicacion");

  switch (filter) {
    // ── TODOS — resumen general ───────────────────────────────
    case "todos":
    default: {
      buildChartDistribucion(resumen);
      if (hasRep) buildChartTasaReporte(registros); // registros individuales → cuenta cada reporte
      if (hasUbic) buildChartUbicacion(registros); // registros individuales para ubicación correcta
      if (hasDia) buildChartAsistenciaPorDia(registros);
      if (hasInst) buildChartInstitucion(registros);
      if (hasProg) buildChartPrograma(registros);
      if (hasEdad) buildChartEdad(resumen);
      if (hasSit && hasRep)
        buildCollapsible("causas", "⚠️ Causas de ausencia reportadas", () =>
          buildChartCausas(registros),
        );
      buildSeccionReportes(registros);
      buildETLSection(registros, resumen);
      buildTabla(registros);
      break;
    }

    // ── PRESENTES ────────────────────────────────────────────
    case "presentes": {
      if (!nPresentes) {
        showEmptyFilter("✅ Presentes");
        break;
      }
      const nExtPres = extras.filter((r) => esSi(val(r, "asistencia"))).length;
      buildViewHeader(
        "✅ Presentes",
        nPresentes,
        nTotal,
        "#22c55e",
        `${nPresentes} de ${nTotal} bebés asistieron${nExtPres ? ` · incluye ${nExtPres} extra(s)` : ""}`,
      );
      if (hasDia)
        buildChartBarFilter(
          presentes,
          "día",
          "📅 Presentes por día",
          "#22c55e",
        );
      if (hasInst)
        buildChartBarFilter(
          presentes,
          "institución",
          "🏛 Presentes por institución",
          "#22c55e",
        );
      if (hasProg)
        buildChartBarFilter(
          presentes,
          "programa",
          "🎓 Presentes por programa",
          "#22c55e",
        );
      if (hasEdad)
        buildChartBarFilter(
          presentes,
          "edad",
          "🍼 Por rango de edad",
          "#22c55e",
        );
      if (hasUbic)
        buildChartBarFilter(
          presentes,
          "ubicacion",
          "📍 Por ubicación",
          "#22c55e",
        );
      buildTabla(presentes);
      break;
    }

    // ── AUSENTES ─────────────────────────────────────────────
    case "ausentes": {
      if (!nAusentes) {
        showEmptyFilter("❌ Ausentes");
        break;
      }
      const nExtAus = extras.filter((r) => !esSi(val(r, "asistencia"))).length;
      buildViewHeader(
        "❌ Ausentes",
        nAusentes,
        nTotal,
        "#ef4444",
        `${nAusentes} de ${nTotal} bebés no asistieron${nExtAus ? ` · incluye ${nExtAus} extra(s)` : ""}`,
      );
      if (hasDia)
        buildChartBarFilter(ausentes, "día", "📅 Ausentes por día", "#ef4444");
      if (hasInst)
        buildChartBarFilter(
          ausentes,
          "institución",
          "🏛 Ausentes por institución",
          "#ef4444",
        );
      if (hasProg)
        buildChartBarFilter(
          ausentes,
          "programa",
          "🎓 Ausentes por programa",
          "#ef4444",
        );
      if (hasEdad)
        buildChartBarFilter(
          ausentes,
          "edad",
          "🍼 Por rango de edad",
          "#ef4444",
        );
      if (hasSit && hasRep)
        buildCollapsible("causas-aus", "⚠️ Causas de ausencia", () =>
          buildChartCausas(ausentes),
        );
      buildTabla(ausentes);
      break;
    }

    // ── CON REPORTE ──────────────────────────────────────────
    case "reportados": {
      if (!nReportes) {
        showEmptyFilter("⚠️ Con reporte");
        break;
      }
      buildViewHeader(
        "⚠️ Con reporte",
        nReportes,
        nTotal,
        "#f97316",
        `${nReportes} de ${nTotal} bebés tienen al menos un reporte`,
      );
      if (hasSit) buildChartCausasDirecto(reportados);
      if (hasDia)
        buildChartBarFilter(
          reportados,
          "día",
          "📅 Reportes por día",
          "#f97316",
        );
      if (hasInst)
        buildChartBarFilter(
          reportados,
          "institución",
          "🏛 Reportes por institución",
          "#f97316",
        );
      if (hasProg)
        buildChartBarFilter(
          reportados,
          "programa",
          "🎓 Reportes por programa",
          "#f97316",
        );
      buildSeccionReportes(reportados);
      buildTabla(reportados);
      break;
    }

    // ── EXTRAS ───────────────────────────────────────────────
    case "extras": {
      if (!nExtras) {
        showEmptyFilter("🔵 Extras");
        break;
      }
      buildViewHeader(
        "🔵 Extras (visitantes)",
        nExtras,
        nTotal,
        "#3b82f6",
        `${nExtras} visitante(s) del día · incluidos en el total de ${nTotal}`,
      );
      if (hasDia)
        buildChartBarFilter(extras, "día", "📅 Extras por día", "#3b82f6");
      if (hasInst)
        buildChartBarFilter(
          extras,
          "institución",
          "🏛 Extras por institución",
          "#3b82f6",
        );
      if (hasProg)
        buildChartBarFilter(
          extras,
          "programa",
          "🎓 Extras por programa",
          "#3b82f6",
        );
      buildTabla(extras);
      break;
    }

    // ── NO CIDI ──────────────────────────────────────────────
    case "nocidi": {
      if (!nNoCidi) {
        showEmptyFilter("🟣 No CIDI");
        break;
      }
      buildViewHeader(
        "🟣 No CIDI",
        nNoCidi,
        nTotal,
        "#8b5cf6",
        `${nNoCidi} bebé(s) fuera del programa CIDI · incluidos en el total`,
      );
      if (hasDia)
        buildChartBarFilter(nocidi, "día", "📅 No CIDI por día", "#8b5cf6");
      if (hasInst)
        buildChartBarFilter(
          nocidi,
          "institución",
          "🏛 No CIDI por institución",
          "#8b5cf6",
        );
      if (hasProg)
        buildChartBarFilter(
          nocidi,
          "programa",
          "🎓 No CIDI por programa",
          "#8b5cf6",
        );
      buildTabla(nocidi);
      break;
    }
  }
}

/** Limpia todos los charts activos y vacía el mapa */
function clearCharts() {
  for (const id of Object.keys(charts)) {
    if (charts[id]?.destroy) charts[id].destroy();
  }
  charts = {}; // resetear el objeto, no solo vaciar las entradas
}

/** Encabezado de vista filtrada con conteo, porcentaje y barra */
function buildViewHeader(label, count, total, color, subtitle) {
  const pct = total > 0 ? Math.round((count / total) * 100) : 0;
  const card = document.createElement("div");
  card.className = "card col-12";
  card.style.cssText = `border-left: 4px solid ${color}; margin-bottom: 4px;`;
  card.innerHTML = `
    <div style="padding:20px 28px;display:flex;align-items:center;gap:28px;flex-wrap:wrap;">
      <div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:2px;color:var(--muted);text-transform:uppercase;margin-bottom:4px">VISTA ACTIVA</div>
        <div style="font-size:22px;font-weight:800;color:${color};letter-spacing:-.5px">${label}</div>
        ${subtitle ? `<div style="font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted);margin-top:4px">${subtitle}</div>` : ""}
      </div>
      <div style="display:flex;gap:24px;flex-wrap:wrap;align-items:center;">
        <div style="text-align:center;">
          <div style="font-size:40px;font-weight:800;color:${color};line-height:1">${count}</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-top:4px">bebés únicos</div>
        </div>
        <div style="text-align:center;">
          <div style="font-size:40px;font-weight:800;color:var(--text-mid);line-height:1">${total}</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-top:4px">total únicos</div>
        </div>
        <div style="text-align:center;">
          <div style="font-size:40px;font-weight:800;color:${color};line-height:1">${pct}%</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-top:4px">de bebés únicos</div>
        </div>
      </div>
      <div style="flex:1;min-width:200px;">
        <div style="background:var(--surface2);border-radius:8px;height:14px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${color};border-radius:8px;transition:width .6s ease;"></div>
        </div>
      </div>
    </div>`;
  dashGrid.appendChild(card);
}

function showEmptyFilter(label) {
  dashGrid.innerHTML = `
    <div class="card col-12" style="padding:48px;text-align:center;">
      <div style="font-size:40px;margin-bottom:12px">📋</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:13px;color:var(--muted);letter-spacing:1px">${label} — sin datos</div>
      <div style="font-size:13px;color:var(--muted);margin-top:8px">No hay registros para esta categoría en los archivos cargados.</div>
    </div>`;
}

// =============================================================================
//  10. GRÁFICAS GENERALES
// =============================================================================

/** Bloque colapsable genérico */
function buildCollapsible(id, title, renderFn) {
  const wrap = document.createElement("div");
  wrap.className = "card col-12 collapsible-card";
  wrap.innerHTML = `
    <button class="collapsible-trigger">
      <span class="collapsible-title">${title}</span>
      <span class="collapsible-arrow">▶</span>
    </button>
    <div class="collapsible-body" id="collapsible-body-${id}"></div>`;
  dashGrid.appendChild(wrap);

  wrap.querySelector(".collapsible-trigger").addEventListener("click", () => {
    const body = document.getElementById(`collapsible-body-${id}`);
    const arrow = wrap.querySelector(".collapsible-arrow");
    if (body.classList.contains("open")) {
      body.classList.remove("open");
      arrow.textContent = "▶";
      const cid = `chart-${id}`;
      if (charts[cid]) {
        charts[cid].destroy();
        delete charts[cid];
      }
    } else {
      body.classList.add("open");
      arrow.textContent = "▼";
      body.innerHTML = "";
      renderFn();
    }
  });
}

/** Dona: Distribución de asistencia */
function buildChartDistribucion(data) {
  const presentes = data.filter((r) => asistio(r)).length;
  const ausentes = data.length - presentes;
  if (!presentes && !ausentes) return;
  const card = createCard(
    "chart-dona",
    "🥧 Distribución de asistencia",
    "col-4",
    "",
    280,
  );
  card.querySelector(".chart-wrap").innerHTML = `
    <canvas id="chart-dona"></canvas>
    <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;pointer-events:none">
      <div style="font-size:26px;font-weight:800;color:var(--text)">${data.length}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1.5px;color:var(--muted)">bebés</div>
    </div>`;
  dashGrid.appendChild(card);
  charts["chart-dona"] = new Chart(
    document.getElementById("chart-dona").getContext("2d"),
    {
      type: "doughnut",
      data: {
        labels: ["Presentes", "Ausentes"],
        datasets: [
          {
            data: [presentes, ausentes],
            backgroundColor: [VERDE + "dd", ROJO + "dd"],
            borderColor: [VERDE, ROJO],
            borderWidth: 2,
            hoverOffset: 8,
          },
        ],
      },
      options: {
        ...chartOpts(true),
        cutout: "68%",
        plugins: {
          ...chartOpts(true).plugins,
          tooltip: {
            ...chartOpts(true).plugins.tooltip,
            callbacks: {
              label: (ctx) => {
                const pct = Math.round((ctx.parsed / data.length) * 100);
                return `  ${ctx.label}: ${ctx.parsed} bebés (${pct}%)`;
              },
            },
          },
        },
      },
    },
  );
}

/** Dona: Reportes de ausencia sobre el total de bebés */
function buildChartTasaReporte(data) {
  // data = registros individuales (bebé × día)
  // Contamos TODOS los registros con reporte=Sí, sin importar si el bebé se repite.
  // Ej: si un bebé reportó el lunes y el miércoles → cuenta 2 reportes.
  const conReporte = data.filter((r) => esSi(val(r, "reporte"))).length;
  const sinReporte = data.length - conReporte;
  if (!data.length) return;
  const card = createCard(
    "chart-reporte",
    "📋 Reportes de ausencia",
    "col-4",
    "",
    280,
  );
  card.querySelector(".chart-wrap").innerHTML = `
    <canvas id="chart-reporte"></canvas>
    <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;pointer-events:none">
      <div style="font-size:26px;font-weight:800;color:var(--naranja)">${conReporte}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1.5px;color:var(--muted)">reportes totales</div>
    </div>`;
  dashGrid.appendChild(card);
  charts["chart-reporte"] = new Chart(
    document.getElementById("chart-reporte").getContext("2d"),
    {
      type: "doughnut",
      data: {
        labels: ["Con reporte", "Sin reporte"],
        datasets: [
          {
            data: [conReporte, sinReporte],
            backgroundColor: [NARANJA + "dd", "#e5e7eb"],
            borderColor: [NARANJA, "#d1d5db"],
            borderWidth: 2,
            hoverOffset: 8,
          },
        ],
      },
      options: {
        ...chartOpts(true),
        cutout: "68%",
        plugins: {
          ...chartOpts(true).plugins,
          tooltip: {
            ...chartOpts(true).plugins.tooltip,
            callbacks: {
              label: (ctx) =>
                `  ${ctx.label}: ${ctx.parsed} (${Math.round((ctx.parsed / data.length) * 100)}%)`,
            },
          },
        },
      },
    },
  );
}

/** Dona: Distribución por ubicación (sobre registros individuales, solo valores no vacíos) */
function buildChartUbicacion(data) {
  // data = registros individuales (bebé × día) — NO el resumen
  // La ubicación se registra cuando un bebé ausente informa desde dónde está.
  // Solo contamos registros con ubicación explícita (juanfe, casa, otro conocido).
  const cUbic = colKey("ubicacion");
  if (!cUbic) return;
  const UBIC = {
    Juanfe: { color: VERDE, count: 0 },
    Casa: { color: AZUL, count: 0 },
    Otro: { color: NARANJA, count: 0 },
  };
  // Valores conocidos — cualquier otra cosa distinta de vacío va a "Otro"
  const CONOCIDOS = ["juanfe", "casa"];
  data.forEach((r) => {
    const raw = norm(String(r[cUbic] || "").trim());
    if (!raw) return; // ignorar vacíos
    if (raw === "juanfe") UBIC.Juanfe.count++;
    else if (raw === "casa") UBIC.Casa.count++;
    else UBIC.Otro.count++;
  });
  const labels = Object.keys(UBIC).filter((l) => UBIC[l].count > 0);
  if (!labels.length) return;
  const total = labels.reduce((s, l) => s + UBIC[l].count, 0);
  const card = createCard(
    "chart-ubicacion",
    "📍 Asistencia por ubicación",
    "col-4",
    `${total} bebés`,
    280,
  );
  card.querySelector(".chart-wrap").innerHTML = `
    <canvas id="chart-ubicacion"></canvas>
    <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;pointer-events:none">
      <div style="font-size:24px;font-weight:800;color:var(--text)">${total}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:var(--muted)">bebés</div>
    </div>`;
  dashGrid.appendChild(card);
  charts["chart-ubicacion"] = new Chart(
    document.getElementById("chart-ubicacion").getContext("2d"),
    {
      type: "doughnut",
      data: {
        labels,
        datasets: [
          {
            data: labels.map((l) => UBIC[l].count),
            backgroundColor: labels.map((l) => UBIC[l].color + "dd"),
            borderColor: labels.map((l) => UBIC[l].color),
            borderWidth: 2,
            hoverOffset: 8,
          },
        ],
      },
      options: { ...chartOpts(true), cutout: "65%" },
    },
  );
}

/** Barras: Asistencia diaria a lo largo de la semana */
function buildChartAsistenciaPorDia(data) {
  const DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];
  const byDia = {};
  DIAS.forEach((d) => (byDia[d] = { total: 0, presentes: 0, reportes: 0 }));
  data.forEach((r) => {
    const d = val(r, "dia");
    if (!byDia[d]) return;
    byDia[d].total++;
    if (esSi(val(r, "asistencia"))) byDia[d].presentes++;
    if (esSi(val(r, "reporte"))) byDia[d].reportes++;
  });
  const labels = DIAS.filter((d) => byDia[d].total > 0);
  if (!labels.length) return;
  const card = createCard(
    "chart-dias",
    "📅 Asistencia por día",
    "col-12",
    `${data.length} registros`,
    300,
  );
  dashGrid.appendChild(card);
  charts["chart-dias"] = new Chart(
    document.getElementById("chart-dias").getContext("2d"),
    {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Presentes",
            data: labels.map((d) => byDia[d].presentes),
            backgroundColor: VERDE + "cc",
            borderRadius: 6,
            borderSkipped: false,
          },
          {
            label: "Ausentes",
            data: labels.map((d) => byDia[d].total - byDia[d].presentes),
            backgroundColor: ROJO + "cc",
            borderRadius: 6,
            borderSkipped: false,
          },
          {
            label: "Con reporte",
            data: labels.map((d) => byDia[d].reportes),
            backgroundColor: NARANJA + "cc",
            borderRadius: 6,
            borderSkipped: false,
          },
        ],
      },
      options: chartOpts(false),
    },
  );
}

/** Barras horizontales: Asistencia por institución */
function buildChartInstitucion(data) {
  const cInst = colKey("institucion");
  if (!cInst) return;
  const INST_ORDER = ["UTE", "ULA 1", "ULA 2", "TSF", "Otra"];
  const INST_COLORS = {
    UTE: AZUL,
    "ULA 1": VERDE,
    "ULA 2": TEAL,
    TSF: MORADO,
    Otra: NARANJA,
  };
  function canonInst(raw) {
    const n = norm(raw);
    if (!n || n === "nan") return null;
    if (n === "ute") return "UTE";
    if (n === "ula 1" || n === "ula1") return "ULA 1";
    if (n === "ula 2" || n === "ula2") return "ULA 2";
    if (n === "tsf") return "TSF";
    return "Otra";
  }
  const counts = {};
  INST_ORDER.forEach((k) => (counts[k] = { presentes: 0, ausentes: 0 }));
  // data = registros individuales (bebé × día) → esSi() es la fuente correcta
  data.forEach((r) => {
    const c = canonInst(String(r[cInst] || ""));
    if (!c) return;
    esSi(val(r, "asistencia")) ? counts[c].presentes++ : counts[c].ausentes++;
  });
  const labels = INST_ORDER.filter(
    (l) => counts[l].presentes + counts[l].ausentes > 0,
  );
  if (!labels.length) return;
  const h = Math.max(240, labels.length * 60 + 60);
  const totalInst = labels.reduce(
    (s, l) => s + counts[l].presentes + counts[l].ausentes,
    0,
  );
  const card = createCard(
    "chart-institucion",
    "🏛 Por:institución",
    "col-6",
    `${labels.length} instituciones · ${totalInst} registros`,
    h,
  );
  dashGrid.appendChild(card);
  const optsInst = {
    ...chartOpts(false),
    indexAxis: "y",
    plugins: {
      ...chartOpts(false).plugins,
      tooltip: {
        ...chartOpts(false).plugins.tooltip,
        callbacks: {
          // indexAxis:"y" → el valor numérico está en parsed.x, no parsed.y
          label: (ctx) => `  ${ctx.dataset.label}: ${ctx.parsed.x}`,
        },
      },
    },
  };
  charts["chart-institucion"] = new Chart(
    document.getElementById("chart-institucion").getContext("2d"),
    {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Presentes",
            data: labels.map((l) => counts[l].presentes),
            backgroundColor: labels.map((l) => (INST_COLORS[l] || AZUL) + "cc"),
            borderColor: labels.map((l) => INST_COLORS[l] || AZUL),
            borderWidth: 1.5,
            borderRadius: 5,
          },
          {
            label: "Ausentes",
            data: labels.map((l) => counts[l].ausentes),
            backgroundColor: ROJO + "33",
            borderColor: ROJO + "99",
            borderWidth: 1,
            borderRadius: 5,
          },
        ],
      },
      options: optsInst,
    },
  );
}

/** Barras horizontales: Asistencia por programa */
function buildChartPrograma(data) {
  const cProg = colKey("programa");
  if (!cProg) return;
  const PROG_ORDER = [
    "Hotelería",
    "Cocina",
    "Belleza",
    "Auxiliar Administrativo",
    "Otro",
  ];
  const PROG_COLORS = {
    Hotelería: AZUL,
    Cocina: NARANJA,
    Belleza: "#ec4899",
    "Auxiliar Administrativo": TEAL,
    Otro: MORADO,
  };
  function canonProg(raw) {
    const n = norm(raw);
    // Ignorar vacíos y "nan" (valor nulo que SheetJS convierte a string "nan")
    if (!n || n === "nan") return null;
    if (n === "hoteleria" || n === "hoteleria" || n.includes("hotel"))
      return "Hotelería";
    if (n === "cocina") return "Cocina";
    if (n === "belleza") return "Belleza";
    if (n.includes("auxiliar") || n.includes("administrativo"))
      return "Auxiliar Administrativo";
    // Solo clasificar como "Otro" si hay un valor real no reconocido
    return "Otro";
  }
  const counts = {};
  PROG_ORDER.forEach((k) => (counts[k] = { presentes: 0, ausentes: 0 }));
  // data = registros individuales → esSi() es la fuente correcta
  data.forEach((r) => {
    const c = canonProg(String(r[cProg] || ""));
    if (!c) return;
    esSi(val(r, "asistencia")) ? counts[c].presentes++ : counts[c].ausentes++;
  });
  const labels = PROG_ORDER.filter(
    (l) => counts[l].presentes + counts[l].ausentes > 0,
  );
  if (!labels.length) return;
  const h = Math.max(240, labels.length * 60 + 60);
  const totalProg = labels.reduce(
    (s, l) => s + counts[l].presentes + counts[l].ausentes,
    0,
  );
  const card = createCard(
    "chart-programa",
    "🎓 Por programa",
    "col-6",
    `${labels.length} programas · ${totalProg} registros`,
    h,
  );
  dashGrid.appendChild(card);
  const optsProg = {
    ...chartOpts(false),
    indexAxis: "y",
    plugins: {
      ...chartOpts(false).plugins,
      tooltip: {
        ...chartOpts(false).plugins.tooltip,
        callbacks: {
          label: (ctx) => `  ${ctx.dataset.label}: ${ctx.parsed.x}`,
        },
      },
    },
  };
  charts["chart-programa"] = new Chart(
    document.getElementById("chart-programa").getContext("2d"),
    {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Presentes",
            data: labels.map((l) => counts[l].presentes),
            backgroundColor: labels.map(
              (l) => (PROG_COLORS[l] || MORADO) + "cc",
            ),
            borderColor: labels.map((l) => PROG_COLORS[l] || MORADO),
            borderWidth: 1.5,
            borderRadius: 5,
          },
          {
            label: "Ausentes",
            data: labels.map((l) => counts[l].ausentes),
            backgroundColor: ROJO + "33",
            borderColor: ROJO + "99",
            borderWidth: 1,
            borderRadius: 5,
          },
        ],
      },
      options: optsProg,
    },
  );
}

/**
 * Barras: Distribución por rango de edad.
 * Maneja tanto valores numéricos ("6","16") como los rangos normalizados ("6-15","16-30").
 */
function buildChartEdad(data) {
  const cEdad = colKey("edad");
  if (!cEdad) return;

  const BUCKETS = {
    "6 – 15 meses": { presentes: 0, ausentes: 0, color: AZUL },
    "16 – 30 meses": { presentes: 0, ausentes: 0, color: VERDE },
    "Otro rango": { presentes: 0, ausentes: 0, color: NARANJA },
  };
  let sinEdad = 0;

  data.forEach((row) => {
    const raw = String(row[cEdad] || "").trim();
    let bucket = null;

    const rawNorm = raw
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
    if (rawNorm === "6-15" || /6.{1,5}15/.test(rawNorm))
      bucket = "6 – 15 meses";
    else if (rawNorm === "16-30" || /16.{1,5}30/.test(rawNorm))
      bucket = "16 – 30 meses";
    else {
      const n = parseFloat(rawNorm.replace(",", ".").replace(/[^\d.]/g, ""));
      if (!isNaN(n) && n > 0) {
        if (n >= 6 && n <= 15) bucket = "6 – 15 meses";
        else if (n >= 16 && n <= 30) bucket = "16 – 30 meses";
        else bucket = "Otro rango";
      } else {
        sinEdad++;
        return;
      }
    }

    asistio(row) ? BUCKETS[bucket].presentes++ : BUCKETS[bucket].ausentes++;
  });

  const labels = Object.keys(BUCKETS).filter(
    (l) => BUCKETS[l].presentes + BUCKETS[l].ausentes > 0,
  );
  if (!labels.length) return;
  const total = labels.reduce(
    (s, l) => s + BUCKETS[l].presentes + BUCKETS[l].ausentes,
    0,
  );
  const card = createCard(
    "chart-edad",
    "🍼 Por rango de edad",
    "col-12",
    `${total} bebés${sinEdad ? ` · ${sinEdad} sin dato` : ""}`,
    300,
  );
  dashGrid.appendChild(card);
  charts["chart-edad"] = new Chart(
    document.getElementById("chart-edad").getContext("2d"),
    {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Presentes",
            data: labels.map((l) => BUCKETS[l].presentes),
            backgroundColor: labels.map((l) => BUCKETS[l].color + "cc"),
            borderColor: labels.map((l) => BUCKETS[l].color),
            borderWidth: 1.5,
            borderRadius: 8,
            borderSkipped: false,
          },
          {
            label: "Ausentes",
            data: labels.map((l) => BUCKETS[l].ausentes),
            backgroundColor: ROJO + "44",
            borderColor: ROJO + "99",
            borderWidth: 1,
            borderRadius: 8,
            borderSkipped: false,
          },
        ],
      },
      options: chartOpts(false),
    },
  );
}

// Tabla de causas de ausencia (canonicaliza los valores para agrupar correctamente)
const SIT_CANON = [
  { key: "SANOS (CIDI)", patterns: ["sanos", "cidi"], color: VERDE },
  { key: "IRA", patterns: ["ira", "gripe", "viral", "resfr"], color: AZUL },
  {
    key: "ALERGIAS",
    patterns: ["alergia", "respiratoria", "piel", "medicamento"],
    color: TEAL,
  },
  { key: "BROTES", patterns: ["brote", "escabiosis", "contagio"], color: ROJO },
  { key: "EDA", patterns: ["eda", "diarrea"], color: NARANJA },
  { key: "VÓMITOS", patterns: ["vomito", "vómito"], color: MORADO },
  { key: "FIEBRE", patterns: ["fiebre"], color: "#f59e0b" },
  { key: "ACCIDENTE CASERO", patterns: ["accidente"], color: "#ef4444" },
  { key: "SITUACIÓN PERSONAL", patterns: ["personal"], color: "#8b5cf6" },
  { key: "ASISTE A FAMI", patterns: ["fami"], color: "#06b6d4" },
  { key: "CITA / VACUNAS", patterns: ["cita", "vacuna"], color: "#14b8a6" },
  { key: "HOSPITALIZACIÓN", patterns: ["hospital"], color: "#dc2626" },
  {
    key: "OTROS",
    patterns: ["otro", "transporte", "mama enferma", "mamá enferma"],
    color: "#94a3b8",
  },
];

function canonizarSituacion(raw) {
  const n = norm(raw);
  for (const c of SIT_CANON) {
    if (c.patterns.some((p) => n.includes(p))) return c.key;
  }
  return raw.trim() ? "OTROS" : null;
}

/** Causas de ausencia en un colapsable (vista Todos) */
function buildChartCausas(data) {
  const cSit = colKey("situacion");
  const body = document.getElementById("collapsible-body-causas");
  if (!cSit || !body) return;

  const counts = {};
  data
    .filter((r) => esSi(val(r, "reporte")))
    .forEach((r) => {
      const canon = canonizarSituacion(String(r[cSit] || "").trim()) || "OTROS";
      counts[canon] = (counts[canon] || 0) + 1;
    });
  const entries = SIT_CANON.filter((c) => counts[c.key] > 0).sort(
    (a, b) => counts[b.key] - counts[a.key],
  );
  if (!entries.length) {
    body.innerHTML =
      '<p style="padding:16px;color:var(--muted)">Sin reportes en este período.</p>';
    return;
  }
  const h = Math.max(320, entries.length * 52 + 60);
  body.innerHTML = `<div class="chart-wrap" style="height:${h}px"><canvas id="chart-causas"></canvas></div>`;
  charts["chart-causas"] = new Chart(
    document.getElementById("chart-causas").getContext("2d"),
    {
      type: "bar",
      data: {
        labels: entries.map((c) => c.key),
        datasets: [
          {
            data: entries.map((c) => counts[c.key]),
            backgroundColor: entries.map((c) => c.color + "cc"),
            borderColor: entries.map((c) => c.color),
            borderWidth: 1.5,
            borderRadius: 5,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        indexAxis: "y",
        plugins: {
          ...chartOpts(false).plugins,
          legend: { display: false },
          tooltip: {
            ...chartOpts(false).plugins.tooltip,
            callbacks: { label: (ctx) => `  ${ctx.label}: ${ctx.parsed.x}` },
          },
        },
        scales: {
          ...chartOpts(false).scales,
          x: {
            ...chartOpts(false).scales.x,
            ticks: { ...chartOpts(false).scales.x.ticks, stepSize: 1 },
          },
        },
      },
    },
  );
}

/** Causas de ausencia directo (sin colapsable, vista Reportados) */
function buildChartCausasDirecto(data) {
  const cSit = colKey("situacion");
  if (!cSit) return;
  // total = TODOS los registros con reporte, incluidos extras y no-cidi
  const totalReportes = data.filter((r) => esSi(val(r, "reporte"))).length;
  const counts = {};
  data
    .filter((r) => esSi(val(r, "reporte")))
    .forEach((r) => {
      // || "OTROS" garantiza que ningun reporte se pierda del conteo
      const canon = canonizarSituacion(String(r[cSit] || "").trim()) || "OTROS";
      counts[canon] = (counts[canon] || 0) + 1;
    });
  const entries = SIT_CANON.filter((c) => counts[c.key] > 0).sort(
    (a, b) => counts[b.key] - counts[a.key],
  );
  if (!entries.length) return;
  // Badge muestra el total real, no solo los que matchearon un patrón
  const total = totalReportes;
  const h = Math.max(280, entries.length * 52 + 60);
  const card = createCard(
    "chart-causas-dir",
    "⚠️ Causas de ausencia",
    "col-12",
    `${total} reportes`,
    h,
  );
  dashGrid.appendChild(card);
  charts["chart-causas-dir"] = new Chart(
    document.getElementById("chart-causas-dir").getContext("2d"),
    {
      type: "bar",
      data: {
        labels: entries.map((c) => c.key),
        datasets: [
          {
            data: entries.map((c) => counts[c.key]),
            backgroundColor: entries.map((c) => c.color + "cc"),
            borderColor: entries.map((c) => c.color),
            borderWidth: 1.5,
            borderRadius: 5,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        indexAxis: "y",
        plugins: {
          ...chartOpts(false).plugins,
          legend: { display: false },
          tooltip: {
            ...chartOpts(false).plugins.tooltip,
            callbacks: { label: (ctx) => `  ${ctx.label}: ${ctx.parsed.x}` },
          },
        },
        scales: {
          ...chartOpts(false).scales,
          x: {
            ...chartOpts(false).scales.x,
            ticks: { ...chartOpts(false).scales.x.ticks, stepSize: 1 },
          },
        },
      },
    },
  );
}

/** Gráfico de barras horizontal genérico para vistas filtradas */
function buildChartBarFilter(data, campo, titulo, color) {
  const campoKey =
    {
      día: "dia",
      institución: "institucion",
      programa: "programa",
      edad: "edad",
      ubicacion: "ubicacion",
    }[campo] || campo;
  const cCol = colKey(campoKey);
  if (!cCol) return;
  const counts = {};
  data.forEach((r) => {
    const v = String(r[cCol] || "").trim() || "(Sin dato)";
    counts[v] = (counts[v] || 0) + 1;
  });
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return;
  const total = entries.reduce((s, [, v]) => s + v, 0);
  const labels = entries.map(([k]) => k);
  const values = entries.map(([, v]) => v);
  const h = Math.max(200, labels.length * 52 + 60);
  const card = createCard(
    `chart-filter-${campoKey}`,
    titulo,
    "col-6",
    `${total} registros`,
    h,
  );
  dashGrid.appendChild(card);
  charts[`chart-filter-${campoKey}`] = new Chart(
    document.getElementById(`chart-filter-${campoKey}`).getContext("2d"),
    {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: titulo,
            data: values,
            backgroundColor: labels.map(
              (_, i) => color + (i === 0 ? "ee" : "99"),
            ),
            borderColor: color,
            borderWidth: 1.5,
            borderRadius: 6,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        indexAxis: "y",
        plugins: {
          ...chartOpts(false).plugins,
          legend: { display: false },
          tooltip: {
            ...chartOpts(false).plugins.tooltip,
            callbacks: {
              label: (ctx) => {
                const v = ctx.parsed.x;
                return `  ${v} registros (${total ? Math.round((v / total) * 100) : 0}%)`;
              },
            },
          },
        },
        scales: {
          ...chartOpts(false).scales,
          x: {
            ...chartOpts(false).scales.x,
            ticks: { ...chartOpts(false).scales.x.ticks, stepSize: 1 },
          },
        },
      },
    },
  );
}

// =============================================================================
//  11. HISTORIAL INDIVIDUAL DE BEBÉ
// =============================================================================

function buildBabyHistorial(data) {
  clearCharts();
  dashGrid.innerHTML = "";

  if (!data.length) {
    dashGrid.innerHTML = `<div class="card col-12" style="padding:48px;text-align:center;color:var(--muted);">Sin registros para este bebé en los archivos cargados.</div>`;
    return;
  }

  const madre = val(data[0], "madre");
  const inst = val(data[0], "institucion");
  const prog = val(data[0], "programa");
  const rawEdad = val(data[0], "edad");
  const ubic = val(data[0], "ubicacion");

  const totalReg = data.length;
  const totalSi = data.filter((r) => esSi(val(r, "asistencia"))).length;
  const totalNo = totalReg - totalSi;
  const totalRep = data.filter((r) => esSi(val(r, "reporte"))).length;
  const pct = totalReg ? Math.round((totalSi / totalReg) * 100) : 0;

  let edadLabel = rawEdad || null;
  if (rawEdad === "6-15") edadLabel = "6-15 meses";
  else if (rawEdad === "16-30") edadLabel = "16-30 meses";

  const CHIP = (icon, txt, bg, color) =>
    `<span style="background:${bg};color:${color};padding:4px 12px;border-radius:20px;font-weight:700;font-size:13px;display:inline-flex;align-items:center;gap:5px">${icon} ${txt}</span>`;

  const notas = data.map((r) => val(r, "nota")).filter((n) => n && n !== "—");

  const infoCard = document.createElement("div");
  infoCard.className = "card col-12";
  infoCard.innerHTML = `
    <div style="padding:22px 26px 18px;display:flex;flex-wrap:wrap;gap:20px;justify-content:space-between;align-items:flex-start">
      <div style="flex:1;min-width:260px">
        <div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:2.5px;color:var(--muted);margin-bottom:8px">👶 PERFIL DEL BEBÉ</div>
        <div style="font-size:24px;font-weight:800;color:var(--text);margin-bottom:12px">${selectedBaby}</div>
        <div style="display:flex;flex-wrap:wrap;gap:7px">
          ${madre ? CHIP("👩", madre, "var(--green-soft)", "var(--green-dark)") : ""}
          ${inst ? CHIP("🏛", inst, "var(--surface2)", "var(--text-mid)") : ""}
          ${prog ? CHIP("🎓", prog, "var(--azul-pale)", "var(--azul)") : ""}
          ${edadLabel ? CHIP("🍼", edadLabel, "#fdf4ff", "#7c3aed") : ""}
          ${ubic ? CHIP("📍", ubic, "var(--naranja-pale)", "var(--naranja)") : ""}
        </div>
      </div>
      <div style="display:flex;gap:16px;flex-wrap:wrap;text-align:center">
        ${[
          [
            totalSi,
            "Asistió",
            "var(--green-soft)",
            "var(--green)",
            "var(--green-dark)",
          ],
          [totalNo, "Faltó", "var(--rojo-pale)", "var(--rojo)", "var(--rojo)"],
          [pct + "%", "Asistencia", "#f0f9ff", "var(--green)", "var(--green)"],
          [
            totalRep,
            "Reportes",
            "var(--naranja-pale)",
            "var(--naranja)",
            "var(--naranja)",
          ],
          [
            totalReg,
            "Registros",
            "var(--surface2)",
            "var(--border)",
            "var(--text-mid)",
          ],
        ]
          .map(
            ([v, l, bg, border, txtColor]) => `
          <div style="background:${bg};border:1.5px solid ${border};border-radius:12px;padding:14px 18px;min-width:72px">
            <div style="font-size:26px;font-weight:800;color:${txtColor};line-height:1">${v}</div>
            <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:var(--muted);margin-top:4px;text-transform:uppercase">${l}</div>
          </div>`,
          )
          .join("")}
      </div>
    </div>
    ${
      notas.length
        ? `<div style="margin:0 26px 20px;padding:14px 18px;background:var(--green-soft);border-left:3px solid var(--green);border-radius:0 8px 8px 0">
      <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:var(--green-dark);margin-bottom:8px">📝 NOTAS</div>
      ${notas.map((n) => `<div style="font-size:13px;color:var(--text);padding:3px 0;border-bottom:1px solid var(--border)">${n}</div>`).join("")}
    </div>`
        : ""
    }`;
  dashGrid.appendChild(infoCard);

  // Gráfico de barras — historial de asistencia día a día
  const sorted = [...data].sort((a, b) => {
    const fa = val(a, "fecha") || val(a, "dia") || "";
    const fb = val(b, "fecha") || val(b, "dia") || "";
    return fa.localeCompare(fb);
  });
  const labelsHist = sorted.map(
    (r, i) => val(r, "dia") || val(r, "fecha") || `Reg ${i + 1}`,
  );
  const valsHist = sorted.map((r) => (esSi(val(r, "asistencia")) ? 1 : 0));
  const hHist = Math.min(Math.max(220, labelsHist.length * 22 + 80), 340);
  const cardHist = createCard(
    "chart-baby-linea",
    "📈 Historial de asistencia",
    "col-8",
    `${data.length} registros`,
    hHist,
  );
  dashGrid.appendChild(cardHist);
  charts["chart-baby-linea"] = new Chart(
    document.getElementById("chart-baby-linea").getContext("2d"),
    {
      type: "bar",
      data: {
        labels: labelsHist,
        datasets: [
          {
            label: "Asistencia",
            data: valsHist,
            backgroundColor: valsHist.map((v) =>
              v ? VERDE + "cc" : ROJO + "cc",
            ),
            borderRadius: 6,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        plugins: { ...chartOpts(false).plugins, legend: { display: false } },
        scales: {
          ...chartOpts(false).scales,
          y: {
            ...chartOpts(false).scales.y,
            min: 0,
            max: 1,
            ticks: {
              ...chartOpts(false).scales.y.ticks,
              callback: (v) => (v === 1 ? "Sí" : v === 0 ? "No" : ""),
            },
          },
        },
      },
    },
  );

  // Donas de resumen
  const cardDona1 = document.createElement("div");
  cardDona1.className = "card col-4";
  const pctColor = pct >= 80 ? VERDE : pct >= 60 ? NARANJA : ROJO;
  cardDona1.innerHTML = `
    <div class="card-head"><div class="card-title"><span class="card-dot"></span>📊 Tasa de asistencia</div></div>
    <div style="padding:24px 26px;text-align:center">
      <div style="font-size:52px;font-weight:800;color:${pctColor};line-height:1">${pct}%</div>
      <div style="margin:16px 0 8px;background:var(--surface2);border-radius:8px;height:10px;overflow:hidden">
        <div style="height:100%;width:${pct}%;background:${pctColor};border-radius:8px;transition:width .6s ease"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted)">
        <span style="color:var(--green-dark);font-weight:700">${totalSi} asistió</span>
        <span style="color:var(--rojo);font-weight:700">${totalNo} faltó</span>
      </div>
    </div>`;
  dashGrid.appendChild(cardDona1);

  // Causas de ausencia (colapsable)
  if (colKey("situacion")) {
    buildCollapsible("baby-causas", "⚠️ Causas de ausencia", () => {
      const body = document.getElementById("collapsible-body-baby-causas");
      const reportados = data.filter((r) => esSi(val(r, "reporte")));
      if (!reportados.length) {
        body.innerHTML =
          '<p style="padding:16px;color:var(--muted)">Sin reportes.</p>';
        return;
      }
      const counts = {};
      reportados.forEach((r) => {
        const s = val(r, "situacion") || "Sin especificar";
        counts[s] = (counts[s] || 0) + 1;
      });
      const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      body.innerHTML = `<div class="chart-wrap" style="height:280px"><canvas id="chart-baby-causas"></canvas></div>`;
      charts["chart-baby-causas"] = new Chart(
        document.getElementById("chart-baby-causas").getContext("2d"),
        {
          type: "doughnut",
          data: {
            labels: entries.map(([k]) => k),
            datasets: [
              {
                data: entries.map(([, v]) => v),
                backgroundColor: entries.map(
                  (_, i) => PALETTE[i % PALETTE.length] + "dd",
                ),
                borderWidth: 2,
                hoverOffset: 6,
              },
            ],
          },
          options: { ...chartOpts(true), cutout: "55%" },
        },
      );
    });
  }

  buildTabla(data);
}

// =============================================================================
//  12. ETL AVANZADO — Tendencia · Ranking · Comparación · Alertas de riesgo
// =============================================================================

/**
 * buildETLSection(registros, resumen)
 * Sección de análisis avanzado. Solo se muestra en la vista "Todos".
 * NOTA: usa parámetros locales — no accede a la variable global allData.
 */
function buildETLSection(registros, resumen) {
  const hasFecha = registros.some((r) => val(r, "fecha"));
  const hasDia = registros.some((r) => val(r, "dia"));
  if (!hasFecha && !hasDia) return;

  const sep = document.createElement("div");
  sep.className = "card col-12";
  sep.style.cssText =
    "background:linear-gradient(90deg,rgba(133,218,26,.08) 0%,transparent 100%);border-left:4px solid var(--green);padding:18px 28px;margin:8px 0 4px;";
  sep.innerHTML = `
    <div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:3px;color:var(--green-dark);text-transform:uppercase;margin-bottom:4px">📊 Análisis ETL</div>
    <div style="font-size:15px;font-weight:700;color:var(--text)">Tendencia · Ranking · Comparación · Alertas de riesgo</div>`;
  dashGrid.appendChild(sep);

  buildETLTendencia(registros);
  buildETLToggle(
    "etl-ranking",
    "🏆",
    "Ranking",
    "Bebés con más ausencias",
    "#3b82f6",
    "#dbeafe",
    "#bfdbfe",
    (body) => buildETLRanking(registros, body),
  );
  buildETLToggle(
    "etl-comparacion",
    "📆",
    "Comparación por período",
    "Asistencia por fecha o día de semana",
    "#8b5cf6",
    "#ede9fe",
    "#ddd6fe",
    (body) => buildETLComparacion(registros, body),
  );
  buildETLAlertas(registros, resumen);
}

// ---------------------------------------------------------------------------
//  Utilidades de fecha para el ETL
// ---------------------------------------------------------------------------

/**
 * parseFechaRobusta(str) → "YYYY-MM-DD" | null
 * Soporta: DD/MM/YYYY · YYYY-MM-DD · DD-MM-YYYY · serial Excel · "15 de enero 2025" · ISO
 */
function parseFechaRobusta(str) {
  if (!str && str !== 0) return null;
  const s = String(str).trim();

  // Serial numérico de Excel (aprox. años 2009–2050)
  const serial = parseFloat(s);
  if (!isNaN(serial) && serial > 40000 && serial < 55000) {
    const d = new Date(Math.round((serial - 25569) * 86400000));
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  // YYYY-MM-DD
  let m = s.match(/^(\d{4})[\/\-\.](\d{2})[\/\-\.](\d{2})$/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  // DD/MM/YYYY o DD-MM-YYYY (Colombia: asumir día/mes)
  m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  // "15 de enero de 2025"
  const MESES = {
    enero: 1,
    febrero: 2,
    marzo: 3,
    abril: 4,
    mayo: 5,
    junio: 6,
    julio: 7,
    agosto: 8,
    septiembre: 9,
    octubre: 10,
    noviembre: 11,
    diciembre: 12,
  };
  m = s.toLowerCase().match(/(\d{1,2})\s+de\s+(\w+)(?:\s+de)?\s+(\d{4})/);
  if (m && MESES[m[2]])
    return `${m[3]}-${String(MESES[m[2]]).padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  // ISO timestamp
  m = s.match(/^(\d{4}-\d{2}-\d{2})T/);
  if (m) return m[1];
  return null;
}

/** Devuelve la semana ISO de una fecha "YYYY-MM-DD" → "2025-S15" */
function semanaISO(fechaStr) {
  const d = new Date(fechaStr + "T12:00:00");
  if (isNaN(d)) return null;
  const jan4 = new Date(d.getFullYear(), 0, 4);
  const ini = new Date(jan4);
  ini.setDate(jan4.getDate() - ((jan4.getDay() + 6) % 7));
  const semana = Math.floor(1 + (d - ini) / (7 * 86400000));
  return `${d.getFullYear()}-S${String(semana).padStart(2, "0")}`;
}

/**
 * buildETLToggle — header con botón Mostrar/Ocultar igual que Alertas de riesgo.
 * Renderiza el contenido bajo demanda (lazy) al abrir por primera vez.
 */
function buildETLToggle(
  id,
  icono,
  titulo,
  subtitulo,
  color,
  bgSoft,
  border,
  renderFn,
) {
  const card = document.createElement("div");
  card.className = "card col-12";
  card.style.cssText = `border:1px solid ${border};border-top:3px solid ${color};overflow:hidden;`;

  card.innerHTML = `
    <div id="${id}-header" style="display:flex;align-items:center;justify-content:space-between;padding:16px 22px;cursor:pointer;background:linear-gradient(135deg,${bgSoft} 0%,#fff 100%);border-bottom:1px solid ${border};user-select:none;">
      <div style="display:flex;align-items:center;gap:8px">
        <div style="width:32px;height:32px;border-radius:8px;background:${bgSoft};border:1.5px solid ${border};display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0">${icono}</div>
        <div>
          <div style="font-weight:800;font-size:14px;color:${color}">${titulo}</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#a0aec0;text-transform:uppercase">${subtitulo}</div>
        </div>
      </div>
      <button id="${id}-toggle-btn" style="display:flex;align-items:center;gap:7px;padding:7px 16px;border-radius:8px;cursor:pointer;border:1.5px solid ${border};background:white;color:${color};font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;white-space:nowrap"
        onmouseover="this.style.background='${bgSoft}'" onmouseout="this.style.background='white'">
        <span id="${id}-arrow" style="transition:transform .25s">▼</span>
        <span id="${id}-text">Mostrar</span>
      </button>
    </div>
    <div id="${id}-body" style="padding:20px 22px 16px;display:none;background:#fafafa;"></div>`;

  dashGrid.appendChild(card);

  let abierto = false;
  let rendered = false;
  const body = card.querySelector(`#${id}-body`);
  const arrow = card.querySelector(`#${id}-arrow`);
  const textBtn = card.querySelector(`#${id}-text`);

  card.querySelector(`#${id}-header`).addEventListener("click", () => {
    abierto = !abierto;
    body.style.display = abierto ? "block" : "none";
    arrow.style.transform = abierto ? "rotate(180deg)" : "rotate(0deg)";
    textBtn.textContent = abierto ? "Ocultar" : "Mostrar";

    // Renderizar contenido solo la primera vez (lazy)
    if (abierto && !rendered) {
      rendered = true;
      renderFn(body);
    }
  });
}

/** Tendencia de asistencia en el tiempo */
function buildETLTendencia(data) {
  const fechasRaw = data.map((r) => val(r, "fecha")).filter(Boolean);
  const diasRaw = data.map((r) => val(r, "dia")).filter(Boolean);
  const porPeriodo = {};

  if (fechasRaw.length > 0) {
    const fechasValidas = [
      ...new Set(fechasRaw.map(parseFechaRobusta).filter(Boolean)),
    ];
    const gran =
      fechasValidas.length <= 20
        ? "dia"
        : fechasValidas.length <= 90
          ? "semana"
          : "mes";
    data.forEach((r) => {
      const p = parseFechaRobusta(val(r, "fecha"));
      if (!p) return;
      const key =
        gran === "semana"
          ? semanaISO(p) || p
          : gran === "mes"
            ? p.slice(0, 7)
            : p;
      if (!porPeriodo[key]) porPeriodo[key] = { total: 0, presentes: 0 };
      porPeriodo[key].total++;
      if (esSi(val(r, "asistencia"))) porPeriodo[key].presentes++;
    });
  } else if (diasRaw.length > 0) {
    const DIAS_ORD = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];
    data.forEach((r) => {
      const d = val(r, "dia");
      if (!d) return;
      if (!porPeriodo[d]) porPeriodo[d] = { total: 0, presentes: 0 };
      porPeriodo[d].total++;
      if (esSi(val(r, "asistencia"))) porPeriodo[d].presentes++;
    });
    // Ordenar por día de semana
    const orden = [
      "lunes",
      "martes",
      "miercoles",
      "miércoles",
      "jueves",
      "viernes",
    ];
    Object.keys(porPeriodo)
      .sort((a, b) => {
        const ia = orden.indexOf(norm(a)),
          ib = orden.indexOf(norm(b));
        return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
      })
      .forEach((k) => {
        const v = porPeriodo[k];
        delete porPeriodo[k];
        porPeriodo[k] = v;
      });
  }

  const periodos = Object.keys(porPeriodo).sort((a, b) => {
    const DIAS = [
      "lunes",
      "martes",
      "miercoles",
      "miércoles",
      "jueves",
      "viernes",
    ];
    const ia = DIAS.indexOf(norm(a)),
      ib = DIAS.indexOf(norm(b));
    if (ia >= 0 && ib >= 0) return ia - ib;
    return a.localeCompare(b);
  });
  if (periodos.length < 2) return;

  const tasas = periodos.map((p) =>
    porPeriodo[p].total
      ? Math.round((porPeriodo[p].presentes / porPeriodo[p].total) * 100)
      : 0,
  );
  const totals = periodos.map((p) => porPeriodo[p].total);
  const h = Math.min(Math.max(260, periodos.length * 30 + 80), 360);
  const card = createCard(
    "chart-etl-tendencia",
    "📈 Tendencia de asistencia",
    "col-12",
    `${periodos.length} períodos`,
    h,
  );
  dashGrid.appendChild(card);
  charts["chart-etl-tendencia"] = new Chart(
    document.getElementById("chart-etl-tendencia").getContext("2d"),
    {
      type: "line",
      data: {
        labels: periodos,
        datasets: [
          {
            label: "% Asistencia",
            data: tasas,
            borderColor: VERDE,
            backgroundColor: VERDE + "22",
            borderWidth: 2.5,
            pointBackgroundColor: tasas.map((t) =>
              t >= 80 ? VERDE : t >= 60 ? NARANJA : ROJO,
            ),
            pointBorderColor: "#fff",
            pointRadius: 6,
            tension: 0.35,
            fill: true,
            yAxisID: "yPct",
          },
          {
            label: "Total registros",
            data: totals,
            borderColor: AZUL + "88",
            backgroundColor: "transparent",
            borderWidth: 1.5,
            borderDash: [5, 4],
            pointRadius: 3,
            tension: 0.3,
            yAxisID: "yTotal",
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 500 },
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: {
            labels: {
              color: "#8ba869",
              font: { family: "JetBrains Mono", size: 11 },
              padding: 14,
              boxWidth: 10,
            },
          },
          tooltip: {
            backgroundColor: "#fff",
            titleColor: "#1a2e05",
            bodyColor: "#555",
            borderColor: "#e0eccc",
            borderWidth: 1,
            padding: 12,
            callbacks: {
              label: (ctx) =>
                ctx.datasetIndex === 0
                  ? `  Asistencia: ${ctx.parsed.y}%`
                  : `  Total: ${ctx.parsed.y} registros`,
            },
          },
        },
        scales: {
          x: {
            ticks: {
              color: "#8ba869",
              font: { family: "JetBrains Mono", size: 10 },
              maxRotation: 35,
            },
            grid: { color: "rgba(133,218,26,.07)" },
            border: { color: "#e0eccc" },
          },
          yPct: {
            type: "linear",
            position: "left",
            min: 0,
            max: 100,
            ticks: {
              color: VERDE,
              font: { family: "JetBrains Mono", size: 10 },
              callback: (v) => v + "%",
            },
            grid: { color: "rgba(133,218,26,.07)" },
            border: { color: "#e0eccc" },
          },
          yTotal: {
            type: "linear",
            position: "right",
            beginAtZero: true,
            ticks: {
              color: AZUL + "99",
              font: { family: "JetBrains Mono", size: 10 },
            },
            grid: { drawOnChartArea: false },
            border: { color: "transparent" },
          },
        },
      },
    },
  );
}

/**
 * Ranking de bebés con más ausencias.
 * Agrupa por nombre+madre para distinguir homónimos correctamente.
 */
function buildETLRanking(data, body) {
  if (!body) return;

  // Clave = nombre + madre para distinguir homónimos
  const statsMap = new Map();
  data.forEach((r) => {
    const nombre = val(r, "bebe");
    if (!nombre) return;
    const key = nombre + "|" + (val(r, "madre") || "");
    if (!statsMap.has(key))
      statsMap.set(key, { nombre, total: 0, ausencias: 0 });
    const s = statsMap.get(key);
    s.total++;
    if (!esSi(val(r, "asistencia"))) s.ausencias++;
  });

  const ranking = Array.from(statsMap.values())
    .filter((s) => s.ausencias > 0)
    .map((s) => ({
      ...s,
      tasa: Math.round(((s.total - s.ausencias) / s.total) * 100),
    }))
    .sort((a, b) => b.ausencias - a.ausencias)
    .slice(0, 15);

  if (!ranking.length) {
    body.innerHTML =
      '<p style="padding:16px;color:var(--muted)">Sin ausencias.</p>';
    return;
  }
  const h = Math.max(300, ranking.length * 44 + 60);
  body.innerHTML = `<div class="chart-wrap" style="height:${h}px"><canvas id="chart-etl-ranking"></canvas></div>`;
  charts["chart-etl-ranking"] = new Chart(
    document.getElementById("chart-etl-ranking").getContext("2d"),
    {
      type: "bar",
      data: {
        labels: ranking.map((d) => d.nombre.split(" ").slice(0, 2).join(" ")),
        datasets: [
          {
            label: "Ausencias",
            data: ranking.map((d) => d.ausencias),
            backgroundColor: ranking.map((d) =>
              d.tasa < 60
                ? ROJO + "dd"
                : d.tasa < 80
                  ? NARANJA + "dd"
                  : VERDE + "dd",
            ),
            borderColor: ranking.map((d) =>
              d.tasa < 60 ? ROJO : d.tasa < 80 ? NARANJA : VERDE,
            ),
            borderWidth: 1.5,
            borderRadius: 5,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        indexAxis: "y",
        plugins: {
          ...chartOpts(false).plugins,
          legend: { display: false },
          tooltip: {
            ...chartOpts(false).plugins.tooltip,
            callbacks: {
              title: (ctx) => ranking[ctx[0].dataIndex]?.nombre || ctx[0].label,
              label: (ctx) => {
                const d = ranking[ctx.dataIndex];
                return [
                  `  Ausencias: ${d.ausencias} de ${d.total}`,
                  `  Tasa: ${d.tasa}%`,
                ];
              },
            },
          },
        },
        scales: {
          ...chartOpts(false).scales,
          x: {
            ...chartOpts(false).scales.x,
            ticks: { ...chartOpts(false).scales.x.ticks, stepSize: 1 },
          },
        },
      },
    },
  );
}

/** Comparación de asistencia por período (semana/mes/día) */
function buildETLComparacion(data, body) {
  if (!body) return;

  const fechasRaw = data.map((r) => val(r, "fecha")).filter(Boolean);
  const diasRaw = data.map((r) => val(r, "dia")).filter(Boolean);
  let grupos = {};

  if (fechasRaw.length > 0) {
    const fechasValidas = [
      ...new Set(fechasRaw.map(parseFechaRobusta).filter(Boolean)),
    ];
    const usarMes = fechasValidas.length >= 60;
    const usarSemana = !usarMes && fechasValidas.length >= 5;
    data.forEach((r) => {
      const p = parseFechaRobusta(val(r, "fecha"));
      if (!p) return;
      const key = usarMes
        ? p.slice(0, 7)
        : usarSemana
          ? semanaISO(p) || p.slice(0, 7)
          : p;
      if (!grupos[key]) grupos[key] = { total: 0, presentes: 0 };
      grupos[key].total++;
      if (esSi(val(r, "asistencia"))) grupos[key].presentes++;
    });
    grupos = Object.fromEntries(
      Object.entries(grupos).sort(([a], [b]) => a.localeCompare(b)),
    );
  } else if (diasRaw.length > 0) {
    const DIAS_ORD = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];
    data.forEach((r) => {
      const d = val(r, "dia");
      if (!d) return;
      if (!grupos[d]) grupos[d] = { total: 0, presentes: 0 };
      grupos[d].total++;
      if (esSi(val(r, "asistencia"))) grupos[d].presentes++;
    });
    const ord = {};
    DIAS_ORD.forEach((d) => {
      if (grupos[d]) ord[d] = grupos[d];
    });
    Object.keys(grupos).forEach((d) => {
      if (!ord[d]) ord[d] = grupos[d];
    });
    grupos = ord;
  }

  const periodos = Object.keys(grupos);
  if (periodos.length < 2) {
    body.innerHTML =
      '<p style="padding:16px;color:var(--muted)">Se necesitan al menos 2 períodos para comparar. Carga más archivos Excel.</p>';
    return;
  }

  const h = Math.max(280, periodos.length * 38 + 80);
  body.innerHTML = `<div class="chart-wrap" style="height:${h}px"><canvas id="chart-etl-comparacion"></canvas></div>`;
  const tasas = periodos.map((p) =>
    grupos[p].total
      ? Math.round((grupos[p].presentes / grupos[p].total) * 100)
      : 0,
  );
  charts["chart-etl-comparacion"] = new Chart(
    document.getElementById("chart-etl-comparacion").getContext("2d"),
    {
      type: "bar",
      data: {
        labels: periodos,
        datasets: [
          {
            label: "Presentes",
            data: periodos.map((p) => grupos[p].presentes),
            backgroundColor: VERDE + "cc",
            borderRadius: 5,
            borderSkipped: false,
          },
          {
            label: "Ausentes",
            data: periodos.map((p) => grupos[p].total - grupos[p].presentes),
            backgroundColor: ROJO + "99",
            borderRadius: 5,
            borderSkipped: false,
          },
        ],
      },
      options: {
        ...chartOpts(false),
        plugins: {
          ...chartOpts(false).plugins,
          tooltip: {
            ...chartOpts(false).plugins.tooltip,
            callbacks: {
              afterBody: (ctx) =>
                ctx.length ? [`  Tasa: ${tasas[ctx[0].dataIndex]}%`] : [],
            },
          },
        },
        scales: {
          ...chartOpts(false).scales,
          x: { ...chartOpts(false).scales.x, stacked: false },
          y: {
            ...chartOpts(false).scales.y,
            stacked: false,
            beginAtZero: true,
          },
        },
      },
    },
  );
}

/** Alertas de riesgo — bebés con baja asistencia (colapsable con toggle) */
function buildETLAlertas(data, resumen) {
  // Calcular stats por bebé (nombre+madre como clave)
  const statsMap = new Map();
  data.forEach((r) => {
    const nombre = val(r, "bebe");
    if (!nombre) return;
    const key = nombre + "|" + (val(r, "madre") || "");
    if (!statsMap.has(key))
      statsMap.set(key, {
        nombre,
        total: 0,
        presentes: 0,
        reportes: 0,
        madre: val(r, "madre"),
        inst: val(r, "institucion"),
        prog: val(r, "programa"),
      });
    const s = statsMap.get(key);
    s.total++;
    if (esSi(val(r, "asistencia"))) s.presentes++;
    if (esSi(val(r, "reporte"))) s.reportes++;
  });

  const enRiesgo = Array.from(statsMap.values())
    .map((s) => ({
      ...s,
      tasa: s.total ? Math.round((s.presentes / s.total) * 100) : 0,
    }))
    .filter((d) => d.tasa < 80 && d.total >= 2)
    .sort((a, b) => a.tasa - b.tasa);

  if (!enRiesgo.length) return;

  const NIVEL = {
    critico: {
      color: "#c0392b",
      bg: "#fdf2f2",
      border: "#f5c6c6",
      tagBg: "#e74c3c22",
      tagText: "#c0392b",
    },
    moderado: {
      color: "#c96a00",
      bg: "#fdf6ee",
      border: "#f8d9b0",
      tagBg: "#f0901022",
      tagText: "#c96a00",
    },
    bajo: {
      color: "#8a7500",
      bg: "#fdfbee",
      border: "#f0e68c",
      tagBg: "#d4ac0022",
      tagText: "#8a7500",
    },
  };
  const criticos = enRiesgo.filter((d) => d.tasa < 50);
  const moderados = enRiesgo.filter((d) => d.tasa >= 50 && d.tasa < 70);
  const bajos = enRiesgo.filter((d) => d.tasa >= 70 && d.tasa < 80);

  const renderFila = (d, n) => {
    const initials = d.nombre
      .split(" ")
      .slice(0, 2)
      .map((w) => w[0] || "")
      .join("")
      .toUpperCase();
    return `
    <div style="display:flex;align-items:center;gap:14px;padding:12px 16px;background:${n.bg};border:1px solid ${n.border};border-radius:12px;margin-bottom:8px;flex-wrap:wrap;"
      onmouseover="this.style.boxShadow='0 2px 12px ${n.color}22'" onmouseout="this.style.boxShadow=''">
      <div style="width:40px;height:40px;border-radius:50%;flex-shrink:0;background:${n.tagBg};border:2px solid ${n.border};display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:800;color:${n.tagText}">${initials}</div>
      <div style="flex:1;min-width:140px">
        <div style="font-weight:700;font-size:13px;color:#2d3748">${d.nombre}</div>
        <div style="font-size:10px;color:#718096;font-family:'JetBrains Mono',monospace;margin-top:2px">${[d.madre, d.inst, d.prog].filter(Boolean).join(" · ")}</div>
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
        ${[
          [d.tasa + "%", "ASIST.", n.color],
          [d.presentes + "/" + d.total, "DÍAS", "#4a5568"],
          [d.total - d.presentes, "AUSENC.", n.color],
          ...(d.reportes > 0 ? [[d.reportes, "REPORT.", "#c96a00"]] : []),
        ]
          .map(
            ([v, l, c]) =>
              `<div style="background:white;border:1.5px solid ${n.border};border-radius:10px;padding:7px 12px;text-align:center;min-width:58px"><div style="font-size:17px;font-weight:800;color:${c};line-height:1">${v}</div><div style="font-family:'JetBrains Mono',monospace;font-size:8px;color:#a0aec0;margin-top:2px">${l}</div></div>`,
          )
          .join("")}
        <div style="min-width:80px"><div style="background:#e2e8f0;border-radius:99px;height:6px;overflow:hidden;width:80px"><div style="height:100%;width:${d.tasa}%;background:${n.color};border-radius:99px"></div></div></div>
      </div>
    </div>`;
  };

  const renderGrupo = (lista, n, icono, titulo, desc) =>
    !lista.length
      ? ""
      : `
    <div style="margin-bottom:20px">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;padding:10px 14px;background:${n.bg};border-left:3px solid ${n.color};border-radius:0 8px 8px 0">
        <span style="font-size:18px">${icono}</span>
        <div>
          <div style="font-weight:800;font-size:13px;color:${n.color}">${titulo} <span style="background:${n.tagBg};color:${n.tagText};padding:2px 10px;border-radius:99px;font-size:11px">${lista.length}</span></div>
          <div style="font-size:10px;color:#718096;font-family:'JetBrains Mono',monospace">${desc}</div>
        </div>
      </div>
      ${lista.map((d) => renderFila(d, n)).join("")}
    </div>`;

  const chips = [
    criticos.length
      ? `<span style="background:#fdf2f2;color:#c0392b;border:1px solid #f5c6c6;padding:4px 12px;border-radius:99px;font-size:11px;font-weight:700">🔴 ${criticos.length} crítico(s)</span>`
      : "",
    moderados.length
      ? `<span style="background:#fdf6ee;color:#c96a00;border:1px solid #f8d9b0;padding:4px 12px;border-radius:99px;font-size:11px;font-weight:700">🟠 ${moderados.length} moderado(s)</span>`
      : "",
    bajos.length
      ? `<span style="background:#fdfbee;color:#8a7500;border:1px solid #f0e68c;padding:4px 12px;border-radius:99px;font-size:11px;font-weight:700">🟡 ${bajos.length} bajo(s)</span>`
      : "",
  ]
    .filter(Boolean)
    .join(" ");

  const card = document.createElement("div");
  card.className = "card col-12";
  card.style.cssText =
    "border:1px solid #fecaca;border-top:3px solid #e57373;overflow:hidden;";
  card.innerHTML = `
    <div id="alertas-header" style="display:flex;align-items:center;justify-content:space-between;padding:16px 22px;cursor:pointer;background:linear-gradient(135deg,#fff5f5 0%,#fff 100%);border-bottom:1px solid #fecaca;user-select:none;">
      <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <div style="display:flex;align-items:center;gap:8px">
          <div style="width:32px;height:32px;border-radius:8px;background:#fdf2f2;border:1.5px solid #fecaca;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0">🚨</div>
          <div>
            <div style="font-weight:800;font-size:14px;color:#c0392b">Alertas de riesgo</div>
            <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#a0aec0;text-transform:uppercase">Bebés con baja asistencia</div>
          </div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap">${chips}</div>
      </div>
      <button id="alertas-toggle-btn" style="display:flex;align-items:center;gap:7px;padding:7px 16px;border-radius:8px;cursor:pointer;border:1.5px solid #fecaca;background:white;color:#c0392b;font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;white-space:nowrap"
        onmouseover="this.style.background='#fdf2f2'" onmouseout="this.style.background='white'">
        <span id="alertas-arrow" style="transition:transform .25s">▼</span>
        <span id="alertas-text">Mostrar ${enRiesgo.length} bebés</span>
      </button>
    </div>
    <div id="alertas-body" style="padding:20px 22px 16px;display:none;background:#fafafa;">
      ${renderGrupo(criticos, NIVEL.critico, "🔴", "Crítico", "Asistencia menor al 50%")}
      ${renderGrupo(moderados, NIVEL.moderado, "🟠", "Moderado", "Asistencia entre 50% y 70%")}
      ${renderGrupo(bajos, NIVEL.bajo, "🟡", "Bajo", "Asistencia entre 70% y 80%")}
      <div style="margin-top:8px;padding:10px 14px;background:white;border:1px solid #e2e8f0;border-radius:8px;font-family:'JetBrains Mono',monospace;font-size:9px;color:#a0aec0;line-height:1.6">
        ℹ️ Se muestran bebés con <strong style="color:#718096">mínimo 2 registros</strong> y <strong style="color:#718096">asistencia menor al 80%</strong>. Porcentajes calculados sobre el total de días registrados de cada bebé.
      </div>
    </div>`;
  dashGrid.appendChild(card);

  let abierto = false;
  const body2 = card.querySelector("#alertas-body");
  const arrow = card.querySelector("#alertas-arrow");
  const textBtn = card.querySelector("#alertas-text");
  card.querySelector("#alertas-header").addEventListener("click", () => {
    abierto = !abierto;
    body2.style.display = abierto ? "block" : "none";
    arrow.style.transform = abierto ? "rotate(180deg)" : "rotate(0deg)";
    textBtn.textContent = abierto
      ? "Ocultar"
      : `Mostrar ${enRiesgo.length} bebés`;
  });
}

// =============================================================================
//  SECCIÓN DE REPORTES DETALLADA
//  Muestra todos los reportes agrupados por bebé.
//  El nombre del bebé aparece UNA sola vez como encabezado de grupo.
//  Cada reporte debajo muestra: día, fecha, situación y nota.
// =============================================================================

function buildSeccionReportes(data) {
  // data = registros individuales — filtrar solo los que tienen reporte=Sí
  const cBebe = colKey("bebe");
  const cMadre = colKey("madre");
  const cInst = colKey("institucion");
  const cProg = colKey("programa");
  const cDia = colKey("dia");
  const cFecha = colKey("fecha");
  const cSit = colKey("situacion");
  const cNota = colKey("nota");
  if (!cBebe) return;

  const reportados = data.filter((r) => esSi(val(r, "reporte")));
  if (!reportados.length) return;

  // Agrupar por bebé (nombre+madre) manteniendo orden de aparición
  const grupos = new Map();
  for (const r of reportados) {
    const k = norm(val(r, "bebe")) + "|" + norm(val(r, "madre") || "");
    if (!grupos.has(k)) {
      grupos.set(k, {
        nombre: r[cBebe] || "",
        madre: cMadre ? r[cMadre] || "" : "",
        inst: cInst ? r[cInst] || "" : "",
        prog: cProg ? r[cProg] || "" : "",
        reportes: [],
      });
    }
    grupos.get(k).reportes.push({
      dia: cDia ? r[cDia] || "" : "",
      fecha: cFecha ? r[cFecha] || "" : "",
      sit: cSit ? r[cSit] || "" : "",
      nota: cNota ? r[cNota] || "" : "",
    });
  }

  const totalBebes = grupos.size;
  const totalReportes = reportados.length;

  // Colores por situación — reutiliza SIT_CANON
  function colorSit(sit) {
    const n = norm(sit);
    const found = SIT_CANON.find((c) => c.patterns.some((p) => n.includes(p)));
    return found ? found.color : "#94a3b8";
  }

  // Construir HTML de la sección
  const card = document.createElement("div");
  card.className = "card col-12";
  card.style.cssText =
    "border:1px solid #e0eccc;border-top:3px solid var(--green);overflow:hidden;";

  // ── Header con toggle ──
  card.innerHTML = `
    <div id="reportes-header" style="display:flex;align-items:center;justify-content:space-between;padding:16px 22px;cursor:pointer;background:linear-gradient(135deg,#f0fad8 0%,#fff 100%);border-bottom:1px solid #e0eccc;user-select:none;">
      <div style="display:flex;align-items:center;gap:8px">
        <div style="width:32px;height:32px;border-radius:8px;background:#f0fad8;border:1.5px solid #d0e8a8;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0">📋</div>
        <div>
          <div style="font-weight:800;font-size:14px;color:var(--green-dark)">Detalle de reportes</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#a0aec0;text-transform:uppercase">
            ${totalBebes} bebés · ${totalReportes} reportes totales
          </div>
        </div>
      </div>
      <button id="reportes-toggle-btn"
        style="display:flex;align-items:center;gap:7px;padding:7px 16px;border-radius:8px;cursor:pointer;border:1.5px solid #d0e8a8;background:white;color:var(--green-dark);font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;white-space:nowrap"
        onmouseover="this.style.background='#f0fad8'" onmouseout="this.style.background='white'">
        <span id="reportes-arrow" style="transition:transform .25s">▼</span>
        <span id="reportes-text">Mostrar</span>
      </button>
    </div>
    <div id="reportes-body" style="display:none;padding:20px 22px 24px;background:#fafcf7;">
    </div>`;

  dashGrid.appendChild(card);

  // ── Contenido (se genera al abrir) ──
  function renderContenido() {
    const body = card.querySelector("#reportes-body");

    let html = `<div style="display:flex;flex-direction:column;gap:16px;">`;

    for (const [, g] of grupos) {
      const initials = g.nombre
        .split(" ")
        .slice(0, 2)
        .map((w) => w[0] || "")
        .join("")
        .toUpperCase();
      const nRep = g.reportes.length;
      const badge =
        nRep > 1
          ? `<span style="background:#fef3c7;color:#d97706;border:1px solid #fde68a;padding:2px 10px;border-radius:99px;font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700">${nRep} reportes</span>`
          : `<span style="background:#f0fbe6;color:var(--green-dark);border:1px solid #d0e8a8;padding:2px 10px;border-radius:99px;font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700">1 reporte</span>`;

      // Filas de reportes
      const filas = g.reportes
        .map((rep) => {
          const color = colorSit(rep.sit);
          const sitLabel = rep.sit || "Sin especificar";
          const notaHtml =
            rep.nota && rep.nota.trim() && rep.nota !== "nan"
              ? `<span style="color:#718096;font-size:11px;margin-left:8px">· ${rep.nota}</span>`
              : "";
          const fechaHtml =
            rep.fecha && rep.fecha !== "nan"
              ? `<span style="font-family:'JetBrains Mono',monospace;font-size:10px;color:#a0aec0;margin-left:6px">${rep.fecha}</span>`
              : "";
          return `
          <div style="display:flex;align-items:center;gap:10px;padding:8px 14px;background:white;border:1px solid #f0f0f0;border-left:3px solid ${color};border-radius:0 8px 8px 0;flex-wrap:wrap;">
            <span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#718096;min-width:72px">${rep.dia || "—"}</span>
            ${fechaHtml}
            <span style="background:${color}22;color:${color};border:1px solid ${color}44;padding:3px 10px;border-radius:99px;font-size:11px;font-weight:700">${sitLabel}</span>
            ${notaHtml}
          </div>`;
        })
        .join("");

      html += `
        <div style="background:white;border:1px solid #e8f0e0;border-radius:12px;overflow:hidden;">
          <!-- Encabezado del bebé -->
          <div style="display:flex;align-items:center;gap:12px;padding:12px 16px;background:#f7faef;border-bottom:1px solid #e8f0e0;flex-wrap:wrap;">
            <div style="width:38px;height:38px;border-radius:50%;background:var(--green);display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:800;color:white;flex-shrink:0">${initials}</div>
            <div style="flex:1;min-width:180px;">
              <div style="font-weight:700;font-size:13px;color:#2d3748">${g.nombre}</div>
              <div style="font-size:11px;color:#718096;margin-top:1px">${[g.madre, g.inst, g.prog].filter((v) => v && v !== "nan").join(" · ")}</div>
            </div>
            ${badge}
          </div>
          <!-- Reportes del bebé -->
          <div style="display:flex;flex-direction:column;gap:6px;padding:10px 14px;">
            ${filas}
          </div>
        </div>`;
    }

    html += `</div>`;
    body.innerHTML = html;
  }

  // ── Toggle ──
  let abierto = false;
  let rendered = false;
  const arrow = card.querySelector("#reportes-arrow");
  const textBtn = card.querySelector("#reportes-text");

  card.querySelector("#reportes-header").addEventListener("click", () => {
    abierto = !abierto;
    card.querySelector("#reportes-body").style.display = abierto
      ? "block"
      : "none";
    arrow.style.transform = abierto ? "rotate(180deg)" : "rotate(0deg)";
    textBtn.textContent = abierto ? "Ocultar" : "Mostrar";
    if (abierto && !rendered) {
      rendered = true;
      renderContenido();
    }
  });
}

// =============================================================================
//  13. TABLA DE REGISTROS CON PAGINACIÓN
//  Muestra todos los registros en páginas de 50 filas.
//  El CSV exporta el total completo (no solo la página visible).
// =============================================================================

function buildTabla(data) {
  const PAGE_SIZE = 50;
  const modoHistorial = !!selectedBaby;
  const clickable = !modoHistorial;

  const COLS = [
    { key: "bebe", label: "NOMBRE BEBÉ" },
    { key: "madre", label: "NOMBRE MADRE" },
    { key: "institucion", label: "INSTITUCIÓN" },
    { key: "programa", label: "PROGRAMA" },
    { key: "edad", label: "EDAD" },
    { key: "asistencia", label: "ASISTENCIA" },
    { key: "ubicacion", label: "UBICACIÓN" },
    { key: "reporte", label: "REPORTE" },
    { key: "situacion", label: "SITUACIÓN" },
    { key: "nota", label: "NOTA" },
    { key: "dia", label: "DÍA" },
  ]
    .map((c) => ({ ...c, header: colKey(c.key) }))
    .filter((c) => c.header);

  if (!COLS.length) return;
  const cBebe = colKey("bebe");

  // Ordenar: ausentes+reportados primero, luego ausentes, luego presentes
  const dataSorted = [...data].sort((a, b) => {
    const aAus = !esSi(val(a, "asistencia")),
      bAus = !esSi(val(b, "asistencia"));
    if (aAus !== bAus) return aAus ? -1 : 1;
    const aRep = esSi(val(a, "reporte")),
      bRep = esSi(val(b, "reporte"));
    if (aRep !== bRep) return aRep ? -1 : 1;
    return 0;
  });

  const totalPages = Math.ceil(dataSorted.length / PAGE_SIZE);
  let currentPage = 1;

  function makeRow(r) {
    const tr = document.createElement("tr");
    if (clickable) {
      tr.style.cursor = "pointer";
      tr.title = "Click para ver el perfil";
    }
    const ausente = !esSi(val(r, "asistencia"));
    const reportado = esSi(val(r, "reporte"));
    const bgBase =
      ausente && reportado
        ? "rgba(249,115,22,.06)"
        : ausente
          ? "rgba(239,68,68,.05)"
          : "";
    if (bgBase) tr.style.background = bgBase;

    tr.innerHTML = COLS.map((c) => {
      const v = String(r[c.header] || "").trim();
      if (c.key === "asistencia")
        return `<td><span class="${esSi(v) ? "badge-si" : "badge-no"}">${esSi(v) ? "Sí ✓" : "No ✕"}</span></td>`;
      if (c.key === "reporte")
        return `<td>${esSi(v) ? '<span class="badge-reporte">Reportado</span>' : '<span style="color:var(--muted)">—</span>'}</td>`;
      if (c.key === "nota" && v)
        return `<td style="max-width:200px;white-space:normal;line-height:1.4;font-size:11px">${v}</td>`;
      if (c.key === "bebe" && clickable && v)
        return `<td style="color:var(--green-dark);font-weight:700">${v} <span style="font-size:10px;opacity:.6">↗</span></td>`;
      return `<td>${v || '<span style="color:var(--muted)">—</span>'}</td>`;
    }).join("");

    if (clickable && cBebe) {
      const nombre = String(r[cBebe] || "").trim();
      if (nombre) {
        tr.addEventListener("mouseenter", () => {
          tr.style.background = "rgba(133,218,26,.12)";
          tr.style.outline = "1.5px solid rgba(133,218,26,.35)";
        });
        tr.addEventListener("mouseleave", () => {
          tr.style.background = bgBase;
          tr.style.outline = "";
        });
        tr.addEventListener("click", () => {
          const inp = document.getElementById("babySearch");
          if (inp) inp.value = nombre;
          document.getElementById("selectedBabyName").textContent = nombre;
          document.getElementById("selectedBabyTag").style.display = "flex";
          selectBaby(nombre);
          window.scrollTo({ top: 0, behavior: "smooth" });
        });
      }
    }
    return tr;
  }

  function renderPage(page, tbody) {
    tbody.innerHTML = "";
    const start = (page - 1) * PAGE_SIZE;
    const end = Math.min(start + PAGE_SIZE, dataSorted.length);
    for (let i = start; i < end; i++) tbody.appendChild(makeRow(dataSorted[i]));
  }

  function buildPagerHTML(page) {
    if (totalPages <= 1) return "";
    const start = (page - 1) * PAGE_SIZE + 1;
    const end = Math.min(page * PAGE_SIZE, dataSorted.length);
    const pagesToShow = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - page) <= 2)
        pagesToShow.push(p);
    }
    const btns = [];
    pagesToShow.forEach((p, idx) => {
      if (idx > 0 && p - pagesToShow[idx - 1] > 1)
        btns.push(
          `<span style="font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted);padding:0 2px">…</span>`,
        );
      btns.push(
        `<button data-page="${p}" style="font-family:'JetBrains Mono',monospace;font-size:10px;padding:5px 10px;border:1.5px solid ${p === page ? "var(--green)" : "var(--border)"};border-radius:6px;background:${p === page ? "var(--green)" : "transparent"};color:${p === page ? "white" : "var(--text-mid)"};cursor:pointer;font-weight:${p === page ? "700" : "400"};min-width:32px">${p}</button>`,
      );
    });
    return `
      <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 22px;border-top:1px solid var(--border);flex-wrap:wrap;gap:10px;">
        <div style="font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--muted)">
          Mostrando <strong style="color:var(--text)">${start}–${end}</strong> de <strong style="color:var(--text)">${dataSorted.length}</strong> · Página <strong style="color:var(--text)">${page}</strong>/${totalPages}
        </div>
        <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap;">
          <button id="btnPrev" ${page <= 1 ? "disabled" : ""} style="font-family:'JetBrains Mono',monospace;font-size:10px;padding:5px 12px;border:1.5px solid var(--border);border-radius:6px;background:transparent;color:${page <= 1 ? "var(--muted)" : "var(--green-dark)"};cursor:${page <= 1 ? "default" : "pointer"};opacity:${page <= 1 ? ".35" : "1"}">← Ant</button>
          ${btns.join("")}
          <button id="btnNext" ${page >= totalPages ? "disabled" : ""} style="font-family:'JetBrains Mono',monospace;font-size:10px;padding:5px 12px;border:1.5px solid var(--border);border-radius:6px;background:transparent;color:${page >= totalPages ? "var(--muted)" : "var(--green-dark)"};cursor:${page >= totalPages ? "default" : "pointer"};opacity:${page >= totalPages ? ".35" : "1"}">Sig →</button>
        </div>
      </div>`;
  }

  const card = document.createElement("div");
  card.className = "card col-12";
  card.innerHTML = `
    <div class="card-head" style="padding:16px 22px">
      <div class="card-title"><span class="card-dot"></span>🗂 Detalle de registros
        ${clickable ? `<span style="font-family:'JetBrains Mono',monospace;font-size:9px;letter-spacing:1px;color:var(--muted);font-weight:400;margin-left:8px">· CLICK EN FILA PARA VER PERFIL</span>` : ""}
      </div>
      <div style="display:flex;align-items:center;gap:10px">
        <span class="card-badge">${dataSorted.length} REGISTROS</span>
        <button id="btnExportCSV" style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:1px;background:transparent;border:1px solid var(--border);color:var(--muted);padding:4px 10px;border-radius:6px;cursor:pointer;transition:.2s"
          onmouseover="this.style.borderColor='var(--green)';this.style.color='var(--green)'"
          onmouseout="this.style.borderColor='var(--border)';this.style.color='var(--muted)'">
          ⬇ Exportar Excel (${dataSorted.length} filas)
        </button>
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr>${COLS.map((c) => `<th>${c.label}</th>`).join("")}</tr></thead>
        <tbody id="tablaBody"></tbody>
      </table>
    </div>
    <div id="tablaPager"></div>`;
  dashGrid.appendChild(card);

  const tablaBody = card.querySelector("#tablaBody");
  const pagerDiv = card.querySelector("#tablaPager");

  function goToPage(page) {
    currentPage = Math.max(1, Math.min(page, totalPages));
    renderPage(currentPage, tablaBody);
    pagerDiv.innerHTML = buildPagerHTML(currentPage);
    pagerDiv.querySelector("#btnPrev")?.addEventListener("click", () => {
      goToPage(currentPage - 1);
      card.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    pagerDiv.querySelector("#btnNext")?.addEventListener("click", () => {
      goToPage(currentPage + 1);
      card.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    pagerDiv.querySelectorAll("[data-page]").forEach((btn) =>
      btn.addEventListener("click", () => {
        goToPage(parseInt(btn.dataset.page));
        card.scrollIntoView({ behavior: "smooth", block: "start" });
      }),
    );
  }

  // Exportar Excel (.xlsx) con formato de tabla profesional
  card.querySelector("#btnExportCSV").addEventListener("click", () => {
    exportarExcel(dataSorted, COLS);
  });

  goToPage(1);
}

// =============================================================================
//  ETL — EXPORTACIÓN DE TABLA CONSOLIDADA (limpia, sin duplicados, lista para dashboards)
//
//  Reglas aplicadas:
//  1. Un registro por bebé (clave: NOMBRE BEBÉ + NOMBRE MADRE)
//  2. Si hay múltiples registros del mismo bebé → conservar el más completo
//     y en caso de empate, el más reciente según el campo DÍA
//  3. Limpieza de texto: trim, capitalización, espacios dobles
//  4. Columnas en orden estándar, con ID único al inicio
//  5. Exporta .xlsx con formato de tabla profesional
// =============================================================================

function exportarTablaConsolidada() {
  if (!allData.length) {
    showToast("No hay datos cargados", true);
    return;
  }

  // ── Helpers de limpieza ────────────────────────────────────

  /** Limpia un string: trim + elimina espacios dobles + capitaliza primera letra de cada palabra */
  function limpiarTexto(str) {
    if (!str) return "";
    return String(str)
      .trim()
      .replace(/\s+/g, " ") // espacios dobles → uno
      .replace(/\b\w/g, (c) => c.toUpperCase()); // capitalizar palabras
  }

  /** Limpia texto en MAYÚSCULAS (para institución, programa) */
  function limpiarMayus(str) {
    if (!str) return "";
    return String(str).trim().replace(/\s+/g, " ").toUpperCase();
  }

  /** Cuenta campos no vacíos en un registro */
  function completitud(r) {
    return [
      r.NombreBebe,
      r.NombreMadre,
      r.InstitucionMadre,
      r.ProgramaMadre,
      r.Edad,
      r.Asistencia,
      r.Ubicacion,
      r.Reporte,
      r.SituacionEspecifica,
      r.Nota,
      r.Dia,
    ].filter((v) => v && String(v).trim() !== "" && String(v).trim() !== "No")
      .length;
  }

  /** Orden de día para comparar cuál es "más reciente" */
  const ORDEN_DIA = {
    lunes: 1,
    martes: 2,
    miercoles: 3,
    miércoles: 3,
    jueves: 4,
    viernes: 5,
  };

  function pesodia(dia) {
    return ORDEN_DIA[norm(dia || "")] || 0;
  }

  // ── Paso 1: consolidar por bebé (nombre + madre) ─────────
  // Para cada bebé único, quedarse con el registro más completo.
  // Si hay empate, el del día más reciente gana.
  const bebeMap = new Map();

  for (const r of allData) {
    const nombre = limpiarTexto(r.NombreBebe || "");
    const madre = limpiarTexto(r.NombreMadre || "");
    if (!nombre) continue;

    const key = norm(nombre) + "|" + norm(madre);

    if (!bebeMap.has(key)) {
      bebeMap.set(key, r);
    } else {
      const prev = bebeMap.get(key);
      const prevComp = completitud(prev);
      const currComp = completitud(r);

      // Gana el más completo; en empate, el del día más reciente
      if (currComp > prevComp) {
        bebeMap.set(key, r);
      } else if (currComp === prevComp && pesodia(r.Dia) > pesodia(prev.Dia)) {
        bebeMap.set(key, r);
      }
    }
  }

  // ── Paso 2: normalizar y limpiar cada registro ────────────
  const ORDEN_COLS = [
    // ID único primero (clave de negocio)
    {
      key: "id",
      label: "ID",
      get: (r) =>
        limpiarTexto(r.NombreBebe) + " - " + limpiarTexto(r.NombreMadre),
    },
    {
      key: "bebe",
      label: "NOMBRE BEBÉ",
      get: (r) => limpiarTexto(r.NombreBebe),
    },
    {
      key: "madre",
      label: "NOMBRE MADRE",
      get: (r) => limpiarTexto(r.NombreMadre),
    },
    {
      key: "institucion",
      label: "INSTITUCIÓN",
      get: (r) => limpiarMayus(r.InstitucionMadre),
    },
    {
      key: "programa",
      label: "PROGRAMA",
      get: (r) => limpiarTexto(r.ProgramaMadre),
    },
    {
      key: "edad",
      label: "EDAD",
      get: (r) => {
        const e = String(r.Edad || "").trim();
        if (e === "6-15") return "6-15 meses";
        if (e === "16-30") return "16-30 meses";
        return e;
      },
    },
    {
      key: "asistencia",
      label: "ASISTENCIA",
      get: (r) => (esSi(r.Asistencia) ? "Sí" : "No"),
    },
    {
      key: "ubicacion",
      label: "UBICACIÓN",
      get: (r) => limpiarTexto(r.Ubicacion),
    },
    {
      key: "reporte",
      label: "REPORTE",
      get: (r) => (esSi(r.Reporte) ? "Reportado" : ""),
    },
    {
      key: "situacion",
      label: "SITUACIÓN",
      get: (r) => limpiarTexto(r.SituacionEspecifica),
    },
    { key: "nota", label: "NOTA", get: (r) => limpiarTexto(r.Nota) },
    { key: "dia", label: "DÍA", get: (r) => limpiarTexto(r.Dia) },
    {
      key: "visitante",
      label: "TIPO",
      get: (r) =>
        esSi(r.Visitante) ? "Extra" : esSi(r.NoCidi) ? "No CIDI" : "Programa",
    },
  ];

  // Ordenar por día de la semana, luego alfabéticamente por nombre
  const consolidados = Array.from(bebeMap.values()).sort((a, b) => {
    const da = pesodia(a.Dia),
      db = pesodia(b.Dia);
    if (da !== db) return da - db;
    return norm(a.NombreBebe || "").localeCompare(norm(b.NombreBebe || ""));
  });

  if (!consolidados.length) {
    showToast("No se encontraron datos válidos", true);
    return;
  }

  // Construir filas limpias
  const headers = ORDEN_COLS.map((c) => c.label);
  const filas = consolidados.map((r) => ORDEN_COLS.map((c) => c.get(r)));

  // ── Paso 3: estadísticas del consolidado ─────────────────
  const nPresentes = consolidados.filter((r) => esSi(r.Asistencia)).length;
  const nAusentes = consolidados.length - nPresentes;
  const nReportes = consolidados.filter((r) => esSi(r.Reporte)).length;
  const nExtras = consolidados.filter((r) => esSi(r.Visitante)).length;
  const nNoCidi = consolidados.filter((r) => esSi(r.NoCidi)).length;

  // ── Paso 4: construir Excel con formato ──────────────────
  const wb = XLSX.utils.book_new();

  // Hoja principal — tabla consolidada
  const ws = XLSX.utils.aoa_to_sheet([headers, ...filas]);

  // Anchos de columna ajustados
  ws["!cols"] = ORDEN_COLS.map((c, i) => ({
    wch: Math.min(
      Math.max(c.label.length, ...filas.map((f) => String(f[i] || "").length)) +
        2,
      45,
    ),
  }));

  // Inmovilizar primera fila (encabezados siempre visibles)
  ws["!freeze"] = {
    xSplit: 0,
    ySplit: 1,
    topLeftCell: "A2",
    activePane: "bottomLeft",
  };

  // Filtros automáticos
  ws["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: 0, c: ORDEN_COLS.length - 1 },
    }),
  };

  // Altura de filas
  ws["!rows"] = [{ hpt: 22 }, ...Array(filas.length).fill({ hpt: 15 })];

  // Paleta de colores
  const C = {
    headerBg: "85DA1A", // verde Juanfe
    headerFg: "FFFFFF",
    presenteBg: "F0FBE6", // verde suave
    ausenteBg: "FEF2F2", // rojo suave
    reportBg: "FFF7ED", // naranja suave
    extraBg: "EFF6FF", // azul suave
    nocidiBg: "F5F3FF", // morado suave
    pairBg: "F7FAF2", // gris muy verde (filas pares)
    text: "1A2E05",
    border: "D0E8A8",
  };

  const nFilas = filas.length + 1;
  const nCols = ORDEN_COLS.length;

  for (let R = 0; R < nFilas; R++) {
    for (let Col = 0; Col < nCols; Col++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: Col });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };

      let bgColor = C.pairBg;
      const esEnc = R === 0;

      if (!esEnc) {
        const rec = consolidados[R - 1];
        const asiste = esSi(rec.Asistencia);
        const reporto = esSi(rec.Reporte);
        const esExtra = esSi(rec.Visitante);
        const esNoCidi = esSi(rec.NoCidi);

        if (esExtra) bgColor = C.extraBg;
        else if (esNoCidi) bgColor = C.nocidiBg;
        else if (!asiste && reporto) bgColor = C.reportBg;
        else if (!asiste) bgColor = C.ausenteBg;
        else bgColor = C.presenteBg;
      }

      ws[addr].s = {
        font: {
          name: "Calibri",
          sz: esEnc ? 11 : 10,
          bold: esEnc,
          color: { rgb: esEnc ? C.headerFg : C.text },
        },
        fill: {
          patternType: "solid",
          fgColor: { rgb: esEnc ? C.headerBg : bgColor },
        },
        border: {
          top: { style: "thin", color: { rgb: C.border } },
          bottom: { style: "thin", color: { rgb: C.border } },
          left: { style: "thin", color: { rgb: C.border } },
          right: { style: "thin", color: { rgb: C.border } },
        },
        alignment: {
          vertical: "center",
          horizontal: esEnc
            ? "center"
            : ["ASISTENCIA", "REPORTE", "EDAD", "TIPO"].includes(
                  ORDEN_COLS[Col]?.label,
                )
              ? "center"
              : "left",
          wrapText: false,
        },
      };
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "Tabla Consolidada");

  // Hoja de resumen
  const fecha = new Date().toLocaleDateString("es-CO", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric",
  });
  const wsRes = XLSX.utils.aoa_to_sheet([
    ["RESUMEN ETL — FUNDACIÓN JUANFE"],
    [""],
    ["Fecha de exportación:", fecha],
    [""],
    ["RESULTADO DE CONSOLIDACIÓN"],
    ["Total bebés únicos:", consolidados.length],
    ["Presentes:", nPresentes],
    ["Ausentes:", nAusentes],
    ["Con reporte:", nReportes],
    ["Extras (visitantes):", nExtras],
    ["No CIDI:", nNoCidi],
    [
      "Tasa de asistencia:",
      consolidados.length
        ? `${Math.round((nPresentes / consolidados.length) * 100)}%`
        : "0%",
    ],
    [""],
    ["FUENTES INCLUIDAS"],
    ...loadedFiles.map((f) => ["", f.name, `${f.rowCount} filas originales`]),
    [""],
    ["REGLAS APLICADAS"],
    ["", "1 registro por bebé (NOMBRE BEBÉ + NOMBRE MADRE)"],
    ["", "Registro más completo conservado"],
    ["", "En empate: registro del día más reciente"],
    ["", "Texto limpiado: espacios, capitalización"],
    ["", "Edad normalizada a rangos estándar"],
    ["", "Asistencia y Reporte estandarizados"],
  ]);
  wsRes["!cols"] = [{ wch: 28 }, { wch: 40 }, { wch: 22 }];
  if (wsRes["A1"])
    wsRes["A1"].s = {
      font: { name: "Calibri", sz: 14, bold: true, color: { rgb: "1A2E05" } },
      fill: { patternType: "solid", fgColor: { rgb: "F0FAD8" } },
    };
  if (wsRes["A5"])
    wsRes["A5"].s = {
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "1A2E05" } },
    };

  XLSX.utils.book_append_sheet(wb, wsRes, "Resumen ETL");

  // Descargar
  const hoy = new Date();
  const ts = `${hoy.getFullYear()}${String(hoy.getMonth() + 1).padStart(2, "0")}${String(hoy.getDate()).padStart(2, "0")}`;
  const nombreArchivo = `juanfe_consolidado_${ts}_${consolidados.length}bebes.xlsx`;

  try {
    XLSX.writeFile(wb, nombreArchivo);
    showToast(
      `✅ Tabla consolidada: ${consolidados.length} bebés únicos exportados`,
    );
  } catch {
    // Fallback sin estilos
    const wbS = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wbS,
      XLSX.utils.aoa_to_sheet([headers, ...filas]),
      "Tabla Consolidada",
    );
    XLSX.writeFile(wbS, nombreArchivo);
    showToast(`Exportado (sin estilos): ${consolidados.length} bebés únicos`);
  }
}

//
//  Usa la librería SheetJS (XLSX) ya cargada en dashboard.html.
//  Genera un archivo .xlsx con:
//    • Encabezados con fondo verde Juanfe y texto blanco en negrita
//    • Filas alternadas en blanco/verde muy claro (zebra)
//    • Ausentes destacados en rojo suave
//    • Presentes en verde suave
//    • Reportados en naranja suave
//    • Columnas con ancho ajustado automáticamente al contenido
//    • Fila de resumen al final (totales)
//    • Filtros automáticos en los encabezados
// =============================================================================

function exportarExcel(datos, COLS) {
  if (!datos.length) {
    showToast("Sin datos para exportar", true);
    return;
  }

  // ── 1. Construir matriz de datos ─────────────────────────
  const headers = COLS.map((c) => c.label);

  // Función para limpiar el valor de una celda
  const celda = (r, c) => String(r[c.header] || "").trim();

  // Normalizar el valor de Asistencia para que sea legible
  const celdaAsistencia = (v) => {
    const n = v
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
    return n === "si" || n === "si" || n === "1" || n === "true" ? "Sí" : "No";
  };
  const celdaReporte = (v) => {
    const n = v
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
    return n === "si" || n === "si" || n === "1" || n === "true"
      ? "Reportado"
      : "";
  };

  const filas = datos.map((r) =>
    COLS.map((c) => {
      if (c.key === "asistencia") return celdaAsistencia(celda(r, c));
      if (c.key === "reporte") return celdaReporte(celda(r, c));
      return celda(r, c);
    }),
  );

  // Fila de resumen al final
  const nPresentes = datos.filter((r) => {
    const v = celda(
      r,
      COLS.find((c) => c.key === "asistencia") || { header: "Asistencia" },
    );
    return celdaAsistencia(v) === "Sí";
  }).length;
  const nAusentes = datos.length - nPresentes;
  const nReportes = datos.filter((r) => {
    const v = celda(
      r,
      COLS.find((c) => c.key === "reporte") || { header: "Reporte" },
    );
    return celdaReporte(v) === "Reportado";
  }).length;

  // ── 2. Crear workbook y worksheet ────────────────────────
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...filas]);

  // ── 3. Calcular anchos de columna ────────────────────────
  const anchos = COLS.map((c, colIdx) => {
    const maxContenido = Math.max(
      c.label.length,
      ...filas.map((fila) => String(fila[colIdx] || "").length),
    );
    // Mínimo 10, máximo 40, con un poco de padding
    return { wch: Math.min(Math.max(maxContenido + 2, 10), 40) };
  });
  ws["!cols"] = anchos;

  // ── 4. Rango de la hoja ────────────────────────────────────
  const nFilas = filas.length + 1; // +1 por el encabezado
  const nColumnas = COLS.length;
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: nFilas - 1, c: nColumnas - 1 },
  });

  // ── 5. Estilos de celdas ─────────────────────────────────
  // SheetJS Community Edition no soporta estilos nativos.
  // Usamos SheetJS con la extensión de estilos inline (xlsx-style compatible).
  // Para máxima compatibilidad, generamos el archivo con datos limpios y
  // agregamos estilos via el objeto de celda si la versión lo permite.

  // Color verde Juanfe
  const VERDE_JUANFE = "85DA1A";
  const VERDE_TEXTO = "FFFFFF";
  const VERDE_CLARO = "F0FAD8"; // fondo par
  const BLANCO = "FFFFFF"; // fondo impar
  const ROJO_SUAVE = "FEF2F2"; // ausente
  const VERDE_SUAVE = "F0FBE6"; // presente
  const NARANJA_SUAVE = "FFF7ED"; // reportado

  // Aplicar estilos a cada celda
  for (let R = 0; R < nFilas; R++) {
    for (let C = 0; C < nColumnas; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };

      // Determinar tipo de fila
      const esEncabezado = R === 0;
      let fgColor = BLANCO;

      if (!esEncabezado) {
        const filaData = datos[R - 1];
        const colAsis = COLS.find((c) => c.key === "asistencia");
        const colRep = COLS.find((c) => c.key === "reporte");
        const asiste = colAsis
          ? celdaAsistencia(celda(filaData, colAsis)) === "Sí"
          : false;
        const reporto = colRep
          ? celdaReporte(celda(filaData, colRep)) === "Reportado"
          : false;

        if (!asiste && reporto) fgColor = NARANJA_SUAVE;
        else if (!asiste) fgColor = ROJO_SUAVE;
        else if (asiste) fgColor = VERDE_SUAVE;
        else fgColor = R % 2 === 0 ? VERDE_CLARO : BLANCO;
      }

      ws[addr].s = {
        font: {
          name: "Calibri",
          sz: esEncabezado ? 11 : 10,
          bold: esEncabezado,
          color: { rgb: esEncabezado ? VERDE_TEXTO : "1A2E05" },
        },
        fill: {
          patternType: "solid",
          fgColor: { rgb: esEncabezado ? VERDE_JUANFE : fgColor },
        },
        border: {
          top: { style: "thin", color: { rgb: "D0E8A8" } },
          bottom: { style: "thin", color: { rgb: "D0E8A8" } },
          left: { style: "thin", color: { rgb: "D0E8A8" } },
          right: { style: "thin", color: { rgb: "D0E8A8" } },
        },
        alignment: {
          vertical: "center",
          horizontal: esEncabezado ? "center" : "left",
          wrapText: false,
        },
      };
    }
  }

  // Altura de filas: encabezado más alto
  ws["!rows"] = [{ hpt: 22 }, ...Array(filas.length).fill({ hpt: 16 })];

  // Filtros automáticos en la fila de encabezado
  ws["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: 0, c: nColumnas - 1 },
    }),
  };

  // ── 6. Segunda hoja: Resumen ─────────────────────────────
  const fecha = new Date().toLocaleDateString("es-CO", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
  const wsResumen = XLSX.utils.aoa_to_sheet([
    ["RESUMEN DE ASISTENCIA — FUNDACIÓN JUANFE"],
    [""],
    ["Fecha de exportación:", fecha],
    ["Total registros:", datos.length],
    ["Presentes:", nPresentes],
    ["Ausentes:", nAusentes],
    ["Con reporte:", nReportes],
    [
      "Tasa de asistencia:",
      datos.length ? `${Math.round((nPresentes / datos.length) * 100)}%` : "0%",
    ],
    [""],
    ["Archivos incluidos:", loadedFiles.map((f) => f.name).join(", ")],
  ]);

  // Estilo del título del resumen
  if (wsResumen["A1"]) {
    wsResumen["A1"].s = {
      font: { name: "Calibri", sz: 14, bold: true, color: { rgb: "1A2E05" } },
      fill: { patternType: "solid", fgColor: { rgb: "F0FAD8" } },
      alignment: { horizontal: "left", vertical: "center" },
    };
  }
  wsResumen["!cols"] = [{ wch: 25 }, { wch: 40 }];

  // ── 7. Agregar hojas y descargar ─────────────────────────
  const nombreHoja = "Asistencia";
  const nombreResumen = "Resumen";

  XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
  XLSX.utils.book_append_sheet(wb, wsResumen, nombreResumen);

  // Nombre del archivo con fecha
  const hoy = new Date();
  const yyyymmdd = `${hoy.getFullYear()}${String(hoy.getMonth() + 1).padStart(2, "0")}${String(hoy.getDate()).padStart(2, "0")}`;
  const nombreArchivo = `asistencia_juanfe_${yyyymmdd}_${datos.length}registros.xlsx`;

  try {
    XLSX.writeFile(wb, nombreArchivo);
    showToast(`Excel exportado: ${datos.length} registros`);
  } catch (e) {
    // Si la versión CDN de XLSX no soporta estilos, exportar sin ellos
    console.warn("Estilos no soportados, exportando sin formato:", e.message);
    const wbSimple = XLSX.utils.book_new();
    const wsSimple = XLSX.utils.aoa_to_sheet([headers, ...filas]);
    wsSimple["!cols"] = anchos;
    XLSX.utils.book_append_sheet(wbSimple, wsSimple, nombreHoja);
    XLSX.writeFile(wbSimple, nombreArchivo);
    showToast(`Excel exportado (sin estilos): ${datos.length} registros`);
  }
}

function chartOpts(isPie) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 400 },
    plugins: {
      legend: {
        labels: {
          color: "#8ba869",
          font: { family: "JetBrains Mono", size: 11 },
          padding: 14,
          boxWidth: 10,
          boxHeight: 10,
        },
      },
      tooltip: {
        backgroundColor: "#fff",
        titleColor: "#1a2e05",
        bodyColor: "#8ba869",
        borderColor: "#e0eccc",
        borderWidth: 1,
        padding: 12,
        titleFont: { family: "Plus Jakarta Sans", size: 12, weight: "700" },
        bodyFont: { family: "JetBrains Mono", size: 11 },
        callbacks: {
          label: (ctx) => {
            const v = ctx.parsed?.y ?? ctx.parsed ?? ctx.raw;
            return `  ${ctx.dataset.label || ctx.label}: ${typeof v === "number" ? v.toLocaleString("es") : v}`;
          },
        },
      },
    },
    scales: isPie
      ? undefined
      : {
          x: {
            ticks: {
              color: "#8ba869",
              font: { family: "JetBrains Mono", size: 10 },
              maxRotation: 40,
            },
            grid: { color: "rgba(133,218,26,.07)" },
            border: { color: "#e0eccc" },
          },
          y: {
            ticks: {
              color: "#8ba869",
              font: { family: "JetBrains Mono", size: 10 },
            },
            grid: { color: "rgba(133,218,26,.07)" },
            border: { color: "#e0eccc" },
            beginAtZero: true,
          },
        },
  };
}

function createCard(canvasId, title, colClass, badge, chartH = 280) {
  const div = document.createElement("div");
  div.className = `card ${colClass}`;
  div.innerHTML = `
    <div class="card-head">
      <div class="card-title"><span class="card-dot"></span>${title}</div>
      ${badge ? `<span class="card-badge">${badge}</span>` : ""}
    </div>
    <div class="chart-wrap" style="height:${chartH}px;position:relative"><canvas id="${canvasId}"></canvas></div>`;
  return div;
}

function showLoading(on) {
  loadingEl.classList.toggle("show", on);
}

function showToast(msg, error = false) {
  toastEl.textContent = (error ? "❌ " : "✅ ") + msg;
  toastEl.className = "toast show" + (error ? " error" : "");
  setTimeout(() => toastEl.classList.remove("show"), 3500);
}
