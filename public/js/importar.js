// importar.js — Importador de historial · Fundación Juanfe
// ── Estado ────────────────────────────────────────────────
let archivos = []; // { nombre, fecha, dia, registros, estado, msg, importado }
let filtroActivo = "todos";
let importando = false;

const DIAS_VALIDOS = [
  "Lunes",
  "Martes",
  "Miercoles",
  "Miércoles",
  "Jueves",
  "Viernes",
];

// Normalizar día — quita tilde de Miércoles
function normDia(d = "") {
  return d
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// ── Drag & Drop ───────────────────────────────────────────
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () =>
  dropZone.classList.remove("drag-over"),
);
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  procesarArchivos([...e.dataTransfer.files]);
});
dropZone.addEventListener("click", () => fileInput.click());
document.getElementById("dropLink").addEventListener("click", (e) => {
  e.stopPropagation();
  fileInput.click();
});
fileInput.addEventListener("change", () =>
  procesarArchivos([...fileInput.files]),
);

// ── Procesar archivos seleccionados ───────────────────────
async function procesarArchivos(files) {
  const xlsxFiles = files.filter((f) => f.name.endsWith(".xlsx"));
  if (!xlsxFiles.length) {
    toast("Solo se aceptan archivos .xlsx", true);
    return;
  }

  // Cargar fechas/dias ya existentes en Supabase para detectar duplicados
  let existentes = new Set();
  try {
    const res = await authFetch("/api/asistencia/fechas");
    if (res.ok) {
      const data = await res.json();
      existentes = new Set(data.fechas.map((f) => `${f.fecha}|${f.dia}`));
    }
  } catch (e) {
    console.warn("No se pudo cargar fechas existentes:", e.message);
  }

  for (const file of xlsxFiles) {
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
        defval: "",
      });

      if (!rows.length) {
        archivos.push({
          nombre: file.name,
          estado: "error",
          msg: "Archivo vacío",
          registros: [],
          fecha: "—",
          dia: "—",
        });
        continue;
      }

      // Detectar fecha y día del archivo
      const primeraFila = rows[0];
      const fecha = String(primeraFila.Fecha || primeraFila.fecha || "").trim();
      const dia = normDia(
        String(
          primeraFila.Dia || primeraFila.Día || primeraFila.dia || "",
        ).trim(),
      );

      if (!fecha || !dia) {
        archivos.push({
          nombre: file.name,
          estado: "error",
          msg: "No se encontró Fecha o Día",
          registros: rows,
          fecha: "—",
          dia: "—",
        });
        continue;
      }

      if (!DIAS_VALIDOS.includes(dia)) {
        archivos.push({
          nombre: file.name,
          estado: "error",
          msg: `Día no reconocido: "${dia}"`,
          registros: rows,
          fecha,
          dia,
        });
        continue;
      }

      const key = `${fecha}|${dia}`;

      // Duplicado en Supabase
      const esDupBD = existentes.has(key);

      // Duplicado local — ya hay otro archivo con la misma fecha+dia en la lista actual
      const dupLocal = archivos.find((a) => a.fecha === fecha && a.dia === dia);

      const esDuplicado = esDupBD || !!dupLocal;
      let msgDup = "";
      if (esDupBD) msgDup = `Ya existe en BD (${rows.length} registros)`;
      else if (dupLocal) msgDup = `Misma fecha que "${dupLocal.nombre}"`;

      archivos.push({
        nombre: file.name,
        fecha,
        dia,
        registros: rows,
        estado: esDuplicado ? "duplicado" : "nuevo",
        msg: esDuplicado
          ? msgDup
          : `Listo para importar (${rows.length} registros)`,
        importado: false,
      });
    } catch (e) {
      archivos.push({
        nombre: file.name,
        estado: "error",
        msg: "Error leyendo el archivo: " + e.message,
        registros: [],
        fecha: "—",
        dia: "—",
      });
    }
  }

  actualizarStats();
  renderLista();
  mostrarUI();
}

// ── Importar ──────────────────────────────────────────────
async function importar() {
  const pendientes = archivos.filter(
    (a) => a.estado === "nuevo" && !a.importado,
  );
  const duplicados = archivos.filter((a) => a.estado === "duplicado");
  const conErrores = archivos.filter((a) => a.estado === "error");

  // Si hay duplicados o errores, mostrar aviso antes de importar
  if (duplicados.length || conErrores.length) {
    let partes = [];
    if (duplicados.length) {
      const lista = duplicados
        .map((a) => `<li><strong>${a.nombre}</strong> — ${a.msg}</li>`)
        .join("");
      partes.push(
        `<p><strong>Archivos duplicados (${duplicados.length}):</strong></p><ul>${lista}</ul><p>Estos archivos <strong>no se importarán</strong>. Si quieres conservar uno, elimina el otro de la lista primero.</p>`,
      );
    }
    if (conErrores.length) {
      const lista = conErrores
        .map((a) => `<li><strong>${a.nombre}</strong> — ${a.msg}</li>`)
        .join("");
      partes.push(
        `<p><strong>Archivos con error (${conErrores.length}):</strong></p><ul>${lista}</ul><p>Estos archivos tampoco se importarán. Revísalos y vuelve a cargarlos.</p>`,
      );
    }
    if (pendientes.length) {
      partes.push(
        `<p>Se importarán <strong>${pendientes.length} archivo${pendientes.length > 1 ? "s" : ""} nuevo${pendientes.length > 1 ? "s" : ""}</strong> sin problema.</p>`,
      );
    } else {
      partes.push(
        `<p>No hay archivos nuevos para importar. Elimina los duplicados y vuelve a cargar el correcto.</p>`,
      );
    }

    document.getElementById("infoTitulo").textContent =
      "⚠️  Revisar antes de importar";
    document.getElementById("infoTitulo").style.color = "#f57f17";
    document.getElementById("infoCuerpo").innerHTML = partes.join("");
    document.getElementById("infoModal").style.background = "rgba(0,0,0,0.45)";
    document.getElementById("infoBox").style.borderTop = "4px solid #f57f17";

    // Si hay pendientes, cambiar el botón del modal a "Continuar de todas formas"
    const btnOk = document.getElementById("btnInfoOk");
    if (pendientes.length) {
      btnOk.textContent = "Importar los nuevos";
      btnOk.onclick = () => {
        cerrarInfo();
        ejecutarImportacion(pendientes);
      };
    } else {
      btnOk.textContent = "Entendido";
      btnOk.onclick = cerrarInfo;
    }

    document.getElementById("infoModal").classList.add("open");
    return;
  }

  if (!pendientes.length) {
    toast("No hay archivos nuevos para importar", true);
    return;
  }
  ejecutarImportacion(pendientes);
}

async function ejecutarImportacion(pendientes) {
  importando = true;
  document.getElementById("btnImportar").disabled = true;
  document.getElementById("btnImportar").textContent = "Importando...";

  let ok = 0,
    errores = 0;

  for (const arch of pendientes) {
    // Marcar como procesando
    arch.estado = "procesando";
    renderLista();

    try {
      const registros = arch.registros
        .map((r) => ({
          NombreBebe: String(
            r["Nombre Bebé"] || r["Nombre Bebe"] || r.NombreBebe || "",
          ).trim(),
          NombreMadre: String(r["Nombre Madre"] || r.NombreMadre || "").trim(),
          Fase: String(
            r["Institución"] || r["Institucion"] || r.Fase || r.fase || "",
          ).trim(),
          ProgramaMadre: String(r["Programa"] || r.ProgramaMadre || "").trim(),
          Edad: String(r["Edad (meses)"] || r.Edad || "").trim(),
          Asistencia: String(r["Asistencia"] || "No").trim(),
          Ubicacion: String(
            r["Ubicación"] || r["Ubicacion"] || r.Ubicacion || "",
          ).trim(),
          Reporte: String(r["Reporte"] || "No").trim(),
          SituacionEspecifica: String(
            r["Situación Específica"] ||
              r["Situacion Especifica"] ||
              r.SituacionEspecifica ||
              "",
          ).trim(),
          Nota: String(r["Nota"] || "").trim(),
          Extras: String(r["Extras"] || r["Visitante"] || "").trim(),
          NoCidi: String(r["No CIDI"] || r["NoCidi"] || "").trim(),
        }))
        .filter((r) => r.NombreBebe);

      const res = await authFetch("/api/asistencia/guardar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fecha: arch.fecha, dia: arch.dia, registros }),
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`HTTP ${res.status}: ${text.slice(0, 200)}`);
      }

      const data = await res.json();
      if (!data.ok) throw new Error(data.error);

      arch.estado = "importado";
      arch.importado = true;

      // El servidor ahora devuelve `omitidos` = registros que ya existían
      // y fueron actualizados (no nuevos). Así la usuaria sabe exactamente
      // qué pasó en lugar de ver solo un número sin contexto.
      if (data.omitidos > 0) {
        arch.msg = `✓ ${data.guardados} nuevos · ${data.omitidos} actualizados`;
      } else {
        arch.msg = `✓ Importado (${data.guardados} registros)`;
      }
      ok++;
    } catch (e) {
      arch.estado = "error";
      arch.msg = "Error: " + e.message;
      errores++;
    }

    renderLista();
    actualizarStats();
    // Pequeña pausa para no saturar el servidor
    await new Promise((r) => setTimeout(r, 150));
  }

  importando = false;
  document.getElementById("btnImportar").disabled = false;
  document.getElementById("btnImportar").textContent = "⬆ Importar nuevos";
  toast(
    `Importación completa — ${ok} archivos guardados${errores ? `, ${errores} errores` : ""}`,
    errores > 0,
  );
  actualizarStats();
}

// ── Mensajes explicativos por situación ───────────────────
const EXPLICACIONES = {
  duplicado_bd: {
    titulo: "Archivo ya importado",
    icono: "🔁",
    color: "#1565c0",
    bg: "#e3f2fd",
    texto: (arch) => `
            <p>El archivo <strong>${arch.nombre}</strong> ya existe en la base de datos de Supabase.</p>
            <p>Corresponde al día <strong>${arch.dia}</strong> del <strong>${arch.fecha}</strong>, que ya fue guardado anteriormente.</p>
            <p><strong>¿Qué puedes hacer?</strong></p>
            <ul>
              <li>Si el archivo es idéntico al que ya está en BD, no necesitas hacer nada.</li>
              <li>Si el archivo tiene correcciones o cambios, usa el botón <em>Forzar</em> para sobreescribir los registros existentes.</li>
            </ul>`,
  },
  duplicado_local: {
    titulo: "Fecha repetida en esta carga",
    icono: "⚠️",
    color: "#f57f17",
    bg: "#fff8e1",
    texto: (arch) => `
            <p>Ya cargaste otro archivo con la misma fecha y día: <strong>${arch.dia} ${arch.fecha}</strong>.</p>
            <p>Esto pasa frecuentemente cuando:</p>
            <ul>
              <li>Descargaste el mismo Excel dos veces y quedó como <em>"archivo(2).xlsx"</em>.</li>
              <li>Tienes dos versiones del mismo día con correcciones.</li>
            </ul>
            <p><strong>¿Qué hacer?</strong> Revisa cuál de los dos archivos es el correcto y elimina el otro de la lista usando <em>Limpiar lista</em> para volver a cargar solo el que necesitas.</p>`,
  },
  error: {
    titulo: "Error al leer el archivo",
    icono: "❌",
    color: "#c62828",
    bg: "#ffebee",
    texto: (arch) => `
            <p>No se pudo procesar <strong>${arch.nombre}</strong>.</p>
            <p><strong>Motivo:</strong> ${arch.msg}</p>
            <p><strong>Causas frecuentes:</strong></p>
            <ul>
              <li>El archivo está corrupto o no es un Excel válido.</li>
              <li>El archivo no tiene las columnas esperadas (Fecha, Día, Nombre Bebé, etc.).</li>
              <li>El archivo es de un formato muy antiguo (.xls en lugar de .xlsx).</li>
            </ul>
            <p>Intenta exportarlo de nuevo desde el sistema de asistencia.</p>`,
  },
};

function mostrarExplicacion(arch) {
  let tipo;
  if (arch.estado === "error") tipo = "error";
  else if (
    arch.estado === "duplicado" &&
    arch.msg.startsWith("Ya existe en BD")
  )
    tipo = "duplicado_bd";
  else if (arch.estado === "duplicado") tipo = "duplicado_local";
  else return;

  const exp = EXPLICACIONES[tipo];
  document.getElementById("infoTitulo").textContent =
    exp.icono + "  " + exp.titulo;
  document.getElementById("infoTitulo").style.color = exp.color;
  document.getElementById("infoCuerpo").innerHTML = exp.texto(arch);
  document.getElementById("infoModal").style.background = exp.bg + "cc";
  document.getElementById("infoBox").style.borderTop = `4px solid ${exp.color}`;
  document.getElementById("infoModal").classList.add("open");
}

// ── Render lista ──────────────────────────────────────────
function renderLista() {
  const lista = document.getElementById("listaArchivos");
  const filtrados =
    filtroActivo === "todos"
      ? archivos
      : archivos.filter((a) => {
          if (filtroActivo === "nuevo")
            return a.estado === "nuevo" || a.estado === "procesando";
          if (filtroActivo === "duplicado") return a.estado === "duplicado";
          if (filtroActivo === "error") return a.estado === "error";
          return true;
        });

  if (!filtrados.length) {
    lista.innerHTML = `<div class="lista-vacia">No hay archivos en esta categoría</div>`;
    return;
  }

  lista.innerHTML = "";
  filtrados.forEach((arch) => {
    const div = document.createElement("div");
    div.className = "archivo-item";

    const estadoClass =
      {
        nuevo: "est-nuevo",
        duplicado: "est-dup",
        error: "est-err",
        importado: "est-ok",
        procesando: "est-proc",
      }[arch.estado] || "";

    const estadoTexto =
      {
        nuevo: "Nuevo",
        duplicado: "Duplicado",
        error: "Error",
        importado: "Importado",
        procesando: "Importando...",
      }[arch.estado] || arch.estado;

    // Botón Eliminar en TODOS los archivos
    const tieneAccion = true;

    div.innerHTML = `
            <div class="arch-icon">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            </div>
            <div class="arch-info">
              <span class="arch-nombre">${arch.nombre}</span>
              <span class="arch-meta">${arch.fecha} · ${arch.dia}</span>
            </div>
            <span class="arch-msg">${arch.msg}</span>
            <span class="estado-badge ${estadoClass}">${estadoTexto}</span>
            <button class="btn-eliminar-arch" data-nombre="${arch.nombre}" title="Eliminar de la lista">Eliminar</button>`;

    // Botón eliminar — quita el archivo y recalcula duplicados locales
    const btnElim = div.querySelector(".btn-eliminar-arch");
    if (btnElim) {
      btnElim.addEventListener("click", () => {
        archivos = archivos.filter((a) => a.nombre !== btnElim.dataset.nombre);

        // Recalcular duplicados locales — el que quedaba marcado como dup
        // puede que ya no lo sea si eliminamos su "gemelo"
        const vistos = new Map(); // key fecha|dia → primer archivo
        archivos.forEach((a) => {
          if (a.estado === "importado" || a.estado === "procesando") return;
          const key = `${a.fecha}|${a.dia}`;
          if (a.msg && a.msg.startsWith("Misma fecha")) {
            // Era dup local — revisar si su gemelo sigue en la lista
            const gemelo = archivos.find(
              (x) => x !== a && x.fecha === a.fecha && x.dia === a.dia,
            );
            if (!gemelo) {
              // Ya no tiene gemelo — pasa a nuevo
              a.estado = "nuevo";
              a.msg = `Listo para importar (${a.registros?.length || 0} registros)`;
            }
          }
        });

        renderLista();
        actualizarStats();
        if (!archivos.length) {
          document.getElementById("statsRow").style.display = "none";
          document.getElementById("impActions").style.display = "none";
          document.getElementById("archivosSection").style.display = "none";
        }
      });
    }

    lista.appendChild(div);
  });
}

// ── Stats ─────────────────────────────────────────────────
function actualizarStats() {
  const total = archivos.length;
  const nuevos = archivos.filter((a) => a.estado === "nuevo").length;
  const dups = archivos.filter((a) => a.estado === "duplicado").length;
  const errs = archivos.filter((a) => a.estado === "error").length;
  const regs = archivos.reduce((s, a) => s + (a.registros?.length || 0), 0);

  document.getElementById("statTotal").textContent = total;
  document.getElementById("statNuevos").textContent = nuevos;
  document.getElementById("statDups").textContent = dups;
  document.getElementById("statErrs").textContent = errs;
  document.getElementById("statRegs").textContent = regs;

  // Actualizar btn importar
  const btn = document.getElementById("btnImportar");
  const pendientes = archivos.filter((a) => a.estado === "nuevo").length;
  btn.textContent = pendientes
    ? `⬆ Importar ${pendientes} archivo${pendientes > 1 ? "s" : ""}`
    : "⬆ Importar nuevos";
  btn.disabled = !pendientes || importando;
}

function mostrarUI() {
  document.getElementById("statsRow").style.display = "flex";
  document.getElementById("impActions").style.display = "flex";
  document.getElementById("archivosSection").style.display = "block";
}

// ── Limpiar ───────────────────────────────────────────────
document.getElementById("btnLimpiarImp").addEventListener("click", () => {
  archivos = [];
  fileInput.value = "";
  document.getElementById("statsRow").style.display = "none";
  document.getElementById("impActions").style.display = "none";
  document.getElementById("archivosSection").style.display = "none";
  document.getElementById("listaArchivos").innerHTML = "";
});

// ── Filtros ───────────────────────────────────────────────
document.querySelectorAll(".filtro-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document
      .querySelectorAll(".filtro-btn")
      .forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    filtroActivo = btn.dataset.f;
    renderLista();
  });
});

document.getElementById("btnImportar").addEventListener("click", importar);

// ── Modal informativo ─────────────────────────────────────
function cerrarInfo() {
  document.getElementById("infoModal").classList.remove("open");
}
document.getElementById("infoClose").addEventListener("click", cerrarInfo);
document.getElementById("btnInfoOk").addEventListener("click", cerrarInfo);
document.getElementById("infoModal").addEventListener("click", (e) => {
  if (e.target === document.getElementById("infoModal")) cerrarInfo();
});

// ── Toast ─────────────────────────────────────────────────
function toast(msg, error = false) {
  const t = document.getElementById("toastImp");
  t.textContent = msg;
  t.className = "toast-imp show" + (error ? " error" : "");
  setTimeout(() => t.classList.remove("show"), 4000);
}

// =============================================================================
//  EXPORTAR — solo admin y coordinadora
// =============================================================================

/**
 * initExportar(rol)
 * Muestra u oculta la sección de exportar según el rol.
 * Conecta los eventos de vista previa y descarga.
 */
function initExportar(rol) {
  if (!["admin", "coordinadora"].includes(rol)) return;

  // Mostrar divisor y sección exportar
  document.getElementById("expDivisor").style.display = "flex";
  document.getElementById("expSection").style.display = "block";

  // Mostrar divisor y sección eliminar
  document.getElementById("delDivisor").style.display = "flex";
  document.getElementById("delSection").style.display = "block";

  // Fecha por defecto: mes actual
  const hoy = new Date();
  const primerDia = new Date(hoy.getFullYear(), hoy.getMonth(), 1)
    .toISOString()
    .slice(0, 10);
  const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0)
    .toISOString()
    .slice(0, 10);
  document.getElementById("expDesde").value = primerDia;
  document.getElementById("expHasta").value = ultimoDia;

  document
    .getElementById("btnExpPreview")
    .addEventListener("click", cargarPreview);
  document
    .getElementById("btnExpDescargar")
    .addEventListener("click", descargarExcel);
  document
    .getElementById("btnDelDia")
    .addEventListener("click", confirmarEliminarDia);

  // Modal eliminar
  document.getElementById("delConfirmClose").addEventListener("click", () => {
    document.getElementById("delConfirmModal").classList.remove("open");
  });
  document.getElementById("delConfirmNo").addEventListener("click", () => {
    document.getElementById("delConfirmModal").classList.remove("open");
  });
}

/** Construye la URL de exportar con los filtros actuales */
function buildExpUrl(extra = "") {
  const desde = document.getElementById("expDesde").value;
  const hasta = document.getElementById("expHasta").value;
  const programa = document.getElementById("expPrograma").value;
  const fase = document.getElementById("expFase").value;
  const dia = document.getElementById("expDia").value;

  if (!desde || !hasta) {
    toast("Selecciona fecha inicio y fecha fin", true);
    return null;
  }
  if (desde > hasta) {
    toast("La fecha inicio no puede ser mayor a la fecha fin", true);
    return null;
  }

  const params = new URLSearchParams({ desde, hasta });
  if (programa) params.append("programa", programa);
  if (fase) params.append("fase", fase);
  if (dia) params.append("dia", dia);
  if (extra) params.append(extra, "true");

  return `/api/exportar?${params.toString()}`;
}

/** Carga la vista previa — primeras 15 filas + total */
async function cargarPreview() {
  const url = buildExpUrl("preview");
  if (!url) return;

  const btn = document.getElementById("btnExpPreview");
  btn.textContent = "Cargando...";
  btn.disabled = true;

  try {
    const res = await authFetch(url);
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || "Error cargando preview");
    if (!data.total) {
      toast("No hay registros en ese rango y filtros seleccionados", true);
      document.getElementById("expPreviewWrap").style.display = "none";
      document.getElementById("btnExpDescargar").disabled = true;
      return;
    }

    // Renderizar tabla
    renderTablaPreview(data.filas, data.total);
    document.getElementById("btnExpDescargar").disabled = false;
    toast(`✓ ${data.total} registros encontrados`);
  } catch (e) {
    toast("Error: " + e.message, true);
  } finally {
    btn.textContent = "👁 Vista previa";
    btn.disabled = false;
  }
}

/** Renderiza la tabla de vista previa */
function renderTablaPreview(filas, total) {
  const wrap = document.getElementById("expPreviewWrap");
  const tabla = document.getElementById("expTabla");
  const totalEl = document.getElementById("expPreviewTotal");

  totalEl.textContent = `${total.toLocaleString()} registros en total`;

  if (!filas.length) {
    wrap.style.display = "none";
    return;
  }

  const cols = Object.keys(filas[0]);

  const thead = `<thead><tr>${cols.map((c) => `<th>${c}</th>`).join("")}</tr></thead>`;
  const tbody =
    "<tbody>" +
    filas
      .map(
        (f) => `<tr>${cols.map((c) => `<td>${f[c] ?? ""}</td>`).join("")}</tr>`,
      )
      .join("") +
    "</tbody>";

  tabla.innerHTML = thead + tbody;
  wrap.style.display = "block";
}

/** Descarga el Excel generado por el servidor */
async function descargarExcel() {
  const url = buildExpUrl();
  if (!url) return;

  const btn = document.getElementById("btnExpDescargar");
  btn.textContent = "Generando...";
  btn.disabled = true;

  try {
    const res = await authFetch(url);

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || "Error generando Excel");
    }

    // Descargar el archivo desde la respuesta binaria
    const blob = await res.blob();
    const nombreArchivo =
      res.headers.get("Content-Disposition")?.match(/filename="(.+)"/)?.[1] ||
      `juanfe_asistencia.xlsx`;

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nombreArchivo;
    a.click();
    URL.revokeObjectURL(a.href);

    toast("✓ Excel descargado correctamente");
  } catch (e) {
    toast("Error: " + e.message, true);
  } finally {
    btn.textContent = "↓ Descargar Excel";
    btn.disabled = false;
  }
}

/** Muestra modal de confirmación antes de eliminar el día */
function confirmarEliminarDia() {
  const fecha = document.getElementById("delFecha").value;
  const dia = document.getElementById("delDia").value;

  if (!fecha || !dia) {
    toast("Selecciona fecha y día antes de eliminar", true);
    return;
  }

  // Formatear fecha legible
  const [y, m, d] = fecha.split("-");
  const fechaLegible = `${d}/${m}/${y}`;

  document.getElementById("delConfirmMsg").innerHTML =
    `¿Estás segura de eliminar <strong>todos los registros</strong> del <strong>${dia} ${fechaLegible}</strong>?<br><br>
     <span style="color:#c62828;font-size:12px">Esta acción no se puede deshacer. La profesora deberá volver a exportar ese día.</span>`;

  const modal = document.getElementById("delConfirmModal");
  modal.classList.add("open");

  document.getElementById("delConfirmSi").onclick = async () => {
    modal.classList.remove("open");
    await eliminarDia(fecha, dia);
  };
}

/** Elimina todos los registros de un día vía API */
async function eliminarDia(fecha, dia) {
  const btn = document.getElementById("btnDelDia");
  btn.textContent = "Eliminando...";
  btn.disabled = true;

  try {
    const res = await authFetch("/api/asistencia/dia", {
      method: "DELETE",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fecha, dia }),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Error eliminando");

    toast(`✓ ${data.eliminados} registros de ${dia} ${fecha} eliminados`);

    // Limpiar campos
    document.getElementById("delFecha").value = "";
    document.getElementById("delDia").value = "";
  } catch (e) {
    toast("Error: " + e.message, true);
  } finally {
    btn.textContent = "🗑 Eliminar listado";
    btn.disabled = false;
  }
}
