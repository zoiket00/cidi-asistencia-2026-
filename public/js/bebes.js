// bebes.js — Base de Datos · Fundación Juanfe
const DIAS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"];
const PROGRAMA_BADGE = {
  Cocina: "badge-cocina",
  Hotelería: "badge-hoteleria",
  Hoteleria: "badge-hoteleria",
  Belleza: "badge-belleza",
  "Auxiliar Administrativo": "badge-auxiliar",
  Otro: "badge-otro",
};
const PAGE_SIZE = 30;

let allBebes = [],
  filtered = [],
  currentPage = 1;

async function cargar() {
  try {
    const [r1, r2] = await Promise.all([
      authFetch("/api/bebes"),
      authFetch("/api/asistencia-dias"),
    ]);
    const { bebes } = await r1.json();
    const { diasMap = {} } = r2.ok ? await r2.json() : {};
    allBebes = bebes
      .map((b) => ({ ...b, dias: diasMap[b.NombreBebe] || [] }))
      .sort((a, b) => a.NombreBebe.localeCompare(b.NombreBebe));
    document.getElementById("listaTitle").textContent =
      `Registros actuales (${allBebes.length})`;
    aplicarFiltro();
  } catch {
    toast("Error cargando datos", true);
  }
}

function norm(s = "") {
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function aplicarFiltro() {
  const t = norm(document.getElementById("listaSearch").value);
  filtered = t
    ? allBebes.filter(
        (b) =>
          norm(b.NombreBebe).includes(t) ||
          norm(b.NombreMadre).includes(t) ||
          norm(b.Fase).includes(t) ||
          norm(b.ProgramaMadre).includes(t),
      )
    : [...allBebes];
  currentPage = 1;
  renderLista();
  renderPag();
}

function renderLista() {
  const el = document.getElementById("listaBebes");
  const slice = filtered.slice(
    (currentPage - 1) * PAGE_SIZE,
    currentPage * PAGE_SIZE,
  );
  if (!slice.length) {
    el.innerHTML = `<div class="lista-empty">No se encontraron registros.</div>`;
    return;
  }
  el.innerHTML = "";
  slice.forEach((b) => {
    const card = document.createElement("div");
    card.className = "bebe-card";
    const badge = PROGRAMA_BADGE[b.ProgramaMadre]
      ? `<span class="badge-programa ${PROGRAMA_BADGE[b.ProgramaMadre]}">${b.ProgramaMadre}</span>`
      : "";
    const diasHtml = b.dias
      .map((d) => `<span class="dia-tag">${d}</span>`)
      .join("");
    card.innerHTML = `
            <div class="card-main">
              <div class="card-row1">
                <span class="card-nombre">${b.NombreBebe}</span>
                ${badge}
                <span class="card-meta">${[b.Fase, b.Edad ? b.Edad + " meses" : ""].filter(Boolean).join(" · ")}</span>
              </div>
              <div class="card-row2">
                <span class="card-madre">${b.NombreMadre || ""}</span>
                ${diasHtml ? `<span class="card-sep">·</span><div class="card-dias">${diasHtml}</div>` : ""}
              </div>
            </div>
            <div class="card-actions">
              <button class="card-btn-edit" title="Editar">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" width="13" height="13"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
              </button>
              <button class="card-btn-del" title="Eliminar">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" width="13" height="13"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
              </button>
            </div>`;
    card
      .querySelector(".card-btn-edit")
      .addEventListener("click", () => cargarEnForm(b));
    card
      .querySelector(".card-btn-del")
      .addEventListener("click", () => confirmar(b.id, b.NombreBebe));
    el.appendChild(card);
  });
}

function renderPag() {
  const total = Math.ceil(filtered.length / PAGE_SIZE);
  const pag = document.getElementById("pagination");
  if (total <= 1) {
    pag.innerHTML = "";
    return;
  }
  let h = `<span class="page-info">${filtered.length} registros</span>`;
  h += `<button class="page-btn" id="pPrev" ${currentPage === 1 ? "disabled" : ""}>←</button>`;
  for (let p = 1; p <= total; p++) {
    if (total > 6 && p > 2 && p < total - 1 && Math.abs(p - currentPage) > 1) {
      if (p === 3 || p === total - 2) h += `<span class="page-dots">…</span>`;
      continue;
    }
    h += `<button class="page-btn ${p === currentPage ? "active" : ""}" data-p="${p}">${p}</button>`;
  }
  h += `<button class="page-btn" id="pNext" ${currentPage === total ? "disabled" : ""}>→</button>`;
  pag.innerHTML = h;
  pag.querySelectorAll(".page-btn[data-p]").forEach((b) =>
    b.addEventListener("click", () => {
      currentPage = +b.dataset.p;
      renderLista();
      renderPag();
    }),
  );
  pag.querySelector("#pPrev")?.addEventListener("click", () => {
    currentPage--;
    renderLista();
    renderPag();
  });
  pag.querySelector("#pNext")?.addEventListener("click", () => {
    currentPage++;
    renderLista();
    renderPag();
  });
}

function cargarEnForm(b) {
  // Abrir modal de editar
  document.getElementById("editId").value = b.id || "";
  document.getElementById("eNombre").value = b.NombreBebe || "";
  document.getElementById("eMadre").value = b.NombreMadre || "";
  document.getElementById("eFase").value = b.Fase || "";
  document.getElementById("ePrograma").value = b.ProgramaMadre || "";
  document.getElementById("eEdad").value = b.Edad || "";
  // TSF deshabilita programa
  const esTSF = b.Fase === "TSF";
  document.getElementById("ePrograma").disabled = esTSF;
  // Pills de días en el modal
  document
    .querySelectorAll("[data-edit-dia]")
    .forEach((p) =>
      p.classList.toggle("active", (b.dias || []).includes(p.dataset.editDia)),
    );
  document.getElementById("editModal").classList.add("open");
  document.getElementById("eNombre").focus();
}

function cerrarModalEditar() {
  document.getElementById("editModal").classList.remove("open");
}

async function guardarEditar() {
  const id = document.getElementById("editId").value;
  const nombre_bebe = document.getElementById("eNombre").value.trim();
  const nombre_madre = document.getElementById("eMadre").value.trim();
  if (!nombre_bebe || !nombre_madre) {
    toast("Nombre del bebé y madre son obligatorios", true);
    return;
  }
  const dias = [...document.querySelectorAll("[data-edit-dia].active")].map(
    (p) => p.dataset.editDia,
  );
  const body = {
    nombre_bebe,
    nombre_madre,
    Fase: document.getElementById("eFase").value,
    programa: document.getElementById("ePrograma").value,
    edad: document.getElementById("eEdad").value,
    dias,
  };

  cerrarModalEditar();
  showMiniModal({
    title: "¿Guardar cambios?",
    desc: `<strong>${nombre_bebe}</strong><br><span style="color:#888">${nombre_madre}</span>`,
    confirmText: "Guardar cambios",
    color: "green",
    onConfirm: async () => {
      try {
        const res = await authFetch(`/api/bebes/${id}`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        toast("Registro actualizado ✓");
        await cargar();
      } catch (e) {
        toast("Error: " + e.message, true);
      }
    },
  });
}

function limpiar() {
  document.getElementById("editingId").value = "";
  document.getElementById("formTitle").textContent = "Nuevo registro";
  ["fNombre", "fMadre"].forEach(
    (id) => (document.getElementById(id).value = ""),
  );
  ["fFase", "fPrograma", "fEdad"].forEach(
    (id) => (document.getElementById(id).value = ""),
  );
  document
    .querySelectorAll(".dia-pill")
    .forEach((p) => p.classList.remove("active"));
}

async function guardar() {
  const id = document.getElementById("editingId").value;
  const nombre_bebe = document.getElementById("fNombre").value.trim();
  const nombre_madre = document.getElementById("fMadre").value.trim();
  if (!nombre_bebe || !nombre_madre) {
    toast("Nombre del bebé y madre son obligatorios", true);
    return;
  }
  const body = {
    nombre_bebe,
    nombre_madre,
    Fase: document.getElementById("fFase").value,
    programa: document.getElementById("fPrograma").value,
    edad: document.getElementById("fEdad").value,
    dias: [...document.querySelectorAll(".dia-pill.active")].map(
      (p) => p.dataset.dia,
    ),
  };

  const accion = id ? "Guardar cambios" : "Registrar bebé";
  const desc = `<strong>${nombre_bebe}</strong><br><span style="color:#888">${nombre_madre}</span>`;

  showMiniModal({
    title: id ? "¿Guardar cambios?" : "¿Confirmar nuevo registro?",
    desc,
    confirmText: accion,
    color: "green",
    onConfirm: async () => {
      try {
        const res = await authFetch(id ? `/api/bebes/${id}` : "/api/bebes", {
          method: id ? "PUT" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error);
        toast(id ? "Registro actualizado ✓" : "Registro guardado ✓");
        limpiar();
        await cargar();
      } catch (e) {
        toast("Error: " + e.message, true);
      }
    },
  });
}

function confirmar(id, nombre) {
  showMiniModal({
    title: "Eliminar registro",
    desc: `¿Eliminar a <strong>${nombre}</strong>? Esta acción no se puede deshacer.`,
    confirmText: "Eliminar",
    color: "red",
    onConfirm: () => eliminar(id),
  });
}

async function eliminar(id) {
  try {
    const res = await authFetch(`/api/bebes/${id}`, { method: "DELETE" });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error);
    toast("Registro eliminado");
    await cargar();
  } catch (e) {
    toast("Error: " + e.message, true);
  }
}

function toast(msg, error = false) {
  const t = document.getElementById("toastBebes");
  t.textContent = msg;
  t.className = "toast-bebes show" + (error ? " error" : "");
  setTimeout(() => t.classList.remove("show"), 3000);
}

// ── Mini modal minimalista ─────────────────────────────────────────────────────
// color: "green" para guardar/editar, "red" para eliminar
function showMiniModal({
  title,
  desc,
  confirmText,
  color = "green",
  onConfirm,
}) {
  const modal = document.getElementById("miniModal");
  const accent = document.getElementById("miniModalAccent");
  const titleEl = document.getElementById("miniModalTitle");
  const descEl = document.getElementById("miniModalDesc");
  const btnOk = document.getElementById("miniModalConfirm");
  const btnNo = document.getElementById("miniModalCancel");

  const bg = color === "red" ? "#c62828" : "#2e7d32";
  accent.style.background = bg;
  btnOk.style.background = bg;
  titleEl.textContent = title;
  descEl.innerHTML = desc;
  btnOk.textContent = confirmText;

  modal.classList.add("open");

  const close = () => modal.classList.remove("open");
  btnNo.onclick = close;
  modal.onclick = (e) => {
    if (e.target === modal) close();
  };
  btnOk.onclick = () => {
    close();
    onConfirm();
  };
}

document.getElementById("btnGuardar").addEventListener("click", guardar);
document.getElementById("btnLimpiar").addEventListener("click", limpiar);
document.getElementById("listaSearch").addEventListener("input", aplicarFiltro);

// Pills del form nuevo
document
  .querySelectorAll(".dia-pill[data-dia]")
  .forEach((p) =>
    p.addEventListener("click", () => p.classList.toggle("active")),
  );

// Pills del modal editar
document
  .querySelectorAll("[data-edit-dia]")
  .forEach((p) =>
    p.addEventListener("click", () => p.classList.toggle("active")),
  );

document.getElementById("fFase").addEventListener("change", (e) => {
  const prog = document.getElementById("fPrograma");
  prog.disabled = e.target.value === "TSF";
  if (e.target.value === "TSF") prog.value = "";
});

document.getElementById("eFase").addEventListener("change", (e) => {
  const prog = document.getElementById("ePrograma");
  prog.disabled = e.target.value === "TSF";
  if (e.target.value === "TSF") prog.value = "";
});

// Modal editar
document
  .getElementById("btnGuardarEdit")
  .addEventListener("click", guardarEditar);
document
  .getElementById("btnCerrarEdit")
  .addEventListener("click", cerrarModalEditar);
document
  .getElementById("btnCancelarEdit")
  .addEventListener("click", cerrarModalEditar);
document.getElementById("editModal").addEventListener("click", (e) => {
  if (e.target === document.getElementById("editModal")) cerrarModalEditar();
});

// cargar() ya NO se llama aquí automáticamente.
// Lo invoca el script inline de bebes.html DESPUÉS de requireAuth().
