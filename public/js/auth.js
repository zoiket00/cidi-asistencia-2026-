// js/auth.js — Módulo compartido de autenticación · Fundación Juanfe

const PERMISOS = {
  admin: ["asistencia", "bebes", "dashboard", "importar"],
  coordinadora: ["asistencia", "bebes", "dashboard", "importar"],
  profesora: ["asistencia", "dashboard"],
};

const PAGINAS = [
  {
    id: "asistencia",
    href: "/",
    label: "Asist.",
    svg: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/><path d="M9 12l2 2 4-4"/></svg>`,
  },
  {
    id: "bebes",
    href: "/bebes",
    label: "Base",
    svg: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v14c0 1.66 4.03 3 9 3s9-1.34 9-3V5"/><path d="M3 12c0 1.66 4.03 3 9 3s9-1.34 9-3"/></svg>`,
  },
  {
    id: "dashboard",
    href: "/dashboard",
    label: "Gráficas",
    svg: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
  },
  {
    id: "importar",
    href: "/importar",
    label: "Import.",
    svg: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`,
  },
];

let _supabase = null;
let _session = null;
let _usuario = null;

// ── Inicializar cliente Supabase ──────────────────────────────────
async function initSupabase() {
  if (_supabase) return _supabase;
  const cfg = await fetch("/api/config").then((r) => r.json());
  const { createClient } = window.supabase;
  _supabase = createClient(cfg.url, cfg.anonKey, {
    auth: {
      storage: window.sessionStorage,
      persistSession: true,
      autoRefreshToken: true,
    },
  });
  return _supabase;
}

// ── Verificar sesión ──────────────────────────────────────────────
async function requireAuth(paginaActual) {
  try {
    const sb = await initSupabase();
    const {
      data: { session },
      error: errS,
    } = await sb.auth.getSession();

    if (errS || !session) {
      window.location.href = "/login";
      return null;
    }

    _session = session;

    const { data: usuario, error: errU } = await sb
      .from("usuarios")
      .select("rol, nombre, email")
      .eq("id", session.user.id)
      .single();

    if (errU || !usuario) {
      await sb.auth.signOut();
      window.location.href = "/login";
      return null;
    }

    _usuario = usuario;

    // Redirigir si el rol no tiene acceso a esta página
    if (paginaActual && !PERMISOS[usuario.rol]?.includes(paginaActual)) {
      window.location.href = "/bienvenida";
      return null;
    }

    return usuario;
  } catch (e) {
    console.error("Error en requireAuth:", e);
    window.location.href = "/login";
    return null;
  }
}

// ── Renderizar sidebar según rol ──────────────────────────────────
function renderSidebar(paginaActual, rol) {
  const nav = document.querySelector(".sidebar-nav");
  if (!nav) return;

  const permitidas = PERMISOS[rol] || [];
  nav.innerHTML = "";

  PAGINAS.filter((p) => permitidas.includes(p.id)).forEach((p) => {
    const a = document.createElement("a");
    a.href = p.href;
    a.className = "sidebar-item" + (p.id === paginaActual ? " active" : "");
    a.title = p.label;
    a.innerHTML = `${p.svg}<span>${p.label}</span>`;
    nav.appendChild(a);
  });

  renderUserSection();
}

// ── Sección de usuario + logout al fondo del sidebar ─────────────
function renderUserSection() {
  const sidebar = document.querySelector(".sidebar");
  if (!sidebar || !_usuario) return;

  // Quitar sección previa si existe
  sidebar.querySelector(".sidebar-user")?.remove();

  const inicial = (_usuario.nombre || _usuario.email || "U")[0].toUpperCase();
  const colores = {
    admin: "#1565c0",
    coordinadora: "#2e7d32",
    profesora: "#6a1b9a",
  };
  const color = colores[_usuario.rol] || "#555";
  const nombre = _usuario.nombre || _usuario.email;

  const div = document.createElement("div");
  div.className = "sidebar-user";
  div.innerHTML = `
    <div class="sidebar-user-circle" style="background:${color}" title="${nombre}">${inicial}</div>
    <div class="sidebar-user-nombre">${nombre.split(" ")[0]}</div>
    <div class="sidebar-user-rol">${_usuario.rol}</div>
    <button class="sidebar-user-logout" id="btnLogout" title="Cerrar sesión">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" width="14" height="14">
        <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/>
        <polyline points="16 17 21 12 16 7"/>
        <line x1="21" y1="12" x2="9" y2="12"/>
      </svg>
      <span>Salir</span>
    </button>`;

  sidebar.appendChild(div);
  document.getElementById("btnLogout").addEventListener("click", logout);
}

// ── Logout ────────────────────────────────────────────────────────
async function logout() {
  try {
    // Limpiar sessionStorage primero para evitar loops
    Object.keys(sessionStorage)
      .filter((k) => k.startsWith("sb-"))
      .forEach((k) => sessionStorage.removeItem(k));

    // Limpiar datos de asistencia del día para que el próximo usuario
    // no vea datos de la sesión anterior si comparte el mismo navegador
    localStorage.removeItem("datos_asistencia");

    const sb = await initSupabase();
    await sb.auth.signOut();
  } catch (e) {}
  window.location.href = "/login";
}

// ── Token para requests a /api/* ──────────────────────────────────
function getToken() {
  return _session?.access_token || "";
}

// ── Fetch autenticado ─────────────────────────────────────────────
async function authFetch(url, options = {}) {
  return fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${getToken()}`,
      "Content-Type": options.headers?.["Content-Type"] || "application/json",
    },
  });
}
