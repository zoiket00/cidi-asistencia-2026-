// js/bienvenida.js — Pantalla de bienvenida · Fundación Juanfe

const OPCIONES_ROL = {
  admin: [
    { href: "/", label: "Listado de asistencias", primary: true },
    { href: "/bebes", label: "Base de datos" },
    { href: "/dashboard", label: "Dashboard" },
    { href: "/importar", label: "Importar archivos" },
  ],
  coordinadora: [
    { href: "/", label: "Listado de asistencias", primary: true },
    { href: "/bebes", label: "Base de datos" },
    { href: "/dashboard", label: "Dashboard" },
    { href: "/importar", label: "Importar archivos" },
  ],
  profesora: [
    { href: "/", label: "Listado de asistencias", primary: true },
    { href: "/dashboard", label: "Dashboard" },
  ],
};

const SALUDO = {
  admin: "Bienvenido, administrador",
  coordinadora: "Bienvenida, coordinadora",
  profesora: "Bienvenida, profesora",
};

(async () => {
  const usuario = await requireAuth("asistencia"); // todos tienen asistencia
  if (!usuario) return;

  // Título personalizado
  document.getElementById("bienvenidaTitulo").textContent = usuario.nombre
    ? `Bienvenida, ${usuario.nombre}`
    : SALUDO[usuario.rol];

  // Opciones según rol
  const opciones = OPCIONES_ROL[usuario.rol] || [];
  const container = document.getElementById("bienvenidaOpciones");
  opciones.forEach((op) => {
    const a = document.createElement("a");
    a.href = op.href;
    a.className = "opcion-btn" + (op.primary ? " opcion-primary" : "");
    a.textContent = op.label;
    container.appendChild(a);
  });

  // Sidebar dinámico
  renderSidebar("bienvenida", usuario.rol);
})();
