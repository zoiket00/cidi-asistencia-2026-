// js/login.js — Lógica del login · Fundación Juanfe
let _supabase = null;

// Limpiar sesión previa al cargar el login — sin llamar APIs de Supabase
// Solo limpiamos las claves del sessionStorage para evitar loops
Object.keys(sessionStorage)
  .filter((k) => k.startsWith("sb-"))
  .forEach((k) => sessionStorage.removeItem(k));

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

function mostrarError(msg) {
  const el = document.getElementById("loginError");
  el.textContent = msg;
  el.classList.add("show");
}

function limpiarError() {
  document.getElementById("loginError").classList.remove("show");
}

async function login() {
  const email = document.getElementById("loginEmail").value.trim();
  const pass = document.getElementById("loginPass").value;
  const btn = document.getElementById("btnLogin");

  if (!email || !pass) {
    mostrarError("Ingresa tu correo y contraseña.");
    return;
  }

  limpiarError();
  btn.disabled = true;
  btn.textContent = "Ingresando...";

  try {
    const sb = await initSupabase();
    const { data, error } = await sb.auth.signInWithPassword({
      email,
      password: pass,
    });

    if (error) throw error;

    // Verificar que el usuario tiene rol asignado
    const { data: usuario, error: errU } = await sb
      .from("usuarios")
      .select("rol, nombre")
      .eq("id", data.user.id)
      .single();

    if (errU || !usuario) {
      await sb.auth.signOut();
      throw new Error(
        "Usuario sin permisos asignados. Contacta al administrador.",
      );
    }

    // Ir a bienvenida
    window.location.href = "/bienvenida";
  } catch (e) {
    const msg = e.message?.includes("Invalid login")
      ? "Correo o contraseña incorrectos."
      : e.message || "Error al iniciar sesión.";
    mostrarError(msg);
    btn.disabled = false;
    btn.textContent = "Ingresar";
  }
}

// Event listeners
document.getElementById("btnLogin").addEventListener("click", login);
document.getElementById("loginPass").addEventListener("keydown", (e) => {
  if (e.key === "Enter") login();
});
document.getElementById("loginEmail").addEventListener("keydown", (e) => {
  if (e.key === "Enter") document.getElementById("loginPass").focus();
});
document.getElementById("loginEmail").addEventListener("input", limpiarError);
document.getElementById("loginPass").addEventListener("input", limpiarError);

// Toggle ver contraseña
document.getElementById("btnVerPass").addEventListener("click", () => {
  const input = document.getElementById("loginPass");
  input.type = input.type === "password" ? "text" : "password";
});
