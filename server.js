/**
 * server.js — Fundación Juanfe · CIDI
 *
 * ENDPOINTS:
 * ─────────────────────────────────────────────────────
 *  GET  /api/sheet/:dia           → lista bebés del día (CSV para app.js)
 *  GET  /api/bebes                → todos los bebés del catálogo
 *  GET  /api/dias                 → días disponibles
 *  POST /api/bebes                → agregar o actualizar bebé
 *
 *  POST /api/asistencia/guardar   → guarda asistencia del día en Supabase
 *  GET  /api/asistencia           → registros para el dashboard (con filtros)
 *
 *  POST /api/historico/guardar    → guarda Excel en disco (respaldo)
 *  GET  /api/historico            → lista Excels guardados
 *  GET  /api/historico/:nombre    → descarga Excel
 *
 * COMPATIBILIDAD DE COLUMNAS:
 *  El campo "fase" (nuevo) era "institucion" en Excels viejos de profesoras.
 *  El server siempre devuelve el campo como "Fase" en los responses JSON,
 *  pero acepta tanto "Fase" como "InstitucionMadre"/"Institucion" en los POST.
 */

require("dotenv").config();
const express = require("express");
const path = require("path");
const XLSX = require("xlsx"); // Generación de Excel en el servidor para /api/exportar
const { createClient } = require("@supabase/supabase-js");

// ── Validar entorno antes de arrancar ──────────────────────────────────────────
if (
  !process.env.SUPABASE_URL ||
  !process.env.SUPABASE_SERVICE_KEY ||
  !process.env.SUPABASE_ANON_KEY
) {
  console.error(
    "❌  Falta SUPABASE_URL, SUPABASE_SERVICE_KEY o SUPABASE_ANON_KEY en el archivo .env",
  );
  process.exit(1);
}

const app = express();
const PORT = process.env.PORT || 3000;

const DIAS_VALIDOS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"];

// ── Conexión a Supabase ────────────────────────────────────────────────────────
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY,
);

supabase
  .from("bebes")
  .select("id", { count: "exact", head: true })
  .then(({ count, error }) => {
    if (error) console.error("❌  Error conectando a Supabase:", error.message);
    else console.log(`✅  Supabase conectado — ${count} bebés en la BD`);
  });

// ── Normalización y fuzzy-matching de nombres ─────────────────────────────────

function normName(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .trim()
    .replace(/\s+/g, " ");
}

// Distancia de Levenshtein normalizada → similitud [0,1]
// Detecta typos: "Churio" → "Chourio", "Epinosa" → "Espinosa"
function strSim(a, b) {
  if (a === b) return 1;
  const m = a.length, n = b.length;
  if (!m || !n) return 0;
  const dp = Array.from({ length: m + 1 }, (_, i) => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] =
        a[i - 1] === b[j - 1]
          ? dp[i - 1][j - 1]
          : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
  return 1 - dp[m][n] / Math.max(m, n);
}

// Solapamiento de tokens: fracción de palabras del input que aparecen en canonical.
// Detecta nombres abreviados: "axel espinosa" ⊂ "axel alejandro espinosa zapata"
function tokenOverlap(input, canonical) {
  const it = normName(input).split(" ").filter(Boolean);
  const ct = new Set(normName(canonical).split(" ").filter(Boolean));
  if (!it.length) return 0;
  return it.filter((t) => ct.has(t)).length / it.length;
}

/**
 * Busca el bebé canónico más parecido en el catálogo.
 * Combina token-overlap + Levenshtein para manejar:
 *   - Nombres abreviados: "axel espinosa" → "Axel Alejandro Espinosa Zapata"
 *   - Typos: "Churio" → "Chourio", "Epinosa" → "Espinosa"
 * Requiere que AMBOS campos (bebé + madre) superen el umbral → mínimos falsos positivos.
 * @returns {object|null} entrada canónica de bebes, o null si no hay match seguro
 */
function findCanonical(nombre_bebe, nombre_madre, catalogo) {
  const nb = normName(nombre_bebe);
  const nm = normName(nombre_madre);

  // 1. Coincidencia exacta post-normalización (más rápida, máxima confianza)
  for (const b of catalogo) {
    if (normName(b.nombre_bebe) === nb && normName(b.nombre_madre) === nm)
      return b;
  }

  // 2. Fuzzy: token-overlap + Levenshtein combinados
  let best = null, bestScore = 0;
  for (const b of catalogo) {
    const cb = normName(b.nombre_bebe);
    const cm = normName(b.nombre_madre);
    const bebeScore  = Math.max(tokenOverlap(nb, cb), strSim(nb, cb));
    const madreScore = Math.max(tokenOverlap(nm, cm), strSim(nm, cm));
    // Umbral conservador: ambos campos deben coincidir bien
    if (bebeScore >= 0.75 && madreScore >= 0.65) {
      const score = bebeScore * madreScore;
      if (score > bestScore) { bestScore = score; best = b; }
    }
  }
  return best;
}

// ── CORS ───────────────────────────────────────────────────────────────────────
// En producción restringe a tu dominio de Render.
// En local permite cualquier origen para facilitar el desarrollo.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// ── Middleware ─────────────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json({ limit: "50mb" }));

// ── Health check — requerido por Render para saber que el servicio está vivo ───
app.get("/health", (req, res) => res.json({ status: "ok" }));

// ── Config pública para el frontend ───────────────────────────────────────────
// Este endpoint es público intencionalmente — el frontend lo necesita antes
// de tener sesión para inicializar el cliente Supabase.
app.get("/api/config", (req, res) => {
  res.json({
    url: process.env.SUPABASE_URL,
    anonKey: process.env.SUPABASE_ANON_KEY,
  });
});

// ── Middleware de autenticación para /api/* ────────────────────────────────────
// Verifica el JWT de Supabase en cada request a la API.
// Sin esto, cualquier persona con la URL del servidor puede leer y escribir
// datos sin haber hecho login, porque el servidor usa service_role key.
async function requireAuthAPI(req, res, next) {
  const auth = req.headers.authorization || "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";

  if (!token) {
    return res.status(401).json({ error: "No autorizado — sesión requerida" });
  }

  try {
    // getUser verifica la firma del JWT contra Supabase Auth
    const {
      data: { user },
      error,
    } = await supabase.auth.getUser(token);

    if (error || !user) {
      return res.status(401).json({ error: "Sesión inválida o expirada" });
    }

    req.user = user; // disponible en los handlers si lo necesitan
    next();
  } catch (e) {
    console.error("Error verificando token:", e.message);
    return res.status(500).json({ error: "Error verificando autenticación" });
  }
}

// Aplicar a todas las rutas /api/* excepto /api/config (ya definido arriba)
app.use("/api", requireAuthAPI);

// ── Páginas públicas ───────────────────────────────────────────────────────────
app.get("/login", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

app.get("/bienvenida", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "bienvenida.html"));
});

// ── Dashboard ──────────────────────────────────────────────────────────────────
app.get("/dashboard", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "dashboard.html"));
});

// ── Base de datos de bebés ─────────────────────────────────────────────────────
app.get("/bebes", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "bebes.html"));
});

// ── Importador de historial ────────────────────────────────────────────────────
app.get("/importar", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "importar.html"));
});

// ── Utilidad: lee "Fase" o "Institucion" indistintamente del body ──────────────
// Cuando app.js manda los datos, puede venir como Fase (nuevo) o InstitucionMadre (viejo).
function leerFase(obj) {
  return String(
    obj.Fase ||
      obj.fase ||
      obj.InstitucionMadre ||
      obj.Institucion ||
      obj.institucion ||
      "",
  ).trim();
}

// =============================================================================
//  ENDPOINTS — CATÁLOGO
// =============================================================================

/**
 * GET /api/sheet/:dia
 * CSV con los bebés del día para que app.js cargue la tabla.
 * La columna se llama "Fase" en el CSV para que app.js la reconozca.
 * Si app.js viejo espera "Institucion", devolvemos ambas por compatibilidad.
 */
app.get("/api/sheet/:dia", async (req, res) => {
  const dia = req.params.dia;
  if (!DIAS_VALIDOS.includes(dia))
    return res.status(404).json({ error: `Día no válido: ${dia}` });

  try {
    const { data: asistencias, error: errA } = await supabase
      .from("asistencias")
      .select("bebe_id")
      .eq("dia", dia);
    if (errA) throw errA;

    if (!asistencias || asistencias.length === 0)
      return res.send(
        "Nombre Bebe,Nombre Madre,Fase,Institucion,Programa,Edad\n",
      );

    const ids = asistencias.map((a) => a.bebe_id);
    const { data: bebes, error: errB } = await supabase
      .from("bebes")
      .select("nombre_bebe, nombre_madre, fase, programa, edad")
      .in("id", ids)
      .order("nombre_bebe", { ascending: true });
    if (errB) throw errB;

    // Devolvemos "Fase" e "Institucion" con el mismo valor para compatibilidad
    // con versiones viejas de app.js que busquen cualquiera de las dos.
    const csvLines = [
      "Nombre Bebe,Nombre Madre,Fase,Institucion,Programa,Edad",
      ...bebes.map((b) =>
        [b.nombre_bebe, b.nombre_madre, b.fase, b.fase, b.programa, b.edad]
          .map((v) => {
            const s = String(v ?? "");
            return s.includes(",") ? `"${s}"` : s;
          })
          .join(","),
      ),
    ];

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache");
    res.send(csvLines.join("\n"));
  } catch (err) {
    console.error(`Error GET /api/sheet/${dia}:`, err.message);
    res.status(500).json({ error: "No se pudo cargar el listado" });
  }
});

/**
 * GET /api/bebes
 * Todos los bebés. Devuelve "Fase" e "InstitucionMadre" con el mismo valor
 * para no romper código viejo que espere "InstitucionMadre".
 */
app.get("/api/bebes", async (req, res) => {
  try {
    const { data, error } = await supabase
      .from("bebes")
      .select("nombre_bebe, nombre_madre, fase, programa, edad")
      .order("nombre_bebe", { ascending: true })
      .limit(5000);
    if (error) throw error;

    if (data.length === 5000)
      console.warn(
        "⚠️  GET /api/bebes: límite de 5000 alcanzado — puede haber bebés sin devolver",
      );

    res.json({
      bebes: data.map((b) => ({
        NombreBebe: b.nombre_bebe,
        NombreMadre: b.nombre_madre,
        Fase: b.fase,
        InstitucionMadre: b.fase, // alias para compatibilidad con app.js viejo
        ProgramaMadre: b.programa,
        Edad: b.edad,
      })),
    });
  } catch (err) {
    console.error("Error GET /api/bebes:", err.message);
    res.status(500).json({ error: "No se pudo cargar bebés" });
  }
});

/**
 * GET /api/dias
 * Días disponibles, ordenados Lunes → Viernes.
 */
app.get("/api/dias", async (req, res) => {
  try {
    const { data, error } = await supabase
      .from("asistencias")
      .select("dia")
      .order("dia");
    if (error) throw error;

    const orden = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"];
    const dias = [...new Set(data.map((r) => r.dia))].sort(
      (a, b) => orden.indexOf(a) - orden.indexOf(b),
    );
    res.json({ dias });
  } catch (err) {
    console.error("Error GET /api/dias:", err.message);
    res.status(500).json({ error: "No se pudo obtener los días" });
  }
});

/**
 * POST /api/bebes
 * Agregar o actualizar bebé. Acepta "fase" o "institucion" en el body.
 *
 * Body: { nombre_bebe, nombre_madre, fase, programa, edad, dias: [] }
 *       también acepta: { InstitucionMadre: "..." } en lugar de fase
 */
app.post("/api/bebes", async (req, res) => {
  const { nombre_bebe, nombre_madre, programa, edad, dias } = req.body;
  const fase = leerFase(req.body); // acepta Fase o InstitucionMadre

  if (!nombre_bebe || !String(nombre_bebe).trim())
    return res.status(400).json({ error: "nombre_bebe es obligatorio" });
  if (!nombre_madre || !String(nombre_madre).trim())
    return res.status(400).json({ error: "nombre_madre es obligatorio" });

  try {
    const { data: existing } = await supabase
      .from("bebes")
      .select("id")
      .ilike("nombre_bebe", nombre_bebe.trim())
      .ilike("nombre_madre", nombre_madre.trim())
      .maybeSingle();

    let bebeId;
    if (existing) {
      const { error: errUp } = await supabase
        .from("bebes")
        .update({ nombre_madre, fase, programa, edad })
        .eq("id", existing.id);
      if (errUp) throw errUp;
      bebeId = existing.id;
    } else {
      const { data: inserted, error: errIns } = await supabase
        .from("bebes")
        .insert({
          nombre_bebe: nombre_bebe.trim(),
          nombre_madre,
          fase,
          programa,
          edad,
        })
        .select("id")
        .single();
      if (errIns) throw errIns;
      bebeId = inserted.id;
    }

    if (Array.isArray(dias) && dias.length > 0) {
      const diasValidos = dias.filter((d) => DIAS_VALIDOS.includes(d));
      if (diasValidos.length > 0) {
        const { error: errAs } = await supabase.from("asistencias").upsert(
          diasValidos.map((dia) => ({ bebe_id: bebeId, dia })),
          { onConflict: "bebe_id,dia", ignoreDuplicates: true },
        );
        if (errAs) throw errAs;
      }
    }

    res.json({ ok: true, id: bebeId });
  } catch (err) {
    console.error("Error POST /api/bebes:", err.message);
    res.status(500).json({ error: "No se pudo guardar el bebé" });
  }
});

/**
 * GET /api/asistencia-dias
 * Devuelve un mapa { nombreBebe: [dias] } para el CRUD de bebes.html
 */
app.get("/api/asistencia-dias", async (req, res) => {
  try {
    const { data, error } = await supabase
      .from("asistencias")
      .select("dia, bebes(nombre_bebe)")
      .limit(10000);
    if (error) throw error;

    const diasMap = {};
    data.forEach((row) => {
      const nombre = row.bebes?.nombre_bebe;
      if (!nombre) return;
      if (!diasMap[nombre]) diasMap[nombre] = [];
      if (!diasMap[nombre].includes(row.dia)) diasMap[nombre].push(row.dia);
    });

    res.json({ ok: true, diasMap });
  } catch (err) {
    console.error("Error GET /api/asistencia-dias:", err.message);
    res.status(500).json({ error: "No se pudo obtener los días" });
  }
});

/**
 * PUT /api/bebes/:id
 * Editar bebé existente por ID.
 * Body: { nombre_bebe, nombre_madre, fase, programa, edad, dias: [] }
 */
app.put("/api/bebes/:id", async (req, res) => {
  const { id } = req.params;
  const { nombre_bebe, nombre_madre, programa, edad, dias } = req.body;
  const fase = leerFase(req.body);

  if (!nombre_bebe || !String(nombre_bebe).trim())
    return res.status(400).json({ error: "nombre_bebe es obligatorio" });
  if (!nombre_madre || !String(nombre_madre).trim())
    return res.status(400).json({ error: "nombre_madre es obligatorio" });

  try {
    const { error: errUp } = await supabase
      .from("bebes")
      .update({
        nombre_bebe: nombre_bebe.trim(),
        nombre_madre,
        fase,
        programa,
        edad,
      })
      .eq("id", id);
    if (errUp) throw errUp;

    // Reemplazar días: validar primero, luego borrar y reinsertar
    if (Array.isArray(dias)) {
      const diasValidos = dias.filter((d) => DIAS_VALIDOS.includes(d));

      // Protección: no permitir dejar un bebé sin ningún día asignado.
      // Antes, un array vacío borraba todos los días y no insertaba nada,
      // dejando el bebé huérfano de asistencias sin ningún aviso.
      if (diasValidos.length === 0) {
        return res
          .status(400)
          .json({ error: "Debe asignar al menos un día de asistencia" });
      }

      const { error: errDel } = await supabase
        .from("asistencias")
        .delete()
        .eq("bebe_id", id);
      if (errDel) throw errDel;

      const { error: errAs } = await supabase
        .from("asistencias")
        .insert(diasValidos.map((dia) => ({ bebe_id: id, dia })));
      if (errAs) throw errAs;
    }

    res.json({ ok: true });
  } catch (err) {
    console.error("Error PUT /api/bebes/:id:", err.message);
    res.status(500).json({ error: "No se pudo actualizar el bebé" });
  }
});

/**
 * DELETE /api/bebes/:id
 * Eliminar bebé por ID (cascade elimina sus asistencias por FK).
 */
app.delete("/api/bebes/:id", async (req, res) => {
  const { id } = req.params;
  try {
    const { error } = await supabase.from("bebes").delete().eq("id", id);
    if (error) throw error;
    res.json({ ok: true });
  } catch (err) {
    console.error("Error DELETE /api/bebes/:id:", err.message);
    res.status(500).json({ error: "No se pudo eliminar el bebé" });
  }
});

// =============================================================================
//  ENDPOINTS — ASISTENCIA DIARIA
// =============================================================================

/**
 * GET /api/asistencia/fechas
 * Devuelve todas las combinaciones fecha+dia ya guardadas en Supabase.
 * El importador las usa para detectar duplicados antes de enviar.
 */
app.get("/api/asistencia/fechas", async (req, res) => {
  try {
    const PAGE_SIZE = 1000;
    let allData = [];
    let from = 0;
    let hasMore = true;
    while (hasMore) {
      const { data, error } = await supabase
        .from("registros_asistencia")
        .select("fecha, dia")
        .range(from, from + PAGE_SIZE - 1);
      if (error) throw error;
      allData = allData.concat(data);
      hasMore = data.length === PAGE_SIZE;
      from += PAGE_SIZE;
    }

    // Deduplicar — solo necesitamos combinaciones únicas fecha+dia
    const seen = new Set();
    const fechas = [];
    for (const r of allData) {
      const key = `${r.fecha}|${r.dia}`;
      if (!seen.has(key)) {
        seen.add(key);
        fechas.push({ fecha: r.fecha, dia: r.dia });
      }
    }

    res.json({ ok: true, fechas });
  } catch (err) {
    console.error("Error GET /api/asistencia/fechas:", err.message);
    res.status(500).json({ error: "No se pudo obtener las fechas" });
  }
});

/**
 * POST /api/asistencia/guardar
 * Guarda la asistencia de un día completo.
 * Acepta "Fase" o "InstitucionMadre" en cada registro.
 *
 * Body: { fecha: "YYYY-MM-DD", dia: "Lunes", registros: [...] }
 */
app.post("/api/asistencia/guardar", async (req, res) => {
  const { fecha, registros } = req.body;
  // Normalizar día — quitar tilde de Miércoles para compatibilidad
  const dia = String(req.body.dia || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();

  if (!fecha || !dia)
    return res.status(400).json({ error: "fecha y dia son obligatorios" });
  if (!Array.isArray(registros) || registros.length === 0)
    return res.status(400).json({ error: "registros no puede estar vacío" });
  if (!DIAS_VALIDOS.includes(dia))
    return res.status(400).json({ error: `Día no válido: "${dia}"` });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha))
    return res
      .status(400)
      .json({ error: "fecha debe tener formato YYYY-MM-DD" });

  try {
    // Cargar catálogo canónico una sola vez para canonicalizar todos los nombres.
    const { data: catalogo } = await supabase
      .from("bebes")
      .select("nombre_bebe, nombre_madre");
    const cat = catalogo || [];

    const filas = registros
      .filter((r) => r.NombreBebe && String(r.NombreBebe).trim())
      .map((r) => {
        let nombre_bebe = String(r.NombreBebe || "").trim();
        let nombre_madre = String(r.NombreMadre || "").trim();

        // Buscar nombre canónico con fuzzy matching (typos + abreviaturas + espacios)
        const canonical = findCanonical(nombre_bebe, nombre_madre, cat);
        if (canonical) {
          nombre_bebe = canonical.nombre_bebe;
          nombre_madre = canonical.nombre_madre;
        }

        return {
          nombre_bebe,
          nombre_madre,
          fase: leerFase(r),
          programa: String(
            r.ProgramaMadre || r.Programa || r.programa || "",
          ).trim(),
          edad: String(r.Edad || "").trim(),
          fecha,
          dia,
          asistencia: String(r.Asistencia || "No").trim(),
          ubicacion: String(r.Ubicacion || "").trim(),
          reporte: String(r.Reporte || "No").trim(),
          situacion_especifica: String(r.SituacionEspecifica || "").trim(),
          nota: String(r.Nota || "").trim(),
          extras: String(r.Extras || r.Visitante || "").trim(),
          no_cidi: String(r.NoCidi || "").trim(),
        };
      });

    if (filas.length === 0)
      return res
        .status(400)
        .json({ error: "Ningún registro válido para guardar" });

    // Upsert en lotes de 20 para evitar timeouts y facilitar debug.
    // Se usa upsert SIN ignoreDuplicates para que una re-importación con datos
    // corregidos sí actualice los registros existentes en lugar de ignorarlos.
    const LOTE = 20;
    let guardados = 0;
    let omitidos = 0;
    const erroresFila = [];

    for (let i = 0; i < filas.length; i += LOTE) {
      const lote = filas.slice(i, i + LOTE);

      // Detectar cuáles ya existen para reportarlos al frontend
      const claves = lote.map(
        (r) => `${r.nombre_bebe}||${r.nombre_madre}||${r.fecha}`,
      );
      const { data: yaExisten } = await supabase
        .from("registros_asistencia")
        .select("nombre_bebe, nombre_madre, fecha")
        .in("fecha", [...new Set(lote.map((r) => r.fecha))]);

      const setExistentes = new Set(
        (yaExisten || []).map(
          (r) => `${r.nombre_bebe}||${r.nombre_madre}||${r.fecha}`,
        ),
      );
      omitidos += claves.filter((k) => setExistentes.has(k)).length;

      const { error: errLote } = await supabase
        .from("registros_asistencia")
        .upsert(lote, { onConflict: "nombre_bebe,nombre_madre,fecha" });

      if (errLote) {
        console.error(`⚠️  Error en lote ${i}-${i + LOTE}:`, errLote.message);
        erroresFila.push(errLote.message);
      } else {
        guardados += lote.length;
      }
    }

    console.log(
      `✅  Asistencia guardada: ${guardados}/${filas.length} registros (${omitidos} ya existían) — ${dia} ${fecha}`,
    );
    res.json({
      ok: true,
      guardados: guardados - omitidos,
      total: filas.length,
      omitidos,
      advertencias: erroresFila.length ? erroresFila : undefined,
    });
  } catch (err) {
    console.error("Error POST /api/asistencia/guardar:", err.message);
    res
      .status(500)
      .json({ error: "No se pudo guardar la asistencia: " + err.message });
  }
});

/**
 * GET /api/asistencia
 * Registros para el dashboard con filtros opcionales.
 * Devuelve "Fase" e "InstitucionMadre" con el mismo valor para compatibilidad.
 *
 *   /api/asistencia?fecha=2026-03-16
 *   /api/asistencia?desde=2026-03-01&hasta=2026-03-31
 *   /api/asistencia?dia=Lunes
 */
app.get("/api/asistencia", async (req, res) => {
  try {
    const { fecha, desde, hasta, dia } = req.query;

    // PostgREST limita a 1000 filas por request — paginamos hasta traer todo
    const PAGE_SIZE = 1000;
    let allData = [];
    let from = 0;
    let hasMore = true;

    while (hasMore) {
      let q = supabase
        .from("registros_asistencia")
        .select("*")
        .order("fecha", { ascending: true })
        .order("nombre_bebe", { ascending: true })
        .range(from, from + PAGE_SIZE - 1);

      if (fecha) q = q.eq("fecha", fecha);
      if (dia) q = q.eq("dia", dia);
      if (desde) q = q.gte("fecha", desde);
      if (hasta) q = q.lte("fecha", hasta);

      const { data, error } = await q;
      if (error) throw error;

      allData = allData.concat(data);
      hasMore = data.length === PAGE_SIZE;
      from += PAGE_SIZE;
    }

    // Canonicalizar nombres en memoria — la BD no se toca, pero el dashboard
    // recibe nombres unificados y el TOTAL ÚNICOS refleja la realidad.
    const { data: catalogo } = await supabase
      .from("bebes")
      .select("nombre_bebe, nombre_madre");
    const cat = catalogo || [];

    const registros = allData.map((r) => {
      const canonical = findCanonical(r.nombre_bebe, r.nombre_madre, cat);
      return ({
      NombreBebe: canonical ? canonical.nombre_bebe : r.nombre_bebe,
      NombreMadre: canonical ? canonical.nombre_madre : r.nombre_madre,
      Fase: r.fase,
      InstitucionMadre: r.fase, // alias para compatibilidad con dashboard viejo
      ProgramaMadre: r.programa,
      Edad: r.edad,
      Fecha: r.fecha,
      Dia: r.dia,
      Asistencia: r.asistencia,
      Ubicacion: r.ubicacion,
      Reporte: r.reporte,
      SituacionEspecifica: r.situacion_especifica,
      Nota: r.nota,
      Extras: r.extras,
      Visitante: r.extras, // alias para compatibilidad
      NoCidi: r.no_cidi,
      });
    });

    res.json({ ok: true, total: registros.length, registros });
  } catch (err) {
    console.error("Error GET /api/asistencia:", err.message);
    res.status(500).json({ error: "No se pudo obtener la asistencia" });
  }
});

/**
 * DELETE /api/asistencia/dia
 * Elimina todos los registros de un día específico.
 * Solo accesible para admin y coordinadora.
 * Body: { fecha: "YYYY-MM-DD", dia: "Lunes" }
 */
app.delete("/api/asistencia/dia", async (req, res) => {
  try {
    const { fecha, dia } = req.body;

    if (!fecha || !dia) {
      return res.status(400).json({ error: "Se requieren fecha y dia" });
    }

    // Verificar que el usuario es admin o coordinadora
    const { data: usuario } = await supabase
      .from("usuarios")
      .select("rol")
      .eq("id", req.user.id)
      .single();

    if (!usuario || !["admin", "coordinadora"].includes(usuario.rol)) {
      return res
        .status(403)
        .json({ error: "Solo admin y coordinadora pueden eliminar registros" });
    }

    const { data, error } = await supabase
      .from("registros_asistencia")
      .delete()
      .eq("fecha", fecha)
      .eq("dia", dia)
      .select("id");

    if (error) throw error;

    const eliminados = data?.length ?? 0;

    if (eliminados === 0) {
      return res
        .status(404)
        .json({ error: `No se encontraron registros para ${dia} ${fecha}` });
    }

    console.log(
      `🗑️  Eliminados ${eliminados} registros de ${dia} ${fecha} por ${usuario.rol}`,
    );
    res.json({ ok: true, eliminados });
  } catch (err) {
    console.error("Error DELETE /api/asistencia/dia:", err.message);
    res.status(500).json({ error: "Error eliminando registros" });
  }
});

/**
 * GET /api/exportar
 * Genera o previsualiza un Excel de registros_asistencia filtrado por rango de fechas.
 * Solo accesible para admin y coordinadora (verificado por rol en tabla usuarios).
 *
 * Query params:
 *   desde      YYYY-MM-DD  (requerido)
 *   hasta      YYYY-MM-DD  (requerido)
 *   programa   string      (opcional)
 *   fase       string      (opcional)
 *   dia        string      (opcional)
 *   preview    "true"      → devuelve JSON con primeras 15 filas + total
 */
app.get("/api/exportar", async (req, res) => {
  try {
    const { desde, hasta, programa, fase, dia, preview } = req.query;

    if (!desde || !hasta) {
      return res
        .status(400)
        .json({ error: "Los parámetros desde y hasta son requeridos" });
    }

    // Verificar que el usuario es admin o coordinadora
    const { data: usuario } = await supabase
      .from("usuarios")
      .select("rol")
      .eq("id", req.user.id)
      .single();

    if (!usuario || !["admin", "coordinadora"].includes(usuario.rol)) {
      return res
        .status(403)
        .json({ error: "Solo admin y coordinadora pueden exportar registros" });
    }

    // Construir query con filtros opcionales — paginada para superar el cap de PostgREST
    const PAGE_SIZE = 1000;
    let allData = [];
    let from = 0;
    let hasMore = true;
    while (hasMore) {
      let q = supabase
        .from("registros_asistencia")
        .select("*")
        .gte("fecha", desde)
        .lte("fecha", hasta)
        .order("fecha", { ascending: true })
        .order("dia", { ascending: true })
        .order("nombre_bebe", { ascending: true })
        .range(from, from + PAGE_SIZE - 1);
      if (programa) q = q.eq("programa", programa);
      if (fase) q = q.eq("fase", fase);
      if (dia) q = q.eq("dia", dia);
      const { data, error } = await q;
      if (error) throw error;
      allData = allData.concat(data);
      hasMore = data.length === PAGE_SIZE;
      from += PAGE_SIZE;
    }
    const data = allData;

    // Mapear columnas al formato exacto del Excel de las profesoras
    const mapearFila = (r) => ({
      Fecha: r.fecha,
      Dia: r.dia,
      "Nombre Bebé": r.nombre_bebe,
      "Nombre Madre": r.nombre_madre,
      Institución: r.fase,
      Programa: r.programa,
      "Edad (meses)": r.edad,
      Asistencia: r.asistencia,
      Ubicación: r.ubicacion,
      Reporte: r.reporte,
      "Situación Específica": r.situacion_especifica,
      Nota: r.nota,
      Extras: r.extras,
      "No CIDI": r.no_cidi,
    });

    // Modo preview — devuelve JSON con primeras 15 filas y total
    if (preview === "true") {
      return res.json({
        total: data.length,
        filas: data.slice(0, 15).map(mapearFila),
      });
    }

    // Modo descarga — genera Excel con mismo formato que las profesoras
    const filas = data.map(mapearFila);
    const ws = XLSX.utils.json_to_sheet(filas);

    ws["!cols"] = [
      { wch: 12 }, // Fecha
      { wch: 11 }, // Dia
      { wch: 30 }, // Nombre Bebé
      { wch: 30 }, // Nombre Madre
      { wch: 9 }, // Institución
      { wch: 24 }, // Programa
      { wch: 13 }, // Edad (meses)
      { wch: 11 }, // Asistencia
      { wch: 11 }, // Ubicación
      { wch: 9 }, // Reporte
      { wch: 26 }, // Situación Específica
      { wch: 20 }, // Nota
      { wch: 8 }, // Extras
      { wch: 8 }, // No CIDI
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Asistencia");

    const nombreArchivo = `juanfe_asistencia_${desde}_${hasta}.xlsx`;
    const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nombreArchivo}"`,
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.send(buffer);

    console.log(
      `✅ Exportado: ${data.length} registros ${desde} → ${hasta} por ${usuario.rol}`,
    );
  } catch (err) {
    console.error("Error GET /api/exportar:", err.message);
    res.status(500).json({ error: "Error generando exportación" });
  }
});

// ── 404 explícito para /api/* — evita que el SPA fallback devuelva HTML ────────
// Sin esto, un fetch a una ruta mal escrita (/api/assistencia) recibe index.html
// con status 200, y el .json() del cliente explota con un error críptico.
app.all("/api/*", (req, res) => {
  res
    .status(404)
    .json({ error: `Ruta no encontrada: ${req.method} ${req.path}` });
});

// ── Fallback SPA — solo para rutas que no sean /api ───────────────────────────
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () =>
  console.log(`\n🚀  Servidor CIDI en http://localhost:${PORT}`),
);
