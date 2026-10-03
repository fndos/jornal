// Jornal en el navegador. El algoritmo es el mismo motor.py del escritorio:
// Pyodide lo corre tal cual sobre un sistema de archivos virtual, así que el
// archivo del usuario nunca sale de su equipo.

const $ = (id) => document.getElementById(id);
const RAIZ = "/jornal";

let py = null;
let archivoElegido = null;
let ultimo = null;

// La tabla de roles no viaja con el sitio: lleva cédulas y nombres. Cada
// persona la carga una vez y queda guardada en su propio navegador.
const BD = "jornal", ALMACEN = "archivos";

function abrirBD() {
  return new Promise((ok, mal) => {
    const s = indexedDB.open(BD, 1);
    s.onupgradeneeded = () => s.result.createObjectStore(ALMACEN);
    s.onsuccess = () => ok(s.result);
    s.onerror = () => mal(s.error);
  });
}
async function guardarLocal(clave, bytes) {
  const bd = await abrirBD();
  return new Promise((ok, mal) => {
    const t = bd.transaction(ALMACEN, "readwrite");
    t.objectStore(ALMACEN).put(bytes, clave);
    t.oncomplete = ok; t.onerror = () => mal(t.error);
  });
}
async function leerLocal(clave) {
  const bd = await abrirBD();
  return new Promise((ok, mal) => {
    const s = bd.transaction(ALMACEN, "readonly").objectStore(ALMACEN).get(clave);
    s.onsuccess = () => ok(s.result || null);
    s.onerror = () => mal(s.error);
  });
}

// ------------------------------------------------------------------ arranque
// Nunca debe quedarse cargando en silencio: si algo se traba, hay que poder
// ver en que paso fue sin abrir la consola del navegador.
let paso = "iniciando";
const enPaso = (t) => { paso = t; };

function fallo(detalle) {
  const caja = $("arranque");
  if (!caja || caja.dataset.fallo) return;
  caja.dataset.fallo = "1";
  caja.innerHTML =
    `<div><strong style="color:var(--rojo)">No pude arrancar.</strong>` +
    `<div class="sub" style="margin-top:6px">Se quedó en: ${paso}</div>` +
    `<div class="sub" style="margin-top:4px">${detalle}</div>` +
    `<div class="sub" style="margin-top:8px">Recarga con ⌘⇧R. Si sigue igual, ` +
    `prueba en otro navegador.</div></div>`;
}

window.addEventListener("error", (e) => fallo(e.message));
window.addEventListener("unhandledrejection", (e) => fallo(String(e.reason)));
setTimeout(() => {
  if ($("trabajo").hidden) fallo("Pasaron 90 segundos sin terminar de cargar.");
}, 90000);

async function arrancar() {
  try {
    enPaso("descargar Python");
    py = await loadPyodide({ indexURL: "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/" });
    enPaso("instalar openpyxl");
    await py.loadPackage("micropip");
    await py.pyimport("micropip").install(["et-xmlfile", "openpyxl"]);

    enPaso("copiar el motor");
    py.FS.mkdirTree(`${RAIZ}/datos`);
    for (const f of ["motor.py", "almacen.py"]) {
      const r = await fetch(`py/${f}`);
      if (!r.ok) throw new Error(`no encontré py/${f}`);
      py.FS.writeFile(`${RAIZ}/${f}`, new Uint8Array(await r.arrayBuffer()));
    }

    // Tabla de respaldo de roles, por si el formato del archivo no los trae.
    // Cuando sí los trae, el motor usa los del propio archivo.
    // Del sitio sólo viene la configuración, que no tiene datos de nadie.
    enPaso("leer la configuración");
    const rc = await fetch("py/datos/configuracion.xlsx");
    if (rc.ok) py.FS.writeFile(`${RAIZ}/datos/configuracion.xlsx`,
                               new Uint8Array(await rc.arrayBuffer()));
    await montarRoles();

    enPaso("importar el motor");
    await py.runPythonAsync(`
import sys
sys.path.insert(0, "${RAIZ}")
import almacen, motor
`);
    enPaso("leer los ajustes");
    await pintarAjustes();

    $("arranque").hidden = true;
    $("trabajo").hidden = false;
    $("ajustes").hidden = false;
  } catch (e) {
    fallo(String(e.message || e));
    console.error(e);
  }
}

async function montarRoles() {
  try {
    const guardada = await leerLocal("roles");
    if (guardada) {
      py.FS.writeFile(`${RAIZ}/datos/roles.xlsx`, guardada);
      return true;
    }
  } catch (e) {
    console.warn("no pude leer la tabla guardada:", e);
  }
  return false;
}

async function cargarRolesDesde(archivo) {
  const bytes = new Uint8Array(await archivo.arrayBuffer());
  py.FS.writeFile(`${RAIZ}/quincena-anterior.xlsx`, bytes);
  const n = await py.runPythonAsync(`
import almacen
regs = almacen.roles_desde_quincena("${RAIZ}/quincena-anterior.xlsx")
almacen.guardar_roles(regs)
len(regs)
`);
  if (!n) throw new Error("No encontré roles ni licencias en ese archivo.");
  await guardarLocal("roles", py.FS.readFile(`${RAIZ}/datos/roles.xlsx`));
  return n;
}

// Los ajustes se muestran, no se editan: son decisiones de negocio que se
// cambian en el Excel de configuración, no al vuelo desde la web.
async function pintarAjustes() {
  const crudo = await py.runPythonAsync(`
import json, almacen
cfg = almacen.cargar_config()
json.dumps([[almacen.ETIQUETAS.get(c, c),
             ("SÍ" if (cfg[c] if "." not in c else cfg[c.split(".")[0]][c.split(".")[1]])
              else "NO") if t == "si_no"
             else str(cfg[c] if "." not in c else cfg[c.split(".")[0]][c.split(".")[1]])]
            for c, _, t, _ in almacen.PARAMETROS
            if c not in almacen.AVANZADOS and c != "carpeta_de_resultados"])
`);
  const filas = JSON.parse(crudo);
  $("ajustes-tabla").innerHTML = filas
    .map(([k, v]) => `<tr><th>${k}</th><td>${v}</td></tr>`).join("");
}

// -------------------------------------------------------------------- flujo
function elegir(f) {
  archivoElegido = f;
  $("elegido").textContent = f.name;
  $("elegido").hidden = false;
  $("generar").disabled = false;
  $("resultado").hidden = true;
  $("estado").textContent = "";
}

async function generar() {
  $("generar").disabled = true;
  $("estado").className = "estado";
  $("estado").textContent = "Procesando…";
  try {
    const bytes = new Uint8Array(await archivoElegido.arrayBuffer());
    py.FS.writeFile(`${RAIZ}/entrada.xlsx`, bytes);

    const crudo = await py.runPythonAsync(`
import json, almacen, motor
entrada = "${RAIZ}/entrada.xlsx"

fecha = motor.primera_fecha(entrada)
if fecha is None:
    raise ValueError("Ese archivo no tiene fechas de marcación legibles.")

hoja_bio, hoja = motor.identificar_hojas(entrada)
if hoja:
    plantilla, hoja_formato = entrada, hoja
elif almacen.hay_plantilla():
    plantilla, hoja_formato = str(almacen.ARCHIVO_PLANTILLA), None
else:
    raise ValueError("El archivo no trae la hoja del formato y no hay plantilla guardada.")

salida = "${RAIZ}/" + motor.nombre_sugerido(fecha)
res = motor.procesar(entrada, salida, plantilla=plantilla, hoja_bio=hoja_bio,
                     hoja_formato=hoja_formato, cfg=almacen.cargar_config(),
                     colaboradores=almacen.indice(almacen.cargar_roles()))
d0, d1, n = res["cobertura"]
json.dumps({
    "salida": res["salida"], "revision": res["salida_revision"],
    "celdas": sum(res["escritas"].values()),
    "a_mano": sum(v for k, v in res["resumen_incidencias"].items()
                  if k in motor.CATEGORIAS_ACCION),
    "inicio": str(res["inicio"]), "fin": str(res["fin"]),
    "d0": str(d0), "d1": str(d1), "dias": n, "origen": res["origen_formato"],
    "dias_periodo": (res["fin"] - res["inicio"]).days + 1,
    "roles": sum(v for k, v in res["escritas"].items() if k.startswith("rol ")),
})
`);
    ultimo = JSON.parse(crudo);
    mostrar(ultimo);
    $("estado").textContent = "";
  } catch (e) {
    const msg = String(e.message || e).split("\n").filter(Boolean).pop();
    $("estado").className = "estado mal";
    $("estado").textContent = msg;
    console.error(e);
  } finally {
    $("generar").disabled = false;
  }
}

function enlazar(a, ruta) {
  const nombre = ruta.split("/").pop();
  const blob = new Blob([py.FS.readFile(ruta)],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  if (a.dataset.url) URL.revokeObjectURL(a.dataset.url);
  const url = URL.createObjectURL(blob);
  a.href = url; a.download = nombre; a.dataset.url = url;
}

function mostrar(r) {
  $("res-titular").textContent = "Formato generado";
  // Sólo se avisa de lo que cambia lo que hay que hacer: que falten días y que
  // haya casos sin resolver. El recuento fino está en el archivo de revisión.
  const avisos = ["Revísalo antes de entregarlo: hay casos que necesitan tu criterio, " +
                  "listados en el archivo de revisión."];
  if (r.dias < r.dias_periodo) {
    avisos.push(`El biométrico sólo cubre ${r.dias} de los ${r.dias_periodo} días ` +
                `de la quincena, del ${r.d0} al ${r.d1}.`);
  }
  $("res-detalle").textContent = avisos.join(" ");
  // El aviso de los roles sólo aparece cuando hace falta, no como sección fija.
  $("sin-roles").hidden = r.roles > 0;
  enlazar($("baja-formato"), r.salida);
  enlazar($("baja-revision"), r.revision);
  $("resultado").hidden = false;
}

// ------------------------------------------------------------------ eventos
$("buscar").onclick = () => $("archivo").click();
$("archivo").onchange = (e) => e.target.files[0] && elegir(e.target.files[0]);
$("generar").onclick = generar;

const zona = $("zona");
["dragenter", "dragover"].forEach((ev) => zona.addEventListener(ev, (e) => {
  e.preventDefault(); zona.classList.add("encima");
}));
["dragleave", "drop"].forEach((ev) => zona.addEventListener(ev, (e) => {
  e.preventDefault(); zona.classList.remove("encima");
}));
zona.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) elegir(f);
});

$("cargar-roles").onclick = () => $("archivo-roles").click();
$("archivo-roles").onchange = async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  $("estado").className = "estado";
  $("estado").textContent = "Leyendo los roles…";
  try {
    const n = await cargarRolesDesde(f);
    $("sin-roles").hidden = true;
    $("estado").textContent = `${n} personas con rol o licencia, guardadas en este navegador.`;
  } catch (err) {
    $("estado").className = "estado mal";
    $("estado").textContent = String(err.message || err);
  }
};

arrancar();
