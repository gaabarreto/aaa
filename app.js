/* =========================================================================
   DataPrice Vzla - app.js
   Puerto modular de la maqueta original. Estructura:
     1. ESTADO GLOBAL
     2. INGESTA DE DATOS (maestros base + carga manual de Excel)
     3. PERFIL DE ASESOR
     4. NAVEGACIÓN Y CAPTURA (cliente -> categoría -> sublínea -> productos)
     5. FILTROS Y VISUALIZADOR ANALYTICS (con fix Estado/Ciudad)
     6. GRÁFICOS (Chart.js)
     7. EXPORTACIÓN Y ADMINISTRACIÓN
     8. INICIALIZACIÓN
   ========================================================================= */

/* ============================== 1. ESTADO GLOBAL ============================== */
let rolActual = null; // ASESOR/DESARROLLADOR: submodo elegido dentro de la pantalla 0 (legado)
let sesionTipo = null; // ADMIN/ASESOR/ANALYTICS: tipo de sesión del login (para el logo/logout)
let MAESTRO_PRODUCTOS = [];
let MAESTRO_VENDEDORES = [];
let MAESTRO_CLIENTES = [];
let registrosPreciosDB = JSON.parse(localStorage.getItem('registrosPrecios') || '[]');
let asesorInfo = JSON.parse(localStorage.getItem('asesorInfo') || 'null') || { codigo: "", nombre: "No configurado" };
let clienteActual = { cod: "", nombre: "", canal: "", region: "", estado: "", ciudad: "", fecha: "" };
let categoriaSeleccionada = "";
let sublineaSeleccionada = "";
let chartBarrasInstancia = null;
let chartSuraVsCompInstancia = null;
let chartSuraVsCompSizeInstancia = null;
let chartTendenciaInstancia = null;

// Normaliza texto para comparar filtros sin que fallen por tildes,
// mayúsculas/minúsculas o espacios de más (ej: registros viejos guardados
// antes de reordenar la maestra podían quedar con "Pañales " vs "PAÑALES").
function normTxt(v) {
  return (v || "").toString().trim().toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

/* ==================== 1.5 MODAL DE MENSAJES (reemplaza alert/confirm) ====================
   Mismo look & feel que el modal "Guardado Localmente": ícono circular, título,
   texto y botón(es). mostrarAlerta() es async y se resuelve al cerrar; mostrarConfirmacion()
   se resuelve en true/false según el botón que el usuario presione.
   ------------------------------------------------------------------------------------- */
function _configurarModalMensaje({ titulo, texto, icono, colorFondo, colorTexto, textoConfirmar, colorBoton, mostrarCancelar }) {
  document.getElementById("modal-mensaje-titulo").textContent = titulo;
  document.getElementById("modal-mensaje-texto").textContent = texto;
  document.getElementById("modal-mensaje-icono").className = `w-16 h-16 rounded-full flex items-center justify-center mx-auto text-3xl ${colorFondo} ${colorTexto}`;
  document.getElementById("modal-mensaje-icono-fa").className = `fa-solid ${icono}`;
  const btnConfirmar = document.getElementById("modal-mensaje-btn-confirmar");
  btnConfirmar.textContent = textoConfirmar;
  btnConfirmar.className = `w-full text-white font-bold py-3 rounded-xl text-xs shadow transition ${colorBoton}`;
  document.getElementById("modal-mensaje-btn-cancelar").classList.toggle("hidden", !mostrarCancelar);
  document.getElementById("modal-mensaje").classList.remove("hidden");
}

function _cerrarModalMensaje() {
  document.getElementById("modal-mensaje").classList.add("hidden");
}

// Sustituto estilizado de alert(). Variantes: 'info' (azul), 'exito' (verde), 'advertencia' (ámbar), 'error' (rojo).
function mostrarAlerta(texto, variante = "info", titulo = null) {
  const estilos = {
    info: { icono: "fa-circle-info", colorFondo: "bg-blue-100", colorTexto: "text-blue-600", colorBoton: "bg-blue-600 hover:bg-blue-700", tituloDefault: "Aviso" },
    exito: { icono: "fa-circle-check", colorFondo: "bg-emerald-100", colorTexto: "text-emerald-600", colorBoton: "bg-emerald-600 hover:bg-emerald-700", tituloDefault: "Listo" },
    advertencia: { icono: "fa-triangle-exclamation", colorFondo: "bg-amber-100", colorTexto: "text-amber-600", colorBoton: "bg-amber-500 hover:bg-amber-600", tituloDefault: "Atención" },
    error: { icono: "fa-circle-exclamation", colorFondo: "bg-rose-100", colorTexto: "text-rose-600", colorBoton: "bg-rose-600 hover:bg-rose-700", tituloDefault: "Error" },
  };
  const e = estilos[variante] || estilos.info;
  return new Promise(resolve => {
    _configurarModalMensaje({ titulo: titulo || e.tituloDefault, texto, icono: e.icono, colorFondo: e.colorFondo, colorTexto: e.colorTexto, textoConfirmar: "Entendido", colorBoton: e.colorBoton, mostrarCancelar: false });
    document.getElementById("modal-mensaje-btn-confirmar").onclick = () => { _cerrarModalMensaje(); resolve(); };
  });
}

// Sustituto estilizado de confirm(). Resuelve true (confirmó) / false (canceló).
function mostrarConfirmacion(texto, opciones = {}) {
  const variante = opciones.variante || "advertencia";
  const estilos = {
    advertencia: { icono: "fa-triangle-exclamation", colorFondo: "bg-amber-100", colorTexto: "text-amber-600", colorBoton: "bg-amber-500 hover:bg-amber-600" },
    peligro: { icono: "fa-trash-can", colorFondo: "bg-rose-100", colorTexto: "text-rose-600", colorBoton: "bg-rose-600 hover:bg-rose-700" },
  };
  const e = estilos[variante] || estilos.advertencia;
  return new Promise(resolve => {
    _configurarModalMensaje({
      titulo: opciones.titulo || "Confirmar",
      texto, icono: e.icono, colorFondo: e.colorFondo, colorTexto: e.colorTexto,
      textoConfirmar: opciones.textoConfirmar || "Confirmar", colorBoton: e.colorBoton, mostrarCancelar: true,
    });
    document.getElementById("modal-mensaje-btn-confirmar").onclick = () => { _cerrarModalMensaje(); resolve(true); };
    document.getElementById("modal-mensaje-btn-cancelar").onclick = () => { _cerrarModalMensaje(); resolve(false); };
  });
}

/* ============================== 2. INGESTA DE DATOS ==============================
   Ingesta base (fetch de data/*.json, generados desde Query_Customers_Salesman.xlsx
   y Data_Entry_Products.xlsx) + ingesta manual (Excel subido desde Administración).
   Ambas rutas normalizan al MISMO schema, así que Filtros/Gráficos/Captura nunca
   necesitan saber de dónde vinieron los datos.
   ------------------------------------------------------------------------------- */
async function cargarMaestrosBase() {
  const cacheProductos = localStorage.getItem('maestroProductos');
  const cacheClientes = localStorage.getItem('maestroClientes');
  const cacheVendedores = localStorage.getItem('maestroVendedores');

  const [productosFetch, clientesFetch, vendedoresFetch] = await Promise.all([
    cacheProductos ? null : fetch('./data/productos.json').then(r => r.json()),
    cacheClientes ? null : fetch('./data/clientes.json').then(r => r.json()),
    cacheVendedores ? null : fetch('./data/vendedores.json').then(r => r.json()),
  ]);

  MAESTRO_PRODUCTOS = cacheProductos ? JSON.parse(cacheProductos) : productosFetch;
  MAESTRO_CLIENTES = cacheClientes ? JSON.parse(cacheClientes) : clientesFetch;
  MAESTRO_VENDEDORES = cacheVendedores ? JSON.parse(cacheVendedores) : vendedoresFetch;
}

function cargarMaestroClientes(event) {
  const file = event.target.files[0]; if (!file) return; const reader = new FileReader();
  document.getElementById('status-clientes').textContent = "Cargando...";
  reader.onload = async function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (rows.length > 0) {
        MAESTRO_CLIENTES = rows.map(r => ({
          cod: (r["CodigoCliente"] || "").toString().trim(),
          nombre: (r["DescripcionCliente"] || "").toString().trim(),
          region: (r["RegionCli"] || "").toString().trim(),
          estado: (r["DescripEstado"] || "").toString().trim(),
          ciudad: (r["ciudad"] || r["Ciudad"] || "").toString().trim(),
          canal: (r["Canal"] || "").toString().trim(),
          codVendedor: (r["CodVendedor"] || "").toString().trim(),
          nombreVendedor: (r["NombreVendedor"] || "").toString().trim(),
        })).filter(c => c.cod);
        localStorage.setItem('maestroClientes', JSON.stringify(MAESTRO_CLIENTES));

        // El mismo Excel de clientes trae Código/Nombre de cada vendedor
        // (columnas CodVendedor/NombreVendedor); se derivan los asesores
        // únicos de ahí para no depender de subir un segundo archivo.
        const vendedoresPorCod = new Map();
        MAESTRO_CLIENTES.forEach(c => {
          if (c.codVendedor && !vendedoresPorCod.has(c.codVendedor)) {
            vendedoresPorCod.set(c.codVendedor, { cod: c.codVendedor, nombre: c.nombreVendedor, canal: c.canal, region: c.region, zona: "" });
          }
        });
        MAESTRO_VENDEDORES = [...vendedoresPorCod.values()];
        localStorage.setItem('maestroVendedores', JSON.stringify(MAESTRO_VENDEDORES));

        const [syncClientes, syncVendedores] = await Promise.all([
          subirMaestroANube("clientes", MAESTRO_CLIENTES),
          subirMaestroANube("vendedores", MAESTRO_VENDEDORES),
        ]);
        const sufijo = (syncClientes && syncVendedores) ? " y sincronizados a todos los dispositivos." : " (solo en este dispositivo; revisa la configuración de Firebase).";
        document.getElementById('status-clientes').textContent = `¡Éxito! ${MAESTRO_CLIENTES.length} clientes y ${MAESTRO_VENDEDORES.length} asesores cargados${sufijo}`;
      } else {
        await mostrarAlerta("Formato incorrecto o vacío.", "error");
        document.getElementById('status-clientes').textContent = "";
      }
    } catch (err) {
      await mostrarAlerta("Error al leer el archivo: " + err.message, "error");
      document.getElementById('status-clientes').textContent = "";
    }
    event.target.value = "";
  }; reader.readAsArrayBuffer(file);
}

function cargarMaestroProductos(event) {
  const file = event.target.files[0]; if (!file) return; const reader = new FileReader();
  document.getElementById('status-productos').textContent = "Cargando...";
  reader.onload = async function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      if (rows.length > 0) {
        MAESTRO_PRODUCTOS = rows.map(r => ({
          id: (r["ID"] ?? "").toString().trim(),
          producto: (r["Producto"] || "").toString().trim(),
          linea: (r["Linea"] || "").toString().trim(),
          sublinea: (r["Sublinea"] || "").toString().trim(),
          sizeReal: (r["Size Real"] || "").toString().trim(),
          marca: (r["Marca"] || "").toString().trim(),
          gama: (r["Gama"] || "").toString().trim(),
          identificador: (r["Identificador"] || "").toString().trim(),
          size: (r["Size Unificado"] || "").toString().trim(),
          tipo: (r["Tipo"] || "COMPETENCIA").toString().trim(),
          categoria: (r["Categoria"] || "GENERAL").toString().trim(),
        })).filter(p => p.producto);
        localStorage.setItem('maestroProductos', JSON.stringify(MAESTRO_PRODUCTOS));
        const sincronizado = await subirMaestroANube("productos", MAESTRO_PRODUCTOS);
        document.getElementById('status-productos').textContent = `¡Éxito! ${MAESTRO_PRODUCTOS.length} productos cargados${sincronizado ? " y sincronizados a todos los dispositivos." : " (solo en este dispositivo; revisa la configuración de Firebase)."}`;
      } else {
        await mostrarAlerta("Formato incorrecto.", "error");
        document.getElementById('status-productos').textContent = "";
      }
    } catch (err) {
      await mostrarAlerta("Error al leer el archivo: " + err.message, "error");
      document.getElementById('status-productos').textContent = "";
    }
    event.target.value = "";
  }; reader.readAsArrayBuffer(file);
}

/* ============================== 3. PERFIL DE ASESOR ============================== */
function abrirModalAsesor() {
  document.getElementById('input-asesor-codigo').value = asesorInfo.codigo || "";
  document.getElementById('input-asesor-nombre').value = asesorInfo.nombre !== "No configurado" ? asesorInfo.nombre : "";
  document.getElementById('modal-config-asesor').classList.remove('hidden');
}

function actualizarInfoAsesorDOM() {
  document.getElementById('lbl-nombre-asesor').textContent = asesorInfo.nombre;
  document.getElementById('lbl-ruta-asesor').textContent = asesorInfo.codigo || "N/A";
}

function buscarVendedorPorCodigo() {
  const val = document.getElementById('input-asesor-codigo').value.trim().toUpperCase();
  if (val === RUTA_GENERAL_CODIGO) {
    document.getElementById('input-asesor-nombre').value = RUTA_GENERAL_NOMBRE;
    return;
  }
  const match = MAESTRO_VENDEDORES.find(v => v.cod.toUpperCase() === val);
  document.getElementById('input-asesor-nombre').value = match ? match.nombre : "";
}

async function guardarConfigAsesor() {
  const c = document.getElementById('input-asesor-codigo').value.trim().toUpperCase();
  const n = document.getElementById('input-asesor-nombre').value.trim();
  if (c && n) {
    asesorInfo = { codigo: c, nombre: n };
    localStorage.setItem('asesorInfo', JSON.stringify(asesorInfo));
    actualizarInfoAsesorDOM();
    document.getElementById('modal-config-asesor').classList.add('hidden');
    document.getElementById("input-search-cliente").value = "";
    document.getElementById("cod-cliente").value = "";
    document.getElementById("select-region").value = "";
    document.getElementById("select-estado").value = "";
    document.getElementById("select-ciudad").value = "";
  } else {
    await mostrarAlerta("Ingresa un código de vendedor válido y registrado en el Maestro.", "error");
  }
}

/* ===================== 3.5. LOGIN (Asesor por código / Admin / BI) =====================
   Login liviano sin contraseña: no hay backend, así que esto es solo
   enrutamiento de rol, no seguridad real.
     - Código de vendedor (existe en MAESTRO_VENDEDORES) -> entra directo como Asesor.
     - "intventascom"  -> Administrador: acceso a las 3 funciones (Asesor/Analytics/Admin).
     - "bigerensura"   -> Solo visualiza Analytics (BI).
   ------------------------------------------------------------------------------------- */
const USUARIO_ADMIN = "INTVENTASCOM";
const USUARIO_BI = "BIGERENSURA";
const RUTA_GENERAL_CODIGO = "INTCOMERCIAL001"; // código compartido de apoyo: ve la cartera COMPLETA de clientes
const RUTA_GENERAL_NOMBRE = "Asesor de Apoyo (Ruta General)";

function intentarLogin() {
  const input = document.getElementById('login-codigo');
  const errorEl = document.getElementById('login-error');
  const val = input.value.trim().toUpperCase();
  errorEl.classList.add('hidden');
  if (!val) return;

  if (val === USUARIO_ADMIN) {
    localStorage.setItem('sesionActiva', JSON.stringify({ tipo: 'ADMIN' }));
    sesionTipo = 'ADMIN';
    input.value = "";
    goToScreen(0);
    return;
  }

  if (val === USUARIO_BI) {
    localStorage.setItem('sesionActiva', JSON.stringify({ tipo: 'ANALYTICS' }));
    sesionTipo = 'ANALYTICS';
    input.value = "";
    goToScreen(5);
    return;
  }

  if (val === RUTA_GENERAL_CODIGO) {
    asesorInfo = { codigo: RUTA_GENERAL_CODIGO, nombre: RUTA_GENERAL_NOMBRE };
    localStorage.setItem('asesorInfo', JSON.stringify(asesorInfo));
    localStorage.setItem('sesionActiva', JSON.stringify({ tipo: 'ASESOR' }));
    actualizarInfoAsesorDOM();
    sesionTipo = 'ASESOR';
    input.value = "";
    goToScreen(1);
    return;
  }

  const vendedor = MAESTRO_VENDEDORES.find(v => v.cod.toUpperCase() === val);
  if (vendedor) {
    asesorInfo = { codigo: vendedor.cod, nombre: vendedor.nombre };
    localStorage.setItem('asesorInfo', JSON.stringify(asesorInfo));
    localStorage.setItem('sesionActiva', JSON.stringify({ tipo: 'ASESOR' }));
    actualizarInfoAsesorDOM();
    sesionTipo = 'ASESOR';
    input.value = "";
    goToScreen(1);
    return;
  }

  errorEl.textContent = "Código no reconocido. Verifica e intenta de nuevo.";
  errorEl.classList.remove('hidden');
}

function cerrarSesion() {
  localStorage.removeItem('sesionActiva');
  // Limpia también el asesor guardado en este dispositivo: si no se hiciera,
  // entrar luego como Administrador y usar la tarjeta "Modo Asesor" (en vez
  // de loguearse con un código) reutilizaría en silencio al ÚLTIMO asesor
  // configurado en este navegador, atribuyéndole mal las capturas nuevas.
  localStorage.removeItem('asesorInfo');
  asesorInfo = { codigo: "", nombre: "No configurado" };
  actualizarInfoAsesorDOM();
  sesionTipo = null;
  rolActual = null;
  goToScreen('login');
}

// Click en el logo: para Admin vuelve al selector de las 3 funciones; para
// Asesor/Analytics vuelve a su propia pantalla (no cierra sesión).
function irAInicioSesion() {
  if (sesionTipo === 'ADMIN') goToScreen(0);
  else if (sesionTipo === 'ASESOR') goToScreen(1);
  else if (sesionTipo === 'ANALYTICS') goToScreen(5);
  else goToScreen('login');
}

/* ===================== 4. NAVEGACIÓN Y CAPTURA ===================== */
function goToScreen(screenId) {
  ["login", 0, 1, 2, 3, 4, 5, 6].forEach(id => document.getElementById(`screen-${id}`).classList.add("hidden"));
  document.getElementById(`screen-${screenId}`).classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: 'smooth' });
  document.getElementById("btn-cerrar-sesion").classList.toggle("hidden", screenId === "login");

  if (screenId === 5) {
    document.getElementById("panel-cobertura-asesor").classList.toggle("hidden", sesionTipo !== 'ADMIN');
    actualizarVisualizadorAnalytics();
    // Admin y BI ven los datos combinados de todos los asesores (nube en vivo).
    if (sesionTipo === 'ADMIN' || sesionTipo === 'ANALYTICS') suscribirseAnalyticsEnVivo();
  } else {
    desuscribirseAnalyticsEnVivo();
  }

  if (screenId === 6) {
    poblarSelectorImagenProductos();
    poblarSelectorOrdenCategorias();
    renderVistas();
  }
}

function seleccionarModo(modo) {
  rolActual = modo;
  if (modo === 'ASESOR') {
    if (!asesorInfo.codigo) abrirModalAsesor();
    goToScreen(1);
  } else {
    goToScreen(6);
  }
}

function mostrarDropdownClientes() {
  document.getElementById("dropdown-clientes").classList.remove("hidden");
  filtrarClientesDropdown();
}

// Nivel 1 de autocompletado: la búsqueda queda acotada a la cartera del
// asesor activo (mismo codVendedor), replicando la ruta asignada.
function filtrarClientesDropdown() {
  if (!asesorInfo.codigo) {
    document.getElementById("dropdown-clientes").innerHTML = `<li class="px-4 py-2 text-xs text-rose-500 font-bold">Por favor configura tu código de vendedor arriba.</li>`;
    return;
  }
  const term = document.getElementById("input-search-cliente").value.toLowerCase();
  const ul = document.getElementById("dropdown-clientes");
  ul.innerHTML = "";

  // Ruta General: código de apoyo compartido, ve TODA la cartera de clientes
  // sin importar a qué vendedor esté asignada (para cubrir sucursales de otros).
  const esRutaGeneral = asesorInfo.codigo.toUpperCase() === RUTA_GENERAL_CODIGO;
  const matches = MAESTRO_CLIENTES.filter(c =>
    c.nombre.toLowerCase().includes(term) &&
    (esRutaGeneral || c.codVendedor.toUpperCase() === asesorInfo.codigo.toUpperCase())
  );

  if (matches.length === 0) {
    ul.innerHTML = `<li class="px-4 py-2 text-xs text-gray-400">Sin resultados en tu ruta...</li>`;
  } else {
    matches.slice(0, 50).forEach(c => {
      const li = document.createElement("li");
      li.className = "px-4 py-2 hover:bg-blue-50 cursor-pointer text-sm border-b border-gray-50";
      li.innerHTML = `<span class="font-bold text-gray-800">${c.nombre}</span> <span class="text-xs text-gray-400">(${c.region} - ${c.ciudad})</span>`;
      li.onclick = () => seleccionarCliente(c);
      ul.appendChild(li);
    });
  }
}

// Nivel 2 de autocompletado: al elegir un cliente, se completan Código,
// Canal, Región, Estado y Ciudad automáticamente.
function seleccionarCliente(cliente) {
  document.getElementById("input-search-cliente").value = cliente.nombre;
  document.getElementById("cod-cliente").value = cliente.cod;

  let cn = "Tradicional";
  const canalUpper = cliente.canal.toUpperCase();
  if (canalUpper.includes("CLAVE")) cn = "Cuentas Clave";
  else if (canalUpper.includes("ESPECIAL")) cn = "Clientes Especiales";
  document.getElementById("select-canal").value = cn;

  document.getElementById("select-region").value = cliente.region;
  document.getElementById("select-estado").value = cliente.estado;
  document.getElementById("select-ciudad").value = cliente.ciudad;

  document.getElementById("dropdown-clientes").classList.add("hidden");
}

async function iniciarCaptura() {
  const codCliente = document.getElementById("cod-cliente").value.trim();
  // El código solo se llena al elegir un cliente del dropdown (seleccionarCliente).
  // Si está vacío, no se seleccionó ninguno: no se puede continuar.
  if (!codCliente) {
    await mostrarAlerta("Debes buscar y seleccionar un cliente de la lista antes de continuar.", "advertencia");
    document.getElementById("input-search-cliente").focus();
    return;
  }
  clienteActual = {
    cod: codCliente,
    nombre: document.getElementById("input-search-cliente").value.trim(),
    canal: document.getElementById("select-canal").value,
    region: document.getElementById("select-region").value || "N/A",
    estado: document.getElementById("select-estado").value || "",
    ciudad: document.getElementById("select-ciudad").value || "",
    fecha: document.getElementById("fecha-captura").value || new Date().toISOString().split('T')[0],
  };
  renderCategoriasMenu();
  goToScreen(2);
}

// "Cambiar Cliente": vuelve a la pantalla 1 con el buscador y los campos
// autocompletados en blanco, para elegir el siguiente cliente sin arrastrar
// los datos del anterior.
function cambiarCliente() {
  document.getElementById("input-search-cliente").value = "";
  document.getElementById("cod-cliente").value = "";
  document.getElementById("select-region").value = "";
  document.getElementById("select-estado").value = "";
  document.getElementById("select-ciudad").value = "";
  document.getElementById("dropdown-clientes").classList.add("hidden");
  goToScreen(1);
}

const ICONOS_CATEGORIA = {
  "CUIDADO DEL HOGAR": "fa-house-chimney-window text-blue-600 bg-blue-50 border-blue-200",
  "CUIDADO PERSONAL": "fa-pump-soap text-emerald-600 bg-emerald-50 border-emerald-200",
  "CUIDADO DEL BEBÉ": "fa-baby text-blue-600 bg-blue-50 border-blue-200",
};

/* ------------------- Orden de visualización (Categorías/Sublíneas/Productos) -------------------
   El Administrador puede reordenar (arrastrar y soltar) cómo se listan las
   categorías, y dentro de cada categoría sus sublíneas, y dentro de cada
   sublínea sus productos. El orden guardado es solo una LISTA DE NOMBRES;
   cualquier categoría/sublínea/producto nuevo que no esté en esa lista
   (por una carga de maestro posterior) simplemente aparece al final.
   -------------------------------------------------------------------------------------------- */
let ORDEN_CATEGORIAS = [];
let ORDEN_SUBLINEAS = {}; // { [categoria]: [nombres...] }
let ORDEN_PRODUCTOS = {}; // { "categoria|sublinea": [nombres...] }

function ordenarConListaGuardada(nombresDisponibles, ordenGuardado) {
  const disponibles = new Set(nombresDisponibles);
  const conocidos = (ordenGuardado || []).filter(n => disponibles.has(n));
  const nuevos = nombresDisponibles.filter(n => !conocidos.includes(n));
  return [...conocidos, ...nuevos];
}

function obtenerCategoriasOrdenadas() {
  const set = [...new Set(MAESTRO_PRODUCTOS.map(p => p.categoria))].filter(Boolean);
  return ordenarConListaGuardada(set, ORDEN_CATEGORIAS);
}

function obtenerSublineasOrdenadas(categoria) {
  const set = [...new Set(MAESTRO_PRODUCTOS.filter(p => p.categoria === categoria).map(p => p.sublinea))].filter(Boolean);
  return ordenarConListaGuardada(set, ORDEN_SUBLINEAS[categoria]);
}

// Extrae el número de un "Size Real" como "160 UND" o "1.5 LT" para poder
// ordenar de mayor a menor; si no hay número reconocible, va al final.
// Extrae el número de un texto tipo "160 UND" o "1.5 LT" para poder ordenar
// de mayor a menor; si no hay número reconocible, va al final.
function numeroDeTexto(texto) {
  const match = (texto || "").toString().replace(",", ".").match(/[\d.]+/);
  return match ? parseFloat(match[0]) : -Infinity;
}

// Orden por defecto de productos dentro de una sublínea:
// 1) SURA antes que Competencia.
// 2) Dentro de cada grupo, Marca de mayor a menor (Z→A), sin importar la empresa.
// 3) Dentro de la misma marca, Size Unificado de mayor a menor.
function ordenAutomaticoProductos(lista) {
  return [...lista].sort((a, b) => {
    const rangoTipo = (p) => (p.tipo === "SURA" ? 0 : 1);
    const diffTipo = rangoTipo(a) - rangoTipo(b);
    if (diffTipo !== 0) return diffTipo;
    const diffMarca = (b.marca || "").localeCompare(a.marca || "", "es", { sensitivity: "base" });
    if (diffMarca !== 0) return diffMarca;
    return numeroDeTexto(b.size) - numeroDeTexto(a.size);
  });
}

function obtenerProductosOrdenados(categoria, sublinea) {
  const lista = MAESTRO_PRODUCTOS.filter(p => p.categoria === categoria && p.sublinea === sublinea);
  const ordenGuardado = ORDEN_PRODUCTOS[categoria + "|" + sublinea];
  // Si el Administrador ya reordenó esta sublínea a mano, se respeta ese
  // orden; si no, el orden por defecto es SURA/Competencia + Size Real.
  if (ordenGuardado && ordenGuardado.length) {
    const nombresOrdenados = ordenarConListaGuardada(lista.map(p => p.producto), ordenGuardado);
    const porNombre = new Map(lista.map(p => [p.producto, p]));
    return nombresOrdenados.map(n => porNombre.get(n)).filter(Boolean);
  }
  return ordenAutomaticoProductos(lista);
}

function renderCategoriasMenu() {
  document.getElementById("resumen-nombre-s2").textContent = clienteActual.nombre;
  document.getElementById("resumen-info-s2").textContent = `Cód: ${clienteActual.cod} | ${clienteActual.ciudad} (${clienteActual.region}) | Canal: ${clienteActual.canal}`;
  const categoriasSet = categoriasParaToma();
  const container = document.getElementById("container-categorias");
  container.innerHTML = "";
  categoriasSet.forEach(cat => {
    const sublineasActivas = sublineasParaToma(cat);
    const sublinesCount = sublineasActivas.length;
    const productsCount = sublineasActivas.reduce((n, s) => n + productosParaToma(cat, s).length, 0);
    const styleClass = ICONOS_CATEGORIA[cat] || "fa-layer-group text-slate-600 bg-slate-50 border-slate-200";
    const card = document.createElement("div");
    card.className = "bg-white rounded-2xl p-6 shadow-sm hover:shadow-xl border border-gray-200 hover:border-blue-500 cursor-pointer transition-all group flex flex-col justify-between space-y-4";
    card.onclick = () => seleccionarCategoria(cat);
    card.innerHTML = `<div class="flex items-center justify-between"><div class="w-12 h-12 rounded-xl flex items-center justify-center text-xl border ${styleClass}"><i class="fa-solid ${styleClass.split(' ')[0]}"></i></div><span class="text-[10px] font-bold bg-gray-100 text-gray-600 px-2.5 py-1 rounded-full group-hover:bg-blue-600 group-hover:text-white transition">${sublinesCount} Sublíneas</span></div><div><h3 class="text-xl font-black text-gray-900 group-hover:text-blue-600 transition">${cat}</h3><p class="text-xs text-gray-500 mt-1">${productsCount} productos en catálogo.</p></div>`;
    container.appendChild(card);
  });
}

function seleccionarCategoria(cat) {
  categoriaSeleccionada = cat;
  document.getElementById("titulo-categoria-activa").textContent = cat;
  document.getElementById("resumen-nombre-s3").textContent = clienteActual.nombre;
  document.getElementById("resumen-info-s3").textContent = `Cód: ${clienteActual.cod} | ${clienteActual.region}`;
  const sublineasSet = sublineasParaToma(cat);
  const container = document.getElementById("container-sublineas");
  container.innerHTML = "";
  sublineasSet.forEach(sub => {
    const prodCount = productosParaToma(cat, sub).length;
    const item = document.createElement("div");
    item.className = "bg-white p-4 rounded-xl shadow-sm hover:shadow-md border border-gray-200 hover:border-blue-500 cursor-pointer transition flex items-center justify-between group";
    item.onclick = () => seleccionarSublinea(sub);
    item.innerHTML = `<div class="flex items-center space-x-3"><div class="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs group-hover:bg-blue-600 group-hover:text-white transition"><i class="fa-solid fa-folder"></i></div><div><h4 class="font-bold text-gray-800 text-sm group-hover:text-blue-600 transition">${sub}</h4><span class="text-[10px] text-gray-400 font-medium">${prodCount} productos en lista</span></div></div><i class="fa-solid fa-chevron-right text-gray-300 text-xs group-hover:text-blue-600 group-hover:translate-x-0.5 transition"></i>`;
    container.appendChild(item);
  });
  goToScreen(3);
}

function seleccionarSublinea(sub) {
  sublineaSeleccionada = sub;
  document.getElementById("bc-categoria").textContent = categoriaSeleccionada;
  document.getElementById("bc-sublinea").textContent = sublineaSeleccionada;
  document.getElementById("titulo-sublinea-activa").textContent = sublineaSeleccionada;
  const productos = productosParaToma(categoriaSeleccionada, sublineaSeleccionada);
  document.getElementById("count-productos-badge").textContent = `${productos.length} Productos`;
  const tbody = document.getElementById("lista-productos-body");
  tbody.innerHTML = "";
  productos.forEach(prod => {
    const existente = registrosPreciosDB.find(r => r.codCliente === clienteActual.cod && r.producto === prod.producto && r.fecha === clienteActual.fecha);
    const valPrecio = existente ? existente.precio.toFixed(2) : "";
    const isSura = prod.tipo === "SURA";
    const slug = slugProducto(prod.producto);
    // Tarjeta (no <tr>): en pantallas chicas queda todo apilado verticalmente
    // (imagen+nombre arriba, precio abajo a lo ancho) para que el asesor no
    // tenga que desplazarse horizontalmente entre el producto y su precio;
    // en pantallas grandes se acomoda en una sola fila.
    const row = document.createElement("div");
    row.className = "flex flex-col sm:flex-row sm:items-center gap-3 p-3 hover:bg-blue-50/50 transition item-producto-row";
    row.setAttribute("data-search", `${prod.producto} ${prod.marca}`.toLowerCase());
    row.innerHTML = `
      <div class="flex items-center gap-3 flex-1 min-w-0">
        <img id="img-prod-${slug}" src="" alt="" onclick="mostrarImagenAmpliada(this.src)" class="w-11 h-11 rounded-lg object-cover border border-gray-200 hidden flex-shrink-0 cursor-zoom-in hover:ring-2 hover:ring-blue-400 transition">
        <div id="img-prod-${slug}-placeholder" class="w-11 h-11 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-300 flex-shrink-0"><i class="fa-solid fa-image text-sm"></i></div>
        <div class="min-w-0">
          <p class="font-bold text-gray-900 text-sm leading-snug">${prod.producto}</p>
          <p class="text-xs text-gray-500 mt-0.5 flex items-center gap-1.5 flex-wrap"><span class="font-semibold">${prod.marca}</span><span class="px-1.5 py-0.5 rounded text-[10px] font-black tracking-wider ${isSura ? 'bg-blue-100 text-blue-800 border border-blue-200' : 'bg-gray-100 text-gray-700 border border-gray-200'}">${prod.size || '-'}</span></p>
        </div>
      </div>
      <div class="w-full sm:w-36 flex-shrink-0">
        <div class="relative rounded-lg shadow-sm">
          <div class="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3"><span class="text-gray-400 font-bold text-sm">$</span></div>
          <input type="text" inputmode="numeric" autocomplete="off" placeholder="0.00" value="${valPrecio}" data-producto="${prod.producto}" data-marca="${prod.marca}" data-tipo="${prod.tipo}" data-size="${prod.size || ''}" onkeydown="manejarTeclaPrecioATM(event)" onpaste="manejarPegadoPrecioATM(event)" class="input-precio-prod w-full border border-gray-300 rounded-lg pl-7 pr-3 py-2.5 sm:py-1.5 text-right font-extrabold text-gray-900 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white text-sm sm:text-xs">
        </div>
      </div>`;
    tbody.appendChild(row);
    cargarImagenProductoEnFila(slug);
  });
  goToScreen(4);
}

function filtrarTablaProductos() {
  const term = document.getElementById("search-producto-input").value.toLowerCase();
  document.querySelectorAll("#lista-productos-body .item-producto-row").forEach(row => {
    row.classList.toggle("hidden", !row.getAttribute("data-search").includes(term));
  });
}

/* ----------------------- Inputs de precio estilo cajero (ATM) -----------------------
   El valor se maneja siempre en centavos: cada dígito nuevo entra por la derecha
   y empuja todo lo demás hacia la izquierda (10.00 -> tecla "5" -> 100.05).
   Backspace hace lo inverso (divide entre 10). No se aceptan letras ni símbolos.
   --------------------------------------------------------------------------------- */
function obtenerCentavosDeInput(input) {
  const digitos = (input.value || "0.00").replace(/[^0-9]/g, "");
  return parseInt(digitos, 10) || 0;
}

const PRECIO_MAXIMO_CENTAVOS = 2000; // $20.00 — tope de negocio para precios en anaquel
const PRECIO_ALERTA_CENTAVOS = 800; // $8.00 — a partir de aquí se pide confirmación al guardar

function formatearCentavos(centavos) {
  return (Math.max(0, centavos) / 100).toFixed(2);
}

// Si al escribir un dígito el valor supera el tope de negocio, se reinicia a
// $0.00 en vez de quedar pegado en el máximo (evita precios "falsos" por
// error de tecleo, ej. seguir tecleando después de llegar a 20.00).
function limitarCentavos(centavos) {
  return centavos > PRECIO_MAXIMO_CENTAVOS ? 0 : centavos;
}

function manejarTeclaPrecioATM(event) {
  const input = event.target;
  const tecla = event.key;

  // Navegación y atajos del navegador: se dejan pasar sin tocar el valor.
  if (["Tab", "ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End", "Enter", "Escape"].includes(tecla)
    || event.ctrlKey || event.metaKey) {
    return;
  }

  event.preventDefault(); // el input nunca se edita "a mano": todo pasa por aquí

  if (tecla === "Backspace" || tecla === "Delete") {
    const nuevosCentavos = Math.floor(obtenerCentavosDeInput(input) / 10);
    input.value = formatearCentavos(nuevosCentavos);
    return;
  }

  if (/^[0-9]$/.test(tecla)) {
    const nuevosCentavos = limitarCentavos(obtenerCentavosDeInput(input) * 10 + Number(tecla));
    input.value = formatearCentavos(nuevosCentavos);
  }
  // cualquier otra tecla (letras, símbolos) queda ignorada por el preventDefault de arriba
}

function manejarPegadoPrecioATM(event) {
  event.preventDefault();
  const texto = (event.clipboardData || window.clipboardData).getData("text");
  const soloDigitos = texto.replace(/[^0-9]/g, "");
  if (!soloDigitos) return;
  let centavos = obtenerCentavosDeInput(event.target);
  for (const digito of soloDigitos) centavos = limitarCentavos(centavos * 10 + Number(digito));
  event.target.value = formatearCentavos(centavos);
}

async function guardarPreciosToma() {
  const inputs = document.querySelectorAll(".input-precio-prod");

  // Primera pasada: solo lee los valores válidos, sin tocar registrosPreciosDB
  // todavía (así un "Cancelar" en la confirmación no deja cambios a medias).
  const candidatos = [];
  inputs.forEach(inp => {
    // Acepta tanto "5.50" como "5,50" (teclado numérico en locale es-VE usa coma)
    const val = parseFloat((inp.value || "").toString().replace(",", "."));
    if (!isNaN(val) && val > 0) {
      candidatos.push({
        prodNombre: inp.getAttribute("data-producto"),
        marca: inp.getAttribute("data-marca"),
        tipo: inp.getAttribute("data-tipo"),
        size: inp.getAttribute("data-size") || "",
        val,
      });
    }
  });
  if (candidatos.length === 0) { await mostrarAlerta("Ingresa al menos un precio.", "advertencia"); return; }

  const hayPrecioAlto = candidatos.some(c => c.val > PRECIO_ALERTA_CENTAVOS / 100);
  if (hayPrecioAlto) {
    const continuar = await mostrarConfirmacion("La tabla presenta valores mayores a 8.00$, ¿Estas Seguro que Deseas Guardar?", { titulo: "Precio elevado", textoConfirmar: "Sí, guardar" });
    if (!continuar) return;
  }

  const nuevosRegistros = [];
  candidatos.forEach(({ prodNombre, marca, tipo, size, val }) => {
    registrosPreciosDB = registrosPreciosDB.filter(r => !(r.codCliente === clienteActual.cod && r.producto === prodNombre && r.fecha === clienteActual.fecha));
    const registro = {
      id: Date.now() + Math.floor(Math.random() * 1000),
      fecha: clienteActual.fecha, codCliente: clienteActual.cod, nombreCliente: clienteActual.nombre,
      rutaAsesor: asesorInfo.codigo || "",
      canal: clienteActual.canal, region: clienteActual.region, estado: clienteActual.estado, ciudad: clienteActual.ciudad,
      categoria: categoriaSeleccionada, sublinea: sublineaSeleccionada,
      producto: prodNombre, marca: marca, size: size, tipo: tipo, precio: val,
      synced: false, // se marca true cuando llega a la nube (ver sincronizarPendientes)
    };
    registrosPreciosDB.push(registro);
    nuevosRegistros.push(registro);
  });

  localStorage.setItem('registrosPrecios', JSON.stringify(registrosPreciosDB));
  poblarFiltrosVisualizador(); // refresca Región/Estado/Ciudad de Analytics con los datos recién capturados
  document.getElementById("modal-guardar-exito").classList.remove("hidden");

  // Intenta subir a la nube en segundo plano; si falla (sin señal) queda
  // pendiente y se reintenta solo (ver sincronizarPendientes / evento 'online').
  nuevosRegistros.forEach(r => subirRegistroANube(r));
}

/* ============ 5. FILTROS Y VISUALIZADOR ANALYTICS ============
   FIX del bug de Estado/Ciudad: la maqueta original declaraba los <select>
   filter-estado / filter-ciudad pero nunca los poblaba ni los leía en
   obtenerDataFiltrada(). Aquí se agregan onFilterRegionChange/EstadoChange
   para encadenar Región -> Estado -> Ciudad, y ambos campos se incluyen en
   el filtro final.
   --------------------------------------------------------------------- */
function poblarFiltrosVisualizador() {
  const regSelect = document.getElementById("filter-region");
  regSelect.innerHTML = '<option value="">Todas</option>';
  [...new Set(registrosPreciosDB.map(r => r.region))].filter(Boolean).sort().forEach(r => regSelect.innerHTML += `<option value="${r}">${r}</option>`);

  const catSelect = document.getElementById("filter-categoria");
  catSelect.innerHTML = '<option value="">Todas</option>';
  [...new Set(MAESTRO_PRODUCTOS.map(p => p.categoria))].filter(Boolean).sort().forEach(c => catSelect.innerHTML += `<option value="${c}">${c}</option>`);

  document.getElementById("filter-canal").innerHTML = `<option value="">Todos</option><option value="Cuentas Clave">Cuentas Claves</option><option value="Clientes Especiales">Clientes Especiales</option><option value="Tradicional">Canal Tradicional</option>`;

  onFilterRegionChange();
  onFilterCategoriaChange();
}

function onFilterRegionChange() {
  const regVal = document.getElementById("filter-region").value;
  const base = regVal ? registrosPreciosDB.filter(r => r.region === regVal) : registrosPreciosDB;
  const estSelect = document.getElementById("filter-estado");
  const prevEstado = estSelect.value;
  estSelect.innerHTML = '<option value="">Todos</option>' +
    [...new Set(base.map(r => r.estado))].filter(Boolean).sort().map(e => `<option value="${e}">${e}</option>`).join('');
  if ([...estSelect.options].some(o => o.value === prevEstado)) estSelect.value = prevEstado;
  onFilterEstadoChange();
}

function onFilterEstadoChange() {
  const regVal = document.getElementById("filter-region").value;
  const estVal = document.getElementById("filter-estado").value;
  let base = registrosPreciosDB;
  if (regVal) base = base.filter(r => r.region === regVal);
  if (estVal) base = base.filter(r => r.estado === estVal);
  const ciuSelect = document.getElementById("filter-ciudad");
  const prevCiudad = ciuSelect.value;
  ciuSelect.innerHTML = '<option value="">Todas</option>' +
    [...new Set(base.map(r => r.ciudad))].filter(Boolean).sort().map(c => `<option value="${c}">${c}</option>`).join('');
  if ([...ciuSelect.options].some(o => o.value === prevCiudad)) ciuSelect.value = prevCiudad;
  aplicarFiltrosDashboard();
}

function onFilterCategoriaChange() {
  const catVal = document.getElementById("filter-categoria").value;
  const subSelect = document.getElementById("filter-sublinea");
  subSelect.innerHTML = '<option value="">Todas</option>';
  const sublineas = catVal
    ? [...new Set(MAESTRO_PRODUCTOS.filter(p => normTxt(p.categoria) === normTxt(catVal)).map(p => p.sublinea))]
    : [...new Set(MAESTRO_PRODUCTOS.map(p => p.sublinea))];
  sublineas.filter(Boolean).sort().forEach(s => subSelect.innerHTML += `<option value="${s}">${s}</option>`);
  onFilterSublineaChange();
}

// El filtro de Size Unificado se acota a la Categoría/Sublínea elegidas
// (igual que Estado se acota a la Región elegida).
function onFilterSublineaChange() {
  const catVal = document.getElementById("filter-categoria").value;
  const subVal = document.getElementById("filter-sublinea").value;
  let productos = MAESTRO_PRODUCTOS;
  if (catVal) productos = productos.filter(p => normTxt(p.categoria) === normTxt(catVal));
  if (subVal) productos = productos.filter(p => normTxt(p.sublinea) === normTxt(subVal));

  const sizeSelect = document.getElementById("filter-size");
  const prevSize = sizeSelect.value;
  sizeSelect.innerHTML = '<option value="">Todos</option>' +
    [...new Set(productos.map(p => p.size))].filter(Boolean).sort().map(s => `<option value="${s}">${s}</option>`).join('');
  if ([...sizeSelect.options].some(o => o.value === prevSize)) sizeSelect.value = prevSize;

  aplicarFiltrosDashboard();
}

function limpiarFiltrosVisualizador() {
  ["region", "estado", "ciudad", "canal", "categoria", "sublinea", "size", "fecha-inicio", "fecha-fin"].forEach(f => document.getElementById("filter-" + f).value = "");
  poblarFiltrosVisualizador();
}

function obtenerDataFiltrada() {
  const reg = document.getElementById("filter-region").value;
  const est = document.getElementById("filter-estado").value;
  const ciu = document.getElementById("filter-ciudad").value;
  const can = document.getElementById("filter-canal").value;
  const cat = document.getElementById("filter-categoria").value;
  const sub = document.getElementById("filter-sublinea").value;
  const size = document.getElementById("filter-size").value;
  const fInicio = document.getElementById("filter-fecha-inicio").value;
  const fFin = document.getElementById("filter-fecha-fin").value;
  return registrosPreciosDB.filter(r => {
    if (reg && r.region !== reg) return false;
    if (est && r.estado !== est) return false;
    if (ciu && r.ciudad !== ciu) return false;
    if (can && r.canal !== can) return false;
    if (cat && normTxt(r.categoria) !== normTxt(cat)) return false;
    if (sub && normTxt(r.sublinea) !== normTxt(sub)) return false;
    if (size && normTxt(r.size) !== normTxt(size)) return false;
    if (fInicio && r.fecha < fInicio) return false;
    if (fFin && r.fecha > fFin) return false;
    return true;
  });
}

function aplicarFiltrosDashboard() { actualizarVisualizadorAnalytics(); }

function actualizarVisualizadorAnalytics() {
  const dataFiltrada = obtenerDataFiltrada();
  if (dataFiltrada.length === 0) {
    ["min", "max", "avg", "sura-avg", "comp-avg"].forEach(id => {
      const el = document.getElementById(`kpi-${id}-precio`) || document.getElementById(`kpi-${id}`);
      if (el) el.textContent = "$0.00";
    });
    document.getElementById("kpi-count-precios").textContent = "0";
    document.getElementById("kpi-diff-pct").textContent = "Diferencial: 0.0%";
  } else {
    const precios = dataFiltrada.map(r => r.precio);
    const minP = Math.min(...precios), maxP = Math.max(...precios);
    const avgP = precios.reduce((a, b) => a + b, 0) / precios.length;
    document.getElementById("kpi-min-precio").textContent = `$${minP.toFixed(2)}`;
    document.getElementById("kpi-max-precio").textContent = `$${maxP.toFixed(2)}`;
    document.getElementById("kpi-avg-precio").textContent = `$${avgP.toFixed(2)}`;
    document.getElementById("kpi-count-precios").textContent = dataFiltrada.length;

    const suraR = dataFiltrada.filter(r => r.tipo === "SURA");
    const compR = dataFiltrada.filter(r => r.tipo === "COMPETENCIA");
    const suraAvg = suraR.length > 0 ? (suraR.reduce((a, b) => a + b.precio, 0) / suraR.length) : 0;
    const compAvg = compR.length > 0 ? (compR.reduce((a, b) => a + b.precio, 0) / compR.length) : 0;
    document.getElementById("kpi-sura-avg").textContent = `$${suraAvg.toFixed(2)}`;
    document.getElementById("kpi-comp-avg").textContent = `$${compAvg.toFixed(2)}`;

    const kpiDiff = document.getElementById("kpi-diff-pct");
    if (compAvg > 0) {
      const diffPct = ((suraAvg - compAvg) / compAvg) * 100;
      const sign = diffPct > 0 ? "+" : "";
      kpiDiff.textContent = `Diff SURA: ${sign}${diffPct.toFixed(1)}%`;
      kpiDiff.className = `text-[11px] font-bold ${diffPct <= 0 ? 'text-emerald-600' : 'text-rose-600'}`;
    } else {
      kpiDiff.textContent = `Diff: N/A`;
    }
  }
  // Los gráficos dependen de Chart.js (CDN externo). Si el CDN fue bloqueado
  // (ad-blocker, firewall corporativo) esto no debe impedir guardar precios
  // ni actualizar los KPIs de arriba.
  if (typeof Chart === "undefined") {
    console.warn("Chart.js no está disponible: los gráficos no se pueden renderizar.");
  } else {
    renderGraficoTendencia(dataFiltrada);
    renderGraficoBarrasOrdenadas(dataFiltrada);
    renderGraficoSuraVsComp(dataFiltrada);
    renderGraficoSuraVsCompSize(dataFiltrada);
  }
  renderMapaCalor(dataFiltrada); // tabla HTML, no depende de Chart.js
  renderTablaDispersionPrecios(dataFiltrada); // tabla HTML, no depende de Chart.js
  if (sesionTipo === 'ADMIN') renderTablaCoberturaAsesor(dataFiltrada); // solo Administrador
}

/* ============================== 6. GRÁFICOS ============================== */
const NOMBRES_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

function renderGraficoTendencia(data) {
  const ctx = document.getElementById("chart-tendencia").getContext("2d");
  if (chartTendenciaInstancia) chartTendenciaInstancia.destroy();

  // Agrupa por mes (YYYY-MM) en vez de por fecha exacta, así la tendencia se
  // lee como un promedio mensual en lugar de un zigzag por cada día capturado.
  const mesesObj = {};
  data.forEach(r => {
    const claveMes = (r.fecha || "").slice(0, 7); // "2026-03-19" -> "2026-03"
    if (!claveMes) return;
    if (!mesesObj[claveMes]) mesesObj[claveMes] = { suraSum: 0, suraC: 0, compSum: 0, compC: 0 };
    if (r.tipo === 'SURA') { mesesObj[claveMes].suraSum += r.precio; mesesObj[claveMes].suraC++; }
    else { mesesObj[claveMes].compSum += r.precio; mesesObj[claveMes].compC++; }
  });
  const sortedMeses = Object.keys(mesesObj).sort();
  const etiquetasMeses = sortedMeses.map(m => {
    const [anio, mes] = m.split("-");
    return `${NOMBRES_MES[parseInt(mes, 10) - 1]} ${anio}`;
  });
  const suraTrend = sortedMeses.map(m => mesesObj[m].suraC > 0 ? (mesesObj[m].suraSum / mesesObj[m].suraC).toFixed(2) : null);
  const compTrend = sortedMeses.map(m => mesesObj[m].compC > 0 ? (mesesObj[m].compSum / mesesObj[m].compC).toFixed(2) : null);

  chartTendenciaInstancia = new Chart(ctx, {
    type: 'line',
    data: {
      labels: etiquetasMeses,
      datasets: [
        { label: 'Promedio SURA ($)', data: suraTrend, borderColor: 'rgb(79, 70, 229)', backgroundColor: 'rgba(79, 70, 229, 0.1)', tension: 0.3, fill: true },
        { label: 'Promedio Competencia ($)', data: compTrend, borderColor: 'rgb(100, 116, 139)', backgroundColor: 'rgba(100, 116, 139, 0.1)', tension: 0.3, fill: true },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: { mode: 'index', intersect: false } }, scales: { y: { beginAtZero: false, ticks: { callback: v => '$' + Number(v).toFixed(2) } } } },
  });
}

function renderGraficoBarrasOrdenadas(data) {
  const ctx = document.getElementById("chart-barras-ordenadas").getContext("2d");
  if (chartBarrasInstancia) chartBarrasInstancia.destroy();
  const prodMap = {};
  data.forEach(r => { if (!prodMap[r.producto]) prodMap[r.producto] = { sum: 0, count: 0 }; prodMap[r.producto].sum += r.precio; prodMap[r.producto].count += 1; });
  let items = Object.keys(prodMap).map(p => ({ producto: p, avg: prodMap[p].sum / prodMap[p].count }));
  items.sort((a, b) => b.avg - a.avg);
  items = items.slice(0, 10);
  // El dataset guarda el nombre COMPLETO (nunca se trunca el dato real); solo
  // el texto pintado sobre el eje X se acorta con un tick callback, así el
  // tooltip siempre puede mostrar el nombre entero, grande y legible.
  chartBarrasInstancia = new Chart(ctx, {
    type: 'bar',
    data: { labels: items.map(i => i.producto), datasets: [{ label: 'Precio Promedio ($)', data: items.map(i => i.avg.toFixed(2)), backgroundColor: 'rgba(37, 99, 235, 0.85)', borderRadius: 4 }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          titleFont: { size: 15, weight: 'bold' },
          bodyFont: { size: 14 },
          padding: 12,
          callbacks: { title: (items) => items[0].label },
        },
      },
      scales: {
        y: { beginAtZero: true, ticks: { callback: v => '$' + Number(v).toFixed(2) } },
        x: {
          ticks: {
            font: { size: 9 }, maxRotation: 45, minRotation: 45,
            callback: function (value) {
              const label = this.getLabelForValue(value);
              return label.length > 20 ? label.substring(0, 20) + '...' : label;
            },
          },
        },
      },
    },
  });
}

function renderGraficoSuraVsComp(data) {
  const ctx = document.getElementById("chart-sura-vs-competencia").getContext("2d");
  if (chartSuraVsCompInstancia) chartSuraVsCompInstancia.destroy();
  const categorias = [...new Set(MAESTRO_PRODUCTOS.map(p => p.categoria))].filter(Boolean);
  const suraAvgs = [], compAvgs = [];
  categorias.forEach(cat => {
    const suraList = data.filter(r => normTxt(r.categoria) === normTxt(cat) && r.tipo === "SURA");
    const compList = data.filter(r => normTxt(r.categoria) === normTxt(cat) && r.tipo === "COMPETENCIA");
    suraAvgs.push(suraList.length ? (suraList.reduce((a, b) => a + b.precio, 0) / suraList.length).toFixed(2) : 0);
    compAvgs.push(compList.length ? (compList.reduce((a, b) => a + b.precio, 0) / compList.length).toFixed(2) : 0);
  });
  chartSuraVsCompInstancia = new Chart(ctx, {
    type: 'bar',
    data: { labels: categorias, datasets: [{ label: 'SURA ($)', data: suraAvgs, backgroundColor: 'rgba(79, 70, 229, 0.85)' }, { label: 'COMPETENCIA ($)', data: compAvgs, backgroundColor: 'rgba(100, 116, 139, 0.75)' }] },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { ticks: { callback: v => '$' + Number(v).toFixed(2) } } } },
  });
}

function renderGraficoSuraVsCompSize(data) {
  const ctx = document.getElementById("chart-sura-vs-competencia-size").getContext("2d");
  if (chartSuraVsCompSizeInstancia) chartSuraVsCompSizeInstancia.destroy();
  const sizes = [...new Set(data.map(r => r.size))].filter(Boolean)
    .sort((a, b) => numeroDeTexto(b) - numeroDeTexto(a));
  const suraAvgs = [], compAvgs = [];
  sizes.forEach(sz => {
    const suraList = data.filter(r => r.size === sz && r.tipo === "SURA");
    const compList = data.filter(r => r.size === sz && r.tipo === "COMPETENCIA");
    suraAvgs.push(suraList.length ? (suraList.reduce((a, b) => a + b.precio, 0) / suraList.length).toFixed(2) : 0);
    compAvgs.push(compList.length ? (compList.reduce((a, b) => a + b.precio, 0) / compList.length).toFixed(2) : 0);
  });
  chartSuraVsCompSizeInstancia = new Chart(ctx, {
    type: 'bar',
    data: { labels: sizes, datasets: [{ label: 'SURA ($)', data: suraAvgs, backgroundColor: 'rgba(37, 99, 235, 0.85)' }, { label: 'COMPETENCIA ($)', data: compAvgs, backgroundColor: 'rgba(100, 116, 139, 0.75)' }] },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { ticks: { callback: v => '$' + Number(v).toFixed(2) } } } },
  });
}

function renderMapaCalor(data) {
  const headerRow = document.getElementById("heatmap-header-row");
  const bodyRows = document.getElementById("heatmap-body-rows");
  const regiones = [...new Set(data.map(r => r.region))].filter(Boolean);
  const categorias = [...new Set(MAESTRO_PRODUCTOS.map(p => p.categoria))].filter(Boolean);
  headerRow.innerHTML = `<th class="py-2.5 px-3 text-left bg-gray-100 font-bold">Región</th>`;
  categorias.forEach(c => headerRow.innerHTML += `<th class="py-2.5 px-3 font-bold bg-gray-100">${c}</th>`);
  bodyRows.innerHTML = "";
  const maxG = Math.max(...data.map(r => r.precio), 1);
  regiones.forEach(reg => {
    let html = `<td class="py-3 px-3 font-bold text-gray-800 text-left bg-gray-50">${reg}</td>`;
    categorias.forEach(cat => {
      const sub = data.filter(r => r.region === reg && normTxt(r.categoria) === normTxt(cat));
      if (!sub.length) {
        html += `<td class="py-3 px-3 text-gray-300 italic">-</td>`;
      } else {
        const avg = sub.reduce((a, b) => a + b.precio, 0) / sub.length;
        const ratio = avg / maxG;
        const bg = ratio > 0.65 ? "bg-rose-100 text-rose-900 border-rose-200" : ratio > 0.35 ? "bg-amber-100 text-amber-900 border-amber-200" : "bg-emerald-100 text-emerald-900 border-emerald-200";
        html += `<td class="py-3 px-3 font-extrabold heatmap-cell rounded-md border ${bg}">$${avg.toFixed(2)}</td>`;
      }
    });
    const tr = document.createElement("tr");
    tr.innerHTML = html;
    bodyRows.appendChild(tr);
  });
}

// Tabla "Dispersión y Control de Precios": filas = Sublínea (expandible) ->
// Producto, valores = Mín/Máx/Promedio de precio. Filtrable por Tipo.
let pivotPreciosExpandido = new Set(); // sublíneas actualmente expandidas

function togglePivotPreciosSublinea(sublinea) {
  if (pivotPreciosExpandido.has(sublinea)) pivotPreciosExpandido.delete(sublinea);
  else pivotPreciosExpandido.add(sublinea);
  actualizarVisualizadorAnalytics();
}

// Agrega { min, max, sum, count } por sublínea y por producto en un solo paso.
function agregarPorSublineaYProducto(dataset) {
  const porSublinea = new Map(); // sublinea -> { agg, productos: Map(producto -> agg) }
  dataset.forEach(r => {
    const sub = r.sublinea || "(Sin sublínea)";
    if (!porSublinea.has(sub)) porSublinea.set(sub, { agg: { min: r.precio, max: r.precio, sum: 0, count: 0 }, productos: new Map() });
    const entrySub = porSublinea.get(sub);
    entrySub.agg.min = Math.min(entrySub.agg.min, r.precio);
    entrySub.agg.max = Math.max(entrySub.agg.max, r.precio);
    entrySub.agg.sum += r.precio; entrySub.agg.count += 1;

    if (!entrySub.productos.has(r.producto)) entrySub.productos.set(r.producto, { min: r.precio, max: r.precio, sum: 0, count: 0 });
    const entryProd = entrySub.productos.get(r.producto);
    entryProd.min = Math.min(entryProd.min, r.precio);
    entryProd.max = Math.max(entryProd.max, r.precio);
    entryProd.sum += r.precio; entryProd.count += 1;
  });
  return porSublinea;
}

function renderTablaDispersionPrecios(data) {
  const modo = document.getElementById("filter-pivot-tipo").value; // SURA | COMPETENCIA | AMBOS
  const header = document.getElementById("pivot-precios-header");
  const body = document.getElementById("pivot-precios-body");
  const celda = v => (v === undefined ? "-" : `$${v.toFixed(2)}`);
  const caret = sub => `<i class="fa-solid ${pivotPreciosExpandido.has(sub) ? "fa-caret-down" : "fa-caret-right"} text-gray-400 mr-1.5 w-3 inline-block"></i>`;

  if (modo === "AMBOS") {
    const suraPorSub = agregarPorSublineaYProducto(data.filter(r => r.tipo === "SURA"));
    const compPorSub = agregarPorSublineaYProducto(data.filter(r => r.tipo === "COMPETENCIA"));
    const sublineas = [...new Set([...suraPorSub.keys(), ...compPorSub.keys()])].sort();

    header.innerHTML = `
      <tr class="bg-slate-100 text-slate-600 text-[10px] font-black uppercase tracking-wider">
        <th rowspan="2" class="py-2 px-3 text-left border-b border-gray-200 align-bottom">Sublínea / Producto</th>
        <th colspan="3" class="py-1.5 px-3 text-center bg-blue-50 text-blue-700 border-b border-l border-gray-200">SURA</th>
        <th colspan="3" class="py-1.5 px-3 text-center bg-slate-200 text-slate-700 border-b border-l border-gray-200">Competencia</th>
      </tr>
      <tr class="bg-slate-50 text-slate-500 text-[10px] font-bold uppercase">
        <th class="py-1.5 px-2 text-right border-l border-gray-200">Mín.</th><th class="py-1.5 px-2 text-right">Máx.</th><th class="py-1.5 px-2 text-right">Prom.</th>
        <th class="py-1.5 px-2 text-right border-l border-gray-200">Mín.</th><th class="py-1.5 px-2 text-right">Máx.</th><th class="py-1.5 px-2 text-right">Prom.</th>
      </tr>`;

    const filaAgg = (nombre, s, c, esPadre) => `<tr class="${esPadre ? "hover:bg-blue-50/50 transition cursor-pointer font-bold text-gray-800" : "bg-gray-50/70 text-gray-600"}"${esPadre ? ` onclick="togglePivotPreciosSublinea('${nombre.replace(/'/g, "\\'")}')"` : ""}>
        <td class="py-2 px-3 text-left ${esPadre ? "" : "pl-8"}">${esPadre ? caret(nombre) : ""}${nombre}</td>
        <td class="py-2 px-2 text-right border-l border-gray-100">${celda(s?.min)}</td>
        <td class="py-2 px-2 text-right">${celda(s?.max)}</td>
        <td class="py-2 px-2 text-right font-black text-blue-700">${celda(s ? s.sum / s.count : undefined)}</td>
        <td class="py-2 px-2 text-right border-l border-gray-100">${celda(c?.min)}</td>
        <td class="py-2 px-2 text-right">${celda(c?.max)}</td>
        <td class="py-2 px-2 text-right font-black text-slate-700">${celda(c ? c.sum / c.count : undefined)}</td>
      </tr>`;

    body.innerHTML = sublineas.length ? sublineas.map(sub => {
      const s = suraPorSub.get(sub), c = compPorSub.get(sub);
      let html = filaAgg(sub, s?.agg, c?.agg, true);
      if (pivotPreciosExpandido.has(sub)) {
        const productos = [...new Set([...(s?.productos.keys() || []), ...(c?.productos.keys() || [])])].sort();
        productos.forEach(p => { html += filaAgg(p, s?.productos.get(p), c?.productos.get(p), false); });
      }
      return html;
    }).join("") : `<tr><td colspan="7" class="py-6 text-center text-gray-400 italic">Sin registros para este tipo.</td></tr>`;
  } else {
    const porSub = agregarPorSublineaYProducto(data.filter(r => r.tipo === modo));
    const sublineas = [...porSub.keys()].sort();

    header.innerHTML = `<tr class="bg-slate-100 text-slate-600 text-[10px] font-black uppercase tracking-wider">
      <th class="py-2.5 px-3 text-left">Sublínea / Producto</th>
      <th class="py-2.5 px-3 text-right">Mín. de Precio ($)</th>
      <th class="py-2.5 px-3 text-right">Máx. de Precio ($)</th>
      <th class="py-2.5 px-3 text-right">Promedio de Precio ($)</th>
    </tr>`;

    const filaSimple = (nombre, e, esPadre) => `<tr class="${esPadre ? "hover:bg-blue-50/50 transition cursor-pointer font-bold text-gray-800" : "bg-gray-50/70 text-gray-600"}"${esPadre ? ` onclick="togglePivotPreciosSublinea('${nombre.replace(/'/g, "\\'")}')"` : ""}>
        <td class="py-2 px-3 text-left ${esPadre ? "" : "pl-8"}">${esPadre ? caret(nombre) : ""}${nombre}</td>
        <td class="py-2 px-3 text-right">${celda(e.min)}</td>
        <td class="py-2 px-3 text-right">${celda(e.max)}</td>
        <td class="py-2 px-3 text-right font-black text-blue-700">${celda(e.sum / e.count)}</td>
      </tr>`;

    body.innerHTML = sublineas.length ? sublineas.map(sub => {
      const entry = porSub.get(sub);
      let html = filaSimple(sub, entry.agg, true);
      if (pivotPreciosExpandido.has(sub)) {
        [...entry.productos.keys()].sort().forEach(p => { html += filaSimple(p, entry.productos.get(p), false); });
      }
      return html;
    }).join("") : `<tr><td colspan="4" class="py-6 text-center text-gray-400 italic">Sin registros para este tipo.</td></tr>`;
  }
}

// Tabla "Cobertura de Levantamiento por Asesor" (solo Administrador): filas =
// Ruta Asesor, columnas = Fecha, valores = clientes ÚNICOS visitados ese día.
// Filtro: Total Asesores / Asesores con Registro / Asesores sin Registro
// (cruza contra el MAESTRO_VENDEDORES completo para detectar quién no capturó nada).
function renderTablaCoberturaAsesor(data) {
  const modo = document.getElementById("filter-cobertura-modo").value; // TODOS | CON_REGISTRO | SIN_REGISTRO
  // Agrupa por mes (YYYY-MM) en vez de por fecha exacta, para leer la
  // cobertura como evolución mensual en lugar de una columna por cada día.
  const meses = [...new Set(data.map(r => (r.fecha || "").slice(0, 7)))].filter(Boolean).sort();
  const etiquetaMes = (m) => { const [anio, mes] = m.split("-"); return `${NOMBRES_MES[parseInt(mes, 10) - 1]} ${anio}`; };

  const porAsesor = new Map(); // ruta -> Map(mes -> cantidad de precios registrados)
  data.forEach(r => {
    const ruta = r.rutaAsesor || "(Sin asignar)";
    const mes = (r.fecha || "").slice(0, 7);
    if (!mes) return;
    if (!porAsesor.has(ruta)) porAsesor.set(ruta, new Map());
    const porMes = porAsesor.get(ruta);
    porMes.set(mes, (porMes.get(mes) || 0) + 1);
  });

  const todosLosAsesores = [...new Set(MAESTRO_VENDEDORES.map(v => v.cod))].filter(Boolean);
  const conRegistro = [...porAsesor.keys()];
  const sinRegistro = todosLosAsesores.filter(cod => !porAsesor.has(cod));

  const rutas = (modo === "CON_REGISTRO" ? conRegistro : modo === "SIN_REGISTRO" ? sinRegistro : [...new Set([...todosLosAsesores, ...conRegistro])]).sort();

  const header = document.getElementById("cobertura-header");
  header.innerHTML = `<th class="py-2.5 px-3 text-left">Ruta Asesor</th>` +
    meses.map(m => `<th class="py-2.5 px-3 text-center">${etiquetaMes(m)}</th>`).join("") +
    `<th class="py-2.5 px-3 text-center bg-slate-200 text-slate-800">Total</th>`;

  const body = document.getElementById("cobertura-body");
  body.innerHTML = rutas.length
    ? rutas.map(ruta => {
      const porMes = porAsesor.get(ruta);
      let totalAsesor = 0;
      let celdas = `<td class="py-2 px-3 font-bold text-gray-800 text-left">${ruta}</td>`;
      meses.forEach(m => {
        const cantidad = porMes && porMes.get(m);
        if (cantidad) totalAsesor += cantidad;
        celdas += `<td class="py-2 px-3 text-center">${cantidad || "-"}</td>`;
      });
      celdas += `<td class="py-2 px-3 text-center bg-slate-50 font-black">${totalAsesor}</td>`;
      return `<tr class="hover:bg-blue-50/50 transition${porMes ? "" : " text-gray-400 italic"}">${celdas}</tr>`;
    }).join("")
    : `<tr><td colspan="${meses.length + 2}" class="py-6 text-center text-gray-400 italic">Sin asesores para este filtro.</td></tr>`;

  let granTotal = 0;
  let totalCeldas = `<td class="py-2.5 px-3 text-left">Total</td>`;
  meses.forEach(m => {
    let totalMes = 0;
    rutas.forEach(ruta => {
      const porMes = porAsesor.get(ruta);
      totalMes += (porMes && porMes.get(m)) || 0;
    });
    granTotal += totalMes;
    totalCeldas += `<td class="py-2.5 px-3 text-center">${totalMes}</td>`;
  });
  totalCeldas += `<td class="py-2.5 px-3 text-center bg-slate-900">${granTotal}</td>`;
  document.getElementById("cobertura-total-row").innerHTML = totalCeldas;
}

/* ===================== 7. EXPORTACIÓN Y ADMINISTRACIÓN ===================== */
async function exportarExcel() {
  const dataFiltrada = obtenerDataFiltrada();
  if (!dataFiltrada.length) { await mostrarAlerta("Sin registros.", "info"); return; }
  // Registros antiguos (capturados antes de guardar "size" por registro) recurren
  // al maestro de productos actual como respaldo para no dejar la columna vacía.
  const sizePorProducto = new Map(MAESTRO_PRODUCTOS.map(p => [p.producto, p.size]));
  const dataExcel = dataFiltrada.map(r => ({
    "Fecha": r.fecha, "Cód. Cliente": r.codCliente, "Cliente": r.nombreCliente, "Ruta Asesor": r.rutaAsesor || "",
    "Canal": r.canal, "Región": r.region, "Estado": r.estado, "Ciudad": r.ciudad,
    "Categoría": r.categoria, "Sublínea": r.sublinea, "Producto": r.producto, "Marca": r.marca,
    "SizeUnificado": r.size || sizePorProducto.get(r.producto) || "",
    "Tipo": r.tipo, "Precio ($)": r.precio,
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataExcel), "BD_Precios");
  XLSX.writeFile(wb, `DataPrice_${new Date().toISOString().split('T')[0]}.xlsx`);
}

/* ===================== 9. SINCRONIZACIÓN CON LA NUBE (Firebase Firestore) =====================
   Todo precio se guarda PRIMERO en localStorage (así el asesor sigue trabajando
   sin señal). Si Firebase está configurado (ver firebase-config.js) y hay
   conexión, cada registro nuevo se sube a la colección "registrosPrecios" de
   Firestore. Administrador y BI leen de ahí (todos los asesores combinados),
   no de su propio localStorage.
   Si no configuras Firebase, la app sigue funcionando 100% local, como antes.
   -------------------------------------------------------------------------------- */
let dbNube = null;
let unsubscribeAnalyticsEnVivo = null;

function inicializarFirebase() {
  if (typeof FIREBASE_CONFIG === "undefined" || !FIREBASE_CONFIG.apiKey || FIREBASE_CONFIG.apiKey.startsWith("TU_")) {
    console.warn("[DataPrice] Firebase no configurado (ver firebase-config.js): Administrador/BI solo verán los datos de ESTE dispositivo.");
    return;
  }
  try {
    firebase.initializeApp(FIREBASE_CONFIG);
    dbNube = firebase.firestore();
  } catch (err) {
    console.error("[DataPrice] Error al inicializar Firebase:", err);
  }
}

async function subirRegistroANube(registro) {
  if (!dbNube) return;
  try {
    await dbNube.collection("registrosPrecios").doc(String(registro.id)).set(registro);
    registro.synced = true;
    localStorage.setItem('registrosPrecios', JSON.stringify(registrosPreciosDB));
  } catch (err) {
    console.warn("[DataPrice] No se pudo sincronizar el registro " + registro.id + " (se reintentará cuando haya conexión):", err.message);
  }
}

async function sincronizarPendientes() {
  if (!dbNube || !navigator.onLine) return;
  const pendientes = registrosPreciosDB.filter(r => !r.synced);
  for (const r of pendientes) {
    await subirRegistroANube(r);
  }
}

// Ante duplicados (mismo cliente+producto+fecha subido más de una vez por
// reintentos), se queda con el registro de "id" más alto -> mismo criterio
// que usa guardarPreciosToma() al sobreescribir localmente.
function deduplicarRegistros(registros) {
  const porClave = new Map();
  registros.forEach(r => {
    const clave = `${r.codCliente}|${r.producto}|${r.fecha}`;
    const actual = porClave.get(clave);
    if (!actual || Number(r.id) > Number(actual.id)) porClave.set(clave, r);
  });
  return [...porClave.values()];
}

// Administrador y BI: escuchan la nube en vivo mientras el Dashboard está
// abierto, así ven las capturas de todos los asesores sin recargar la página.
function suscribirseAnalyticsEnVivo() {
  if (!dbNube) { cargarRegistrosDesdeNubeUnaVez(); return; }
  desuscribirseAnalyticsEnVivo();
  unsubscribeAnalyticsEnVivo = dbNube.collection("registrosPrecios").onSnapshot(
    snapshot => {
      registrosPreciosDB = deduplicarRegistros(snapshot.docs.map(d => d.data()));
      poblarFiltrosVisualizador();
      actualizarVisualizadorAnalytics();
    },
    err => console.error("[DataPrice] Error en la suscripción en vivo a Analytics:", err)
  );
}

function desuscribirseAnalyticsEnVivo() {
  if (unsubscribeAnalyticsEnVivo) { unsubscribeAnalyticsEnVivo(); unsubscribeAnalyticsEnVivo = null; }
}

// Fallback si Firebase no está configurado: no hay nube, así que Admin/BI
// simplemente ven lo que haya en el localStorage de ESTE dispositivo.
function cargarRegistrosDesdeNubeUnaVez() {
  poblarFiltrosVisualizador();
  actualizarVisualizadorAnalytics();
}

/* ================== 10. SINCRONIZACIÓN DE MAESTROS Y ORDEN (Admin) ==================
   Los Excel que sube el Administrador (Clientes, Vendedores, Productos), el
   orden de listados, y las imágenes de referencia de productos, se guardan
   también en Firestore para que TODOS los dispositivos los reciban, no solo
   el que hizo la carga. Los maestros grandes (clientes, productos) se parten
   en "chunks" para no pasar el límite de 1MB por documento de Firestore.
   Cada dispositivo revisa al iniciar un documento liviano de versión; si
   cambió desde la última vez, recién ahí descarga los datos nuevos.
   Si Firebase no está configurado, todo sigue funcionando 100% local: el
   Excel/imagen/orden solo se aplica en ESTE navegador, como ya pasaba antes.
   ------------------------------------------------------------------------------------ */
const TAMANO_CHUNK_MAESTRO = 800;

async function subirMaestroANube(nombreMaestro, datosArray) {
  if (!dbNube) return false;
  try {
    const metaRef = dbNube.collection("maestroMeta").doc(nombreMaestro);
    const metaActual = await metaRef.get();
    const chunksAnteriores = metaActual.exists ? (metaActual.data().chunks || 0) : 0;

    const chunks = [];
    for (let i = 0; i < datosArray.length; i += TAMANO_CHUNK_MAESTRO) chunks.push(datosArray.slice(i, i + TAMANO_CHUNK_MAESTRO));

    await Promise.all(chunks.map((c, i) => dbNube.collection("maestro_" + nombreMaestro).doc("chunk_" + i).set({ items: c })));
    // Borra chunks sobrantes si esta carga tiene menos filas que la anterior.
    const borrados = [];
    for (let i = chunks.length; i < chunksAnteriores; i++) borrados.push(dbNube.collection("maestro_" + nombreMaestro).doc("chunk_" + i).delete());
    await Promise.all(borrados);

    const nuevaVersion = (metaActual.exists ? (metaActual.data().version || 0) : 0) + 1;
    await metaRef.set({ version: nuevaVersion, chunks: chunks.length, total: datosArray.length, actualizadoEn: new Date().toISOString() });
    return true;
  } catch (err) {
    console.error(`[DataPrice] No se pudo sincronizar el maestro "${nombreMaestro}" a la nube:`, err);
    return false;
  }
}

// Revisa si hay una versión más nueva en Firestore que la que ya se tiene
// cacheada en este dispositivo; solo descarga los chunks si cambió.
async function descargarMaestroDeNubeSiCambio(nombreMaestro) {
  if (!dbNube) return null;
  try {
    const metaSnap = await dbNube.collection("maestroMeta").doc(nombreMaestro).get();
    if (!metaSnap.exists) return null;
    const meta = metaSnap.data();
    const versionLocal = parseInt(localStorage.getItem(`maestro_${nombreMaestro}_version`) || "0", 10);
    if (meta.version === versionLocal) return null; // ya está al día

    const chunkDocs = await Promise.all(
      Array.from({ length: meta.chunks }, (_, i) => dbNube.collection("maestro_" + nombreMaestro).doc("chunk_" + i).get())
    );
    const datos = chunkDocs.flatMap(d => (d.exists ? d.data().items : []));
    localStorage.setItem(`maestro_${nombreMaestro}_version`, String(meta.version));
    return datos;
  } catch (err) {
    console.warn(`[DataPrice] No se pudo revisar actualizaciones del maestro "${nombreMaestro}":`, err.message);
    return null;
  }
}

// Se llama una vez al iniciar la app: si Administrador subió algo nuevo desde
// otro dispositivo, actualiza los maestros locales antes de mostrar nada.
async function sincronizarMaestrosDesdeNube() {
  if (!dbNube) return;
  const [nubeProductos, nubeClientes, nubeVendedores] = await Promise.all([
    descargarMaestroDeNubeSiCambio("productos"),
    descargarMaestroDeNubeSiCambio("clientes"),
    descargarMaestroDeNubeSiCambio("vendedores"),
  ]);
  if (nubeProductos) { MAESTRO_PRODUCTOS = nubeProductos; localStorage.setItem('maestroProductos', JSON.stringify(MAESTRO_PRODUCTOS)); }
  if (nubeClientes) { MAESTRO_CLIENTES = nubeClientes; localStorage.setItem('maestroClientes', JSON.stringify(MAESTRO_CLIENTES)); }
  if (nubeVendedores) { MAESTRO_VENDEDORES = nubeVendedores; localStorage.setItem('maestroVendedores', JSON.stringify(MAESTRO_VENDEDORES)); }
}

/* ----- Orden de listados: un solo documento liviano (nombres, no objetos) ----- */
async function cargarOrdenListados() {
  const cache = JSON.parse(localStorage.getItem('ordenListados') || 'null');
  if (cache) { ORDEN_CATEGORIAS = cache.categorias || []; ORDEN_SUBLINEAS = cache.sublineas || {}; ORDEN_PRODUCTOS = cache.productos || {}; }
  if (!dbNube) return;
  try {
    const snap = await dbNube.collection("maestroMeta").doc("orden").get();
    if (snap.exists) {
      const d = snap.data();
      ORDEN_CATEGORIAS = d.categorias || [];
      ORDEN_SUBLINEAS = d.sublineas || {};
      ORDEN_PRODUCTOS = d.productos || {};
      localStorage.setItem('ordenListados', JSON.stringify({ categorias: ORDEN_CATEGORIAS, sublineas: ORDEN_SUBLINEAS, productos: ORDEN_PRODUCTOS }));
    }
  } catch (err) {
    console.warn("[DataPrice] No se pudo cargar el orden de listados desde la nube:", err.message);
  }
}

async function guardarOrdenEnNube() {
  localStorage.setItem('ordenListados', JSON.stringify({ categorias: ORDEN_CATEGORIAS, sublineas: ORDEN_SUBLINEAS, productos: ORDEN_PRODUCTOS }));
  if (!dbNube) return false;
  try {
    await dbNube.collection("maestroMeta").doc("orden").set({
      categorias: ORDEN_CATEGORIAS, sublineas: ORDEN_SUBLINEAS, productos: ORDEN_PRODUCTOS, actualizadoEn: new Date().toISOString(),
    });
    return true;
  } catch (err) {
    console.error("[DataPrice] No se pudo sincronizar el orden de listados a la nube:", err);
    return false;
  }
}

/* ----- Imágenes de referencia de productos -----
   Un índice liviano (solo la lista de nombres de producto que SÍ tienen
   imagen) se descarga una vez al iniciar. La imagen en sí (base64,
   comprimida a ~300px) se descarga bajo demanda, la primera vez que ese
   producto se muestra en la tabla de captura, y se cachea en localStorage
   para no volver a pedirla. */
let IMAGENES_PRODUCTOS_INDICE = new Set();

function slugProducto(nombreProducto) {
  return (nombreProducto || "").toString().trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_").slice(0, 300) || "PRODUCTO";
}

async function cargarIndiceImagenesProductos() {
  const cache = JSON.parse(localStorage.getItem('imagenesProductosIndice') || 'null');
  if (cache) IMAGENES_PRODUCTOS_INDICE = new Set(cache);
  if (!dbNube) return;
  try {
    const snap = await dbNube.collection("maestroMeta").doc("imagenesProductos").get();
    if (snap.exists) {
      IMAGENES_PRODUCTOS_INDICE = new Set(snap.data().slugs || []);
      localStorage.setItem('imagenesProductosIndice', JSON.stringify([...IMAGENES_PRODUCTOS_INDICE]));
    }
  } catch (err) {
    console.warn("[DataPrice] No se pudo cargar el índice de imágenes de productos:", err.message);
  }
}

// Redimensiona/comprime una imagen subida por el Administrador a un JPEG
// pequeño en base64, para que quepa cómodo en un documento de Firestore.
// Muchas fotos de catálogo traen el producto chiquito, rodeado de un margen
// blanco grande. Esto detecta ese margen para recortarlo antes de comprimir,
// usando un "relleno por conexión" (flood fill) que arranca desde el marco
// exterior de la foto: solo se marca como fondo el blanco que está REALMENTE
// conectado al borde. Así, si el producto tiene partes blancas (ej. una tapa
// blanca), esas partes NO se recortan por error -- a diferencia de escanear
// fila por fila, que sí las confundía con el margen y las cortaba.
// Si no encuentra margen que recortar (fondo no blanco, foto ya ajustada,
// etc.) devuelve null y se usa la imagen completa, sin riesgo de romper nada.
function recortarFondoBlanco(canvas, tolerancia = 10) {
  const w = canvas.width, h = canvas.height;
  if (w < 20 || h < 20) return null;
  const ctx = canvas.getContext("2d");
  const { data } = ctx.getImageData(0, 0, w, h);
  const esParecidoAlBlanco = (idxPixel) => {
    const i = idxPixel * 4;
    return data[i] >= 255 - tolerancia && data[i + 1] >= 255 - tolerancia && data[i + 2] >= 255 - tolerancia;
  };

  const esFondo = new Uint8Array(w * h);
  const visitado = new Uint8Array(w * h);
  const colaX = new Int32Array(w * h);
  const colaY = new Int32Array(w * h);
  let colaFin = 0;

  const encolarSiBlanco = (x, y) => {
    const idx = y * w + x;
    if (visitado[idx]) return;
    visitado[idx] = 1;
    if (!esParecidoAlBlanco(idx)) return;
    esFondo[idx] = 1;
    colaX[colaFin] = x; colaY[colaFin] = y; colaFin++;
  };

  // Semillas: todo el marco exterior de la foto.
  for (let x = 0; x < w; x++) { encolarSiBlanco(x, 0); encolarSiBlanco(x, h - 1); }
  for (let y = 0; y < h; y++) { encolarSiBlanco(0, y); encolarSiBlanco(w - 1, y); }

  for (let i = 0; i < colaFin; i++) {
    const x = colaX[i], y = colaY[i];
    if (x > 0) encolarSiBlanco(x - 1, y);
    if (x < w - 1) encolarSiBlanco(x + 1, y);
    if (y > 0) encolarSiBlanco(x, y - 1);
    if (y < h - 1) encolarSiBlanco(x, y + 1);
  }

  let left = w, right = -1, top = h, bottom = -1;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (!esFondo[y * w + x]) {
        if (x < left) left = x;
        if (x > right) right = x;
        if (y < top) top = y;
        if (y > bottom) bottom = y;
      }
    }
  }
  if (right < left || bottom < top) return null; // toda la imagen se detectó como fondo

  // Deja un margen pequeño alrededor del producto (no pegarlo al borde).
  const margenX = Math.round((right - left) * 0.05) + 2;
  const margenY = Math.round((bottom - top) * 0.05) + 2;
  left = Math.max(0, left - margenX);
  right = Math.min(w - 1, right + margenX);
  top = Math.max(0, top - margenY);
  bottom = Math.min(h - 1, bottom + margenY);
  const cropW = right - left + 1, cropH = bottom - top + 1;
  if (cropW < 10 || cropH < 10 || (cropW === w && cropH === h)) return null;
  return { x: left, y: top, w: cropW, h: cropH };
}

function comprimirImagenABase64(archivo, ladoMax = 320, calidad = 0.6) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const lector = new FileReader();
    lector.onerror = reject;
    lector.onload = () => {
      img.onerror = reject;
      img.onload = () => {
        // 1) Dibuja la imagen a un tamaño de trabajo (para detectar el
        // margen blanco sin escanear fotos gigantes).
        const escalaTrabajo = Math.min(1, 640 / Math.max(img.width, img.height));
        const canvasTrabajo = document.createElement("canvas");
        canvasTrabajo.width = Math.round(img.width * escalaTrabajo);
        canvasTrabajo.height = Math.round(img.height * escalaTrabajo);
        const ctxTrabajo = canvasTrabajo.getContext("2d");
        ctxTrabajo.fillStyle = "#FFFFFF";
        ctxTrabajo.fillRect(0, 0, canvasTrabajo.width, canvasTrabajo.height);
        ctxTrabajo.drawImage(img, 0, 0, canvasTrabajo.width, canvasTrabajo.height);

        // 2) Recorta el margen blanco (si hay) para que el producto llene
        // el cuadro, en vez de quedar chiquito en el centro.
        const recorte = recortarFondoBlanco(canvasTrabajo) || { x: 0, y: 0, w: canvasTrabajo.width, h: canvasTrabajo.height };

        // 3) Centra ese recorte dentro de un lienzo CUADRADO (rellenando lo
        // que falte con blanco) y lo redimensiona al tamaño final. Si no se
        // forzara a cuadrado, un recorte muy alargado (ej. un frasco alto)
        // se vería "con zoom"/cortado al mostrarse en las miniaturas
        // cuadradas de la lista (que usan object-cover).
        const ladoRecorte = Math.max(recorte.w, recorte.h);
        const escalaFinal = Math.min(1, ladoMax / ladoRecorte);
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(ladoRecorte * escalaFinal);
        canvas.height = Math.round(ladoRecorte * escalaFinal);
        const ctx = canvas.getContext("2d");
        // El JPEG de salida no soporta transparencia: si la imagen original
        // (ej. un PNG con fondo transparente) no rellena esto primero, el
        // canvas exporta ese fondo como negro. Se rellena de blanco antes.
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        const anchoDestino = Math.round(recorte.w * escalaFinal);
        const altoDestino = Math.round(recorte.h * escalaFinal);
        const offsetX = Math.round((canvas.width - anchoDestino) / 2);
        const offsetY = Math.round((canvas.height - altoDestino) / 2);
        ctx.drawImage(canvasTrabajo, recorte.x, recorte.y, recorte.w, recorte.h, offsetX, offsetY, anchoDestino, altoDestino);
        resolve(canvas.toDataURL("image/jpeg", calidad));
      };
      img.src = lector.result;
    };
    lector.readAsDataURL(archivo);
  });
}

async function subirImagenProducto(nombreProducto, archivo) {
  if (!dbNube) return false;
  await subirImagenProductoSinActualizarIndice(nombreProducto, archivo);
  await dbNube.collection("maestroMeta").doc("imagenesProductos").set({ slugs: [...IMAGENES_PRODUCTOS_INDICE] });
  localStorage.setItem('imagenesProductosIndice', JSON.stringify([...IMAGENES_PRODUCTOS_INDICE]));
  return true;
}

// Sube la imagen y actualiza el índice EN MEMORIA, pero no escribe el
// documento de índice en Firestore todavía. La usa la subida masiva para
// no reescribir ese documento una vez por cada imagen (sería lento con
// muchos archivos); al final del lote se guarda una sola vez.
async function subirImagenProductoSinActualizarIndice(nombreProducto, archivo) {
  const slug = slugProducto(nombreProducto);
  const base64 = await comprimirImagenABase64(archivo);
  await dbNube.collection("productoImagenes").doc(slug).set({ imagenBase64: base64, actualizadoEn: new Date().toISOString() });
  IMAGENES_PRODUCTOS_INDICE.add(slug);
  localStorage.setItem(`imgProducto_${slug}`, base64); // ya que lo tenemos, lo cacheamos de una vez en este dispositivo
  return slug;
}

async function quitarImagenProducto(nombreProducto) {
  if (!dbNube) return false;
  const slug = slugProducto(nombreProducto);
  await dbNube.collection("productoImagenes").doc(slug).delete();
  IMAGENES_PRODUCTOS_INDICE.delete(slug);
  await dbNube.collection("maestroMeta").doc("imagenesProductos").set({ slugs: [...IMAGENES_PRODUCTOS_INDICE] });
  localStorage.setItem('imagenesProductosIndice', JSON.stringify([...IMAGENES_PRODUCTOS_INDICE]));
  localStorage.removeItem(`imgProducto_${slug}`);
  return true;
}

function mostrarImagenAmpliada(base64) {
  if (!base64) return;
  document.getElementById("img-ampliada").src = base64;
  document.getElementById("modal-imagen-ampliada").classList.remove("hidden");
}

function cerrarImagenAmpliada() {
  document.getElementById("modal-imagen-ampliada").classList.add("hidden");
}

// Se llama por cada fila de la tabla de captura (screen-4): si el producto
// tiene imagen, la muestra (desde caché local o, si no, la pide una sola vez
// a Firestore y la cachea para la próxima).
async function cargarImagenProductoEnFila(slug) {
  if (!IMAGENES_PRODUCTOS_INDICE.has(slug)) return;
  const img = document.getElementById(`img-prod-${slug}`);
  const placeholder = document.getElementById(`img-prod-${slug}-placeholder`);
  if (!img) return;

  const mostrar = (base64) => {
    img.src = base64;
    img.classList.remove("hidden");
    if (placeholder) placeholder.classList.add("hidden");
  };

  const cacheado = localStorage.getItem(`imgProducto_${slug}`);
  if (cacheado) { mostrar(cacheado); return; }
  if (!dbNube) return;
  try {
    const snap = await dbNube.collection("productoImagenes").doc(slug).get();
    if (snap.exists) {
      const base64 = snap.data().imagenBase64;
      localStorage.setItem(`imgProducto_${slug}`, base64);
      // La fila pudo haberse vuelto a dibujar mientras esperábamos la red
      // (usuario cambió de sublínea rápido); re-consulta el elemento actual.
      const imgActual = document.getElementById(`img-prod-${slug}`);
      const placeholderActual = document.getElementById(`img-prod-${slug}-placeholder`);
      if (imgActual) { imgActual.src = base64; imgActual.classList.remove("hidden"); }
      if (placeholderActual) placeholderActual.classList.add("hidden");
    }
  } catch (err) {
    console.warn(`[DataPrice] No se pudo cargar la imagen del producto (${slug}):`, err.message);
  }
}

/* ----- Panel de Administrador: imágenes de producto ----- */
function poblarSelectorImagenProductos() {
  const sel = document.getElementById("img-select-producto");
  const productos = MAESTRO_PRODUCTOS.filter(p => p.producto)
    .sort((a, b) => (Number(a.id) || Infinity) - (Number(b.id) || Infinity) || a.producto.localeCompare(b.producto));
  sel.innerHTML = productos.map(p => `<option value="${p.producto}">${p.id ? p.id + " - " : ""}${p.producto}</option>`).join("");
  onCambioSelectorImagenProducto();
}

function onCambioSelectorImagenProducto() {
  const nombre = document.getElementById("img-select-producto").value;
  const slug = slugProducto(nombre);
  const base64 = localStorage.getItem(`imgProducto_${slug}`);
  const img = document.getElementById("img-preview-actual");
  const vacio = document.getElementById("img-preview-vacio");
  const label = document.getElementById("img-preview-label");
  if (base64) {
    img.src = base64; img.classList.remove("hidden"); vacio.classList.add("hidden"); label.textContent = "Imagen actual";
  } else {
    img.classList.add("hidden"); vacio.classList.remove("hidden"); label.textContent = "Sin imagen";
  }
  document.getElementById("status-imagen-producto").textContent = "";
  renderProductosSinImagen();
}

// Desplegable con los productos que todavía no tienen imagen; elegir uno lo
// selecciona arriba para subirle la foto directo.
function renderProductosSinImagen() {
  const sel = document.getElementById("img-select-sin-imagen");
  if (!sel) return;
  const sinImagen = MAESTRO_PRODUCTOS.filter(p => p.producto && !IMAGENES_PRODUCTOS_INDICE.has(slugProducto(p.producto)))
    .sort((a, b) => (Number(a.id) || Infinity) - (Number(b.id) || Infinity) || a.producto.localeCompare(b.producto));
  sel.innerHTML = `<option value="">${sinImagen.length ? `Productos sin imagen (${sinImagen.length})` : "Todos los productos tienen imagen"}</option>` +
    sinImagen.map(p => `<option value="${escHtml(p.producto)}">${p.id ? escHtml(p.id) + " - " : ""}${escHtml(p.producto)}</option>`).join("");
  sel.disabled = !sinImagen.length;
}

function elegirProductoSinImagen() {
  const nombre = document.getElementById("img-select-sin-imagen").value;
  if (!nombre) return;
  document.getElementById("img-select-producto").value = nombre;
  onCambioSelectorImagenProducto();
}

async function onSubirImagenProducto(event) {
  const archivo = event.target.files[0];
  event.target.value = "";
  if (!archivo) return;
  const nombre = document.getElementById("img-select-producto").value;
  if (!nombre) return;
  if (!dbNube) { await mostrarAlerta("Firebase no está configurado: no se puede sincronizar la imagen a todos los dispositivos.", "error"); return; }
  document.getElementById("status-imagen-producto").textContent = "Subiendo...";
  try {
    await subirImagenProducto(nombre, archivo);
    onCambioSelectorImagenProducto();
    document.getElementById("status-imagen-producto").textContent = "¡Imagen subida y sincronizada a todos los dispositivos!";
  } catch (err) {
    document.getElementById("status-imagen-producto").textContent = "";
    await mostrarAlerta("Error al subir la imagen: " + err.message, "error");
  }
}

async function onSubirImagenesMasivo(event) {
  const archivos = [...event.target.files];
  event.target.value = "";
  if (!archivos.length) return;
  if (!dbNube) { await mostrarAlerta("Firebase no está configurado: no se puede sincronizar las imágenes a todos los dispositivos.", "error"); return; }

  const statusEl = document.getElementById("status-imagen-masivo");
  statusEl.textContent = `Subiendo 0 de ${archivos.length}...`;
  try {
    // Sube de a poco (no todo en paralelo) para poder mostrar el progreso y
    // no saturar la conexión si son muchos archivos.
    // Archivo con nombre numérico (ej. "45.jpg", "045.png") -> producto con
    // ese ID de la maestra; si no, se intenta por nombre del producto.
    const porId = new Map(MAESTRO_PRODUCTOS.filter(p => p.id).map(p => [String(Number(p.id)), p]));
    const porSlug = new Map(MAESTRO_PRODUCTOS.map(p => [slugProducto(p.producto), p]));
    const emparejados = [];
    const sinCoincidencia = [];
    archivos.forEach(archivo => {
      const nombreSinExtension = archivo.name.replace(/\.[^.]+$/, "").trim();
      const producto = /^\d+$/.test(nombreSinExtension)
        ? porId.get(String(Number(nombreSinExtension)))
        : porSlug.get(slugProducto(nombreSinExtension));
      if (producto) emparejados.push({ archivo, producto });
      else sinCoincidencia.push(archivo.name);
    });

    for (let i = 0; i < emparejados.length; i++) {
      await subirImagenProductoSinActualizarIndice(emparejados[i].producto.producto, emparejados[i].archivo);
      statusEl.textContent = `Subiendo ${i + 1} de ${emparejados.length}...`;
    }
    if (emparejados.length) {
      await dbNube.collection("maestroMeta").doc("imagenesProductos").set({ slugs: [...IMAGENES_PRODUCTOS_INDICE] });
      localStorage.setItem('imagenesProductosIndice', JSON.stringify([...IMAGENES_PRODUCTOS_INDICE]));
    }

    onCambioSelectorImagenProducto();
    let resumen = `¡Listo! ${emparejados.length} imagen(es) subida(s) y sincronizada(s).`;
    if (sinCoincidencia.length) {
      resumen += ` ${sinCoincidencia.length} sin coincidencia en la maestra (revisa el nombre del archivo): ${sinCoincidencia.slice(0, 5).join(", ")}${sinCoincidencia.length > 5 ? "..." : ""}`;
    }
    statusEl.textContent = resumen;
  } catch (err) {
    statusEl.textContent = "";
    await mostrarAlerta("Error en la subida masiva: " + err.message, "error");
  }
}

async function onQuitarImagenProducto() {
  const nombre = document.getElementById("img-select-producto").value;
  if (!nombre) return;
  if (!dbNube) { await mostrarAlerta("Firebase no está configurado.", "error"); return; }
  const ok = await mostrarConfirmacion(`¿Quitar la imagen de "${nombre}" para todos los dispositivos?`, { variante: "peligro", textoConfirmar: "Sí, quitar" });
  if (!ok) return;
  await quitarImagenProducto(nombre);
  onCambioSelectorImagenProducto();
  document.getElementById("status-imagen-producto").textContent = "Imagen eliminada.";
}

/* ----- Panel de Administrador: orden de listados (drag & drop) ----- */
let ordenNivelActual = null; // {tipo:'categorias'} | {tipo:'sublineas', categoria} | {tipo:'productos', categoria, sublinea}
let ordenListaActual = [];

function poblarSelectorOrdenCategorias() {
  const sel = document.getElementById("orden-select-categoria");
  const cats = obtenerCategoriasOrdenadas();
  sel.innerHTML = `<option value="">-- Ordenar Categorías --</option>` + cats.map(c => `<option value="${c}">${c}</option>`).join("");
  document.getElementById("orden-contenedor-sublinea").classList.add("hidden");
  ordenNivelActual = { tipo: "categorias" };
  ordenListaActual = cats;
  renderOrdenListaActual();
}

function cambiarNivelOrdenCategoria() {
  const cat = document.getElementById("orden-select-categoria").value;
  const contenedorSub = document.getElementById("orden-contenedor-sublinea");
  if (!cat) {
    contenedorSub.classList.add("hidden");
    ordenNivelActual = { tipo: "categorias" };
    ordenListaActual = obtenerCategoriasOrdenadas();
  } else {
    contenedorSub.classList.remove("hidden");
    const subs = obtenerSublineasOrdenadas(cat);
    document.getElementById("orden-select-sublinea").innerHTML = `<option value="">-- Ordenar Sublíneas de "${cat}" --</option>` + subs.map(s => `<option value="${s}">${s}</option>`).join("");
    ordenNivelActual = { tipo: "sublineas", categoria: cat };
    ordenListaActual = subs;
  }
  renderOrdenListaActual();
}

function cambiarNivelOrdenSublinea() {
  const cat = document.getElementById("orden-select-categoria").value;
  const sub = document.getElementById("orden-select-sublinea").value;
  if (!sub) {
    ordenNivelActual = { tipo: "sublineas", categoria: cat };
    ordenListaActual = obtenerSublineasOrdenadas(cat);
  } else {
    ordenNivelActual = { tipo: "productos", categoria: cat, sublinea: sub };
    ordenListaActual = obtenerProductosOrdenados(cat, sub).map(p => p.producto);
  }
  renderOrdenListaActual();
}

function renderOrdenListaActual() {
  const ul = document.getElementById("orden-lista");
  ul.innerHTML = "";
  if (!ordenListaActual.length) {
    ul.innerHTML = `<li class="py-4 text-center text-xs text-gray-400 italic">Nada para ordenar en este nivel.</li>`;
    return;
  }
  let indiceArrastrado = null;
  ordenListaActual.forEach((nombre, idx) => {
    const li = document.createElement("li");
    li.draggable = true;
    li.className = "px-4 py-2.5 flex items-center justify-between bg-white hover:bg-blue-50 text-sm cursor-move transition";
    li.innerHTML = `<span class="flex items-center space-x-3"><i class="fa-solid fa-grip-lines text-gray-400"></i><span class="font-semibold text-gray-800">${nombre}</span></span><span class="text-[10px] text-gray-400 font-bold">#${idx + 1}</span>`;
    li.addEventListener("dragstart", () => { indiceArrastrado = idx; li.classList.add("opacity-40"); });
    li.addEventListener("dragend", () => li.classList.remove("opacity-40"));
    li.addEventListener("dragover", (e) => e.preventDefault());
    li.addEventListener("drop", (e) => {
      e.preventDefault();
      if (indiceArrastrado === null || indiceArrastrado === idx) return;
      const nuevo = [...ordenListaActual];
      const [movido] = nuevo.splice(indiceArrastrado, 1);
      nuevo.splice(idx, 0, movido);
      ordenListaActual = nuevo;
      renderOrdenListaActual();
    });
    ul.appendChild(li);
  });
}

async function guardarOrdenActual() {
  if (!ordenNivelActual) return;
  if (ordenNivelActual.tipo === "categorias") ORDEN_CATEGORIAS = ordenListaActual;
  else if (ordenNivelActual.tipo === "sublineas") ORDEN_SUBLINEAS[ordenNivelActual.categoria] = ordenListaActual;
  else ORDEN_PRODUCTOS[ordenNivelActual.categoria + "|" + ordenNivelActual.sublinea] = ordenListaActual;

  const sincronizado = await guardarOrdenEnNube();
  document.getElementById("status-orden").textContent = sincronizado
    ? "¡Orden guardado y sincronizado a todos los dispositivos!"
    : "Orden guardado (solo en este dispositivo; revisa la configuración de Firebase).";
}

/* ----- Vistas de toma de precio: activar/inactivar categorías, sublíneas y productos -----
   Se guarda solo lo INACTIVO (por nombre), así todo lo nuevo que llegue en
   una carga de maestra posterior aparece activo por defecto. */
let INACTIVOS_CAT = new Set();
let INACTIVOS_SUB = new Set(); // "categoria|sublinea"
let INACTIVOS_PROD = new Set(); // "id:45" (por ID de la maestra) o el nombre si el producto no tiene ID
const claveSub = (cat, sub) => cat + "|" + sub;

// Por ID primero, así renombrar un producto en la maestra no lo reactiva;
// el nombre queda como respaldo para productos sin ID (o marcados antes de existir el ID).
function claveProd(p) {
  return p.id ? "id:" + String(Number(p.id) || p.id) : p.producto;
}
function productoInactivo(p) {
  return INACTIVOS_PROD.has(claveProd(p)) || INACTIVOS_PROD.has(p.producto);
}

function productosParaToma(cat, sub) {
  if (INACTIVOS_CAT.has(cat) || INACTIVOS_SUB.has(claveSub(cat, sub))) return [];
  return obtenerProductosOrdenados(cat, sub).filter(p => !productoInactivo(p));
}
function sublineasParaToma(cat) {
  if (INACTIVOS_CAT.has(cat)) return [];
  return obtenerSublineasOrdenadas(cat).filter(s => productosParaToma(cat, s).length);
}
function categoriasParaToma() {
  return obtenerCategoriasOrdenadas().filter(c => sublineasParaToma(c).length);
}

function aplicarVistas(d) {
  INACTIVOS_CAT = new Set((d && d.categorias) || []);
  INACTIVOS_SUB = new Set((d && d.sublineas) || []);
  INACTIVOS_PROD = new Set((d && d.productos) || []);
}
function vistasComoObjeto() {
  return { categorias: [...INACTIVOS_CAT], sublineas: [...INACTIVOS_SUB], productos: [...INACTIVOS_PROD] };
}

async function cargarVistas() {
  aplicarVistas(JSON.parse(localStorage.getItem('vistasToma') || 'null'));
  if (!dbNube) return;
  try {
    const snap = await dbNube.collection("maestroMeta").doc("vistas").get();
    if (snap.exists) {
      aplicarVistas(snap.data());
      localStorage.setItem('vistasToma', JSON.stringify(vistasComoObjeto()));
    }
  } catch (err) {
    console.warn("[DataPrice] No se pudieron cargar las vistas desde la nube:", err.message);
  }
}

async function guardarVistasEnNube() {
  localStorage.setItem('vistasToma', JSON.stringify(vistasComoObjeto()));
  if (!dbNube) return false;
  try {
    await dbNube.collection("maestroMeta").doc("vistas").set({ ...vistasComoObjeto(), actualizadoEn: new Date().toISOString() });
    return true;
  } catch (err) {
    console.error("[DataPrice] No se pudieron sincronizar las vistas a la nube:", err);
    return false;
  }
}

/* Árbol del Administrador. Los nombres se referencian por índice (vistasRefs)
   en vez de meterlos en el onclick, para que comillas o símbolos en un
   nombre de producto no rompan el HTML. */
let vistasRefs = [];
let vistasExpandidas = new Set(); // "c:cat" / "s:cat|sub"

function escHtml(t) {
  return (t || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function switchVista(idx, encendido) {
  return `<button type="button" onclick="event.stopPropagation(); vistasToggle(${idx})" title="${encendido ? 'Activo: clic para inactivar' : 'Inactivo: clic para activar'}" class="relative inline-flex h-5 w-9 flex-shrink-0 rounded-full transition-colors ${encendido ? 'bg-blue-600' : 'bg-gray-300'}"><span class="absolute top-0.5 h-4 w-4 rounded-full bg-white shadow transition-all" style="left:${encendido ? '18px' : '2px'}"></span></button>`;
}

function renderVistas() {
  const cont = document.getElementById("vistas-arbol");
  if (!cont) return;
  const q = normTxt(document.getElementById("vistas-buscar").value);
  const coincide = (t) => normTxt(t).includes(q);
  vistasRefs = [];
  let html = "";
  let totalProductos = 0;
  let visiblesProductos = 0;

  obtenerCategoriasOrdenadas().forEach(cat => {
    const catOn = !INACTIVOS_CAT.has(cat);
    const catCoincide = !q || coincide(cat);
    let htmlSubs = "";
    let catTieneCoincidencia = catCoincide;
    let catTotal = 0, catVisibles = 0;

    obtenerSublineasOrdenadas(cat).forEach(sub => {
      const subOn = !INACTIVOS_SUB.has(claveSub(cat, sub));
      const subCoincide = catCoincide || coincide(sub);
      const productos = obtenerProductosOrdenados(cat, sub);
      const subVisibles = productos.filter(p => catOn && subOn && !productoInactivo(p)).length;
      catTotal += productos.length;
      catVisibles += subVisibles;

      const productosMostrar = subCoincide ? productos : productos.filter(p => coincide(p.producto) || coincide(p.marca) || (p.id && normTxt(p.id) === q));
      if (!subCoincide && !productosMostrar.length) return;
      catTieneCoincidencia = true;

      const subIdx = vistasRefs.push({ tipo: "sub", cat, sub }) - 1;
      const subExpandida = !!q || vistasExpandidas.has("s:" + claveSub(cat, sub));
      const subAtenuada = !catOn;
      htmlSubs += `<div onclick="vistasExpandir(${subIdx})" class="flex items-center justify-between gap-3 pl-8 pr-4 py-2 bg-gray-50/60 hover:bg-blue-50 cursor-pointer ${subAtenuada ? 'opacity-50' : ''}">
        <span class="flex items-center gap-2 min-w-0"><i class="fa-solid fa-chevron-${subExpandida ? 'down' : 'right'} text-gray-400 text-[10px] w-3"></i><i class="fa-solid fa-folder text-blue-500 text-xs"></i><span class="font-semibold text-sm text-gray-800 truncate ${subOn ? '' : 'line-through text-gray-400'}">${escHtml(sub)}</span><span class="text-[10px] text-gray-400 font-bold flex-shrink-0">${subVisibles}/${productos.length}</span></span>
        ${switchVista(subIdx, subOn)}
      </div>`;

      if (subExpandida) {
        productosMostrar.forEach(p => {
          const prodOn = !productoInactivo(p);
          const prodIdx = vistasRefs.push({ tipo: "prod", p }) - 1;
          const prodAtenuado = !catOn || !subOn;
          htmlSubs += `<div class="flex items-center justify-between gap-3 pl-14 pr-4 py-2 hover:bg-blue-50/50 ${prodAtenuado ? 'opacity-50' : ''}">
            <span class="min-w-0"><span class="block text-xs font-semibold text-gray-800 truncate ${prodOn ? '' : 'line-through text-gray-400'}">${p.id ? `<span class="text-gray-400 font-bold mr-1">#${escHtml(p.id)}</span>` : ""}${escHtml(p.producto)}</span><span class="block text-[10px] text-gray-500">${escHtml(p.marca)} · ${escHtml(p.size || "-")} · ${p.tipo === "SURA" ? '<span class="text-blue-600 font-bold">SURA</span>' : 'Competencia'}</span></span>
            ${switchVista(prodIdx, prodOn)}
          </div>`;
        });
      }
    });

    totalProductos += catTotal;
    visiblesProductos += catVisibles;
    if (!catTieneCoincidencia) return;

    const catIdx = vistasRefs.push({ tipo: "cat", cat }) - 1;
    const catExpandida = !!q || vistasExpandidas.has("c:" + cat);
    html += `<div onclick="vistasExpandir(${catIdx})" class="flex items-center justify-between gap-3 px-4 py-3 hover:bg-blue-50 cursor-pointer">
      <span class="flex items-center gap-2 min-w-0"><i class="fa-solid fa-chevron-${catExpandida ? 'down' : 'right'} text-gray-400 text-[10px] w-3"></i><i class="fa-solid fa-layer-group text-blue-600 text-sm"></i><span class="font-black text-sm text-gray-900 truncate ${catOn ? '' : 'line-through text-gray-400'}">${escHtml(cat)}</span><span class="text-[10px] text-gray-400 font-bold flex-shrink-0">${catVisibles}/${catTotal}</span></span>
      ${switchVista(catIdx, catOn)}
    </div>` + (catExpandida ? htmlSubs : "");
  });

  cont.innerHTML = html || `<div class="py-6 text-center text-xs text-gray-400 italic">${MAESTRO_PRODUCTOS.length ? "Sin resultados para la búsqueda." : "Carga primero la Maestra de Productos."}</div>`;
  document.getElementById("vistas-resumen").textContent = `${visiblesProductos} de ${totalProductos} productos visibles para la toma`;
}

function vistasExpandir(idx) {
  const r = vistasRefs[idx];
  if (!r) return;
  const clave = r.tipo === "cat" ? "c:" + r.cat : "s:" + claveSub(r.cat, r.sub);
  if (vistasExpandidas.has(clave)) vistasExpandidas.delete(clave); else vistasExpandidas.add(clave);
  renderVistas();
}

async function vistasToggle(idx) {
  const r = vistasRefs[idx];
  if (!r) return;
  const alternar = (set, clave) => { if (set.has(clave)) set.delete(clave); else set.add(clave); };
  if (r.tipo === "cat") alternar(INACTIVOS_CAT, r.cat);
  else if (r.tipo === "sub") alternar(INACTIVOS_SUB, claveSub(r.cat, r.sub));
  else if (productoInactivo(r.p)) { INACTIVOS_PROD.delete(claveProd(r.p)); INACTIVOS_PROD.delete(r.p.producto); }
  else INACTIVOS_PROD.add(claveProd(r.p));
  renderVistas();

  const status = document.getElementById("status-vistas");
  status.className = "text-xs font-bold text-gray-500";
  status.textContent = "Guardando...";
  const sincronizado = await guardarVistasEnNube();
  status.className = `text-xs font-bold ${sincronizado ? "text-emerald-600" : "text-amber-600"}`;
  status.textContent = sincronizado
    ? "Cambios guardados y sincronizados a todos los dispositivos."
    : "Guardado solo en este dispositivo (revisa la configuración de Firebase).";
}

/* ============================== 8. INICIALIZACIÓN ============================== */
window.onload = async function () {
  inicializarFirebase();
  await cargarMaestrosBase();
  await sincronizarMaestrosDesdeNube();
  await Promise.all([cargarOrdenListados(), cargarIndiceImagenesProductos(), cargarVistas()]);
  document.getElementById("fecha-captura").value = new Date().toISOString().split('T')[0];
  actualizarInfoAsesorDOM();
  poblarFiltrosVisualizador();
  sincronizarPendientes(); // sube en segundo plano lo que quedó pendiente de una sesión anterior sin señal
  window.addEventListener('online', sincronizarPendientes);

  // Restaura la sesión activa (si la hay) para que el usuario no tenga que
  // volver a loguearse en cada recarga; "Cerrar sesión" la borra.
  const sesion = JSON.parse(localStorage.getItem('sesionActiva') || 'null');
  if (sesion && sesion.tipo === 'ASESOR' && asesorInfo.codigo) {
    sesionTipo = 'ASESOR';
    goToScreen(1);
  } else if (sesion && sesion.tipo === 'ADMIN') {
    sesionTipo = 'ADMIN';
    goToScreen(0);
  } else if (sesion && sesion.tipo === 'ANALYTICS') {
    sesionTipo = 'ANALYTICS';
    goToScreen(5);
  } else {
    goToScreen('login');
  }
};
