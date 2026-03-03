// ═══════════════════════════════════════════════════════════════════
// PAGOS CDP — Google Apps Script Backend  v4.0
// ═══════════════════════════════════════════════════════════════════
// NUEVAS HOJAS REQUERIDAS:
//   SolicitudesAlta  — Cols: ID | EquipoID | Nombre | DNI | FechaNac | TipoSocio |
//                            Mas35 | IC | TipoTitulo | FechasSolicitud | Estado |
//                            ObsAdmin | DriveFolder | FechaResolucion
//   HistorialBajas   — Cols: JugadorID | EquipoID | Nombre | FechaBaja | UltimoMesPago | Motivo
//   Parametros       — Fila 2: PinAdmin | CuotaSocial | CuotaDeportiva
//   Jugadores        — Agregar cols: DNI(F) | FechaNac(G) | TipoSocio(H) | HabilitadoManual(I)
// ═══════════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

const TABS = {
  equipos:          'Equipos',
  comprobantes:     'CargaComprobantes',
  jugadores:        'Jugadores',
  auth:             'AuthControl',
  parametros:       'Parametros',
  fixture:          'Fixture',
  sanciones:        'Sanciones',
  estudiosMedicos:  'EstudiosMedicos',
  solicitudesAlta:  'SolicitudesAlta',
  historialBajas:   'HistorialBajas',
  saludJugadores:   'SaludJugadores',
};

// Correos destino para solicitudes de alta
const CORREOS_ADMIN = ['torneos.cdp@gmail.com', 'germanalbertorodriguez@gmail.com'];

// ═══════════════════════════════════════════════════════════════════
// SISTEMA DE SESIONES (Token-based auth)
// ═══════════════════════════════════════════════════════════════════
const SESSION_DURATION = 21600; // 6 horas en segundos (máximo de CacheService)

// Helper: normaliza IDs de equipo para comparación flexible
// Maneja diferencias entre "Abogados_A" (ID) vs "Abogados A" (nombre) vs "Abogados A (L)" etc.
function normId(s) {
  return String(s || '').trim().replace(/[\s_]+/g, '_').toLowerCase();
}
function matchEquipo(a, b) {
  return normId(a) === normId(b);
}
// Normalizar nombre para comparación (quita acentos, mayúsculas, espacios extra)
function normNombre(s) {
  return String(s||'').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}
// Leer sanciones activas — lee nombre (col A) + equipo (col B)
function leerSancionesActivas() {
  const ids = new Set();
  const registros = [];   // {nombre, equipo, partidos, fechaHab} para matching
  const wsSanc = SS.getSheetByName(TABS.sanciones);
  if (wsSanc) {
    const ds = wsSanc.getDataRange().getValues();
    for (let i = 2; i < ds.length; i++) {
      const val = String(ds[i][0] || '').trim();
      if (!val) continue;
      if (/^\d+$/.test(val)) ids.add(val);       // ID numérico
      const equipo = normNombre(ds[i][1]);         // col B = equipo
      const partidos = String(ds[i][2] || '').trim(); // col C = partidos/motivo
      const fechaHab = ds[i][4] ? formatFecha(ds[i][4]) : ''; // col E = fecha habilitación
      registros.push({ nombre: normNombre(val), equipo, partidos, fechaHab });
    }
  }
  return { ids, registros };
}
// Matching por nombre+equipo, devuelve info de la sanción o null
function buscarSancion(sanc, jugId, jugNombre, jugEquipo) {
  const normJug = normNombre(jugNombre);
  const palabrasJug = normJug.split(' ');
  const normEq = normNombre(jugEquipo);
  for (const reg of sanc.registros) {
    if (reg.equipo && !matchEquipo(reg.equipo, normEq)) continue;
    const palabrasSanc = reg.nombre.split(' ');
    if (palabrasSanc.length >= 2 && palabrasSanc.every(p => palabrasJug.includes(p))) {
      return reg; // devuelve {nombre, equipo, partidos, fechaHab}
    }
  }
  return null;
}
function jugadorSancionado(sanc, jugId, jugNombre, jugEquipo) {
  if (sanc.ids.has(String(jugId).trim())) return true;
  return buscarSancion(sanc, jugId, jugNombre, jugEquipo) !== null;
}
const CACHE = CacheService.getScriptCache();

function generarToken() {
  return Utilities.getUuid();
}

function crearSesion(rol, equipoId, equipoNombre) {
  const token = generarToken();
  const sesion = JSON.stringify({
    rol: rol,               // 'delegado' | 'admin'
    equipoId: equipoId || null,
    equipoNombre: equipoNombre || null,
    creada: new Date().toISOString()
  });
  CACHE.put('ses_' + token, sesion, SESSION_DURATION);
  return token;
}

function obtenerSesion(token) {
  if (!token) return null;
  const raw = CACHE.get('ses_' + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch(e) {
    return null;
  }
}

function destruirSesion(token) {
  if (token) CACHE.remove('ses_' + token);
}

// Middleware de autenticación: valida token y rol requerido
function verificarAuth(params, rolRequerido) {
  const token = params.token;
  if (!token) {
    return { ok: false, error: 'NO_TOKEN', msg: 'Sesión no iniciada. Por favor ingresá nuevamente.' };
  }
  const sesion = obtenerSesion(token);
  if (!sesion) {
    return { ok: false, error: 'TOKEN_EXPIRED', msg: 'Tu sesión expiró. Por favor ingresá nuevamente.' };
  }
  if (rolRequerido === 'admin' && sesion.rol !== 'admin') {
    return { ok: false, error: 'FORBIDDEN', msg: 'No tenés permisos para esta acción.' };
  }
  // Para endpoints de delegado, verificar que accede solo a datos de su equipo
  if (sesion.rol === 'delegado' && params.equipoId) {
    if (!matchEquipo(params.equipoId, sesion.equipoId)) {
      return { ok: false, error: 'FORBIDDEN', msg: 'No tenés permisos para acceder a datos de otro equipo.' };
    }
  }
  return { ok: true, sesion };
}

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const accion = e.parameter.accion;
    const p = e.parameter;

    // ── ACCIONES PÚBLICAS (sin token) ──────────────────────────
    switch (accion) {
      case 'getEquipos':   return makeResponse(getEquipos());       // sanitizado: sin correos
      case 'login':        return makeResponse(login(p));
      case 'loginAdmin':   return makeResponse(loginAdmin(p));
      case 'logout':       return makeResponse(logout(p));
    }

    // ── TODAS LAS DEMÁS ACCIONES REQUIEREN TOKEN ──────────────
    const auth = verificarAuth(p, null); // verificar que hay sesión válida
    if (!auth.ok) return makeResponse(auth);
    const sesion = auth.sesion;

    // ── ACCIONES DE ADMIN ─────────────────────────────────────
    const ADMIN_ACTIONS = [
      'getJugadoresAdmin', 'getAllComprobantes', 'getResumen',
      'aprobar', 'rechazar',
      'altaJugador', 'bajaJugador', 'editarJugador', 'traspasarJugador',
      'habilitarJugador', 'inhabilitarJugador', 'resetHabManual',
      'resolverSolicitudAlta',
      'setCuotas',
      'getPlanillaData',
      'getSaludAdmin'
    ];

    if (ADMIN_ACTIONS.indexOf(accion) >= 0) {
      if (sesion.rol !== 'admin') {
        return makeResponse({ ok: false, error: 'FORBIDDEN', msg: 'Acción solo para administradores.' });
      }
    }

    // ── ACCIONES DE DELEGADO: verificar propiedad del equipo ──
    const DELEGADO_TEAM_ACTIONS = [
      'getJugadores', 'getComprobantes', 'getDelegadoInit', 'getHabilitaciones',
      'getSaludJugador', 'getSaludEquipo', 'setSaludJugador',
      'getEstudiosMedicos', 'getArchivosAdjuntosSolicitud',
      'getSolicitudesAlta', 'getHistorialBajas'
    ];

    if (sesion.rol === 'delegado' && DELEGADO_TEAM_ACTIONS.indexOf(accion) >= 0) {
      // Para delegados, forzar que equipoId sea el de su sesión
      if (p.equipoId && !matchEquipo(p.equipoId, sesion.equipoId)) {
        return makeResponse({ ok: false, error: 'FORBIDDEN', msg: 'Solo podés acceder a datos de tu equipo.' });
      }
    }

    // ── DISPATCH ───────────────────────────────────────────────
    switch (accion) {
      // Autenticado (cualquier rol)
      case 'getJugadores':            return makeResponse(getJugadores(p));
      case 'getComprobantes':         return makeResponse(getComprobantes(p));
      case 'getDelegadoInit':         return makeResponse(getDelegadoInit(p));
      case 'getComprobante':          return makeResponse(getComprobante(p));
      case 'getHabilitaciones':       return makeResponse(getHabilitaciones(p));
      case 'getFixture':              return makeResponse(getFixture());
      case 'getEstudiosMedicos':      return makeResponse(getEstudiosMedicos(p));
      case 'getArchivosAdjuntosSolicitud': return makeResponse(getArchivosAdjuntosSolicitud(p));
      case 'getSaludJugador':         return makeResponse(getSaludJugador(p));
      case 'getSaludEquipo':          return makeResponse(getSaludEquipo(p));
      case 'setSaludJugador':         return makeResponse(setSaludJugador(p));
      case 'getSolicitudesAlta':      return makeResponse(getSolicitudesAlta(p));
      case 'getHistorialBajas':       return makeResponse(getHistorialBajas(p));
      // Admin
      case 'getJugadoresAdmin':       return makeResponse(getJugadoresAdmin(p));
      case 'getAllComprobantes':       return makeResponse(getAllComprobantes(p));
      case 'getComprobante':          return makeResponse(getComprobante(p));
      case 'getResumen':              return makeResponse(getResumen(p));
      case 'aprobar':                 return makeResponse(aprobar(p));
      case 'rechazar':                return makeResponse(rechazar(p));
      case 'getPlanillaData':         return makeResponse(getPlanillaData(p));
      case 'altaJugador':             return makeResponse(altaJugador(p));
      case 'bajaJugador':             return makeResponse(bajaJugador(p));
      case 'editarJugador':           return makeResponse(editarJugador(p));
      case 'traspasarJugador':        return makeResponse(traspasarJugador(p));
      case 'habilitarJugador':        return makeResponse(habilitarJugador(p));
      case 'inhabilitarJugador':      return makeResponse(inhabilitarJugador(p));
      case 'resetHabManual':          return makeResponse(resetHabManual(p));
      case 'resolverSolicitudAlta':   return makeResponse(resolverSolicitudAlta(p));
      case 'getParametros':           return makeResponse(getParametros());
      case 'setCuotas':               return makeResponse(setCuotas(p));
      case 'getSaludAdmin':           return makeResponse(getSaludAdmin(p));
      default: return makeResponse({ ok: false, msg: 'Accion no reconocida: ' + accion });
    }
  } catch(err) {
    return makeResponse({ ok: false, msg: err.toString() });
  }
}

function doPost(e) {
  try {
    const accion = e.parameter.accion;
    
    // Todas las acciones POST requieren autenticación
    const auth = verificarAuth(e.parameter, null);
    if (!auth.ok) return makeResponse(auth);
    const sesion = auth.sesion;
    
    // Para delegados, verificar propiedad del equipo en uploads
    if (sesion.rol === 'delegado' && e.parameter.equipoId) {
      if (!matchEquipo(e.parameter.equipoId, sesion.equipoId)) {
        return makeResponse({ ok: false, error: 'FORBIDDEN', msg: 'Solo podés cargar datos de tu equipo.' });
      }
    }
    
    if (accion === 'cargarComprobante')    return makeResponse(cargarComprobante(e.parameter));
    if (accion === 'solicitarAlta')        return makeResponse(solicitarAlta(e.parameter));
    if (accion === 'subirEstudioMedico')   return makeResponse(subirEstudioMedico(e.parameter));
    return makeResponse({ ok: false, msg: 'POST accion no reconocida' });
  } catch(err) {
    return makeResponse({ ok: false, msg: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════════
// AUTH
// ═══════════════════════════════════════════════════════════════════

function login({ equipoId, pin }) {
  const ws = SS.getSheetByName(TABS.equipos);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (matchEquipo(row[0], equipoId)) {
      const pinSheet = String(row[4] || '').trim();
      if (pinSheet === String(pin).trim()) {
        const nombre = String(row[1]).trim();
        const cat = String(row[2] || '').trim();
        const token = crearSesion('delegado', String(row[0]).trim(), nombre);
        // NO devolver correo ni PIN al frontend
        return { ok: true, token: token, equipo: { id: row[0], nombre: nombre, cat: cat } };
      }
      return { ok: false, msg: 'PIN incorrecto.' };
    }
  }
  return { ok: false, msg: 'Equipo no encontrado.' };
}

function loginAdmin({ pin }) {
  const ws = SS.getSheetByName(TABS.parametros);
  const data = ws.getDataRange().getValues();
  const pinAdmin = String(data[1] ? (data[1][1] || '') : '').trim(); // col B = ClaveControl
  if (String(pin).trim() === pinAdmin) {
    const token = crearSesion('admin', null, null);
    return { ok: true, token: token };
  }
  return { ok: false, msg: 'PIN incorrecto.' };
}

function logout({ token }) {
  destruirSesion(token);
  return { ok: true, msg: 'Sesión cerrada.' };
}

// ═══════════════════════════════════════════════════════════════════
// EQUIPOS
// ═══════════════════════════════════════════════════════════════════

function getEquipos() {
  const ws = SS.getSheetByName(TABS.equipos);
  const data = ws.getDataRange().getValues();
  const equipos = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // SEGURIDAD: NO exponer correo (col D) ni PIN (col E) al frontend
    equipos.push({ id: String(data[i][0]).trim(), nombre: String(data[i][1]).trim(), cat: String(data[i][2]||'') });
  }
  return { ok: true, equipos };
}

// ═══════════════════════════════════════════════════════════════════
// PARÁMETROS / CUOTAS
// ═══════════════════════════════════════════════════════════════════

function getParametros() {
  const ws = SS.getSheetByName(TABS.parametros);
  if (!ws) return { ok: false, msg: 'Hoja Parametros no encontrada' };
  const data = ws.getDataRange().getValues();
  // Estructura real: A=ParametroID | B=ClaveControl | C=Descripcion | D=CuotaSocial | E=CuotaDeportiva
  const fila = data[1] || [];
  // SEGURIDAD: NO exponer pinAdmin (col B) al frontend
  return {
    ok: true,
    cuotaSocial:     Number(fila[3] || 0),  // col D
    cuotaDeportiva:  Number(fila[4] || 0),  // col E
  };
}

function setCuotas({ cuotaSocial, cuotaDeportiva }) {
  const ws = SS.getSheetByName(TABS.parametros);
  if (!ws) return { ok: false, msg: 'Hoja Parametros no encontrada' };
  if (cuotaSocial    !== undefined) ws.getRange(2, 4).setValue(Number(cuotaSocial)); // col D
  if (cuotaDeportiva !== undefined) ws.getRange(2, 5).setValue(Number(cuotaDeportiva)); // col E
  return { ok: true, msg: 'Cuotas actualizadas.' };
}

// ═══════════════════════════════════════════════════════════════════
// JUGADORES
// ═══════════════════════════════════════════════════════════════════
// Columnas Jugadores: A=ID B=EquipoID C=Nombre D=Mas35 E=IC F=DNI G=FechaNac H=TipoSocio I=HabilitadoManual

function getJugadores({ equipoId }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  const jugadores = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0] || !matchEquipo(r[1], equipoId)) continue;
    jugadores.push({
      id:         String(r[0]).trim(),
      equipoId:   String(r[1]).trim(),
      nombre:     String(r[2]||'').trim(),
      mas35:      String(r[3]||'').toUpperCase() === 'X',
      ic:         String(r[4]||'').toUpperCase() === 'X',
      dni:        String(r[5]||'').trim(),
      fechaNac:   r[6] ? formatFecha(r[6]) : '',
      tipoSocio:  String(r[7]||'Activo').trim(),
      habManual:  String(r[8]||'').trim(), // '' | 'SI' | 'NO'
    });
  }
  jugadores.sort((a,b) => a.nombre.localeCompare(b.nombre, 'es'));
  return { ok: true, jugadores };
}

function getJugadoresAdmin({ equipoId }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  const jugadores = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (equipoId && !matchEquipo(r[1], equipoId)) continue;
    jugadores.push({
      id:         String(r[0]).trim(),
      equipoId:   String(r[1]).trim(),
      nombre:     String(r[2]||'').trim(),
      mas35:      String(r[3]||'').toUpperCase() === 'X',
      ic:         String(r[4]||'').toUpperCase() === 'X',
      dni:        String(r[5]||'').trim(),
      fechaNac:   r[6] ? formatFecha(r[6]) : '',
      tipoSocio:  String(r[7]||'Activo').trim(),
      habManual:  String(r[8]||'').trim(),
    });
  }
  jugadores.sort((a,b) => String(a.equipoId).localeCompare(String(b.equipoId)) || a.nombre.localeCompare(b.nombre,'es'));
  return { ok: true, jugadores };
}

function altaJugador({ equipoId, nombre, mas35, ic, dni, fechaNac, tipoSocio }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = parseFloat(data[i][0]) || 0;
    if (id > maxId) maxId = id;
  }
  const nuevoId = maxId + 1;
  ws.appendRow([
    nuevoId,
    String(equipoId).trim(),
    String(nombre).trim().toUpperCase(),
    (mas35==='true'||mas35===true)?'X':'',
    (ic==='true'||ic===true)?'X':'',
    String(dni||'').trim(),
    fechaNac || '',
    String(tipoSocio||'Activo').trim(),
    '',  // habManual vacío
  ]);
  return { ok: true, id: nuevoId, msg: 'Jugador dado de alta correctamente.' };
}

function bajaJugador({ jugadorId, ultimoMesPago, motivo }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      // Guardar en historial
      const wsH = SS.getSheetByName(TABS.historialBajas);
      if (wsH) {
        wsH.appendRow([
          String(data[i][0]).trim(),
          String(data[i][1]).trim(),
          String(data[i][2]||'').trim(),
          new Date(),
          String(ultimoMesPago||'').trim(),
          String(motivo||'').trim(),
        ]);
      }
      ws.deleteRow(i+1);
      return { ok: true, msg: 'Jugador dado de baja y guardado en historial.' };
    }
  }
  return { ok: false, msg: 'Jugador no encontrado.' };
}

function editarJugador({ jugadorId, nombre, mas35, ic, dni, fechaNac, tipoSocio }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      if (nombre    !== undefined) ws.getRange(i+1,3).setValue(String(nombre).trim().toUpperCase());
      if (mas35     !== undefined) ws.getRange(i+1,4).setValue((mas35==='true'||mas35===true)?'X':'');
      if (ic        !== undefined) ws.getRange(i+1,5).setValue((ic==='true'||ic===true)?'X':'');
      if (dni       !== undefined) ws.getRange(i+1,6).setValue(String(dni).trim());
      if (fechaNac  !== undefined) ws.getRange(i+1,7).setValue(fechaNac);
      if (tipoSocio !== undefined) ws.getRange(i+1,8).setValue(String(tipoSocio).trim());
      return { ok: true, msg: 'Jugador actualizado.' };
    }
  }
  return { ok: false, msg: 'Jugador no encontrado.' };
}

function traspasarJugador({ jugadorId, nuevoEquipoId }) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      ws.getRange(i+1,2).setValue(String(nuevoEquipoId).trim());
      return { ok: true, msg: 'Traspaso realizado correctamente.' };
    }
  }
  return { ok: false, msg: 'Jugador no encontrado.' };
}

function habilitarJugador({ jugadorId }) {
  return setHabilitacionManual(jugadorId, 'SI');
}

function inhabilitarJugador({ jugadorId }) {
  return setHabilitacionManual(jugadorId, 'NO');
}

function resetHabManual({ jugadorId }) {
  return setHabilitacionManual(jugadorId, '');
}

function setHabilitacionManual(jugadorId, valor) {
  const ws = SS.getSheetByName(TABS.jugadores);
  if (!ws) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      ws.getRange(i+1, 9).setValue(valor);
      return { ok: true, msg: valor === 'SI' ? 'Jugador habilitado manualmente.' : 'Jugador inhabilitado manualmente.' };
    }
  }
  return { ok: false, msg: 'Jugador no encontrado.' };
}

function getHistorialBajas({ equipoId } = {}) {
  const ws = SS.getSheetByName(TABS.historialBajas);
  if (!ws) return { ok: true, historial: [] };
  const data = ws.getDataRange().getValues();
  const historial = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (equipoId && !matchEquipo(r[1], equipoId)) continue;
    historial.push({
      jugadorId:     String(r[0]).trim(),
      equipoId:      String(r[1]).trim(),
      nombre:        String(r[2]||'').trim(),
      fechaBaja:     r[3] ? formatFecha(r[3]) : '',
      ultimoMesPago: String(r[4]||'').trim(),
      motivo:        String(r[5]||'').trim(),
    });
  }
  historial.sort((a,b) => new Date(b.fechaBaja) - new Date(a.fechaBaja));
  return { ok: true, historial };
}

// ═══════════════════════════════════════════════════════════════════
// HABILITACIONES
// ═══════════════════════════════════════════════════════════════════

function getHabilitaciones({ equipoId }) {
  const mesActual   = obtenerMesActual();
  const mesAnterior = obtenerMesAnterior();

  const comps = leerComprobantes().filter(c =>
    (c.mes===mesActual || c.mes===mesAnterior) && c.estado==='Aprobado');
  const jugAdminOk = new Set(comps.map(c => String(c.jugadorId).trim()));

  const sanciones = leerSancionesActivas();

  const sinMedico = {};
  const wsEM = SS.getSheetByName(TABS.estudiosMedicos);
  if (wsEM) {
    const dem = wsEM.getDataRange().getValues();
    for (let i=2;i<dem.length;i++) {
      if (!dem[i][0]) continue;
      const est = String(dem[i][3]||'').toLowerCase();
      if (est==='vencido'||est==='no presentado') {
        sinMedico[String(dem[i][0]).trim()] = {
          estado: String(dem[i][3]),
          vencimiento: dem[i][4] ? formatFecha(dem[i][4]) : '',
        };
      }
    }
  }

  const wsJug = SS.getSheetByName(TABS.jugadores);
  if (!wsJug) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const dataJug = wsJug.getDataRange().getValues();

  const jugadores = [];
  for (let i=1;i<dataJug.length;i++) {
    const r = dataJug[i];
    if (!r[0] || !matchEquipo(r[1], equipoId)) continue;
    const jugId    = String(r[0]).trim();
    const jugNombre = String(r[2]||'').trim();
    const habManual = String(r[8]||'').trim().toUpperCase();
    let estado='Habilitado', detalle='';

    // HabilitadoManual tiene prioridad
    if (habManual === 'NO') {
      estado = 'I. Administrativa'; detalle = 'Inhabilitado manualmente por administración';
    } else if (habManual === 'SI') {
      estado = 'Habilitado'; detalle = 'Habilitado manualmente por administración';
    } else {
      const infoSanc = buscarSancion(sanciones, jugId, jugNombre, String(r[1]||''));
      if (infoSanc) {
        estado='I. Deportiva';
        detalle='Sanción: ' + (infoSanc.partidos||'s/d') + (infoSanc.fechaHab ? ' · Hábil: '+infoSanc.fechaHab : ' · Sin fecha hábil');
      } else if (sinMedico[jugId]) {
        estado='I. Medica';
        detalle=sinMedico[jugId].estado+(sinMedico[jugId].vencimiento?' · Venc: '+sinMedico[jugId].vencimiento:'');
      } else {
        const tipoSocio = String(r[7]||'Activo').trim().toLowerCase();
        if (tipoSocio === 'pasivo') {
          if (!jugAdminOk.has(jugId)) { estado='I. Administrativa'; detalle='Sin pago aprobado en '+mesActual; }
        } else {
          if (!jugAdminOk.has(jugId)) { estado='I. Administrativa'; detalle='Sin pago aprobado en '+mesActual; }
        }
      }
    }

    jugadores.push({ id:jugId, nombre:String(r[2]||'').trim(),
      mas35:String(r[3]||'').toUpperCase()==='X', ic:String(r[4]||'').toUpperCase()==='X',
      tipoSocio: String(r[7]||'Activo').trim(),
      habManual,
      estado, detalle });
  }

  const orden = {'Habilitado':0,'I. Deportiva':1,'I. Medica':2,'I. Administrativa':3};
  jugadores.sort((a,b) => (orden[a.estado]||9)-(orden[b.estado]||9) || a.nombre.localeCompare(b.nombre,'es'));
  return { ok: true, jugadores, mesActual, mesAnterior };
}

// ═══════════════════════════════════════════════════════════════════
// SOLICITUDES DE ALTA (delegado → admin)
// ═══════════════════════════════════════════════════════════════════
// Cols SolicitudesAlta: A=ID | B=EquipoID | C=NombreEquipo | D=Nombre | E=DNI | F=FechaNac
//                       G=TipoSocio | H=Mas35 | I=IC | J=TipoTitulo | K=FechaSolicitud
//                       L=Estado | M=ObsAdmin | N=DriveFolder | O=FechaResolucion

function getSolicitudesAlta({ equipoId, estado } = {}) {
  const ws = SS.getSheetByName(TABS.solicitudesAlta);
  if (!ws) return { ok: true, solicitudes: [] };
  const data = ws.getDataRange().getValues();
  const solicitudes = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (equipoId && !matchEquipo(r[1], equipoId)) continue;
    if (estado && String(r[11]||'Pendiente') !== estado) continue;
    solicitudes.push({
      id:              String(r[0]).trim(),
      equipoId:        String(r[1]).trim(),
      nombreEquipo:    String(r[2]||'').trim(),
      nombre:          String(r[3]||'').trim(),
      dni:             String(r[4]||'').trim(),
      fechaNac:        String(r[5]||'').trim(),
      tipoSocio:       String(r[6]||'Activo').trim(),
      mas35:           String(r[7]||'').toUpperCase() === 'X',
      ic:              String(r[8]||'').toUpperCase() === 'X',
      tipoTitulo:      String(r[9]||'').trim(),
      fechaSolicitud:  r[10] ? formatFecha(r[10]) : '',
      estado:          String(r[11]||'Pendiente').trim(),
      obsAdmin:        String(r[12]||'').trim(),
      driveFolder:     String(r[13]||'').trim(),
      fechaResolucion: r[14] ? formatFecha(r[14]) : '',
    });
  }
  solicitudes.sort((a,b) => new Date(b.fechaSolicitud) - new Date(a.fechaSolicitud));
  return { ok: true, solicitudes };
}

function solicitarAlta(params) {
  const { equipoId, nombreEquipo, nombre, dni, fechaNac, tipoSocio, mas35, ic,
          tipoTitulo, archivos } = params;

  if (!nombre || !dni) return { ok: false, msg: 'Nombre y DNI son obligatorios.' };

  // Crear carpeta en Drive para los archivos
  const uid = Utilities.getUuid().substring(0,8).toUpperCase();
  let driveFolderUrl = '';
  try {
    const base = 'SolicitudesAlta_CDP';
    const iterBase = DriveApp.getFoldersByName(base);
    const carpetaBase = iterBase.hasNext() ? iterBase.next() : DriveApp.createFolder(base);
    const carpetaEq = obtenerSubcarpeta(carpetaBase, equipoId);
    const carpetaSol = carpetaBase.createFolder(uid + '_' + String(nombre).trim().replace(/\s/g,'_'));
    driveFolderUrl = carpetaSol.getUrl();

    // Subir archivos adjuntos
    if (archivos) {
      const arr = JSON.parse(archivos);
      arr.forEach(a => {
        const blob = Utilities.newBlob(Utilities.base64Decode(a.data), a.tipo, a.nombre);
        const f = carpetaSol.createFile(blob);
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      });
    }
  } catch(e) {
    Logger.log('Drive error: ' + e);
  }

  // Guardar en sheet
  const ws = SS.getSheetByName(TABS.solicitudesAlta);
  if (!ws) return { ok: false, msg: 'Hoja SolicitudesAlta no encontrada.' };
  ws.appendRow([
    uid,
    String(equipoId).trim(),
    String(nombreEquipo||'').trim(),
    String(nombre).trim().toUpperCase(),
    String(dni).trim(),
    String(fechaNac||'').trim(),
    String(tipoSocio||'Activo').trim(),
    (mas35==='true'||mas35===true)?'X':'',
    (ic==='true'||ic===true)?'X':'',
    String(tipoTitulo||'').trim(),
    new Date(),
    'Pendiente',
    '',
    driveFolderUrl,
    '',
  ]);

  // Enviar mail a admins
  try {
    const asunto = `[CDP] Solicitud de Alta — ${nombre} (${nombreEquipo})`;
    const cuerpo = `Estimado administrador,\n\n` +
      `Se recibió una solicitud de alta nueva:\n\n` +
      `Equipo: ${nombreEquipo} (${equipoId})\n` +
      `Jugador: ${nombre}\n` +
      `DNI: ${dni}\n` +
      `Fecha de Nacimiento: ${fechaNac}\n` +
      `Tipo de Socio: ${tipoSocio}\n` +
      `Tipo de Título: ${tipoTitulo}\n` +
      `Mayor de 35: ${mas35==='true'?'Sí':'No'}\n` +
      `Intercategoría: ${ic==='true'?'Sí':'No'}\n\n` +
      `Archivos adjuntos en Drive: ${driveFolderUrl}\n\n` +
      `Por favor ingresá a la app para resolver la solicitud (ID: ${uid}).\n\n` +
      `Sistema de Pagos — Club de Profesionales`;
    CORREOS_ADMIN.forEach(correo => {
      MailApp.sendEmail(correo, asunto, cuerpo);
    });
  } catch(mailErr) { Logger.log('Email error: ' + mailErr); }

  return { ok: true, id: uid, msg: 'Solicitud enviada correctamente. El administrador la revisará a la brevedad.' };
}

function resolverSolicitudAlta({ id, decision, obs, equipoId, nombre, dni, fechaNac, tipoSocio, mas35, ic }) {
  // decision: 'Aprobada' | 'Rechazada' | 'Observada'
  const ws = SS.getSheetByName(TABS.solicitudesAlta);
  if (!ws) return { ok: false, msg: 'Hoja SolicitudesAlta no encontrada.' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(id).trim()) continue;
    ws.getRange(i+1, 12).setValue(decision);
    ws.getRange(i+1, 13).setValue(obs || '');
    ws.getRange(i+1, 15).setValue(new Date());

    if (decision === 'Aprobada') {
      // Dar de alta al jugador automáticamente
      const res = altaJugador({
        equipoId: equipoId || data[i][1],
        nombre:   nombre   || data[i][3],
        mas35:    mas35    !== undefined ? mas35 : (String(data[i][7]).toUpperCase()==='X'),
        ic:       ic       !== undefined ? ic    : (String(data[i][8]).toUpperCase()==='X'),
        dni:      dni      || data[i][4],
        fechaNac: fechaNac || data[i][5],
        tipoSocio: tipoSocio || data[i][6],
      });
      if (!res.ok) return res;
    }

    // Notificar al delegado
    try {
      const correoEq = getCorreoEquipo(data[i][1]);
      if (correoEq) {
        const msgs = {
          'Aprobada':  'fue APROBADA. El jugador fue dado de alta en el sistema.',
          'Rechazada': 'fue RECHAZADA.' + (obs ? '\nMotivo: ' + obs : ''),
          'Observada': 'requiere aclaraciones:\n' + (obs || ''),
        };
        MailApp.sendEmail(correoEq,
          `[CDP] Solicitud de Alta — ${data[i][3]} — ${decision}`,
          `Estimado delegado,\n\nLa solicitud de alta del jugador ${data[i][3]} ${msgs[decision]||''}\n\nSistema de Pagos — Club de Profesionales`
        );
      }
    } catch(mailErr) { Logger.log('Email error: ' + mailErr); }

    return { ok: true, msg: `Solicitud ${decision.toLowerCase()} correctamente.` };
  }
  return { ok: false, msg: 'Solicitud no encontrada.' };
}

// ═══════════════════════════════════════════════════════════════════
// ESTUDIOS MÉDICOS
// ═══════════════════════════════════════════════════════════════════
// Cols EstudiosMedicos: A=JugadorID | B=EquipoID | C=FechaEstudio | D=Estado | E=Vencimiento | F=FileId | G=EsSenior

function getEstudiosMedicos({ equipoId } = {}) {
  const ws = SS.getSheetByName(TABS.estudiosMedicos);
  if (!ws) return { ok: true, estudios: [] };
  const data = ws.getDataRange().getValues();
  const estudios = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (equipoId && !matchEquipo(r[1], equipoId)) continue;
    estudios.push({
      jugadorId:    String(r[0]).trim(),
      equipoId:     String(r[1]).trim(),
      fechaEstudio: r[2] ? formatFecha(r[2]) : '',
      estado:       String(r[3]||'').trim(),
      vencimiento:  r[4] ? formatFecha(r[4]) : '',
      fileId:       String(r[5]||'').trim(),
      esSenior:     String(r[6]||'').toUpperCase() === 'X',
    });
  }
  return { ok: true, estudios };
}

function subirEstudioMedico(params) {
  const { jugadorId, equipoId, esSenior, archivo, nombreArchivo, tipoArchivo } = params;
  if (!archivo) return { ok: false, msg: 'No se recibió el archivo.' };

  const carpeta = obtenerCarpetaEstudios(equipoId);
  const ext = nombreArchivo.split('.').pop().toLowerCase();
  const uid = Utilities.getUuid().substring(0,8);
  const nombreFinal = `${uid}_${jugadorId}_estudio.${ext}`;
  const blob = Utilities.newBlob(Utilities.base64Decode(archivo), tipoArchivo, nombreFinal);
  const file = carpeta.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();

  const hoy = new Date();
  // Vencimiento: Senior o mayor de 40 → 1 año; resto → no vence (ingreso)
  let vencimiento = '';
  if (esSenior === 'true' || esSenior === true) {
    const v = new Date(hoy);
    v.setFullYear(v.getFullYear() + 1);
    vencimiento = v;
  }

  const ws = SS.getSheetByName(TABS.estudiosMedicos);
  if (!ws) return { ok: false, msg: 'Hoja EstudiosMedicos no encontrada.' };

  // Buscar si ya existe para este jugador → actualizar
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      ws.getRange(i+1, 3).setValue(hoy);
      ws.getRange(i+1, 4).setValue('Vigente');
      ws.getRange(i+1, 5).setValue(vencimiento || '');
      ws.getRange(i+1, 6).setValue(fileId);
      ws.getRange(i+1, 7).setValue((esSenior==='true'||esSenior===true)?'X':'');
      return { ok: true, msg: 'Estudio médico actualizado.' };
    }
  }
  // Nuevo registro
  ws.appendRow([
    String(jugadorId).trim(),
    String(equipoId).trim(),
    hoy,
    'Vigente',
    vencimiento || '',
    fileId,
    (esSenior==='true'||esSenior===true)?'X':'',
  ]);
  return { ok: true, msg: 'Estudio médico registrado.' };
}

// ═══════════════════════════════════════════════════════════════════
// COMPROBANTES
// ═══════════════════════════════════════════════════════════════════

function getComprobantes({ equipoId }) {
  const data = leerComprobantes().filter(c => matchEquipo(c.equipoId, equipoId));
  return { ok: true, comprobantes: data };
}

// Endpoint combinado OPTIMIZADO: lee cada hoja una sola vez con filtrado inline
function getDelegadoInit({ equipoId }) {
  // 1) Parámetros — solo necesitamos fila 2
  let parametros = null;
  try {
    const wsP = SS.getSheetByName(TABS.parametros);
    if (wsP) {
      const fila = wsP.getRange(2, 1, 1, 5).getValues()[0] || [];
      parametros = { ok: true, cuotaSocial: Number(fila[3]||0), cuotaDeportiva: Number(fila[4]||0) };
    }
  } catch(e) {}

  // 2) Jugadores — leer y filtrar en una sola pasada
  const jugadores = [];
  const wsJ = SS.getSheetByName(TABS.jugadores);
  if (wsJ) {
    const dataJ = wsJ.getDataRange().getValues();
    for (let i = 1; i < dataJ.length; i++) {
      const r = dataJ[i];
      if (!r[0] || !matchEquipo(r[1], equipoId)) continue;
      jugadores.push({
        id:        String(r[0]).trim(),
        equipoId:  String(r[1]).trim(),
        nombre:    String(r[2]||'').trim(),
        mas35:     String(r[3]||'').toUpperCase() === 'X',
        ic:        String(r[4]||'').toUpperCase() === 'X',
        dni:       String(r[5]||'').trim(),
        fechaNac:  r[6] ? formatFecha(r[6]) : '',
        tipoSocio: String(r[7]||'Activo').trim(),
        habManual: String(r[8]||'').trim(),
      });
    }
    jugadores.sort((a,b) => a.nombre.localeCompare(b.nombre, 'es'));
  }

  // 3) Comprobantes — leer y filtrar por equipo inline (evita parsear filas irrelevantes)
  const comprobantes = [];
  const wsC = SS.getSheetByName(TABS.comprobantes);
  if (wsC) {
    const dataC = wsC.getDataRange().getValues();
    for (let i = 2; i < dataC.length; i++) {
      const r = dataC[i];
      if (!r[0] || r[0] === 'ResumenBot') continue;
      if (!matchEquipo(r[2], equipoId)) continue; // filtrar temprano
      const archivo = String(r[6]||'');
      comprobantes.push({
        id:            String(r[0]),
        equipoId:      r[2],
        jugadorId:     String(r[3]||'').trim(),
        jugadorNombre: r[4],
        mes:           r[5],
        archivo,
        esImagen:      /\.(jpg|jpeg|png)$/i.test(archivo),
        fecha:         r[7] ? new Date(r[7]).toISOString() : null,
        estado:        r[8] || 'Pendiente',
        obsAdmin:      r[9] || '',
        esSubsanacion: r[10],
        obsUsuario:    r[11] || '',
        fileId:        String(r[12]||''),
      });
    }
  }

  return { ok: true, jugadores, comprobantes, parametros };
}

function getAllComprobantes({ mes, estado, equipoId }) {
  let data = leerComprobantes().filter(c => c.estado !== 'Reemplazado');
  if (mes)    data = data.filter(c => c.mes === mes);
  if (estado && estado !== 'todos') data = data.filter(c => c.estado === estado);
  if (equipoId) data = data.filter(c => matchEquipo(c.equipoId, equipoId));
  data.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));
  return { ok: true, comprobantes: data };
}

function getComprobante({ id }) {
  const data = leerComprobantes();
  const c = data.find(c => c.id === id);
  if (!c) return { ok: false, msg: 'Comprobante no encontrado' };
  if (c.fileId) {
    try {
      const file = DriveApp.getFileById(c.fileId);
      const blob  = file.getBlob();
      const mime  = blob.getContentType();
      c.mimeType    = mime;
      c.esImagen    = mime.startsWith('image/');
      c.archivoB64  = Utilities.base64Encode(blob.getBytes());
      c.urlDescarga = file.getDownloadUrl();
    } catch(e) { c.archivoB64 = null; }
  }
  return { ok: true, comprobante: c };
}

function leerComprobantes() {
  const ws = SS.getSheetByName(TABS.comprobantes);
  const data = ws.getDataRange().getValues();
  const comps = [];
  for (let i = 2; i < data.length; i++) {
    const r = data[i];
    if (!r[0] || r[0] === 'ResumenBot') continue;
    const archivo = String(r[6]||'');
    comps.push({
      id:            String(r[0]),
      equipoId:      r[2],
      jugadorId:     String(r[3]||'').trim(),
      jugadorNombre: r[4],
      mes:           r[5],
      archivo,
      esImagen:      /\.(jpg|jpeg|png)$/i.test(archivo),
      fecha:         r[7] ? new Date(r[7]).toISOString() : null,
      estado:        r[8] || 'Pendiente',
      obsAdmin:      r[9] || '',
      esSubsanacion: r[10],
      obsUsuario:    r[11] || '',
      fileId:        String(r[12]||''),
    });
  }
  return comps;
}

function cargarComprobante(params) {
  const { equipoId, jugadorId, jugadorNombre, mes, archivo, nombreArchivo,
          tipoArchivo, obsUsuario, reemplazaId } = params;
  if (!archivo) return { ok: false, msg: 'No se recibio el archivo' };

  const ws = SS.getSheetByName(TABS.comprobantes);
  const data = ws.getDataRange().getValues();

  if (!reemplazaId) {
    for (let i = 2; i < data.length; i++) {
      const r = data[i];
      if (matchEquipo(r[2], equipoId) &&
          String(r[3]).trim()===String(jugadorId).trim() &&
          r[5]===mes && r[0]!=='ResumenBot' &&
          r[8]!=='Rechazado' && r[8]!=='Reemplazado') {
        return { ok: false, msg: 'Ya existe un comprobante para este jugador en ' + mes };
      }
    }
  }

  const carpeta = obtenerCarpeta(equipoId, mes);
  const ext = nombreArchivo.split('.').pop().toLowerCase();
  const uid = Utilities.getUuid().substring(0,8);
  const nombreFinal = `${uid}_${jugadorId}_${mes}.${ext}`;
  const blob = Utilities.newBlob(Utilities.base64Decode(archivo), tipoArchivo, nombreFinal);
  const file = carpeta.createFile(blob);
  const fileId = file.getId();
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  if (reemplazaId) {
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0])===String(reemplazaId)) {
        ws.getRange(i+1,9).setValue('Reemplazado');
        break;
      }
    }
  }

  ws.appendRow([uid,'',equipoId,jugadorId,jugadorNombre,mes,nombreFinal,
    new Date(),'Pendiente','',reemplazaId?'X':'',obsUsuario,fileId]);
  return { ok: true, id: uid };
}

// ═══════════════════════════════════════════════════════════════════
// APROBAR / RECHAZAR / HABILITAR
// ═══════════════════════════════════════════════════════════════════

function aprobar({ id }) {
  return cambiarEstado(id, 'Aprobado', '');
}

function rechazar({ id, obs }) {
  if (!obs) return { ok: false, msg: 'El motivo es requerido' };
  const resultado = cambiarEstado(id, 'Rechazado', obs);
  if (resultado.ok) {
    var waData = null;
    try {
      const c = leerComprobantes().find(x => x.id === id);
      if (c) {
        var equipoNombre = c.equipoId.replace(/_/g,' ');
        var telefono = getTelefonoEquipo(c.equipoId);
        waData = { telefono: telefono, jugador: c.jugadorNombre, equipo: equipoNombre, mes: c.mes, motivo: obs };
        const correo = getCorreoEquipo(c.equipoId);
        if (correo) {
          var asunto = '[CDP] Comprobante rechazado — ' + c.jugadorNombre;
          var textoPlano = 'Estimado Delegado,\n\nLe informamos que el comprobante correspondiente al jugador:\n\n' +
            c.jugadorNombre + ' – Equipo: ' + equipoNombre + ' (mes: ' + c.mes + ') ha sido RECHAZADO.\nMotivo: ' + obs +
            '\n\nLe solicitamos ingresar a la aplicación para cargar un nuevo comprobante y regularizar la situación a la brevedad.\n\nAnte cualquier duda, quedamos a disposición.\n\n---\nSistema de Pagos\nClub de Profesionales Justo José de Urquiza';
          var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.6">' +
            '<p>Estimado Delegado,</p>' +
            '<p>Le informamos que el comprobante correspondiente al jugador:</p>' +
            '<p style="margin:16px 0 16px 20px"><b>' + c.jugadorNombre + ' – Equipo: ' + equipoNombre + ' (mes: ' + c.mes + ') ha sido RECHAZADO.</b><br>' +
            '<b>Motivo: ' + obs + '</b></p>' +
            '<p>Le solicitamos ingresar a la aplicación para cargar un nuevo comprobante y regularizar la situación a la brevedad.</p>' +
            '<p>Ante cualquier duda, quedamos a disposición.</p>' +
            '<hr style="border:none;border-top:1px solid #ccc;margin:24px 0 12px">' +
            '<p style="color:#555;font-size:13px">Sistema de Pagos<br>Club de Profesionales Justo José de Urquiza</p>' +
            '</div>';
          MailApp.sendEmail(correo, asunto, textoPlano, { htmlBody: html });
        }
      }
    } catch(mailErr) { Logger.log('Email error: ' + mailErr); }
    resultado.whatsapp = waData;
  }
  return resultado;
}

function cambiarEstado(id, estado, obs) {
  const ws = SS.getSheetByName(TABS.comprobantes);
  const data = ws.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][0])===String(id)) {
      ws.getRange(i+1,9).setValue(estado);
      ws.getRange(i+1,10).setValue(obs);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'Comprobante no encontrado' };
}

// ═══════════════════════════════════════════════════════════════════
// RESUMEN ADMIN
// ═══════════════════════════════════════════════════════════════════

function getResumen({ mes } = {}) {
  const mesActivo = mes || obtenerMesActual();
  const comps = leerComprobantes().filter(c => c.estado!=='Reemplazado');
  const compsMes = comps.filter(c => c.mes===mesActivo);
  const aprobados  = compsMes.filter(c=>c.estado==='Aprobado').length;
  const rechazados = compsMes.filter(c=>c.estado==='Rechazado').length;
  const pendientes = compsMes.filter(c=>c.estado==='Pendiente').length;
  const total = aprobados+rechazados+pendientes;

  // Solicitudes de alta pendientes
  const wsSol = SS.getSheetByName(TABS.solicitudesAlta);
  let solicitudesPendientes = 0;
  if (wsSol) {
    const dSol = wsSol.getDataRange().getValues();
    for (let i=1;i<dSol.length;i++) if (dSol[i][0] && String(dSol[i][11]||'Pendiente')==='Pendiente') solicitudesPendientes++;
  }

  const wsEq   = SS.getSheetByName(TABS.equipos);
  const dataEq  = wsEq.getDataRange().getValues();
  const wsJug   = SS.getSheetByName(TABS.jugadores);
  const dataJug = wsJug ? wsJug.getDataRange().getValues() : [];

  const porEquipo = [];
  for (let i = 1; i < dataEq.length; i++) {
    const eqId = String(dataEq[i][0]||'').trim();
    if (!eqId) continue;
    const totalJug = dataJug.filter(r => matchEquipo(r[1], eqId) && r[0]).length;
    const eqComps  = compsMes.filter(c => matchEquipo(c.equipoId, eqId));
    porEquipo.push({ id:eqId, nombre:dataEq[i][1], cat:dataEq[i][2], total:totalJug,
      aprobados: eqComps.filter(c=>c.estado==='Aprobado').length,
      pendientes: eqComps.filter(c=>c.estado==='Pendiente').length,
      rechazados: eqComps.filter(c=>c.estado==='Rechazado').length });
  }

  // Sanciones activas desde la hoja Sanciones
  const sancionesData = [];
  const wsSanc = SS.getSheetByName(TABS.sanciones);
  if (wsSanc) {
    const ds = wsSanc.getDataRange().getValues();
    for (let i = 2; i < ds.length; i++) {
      const nombre = String(ds[i][0] || '').trim();
      if (!nombre) continue;
      sancionesData.push({
        nombre: nombre,
        equipo: String(ds[i][1] || '').trim(),
        partidos: String(ds[i][2] || '').trim(),
        articulo: String(ds[i][3] || '').trim(),
        fechaHab: ds[i][4] ? formatFecha(ds[i][4]) : ''
      });
    }
  }

  return { ok: true, total, aprobados, pendientes, rechazados, porEquipo, mes: mesActivo, solicitudesPendientes, sancionesData };
}

// ═══════════════════════════════════════════════════════════════════
// REPORTE SEMANAL (trigger jueves 7AM)
// ═══════════════════════════════════════════════════════════════════
// Para configurar: En Apps Script → Triggers → + Agregar trigger
//   Función: enviarResumenSemanal
//   Tipo: Temporizador semanal → Martes/Miércoles → cualquier hora cercana a las 7AM

function enviarResumenSemanal() {
  const mesActual   = obtenerMesActual();
  const mesAnterior = obtenerMesAnterior();

  const wsEq = SS.getSheetByName(TABS.equipos);
  const dataEq = wsEq.getDataRange().getValues();

  const comps = leerComprobantes().filter(c =>
    (c.mes===mesActual || c.mes===mesAnterior) && c.estado==='Aprobado');
  const jugAdminOk = new Set(comps.map(c => String(c.jugadorId).trim()));

  const sanciones = leerSancionesActivas();

  const sinMedico = new Set();
  const wsEM = SS.getSheetByName(TABS.estudiosMedicos);
  if (wsEM) {
    const dem = wsEM.getDataRange().getValues();
    for (let i=2;i<dem.length;i++) {
      if (!dem[i][0]) continue;
      const est = String(dem[i][3]||'').toLowerCase();
      if (est==='vencido'||est==='no presentado') sinMedico.add(String(dem[i][0]).trim());
    }
  }

  const wsJug = SS.getSheetByName(TABS.jugadores);
  const dataJug = wsJug ? wsJug.getDataRange().getValues() : [];

  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  for (let i = 1; i < dataEq.length; i++) {
    const eqId     = String(dataEq[i][0]||'').trim();
    const eqNombre = String(dataEq[i][1]||'').trim();
    const correo   = String(dataEq[i][3]||'').trim();
    if (!eqId || !correo) continue;

    const jugadores = dataJug.filter(r => r[0] && matchEquipo(r[1], eqId));
    if (!jugadores.length) continue;

    const habilitados   = [], iAdmin = [], iMedica = [], iDeportiva = [];
    jugadores.forEach(r => {
      const jugId    = String(r[0]).trim();
      const nombre   = String(r[2]||'').trim();
      const habManual = String(r[8]||'').toUpperCase();
      if (habManual === 'SI')                                    habilitados.push(nombre);
      else if (habManual === 'NO')                               iAdmin.push(nombre + ' (inh. manual)');
      else if (jugadorSancionado(sanciones, jugId, nombre, String(r[1]||'')))      iDeportiva.push(nombre);
      else if (sinMedico.has(jugId))                             iMedica.push(nombre);
      else if (!jugAdminOk.has(jugId))                           iAdmin.push(nombre);
      else                                                       habilitados.push(nombre);
    });

    const cuerpo = `Estimado delegado de ${eqNombre},\n\n` +
      `Este es el resumen de habilitaciones de tu equipo al ${hoy}:\n\n` +
      `✅ HABILITADOS (${habilitados.length}): ${habilitados.join(', ') || 'ninguno'}\n\n` +
      `💰 I. ADMINISTRATIVA (${iAdmin.length}): ${iAdmin.join(', ') || 'ninguno'}\n\n` +
      `🏥 I. MÉDICA (${iMedica.length}): ${iMedica.join(', ') || 'ninguno'}\n\n` +
      `🚫 I. DEPORTIVA (${iDeportiva.length}): ${iDeportiva.join(', ') || 'ninguno'}\n\n` +
      `Mes de referencia: ${mesActual} (también acepta ${mesAnterior})\n\n` +
      `Club de Profesionales — Sistema de Pagos CDP`;

    try {
      MailApp.sendEmail(correo,
        `[CDP] Estado del plantel — ${eqNombre} — ${hoy}`,
        cuerpo
      );
    } catch(e) { Logger.log('Error email ' + correo + ': ' + e); }
  }
}

// ═══════════════════════════════════════════════════════════════════
// FIXTURE + PLANILLA
// ═══════════════════════════════════════════════════════════════════

function getFixture() {
  const ws = SS.getSheetByName(TABS.fixture);
  if (!ws) return { ok: false, msg: 'Pestana "Fixture" no encontrada.' };
  const data = ws.getDataRange().getValues();
  const partidos = [];
  for (let i = 2; i < data.length; i++) {
    const r = data[i];
    if (!r[7] || String(r[8]).trim().toUpperCase()==='LIBRE') continue;
    let hora = '';
    if (r[5]) {
      if (typeof r[5]==='object' && r[5].getHours) {
        hora = Utilities.formatDate(r[5], 'UTC', 'HH:mm');
      } else hora = String(r[5]);
    }
    let fechaStr='', diaStr='';
    if (r[4]) {
      if (typeof r[4]==='object' && r[4].getDate) {
        fechaStr = Utilities.formatDate(r[4], 'UTC', 'dd/MM/yyyy');
        const dias = ['Domingo','Lunes','Martes','Miercoles','Jueves','Viernes','Sabado'];
        const parsed = new Date(Utilities.formatDate(r[4], 'UTC', 'yyyy-MM-dd') + 'T00:00:00Z');
        diaStr = dias[parsed.getUTCDay()];
      } else fechaStr = String(r[4]);
    }
    if (!diaStr && r[3]) diaStr = String(r[3]);
    partidos.push({
      fechaNro:  String(r[0]||'').trim(), torneo: String(r[1]||'').trim(),
      categoria: String(r[2]||'').trim(), dia: diaStr, fecha: fechaStr,
      hora, cancha: String(r[6]||'').trim(),
      local: String(r[7]||'').trim(), visitante: String(r[8]||'').trim(),
      resultado: String(r[9]||'').trim(), obs: String(r[10]||'').trim(),
    });
  }
  return { ok: true, partidos };
}

function getPlanillaData({ local, visitante, fechaNro, torneo, categoria, fecha, hora, cancha }) {
  const mesActual   = obtenerMesActual();
  const mesAnterior = obtenerMesAnterior();

  const nombreToId = buildNombreToIdMap();
  const localId     = nombreToId[local]     || local;
  const visitanteId = nombreToId[visitante] || visitante;

  const wsJug = SS.getSheetByName(TABS.jugadores);
  if (!wsJug) return { ok: false, msg: 'Pestana Jugadores no encontrada' };
  const dataJug = wsJug.getDataRange().getValues();

  const comps = leerComprobantes().filter(c =>
    (c.mes===mesActual||c.mes===mesAnterior) && c.estado==='Aprobado');
  const jugAdminOk = new Set(comps.map(c => String(c.jugadorId).trim()));

  const sanciones = leerSancionesActivas();

  const sinMedico = new Set();
  const wsEM = SS.getSheetByName(TABS.estudiosMedicos);
  if (wsEM) {
    const dem = wsEM.getDataRange().getValues();
    for (let i=2;i<dem.length;i++) {
      if (!dem[i][0]) continue;
      const est = String(dem[i][3]||'').toLowerCase();
      if (est==='vencido'||est==='no presentado') sinMedico.add(String(dem[i][0]).trim());
    }
  }

  function determinarEstado(jugId, jugNombre, jugEquipo, habManual) {
    if (habManual === 'SI') return 'Habilitado';
    if (habManual === 'NO') return 'I. Administrativa';
    if (jugadorSancionado(sanciones, jugId, jugNombre, jugEquipo)) return 'I. Deportiva';
    if (sinMedico.has(jugId))   return 'I. Medica';
    if (!jugAdminOk.has(jugId)) return 'I. Administrativa';
    return 'Habilitado';
  }

  function getJugadoresEquipo(equipoId) {
    const jug = [];
    for (let i=1;i<dataJug.length;i++) {
      const r = dataJug[i];
      if (!r[0]||!matchEquipo(r[1], equipoId)) continue;
      const jugId = String(r[0]).trim();
      const habManual = String(r[8]||'').trim().toUpperCase();
      jug.push({ id:jugId, nombre:String(r[2]||'').trim(),
        mas35:String(r[3]||'').toUpperCase()==='X', ic:String(r[4]||'').toUpperCase()==='X',
        estado:determinarEstado(jugId, String(r[2]||'').trim(), String(r[1]||'').trim(), habManual) });
    }
    jug.sort((a,b)=>a.nombre.localeCompare(b.nombre,'es'));
    return jug;
  }

  return {
    ok: true,
    partido: { fechaNro, torneo, categoria, fecha, hora, cancha, local, visitante },
    jugadoresLocal:     getJugadoresEquipo(localId),
    jugadoresVisitante: getJugadoresEquipo(visitanteId),
    mesActual, localId, visitanteId,
  };
}

// ═══════════════════════════════════════════════════════════════════
// SALUD DE JUGADORES
// ═══════════════════════════════════════════════════════════════════
// Cols SaludJugadores: A=JugadorID | B=EquipoID | C=NombreJugador | D=GrupoSanguineo
//   E=CondicionesMedicas | F=Medicacion | G=ContactoNombre | H=ContactoTel
//   I=ObraSocial | J=FechaActualizacion

function getSaludJugador({ jugadorId }) {
  const ws = SS.getSheetByName(TABS.saludJugadores);
  if (!ws) return { ok: true, salud: null };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      return { ok: true, salud: filaToSalud(data[i]) };
    }
  }
  return { ok: true, salud: null };
}

function getSaludEquipo({ equipoId }) {
  const ws = SS.getSheetByName(TABS.saludJugadores);
  if (!ws) return { ok: true, salud: [] };
  const data = ws.getDataRange().getValues();
  const salud = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (matchEquipo(data[i][1], equipoId)) {
      salud.push(filaToSalud(data[i]));
    }
  }
  return { ok: true, salud };
}

function getSaludAdmin({ equipoId } = {}) {
  const ws = SS.getSheetByName(TABS.saludJugadores);
  if (!ws) return { ok: true, salud: [] };
  const data = ws.getDataRange().getValues();
  const salud = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (equipoId && !matchEquipo(data[i][1], equipoId)) continue;
    salud.push(filaToSalud(data[i]));
  }
  salud.sort((a,b) => String(a.equipoId).localeCompare(String(b.equipoId)) || a.nombre.localeCompare(b.nombre,'es'));
  return { ok: true, salud };
}

function filaToSalud(r) {
  return {
    jugadorId:          String(r[0]).trim(),
    equipoId:           String(r[1]).trim(),
    nombre:             String(r[2]||'').trim(),
    grupoSanguineo:     String(r[3]||'').trim(),
    condicionesMedicas: String(r[4]||'').trim(),
    medicacion:         String(r[5]||'').trim(),
    contactoNombre:     String(r[6]||'').trim(),
    contactoTel:        String(r[7]||'').trim(),
    obraSocial:         String(r[8]||'').trim(),
    fechaActualizacion: r[9] ? formatFecha(r[9]) : '',
  };
}

function setSaludJugador({ jugadorId, equipoId, nombre, grupoSanguineo, condicionesMedicas,
                            medicacion, contactoNombre, contactoTel, obraSocial }) {
  const ws = SS.getSheetByName(TABS.saludJugadores);
  if (!ws) return { ok: false, msg: 'Hoja SaludJugadores no encontrada. Ejecutá setupHojasNuevas().' };
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jugadorId).trim()) {
      ws.getRange(i+1, 3).setValue(String(nombre||'').trim().toUpperCase());
      ws.getRange(i+1, 4).setValue(String(grupoSanguineo||'').trim());
      ws.getRange(i+1, 5).setValue(String(condicionesMedicas||'').trim());
      ws.getRange(i+1, 6).setValue(String(medicacion||'').trim());
      ws.getRange(i+1, 7).setValue(String(contactoNombre||'').trim());
      ws.getRange(i+1, 8).setValue(String(contactoTel||'').trim());
      ws.getRange(i+1, 9).setValue(String(obraSocial||'').trim());
      ws.getRange(i+1,10).setValue(new Date());
      return { ok: true, msg: 'Información de salud actualizada.' };
    }
  }
  // Nueva fila
  ws.appendRow([
    String(jugadorId).trim(),
    String(equipoId).trim(),
    String(nombre||'').trim().toUpperCase(),
    String(grupoSanguineo||'').trim(),
    String(condicionesMedicas||'').trim(),
    String(medicacion||'').trim(),
    String(contactoNombre||'').trim(),
    String(contactoTel||'').trim(),
    String(obraSocial||'').trim(),
    new Date(),
  ]);
  return { ok: true, msg: 'Información de salud guardada.' };
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════

function obtenerCarpeta(equipoId, mes) {
  const base = 'PagosComprobantes_CDP';
  const iter = DriveApp.getFoldersByName(base);
  const carpetaBase = iter.hasNext() ? iter.next() : DriveApp.createFolder(base);
  const iterEq = carpetaBase.getFoldersByName(equipoId);
  const carpetaEq = iterEq.hasNext() ? iterEq.next() : carpetaBase.createFolder(equipoId);
  const iterMes = carpetaEq.getFoldersByName(mes);
  return iterMes.hasNext() ? iterMes.next() : carpetaEq.createFolder(mes);
}

function obtenerCarpetaEstudios(equipoId) {
  const base = 'EstudiosMedicos_CDP';
  const iter = DriveApp.getFoldersByName(base);
  const carpetaBase = iter.hasNext() ? iter.next() : DriveApp.createFolder(base);
  return obtenerSubcarpeta(carpetaBase, equipoId);
}

function obtenerSubcarpeta(carpetaPadre, nombre) {
  const iter = carpetaPadre.getFoldersByName(nombre);
  return iter.hasNext() ? iter.next() : carpetaPadre.createFolder(nombre);
}

function getCorreoEquipo(equipoId) {
  const ws = SS.getSheetByName(TABS.equipos);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (matchEquipo(data[i][0], equipoId)) return String(data[i][3]||'').trim();
  }
  return '';
}

// Columna F (índice 5) de Equipos = TelefonoDelegado (formato: 5493442XXXXXX sin +)
function getTelefonoEquipo(equipoId) {
  const ws = SS.getSheetByName(TABS.equipos);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (matchEquipo(data[i][0], equipoId)) {
      var tel = String(data[i][5]||'').replace(/[^0-9]/g, '');
      return tel || '';
    }
  }
  return '';
}

function buildNombreToIdMap() {
  const wsEq = SS.getSheetByName(TABS.equipos);
  const data = wsEq.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]||'').trim();
    const nombre = String(data[i][1]||'').trim();
    if (id && nombre) map[nombre] = id;
  }
  return map;
}

function obtenerMesActual() {
  const m = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return m[new Date().getMonth()];
}

function obtenerMesAnterior() {
  const m = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const i = new Date().getMonth();
  return m[i===0?11:i-1];
}

function formatFecha(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    return String(dt.getDate()).padStart(2,'0')+'/'+String(dt.getMonth()+1).padStart(2,'0')+'/'+dt.getFullYear();
  } catch(e) { return String(d); }
}

// ═══════════════════════════════════════════════════════════════════
// SETUP (ejecutar una sola vez)
// ═══════════════════════════════════════════════════════════════════

function setupPINs() {
  const ws = SS.getSheetByName(TABS.equipos);
  const data = ws.getDataRange().getValues();
  if (data[0][4]!=='PIN') ws.getRange(1,5).setValue('PIN');
  for (let i=1;i<data.length;i++) if (!data[i][4]&&data[i][0]) ws.getRange(i+1,5).setValue('1234');
  return 'PINs configurados en columna E del sheet Equipos.';
}

function setupHojasNuevas() {
  // Crea las hojas nuevas si no existen
  const hojas = [
    { nombre: TABS.solicitudesAlta,
      headers: ['ID','EquipoID','NombreEquipo','Nombre','DNI','FechaNac','TipoSocio','Mas35','IC','TipoTitulo','FechaSolicitud','Estado','ObsAdmin','DriveFolder','FechaResolucion'] },
    { nombre: TABS.historialBajas,
      headers: ['JugadorID','EquipoID','Nombre','FechaBaja','UltimoMesPago','Motivo'] },
    { nombre: TABS.saludJugadores,
      headers: ['JugadorID','EquipoID','NombreJugador','GrupoSanguineo','CondicionesMedicas','Medicacion','ContactoNombre','ContactoTel','ObraSocial','FechaActualizacion'] },
  ];
  const resultados = [];
  hojas.forEach(h => {
    let ws = SS.getSheetByName(h.nombre);
    if (!ws) {
      ws = SS.insertSheet(h.nombre);
      ws.getRange(1,1,1,h.headers.length).setValues([h.headers]);
      resultados.push('Creada: ' + h.nombre);
    } else {
      resultados.push('Ya existe: ' + h.nombre);
    }
  });

  // Asegurarse que Parametros tenga columnas de cuotas en D y E
  const wsP = SS.getSheetByName(TABS.parametros);
  if (wsP) {
    const headers = wsP.getRange(1,1,1,5).getValues()[0];
    if (!headers[3]) wsP.getRange(1,4).setValue('CuotaSocial');
    if (!headers[4]) wsP.getRange(1,5).setValue('CuotaDeportiva');
  }

  return resultados.join('\n');
}

function setupFileIdColumn() {
  const ws = SS.getSheetByName(TABS.comprobantes);
  const data = ws.getDataRange().getValues();
  if (data[0][12]!=='FileId') ws.getRange(1,13).setValue('FileId');
  return 'Columna FileId verificada.';
}

// Debug fixture (sin cambios)
function debugFixture() {
  const ws = SS.getSheetByName('Fixture');
  const data = ws.getDataRange().getValues();
  const r = data[2];
  const tz = Session.getScriptTimeZone();
  Logger.log('Timezone: ' + tz);
  Logger.log('Local: ' + r[7] + ' vs ' + r[8]);
}
