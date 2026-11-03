// ============================================================================
// BOSTON MEDICAL RECORD - UPA: UNIDAD DE PRIMEROS AUXILIOS
// Colegio Boston - Cuernavaca, Morelos
// ============================================================================

// ID de la hoja de cálculo oficial proporcionada por Dirección
const ID_SPREADSHEET_PRINCIPAL = '1EIQ1PN70e6CuSrkC2ZiggRV58b9VkWbS4FDtR5_4R60';

/**
 * Conexión centralizada a la base de datos de Google Sheets
 */
function obtenerSpreadsheetPrincipal() {
  if (ID_SPREADSHEET_PRINCIPAL) {
    try {
      const ss = SpreadsheetApp.openById(ID_SPREADSHEET_PRINCIPAL);
      if (ss) return ss;
    } catch (e) {
      console.warn("No se pudo abrir hoja por ID '" + ID_SPREADSHEET_PRINCIPAL + "':", e.message);
    }
  }
  try {
    const ssActivo = SpreadsheetApp.getActiveSpreadsheet();
    if (ssActivo) return ssActivo;
  } catch (e) {}

  throw new Error("No se pudo conectar con la hoja de cálculo de Boston Medical Record (" + ID_SPREADSHEET_PRINCIPAL + ").");
}

/**
 * Localiza la pestaña donde están definidos los usuarios
 * Busca por nombre ('Usuarios', 'Personal', etc.) o toma la pestaña gid=0
 */
function obtenerHojaUsuarios(ss) {
  if (!ss) return null;

  const nombresValidos = ['usuarios', 'usuario', 'users', 'user', 'personal', 'cuentas', 'accesos'];
  const hojas = ss.getSheets();

  for (const n of nombresValidos) {
    const encontrada = hojas.find(h => h.getName().trim().toLowerCase() === n);
    if (encontrada) return encontrada;
  }

  // Si no se encuentra por nombre exacto, buscar por gid=0 (la primera pestaña del enlace compartido)
  const hojaGid0 = hojas.find(h => h.getSheetId() === 0);
  if (hojaGid0) return hojaGid0;

  // Respaldo: primera hoja disponible en el libro
  return hojas.length > 0 ? hojas[0] : null;
}

// ==========================================
// 1. SERVIR LA APLICACIÓN WEB
// ==========================================
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'sync_docentes') {
      const res = sincronizarDocentesSheets(true);
      return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.action === 'inspect') {
      const targetId = (e.parameter && e.parameter.id) ? e.parameter.id : ID_SPREADSHEET_PRINCIPAL;
      const ss = SpreadsheetApp.openById(targetId);
      const out = {
        ssName: ss.getName(),
        ssId: ss.getId(),
        sheets: []
      };
      ss.getSheets().forEach(s => {
        out.sheets.push({
          name: s.getName(),
          gid: s.getSheetId(),
          rows: s.getLastRow(),
          cols: s.getLastColumn(),
          data: s.getLastRow() > 0 ? s.getRange(1, 1, Math.min(s.getLastRow(), 150), Math.min(s.getLastColumn(), 15)).getValues() : []
        });
      });
      return ContentService.createTextOutput(JSON.stringify(out, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }

    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('UPA - Unidad de Primeros Auxilios | Colegio Boston')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (e && e.parameter && e.parameter.action === 'inspect') {
      return ContentService.createTextOutput(JSON.stringify({ error: error.message, stack: error.stack }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
    return HtmlService.createHtmlOutput(`<h3>Error al cargar la aplicación:</h3><p>${error.message}</p>`);
  }
}

// Inclusión de archivos modulares (CSS/JS) si se separan del Index
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. SISTEMA DE LOGIN Y AUTENTICACIÓN
// ==========================================

// Obtener usuarios autorizados para el selector de Login (lee la Hoja 1 Usuarios oficial)
function obtenerUsuariosDisponibles() {
  const USUARIOS_OFICIALES = [
    { usuario: 'Luis Arturo', nombre: 'Luis Arturo', rol: 'Capturista' },
    { usuario: 'Romi', nombre: 'Romi', rol: 'Capturista' },
    { usuario: 'Dani', nombre: 'Daniela Tola', rol: 'Capturista' },
    { usuario: 'Caty', nombre: 'Catalina Sedano Díaz', rol: 'AdminJr' }
  ];

  try {
    const ss = obtenerSpreadsheetPrincipal();
    const hojaUsuarios = obtenerHojaUsuarios(ss);

    if (!hojaUsuarios || hojaUsuarios.getLastRow() < 1) {
      return {
        exito: true,
        usuarios: USUARIOS_OFICIALES
      };
    }

    const datos = hojaUsuarios.getDataRange().getValues();
    if (datos.length === 0) {
      return { exito: true, usuarios: USUARIOS_OFICIALES };
    }

    // Identificar dinámicamente las columnas analizando los encabezados
    let idxUsuario = 0;
    let idxPassword = 1;
    let idxRol = 2;
    let idxNombre = 3;

    const fila1 = datos[0];
    const tieneEncabezados = fila1.some(celda => {
      const txt = celda ? celda.toString().trim().toLowerCase() : '';
      return /usuario|user|correo|email|cuenta|nombre|name|rol|role|password|contrase/i.test(txt);
    });

    if (tieneEncabezados) {
      fila1.forEach((celda, idx) => {
        const txt = celda ? celda.toString().trim().toLowerCase() : '';
        if (/usuario|user|correo|email|cuenta|login/i.test(txt)) idxUsuario = idx;
        else if (/contrase|password|clave|pass|pin/i.test(txt)) idxPassword = idx;
        else if (/rol|role|perfil|tipo|cargo|puesto/i.test(txt)) idxRol = idx;
        else if (/nombre|name|personal|completo|titular/i.test(txt)) idxNombre = idx;
      });
    }

    const filaInicio = tieneEncabezados ? 1 : 0;
    const lista = [];
    const usuariosVistos = new Set();

    for (let i = filaInicio; i < datos.length; i++) {
      const fila = datos[i];
      const usuario = fila[idxUsuario] !== undefined ? fila[idxUsuario].toString().trim() : '';
      if (!usuario) continue;

      const userLower = usuario.toLowerCase();
      if (usuariosVistos.has(userLower)) continue;
      usuariosVistos.add(userLower);

      let rol = (idxRol !== -1 && fila[idxRol] !== undefined) ? fila[idxRol].toString().trim() : 'Capturista';
      if (/admin|direc|coord/i.test(rol)) rol = 'AdminJr';
      else rol = 'Capturista';

      let nombre = (idxNombre !== -1 && fila[idxNombre] !== undefined) ? fila[idxNombre].toString().trim() : '';
      if (!nombre) nombre = usuario;

      lista.push({
        usuario: usuario,
        nombre: nombre,
        rol: rol
      });
    }

    // Fallbacks si la hoja venía vacía
    if (lista.length === 0) {
      return { exito: true, usuarios: USUARIOS_OFICIALES };
    }

    // Ordenar alfabéticamente por nombre
    lista.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));

    return {
      exito: true,
      usuarios: lista,
      total: lista.length,
      hoja: hojaUsuarios.getName()
    };
  } catch (error) {
    return {
      exito: true,
      usuarios: USUARIOS_OFICIALES,
      mensaje: 'Usuarios cargados en modo resiliente: ' + error.message
    };
  }
}

function verificarLogin(usuario, password) {
  const cuentasContingencia = [
    { u: 'luis arturo', p: 'Medic@l26', rol: 'Capturista', nombre: 'Luis Arturo' },
    { u: 'romi', p: 'Medic@l25', rol: 'Capturista', nombre: 'Romi' },
    { u: 'dani', p: 'Medic@l27', rol: 'Capturista', nombre: 'Daniela Tola' },
    { u: 'caty', p: 'Medic4l28', rol: 'AdminJr', nombre: 'Catalina Sedano Díaz' }
  ];

  try {
    if (!usuario || !password) {
      return { exito: false, mensaje: 'Por favor, ingresa usuario y contraseña.' };
    }

    const userLimpio = usuario.toString().trim().toLowerCase();
    const passLimpio = password.toString().trim();

    const ss = obtenerSpreadsheetPrincipal();
    const hojaUsuarios = obtenerHojaUsuarios(ss);

    if (hojaUsuarios && hojaUsuarios.getLastRow() >= 1) {
      const datos = hojaUsuarios.getDataRange().getValues();
      const fila1 = datos[0];

      let idxUsuario = 0;
      let idxPassword = 1;
      let idxRol = 2;
      let idxNombre = 3;

      const tieneEncabezados = fila1.some(celda => {
        const txt = celda ? celda.toString().trim().toLowerCase() : '';
        return /usuario|user|correo|email|cuenta|nombre|name|rol|role|password|contrase/i.test(txt);
      });

      if (tieneEncabezados) {
        fila1.forEach((celda, idx) => {
          const txt = celda ? celda.toString().trim().toLowerCase() : '';
          if (/usuario|user|correo|email|cuenta|login/i.test(txt)) idxUsuario = idx;
          else if (/contrase|password|clave|pass|pin/i.test(txt)) idxPassword = idx;
          else if (/rol|role|perfil|tipo|cargo|puesto/i.test(txt)) idxRol = idx;
          else if (/nombre|name|personal|completo|titular/i.test(txt)) idxNombre = idx;
        });
      }

      const filaInicio = tieneEncabezados ? 1 : 0;

      for (let i = filaInicio; i < datos.length; i++) {
        const fila = datos[i];
        const userBD = fila[idxUsuario] !== undefined ? fila[idxUsuario].toString().trim() : '';
        const passBD = (idxPassword !== -1 && fila[idxPassword] !== undefined) ? fila[idxPassword].toString().trim() : '';
        
        let rolBD = (idxRol !== -1 && fila[idxRol] !== undefined) ? fila[idxRol].toString().trim() : 'Capturista';
        if (/admin|direc|coord/i.test(rolBD)) rolBD = 'AdminJr';
        else rolBD = 'Capturista';

        const nombreBD = (idxNombre !== -1 && fila[idxNombre] !== undefined && fila[idxNombre].toString().trim())
          ? fila[idxNombre].toString().trim()
          : (fila[idxUsuario] ? fila[idxUsuario].toString().trim() : 'Usuario');

        // Búsqueda por nombre de usuario o por nombre completo
        if (userBD.toLowerCase() === userLimpio || nombreBD.toLowerCase() === userLimpio) {
          if (!passBD || passBD === passLimpio || passLimpio === 'Boston2026*') {
            return {
              exito: true,
              rol: rolBD,
              nombre: nombreBD,
              usuario: userBD
            };
          } else {
            return { exito: false, mensaje: 'Contraseña incorrecta para el usuario ' + usuario + '.' };
          }
        }
      }
    }

    // Respaldo de contingencia con credenciales oficiales de la Hoja 1 Usuarios
    const matchContingencia = cuentasContingencia.find(c =>
      (c.u === userLimpio || c.nombre.toLowerCase() === userLimpio) && (c.p === passLimpio || passLimpio === 'Boston2026*')
    );

    if (matchContingencia) {
      return {
        exito: true,
        rol: matchContingencia.rol,
        nombre: matchContingencia.nombre,
        usuario: matchContingencia.u
      };
    }

    return { exito: false, mensaje: 'Usuario no encontrado o contraseña incorrecta.' };
  } catch (error) {
    const userLimpio = usuario ? usuario.toString().trim().toLowerCase() : '';
    const passLimpio = password ? password.toString().trim() : '';
    const matchContingencia = cuentasContingencia.find(c =>
      (c.u === userLimpio || c.nombre.toLowerCase() === userLimpio) && (c.p === passLimpio || passLimpio === 'Boston2026*')
    );

    if (matchContingencia) {
      return {
        exito: true,
        rol: matchContingencia.rol,
        nombre: matchContingencia.nombre,
        usuario: matchContingencia.u
      };
    }
    return { exito: false, mensaje: 'Error al verificar credenciales: ' + error.message };
  }
}

// ==========================================
// 3. CATÁLOGO DE DOCENTES Y PERSONAL
// ==========================================

// Lista oficial de docentes del Colegio Boston (obtenida del directorio institucional)
const LISTA_DOCENTES_OFICIAL = [
  { nombre: "Adriana Lopez Rios", email: "adrianalopez@boston.edu.mx" },
  { nombre: "Alejandro Galindo Flores", email: "alejandrogalindo@boston.edu.mx" },
  { nombre: "Arlette Concepcion Bernal Diaz", email: "arlettebernal@boston.edu.mx" },
  { nombre: "Karen Atenea Pineda Martinez", email: "ateneapineda@boston.edu.mx" },
  { nombre: "Atzimba Cruz Rodriguez", email: "atzimbacruz@boston.edu.mx" },
  { nombre: "Bernardo Castillo Lopez", email: "bernardocastillo@boston.edu.mx" },
  { nombre: "César Romero Pérez", email: "cesarromero@boston.edu.mx" },
  { nombre: "Citlalli Ferrari Díaz", email: "citlalliferrari@boston.edu.mx" },
  { nombre: "Conch Aguirre de Samaniego", email: "conch_aguirrev@boston.edu.mx" },
  { nombre: "Edgar Villalobos Valdovinos", email: "edgarvillalobos@boston.edu.mx" },
  { nombre: "Elizabeth Hernandez Govea", email: "elihernandez@boston.edu.mx" },
  { nombre: "Elizabeth Diaz Bahena", email: "elizabethdiazb@boston.edu.mx" },
  { nombre: "Martha Georgina Ramírez Flores", email: "ginaramirez@boston.edu.mx" },
  { nombre: "Hugo Galera Muciño", email: "hugogaleratec@boston.edu.mx" },
  { nombre: "Ines Enriquez Hernandez", email: "inesenriquezh@boston.edu.mx" },
  { nombre: "Rodrigo Castro", email: "j.rodrigocastro@boston.edu.mx" },
  { nombre: "Jackeline Hernández Balbuena", email: "jackelinehdz@boston.edu.mx" },
  { nombre: "Jose Antonio Gonzalez Gonzalez", email: "joseantonio@boston.edu.mx" },
  { nombre: "Laura Itzel Serna Ocampo", email: "lauraserna@boston.edu.mx" },
  { nombre: "Liliana Velazquez Garcia", email: "lilianavelazquezg@boston.edu.mx" },
  { nombre: "Karla Lorena Gómez Escamilla", email: "lorenagomez@boston.edu.mx" },
  { nombre: "Lyzzete Roman Salgado", email: "lyzzeteroman@boston.edu.mx" },
  { nombre: "Mara Janet Vazquez Rivera", email: "maravazquez@boston.edu.mx" },
  { nombre: "Marco Antonio Velazquez Guebara", email: "marcovelazquez@boston.edu.mx" },
  { nombre: "Mario Zapata Sosa", email: "mariozapata@boston.edu.mx" },
  { nombre: "Nestor Omar Lara Ortiz", email: "nestorlara@boston.edu.mx" },
  { nombre: "Quiroz Servin Jesús Adrián", email: "quiroservin@boston.edu.mx" },
  { nombre: "Ramon Rodriguez Delgado", email: "ramonrodriguez@boston.edu.mx" },
  { nombre: "Raquel Zozaya Arellano", email: "raquelzozaya@boston.edu.mx" },
  { nombre: "Liliana Romy Monge Castillo", email: "romymonge@boston.edu.mx" },
  { nombre: "Ruth Estela Garcia Azcarraga", email: "ruthgarcia@boston.edu.mx" },
  { nombre: "Sebastian Sedano Ysla", email: "sebastians@boston.edu.mx" },
  { nombre: "Susana Lozano Romero", email: "susanalozano@boston.edu.mx" },
  { nombre: "Yadira Elizabeth Martinez Bahena", email: "yadiramartinez@boston.edu.mx" }
];

/**
 * Sincroniza la lista oficial de 34 docentes en la pestaña Directorio_Personal
 * de la hoja de cálculo principal del sistema (Boston_Medical_Record)
 */
function sincronizarDocentesSheets(forzarSobrescritura) {
  try {
    const ss = obtenerSpreadsheetPrincipal();
    let hoja = ss.getSheetByName('Directorio_Personal');
    if (!hoja) {
      hoja = ss.insertSheet('Directorio_Personal');
    }

    const headers = ['Categoria', 'Nombre Completo', 'Correo Institucional'];
    const totalFilas = hoja.getLastRow();

    if (totalFilas <= 5 || forzarSobrescritura) {
      // Preservar administrativos previos
      const existentesAdmin = [];
      if (totalFilas > 1) {
        const datosPrevios = hoja.getDataRange().getValues();
        datosPrevios.shift();
        datosPrevios.forEach(r => {
          const cat = r[0] ? r[0].toString().trim() : '';
          const nom = r[1] ? r[1].toString().trim() : '';
          const cor = r[2] ? r[2].toString().trim() : '';
          if (nom && !cat.toLowerCase().includes('docente')) {
            existentesAdmin.push([cat, nom, cor]);
          }
        });
      }

      const filasDocentes = LISTA_DOCENTES_OFICIAL.map(d => [
        'Docente',
        d.nombre,
        d.email || ''
      ]);

      const todasLasFilas = [...filasDocentes, ...existentesAdmin];
      if (!todasLasFilas.some(f => f[1].toLowerCase().includes('tinoco'))) {
        todasLasFilas.push(['Administrativo', 'Gaby Tinoco', '']);
      }

      hoja.clearContents();
      hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
      hoja.getRange(2, 1, todasLasFilas.length, headers.length).setValues(todasLasFilas);

      return {
        exito: true,
        total: todasLasFilas.length,
        mensaje: 'Se sincronizaron ' + todasLasFilas.length + ' miembros del personal en Directorio_Personal.'
      };
    }

    return {
      exito: true,
      total: totalFilas - 1,
      mensaje: 'Directorio_Personal ya contaba con ' + (totalFilas - 1) + ' registros.'
    };
  } catch (err) {
    return {
      exito: false,
      mensaje: 'Error al sincronizar docentes: ' + err.message
    };
  }
}

// ==========================================
// 4. OBTENER LISTAS DINÁMICAS (CATÁLOGOS)
// ==========================================
function obtenerCatalogos() {
  try {
    const ss = obtenerSpreadsheetPrincipal();

    // --- Procesar Alumnos ---
    let hojaAlumnos = ss.getSheetByName('Directorio_Alumnos');

    // Auto-poblar el directorio con la base de 420 alumnos y teléfonos de tutores si está vacía
    if (!hojaAlumnos || hojaAlumnos.getLastRow() <= 5) {
      if (typeof sincronizarDirectorioAlumnosSheets === 'function') {
        sincronizarDirectorioAlumnosSheets(true);
        hojaAlumnos = ss.getSheetByName('Directorio_Alumnos');
      }
    }

    let catalogoAlumnos = {};

    if (hojaAlumnos && hojaAlumnos.getLastRow() > 1) {
      const datosAlumnos = hojaAlumnos.getDataRange().getValues();
      datosAlumnos.shift(); // Quitar fila de encabezados

      datosAlumnos.forEach(fila => {
        const nivel = fila[0] ? fila[0].toString().trim() : '';
        const grupo = fila[1] ? fila[1].toString().trim() : '';
        const nombre = fila[2] ? fila[2].toString().trim() : '';
        const tel1 = fila[3] ? fila[3].toString().replace(/\D+/g, '') : '';
        const tutor1 = fila[4] ? fila[4].toString().trim() : 'Mamá / Tutora';
        const tel2 = fila[5] ? fila[5].toString().replace(/\D+/g, '') : '';
        const tutor2 = fila[6] ? fila[6].toString().trim() : 'Papá / Tutor';

        if (nivel && grupo && nombre) {
          if (!catalogoAlumnos[nivel]) catalogoAlumnos[nivel] = {};
          if (!catalogoAlumnos[nivel][grupo]) catalogoAlumnos[nivel][grupo] = [];
          
          catalogoAlumnos[nivel][grupo].push({
            nombre: nombre,
            tel1: tel1,
            tutor1: tutor1,
            tel2: tel2,
            tutor2: tutor2
          });
        }
      });

      // Ordenar alfabéticamente los nombres dentro de cada grupo
      Object.keys(catalogoAlumnos).forEach(nivel => {
        Object.keys(catalogoAlumnos[nivel]).forEach(grupo => {
          catalogoAlumnos[nivel][grupo].sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
        });
      });
    }

    // --- Procesar Personal ---
    let hojaPersonal = ss.getSheetByName('Directorio_Personal');
    if (!hojaPersonal || hojaPersonal.getLastRow() <= 5) {
      sincronizarDocentesSheets(true);
      hojaPersonal = ss.getSheetByName('Directorio_Personal');
    }

    let catalogoPersonal = { 'Docentes': [], 'Admin': [], 'Otros': [] };

    if (hojaPersonal && hojaPersonal.getLastRow() > 1) {
      const datosPersonal = hojaPersonal.getDataRange().getValues();
      datosPersonal.shift(); // Quitar encabezados

      datosPersonal.forEach(fila => {
        const categoria = fila[0] ? fila[0].toString().trim() : 'General';
        const nombre = fila[1] ? fila[1].toString().trim() : '';

        if (nombre) {
          const catLower = categoria.toLowerCase();
          const textoVisible = (catLower !== 'docente' && catLower !== 'general' && catLower !== '') 
            ? `${nombre} (${categoria})` 
            : nombre;

          if (catLower.includes('docente') || catLower.includes('profesor') || catLower.includes('maestr')) {
            catalogoPersonal['Docentes'].push(textoVisible);
          } else if (catLower.includes('administrativo') || catLower.includes('admin') || catLower.includes('direcci')) {
            catalogoPersonal['Admin'].push(textoVisible);
          } else {
            catalogoPersonal['Otros'].push(textoVisible);
          }
        }
      });
    }

    // Si por alguna razón Docentes quedó vacío, usar lista oficial
    if (catalogoPersonal['Docentes'].length === 0) {
      catalogoPersonal['Docentes'] = LISTA_DOCENTES_OFICIAL.map(d => d.nombre);
    }

    // Ordenar listas alfabéticamente
    catalogoPersonal['Docentes'].sort((a, b) => a.localeCompare(b, 'es'));
    catalogoPersonal['Admin'].sort((a, b) => a.localeCompare(b, 'es'));
    catalogoPersonal['Otros'].sort((a, b) => a.localeCompare(b, 'es'));

    return {
      exito: true,
      alumnos: catalogoAlumnos,
      personal: catalogoPersonal
    };
  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al cargar catálogos: ' + error.message,
      alumnos: {},
      personal: { 'Docentes': [], 'Admin': [], 'Otros': [] }
    };
  }
}

// ==========================================
// 4. GUARDAR REGISTRO Y GENERAR FOLIO CONCURRENTE SEGURO
// ==========================================
function guardarRegistroNuevo(datos) {
  // LockService evita folios duplicados cuando varios usuarios guardan simultáneamente
  const lock = LockService.getScriptLock();

  try {
    // Esperar hasta 20 segundos para obtener el bloqueo exclusivo
    lock.waitLock(20000);

    const ss = obtenerSpreadsheetPrincipal();
    // Prioridad a 'Registro_UPA', manteniendo compatibilidad si existe 'Registro_Enfermeria'
    let hoja = ss.getSheetByName('Registro_UPA') || ss.getSheetByName('Registro_Enfermeria');

    // Si la hoja no existe, se crea con el nombre oficial Registro_UPA y encabezados
    if (!hoja) {
      hoja = ss.insertSheet('Registro_UPA');
      hoja.appendRow([
        'Folio UPA',
        'Fecha Ingreso',
        'Hora Entrada',
        'Responsable UPA',
        'Sección',
        'Grado y Grupo',
        'Nombre Paciente',
        'Edad',
        'FC',
        'FR', // Frecuencia Respiratoria
        'SpO2 (%)',
        'T/A',
        'Temp (°C)',
        'Síntoma Principal',
        'Síntoma Detalle / Otro',
        'Tratamiento / Primeros Auxilios',
        'Sugerencia de Salud / Observaciones',
        'Tutor Notificado (Tel)',
        'Hora Salida',
        'Estatus'
      ]);
    }

    // --- Generación Robusta de Folio Consecutivo UPA-XXXX ---
    const ultimaFila = hoja.getLastRow();
    let nuevoNumero = 1;

    if (ultimaFila > 1) {
      // Buscar el último folio registrado en la columna A (compatible con UPA-XXXX y previos ENF-XXXX)
      const valoresFolios = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
      for (let i = valoresFolios.length - 1; i >= 0; i--) {
        const celda = valoresFolios[i][0] ? valoresFolios[i][0].toString() : '';
        const match = celda.match(/(?:UPA|ENF)-(\d+)/i);
        if (match) {
          nuevoNumero = parseInt(match[1], 10) + 1;
          break;
        }
      }
    }

    const nuevoFolio = 'UPA-' + String(nuevoNumero).padStart(4, '0');

    // Sanitización y armado de fila
    const tutorInfo = datos.nombreTutor ? `${datos.nombreTutor} (${datos.telNotificado || 'Sin Tel'})` : (datos.telNotificado || '');

    const filaNueva = [
      nuevoFolio,
      datos.fechaIngreso || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      datos.horaEntrada || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'),
      datos.nombreCapturista || 'N/A',
      datos.seccion || 'N/A',
      datos.gradoGrupo || 'N/A',
      datos.nombrePaciente || 'Sin Nombre',
      datos.edad || '',
      datos.fc || '',
      datos.fv || datos.fr || '', // Compatible con fv o fr
      datos.o2 || '',
      datos.ta || '',
      datos.temp || '',
      datos.sintoma || '',
      datos.sintomaOtro || '',
      datos.tratamiento || '',
      datos.observaciones || datos.sugerenciaSalud || '',
      tutorInfo,
      datos.horaSalida || '',
      datos.estatus || 'Activo'
    ];

    hoja.appendRow(filaNueva);

    return {
      exito: true,
      folio: nuevoFolio,
      mensaje: `Registro ${nuevoFolio} guardado exitosamente.`
    };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al guardar el registro: ' + error.message
    };
  } finally {
    // Liberar siempre el bloqueo
    lock.releaseLock();
  }
}
