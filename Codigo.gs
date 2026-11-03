// ============================================================================
// BOSTON MEDICAL RECORD - UPA: UNIDAD DE PRIMEROS AUXILIOS
// Colegio Boston - Cuernavaca, Morelos
// ============================================================================

// ==========================================
// 1. SERVIR LA APLICACIÓN WEB
// ==========================================
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('UPA - Unidad de Primeros Auxilios | Colegio Boston')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
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

// Obtener usuarios autorizados para el selector de Login (evita errores tipográficos)
function obtenerUsuariosDisponibles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaUsuarios = ss.getSheetByName('Usuarios');

    // Búsqueda flexible (insensible a mayúsculas/minúsculas y espacios)
    if (!hojaUsuarios) {
      const hojas = ss.getSheets();
      hojaUsuarios = hojas.find(h => h.getName().trim().toLowerCase() === 'usuarios');
    }

    // Si la hoja no existe o está vacía, crearla y auto-poblarla con cuentas predeterminadas
    if (!hojaUsuarios) {
      hojaUsuarios = ss.insertSheet('Usuarios');
      hojaUsuarios.appendRow(['Usuario', 'Contraseña', 'Rol', 'Nombre Completo']);
      hojaUsuarios.appendRow(['upa', 'Boston2026*', 'Capturista', 'Personal UPA']);
      hojaUsuarios.appendRow(['enfermeria', 'Boston2026*', 'Capturista', 'Enfermería Boston']);
      hojaUsuarios.appendRow(['direccion', 'Admin2026*', 'AdminJr', 'Dirección Académica']);
      hojaUsuarios.appendRow(['admin', 'Admin2026*', 'AdminJr', 'Administrador']);
    } else if (hojaUsuarios.getLastRow() <= 1) {
      if (hojaUsuarios.getLastRow() === 0) {
        hojaUsuarios.appendRow(['Usuario', 'Contraseña', 'Rol', 'Nombre Completo']);
      }
      hojaUsuarios.appendRow(['upa', 'Boston2026*', 'Capturista', 'Personal UPA']);
      hojaUsuarios.appendRow(['enfermeria', 'Boston2026*', 'Capturista', 'Enfermería Boston']);
      hojaUsuarios.appendRow(['direccion', 'Admin2026*', 'AdminJr', 'Dirección Académica']);
      hojaUsuarios.appendRow(['admin', 'Admin2026*', 'AdminJr', 'Administrador']);
    }

    const datos = hojaUsuarios.getDataRange().getValues();
    const lista = [];

    // Omitir encabezados (fila 1 / índice 0)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const usuario = fila[0] ? fila[0].toString().trim() : '';
      const rol = fila[2] ? fila[2].toString().trim() : 'Capturista';
      const nombre = fila[3] ? fila[3].toString().trim() : usuario;

      if (usuario) {
        lista.push({
          usuario: usuario,
          nombre: nombre,
          rol: rol
        });
      }
    }

    if (lista.length === 0) {
      lista.push({ usuario: 'upa', nombre: 'Personal UPA', rol: 'Capturista' });
      lista.push({ usuario: 'enfermeria', nombre: 'Enfermería Boston', rol: 'Capturista' });
      lista.push({ usuario: 'direccion', nombre: 'Dirección Académica', rol: 'AdminJr' });
    }

    // Ordenar alfabéticamente por nombre para facilitar la búsqueda
    lista.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));

    return { exito: true, usuarios: lista };
  } catch (error) {
    // Si ocurre un error, regresar usuarios por defecto para no bloquear el acceso
    return {
      exito: true,
      usuarios: [
        { usuario: 'upa', nombre: 'Personal UPA', rol: 'Capturista' },
        { usuario: 'enfermeria', nombre: 'Enfermería Boston', rol: 'Capturista' },
        { usuario: 'direccion', nombre: 'Dirección Académica', rol: 'AdminJr' }
      ],
      mensaje: 'Usuarios cargados en modo resiliente: ' + error.message
    };
  }
}

function verificarLogin(usuario, password) {
  try {
    if (!usuario || !password) {
      return { exito: false, mensaje: 'Por favor, ingresa usuario y contraseña.' };
    }

    const userLimpio = usuario.toString().trim().toLowerCase();
    const passLimpio = password.toString().trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaUsuarios = ss.getSheetByName('Usuarios');

    if (!hojaUsuarios) {
      const hojas = ss.getSheets();
      hojaUsuarios = hojas.find(h => h.getName().trim().toLowerCase() === 'usuarios');
    }

    // Si aún no existe, inicializarla
    if (!hojaUsuarios || hojaUsuarios.getLastRow() <= 1) {
      obtenerUsuariosDisponibles();
      hojaUsuarios = ss.getSheetByName('Usuarios') || ss.getSheets().find(h => h.getName().trim().toLowerCase() === 'usuarios');
    }

    if (hojaUsuarios && hojaUsuarios.getLastRow() > 1) {
      const datos = hojaUsuarios.getDataRange().getValues();

      // Búsqueda a partir de la fila 2 (índice 1) para omitir encabezados
      for (let i = 1; i < datos.length; i++) {
        const fila = datos[i];
        const userBD = fila[0] ? fila[0].toString().trim().toLowerCase() : '';
        const passBD = fila[1] ? fila[1].toString().trim() : '';
        const rolBD = fila[2] ? fila[2].toString().trim() : 'Capturista';
        const nombreBD = fila[3] ? fila[3].toString().trim() : 'Usuario';

        if (userBD === userLimpio && passBD === passLimpio) {
          return {
            exito: true,
            rol: rolBD,
            nombre: nombreBD,
            usuario: userBD
          };
        }
      }
    }

    // Respaldo de contingencia con credenciales oficiales por defecto
    if ((userLimpio === 'upa' || userLimpio === 'enfermeria') && passLimpio === 'Boston2026*') {
      return {
        exito: true,
        rol: 'Capturista',
        nombre: userLimpio === 'upa' ? 'Personal UPA' : 'Enfermería Boston',
        usuario: userLimpio
      };
    }
    if ((userLimpio === 'direccion' || userLimpio === 'admin') && passLimpio === 'Admin2026*') {
      return {
        exito: true,
        rol: 'AdminJr',
        nombre: userLimpio === 'direccion' ? 'Dirección Académica' : 'Administrador',
        usuario: userLimpio
      };
    }

    return { exito: false, mensaje: 'Usuario o contraseña incorrectos.' };
  } catch (error) {
    const userLimpio = usuario ? usuario.toString().trim().toLowerCase() : '';
    const passLimpio = password ? password.toString().trim() : '';
    if ((userLimpio === 'upa' || userLimpio === 'enfermeria') && passLimpio === 'Boston2026*') {
      return {
        exito: true,
        rol: 'Capturista',
        nombre: userLimpio === 'upa' ? 'Personal UPA' : 'Enfermería Boston',
        usuario: userLimpio
      };
    }
    if ((userLimpio === 'direccion' || userLimpio === 'admin') && passLimpio === 'Admin2026*') {
      return {
        exito: true,
        rol: 'AdminJr',
        nombre: userLimpio === 'direccion' ? 'Dirección Académica' : 'Administrador',
        usuario: userLimpio
      };
    }
    return { exito: false, mensaje: 'Error al verificar credenciales: ' + error.message };
  }
}

// ==========================================
// 3. OBTENER LISTAS DINÁMICAS (CATÁLOGOS)
// ==========================================
function obtenerCatalogos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

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
    const hojaPersonal = ss.getSheetByName('Directorio_Personal');
    let catalogoPersonal = { 'Docentes': [], 'Admin': [], 'Otros': [] };

    if (hojaPersonal && hojaPersonal.getLastRow() > 1) {
      const datosPersonal = hojaPersonal.getDataRange().getValues();
      datosPersonal.shift(); // Quitar encabezados

      datosPersonal.forEach(fila => {
        const categoria = fila[0] ? fila[0].toString().trim() : 'General';
        const nombre = fila[1] ? fila[1].toString().trim() : '';

        if (nombre) {
          const catLower = categoria.toLowerCase();
          if (catLower.includes('docente') || catLower.includes('profesor') || catLower.includes('maestr')) {
            catalogoPersonal['Docentes'].push(`${nombre} (${categoria})`);
          } else if (catLower.includes('administrativo') || catLower.includes('admin') || catLower.includes('direcci')) {
            catalogoPersonal['Admin'].push(`${nombre} (${categoria})`);
          } else {
            catalogoPersonal['Otros'].push(`${nombre} (${categoria})`);
          }
        }
      });

      // Ordenar listas alfabéticamente
      catalogoPersonal['Docentes'].sort((a, b) => a.localeCompare(b, 'es'));
      catalogoPersonal['Admin'].sort((a, b) => a.localeCompare(b, 'es'));
      catalogoPersonal['Otros'].sort((a, b) => a.localeCompare(b, 'es'));
    }

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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
