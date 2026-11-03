// ============================================================================
// BOSTON MEDICAL RECORD - SISTEMA DE REGISTRO CLÍNICO Y ENFERMERÍA
// Colegio Boston - Cuernavaca, Morelos
// ============================================================================

// ==========================================
// 1. SERVIR LA APLICACIÓN WEB
// ==========================================
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Boston Medical Record | Colegio Boston')
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
function verificarLogin(usuario, password) {
  try {
    if (!usuario || !password) {
      return { exito: false, mensaje: 'Por favor, ingresa usuario y contraseña.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaUsuarios = ss.getSheetByName('Usuarios');

    if (!hojaUsuarios) {
      return { exito: false, mensaje: 'Error de configuración: No existe la hoja "Usuarios".' };
    }

    const datos = hojaUsuarios.getDataRange().getValues();
    if (datos.length <= 1) {
      return { exito: false, mensaje: 'No hay usuarios registrados en el sistema.' };
    }

    const userLimpio = usuario.toString().trim().toLowerCase();
    const passLimpio = password.toString().trim();

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

    return { exito: false, mensaje: 'Usuario o contraseña incorrectos.' };
  } catch (error) {
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
    const hojaAlumnos = ss.getSheetByName('Directorio_Alumnos');
    let catalogoAlumnos = {};

    if (hojaAlumnos && hojaAlumnos.getLastRow() > 1) {
      const datosAlumnos = hojaAlumnos.getDataRange().getValues();
      datosAlumnos.shift(); // Quitar fila de encabezados

      datosAlumnos.forEach(fila => {
        const nivel = fila[0] ? fila[0].toString().trim() : '';
        const grupo = fila[1] ? fila[1].toString().trim() : '';
        const nombre = fila[2] ? fila[2].toString().trim() : '';

        if (nivel && grupo && nombre) {
          if (!catalogoAlumnos[nivel]) catalogoAlumnos[nivel] = {};
          if (!catalogoAlumnos[nivel][grupo]) catalogoAlumnos[nivel][grupo] = [];
          catalogoAlumnos[nivel][grupo].push(nombre);
        }
      });

      // Ordenar alfabéticamente los nombres dentro de cada grupo
      Object.keys(catalogoAlumnos).forEach(nivel => {
        Object.keys(catalogoAlumnos[nivel]).forEach(grupo => {
          catalogoAlumnos[nivel][grupo].sort((a, b) => a.localeCompare(b, 'es'));
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
    let hoja = ss.getSheetByName('Registro_Enfermeria');

    // Si la hoja no existe, se crea con sus encabezados oficiales
    if (!hoja) {
      hoja = ss.insertSheet('Registro_Enfermeria');
      hoja.appendRow([
        'Folio',
        'Fecha Ingreso',
        'Hora Entrada',
        'Capturista',
        'Sección',
        'Grado y Grupo',
        'Nombre Paciente',
        'Edad',
        'FC',
        'FR', // Corregido FV por FR (Frecuencia Respiratoria habitual en clínica)
        'SpO2 (%)',
        'T/A',
        'Temp (°C)',
        'Síntoma Principal',
        'Síntoma Detalle / Otro',
        'Tratamiento Aplicado',
        'Observaciones',
        'Hora Salida',
        'Estatus'
      ]);
    }

    // --- Generación Robusta de Folio Consecutivo ---
    const ultimaFila = hoja.getLastRow();
    let nuevoNumero = 1;

    if (ultimaFila > 1) {
      // Buscar el último folio registrado en la columna A
      const valoresFolios = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
      for (let i = valoresFolios.length - 1; i >= 0; i--) {
        const celda = valoresFolios[i][0] ? valoresFolios[i][0].toString() : '';
        const match = celda.match(/ENF-(\d+)/);
        if (match) {
          nuevoNumero = parseInt(match[1], 10) + 1;
          break;
        }
      }
    }

    const nuevoFolio = 'ENF-' + String(nuevoNumero).padStart(4, '0');

    // Sanitización y armado de fila
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
      datos.observaciones || '',
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
