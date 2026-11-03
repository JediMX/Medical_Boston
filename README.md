# Boston Medical Record 🏥
### Sistema Web de Registro Clínico y Enfermería Escolar
**Colegio Boston - Cuernavaca, Morelos**

---

## 📌 Descripción General
Aplicación Web desarrollada sobre **Google Apps Script (GAS)** y **Google Sheets** para la gestión y captura en tiempo real de atenciones médicas en la enfermería escolar. Brinda control de acceso por roles, listas dinámicas dependientes de alumnos y personal, control de concurrencia y generación automática de folios consecutivos únicos (`ENF-XXXX`).

---

## 📂 Estructura de Archivos del Proyecto
```
Boston_Medical_Record/
│
├── Codigo.gs        # Backend (Google Apps Script): APIs, Locks, Login y DB Sheets
├── Index.html       # Frontend (HTML5, CSS3 institucional, JavaScript reactivo)
└── README.md        # Documentación de configuración y despliegue
```

---

## 📊 Estructura de la Hoja de Cálculo (Google Sheets)
El script requiere un archivo de Google Sheets activo vinculado al proyecto con las siguientes hojas y nombres exactos:

### 1. Hoja: `Usuarios`
| Columna A | Columna B | Columna C | Columna D |
| :--- | :--- | :--- | :--- |
| **Usuario** | **Password** | **Rol** | **Nombre Completo** |
| enfermeria | Boston2026* | Capturista | Lic. Ana Morales |
| direccion | Admin2026* | AdminJr | Lic. Hugo García |

### 2. Hoja: `Directorio_Alumnos`
| Columna A | Columna B | Columna C |
| :--- | :--- | :--- |
| **Nivel** | **Grupo** | **Nombre del Alumno** |
| Secundaria | 1°A | Álvarez Pérez Juan Carlos |
| Secundaria | 3°B | Domínguez Soto Mariana |
| Preparatoria | 1° Semestre | Morales Reyes Emiliano |

### 3. Hoja: `Directorio_Personal`
| Columna A | Columna B |
| :--- | :--- |
| **Categoría** | **Nombre Completo** |
| Docente Secundaria | Prof. Ricardo Méndez |
| Administrativo | Lic. Patricia Fuentes |
| Docente Preparatoria | Mtra. Sofía Valenzuela |

### 4. Hoja: `Registro_Enfermeria`
*(Se crea automáticamente si no existe la primera vez que se guarda un registro).*
- **Columnas:** Folio (`ENF-0001`), Fecha Ingreso, Hora Entrada, Capturista, Sección, Grado y Grupo, Nombre Paciente, Edad, FC, FR, SpO2 (%), T/A, Temp (°C), Síntoma Principal, Síntoma Detalle, Tratamiento Aplicado, Observaciones, Hora Salida, Estatus.

---

## 📱 Optimización para Dispositivos Móviles (Persistencia de Sesión)
En dispositivos móviles (Safari en iOS y Chrome en Android), los navegadores descargan de memoria RAM las pestañas en segundo plano cuando el usuario:
- Bloquea la pantalla o se apaga el teléfono.
- Abre otra aplicación (como WhatsApp, Cámara, Galería o Llamadas).
- Cambia entre pestañas.

### Soluciones Implementadas:
1. **Persistencia Local (`localStorage`)**: Al marcar la casilla *"Mantener mi sesión abierta en este celular"*, la sesión se almacena de forma segura en el almacenamiento local del dispositivo con vigencia renovable de 30 días. Al volver a abrir la página, el usuario entra directo a la captura sin pedir usuario ni contraseña.
2. **Auto-guardado de Borrador (`sessionStorage`)**: Si la enfermera está capturando signos vitales o notas clínicas y el celular se bloquea o recarga la pestaña, los campos capturados se restauran automáticamente sin perder información.
3. **Modo Web App / PWA (Pantalla de Inicio)**:
   - **En iPhone (Safari)**: Presionar el botón *Compartir* (cuadro con flecha arriba) y seleccionar **"Agregar a pantalla de inicio"**.
   - **En Android (Chrome)**: Tocar los 3 puntos superiores y elegir **"Agregar a la pantalla principal"** o **"Instalar aplicación"**.
   Esto abre la aplicación a pantalla completa como una App nativa del Colegio Boston, eliminando las recargas accidentales del navegador web.

---

## 🛠️ Mejoras y Correcciones Aplicadas en la Depuración
1. **Prevención de Folios Duplicados (`LockService`)**: Se implementó un bloqueo exclusivo con espera de hasta 20 segundos para evitar colisiones cuando varias enfermeras o capturistas guarden registros de forma simultánea.
2. **Eliminación de Excepciones Fatal por Nulos**: Las llamadas a `.toString()` sobre celdas vacías o filas sin contenido ahora están protegidas con validación defensiva.
3. **Manejo Asíncrono Robusto de Catálogos**: `Index.html` ya no falla si se selecciona una sección antes de que termine la descarga del directorio; cuenta con estado inicial defensivo y auto-recarga reactiva.
4. **Validación Completa en "Público en General"**: Se corrigió la fuga donde se podían guardar pacientes públicos en blanco.
5. **Manejo Centralizado de Errores (`withFailureHandler`)**: Todos los llamados a `google.script.run` tienen ahora manejadores de fallo que informan al usuario en pantalla en lugar de congelar el botón de login o guardado.
6. **Interfaz Institucional Boston**: Diseño adaptado con la paleta de colores azul marino y dorado, tipografía Inter, alertas visuales integradas y adaptación responsiva a pantallas móviles y tablets.
