# Recon v2 — Especificación de Diseño
**Fecha:** 2026-04-12  
**Proyecto:** Sistema de Reconocimiento de Créditos — UPM ETSII  
**Autor:** djfer

---

## 1. Resumen

Recon v2 es una refactorización completa de `Recon v1.html`. Mantiene toda la lógica existente, corrige todos los errores conocidos, añade una base de datos local persistente (IndexedDB), rediseña la interfaz al estilo Dark Dashboard, e incorpora nuevos módulos: gestión de convocatorias, seguimiento de informes de profesores, envío de correos a Thunderbird, estadísticas y administración de datos.

---

## 2. Arquitectura

### Archivos
- `Recon.html` — aplicación completa (interfaz + lógica)
- `recon-data.js` — datos de seed: 1.193 precedentes + 3 planes de estudio (05IQ, 05IR, 05TI) con datos de profesores

Ambos archivos deben estar en la misma carpeta. No se requiere servidor ni instalación.

### Base de datos — IndexedDB
Cuatro almacenes (`object stores`):

| Store | Contenido |
|---|---|
| `solicitudes` | Cada alumno procesado con todos sus datos, asignaturas, estado y historial |
| `precedentes` | Registros históricos de reconocimiento (inicialmente 1.193, ampliables) |
| `planes` | Datos de los 3 planes de estudio con info de profesores por asignatura |
| `config` | Versión de datos, fecha último backup, ajustes de la app |

### Backup automático
- Al cerrar el navegador (`beforeunload`): exporta `recon-backup-YYYY-MM-DD.json`
- Al abrir, si ya existen datos en IndexedDB: no pregunta nada, arranca normal
- Al abrir sin datos (primera vez o nuevo PC): ofrece importar un `.json` de backup
- Botón manual "Exportar BD" siempre disponible en la barra superior

### Librerías
Todas las librerías (pdf.js, xlsx.js, docx.js, pako) se embeben directamente en el HTML — sin dependencias de CDN, funciona sin internet salvo las llamadas a Claude API.

---

## 3. Interfaz — Dark Dashboard

**Paleta de colores:**
- Fondo: `#0f172a`
- Superficie: `#1e293b`
- Acento principal: `#38bdf8` (azul claro)
- Texto: `#cbd5e1`
- Verde: `#4ade80`, Rojo: `#f87171`, Ámbar: `#fbbf24`

**Estructura:**
- Barra superior fija con nombre, estado de BD, botón exportar
- Navegación lateral con las 7 secciones
- Área de contenido principal
- Alertas de intervención humana: borde rojo pulsante + icono ⚠️ prominente

**Principio de edición:** todo es editable. Cualquier campo de cualquier alumno puede modificarse manualmente en cualquier momento.

---

## 4. Secciones de la app

### 4.1 Nueva Solicitud
Flujo de procesamiento de una nueva solicitud de alumno.

**Pasos:**
1. Subir PDF de solicitud (obligatorio) + certificado académico (opcional)
2. Detectar si el PDF tiene texto extraíble (pdf.js):
   - Si sí → enviar solo texto a Claude (modo económico)
   - Si no → enviar imágenes de las páginas relevantes a Claude (modo visión)
3. Claude extrae datos personales y tabla de asignaturas
4. Cruce automático con planes de estudio
5. Cruce con precedentes (lógica estricta, ver sección 6)
6. Clasificación automática: Automático / Comisión / En duda
7. Guardar en IndexedDB
8. Mostrar resultados editables

**Detección de duplicados:**
Si el DNI/NIE ya existe en la BD → alerta prominente con opciones:
- "Reemplazar" (borra la anterior, guarda la nueva)
- "Fusionar" (combina asignaturas de ambas, el humano revisa)
- "Cancelar"

### 4.2 Solicitudes
Listado general de todos los alumnos procesados.

**Filtros:** estado, convocatoria, universidad de origen, titulación destino, fecha
**Vista:** tabla con nombre, DNI, universidad, nº asignaturas, estado, convocatoria asignada
**Acciones por alumno:** ver detalle, editar, asignar a convocatoria, eliminar
**Vista detalle:** todos los datos del alumno + tabla de asignaturas + historial de cambios

**Estados posibles de una solicitud:**

| Estado | Descripción |
|---|---|
| `nuevo` | Recién procesado, sin revisar |
| `automatico` | Todas las asignaturas con precedente válido |
| `comision` | Alguna asignatura necesita informe de profesor |
| `en_duda` ⚠️ | Similitud baja, requiere revisión humana |
| `correo_enviado` | Informe solicitado a profesores |
| `informe_recibido` | Al menos un informe recibido, pendiente el resto |
| `resuelto` | Todos los informes recibidos o resuelto manualmente |

**Estados posibles de una asignatura:**

| Estado | Descripción |
|---|---|
| `automatico` | Precedente válido confirmado |
| `pendiente` | Sin precedente, a la espera de informe profesor |
| `favorable` | Profesor aprueba reconocimiento |
| `desfavorable` | Profesor deniega reconocimiento |
| `en_duda` ⚠️ | Similitud insuficiente, humano debe decidir |
| `manual` | Resuelto manualmente por el humano |

### 4.3 Convocatorias
Gestión de sesiones de la comisión de reconocimiento.

**Operaciones:**
- Crear nueva convocatoria (nombre + fecha, ej: "Comisión junio 2026")
- Asignar alumnos a una convocatoria manualmente o por filtro de estado
- Ver alumnos asignados a cada convocatoria

**Documentos generables por convocatoria:**
- **Listado de comisión**: según plantilla del usuario (formato exacto preservado)
- **Acta de comisión**: según plantilla del usuario (formato exacto preservado)

**Estructura del Acta (formato analizado del modelo real):**
1. Cabecera: "SUBCOMISIÓN DE RECONOCIMIENTO DE ESTUDIOS — Acta nº XX"
2. Párrafo de apertura: fecha, hora, lugar
3. Lista de miembros con estado "Asiste / No asiste / Excusa asistencia" (configurable en Administración)
4. Orden del día (5 puntos fijos)
5. Puntos 1 y 2 con texto fijo configurable
6. Punto 3 — bloques por grado (05TI → 05IR → 05IQ), ordenados alfabéticamente:
   - **Alumnos de comisión**: nombre, titulación origen/destino, tabla de asignaturas reconocidas (código UPM | nombre UPM | asignatura origen), informes favorables/desfavorables por profesor
   - **Alumnos automáticos**: sección separada "PROPUESTAS DE RECONOCIMIENTO POR PRECEDENTES" — solo nombre + titulación origen + titulación destino

**Estructura de la Relación de Solicitantes (formato analizado del modelo real):**
1. Cabecera: fecha, nº sesión, resumen de solicitudes (completas / falta documentación)
2. Por grado (05TI → 05IR → 05IQ), por alumno, ordenado alfabéticamente:
   - Nombre, titulación origen/destino
   - Tabla ✓ Informe favorable: código UPM | nombre UPM | asignaturas origen
   - Tabla ✗ Informe desfavorable: código UPM | nombre UPM | asignaturas origen
   - Tabla ◌ Pendiente de informe: código UPM | nombre UPM | asignaturas origen
   - Desglose de pendientes por profesor
3. Sección final "INFORMES PENDIENTES DE EMISIÓN": tabla consolidada profesor → alumno → asignaturas pendientes

### 4.4 Correos a Profesores
Gestión centralizada del envío de solicitudes de informe a profesores.

**Vista principal:** tabla de profesores con:
- Nombre, departamento, email
- Nº de asignaturas pendientes de informe
- Nº de alumnos afectados
- Estado general: `sin_contactar / contactado / parcial / completo`
- Días transcurridos desde el primer envío (contador de tiempo)

**Flujo de envío:**
1. Seleccionar profesor
2. Elegir modo: **individual** (un correo por alumno) o **consolidado** (un correo con todos los alumnos)
3. Vista previa del correo generado (asunto + cuerpo prefabricado con datos del alumno/s)
4. Al confirmar:
   - Se descarga automáticamente el PDF de la carta
   - Se abre Thunderbird con `mailto:` prefabricado (to, subject, body)
   - El estado de la asignatura pasa a `correo_enviado`
   - Se registra la fecha de envío

**Documento "Informe de pendientes":**
Genera un documento con todos los profesores que no han respondido, cuántos días llevan sin responder y qué asignaturas tienen pendientes. Para enviar al jefe de estudios y reclamar respuestas.

**Configuración del correo (en Administración):**
- Plantilla de asunto
- Plantilla de cuerpo del correo
- Todo editable con marcadores de posición

### 4.5 Informes de Profesores
Módulo para procesar las resoluciones recibidas de los profesores.

**Subida:** drag & drop o selector. Acepta PDF, Word (.docx), imagen, o cualquier formato de texto.

**Procesamiento (Claude Haiku — coste mínimo):**
1. Extraer texto del documento (pdf.js / texto directo)
2. Enviar solo texto a Claude Haiku con prompt mínimo:
   - ¿Se reconoce o deniega? ¿Cuál es el motivo?
3. Claude devuelve: `{decision: "favorable"|"desfavorable"|"duda", motivo: "..."}`
4. Si `duda` → alerta de intervención humana ⚠️
5. Asociar resultado al alumno y asignatura correspondientes
6. Actualizar estado en BD

**Asociación del informe al alumno:**
- El sistema intenta detectar automáticamente el nombre del alumno / código de asignatura en el documento
- Si no puede → muestra selector manual para que el humano elija a qué alumno y asignatura corresponde

**Motivo:** se guarda en BD, no visible por defecto, expandible con un clic.

### 4.6 Estadísticas
Panel de métricas y gráficos.

**Métricas:**
- Total solicitudes por estado (automático / comisión / resuelto / pendiente)
- Tasa de reconocimiento automático vs manual
- Top 10 universidades de origen
- Distribución por titulación de destino (05IQ / 05IR / 05TI)
- Solicitudes por convocatoria
- Tiempo medio de resolución
- Profesores con mayor nº de informes pendientes
- Asignaturas más frecuentemente reconocidas / denegadas

**Implementación:** gráficos con Canvas API nativo (sin librería externa).

### 4.7 Administración
Gestión de datos del sistema.

**Actualizar precedentes:**
- Subir nuevos PDFs de reconocimiento histórico
- El sistema los procesa con pdf.js + parsePrecBlock()
- Los añade a la BD (sin borrar los existentes)
- Muestra contador: "X precedentes añadidos, Y ya existían"

**Actualizar planes de estudio:**
- Subir nuevo Excel (.xlsx) para cualquiera de los 3 planes
- El sistema lo procesa y reemplaza el plan correspondiente en BD
- Validación: muestra preview de los datos detectados antes de confirmar

**Configuración de correos:**
- Plantilla de asunto y cuerpo para solicitud de informe a profesor
- Marcadores disponibles: `{{nombre_profesor}}`, `{{tratamiento}}`, `{{alumno}}`, `{{asignatura}}`, `{{fecha}}`, etc.

**Plantillas de documentos:**
- Subir/editar plantilla de listado de comisión
- Subir/editar plantilla de acta de comisión
- Subir/editar plantilla de informe de pendientes

**Gestión de API:**
- Campo para la clave de Anthropic
- Selector de modelo para extracción (Sonnet por defecto)
- Selector de modelo para informes de profesores (Haiku por defecto)

---

## 5. Optimización de tokens

| Llamada | Modelo | Estrategia |
|---|---|---|
| Extracción datos alumno | claude-sonnet-4-20250514 | Texto si pdf.js extrae bien; imágenes solo si el texto es ilegible. Solo páginas necesarias (p.1 para datos personales, páginas con tabla de asignaturas). |
| Interpretación informe profesor | claude-haiku-4-5-20251001 | Solo texto extraído. Prompt mínimo (< 200 tokens). |
| Prompts | — | Sin texto de relleno. JSON schema en el prompt de extracción. |

---

## 6. Lógica de precedentes (refinada)

### Umbrales de similitud

| Comparación | Umbral confirmado | Umbral duda | Sin match |
|---|---|---|---|
| Universidad (palabras distintivas) | ≥ 0.6 | 0.4 – 0.59 | < 0.4 |
| Titulación de origen | ≥ 0.65 | 0.50 – 0.64 | < 0.50 |
| Asignatura de origen | ≥ 0.45 | — | < 0.45 |

**Zona de duda:** si universidad o titulación cae en rango intermedio → estado `en_duda`, alerta prominente, el humano confirma o descarta manualmente. El sistema NO aplica el precedente automáticamente en caso de duda.

**Principio:** mejor falso negativo (marcar como pendiente algo que sí tenía precedente) que falso positivo (aplicar un precedente incorrecto).

### Clasificación automática de la solicitud
- **Automático**: TODAS las asignaturas tienen precedente confirmado (ninguna en duda, ninguna sin precedente) → solicitud resuelta sin pasar por comisión
- **Comisión**: alguna asignatura sin precedente confirmado → necesita informe de profesor
- **En duda** ⚠️: alguna asignatura con similitud en rango intermedio → intervención humana requerida antes de continuar

---

## 7. Correcciones de bugs (v1 → v2)

| Bug | Línea v1 | Corrección |
|---|---|---|
| `HEADER_B64` vs `HEADER_IMG_B64` | 1143, 1161 | Unificar nombre de variable |
| `window.open()` sin null check | 1104 | Comprobar `if(!win)` antes de usar |
| `a.origen.filter()` sin validación | 1017, 1123, 1265 | Usar `(a.origen\|\|[]).filter()` |
| Umbral titulación 0.4 demasiado bajo | 773 | Subir a 0.65 con zona de duda |
| `anthropic-version` desactualizada | 641 | Actualizar a versión actual |
| `updatePreview()` múltiples llamadas | 980–986 | Debounce o llamada única |
| Errores Word silenciados | 1078–1081 | Mostrar error al usuario |
| Sin validación campos modal | 382–388 | Validar fecha y titulación destino |

---

## 8. Plantillas de documentos

Analizadas a partir de los modelos reales proporcionados por el usuario. Generadas programáticamente en JavaScript (sin dependencia de plantillas externas).

Los miembros de la subcomisión (Presidente, Secretaria, Vocales, Invitados) son configurables desde la sección Administración y se persisten en IndexedDB.

El nº de acta y el nº de sesión se gestionan como campo editable en cada convocatoria.

---

## 9. Fuera de alcance (decidido explícitamente)

- Notificación de resultado al alumno
- Log de auditoría de cambios
- Modo offline sin Claude API
- Acceso multiusuario / autenticación
- Migración a Access (se hará manualmente como proyecto separado)
