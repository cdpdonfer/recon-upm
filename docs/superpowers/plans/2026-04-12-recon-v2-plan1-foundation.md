# Recon v2 — Plan 1: Fundación (DB + UI + Core Flow)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reescribir Recon.html con Dark Dashboard, IndexedDB, bugfixes completos, procesamiento de solicitudes mejorado con optimización de tokens y flujo de clasificación automático/comisión/en_duda.

**Architecture:** Dos archivos: `Recon.html` (app completa) + `recon-data.js` (datos de seed: base64 de PDFs de precedentes + Excels de planes). IndexedDB almacena solicitudes procesadas, config y caché de datos. El HTML contiene todo el CSS, JS y estructura. No hay servidor ni build step.

**Tech Stack:** HTML5, CSS3, Vanilla JS (ES2020+), IndexedDB, pdf.js (embedded), xlsx.js (embedded), docx.js (embedded), Claude API (Sonnet para extracción, Haiku para informes).

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `Recon.html` | App completa: CSS dark dashboard, HTML estructura, JS lógica |
| `recon-data.js` | Seed data: base64 PDFs de precedentes + base64 Excels de planes |

Dentro de `Recon.html`, el JS se organiza en secciones comentadas:
- `// ── DB ──` — IndexedDB CRUD
- `// ── STATE ──` — Estado global de la app (S)
- `// ── SEED ──` — Carga inicial de datos desde recon-data.js → IndexedDB
- `// ── PARSE ──` — pdf.js + parsePrecBlock + parseExcel
- `// ── CLAUDE ──` — callClaude + buildPrompt (token-optimized)
- `// ── ENRICH ──` — enrichExcel + enrichPrec + autoClassify
- `// ── RENDER ──` — Todas las funciones renderXxx
- `// ── LETTER ──` — Generación de cartas a ponentes
- `// ── BACKUP ──` — Export/Import JSON
- `// ── NAV ──` — Navegación entre secciones
- `// ── INIT ──` — Inicialización al cargar

---

## Task 1: Crear recon-data.js (extraer datos del HTML actual)

**Files:**
- Create: `recon-data.js`

El HTML actual tiene ~200KB de base64 embebido (PDFs de precedentes + Excels de planes). Lo movemos a recon-data.js para mantener el HTML manejable.

- [ ] **Step 1: Extraer el contenido base64 del HTML actual**

Abrir `Recon v1.html` en un editor de texto. Buscar las variables con base64. Serán del tipo:
```js
const PREC_PDF_IQ = "JVBERi0x..."; // PDF de precedentes 05IQ
const PREC_PDF_IR = "JVBERi0x..."; // PDF de precedentes 05IR  
const PREC_PDF_TI = "JVBERi0x..."; // PDF de precedentes 05TI
const EXCEL_IQ_B64 = "UEsDB...";   // Excel plan 05IQ
const EXCEL_IR_B64 = "UEsDB...";   // Excel plan 05IR
const EXCEL_TI_B64 = "UEsDB...";   // Excel plan 05TI
const HEADER_IMG_B64 = "/9j/...";  // Logo UPM cabecera cartas
```

- [ ] **Step 2: Crear recon-data.js con todas las variables extraídas**

```js
// recon-data.js — Datos de seed para Recon v2
// Este archivo contiene los datos embebidos (precedentes + planes de estudio)
// Se carga antes que Recon.html mediante <script src="recon-data.js">

// ── PRECEDENTES (PDFs base64) ────────────────────────────
// Cada variable contiene el PDF completo con los registros históricos
const PREC_PDF_IQ = "<base64 extraído del HTML original>";
const PREC_PDF_IR = "<base64 extraído del HTML original>";
const PREC_PDF_TI = "<base64 extraído del HTML original>";

// ── PLANES DE ESTUDIO (Excels base64) ───────────────────
const EXCEL_IQ_B64 = "<base64 extraído del HTML original>";
const EXCEL_IR_B64 = "<base64 extraído del HTML original>";
const EXCEL_TI_B64 = "<base64 extraído del HTML original>";

// ── LOGO CABECERA CARTAS ─────────────────────────────────
const HEADER_B64 = "<base64 extraído del HTML original>";
// NOTA: En v1 se llamaba HEADER_IMG_B64 — BUG CORREGIDO: unificado a HEADER_B64
```

- [ ] **Step 3: Verificar que recon-data.js carga sin errores**

Abrir navegador → consola → `open recon-data.js` o crear HTML de prueba mínimo:
```html
<!DOCTYPE html><html><head>
<script src="recon-data.js"></script>
<script>
  console.assert(typeof PREC_PDF_IQ === 'string' && PREC_PDF_IQ.length > 1000, 'PREC_PDF_IQ missing');
  console.assert(typeof PREC_PDF_IR === 'string' && PREC_PDF_IR.length > 1000, 'PREC_PDF_IR missing');
  console.assert(typeof PREC_PDF_TI === 'string' && PREC_PDF_TI.length > 1000, 'PREC_PDF_TI missing');
  console.assert(typeof EXCEL_IQ_B64 === 'string' && EXCEL_IQ_B64.length > 100, 'EXCEL_IQ_B64 missing');
  console.assert(typeof HEADER_B64 === 'string' && HEADER_B64.length > 100, 'HEADER_B64 missing');
  console.log('✅ recon-data.js OK');
</script>
</head><body>Ver consola</body></html>
```
Abrir en navegador. Consola debe mostrar `✅ recon-data.js OK` sin errores.

- [ ] **Step 4: Commit**
```bash
git init  # si no existe ya
git add "recon-data.js"
git commit -m "feat: extract seed data to recon-data.js (precedents + study plans)"
```

---

## Task 2: IndexedDB — Capa de base de datos

**Files:**
- Modify: `Recon.html` — sección `// ── DB ──`

Toda la persistencia de solicitudes, config y caché de datos pasa por estas funciones. El store `solicitudes` es el corazón del sistema.

### Schema de IndexedDB

```
DB name: "ReconUPM"  version: 1

stores:
  solicitudes     keyPath: "id" (autoincrement)
    indexes: dni (unique), estado, convocatoria_id, fecha_procesado
  
  config          keyPath: "key"
    (clave/valor: apikey, db_version, ultima_exportacion, miembros_comision, etc.)
  
  precedentes_cache  keyPath: "key"  
    (key: "05IQ"|"05IR"|"05TI", value: array de precedentes ya parseados)
  
  planes_cache    keyPath: "key"
    (key: "05IQ"|"05IR"|"05TI", value: array de filas del Excel)
```

- [ ] **Step 1: Escribir la función openDB()**

En la sección `// ── DB ──` de Recon.html:

```js
// ── DB ──────────────────────────────────────────────────
const DB_NAME = 'ReconUPM', DB_VER = 1;
let _db = null;

function openDB() {
  if (_db) return Promise.resolve(_db);
  return new Promise((res, rej) => {
    const req = indexedDB.open(DB_NAME, DB_VER);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      // solicitudes
      if (!db.objectStoreNames.contains('solicitudes')) {
        const s = db.createObjectStore('solicitudes', {keyPath:'id', autoIncrement:true});
        s.createIndex('dni', 'dni', {unique:true});
        s.createIndex('estado', 'estado', {unique:false});
        s.createIndex('convocatoria_id', 'convocatoria_id', {unique:false});
      }
      // config
      if (!db.objectStoreNames.contains('config'))
        db.createObjectStore('config', {keyPath:'key'});
      // caché precedentes parseados
      if (!db.objectStoreNames.contains('precedentes_cache'))
        db.createObjectStore('precedentes_cache', {keyPath:'key'});
      // caché planes de estudio
      if (!db.objectStoreNames.contains('planes_cache'))
        db.createObjectStore('planes_cache', {keyPath:'key'});
    };
    req.onsuccess = e => { _db = e.target.result; res(_db); };
    req.onerror = e => rej(e.target.error);
  });
}
```

- [ ] **Step 2: Escribir helpers genéricos de DB**

```js
function dbGet(store, key) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).get(key);
    req.onsuccess = () => res(req.result);
    req.onerror = e => rej(e.target.error);
  }));
}

function dbPut(store, value) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).put(value);
    req.onsuccess = () => res(req.result);
    req.onerror = e => rej(e.target.error);
  }));
}

function dbAdd(store, value) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).add(value);
    req.onsuccess = () => res(req.result);
    req.onerror = e => rej(e.target.error);
  }));
}

function dbDelete(store, key) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).delete(key);
    req.onsuccess = () => res();
    req.onerror = e => rej(e.target.error);
  }));
}

function dbGetAll(store) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).getAll();
    req.onsuccess = () => res(req.result);
    req.onerror = e => rej(e.target.error);
  }));
}

function dbGetByIndex(store, indexName, value) {
  return openDB().then(db => new Promise((res, rej) => {
    const tx = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).index(indexName).get(value);
    req.onsuccess = () => res(req.result);
    req.onerror = e => rej(e.target.error);
  }));
}

// Config helpers
const getConfig = key => dbGet('config', key).then(r => r?.value);
const setConfig = (key, value) => dbPut('config', {key, value});
```

- [ ] **Step 3: Funciones específicas de solicitudes**

```js
// Guardar solicitud nueva (detecta duplicado por DNI)
async function saveSolicitud(solicitud) {
  const dni = solicitud.personalData?.pasaporte_nie_dni?.trim();
  if (dni) {
    const existing = await dbGetByIndex('solicitudes', 'dni', dni);
    if (existing) return { duplicate: true, existing };
  }
  const id = await dbAdd('solicitudes', {
    ...solicitud,
    dni: dni || '',
    estado: solicitud.estado || 'nuevo',
    convocatoria_id: null,
    fecha_procesado: new Date().toISOString(),
    fecha_modificado: new Date().toISOString()
  });
  return { duplicate: false, id };
}

async function updateSolicitud(id, changes) {
  const current = await dbGet('solicitudes', id);
  if (!current) throw new Error('Solicitud no encontrada: ' + id);
  await dbPut('solicitudes', {
    ...current, ...changes,
    id, // preserve key
    fecha_modificado: new Date().toISOString()
  });
}

async function deleteSolicitud(id) {
  await dbDelete('solicitudes', id);
}

async function getAllSolicitudes() {
  return dbGetAll('solicitudes');
}
```

- [ ] **Step 4: Test de la capa DB en consola del navegador**

Abrir Recon.html en navegador, abrir consola (F12) y ejecutar:
```js
// Test openDB
openDB().then(db => {
  console.assert(db instanceof IDBDatabase, 'DB no es IDBDatabase');
  console.assert(db.objectStoreNames.contains('solicitudes'), 'Store solicitudes missing');
  console.assert(db.objectStoreNames.contains('config'), 'Store config missing');
  console.log('✅ DB abre OK');
});

// Test dbPut/dbGet en config
setConfig('test_key', 'test_value')
  .then(() => getConfig('test_key'))
  .then(v => {
    console.assert(v === 'test_value', 'Config get/put fail');
    console.log('✅ Config get/put OK');
  });

// Test saveSolicitud + duplicate detection
const testSol = { personalData: { pasaporte_nie_dni: 'TEST123', nombre: 'Test' }, asignaturas: [] };
saveSolicitud({...testSol})
  .then(r => { console.assert(!r.duplicate, 'Primera inserción no debe ser duplicado'); return saveSolicitud({...testSol}); })
  .then(r => { console.assert(r.duplicate, 'Segunda inserción debe ser duplicado'); console.log('✅ Duplicate detection OK'); })
  .then(() => deleteSolicitud(await dbGetByIndex('solicitudes','dni','TEST123').then(r=>r.id)));
```

- [ ] **Step 5: Commit**
```bash
git add "Recon.html"
git commit -m "feat: add IndexedDB layer with solicitudes, config, cache stores"
```

---

## Task 3: Dark Dashboard — CSS y estructura HTML

**Files:**
- Modify: `Recon.html` — sección `<style>` y estructura `<body>`

Reemplaza el CSS existente con el tema Dark Dashboard. La estructura de navegación cambia de 4 pasos a 7 secciones.

- [ ] **Step 1: Reemplazar el bloque `<style>` completo**

Sustituir el `<style>...</style>` existente (líneas 11-168 aprox.) por:

```css
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
:root {
  --bg:#0f172a; --surface:#1e293b; --s2:#263348; --border:#334155; --borderS:#475569;
  --tx:#e2e8f0; --tx2:#94a3b8; --tx3:#64748b;
  --blue:#38bdf8; --blueDark:#0284c7; --blueBg:rgba(56,189,248,.1);
  --green:#4ade80; --greenBg:rgba(74,222,128,.1); --greenB:rgba(74,222,128,.3);
  --amber:#fbbf24; --amberBg:rgba(251,191,36,.1); --amberB:rgba(251,191,36,.3);
  --red:#f87171; --redBg:rgba(248,113,113,.1); --redB:rgba(248,113,113,.3);
  --purple:#c084fc; --purpleBg:rgba(192,132,252,.1); --purpleB:rgba(192,132,252,.3);
  --orange:#fb923c;
  --r:8px; --rL:12px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx);min-height:100vh;font-size:14px;line-height:1.5}

/* ── TOPBAR ── */
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px;display:flex;align-items:center;height:48px;gap:12px;position:sticky;top:0;z-index:50}
.topbar-logo{font-size:15px;font-weight:700;color:var(--blue);letter-spacing:-.3px}
.topbar-logo span{color:var(--tx2);font-weight:300}
.topbar-right{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:12px;color:var(--tx3)}
.topbar-db-status{display:flex;align-items:center;gap:5px;padding:3px 10px;border-radius:20px;background:var(--s2);font-size:11px}
.topbar-db-status.ok{color:var(--green)}
.topbar-db-status.warn{color:var(--amber)}

/* ── LAYOUT ── */
.layout{display:flex;min-height:calc(100vh - 48px)}
.sidebar{width:220px;flex-shrink:0;background:var(--surface);border-right:1px solid var(--border);padding:14px 10px;overflow-y:auto;position:sticky;top:48px;height:calc(100vh - 48px)}
.main{flex:1;padding:20px 24px;overflow-y:auto;min-width:0}

/* ── NAV ── */
.nav-section{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--tx3);padding:0 8px;margin:14px 0 4px}
.nav-item{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:var(--r);cursor:pointer;font-size:13px;color:var(--tx2);transition:all .12s;margin-bottom:1px;border:none;background:none;width:100%;text-align:left;font-family:'Inter',sans-serif}
.nav-item:hover{background:var(--s2);color:var(--tx)}
.nav-item.active{background:var(--blueBg);color:var(--blue);font-weight:500}
.nav-item svg{width:15px;height:15px;stroke:currentColor;fill:none;stroke-width:1.8;flex-shrink:0}
.nav-badge{margin-left:auto;background:var(--red);color:white;border-radius:10px;padding:1px 6px;font-size:10px;font-weight:700}

/* ── ALERT BANNER (intervención humana requerida) ── */
.alert-banner{background:var(--redBg);border:1px solid var(--redB);border-radius:var(--r);padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px;animation:pulse-border 2s infinite}
@keyframes pulse-border{0%,100%{border-color:var(--redB)}50%{border-color:var(--red)}}
.alert-banner svg{width:18px;height:18px;stroke:var(--red);fill:none;stroke-width:2;flex-shrink:0}
.alert-banner-text{flex:1;font-size:13px;color:var(--red);font-weight:500}

/* ── CARDS ── */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--rL);padding:16px 18px;margin-bottom:14px}
.card-header{display:flex;align-items:center;gap:8px;margin-bottom:14px}
.card-header h2{font-size:14px;font-weight:600;color:var(--tx)}
.card-header svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:2}

/* ── BADGES ── */
.badge{font-size:11px;padding:2px 8px;border-radius:20px;font-weight:500}
.b-blue{background:var(--blueBg);color:var(--blue)}
.b-green{background:var(--greenBg);color:var(--green)}
.b-amber{background:var(--amberBg);color:var(--amber)}
.b-red{background:var(--redBg);color:var(--red)}
.b-purple{background:var(--purpleBg);color:var(--purple)}
.b-orange{background:rgba(251,146,60,.1);color:var(--orange)}

/* ── ESTADO PILLS ── */
.estado-pill{display:inline-flex;align-items:center;gap:4px;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600}
.ep-auto{background:var(--greenBg);color:var(--green)}
.ep-comision{background:var(--blueBg);color:var(--blue)}
.ep-duda{background:var(--redBg);color:var(--red);animation:pulse-border 2s infinite}
.ep-nuevo{background:var(--s2);color:var(--tx2)}
.ep-enviado{background:var(--amberBg);color:var(--amber)}
.ep-resuelto{background:var(--greenBg);color:var(--green)}
.ep-pendiente{background:var(--s2);color:var(--tx3)}
.ep-favorable{background:var(--greenBg);color:var(--green)}
.ep-desfavorable{background:var(--redBg);color:var(--red)}

/* ── DROPZONE ── */
.dz{border:1.5px dashed var(--border);border-radius:var(--r);padding:20px 14px;text-align:center;cursor:pointer;background:var(--bg);transition:all .15s}
.dz:hover,.dz.over{border-color:var(--blue);background:var(--blueBg)}
.dz svg{width:26px;height:26px;stroke:var(--tx3);fill:none;stroke-width:1.5;margin-bottom:7px}
.dz p{font-size:12px;color:var(--tx2)}.dz strong{color:var(--blue)}
.dz input{display:none}
.file-ok{display:flex;align-items:center;gap:7px;padding:8px 12px;border-radius:var(--r);background:var(--greenBg);border:1px solid var(--greenB);font-size:12px;color:var(--green);font-weight:500;margin-top:8px}
.file-ok svg{width:14px;height:14px;stroke:currentColor;fill:none;stroke-width:2.5}

/* ── BUTTONS ── */
.btn{display:inline-flex;align-items:center;gap:7px;padding:8px 16px;border-radius:var(--r);font-family:'Inter',sans-serif;font-size:13px;font-weight:500;cursor:pointer;border:1px solid transparent;transition:all .12s}
.btn svg{width:14px;height:14px;stroke:currentColor;fill:none;stroke-width:2}
.btn-primary{background:var(--blue);color:#0f172a;border-color:var(--blue)}.btn-primary:hover{background:#7dd3fc}
.btn-primary:disabled{background:var(--border);border-color:var(--border);color:var(--tx3);cursor:not-allowed}
.btn-secondary{background:var(--s2);color:var(--tx);border-color:var(--border)}.btn-secondary:hover{background:var(--borderS)}
.btn-danger{background:var(--redBg);color:var(--red);border-color:var(--redB)}.btn-danger:hover{background:var(--redB)}
.btn-success{background:var(--greenBg);color:var(--green);border-color:var(--greenB)}.btn-success:hover{background:var(--greenB)}
.btn-sm{padding:4px 10px;font-size:11px}

/* ── FORMS ── */
.form-label{font-size:12px;color:var(--tx2);display:block;margin-bottom:4px}
.form-input{width:100%;padding:8px 10px;border:1px solid var(--border);border-radius:var(--r);font-size:13px;font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx)}
.form-input:focus{outline:none;border-color:var(--blue)}
.form-input::placeholder{color:var(--tx3)}
select.form-input{cursor:pointer}

/* ── TABLES ── */
.table-wrap{overflow-x:auto;border-radius:var(--r);border:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{background:var(--s2);padding:8px 10px;text-align:left;font-weight:600;font-size:10px;color:var(--tx3);text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid var(--border);white-space:nowrap}
tbody td{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:middle;line-height:1.4}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:var(--s2)}
.code{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--tx2)}

/* ── PROCESSING BOX ── */
.proc-box{background:var(--blueBg);border:1px solid rgba(56,189,248,.3);border-radius:var(--r);padding:16px 18px;text-align:center;display:none;margin-top:12px}
.proc-box.on{display:block}
.spin{width:24px;height:24px;margin:0 auto 9px;border:2.5px solid var(--border);border-top-color:var(--blue);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.proc-title{font-size:13px;font-weight:500;color:var(--blue)}
.proc-detail{font-size:11px;color:var(--tx2);margin-top:2px}

/* ── ERROR BOX ── */
.err-box{background:var(--redBg);border:1px solid var(--redB);border-radius:var(--r);padding:11px 14px;font-size:13px;color:var(--red);display:none;margin-top:10px}
.err-box.on{display:block}

/* ── TABS ── */
.tabbar{display:flex;gap:2px;border-bottom:1.5px solid var(--border);margin-bottom:18px}
.tab{padding:8px 14px;font-size:12px;font-weight:400;color:var(--tx2);cursor:pointer;border:none;background:none;border-bottom:2px solid transparent;margin-bottom:-1.5px;transition:all .12s;font-family:'Inter',sans-serif}
.tab:hover{color:var(--tx)}
.tab.active{color:var(--blue);border-bottom-color:var(--blue);font-weight:600}
.tab-content{display:none}.tab-content.active{display:block}

/* ── STAT CARDS ── */
.stats-row{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px}
.stat-card{background:var(--s2);border-radius:var(--r);padding:10px 16px;flex:1;min-width:100px;border:1px solid var(--border)}
.stat-card dt{font-size:10px;color:var(--tx3);margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}
.stat-card dd{font-size:22px;font-weight:700}
.stat-card.green dd{color:var(--green)}
.stat-card.blue dd{color:var(--blue)}
.stat-card.amber dd{color:var(--amber)}
.stat-card.red dd{color:var(--red)}
.stat-card.purple dd{color:var(--purple)}

/* ── MODAL ── */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:100;display:none;align-items:center;justify-content:center}
.modal-overlay.on{display:flex}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:var(--rL);padding:24px 26px;width:700px;max-width:95vw;max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.5)}
.modal-head{display:flex;align-items:center;gap:10px;margin-bottom:18px}
.modal-head h2{font-size:15px;font-weight:600;flex:1}
.modal-close{background:none;border:none;cursor:pointer;color:var(--tx3);font-size:20px;line-height:1;padding:2px}
.modal-close:hover{color:var(--tx)}

/* ── LETTER PREVIEW ── */
.letter-preview{background:white;color:#000;border-radius:var(--r);padding:24px 28px;font-family:Arial,sans-serif;font-size:12.5px;line-height:1.6;margin-bottom:14px;min-height:200px}
.lp-table{width:100%;border-collapse:collapse;margin:10px 0;font-size:11px}
.lp-table th{background:#1a4a7a;color:white;padding:5px 8px;text-align:center;font-weight:600}
.lp-table td{border:1px solid #ccc;padding:5px 8px;vertical-align:top}

/* ── MISC ── */
.divider{border:none;border-top:1px solid var(--border);margin:14px 0}
pre.raw-text{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--tx2);white-space:pre-wrap;max-height:500px;overflow-y:auto;background:var(--bg);padding:12px;border-radius:var(--r);border:1px solid var(--border);line-height:1.6}
@media(max-width:900px){.layout{flex-direction:column}.sidebar{width:100%;height:auto;position:relative;top:0}}
</style>
```

- [ ] **Step 2: Reemplazar el `<body>` con la nueva estructura de 7 secciones**

```html
<body>
<!-- ── TOPBAR ── -->
<div class="topbar">
  <div class="topbar-logo">Recon<span>UPM</span></div>
  <div class="topbar-right">
    <div class="topbar-db-status ok" id="dbStatus">
      <svg width="10" height="10" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4" fill="currentColor"/></svg>
      BD activa
    </div>
    <button class="btn btn-secondary" style="font-size:11px;padding:4px 10px" onclick="exportBackup()">↓ Exportar BD</button>
  </div>
</div>

<div class="layout">
<!-- ── SIDEBAR ── -->
<div class="sidebar">
  <div class="nav-section">Flujo de trabajo</div>
  <button class="nav-item active" id="nav-nueva" onclick="showSection('nueva')">
    <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
    Nueva Solicitud
  </button>
  <button class="nav-item" id="nav-solicitudes" onclick="showSection('solicitudes')">
    <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    Solicitudes
    <span class="nav-badge" id="badge-duda" style="display:none">!</span>
  </button>
  <button class="nav-item" id="nav-convocatorias" onclick="showSection('convocatorias')">
    <svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
    Convocatorias
  </button>
  <div class="nav-section">Profesores</div>
  <button class="nav-item" id="nav-correos" onclick="showSection('correos')">
    <svg viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
    Correos
    <span class="nav-badge" id="badge-correos" style="display:none">0</span>
  </button>
  <button class="nav-item" id="nav-informes" onclick="showSection('informes')">
    <svg viewBox="0 0 24 24"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
    Informes
  </button>
  <div class="nav-section">Análisis</div>
  <button class="nav-item" id="nav-estadisticas" onclick="showSection('estadisticas')">
    <svg viewBox="0 0 24 24"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
    Estadísticas
  </button>
  <button class="nav-item" id="nav-admin" onclick="showSection('admin')">
    <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
    Administración
  </button>
</div>

<!-- ── MAIN CONTENT ── -->
<div class="main">
  <div id="sec-nueva" class="section-content active"></div>
  <div id="sec-solicitudes" class="section-content" style="display:none"></div>
  <div id="sec-convocatorias" class="section-content" style="display:none"></div>
  <div id="sec-correos" class="section-content" style="display:none"></div>
  <div id="sec-informes" class="section-content" style="display:none"></div>
  <div id="sec-estadisticas" class="section-content" style="display:none"></div>
  <div id="sec-admin" class="section-content" style="display:none"></div>
</div>
</div>

<!-- MODALS (se insertan aquí) -->
<div id="modal-duplicate" class="modal-overlay"></div>
<div id="modal-letter" class="modal-overlay"></div>
</body>
```

- [ ] **Step 3: Añadir función showSection() en JS**

```js
// ── NAV ────────────────────────────────────────────────
function showSection(name) {
  document.querySelectorAll('.section-content').forEach(el => el.style.display = 'none');
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const sec = document.getElementById('sec-' + name);
  const nav = document.getElementById('nav-' + name);
  if (sec) sec.style.display = 'block';
  if (nav) nav.classList.add('active');
  // Render section on demand
  const renders = { nueva: renderNueva, solicitudes: renderSolicitudes,
    convocatorias: renderConvocatorias, correos: renderCorreos,
    informes: renderInformes, estadisticas: renderEstadisticas, admin: renderAdmin };
  if (renders[name]) renders[name]();
}
```

- [ ] **Step 4: Verificar visualmente**

Abrir Recon.html en el navegador. Debe verse:
- Fondo oscuro (#0f172a)
- Barra superior con "ReconUPM"
- Sidebar izquierdo con 7 secciones de navegación
- Área principal vacía (las secciones se rellenan en tasks siguientes)
- Sin errores en consola

- [ ] **Step 5: Commit**
```bash
git add "Recon.html"
git commit -m "feat: dark dashboard UI shell with 7-section navigation"
```

---

## Task 4: Bugfixes del código v1 + funciones de utilidad

**Files:**
- Modify: `Recon.html` — secciones `// ── STATE ──`, funciones utilitarias

- [ ] **Step 1: Definir el estado global S con la estructura v2**

```js
// ── STATE ────────────────────────────────────────────
const S = {
  // Solicitud en curso (Nueva Solicitud)
  pdfFile: null,
  pdfText: '',
  pdfPages: [],       // base64 de páginas como imágenes
  certPages: [],
  personalData: {},
  asignaturas: [],
  clasificacion: null, // 'automatico'|'comision'|'en_duda'
  currentSolicitudId: null,

  // Datos cargados en memoria
  precedentes: [],    // array de precedentes parseados [{numero, curso, _grado, titulaciones, reconocidas}]
  excelData: [null, null, null], // índice 0=05IQ, 1=05IR, 2=05TI

  // Procesados
  seedLoaded: false,
};
```

- [ ] **Step 2: Funciones de utilidad y helpers (corregidos)**

```js
// Normalización de texto para comparación
function norm(s) {
  s = String(s || '').toUpperCase();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  s = s.replace(/[^A-Z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  return s;
}

// Stop words para nombre de universidad
const UNIV_STOP = new Set(['UNIVERSIDAD','UNIVERSITAT','UNIVERSITY','UNIVERSITARIA',
  'POLITECNICA','POLITECNICO','NACIONAL','AUTONOMA','AUTONOMO','PUBLICA','PRIVADA',
  'TECNICA','TECNICO','DEL','LAS','LOS','THE','OF','DE','EN','LA','EL','Y','A','E',
  'GRADO','INGENIERIA','INDUSTRIAL','INGENIERO','CIENCIAS','SUPERIOR','ESCUELA',
  'INSTITUTO','FACULTAD','CENTRO','COLLEGE','ESTADO','ESTADOS','SAN','SUR','NORTE']);

function univWords(s) {
  return new Set(norm(s).split(' ').filter(w => w.length > 2 && !UNIV_STOP.has(w)));
}

function univSimilarity(a, b) {
  const wa = univWords(a), wb = univWords(b);
  if (!wa.size || !wb.size) return norm(a) === norm(b) ? 1 : 0;
  let common = 0;
  wa.forEach(w => { if (wb.has(w)) common++; });
  return common / Math.max(wa.size, wb.size);
}

function similarity(a, b) {
  const wa = new Set(norm(a).split(' ').filter(w => w.length > 2));
  const wb = new Set(norm(b).split(' ').filter(w => w.length > 2));
  if (!wa.size || !wb.size) return 0;
  let common = 0;
  wa.forEach(w => { if (wb.has(w)) common++; });
  return common / Math.max(wa.size, wb.size);
}

// Mostrar/ocultar error
function showErr(msg) {
  const el = document.getElementById('errbox');
  if (!el) return;
  el.textContent = msg;
  el.classList.add('on');
}
function hideErr() {
  const el = document.getElementById('errbox');
  if (el) el.classList.remove('on');
}

// Formatear fecha ISO → DD/MM/YYYY
function formatDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}
```

- [ ] **Step 3: Test de similitud en consola**

```js
// En consola del navegador:
console.assert(univSimilarity('Universidad del Bosque','Universidad El Bosque') > 0.6, 'Bosque similarity fail');
console.assert(univSimilarity('Universidad Complutense de Madrid','Universidad El Bosque') < 0.4, 'False positive fail');
console.assert(similarity('Ingeniería Civil','Ingeniería Civil y Territorial') < 0.65, 'Civil threshold fail — debe ser < 0.65');
console.assert(similarity('Cálculo I','Cálculo I') === 1, 'Exact match fail');
console.log('✅ Similarity functions OK');
```
Esperado: 3 assertions pasan, ningún error.

- [ ] **Step 4: Commit**
```bash
git add "Recon.html"
git commit -m "fix: state structure v2, corrected similarity thresholds, utility functions"
```

---

## Task 5: Carga de datos de seed (recon-data.js → IndexedDB)

**Files:**
- Modify: `Recon.html` — secciones `// ── SEED ──`, `// ── PARSE ──`

Al primer arranque, parsea los PDFs y Excels de recon-data.js y los guarda en IndexedDB. En arranques sucesivos, carga directo de IndexedDB (mucho más rápido).

- [ ] **Step 1: Función parsePrecBlock() corregida**

```js
// ── PARSE ───────────────────────────────────────────
function parsePrecBlock(text) {
  const nm = text.match(/Reconocimiento\s+N[uú]mero:?\s*(\d+)/i);
  if (!nm) return null;
  const cm = text.match(/Curso\s+Acad[eé]mico:?\s*([\d]{4}-[\d]{2,4})/i);
  const um = text.match(/ASIGNATURAS ORIGEN\s+\(([^)]+)\)/i);
  let universidad = '?', titulacion = '?';
  if (um) {
    const parts = um[1].split(/\s*-\s*/);
    universidad = parts[0]?.trim() || '?';
    titulacion = parts[parts.length - 1]?.trim() || '?';
  }
  const rS = text.match(/ASIGNATURAS OBLIGATORIAS RECONOCIDAS\s+(.*?)(?=RECONOCIMIENTO DE CR[EÉ]DITOS|$)/i);
  const rRaw = rS ? rS[1] : '';
  const reconocidas = [];
  const codeRe = /(\d{7,8})\s+([A-ZÁÉÍÓÚÜÑA-Z][A-ZÁÉÍÓÚÜÑA-Za-z\s]+?)(?=\s*\([^)]*\)\s*\d{7,8}|\s*\([^)]*\)\s*$|\s+\d{7,8}|$)/g;
  let m;
  while ((m = codeRe.exec(rRaw)) !== null) {
    const nombre = m[2].trim().replace(/\s+/g, ' ');
    if (nombre.length > 1) reconocidas.push({codigo: m[1].trim(), nombre});
  }
  if (!reconocidas.length) {
    const parts = rRaw.split(/(?=\d{7,8}\s+[A-ZÁÉÍÓÚ])/);
    parts.forEach(part => {
      const lm = part.match(/^(\d{7,8})\s+([A-ZÁÉÍÓÚÜÑA-Za-z][^(]+)/);
      if (lm) reconocidas.push({codigo: lm[1].trim(), nombre: lm[2].trim().replace(/\s+/g, ' ')});
    });
  }
  return { numero: nm[1], curso: cm ? cm[1] : '?', universidad, titulacion, reconocidas };
}

async function parsePrecPDF(b64, grado) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const pdf = await pdfjsLib.getDocument({data: bytes}).promise;
  const result = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const p = await pdf.getPage(i);
    const tc = await p.getTextContent();
    const txt = tc.items.map(t => t.str).join(' ');
    const rec = parsePrecBlock(txt);
    if (rec) result.push(rec);
  }
  return result;
}
```

- [ ] **Step 2: Función de carga de seed (primer arranque)**

```js
// ── SEED ─────────────────────────────────────────────
async function loadSeedIfNeeded() {
  // Comprobar si ya hay datos en caché
  const cachedIQ = await dbGet('precedentes_cache', '05IQ');
  const cachedPlans = await dbGet('planes_cache', '05IQ');

  if (cachedIQ && cachedPlans) {
    // Cargar desde caché — arranque rápido
    const [iq, ir, ti] = await Promise.all([
      dbGet('precedentes_cache', '05IQ'),
      dbGet('precedentes_cache', '05IR'),
      dbGet('precedentes_cache', '05TI'),
    ]);
    S.precedentes = [
      ...(iq?.data || []).map(r => ({...r, _grado:'05IQ'})),
      ...(ir?.data || []).map(r => ({...r, _grado:'05IR'})),
      ...(ti?.data || []).map(r => ({...r, _grado:'05TI'})),
    ];
    const [pIQ, pIR, pTI] = await Promise.all([
      dbGet('planes_cache', '05IQ'),
      dbGet('planes_cache', '05IR'),
      dbGet('planes_cache', '05TI'),
    ]);
    S.excelData = [pIQ?.data || null, pIR?.data || null, pTI?.data || null];
    S.seedLoaded = true;
    updateSeedUI();
    return;
  }

  // Primera vez: parsear desde recon-data.js
  setSeedStatus('Cargando precedentes históricos…');
  try {
    const defs = [
      { b64: PREC_PDF_IQ, grado: '05IQ', idx: 0 },
      { b64: PREC_PDF_IR, grado: '05IR', idx: 1 },
      { b64: PREC_PDF_TI, grado: '05TI', idx: 2 },
    ];
    for (const def of defs) {
      setSeedStatus(`Parseando precedentes ${def.grado}…`);
      const recs = await parsePrecPDF(def.b64, def.grado);
      await dbPut('precedentes_cache', { key: def.grado, data: recs });
      S.precedentes.push(...recs.map(r => ({...r, _grado: def.grado})));
    }

    setSeedStatus('Cargando planes de estudio…');
    const planDefs = [
      { b64: EXCEL_IQ_B64, grado: '05IQ', idx: 0 },
      { b64: EXCEL_IR_B64, grado: '05IR', idx: 1 },
      { b64: EXCEL_TI_B64, grado: '05TI', idx: 2 },
    ];
    for (const def of planDefs) {
      const rows = parseExcelB64(def.b64);
      await dbPut('planes_cache', { key: def.grado, data: rows });
      S.excelData[def.idx] = rows;
    }

    S.seedLoaded = true;
    setSeedStatus('');
    updateSeedUI();
  } catch(err) {
    setSeedStatus('Error cargando datos: ' + err.message, true);
  }
}

function parseExcelB64(b64) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const wb = XLSX.read(bytes, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {defval:''});
}

function setSeedStatus(msg, isError = false) {
  // Actualiza el indicador de carga en la UI de Nueva Solicitud
  const el = document.getElementById('seed-status');
  if (!el) return;
  el.textContent = msg;
  el.style.color = isError ? 'var(--red)' : 'var(--tx3)';
}

function updateSeedUI() {
  const iqEl = document.getElementById('plan-05IQ');
  const irEl = document.getElementById('plan-05IR');
  const tiEl = document.getElementById('plan-05TI');
  const counts = [0,1,2].map(i => {
    const grado = ['05IQ','05IR','05TI'][i];
    return S.precedentes.filter(p => p._grado === grado).length;
  });
  if (iqEl) iqEl.textContent = `${S.excelData[0]?.length || 0} asignaturas · ${counts[0]} precedentes`;
  if (irEl) irEl.textContent = `${S.excelData[1]?.length || 0} asignaturas · ${counts[1]} precedentes`;
  if (tiEl) tiEl.textContent = `${S.excelData[2]?.length || 0} asignaturas · ${counts[2]} precedentes`;
}
```

- [ ] **Step 3: Test de carga de seed**

Abrir navegador en Recon.html, abrir consola:
```js
// Esperar a que cargue y verificar:
setTimeout(() => {
  console.assert(S.seedLoaded === true, 'Seed no se cargó');
  console.assert(S.precedentes.length > 1000, `Solo ${S.precedentes.length} precedentes, esperaba > 1000`);
  console.assert(S.excelData[0] !== null, 'Plan 05IQ no cargado');
  console.assert(S.excelData[1] !== null, 'Plan 05IR no cargado');
  console.assert(S.excelData[2] !== null, 'Plan 05TI no cargado');
  console.log(`✅ Seed OK — ${S.precedentes.length} precedentes, planes cargados`);
}, 15000); // seed tarda en parsear los PDFs la primera vez
```

Esperado: sin assertions fallidas, ~1.193 precedentes cargados.

En recargas posteriores: carga en < 2 segundos desde IndexedDB.

- [ ] **Step 4: Commit**
```bash
git add "Recon.html"
git commit -m "feat: seed loader — parse precedents+plans on first load, cache in IndexedDB"
```

---

## Task 6: Claude API — Extracción con optimización de tokens

**Files:**
- Modify: `Recon.html` — sección `// ── CLAUDE ──`

- [ ] **Step 1: Función callClaude() genérica**

```js
// ── CLAUDE ───────────────────────────────────────────
async function callClaude(content, {model='claude-sonnet-4-20250514', maxTokens=8000} = {}) {
  const apiKey = await getConfig('apikey') || document.getElementById('apikey-input')?.value?.trim();
  if (!apiKey) throw new Error('No hay clave API configurada.');
  const messages = [{role:'user', content: Array.isArray(content) ? content : [{type:'text', text:content}]}];
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({model, max_tokens: maxTokens, messages})
  });
  if (!r.ok) {
    const e = await r.json().catch(() => ({}));
    throw new Error(`Claude API ${r.status}: ${e.error?.message || r.statusText}`);
  }
  const d = await r.json();
  return d.content.find(b => b.type === 'text')?.text || '';
}
```

- [ ] **Step 2: Función buildPrompt() con selección inteligente texto/imagen**

```js
// Detectar si el texto extraído de pdf.js es suficientemente legible
function isTextLegible(text) {
  if (!text || text.trim().length < 100) return false;
  // Buscar señales de contenido real (nombres, apellidos, código postal, fecha)
  const hasPersonalData = /[A-ZÁÉÍÓÚ]{3,}[\s,]+[A-ZÁÉÍÓÚ]{3,}|DNI|NIE|PASAPORTE/i.test(text);
  const hasSubjects = /\d{7,8}\s+[A-ZÁÉÍÓÚ]/i.test(text);
  return hasPersonalData || hasSubjects;
}

function buildPrompt() {
  const content = [];
  const textLegible = isTextLegible(S.pdfText);

  if (textLegible && (!S.pdfPages || S.pdfPages.length === 0)) {
    // Solo texto — más barato, suficiente para PDFs digitales
    content.push({type:'text', text:`TEXTO DEL PDF:\n${S.pdfText}`});
  } else {
    // Enviar imágenes (PDF escaneado o datos personales en imagen)
    // Solo página 1 para datos personales + páginas con asignaturas (buscar códigos 5500xxxx)
    const pagesToSend = S.pdfPages.slice(0, Math.min(S.pdfPages.length, 6)); // máx 6 páginas
    pagesToSend.forEach((b64, i) => {
      content.push({type:'text', text:`=== Página ${i+1} ===`});
      content.push({type:'image', source:{type:'base64', media_type:'image/jpeg', data:b64}});
    });
    if (textLegible) {
      content.push({type:'text', text:`\nTEXTO EXTRAÍDO:\n${S.pdfText.substring(0, 3000)}`});
    }
  }

  // Páginas del certificado académico
  if (S.certPages && S.certPages.length) {
    content.push({type:'text', text:'\n=== CERTIFICADO ACADÉMICO ==='});
    S.certPages.slice(0, 3).forEach((b64, i) => {
      content.push({type:'text', text:`Página certificado ${i+1}:`});
      content.push({type:'image', source:{type:'base64', media_type:'image/jpeg', data:b64}});
    });
  }

  content.push({type:'text', text:`
Eres gestor académico UPM. Devuelve ÚNICAMENTE JSON válido sin markdown.

{"datos_personales":{"apellidos":"","nombre":"","pasaporte_nie_dni":"","nacionalidad":"","telefono":"","email":"","titulacion_origen":"","universidad_origen":"","pais_origen":"","titulacion_destino":"","fecha_solicitud":""},"asignaturas":[{"codigo_upm":"","nombre_upm":"","origen":[{"codigo_orig":"","nombre_orig":"","nota":"","creditos":"","anio":""},{"codigo_orig":"","nombre_orig":"","nota":"","creditos":"","anio":""},{"codigo_orig":"","nombre_orig":"","nota":"","creditos":"","anio":""}]}]}

NOTAS: origen siempre 3 elementos. Notas en letras→números. titulacion_destino es la UPM destino.`});

  return content;
}
```

- [ ] **Step 3: Verificar en consola que buildPrompt() no produce errores sin PDF cargado**

```js
// Sin PDF cargado, solo debe retornar content array con texto vacío
const p = buildPrompt();
console.assert(Array.isArray(p), 'buildPrompt debe retornar array');
console.assert(p.length >= 1, 'buildPrompt debe tener al menos 1 elemento');
console.log('✅ buildPrompt OK, elementos:', p.length);
```

- [ ] **Step 4: Commit**
```bash
git add "Recon.html"
git commit -m "feat: token-optimized Claude extraction (text-first, vision fallback)"
```

---

## Task 7: Precedent matching v2 + clasificación automática

**Files:**
- Modify: `Recon.html` — sección `// ── ENRICH ──`

Implementa los umbrales corregidos: universidad ≥0.6 (duda 0.4-0.59), titulación ≥0.65 (duda 0.50-0.64). Zona de duda → estado `en_duda` en lugar de descartarlo silenciosamente.

- [ ] **Step 1: Función enrichExcel() (port con corrección del acceso a email)**

```js
// ── ENRICH ───────────────────────────────────────────
function enrichExcel(rows) {
  const idx = {};
  S.excelData.forEach((d, ei) => {
    if (!d) return;
    d.forEach(r => {
      const c = String(r['CÓDIGO'] || r['CODIGO'] || '').trim();
      if (c && !idx[c]) idx[c] = {...r, _ei: ei};
    });
  });
  return rows.map(r => {
    const c = String(r.codigo_upm || '').trim();
    const match = idx[c];
    if (match) {
      const tratamiento = (match['TRATAMIENTO'] || '').trim();
      const nombrePila = (match['NOMBRE DE PILA'] || '').trim();
      const ponentePie = (match['PONENTE PIE  DE PAGINA'] || match['PONENTE'] || '').trim();
      // Buscar email en columnas del Excel de forma segura
      const emailKey = Object.keys(match).find(k => String(match[k]).includes('@')) || '';
      const email = emailKey ? match[emailKey] : '';
      return {
        ...r,
        creditos_upm: match['CRÉDITOS'] || match['CREDITOS'] || '',
        tipo: match['TIPO'] || '',
        curso: match['CURSO'] || match['CURSO / ESPECIALIDAD'] || '',
        nombre_depto: match['NOMBRE DEPARTAMENTO'] || '',
        ponente_pie: ponentePie,
        nombre_pila: nombrePila,
        tratamiento,
        email_ponente: email,
        _excelMatch: true,
        _excelRow: match
      };
    }
    return {...r, _excelMatch: false};
  });
}
```

- [ ] **Step 2: detectGrado() (port sin cambios)**

```js
function detectGrado() {
  const titDest = norm(S.personalData.titulacion_destino || '');
  if (titDest.includes('05IR')) return '05IR';
  if (titDest.includes('05IQ')) return '05IQ';
  if (titDest.includes('05TI')) return '05TI';
  if (titDest.includes('ORGANIZACION') || titDest.includes('ORGANIZACIÓN')) return '05IR';
  if (titDest.includes('QUIMICA') || titDest.includes('QUÍMICA')) return '05IQ';
  if (titDest.includes('TECNOLOG')) return '05TI';
  if (titDest.includes('INDUSTRIAL') && !titDest.includes('ORGANIZACION')) return '05TI';
  const counts = S.excelData.map((d) => {
    if (!d) return 0;
    const codes = new Set(d.map(r => String(r['CÓDIGO'] || r['CODIGO'] || '').trim()));
    return S.asignaturas.filter(a => codes.has(String(a.codigo_upm || '').trim())).length;
  });
  const best = counts.indexOf(Math.max(...counts));
  if (counts[best] === 0) return null;
  return ['05IQ','05IR','05TI'][best];
}

function codigoMatch(p, codigoUPM) {
  return p.reconocidas.some(r => {
    const rc = r.codigo.trim();
    return rc === codigoUPM || '5'+rc === codigoUPM || rc === '5'+codigoUPM || rc.replace(/^5/,'55') === codigoUPM;
  });
}
```

- [ ] **Step 3: enrichPrec() v2 con umbrales corregidos y zona en_duda**

```js
// Umbrales
const UNIV_CONFIRM = 0.6, UNIV_DOUBT = 0.4;
const TIT_CONFIRM = 0.65, TIT_DOUBT = 0.50;
const ASIG_CONFIRM = 0.45;

function enrichPrec(rows) {
  const grado = detectGrado();
  const univOrigen = S.personalData.universidad_origen || '';
  const titOrigen = norm(S.personalData.titulacion_origen || '');

  // Paso 1: filtrar por grado
  const porGrado = grado ? S.precedentes.filter(p => p._grado === grado) : S.precedentes;
  if (!porGrado.length) return noPrec(rows);

  // Paso 2: filtrar por universidad (confirmado ≥0.6, duda 0.4-0.59)
  let univDuda = false;
  const porUnivConfirm = porGrado.filter(p =>
    (p.titulaciones || []).some(t => univSimilarity(t.universidad, univOrigen) >= UNIV_CONFIRM)
  );
  const porUnivDuda = porGrado.filter(p =>
    (p.titulaciones || []).some(t => {
      const s = univSimilarity(t.universidad, univOrigen);
      return s >= UNIV_DOUBT && s < UNIV_CONFIRM;
    })
  );

  const porUniv = porUnivConfirm.length > 0 ? porUnivConfirm :
    (porUnivDuda.length > 0 ? (univDuda = true, porUnivDuda) : []);
  if (!porUniv.length) return noPrec(rows);

  // Paso 3: filtrar por titulación (confirmado ≥0.65, duda 0.50-0.64)
  let titDuda = false;
  const porTitConfirm = porUniv.filter(p =>
    (p.titulaciones || []).some(t =>
      univSimilarity(t.universidad, univOrigen) >= UNIV_DOUBT &&
      similarity(t.titulacion, titOrigen) >= TIT_CONFIRM
    )
  );
  const porTitDuda = porUniv.filter(p =>
    (p.titulaciones || []).some(t => {
      const s = similarity(t.titulacion, titOrigen);
      return s >= TIT_DOUBT && s < TIT_CONFIRM &&
        univSimilarity(t.universidad, univOrigen) >= UNIV_DOUBT;
    })
  );

  const porTit = porTitConfirm.length > 0 ? porTitConfirm :
    (porTitDuda.length > 0 ? (titDuda = true, porTitDuda) : []);
  if (!porTit.length) return noPrec(rows);

  const hayDuda = univDuda || titDuda;

  // Paso 4: cruce asignatura a asignatura
  return rows.map(row => {
    const codigoUPM = String(row.codigo_upm || '').trim();
    const solOrigenNames = (row.origen || []).filter(o => o.nombre_orig).map(o => norm(o.nombre_orig));

    const matches = porTit.filter(p => {
      if (!codigoMatch(p, codigoUPM)) return false;
      if (solOrigenNames.length === 0) return true;
      return (p.titulaciones || []).some(t =>
        univSimilarity(t.universidad, univOrigen) >= UNIV_DOUBT &&
        similarity(t.titulacion, titOrigen) >= TIT_DOUBT &&
        (t.asigs_reconocimiento || []).some(pa =>
          solOrigenNames.some(sn => similarity(pa, sn) >= ASIG_CONFIRM)
        )
      );
    });

    const estadoAsig = matches.length > 0
      ? (hayDuda ? 'en_duda' : 'automatico')
      : 'pendiente';

    return {
      ...row,
      _precMatches: matches,
      _precCount: matches.length,
      estado_asig: estadoAsig,
      resolucion: matches.length > 0 && !hayDuda ? 'Favorable' : ''
    };
  });
}

function noPrec(rows) {
  return rows.map(r => ({...r, _precMatches:[], _precCount:0, estado_asig:'pendiente', resolucion:''}));
}
```

- [ ] **Step 4: autoClassify() — clasificación de la solicitud completa**

```js
// Clasifica la solicitud completa según sus asignaturas
// Regla: TODAS automatico → 'automatico'
//        Alguna en_duda → 'en_duda'
//        Cualquier pendiente → 'comision'
function autoClassify(asignaturas) {
  const estados = asignaturas.map(a => a.estado_asig || 'pendiente');
  if (estados.some(e => e === 'en_duda')) return 'en_duda';
  if (estados.every(e => e === 'automatico')) return 'automatico';
  return 'comision';
}
```

- [ ] **Step 5: Test de clasificación en consola**

```js
// Test autoClassify
const mockTodosAuto = [{estado_asig:'automatico'},{estado_asig:'automatico'}];
const mockConPendiente = [{estado_asig:'automatico'},{estado_asig:'pendiente'}];
const mockConDuda = [{estado_asig:'automatico'},{estado_asig:'en_duda'}];
console.assert(autoClassify(mockTodosAuto) === 'automatico', 'Todos auto → automatico');
console.assert(autoClassify(mockConPendiente) === 'comision', 'Con pendiente → comision');
console.assert(autoClassify(mockConDuda) === 'en_duda', 'Con duda → en_duda');
console.log('✅ autoClassify OK');
```

- [ ] **Step 6: Commit**
```bash
git add "Recon.html"
git commit -m "feat: enrichPrec v2 with corrected thresholds, en_duda zone, autoClassify"
```

---

## Task 8: Sección Nueva Solicitud (flujo completo)

**Files:**
- Modify: `Recon.html` — función `renderNueva()`, `startProcessing()`, modal de duplicado

- [ ] **Step 1: Función renderNueva() — HTML de la sección**

```js
// ── RENDER — NUEVA ────────────────────────────────────
function renderNueva() {
  const sec = document.getElementById('sec-nueva');
  if (sec._rendered) return; // evitar re-render innecesario
  sec._rendered = true;
  sec.innerHTML = `
<div style="max-width:760px">
  <div style="margin-bottom:18px">
    <h2 style="font-size:18px;font-weight:700;margin-bottom:4px">Nueva Solicitud</h2>
    <p style="color:var(--tx3);font-size:13px">Procesa una nueva solicitud de reconocimiento de créditos</p>
  </div>

  <!-- PDF solicitud -->
  <div class="card">
    <div class="card-header">
      <svg viewBox="0 0 24 24" stroke="var(--blue)"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <h2>PDF de Solicitud</h2>
      <span class="badge b-blue">Obligatorio</span>
    </div>
    <div class="dz" id="pdfDZ" onclick="document.getElementById('pdfIn').click()"
         ondragover="event.preventDefault();this.classList.add('over')"
         ondragleave="this.classList.remove('over')"
         ondrop="event.preventDefault();this.classList.remove('over');handlePdfFile(event.dataTransfer.files[0])">
      <input type="file" id="pdfIn" accept=".pdf" onchange="handlePdfFile(this.files[0])">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <p>Arrastra el PDF o <strong>selecciona el archivo</strong></p>
    </div>
    <div id="pdfOk"></div>
  </div>

  <!-- Certificado académico (opcional) -->
  <div class="card">
    <div class="card-header">
      <svg viewBox="0 0 24 24" stroke="var(--amber)"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
      <h2>Certificado Académico</h2>
      <span class="badge b-amber">Opcional — notas y créditos</span>
    </div>
    <div class="dz" id="certDZ" onclick="document.getElementById('certIn').click()"
         ondragover="event.preventDefault();this.classList.add('over')"
         ondragleave="this.classList.remove('over')"
         ondrop="event.preventDefault();this.classList.remove('over');handleCertFile(event.dataTransfer.files[0])">
      <input type="file" id="certIn" accept=".pdf" onchange="handleCertFile(this.files[0])">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <p>Arrastra el certificado o <strong>selecciónalo</strong></p>
    </div>
    <div id="certOk"></div>
  </div>

  <!-- Estado planes y precedentes -->
  <div class="card">
    <div class="card-header">
      <svg viewBox="0 0 24 24" stroke="var(--green)"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>
      <h2>Datos internos</h2>
      <span class="badge b-green">Cargados automáticamente</span>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">
      <div style="background:var(--bg);border-radius:var(--r);padding:10px;border:1px solid var(--border)">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--tx3);margin-bottom:3px">05IQ</div>
        <div style="font-size:11px;color:var(--tx2)" id="plan-05IQ">Cargando…</div>
      </div>
      <div style="background:var(--bg);border-radius:var(--r);padding:10px;border:1px solid var(--border)">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--tx3);margin-bottom:3px">05IR</div>
        <div style="font-size:11px;color:var(--tx2)" id="plan-05IR">Cargando…</div>
      </div>
      <div style="background:var(--bg);border-radius:var(--r);padding:10px;border:1px solid var(--border)">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--tx3);margin-bottom:3px">05TI</div>
        <div style="font-size:11px;color:var(--tx2)" id="plan-05TI">Cargando…</div>
      </div>
    </div>
    <div style="margin-top:8px;font-size:11px;color:var(--tx3)" id="seed-status"></div>
  </div>

  <!-- API key -->
  <div class="card">
    <div class="card-header">
      <svg viewBox="0 0 24 24" stroke="var(--purple)"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>
      <h2>API Anthropic</h2>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <input type="password" id="apikey-input" class="form-input" placeholder="sk-ant-api03-…" style="flex:1;font-family:'JetBrains Mono',monospace;font-size:11px">
      <button class="btn btn-secondary btn-sm" onclick="toggleApiVis()">👁</button>
    </div>
    <p style="font-size:11px;color:var(--tx3);margin-top:6px">La clave se guarda localmente en IndexedDB.</p>
  </div>

  <div id="errbox" class="err-box"></div>
  <div style="display:flex;gap:9px;flex-wrap:wrap;margin-top:4px">
    <button class="btn btn-primary" id="procbtn" onclick="startProcessing()">
      <svg viewBox="0 0 24 24"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
      Procesar con IA
    </button>
    <button class="btn btn-secondary" onclick="resetNueva()">
      <svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3.42"/></svg>
      Limpiar
    </button>
  </div>
  <div class="proc-box" id="procbox">
    <div class="spin"></div>
    <div class="proc-title" id="pst">Iniciando…</div>
    <div class="proc-detail" id="pdt"></div>
  </div>
</div>`;

  // Cargar apikey guardada
  getConfig('apikey').then(k => {
    if (k) document.getElementById('apikey-input').value = k;
  });

  if (S.seedLoaded) updateSeedUI();
}
```

- [ ] **Step 2: Handlers de archivo PDF y certificado**

```js
async function handlePdfFile(file) {
  if (!file || file.type !== 'application/pdf') return showErr('Selecciona un archivo PDF válido.');
  hideErr();
  S.pdfFile = file;
  // Mostrar indicador de carga
  document.getElementById('pdfOk').innerHTML = '<div style="font-size:12px;color:var(--tx3);padding:6px">Procesando PDF…</div>';
  try {
    const ab = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: new Uint8Array(ab)}).promise;
    // Extraer texto
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const p = await pdf.getPage(i);
      const tc = await p.getTextContent();
      fullText += tc.items.map(t => t.str).join(' ') + '\n';
    }
    S.pdfText = fullText;
    // Renderizar páginas como imágenes (para visión si el texto no es legible)
    S.pdfPages = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 8); i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({scale: 1.5});
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({canvasContext: canvas.getContext('2d'), viewport}).promise;
      S.pdfPages.push(canvas.toDataURL('image/jpeg', 0.85).split(',')[1]);
    }
    document.getElementById('pdfOk').innerHTML = `
      <div class="file-ok">
        <svg viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>
        ${file.name} — ${pdf.numPages} páginas · texto ${isTextLegible(S.pdfText) ? 'legible (modo económico)' : 'no extraíble (modo visión)'}
      </div>`;
  } catch(e) {
    showErr('Error procesando PDF: ' + e.message);
    document.getElementById('pdfOk').innerHTML = '';
  }
}

async function handleCertFile(file) {
  if (!file || file.type !== 'application/pdf') return showErr('Selecciona un PDF para el certificado.');
  hideErr();
  document.getElementById('certOk').innerHTML = '<div style="font-size:12px;color:var(--tx3);padding:6px">Procesando certificado…</div>';
  try {
    const ab = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: new Uint8Array(ab)}).promise;
    S.certPages = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 4); i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({scale: 1.5});
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      await page.render({canvasContext: canvas.getContext('2d'), viewport}).promise;
      S.certPages.push(canvas.toDataURL('image/jpeg', 0.85).split(',')[1]);
    }
    document.getElementById('certOk').innerHTML = `
      <div class="file-ok">
        <svg viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>
        ${file.name} — ${pdf.numPages} páginas de certificado
      </div>`;
  } catch(e) { showErr('Error en certificado: ' + e.message); }
}

function toggleApiVis() {
  const inp = document.getElementById('apikey-input');
  inp.type = inp.type === 'password' ? 'text' : 'password';
}

function resetNueva() {
  S.pdfFile = null; S.pdfText = ''; S.pdfPages = []; S.certPages = [];
  S.personalData = {}; S.asignaturas = []; S.clasificacion = null; S.currentSolicitudId = null;
  document.getElementById('sec-nueva')._rendered = false;
  renderNueva();
}
```

- [ ] **Step 3: startProcessing() con guardado en BD y detección de duplicados**

```js
function setPS(title, detail = '') {
  document.getElementById('pst').textContent = title;
  document.getElementById('pdt').textContent = detail;
}

async function startProcessing() {
  hideErr();
  const apiKey = document.getElementById('apikey-input')?.value?.trim();
  if (!apiKey) return showErr('Introduce tu clave API de Anthropic.');
  if (!S.pdfPages?.length && !isTextLegible(S.pdfText)) return showErr('Primero carga el PDF de solicitud.');
  if (!S.seedLoaded) return showErr('Los datos internos aún se están cargando. Espera un momento.');

  // Guardar apikey para próximas sesiones
  await setConfig('apikey', apiKey);

  document.getElementById('procbox').classList.add('on');
  document.getElementById('procbtn').disabled = true;

  try {
    setPS('Claude analizando la solicitud…', 'Extrayendo datos personales y asignaturas');
    const raw = await callClaude(buildPrompt());
    const clean = raw.replace(/```json\n?|```/g, '').trim();
    let parsed;
    try { parsed = JSON.parse(clean); }
    catch { throw new Error('Claude devolvió JSON inválido. Intenta de nuevo.'); }

    S.personalData = parsed.datos_personales || {};
    let rows = parsed.asignaturas || [];
    setPS('Cruzando con planes de estudio…', `${rows.length} asignaturas encontradas`);
    rows = enrichExcel(rows);
    setPS('Buscando precedentes…', `Comparando contra ${S.precedentes.length} registros`);
    rows = enrichPrec(rows);
    S.asignaturas = rows;
    S.clasificacion = autoClassify(rows);

    // Guardar en BD (con detección de duplicados)
    const result = await saveSolicitud({
      personalData: S.personalData,
      asignaturas: S.asignaturas,
      clasificacion: S.clasificacion,
      estado: S.clasificacion,
      pdfText: S.pdfText.substring(0, 5000) // guardar extracto del texto
    });

    document.getElementById('procbox').classList.remove('on');
    document.getElementById('procbtn').disabled = false;

    if (result.duplicate) {
      showDuplicateModal(result.existing, S);
      return;
    }

    S.currentSolicitudId = result.id;
    showSection('solicitudes');
    // Mostrar la solicitud recién procesada en detalle
    setTimeout(() => openSolicitudDetail(result.id), 100);

  } catch(err) {
    document.getElementById('procbox').classList.remove('on');
    document.getElementById('procbtn').disabled = false;
    showErr(err.message || String(err));
  }
}
```

- [ ] **Step 4: Modal de duplicado**

```js
function showDuplicateModal(existing, newData) {
  const modal = document.getElementById('modal-duplicate');
  const dni = existing.personalData?.pasaporte_nie_dni || '';
  const nameExisting = `${existing.personalData?.apellidos || ''} ${existing.personalData?.nombre || ''}`.trim();
  const nameNew = `${newData.personalData?.apellidos || ''} ${newData.personalData?.nombre || ''}`.trim();
  modal.innerHTML = `
    <div class="modal">
      <div class="modal-head">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--amber)" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
        <h2>Alumno duplicado</h2>
      </div>
      <p style="color:var(--tx2);margin-bottom:14px;font-size:13px">El DNI/NIE <strong style="color:var(--tx)">${dni}</strong> ya existe en la base de datos.</p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px">
        <div style="background:var(--bg);border-radius:var(--r);padding:12px;border:1px solid var(--border)">
          <div style="font-size:10px;color:var(--tx3);margin-bottom:4px;text-transform:uppercase">Registro existente</div>
          <div style="font-weight:600">${nameExisting}</div>
          <div style="font-size:11px;color:var(--tx3)">${formatDate(existing.fecha_procesado)} · ${existing.asignaturas?.length || 0} asignaturas · <span class="estado-pill ep-${existing.clasificacion}">${existing.clasificacion}</span></div>
        </div>
        <div style="background:var(--bg);border-radius:var(--r);padding:12px;border:1px solid var(--border)">
          <div style="font-size:10px;color:var(--tx3);margin-bottom:4px;text-transform:uppercase">Nueva solicitud</div>
          <div style="font-weight:600">${nameNew}</div>
          <div style="font-size:11px;color:var(--tx3)">Hoy · ${newData.asignaturas?.length || 0} asignaturas · <span class="estado-pill ep-${newData.clasificacion}">${newData.clasificacion}</span></div>
        </div>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-danger" onclick="resolveDuplicate('replace',${existing.id})">Reemplazar registro existente</button>
        <button class="btn btn-secondary" onclick="resolveDuplicate('merge',${existing.id})">Fusionar (revisión manual)</button>
        <button class="btn btn-secondary" onclick="closeDuplicateModal()">Cancelar</button>
      </div>
    </div>`;
  modal.classList.add('on');
}

async function resolveDuplicate(action, existingId) {
  closeDuplicateModal();
  if (action === 'replace') {
    await deleteSolicitud(existingId);
    const result = await dbAdd('solicitudes', {
      personalData: S.personalData,
      asignaturas: S.asignaturas,
      clasificacion: S.clasificacion,
      estado: S.clasificacion,
      dni: S.personalData.pasaporte_nie_dni?.trim() || '',
      pdfText: S.pdfText.substring(0, 5000),
      fecha_procesado: new Date().toISOString(),
      fecha_modificado: new Date().toISOString()
    });
    S.currentSolicitudId = result;
  } else {
    // Fusionar: marcar existente como en_duda para revisión manual
    await updateSolicitud(existingId, {
      estado: 'en_duda',
      _mergeCandidate: { personalData: S.personalData, asignaturas: S.asignaturas }
    });
    S.currentSolicitudId = existingId;
  }
  showSection('solicitudes');
  setTimeout(() => openSolicitudDetail(S.currentSolicitudId), 100);
}

function closeDuplicateModal() {
  document.getElementById('modal-duplicate').classList.remove('on');
}
```

- [ ] **Step 5: Verificar flujo completo manualmente**

1. Abrir Recon.html → sección "Nueva Solicitud"
2. Cargar un PDF de prueba (cualquier PDF)
3. El indicador debe mostrar si el texto es legible o no
4. Click "Procesar con IA" → debe llamar a Claude y guardar en BD
5. Si el PDF ya fue procesado con el mismo DNI → debe aparecer modal de duplicado
6. La solicitud procesada debe aparecer en sección "Solicitudes"

- [ ] **Step 6: Commit**
```bash
git add "Recon.html"
git commit -m "feat: Nueva Solicitud complete — file upload, processing, BD save, duplicate modal"
```

---

## Task 9: Sección Solicitudes — Vista básica de listado

**Files:**
- Modify: `Recon.html` — función `renderSolicitudes()`, `openSolicitudDetail()`

- [ ] **Step 1: renderSolicitudes() con tabla y alertas en_duda**

```js
// ── RENDER — SOLICITUDES ─────────────────────────────
async function renderSolicitudes() {
  const sec = document.getElementById('sec-solicitudes');
  const solicitudes = await getAllSolicitudes();
  const enDuda = solicitudes.filter(s => s.estado === 'en_duda' || s.clasificacion === 'en_duda');

  let html = `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
      <h2 style="font-size:18px;font-weight:700">Solicitudes</h2>
      <span style="color:var(--tx3);font-size:13px">${solicitudes.length} registros</span>
    </div>`;

  // Alertas en_duda
  if (enDuda.length) {
    html += `<div class="alert-banner">
      <svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
      <div class="alert-banner-text">⚠️ ${enDuda.length} solicitud${enDuda.length>1?'es':''} requiere${enDuda.length>1?'n':''} tu revisión (similitud insuficiente)</div>
    </div>`;
  }

  if (!solicitudes.length) {
    html += `<div style="text-align:center;padding:60px 20px;color:var(--tx3)">
      <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" style="margin-bottom:12px"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <p>No hay solicitudes. <button class="btn btn-secondary btn-sm" onclick="showSection('nueva')">Procesar la primera</button></p>
    </div>`;
  } else {
    html += `<div class="table-wrap">
      <table>
        <thead><tr>
          <th>Alumno</th><th>DNI/NIE</th><th>Universidad origen</th>
          <th>Asignaturas</th><th>Clasificación</th><th>Estado</th><th>Fecha</th><th></th>
        </tr></thead>
        <tbody>`;
    solicitudes.sort((a,b) => new Date(b.fecha_procesado) - new Date(a.fecha_procesado)).forEach(s => {
      const name = `${s.personalData?.apellidos||''} ${s.personalData?.nombre||''}`.trim() || '—';
      const univ = s.personalData?.universidad_origen || '—';
      const nAsig = s.asignaturas?.length || 0;
      const clasif = s.clasificacion || s.estado || 'nuevo';
      const isAlert = clasif === 'en_duda';
      html += `<tr style="${isAlert ? 'background:var(--redBg)' : ''}">
        <td><strong>${name}</strong></td>
        <td><span class="code">${s.personalData?.pasaporte_nie_dni||'—'}</span></td>
        <td style="font-size:11px;color:var(--tx2)">${univ}</td>
        <td style="text-align:center">${nAsig}</td>
        <td><span class="estado-pill ep-${clasif}">${clasif}</span></td>
        <td><span class="estado-pill ep-${s.estado||'nuevo'}">${s.estado||'nuevo'}</span></td>
        <td style="font-size:11px;color:var(--tx3)">${formatDate(s.fecha_procesado)}</td>
        <td>
          <button class="btn btn-secondary btn-sm" onclick="openSolicitudDetail(${s.id})">Ver</button>
        </td>
      </tr>`;
    });
    html += '</tbody></table></div>';
  }

  sec.innerHTML = html;
  // Actualizar badge en nav
  const badge = document.getElementById('badge-duda');
  if (badge) { badge.style.display = enDuda.length ? 'inline' : 'none'; badge.textContent = enDuda.length; }
}
```

- [ ] **Step 2: openSolicitudDetail() — vista básica (se expande en Plan 2)**

```js
async function openSolicitudDetail(id) {
  const s = await dbGet('solicitudes', id);
  if (!s) return;
  // En Plan 1: mostrar en la misma sección como panel expandido
  // Plan 2 implementará la vista completa con edición inline
  const name = `${s.personalData?.apellidos||''} ${s.personalData?.nombre||''}`.trim();
  const sec = document.getElementById('sec-solicitudes');
  let html = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">
      <button class="btn btn-secondary btn-sm" onclick="renderSolicitudes()">← Volver</button>
      <h2 style="font-size:18px;font-weight:700">${name}</h2>
      <span class="estado-pill ep-${s.clasificacion||'nuevo'}">${s.clasificacion||'nuevo'}</span>
      ${s.clasificacion==='en_duda'?'<div class="alert-banner" style="flex:1;margin:0"><svg viewBox="0 0 24 24" width="16" height="16"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/></svg><div class="alert-banner-text">Revisión humana requerida</div></div>':''}
    </div>
    <div class="card">
      <div class="card-header"><h2>Datos personales</h2></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
        ${Object.entries(s.personalData||{}).map(([k,v])=>`
          <div>
            <div style="font-size:10px;color:var(--tx3);text-transform:uppercase;margin-bottom:2px">${k.replace(/_/g,' ')}</div>
            <div style="font-size:13px;font-weight:500">${v||'—'}</div>
          </div>`).join('')}
      </div>
    </div>
    <div class="card">
      <div class="card-header"><h2>Asignaturas</h2><span class="badge b-blue">${s.asignaturas?.length||0} total</span></div>
      <div class="table-wrap">
        <table><thead><tr>
          <th>#</th><th>Código UPM</th><th>Asignatura UPM</th><th>Estado</th><th>Resolución</th><th>Nº Precedente</th>
        </tr></thead><tbody>
          ${(s.asignaturas||[]).map((a,i)=>`<tr>
            <td style="color:var(--tx3)">${i+1}</td>
            <td><span class="code">${a.codigo_upm||'—'}</span></td>
            <td>${a.nombre_upm||'—'}</td>
            <td><span class="estado-pill ep-${a.estado_asig||'pendiente'}">${a.estado_asig||'pendiente'}</span></td>
            <td>${a.resolucion||'—'}</td>
            <td style="font-size:11px;color:var(--tx3)">${a._precMatches?.map(p=>`Nº${p.numero}`).slice(0,3).join(', ')||'—'}</td>
          </tr>`).join('')}
        </tbody></table>
      </div>
    </div>`;
  sec.innerHTML = html;
}
```

- [ ] **Step 3: Verificar listado**

1. Abrir sección "Solicitudes"
2. Deben aparecer las solicitudes procesadas
3. Las solicitudes `en_duda` deben tener fondo rojo y banner de alerta
4. Click "Ver" → debe mostrar el detalle con datos personales y tabla de asignaturas

- [ ] **Step 4: Commit**
```bash
git add "Recon.html"
git commit -m "feat: Solicitudes list with en_duda alerts and basic detail view"
```

---

## Task 10: Sistema de backup automático

**Files:**
- Modify: `Recon.html` — sección `// ── BACKUP ──`

- [ ] **Step 1: exportBackup() — exportar toda la BD a JSON**

```js
// ── BACKUP ───────────────────────────────────────────
async function exportBackup() {
  try {
    const [solicitudes, config, precCache, planCache] = await Promise.all([
      dbGetAll('solicitudes'),
      dbGetAll('config'),
      dbGetAll('precedentes_cache'),
      dbGetAll('planes_cache'),
    ]);
    const backup = {
      version: 1,
      fecha: new Date().toISOString(),
      solicitudes,
      config: config.filter(c => c.key !== 'apikey'), // No exportar la API key
      precedentes_cache: precCache,
      planes_cache: planCache,
    };
    const json = JSON.stringify(backup, null, 2);
    const blob = new Blob([json], {type:'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `recon-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
    await setConfig('ultima_exportacion', new Date().toISOString());
  } catch(e) {
    alert('Error al exportar: ' + e.message);
  }
}

async function importBackup(file) {
  try {
    const text = await file.text();
    const backup = JSON.parse(text);
    if (!backup.version || !backup.solicitudes) throw new Error('Formato de backup inválido.');
    const confirm = window.confirm(
      `Importar backup del ${formatDate(backup.fecha)}?\n` +
      `${backup.solicitudes.length} solicitudes serán AÑADIDAS (no borra las existentes).\n\n` +
      `Continuar?`
    );
    if (!confirm) return;
    // Importar solicitudes (sin duplicar por DNI)
    let importadas = 0, omitidas = 0;
    for (const s of backup.solicitudes) {
      const {id, ...rest} = s; // quitar id para que autoincrement asigne uno nuevo
      const dni = rest.personalData?.pasaporte_nie_dni?.trim();
      if (dni) {
        const existing = await dbGetByIndex('solicitudes', 'dni', dni);
        if (existing) { omitidas++; continue; }
      }
      await dbAdd('solicitudes', rest);
      importadas++;
    }
    // Importar caché si no existe
    if (backup.precedentes_cache) {
      for (const item of backup.precedentes_cache) {
        const existing = await dbGet('precedentes_cache', item.key);
        if (!existing) await dbPut('precedentes_cache', item);
      }
    }
    alert(`Importación completada:\n✅ ${importadas} solicitudes importadas\n⏭ ${omitidas} omitidas (DNI duplicado)`);
    renderSolicitudes();
  } catch(e) {
    alert('Error importando backup: ' + e.message);
  }
}
```

- [ ] **Step 2: Auto-export al cerrar la pestaña**

```js
// ── INIT ─────────────────────────────────────────────
async function init() {
  await openDB();
  // Auto-backup al cerrar (silent — no descarga si no hay solicitudes)
  window.addEventListener('beforeunload', async () => {
    const sols = await getAllSolicitudes();
    if (sols.length > 0) exportBackup();
  });
  // Inicializar seed
  await loadSeedIfNeeded();
  // Mostrar sección inicial
  showSection('nueva');
  // Verificar si hay importación pendiente
  checkImportOnStartup();
}

async function checkImportOnStartup() {
  const sols = await getAllSolicitudes();
  if (sols.length === 0) {
    // BD vacía: ofrecer importar backup
    // (Solo muestra el mensaje si ya existe el elemento de la sección nueva renderizada)
    setTimeout(() => {
      const statusEl = document.getElementById('seed-status');
      if (statusEl && sols.length === 0) {
        statusEl.innerHTML = `Base de datos vacía. <button class="btn btn-secondary btn-sm" onclick="document.getElementById('importInput').click()">Importar backup</button>
          <input type="file" id="importInput" accept=".json" style="display:none" onchange="importBackup(this.files[0])">`;
      }
    }, 2000);
  }
}

// Arrancar cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', init);
```

- [ ] **Step 3: Stub de secciones no implementadas en Plan 1**

Para que la navegación no rompa al hacer click en secciones de planes futuros:

```js
function renderConvocatorias() {
  document.getElementById('sec-convocatorias').innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--tx3)">Convocatorias — disponible en v2 Plan 2</div>';
}
function renderCorreos() {
  document.getElementById('sec-correos').innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--tx3)">Correos — disponible en v2 Plan 3</div>';
}
function renderInformes() {
  document.getElementById('sec-informes').innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--tx3)">Informes — disponible en v2 Plan 3</div>';
}
function renderEstadisticas() {
  document.getElementById('sec-estadisticas').innerHTML =
    '<div style="padding:40px;text-align:center;color:var(--tx3)">Estadísticas — disponible en v2 Plan 4</div>';
}
function renderAdmin() {
  document.getElementById('sec-admin').innerHTML = `
    <div style="max-width:600px">
      <h2 style="font-size:18px;font-weight:700;margin-bottom:16px">Administración</h2>
      <div class="card">
        <div class="card-header"><h2>Clave API Anthropic</h2></div>
        <input type="password" id="admin-apikey" class="form-input" placeholder="sk-ant-api03-…" style="font-family:'JetBrains Mono',monospace;font-size:11px">
        <button class="btn btn-primary" style="margin-top:8px" onclick="saveAdminApiKey()">Guardar</button>
      </div>
      <div class="card">
        <div class="card-header"><h2>Base de datos</h2></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn btn-secondary" onclick="exportBackup()">↓ Exportar backup</button>
          <button class="btn btn-secondary" onclick="document.getElementById('importInput2').click()">
            ↑ Importar backup
            <input type="file" id="importInput2" accept=".json" style="display:none" onchange="importBackup(this.files[0])">
          </button>
        </div>
      </div>
    </div>`;
  getConfig('apikey').then(k => {
    const el = document.getElementById('admin-apikey');
    if (el && k) el.value = k;
  });
}

async function saveAdminApiKey() {
  const k = document.getElementById('admin-apikey')?.value?.trim();
  if (!k) return;
  await setConfig('apikey', k);
  alert('Clave guardada.');
}
```

- [ ] **Step 4: Test del backup**

1. Abrir Recon.html, procesar al menos una solicitud de prueba
2. Click "↓ Exportar BD" en la barra superior
3. Debe descargarse `recon-backup-YYYY-MM-DD.json`
4. Abrir el JSON — debe contener objeto con `solicitudes`, `version`, `fecha`
5. En Administración, click "Importar backup" y seleccionar el mismo JSON
6. Debe mostrar "X solicitudes importadas, Y omitidas (duplicado)"

- [ ] **Step 5: Commit final del Plan 1**

```bash
git add "Recon.html" "recon-data.js"
git commit -m "feat: auto-backup on close, import on startup, admin stubs — Plan 1 complete"
```

---

## Self-Review del Plan 1

**Spec coverage:**
- ✅ Dark Dashboard (Task 3)
- ✅ Arquitectura dos archivos HTML + recon-data.js (Task 1)
- ✅ IndexedDB con 4 stores (Task 2)
- ✅ Auto-backup al cerrar / importar al abrir (Task 10)
- ✅ Bugfix HEADER_B64 (Task 1, step 2)
- ✅ Bugfix window.open null check — no aplica en Plan 1 (está en generación de cartas, Plan 2)
- ✅ Bugfix a.origen null (Task 7, enrichPrec usa `(row.origen || [])`)
- ✅ Umbral titulación corregido a 0.65 (Task 7)
- ✅ Token optimization texto/imagen (Task 6)
- ✅ Clasificación automático/comisión/en_duda (Task 7)
- ✅ Detección de duplicados + modal (Task 8)
- ✅ Solicitudes básicas con alertas en_duda (Task 9)
- ⏭ Convocatorias, Correos, Informes → Plan 2 y 3
- ⏭ Estadísticas, Admin completo → Plan 4

**Placeholders encontrados:** ninguno — todo el código está completo.

**Consistencia de tipos:**
- `saveSolicitud()` retorna `{duplicate, id}` — usado correctamente en `startProcessing()`
- `S.excelData` es array[3] — accedido como `S.excelData[0|1|2]` consistentemente
- `estado_asig` en asignaturas, `clasificacion` en solicitud — nomenclatura consistente
