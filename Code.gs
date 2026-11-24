/**
 * SMART · Repositorio de Clubes de Conversación y SmartZone
 * Backend en Google Apps Script (va pegado en script.google.com)
 *
 * Qué hace:
 *  - Crea (si no existe) toda la jerarquía de carpetas en tu Google Drive.
 *  - Le entrega a la página HTML los enlaces de cada carpeta.
 *  - Recibe los planeadores que suben los docentes y los guarda en la carpeta elegida.
 *  - Lleva un registro de cada subida en una hoja de cálculo (03_GESTION.../Bitacoras_y_Actas).
 *
 * Configura las 3 constantes de abajo, ejecuta setup() una vez y luego
 * Implementar > Nueva implementación > Aplicación web
 *   (Ejecutar como: Yo · Quién tiene acceso: Cualquier persona).
 */

// ============ CONFIGURACIÓN ============
const ROOT_FOLDER_ID = '';            // ID de la carpeta principal si ya existe. Vacío = se crea en "Mi unidad".
const ACCESS_CODE    = 'CAMBIA-ESTE-CODIGO'; // Código que comparten los docentes para poder subir.
const MAX_MB         = 15;            // Tamaño máximo por archivo.
// ======================================

const ROOT_NAME = 'CLUBES DE CONVERSACIÓN Y SMARTZONE';
const MASTER = '00_DOCUMENTOS_MAESTROS_Y_GUIAS';
const ALLOWED_EXT = ['docx', 'doc', 'pdf'];
const LOG_PATH = '03_GESTION_Y_SEGUIMIENTO_CELULAS/Bitacoras_y_Actas';
const LOG_NAME = 'Registro_Subidas_Planeadores';

const CLUBS_EN = ['01_Talk_Time', '02_Cultural_Day', '03_Tongue_Gym', '04_Beyond_The_Classroom', '05_Cinemas_Day'];
const CLUBS_FR = ['01_Prenez_La_Parole', '02_Journee_Culturelle', '03_Entrainez_La_Langue', '04_Au_Dela_De_La_Classe', '05_Cinema_Francophone'];

function sub_(names) { const o = {}; names.forEach(n => o[n] = {}); return o; }

const STRUCTURE = {
  [MASTER]: {},
  '01_ENGLISH': {
    'A1-A2_Basico': sub_(CLUBS_EN),
    'B1-B2-C1_Avanzado': sub_(CLUBS_EN),
    'ALL-LEVELS_Extracurriculares_180min': sub_(['Megaclub', 'Salida_Inmersion', 'Realidad_Virtual'])
  },
  '02_FRANCAIS': {
    'A1-A2_Debutant': sub_(CLUBS_FR),
    'B1-B2_Avance': sub_(CLUBS_FR),
    'ALL-LEVELS_Activites_Parascolaires_180min': sub_(['Megaclub', 'Sortie_Immersion', 'Realite_Virtuelle'])
  },
  '03_GESTION_Y_SEGUIMIENTO_CELULAS': sub_(['Bitacoras_y_Actas', 'Kit_Publicidad_y_Recompensas', 'Tableros_KPI_y_Encuestas'])
};

/** Ejecútalo una vez desde el editor para crear las carpetas y autorizar permisos. */
function setup() {
  const map = buildTree_();
  Logger.log('Carpeta principal: https://drive.google.com/drive/folders/' + map[''].id);
  Logger.log(Object.keys(map).length + ' carpetas listas.');
}

// ---------- Árbol de carpetas ----------
function getRoot_() {
  if (ROOT_FOLDER_ID) return DriveApp.getFolderById(ROOT_FOLDER_ID);
  const it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME);
}

function child_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function buildTree_() {
  const root = getRoot_();
  const map = { '': { id: root.getId(), name: root.getName() } };
  (function walk(folder, node, prefix) {
    Object.keys(node).forEach(name => {
      const f = child_(folder, name);
      const path = prefix ? prefix + '/' + name : name;
      map[path] = { id: f.getId(), name: name };
      walk(f, node[name], path);
    });
  })(root, STRUCTURE, '');
  PropertiesService.getScriptProperties().setProperty('TREE', JSON.stringify(map));
  return map;
}

function getTree_(refresh) {
  if (!refresh) {
    const cached = PropertiesService.getScriptProperties().getProperty('TREE');
    if (cached) return JSON.parse(cached);
  }
  return buildTree_();
}

// ---------- Endpoints ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'files') {
      if (p.code !== ACCESS_CODE) return json_({ ok: false, error: 'Código de acceso incorrecto.' });
      return json_({ ok: true, files: listFiles_(p.path || '') });
    }
    const tree = getTree_(p.refresh === '1');
    const folders = {};
    Object.keys(tree).forEach(k => folders[k] = tree[k].id);
    return json_({ ok: true, folders: folders, masterDocs: listFiles_(MASTER, tree), maxMB: MAX_MB });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.code !== ACCESS_CODE) return json_({ ok: false, error: 'Código de acceso incorrecto.' });

    const tree = getTree_();
    const path = String(body.path || '');
    if (!path || !tree[path]) return json_({ ok: false, error: 'Carpeta de destino no válida.' });
    if (path === MASTER) return json_({ ok: false, error: 'Los documentos maestros solo los actualiza coordinación.' });

    const teacher = clean_(body.teacher || '').slice(0, 60);
    if (!teacher) return json_({ ok: false, error: 'Falta el nombre del docente.' });

    const original = String(body.filename || 'planeador');
    const ext = (original.split('.').pop() || '').toLowerCase();
    if (ALLOWED_EXT.indexOf(ext) === -1) return json_({ ok: false, error: 'Solo se aceptan archivos .docx, .doc o .pdf.' });

    const bytes = Utilities.base64Decode(body.data || '');
    if (!bytes.length) return json_({ ok: false, error: 'El archivo está vacío.' });
    if (bytes.length > MAX_MB * 1024 * 1024) return json_({ ok: false, error: 'El archivo supera ' + MAX_MB + ' MB.' });

    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const topic = clean_(body.topic || original.replace(/\.[^.]+$/, '')).slice(0, 80);
    const finalName = [date, teacher, topic].join('_').replace(/\s+/g, '_') + '.' + ext;

    const folder = DriveApp.getFolderById(tree[path].id);
    const file = folder.createFile(Utilities.newBlob(bytes, body.mimeType || MimeType.MICROSOFT_WORD, finalName));
    file.setDescription('Subido por: ' + teacher + ' · ' + new Date().toLocaleString() + (body.note ? ' · Nota: ' + clean_(body.note).slice(0, 300) : ''));

    log_(tree, [new Date(), teacher, path, finalName, file.getUrl(), clean_(body.note || '')]);
    return json_({ ok: true, name: finalName, url: file.getUrl(), folderUrl: folder.getUrl() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- Utilidades ----------
function listFiles_(path, tree) {
  tree = tree || getTree_();
  if (!tree[path]) return [];
  const out = [];
  const it = DriveApp.getFolderById(tree[path].id).getFiles();
  while (it.hasNext() && out.length < 200) {
    const f = it.next();
    out.push({ name: f.getName(), url: f.getUrl(), updated: f.getLastUpdated().toISOString(), note: f.getDescription() || '' });
  }
  out.sort((a, b) => b.updated.localeCompare(a.updated));
  return out.slice(0, 60);
}

function log_(tree, row) {
  try {
    const folder = DriveApp.getFolderById(tree[LOG_PATH].id);
    const it = folder.getFilesByName(LOG_NAME);
    let ss;
    if (it.hasNext()) ss = SpreadsheetApp.open(it.next());
    else {
      ss = SpreadsheetApp.create(LOG_NAME);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
      ss.getSheets()[0].appendRow(['Fecha', 'Docente', 'Carpeta', 'Archivo', 'Enlace', 'Nota']);
    }
    ss.getSheets()[0].appendRow(row);
  } catch (err) { /* el registro es opcional; la subida no falla por esto */ }
}

function clean_(s) { return String(s).replace(/[\\/:*?"<>|#%]/g, '').trim(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
