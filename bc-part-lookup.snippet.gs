// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  STARK FUTURE — BC Part Lookup (Apps Script snippet)                    ║
// ║  Versión: 1.0.0                                                         ║
// ║                                                                          ║
// ║  INSTRUCCIONES:                                                          ║
// ║  1. Abre el proyecto "Stark Future · Atlas Copco Torque Traceability"   ║
// ║  2. Crea un nuevo archivo "BCPartLookup.gs"                             ║
// ║  3. Pega TODO este contenido en ese archivo                             ║
// ║  4. Actualiza PA_ITEMS_URL con la URL de tu Power Automate flow         ║
// ║     para búsqueda de ítems en BC                                        ║
// ║  5. Actualiza TAGS_SHEET_ID con el ID del Google Sheet donde se         ║
// ║     guardarán los tags de malla (crea uno si no existe)                 ║
// ╚══════════════════════════════════════════════════════════════════════════╝

// ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

// URL del Power Automate flow para buscar ítems en BC.
// El flow debe aceptar el parámetro ?query=<texto> y retornar JSON:
// { items: [{ no, description, unitOfMeasure, inventory, category, bcUrl }] }
//
// Si usas el mismo PA_URL del VIN Lookup, añade lógica al flow para manejar
// el parámetro "searchItem" además de "vin" y "orderNo".
const PA_ITEMS_URL = 'REEMPLAZA_CON_URL_DE_TU_POWER_AUTOMATE_FLOW';

// Google Sheet donde se guardan los tags: mesh_name | bc_item_no | category | tagged_by | timestamp
// Crea un Sheet con esas columnas en la primera fila (encabezados).
const TAGS_SHEET_ID = 'REEMPLAZA_CON_ID_DE_TU_GOOGLE_SHEET';

// Caché de tags en memoria (se resetea al refrescar el script, pero evita
// múltiples lecturas de Sheet en la misma sesión de usuario).
const _tagsCache = {};
let _tagsCacheLoaded = false;

// ── FUNCIÓN PÚBLICA: Buscar ítems en BC ───────────────────────────────────────
/**
 * Busca ítems en Business Central por número o descripción.
 * Llamado desde el frontend: google.script.run.searchBcItems(query)
 *
 * @param {string} query - Texto a buscar (mínimo 2 caracteres)
 * @returns {Array} Lista de ítems: [{ no, description, unitOfMeasure, inventory, category, bcUrl }]
 */
function searchBcItems(query) {
  if (!query || query.trim().length < 2) return [];

  const url = PA_ITEMS_URL + '&searchItem=' + encodeURIComponent(query.trim());

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('searchBcItems PA error: ' + resp.getResponseCode());
      return [];
    }

    const data = JSON.parse(resp.getContentText());
    return Array.isArray(data.items) ? data.items : [];

  } catch (err) {
    Logger.log('searchBcItems error: ' + err);
    return [];
  }
}

// ── FUNCIÓN PÚBLICA: Obtener info de un ítem por número ───────────────────────
/**
 * Retorna los datos de un ítem de BC dado su número de parte.
 * Llamado desde el frontend: google.script.run.getPartByItemNo(itemNo)
 *
 * @param {string} itemNo - Número de parte BC (ej. "SMX1-SF-D-36")
 * @returns {Object|null} { no, description, unitOfMeasure, inventory, category, bcUrl }
 */
function getPartByItemNo(itemNo) {
  if (!itemNo) return null;

  const url = PA_ITEMS_URL + '&itemNo=' + encodeURIComponent(itemNo.trim());

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (resp.getResponseCode() !== 200) return null;

    const data = JSON.parse(resp.getContentText());

    // PA puede retornar directamente el ítem o dentro de data.item
    const item = data.item || (data.no ? data : null);
    return item || null;

  } catch (err) {
    Logger.log('getPartByItemNo error: ' + err);
    return null;
  }
}

// ── FUNCIÓN PÚBLICA: Tagear un mesh con un número de parte ────────────────────
/**
 * Guarda el mapeo mesh_name → bc_item_no en el Google Sheet de tags.
 * Llamado desde el frontend: google.script.run.tagMeshAsPart(meshName, bcItemNo, category)
 *
 * @param {string} meshName  - Nombre del mesh en Three.js (ej. "FF-KYB-220726-16.007")
 * @param {string} bcItemNo  - Número de parte BC (ej. "SMX1-SF-D-36")
 * @param {string} category  - Categoría opcional (ej. "Suspension / Fork")
 */
function tagMeshAsPart(meshName, bcItemNo, category) {
  if (!meshName || !bcItemNo) throw new Error('meshName y bcItemNo son requeridos');

  // Guardar en cache en memoria
  _tagsCache[meshName] = { bcItemNo: bcItemNo, category: category || '' };

  // Guardar en Google Sheet
  try {
    const ss    = SpreadsheetApp.openById(TAGS_SHEET_ID);
    let sheet   = ss.getSheetByName('PartTags');

    if (!sheet) {
      sheet = ss.insertSheet('PartTags');
      sheet.getRange(1, 1, 1, 5).setValues([['MeshName', 'BCItemNo', 'Category', 'TaggedBy', 'Timestamp']]);
      sheet.setFrozenRows(1);
    }

    // Si ya existe un tag para este mesh, actualizarlo
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === meshName) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[bcItemNo, category || '', Session.getEffectiveUser().getEmail(), new Date()]]);
        return;
      }
    }

    // Nueva entrada
    sheet.appendRow([meshName, bcItemNo, category || '', Session.getEffectiveUser().getEmail(), new Date()]);

  } catch (err) {
    Logger.log('tagMeshAsPart Sheet error: ' + err);
    // No lanzar error — el tag en memoria ya se guardó
  }
}

// ── FUNCIÓN AUXILIAR: Cargar tags persistentes al arrancar ────────────────────
/**
 * Carga todos los tags guardados en el Google Sheet.
 * Se puede llamar desde el doGet si quieres pre-cargar los tags al servir la app.
 * No es obligatorio — los tags se cargan perezosamente.
 */
function loadAllTags() {
  if (_tagsCacheLoaded) return _tagsCache;
  try {
    const ss    = SpreadsheetApp.openById(TAGS_SHEET_ID);
    const sheet = ss.getSheetByName('PartTags');
    if (!sheet) return {};
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        _tagsCache[data[i][0]] = { bcItemNo: data[i][1], category: data[i][2] };
      }
    }
    _tagsCacheLoaded = true;
  } catch (err) {
    Logger.log('loadAllTags error: ' + err);
  }
  return _tagsCache;
}

// ── NOTAS DE INTEGRACIÓN CON POWER AUTOMATE ───────────────────────────────────
/*
  El Power Automate flow necesita manejar 3 tipos de requests:

  1. Búsqueda: POST con param ?searchItem=<query>
     → BC API: GET /v2.0/{tenant}/{env}/api/v2.0/items?$filter=contains(no,'{query}') or contains(description,'{query}')&$top=10
     → Retorna: { items: [{ no, description, unitOfMeasure, inventory, category }] }

  2. Lookup por número: POST con param ?itemNo=<no>
     → BC API: GET /v2.0/{tenant}/{env}/api/v2.0/items?$filter=no eq '{itemNo}'
     → Retorna: { item: { no, description, unitOfMeasure, inventory, category, bcUrl } }

  Los mismos parámetros pueden añadirse al PA_URL del VIN Lookup existente,
  añadiendo condiciones Switch en el flow de Power Automate.

  BC Web Client URL para un ítem:
  https://businesscentral.dynamics.com/{tenant}/{environment}?page=30&filter=No. IS {itemNo}
*/
