// ════════════════════════════════════════════════════════════════════
//  STARK FUTURE — 3D Parts Identifier
//  Conecta el visor 3D de la moto con Business Central via Power Automate
//
//  SETUP:
//  1. En Script Properties (Proyecto > Propiedades), añade:
//     PA_ITEMS_URL  → URL de tu Power Automate flow para ítems BC
//     PA_TAGS_SHEET_ID → 1ZxDdid7ETc6hoGJ1IlDwSG1LIHEbTirDSZ4fHtdtD6g
//  2. Despliega como Web App (Ejecutar como: Yo, Acceso: Stark Future)
// ════════════════════════════════════════════════════════════════════

const PROPS = PropertiesService.getScriptProperties();

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Stark Future | 3D Parts Identifier')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── HELPER: Sheet de tags ──────────────────────────────────────────
function _getTagsSheet_() {
  const id = PROPS.getProperty('PA_TAGS_SHEET_ID') || '1ZxDdid7ETc6hoGJ1IlDwSG1LIHEbTirDSZ4fHtdtD6g';
  try {
    const ss = SpreadsheetApp.openById(id);
    let sh = ss.getSheetByName('PartTags');
    if (!sh) {
      sh = ss.insertSheet('PartTags');
      sh.getRange(1,1,1,5).setValues([['MeshName','BCItemNo','Category','TaggedBy','Timestamp']]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#111111').setFontColor('#C9A84C');
    }
    return sh;
  } catch(e) { Logger.log('_getTagsSheet_ error: '+e); return null; }
}

/**
 * Busca ítems en BC por número o descripción.
 * @param {string} query
 * @returns {Array} [{ no, description, unitOfMeasure, inventory, category }]
 */
function searchBcItems(query) {
  if (!query || query.trim().length < 2) return [];
  const url = (PROPS.getProperty('PA_ITEMS_URL') || '') + '&searchItem=' + encodeURIComponent(query.trim());
  try {
    const resp = UrlFetchApp.fetch(url, { method:'POST', muteHttpExceptions:true });
    if (resp.getResponseCode() !== 200) return [];
    const data = JSON.parse(resp.getContentText());
    return Array.isArray(data.items) ? data.items : [];
  } catch(e) { Logger.log('searchBcItems: '+e); return []; }
}

/**
 * Retorna datos completos de un ítem BC por su número de parte.
 * @param {string} itemNo
 * @returns {Object|null}
 */
function getPartByItemNo(itemNo) {
  if (!itemNo) return null;
  const url = (PROPS.getProperty('PA_ITEMS_URL') || '') + '&itemNo=' + encodeURIComponent(itemNo.trim());
  try {
    const resp = UrlFetchApp.fetch(url, { method:'POST', muteHttpExceptions:true });
    if (resp.getResponseCode() !== 200) return null;
    const data = JSON.parse(resp.getContentText());
    return data.item || (data.no ? data : null);
  } catch(e) { Logger.log('getPartByItemNo: '+e); return null; }
}

/**
 * Guarda el mapeo mesh → parte BC en el Sheet de tags.
 * @param {string} meshName
 * @param {string} bcItemNo
 * @param {string} category
 */
function tagMeshAsPart(meshName, bcItemNo, category) {
  if (!meshName || !bcItemNo) return;
  const sh = _getTagsSheet_();
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === meshName) {
      sh.getRange(i+1,2,1,4).setValues([[bcItemNo, category||'', Session.getEffectiveUser().getEmail(), new Date()]]);
      return;
    }
  }
  sh.appendRow([meshName, bcItemNo, category||'', Session.getEffectiveUser().getEmail(), new Date()]);
}
