// --- CONFIGURATION: LOCAL Ref TABS ---
const SHEET_NAME_BRANDS            = "Ref_Brands_list";
const SHEET_NAME_RETAIL_MAP   = "Ref_Category_Map_Retail";
const SHEET_NAME_HEALTH_MAP        = "Ref_Category_Map_Health";
const SHEET_NAME_MASTER_IDS        = "Ref_category_id_list"; 

// --- CONFIGURATION: COLUMN HEADER NAMES ---
const HEADERS = {
  // Required Inputs
  INPUT_TITLE_EN:     "title_en", 
  INPUT_TITLE_GR:     "title_gr", 
  INPUT_CATEGORY:     "category_id",
  INPUT_BARCODE:      "barcode",
  
  // Optional Search Inputs
  INPUT_SEARCH_TERM:  "Αναζήτηση category_id", 
  OUT_SEARCH_RESULT:  "category_id_search",    

  // Outputs (General)
  OUT_BRAND_ID:       "brand_id",
  OUT_CAT_EL:         "category_gr",
  OUT_CAT_EN:         "category_en",
  OUT_PROD_TYPE:      "product_type",
  OUT_VAT:            "vat_rate",
  OUT_TAGS:           "tags",
  OUT_AGE_YESNO:      "age_restricted_true_false",
  OUT_AGE_LIMIT:      "age_minimum",
  OUT_WEIGHT_VAL:     "weightValue",
  OUT_WEIGHT_UNIT:    "weightUnit",
  OUT_NUM_UNITS:      "numberOfUnits",
  OUT_CONT_VAL:       "contentsValue",
  OUT_CONT_UNIT:      "contentsUnit"
};

// List of headers that are allowed to be missing
const OPTIONAL_HEADERS = [
  HEADERS.INPUT_SEARCH_TERM,
  HEADERS.OUT_SEARCH_RESULT
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Menu 1: General Tools
  ui.createMenu('⚡ Autocomplete Tools')
    .addItem('▶️ Run All Processing', 'runFast')
    .addSeparator()
    .addItem('1. Get Contents/Unit Only', 'get_cont')
    .addItem('2. Get Brand ID Only', 'get_brandID')
    .addItem('3. Get Product Titles Only', 'get_productTitle')
    .addItem('4. Get Categories Only', 'get_categories')
    .addItem('5. Get Extended Data (Type/Tags/Age)', 'get_product_type')
    .addItem('6. Get VAT Rate Only', 'get_vat_rate')
    .addItem('7. Validate Master IDs', 'validate_master_ids')
    .addSeparator()
    .addItem('🗑️ Clear Data', 'clearData')
    .addToUi();

 // Menu 2: Link Tools
  ui.createMenu('🔗 Link Tools')
    .addItem('🔗 Convert Barcodes & Titles to Search Links', 'generate_google_search_links')
    .addToUi();

  restoreDropdown();
}

/** --- HELPER: RESTORE A1 DROPDOWN --- */
function restoreDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const cell = sheet.getRange("A1");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Retail', 'Health'], true)
    .setAllowInvalid(false)
    .setHelpText('Please select "Retail" or "Health".')
    .build();
  cell.setDataValidation(rule);
}

/** --- HELPER: GET COL MAP --- */
function getColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h.toString().trim()] = i; });

  const missing = [];
  Object.values(HEADERS).forEach(hName => {
    if (map[hName] === undefined && !OPTIONAL_HEADERS.includes(hName)) { missing.push(hName); }
  });

  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert("❌ Error: Missing Columns!\n\nPlease ensure these headers exist in Row 1:\n\n" + missing.join("\n"));
    return null;
  }
  return map;
}

/** * --- FUNCTION: GENERATE GOOGLE SEARCH LINKS (Barcodes & Titles) --- 
 * 1. Converts Barcodes (Col barcode) to Google Search Links (Handling multiple commas).
 * 2. Converts Titles (Col title_gr) to Google Search Links.
 */
function generate_google_search_links() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast("No data."); return; }
  
  // --- PART 1: PROCESS BARCODES ---
  const barcodeColIdx = map[HEADERS.INPUT_BARCODE] + 1;
  const barcodeRange = sheet.getRange(2, barcodeColIdx, lastRow - 1, 1);
  const barcodeData = barcodeRange.getValues();
  const barcodeRichText = [];

  barcodeData.forEach(row => {
    const rawVal = row[0] ? row[0].toString() : "";
    if (!rawVal) {
      barcodeRichText.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
      return;
    }

    const parts = rawVal.split(",").map(s => s.trim()).filter(s => s !== "");
    if (parts.length === 0) {
       barcodeRichText.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
       return;
    }

    let displayText = parts.join(", ");
    const builder = SpreadsheetApp.newRichTextValue().setText(displayText);
    let currentPos = 0;
    
    parts.forEach((code, index) => {
      // Standard Google Search link
      const url = `https://www.google.com/search?q=${code}`;
      const len = code.length;
      builder.setLinkUrl(currentPos, currentPos + len, url);
      currentPos += len;
      if (index < parts.length - 1) currentPos += 2; // skip ", "
    });

    barcodeRichText.push([builder.build()]);
  });

  // --- PART 2: PROCESS TITLES ---
  const titleColIdx = map[HEADERS.INPUT_TITLE_GR] + 1;
  const titleRange = sheet.getRange(2, titleColIdx, lastRow - 1, 1);
  const titleData = titleRange.getValues();
  const titleRichText = [];

  titleData.forEach(row => {
    const title = row[0] ? row[0].toString().trim() : "";
    
    if (!title) {
      titleRichText.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
      return;
    }

    // Link the whole title to standard Google search
    const url = `https://www.google.com/search?q=${encodeURIComponent(title)}`;
    const builder = SpreadsheetApp.newRichTextValue()
      .setText(title)
      .setLinkUrl(url)
      .build();

    titleRichText.push([builder]);
  });

  // Write Updates
  barcodeRange.setRichTextValues(barcodeRichText);
  titleRange.setRichTextValues(titleRichText);
  
  ss.toast("✅ Converted Barcodes & Titles to Search Links", "Success");
}

/** --- FUNCTION: RUN ALL (FAST) --- */
function runFast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); 
  if (lastRow < 2) { ss.toast("No data to process."); return; }
  
  const shopType = sheet.getRange("A1").getValue();
  let fullMap = null;
  let statusMsg = "Loading data...";
  if (!shopType || shopType.toString().trim() === "") { statusMsg = "Loading data... (No Shop Type Selected!)"; } 
  else { fullMap = getAllCategoryDataMap(shopType); if (!fullMap) return; }

  ss.toast(statusMsg, "Step 1/3");
  const brandMap = getBrandMap(); if (!brandMap) return; 
  
  const masterList = getMasterCategoryList();
  const masterSet = new Set(masterList.map(id => id.toString().trim())); 

  const fullDataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const data = fullDataRange.getValues();

  // --- CONFIG ---
  const ALLOWED_INPUT_ENDINGS = ["τεμ", "g", "kg", "mg", "ml", "l", "m", "cm"];
  const OUTPUT_TRANSLATION = { "τεμ": "pieces" };
  const KEYWORDS_SPECIAL = ["Δισκία", "Ταμπλέτες", "Κάψουλες", "Μεζούρες", "Ρολά", "Φύλλα", "Φακελάκια"];

  const results = {
    brands: [], catsEL: [], catsEN: [], pTypes: [], vats: [], tags: [], ageYs: [], ageLs: [],
    titlesEN: [], titlesENCY: [], titlesELCY: [],
    weights: [], wUnits: [], nUnits: [], cVals: [], cUnits: [],
    colorsBrand: [], colorsCat: [], colorsUnit: [], colorsID: []
  };

  ss.toast("Processing rows...", "Step 2/3");
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const colI = (row[map[HEADERS.INPUT_TITLE_EN]] || "").toString().trim(); 
    const colJ_Raw = (row[map[HEADERS.INPUT_TITLE_GR]] || "").toString().trim(); 
    const colT = (row[map[HEADERS.INPUT_CATEGORY]] || "").toString().trim();

    // 1. CLEAN TITLE
    const colJ = colJ_Raw.replace(/\s+-[0-9.,]+[%€].*$/i, "").trim();

    results.titlesEN.push([colI]); results.titlesENCY.push([colI]); results.titlesELCY.push([colJ]);

    // 2. VALIDATE MASTER ID
    let colorID = null;
    if (colT && !masterSet.has(colT)) { colorID = "#ffc7ce"; }
    results.colorsID.push([colorID]);

    // 3. BRAND LOGIC
    let brandID = "", colorM = null; 
    if (colI) {
      const words = colI.split(/\s+/); const maxWords = Math.min(words.length, 5);
      for (let k = maxWords; k > 0; k--) {
        const phrase = words.slice(0, k).join(" ").toLowerCase();
        if (brandMap.has(phrase)) { brandID = brandMap.get(phrase); break; }
      }
      if (!brandID) colorM = "#ffc7ce"; 
    }
    results.brands.push([brandID]); results.colorsBrand.push([colorM]);

    // 4. CATEGORY LOGIC
    let cEL = "", cEN = "", pT = "", vt = "", tg = "", aY = "", aL = "", colorCD = null; 
    if (colJ || colT) { 
      if (!fullMap) {
        colorCD = "#ffc7ce"; 
      } else {
        const key = colT.toLowerCase();
        if (!key) {
           colorCD = "#ffc7ce"; 
        } else if (fullMap.has(key)) {
           const m = fullMap.get(key); 
           cEL = m.catEL; cEN = m.catEN; pT = m.pType; vt = m.vat; tg = m.tags; aY = m.ageY; aL = m.ageL;
           if (cEL === "" || cEL.toString().toUpperCase() === "CHECK") { colorCD = "#ffc7ce"; }
        } else {
           colorCD = "#ffc7ce"; 
        }
      }
    }
    results.catsEL.push([cEL]); results.catsEN.push([cEN]); results.colorsCat.push([colorCD]);
    results.pTypes.push([pT]); results.vats.push([vt]); results.tags.push([tg]); results.ageYs.push([aY]); results.ageLs.push([aL]);

    // 5. CONTENTS LOGIC
    let qty = 0, unit = "", normWeight = "", normUnit = "";
    let isError = false;

    if (colJ) {
      const words = colJ.split(/\s+/); 
      let lastWord = words[words.length - 1]; 
      let secondLastWord = words.length > 1 ? words[words.length - 2] : "";
      
      const extractNum = (s) => { const m = s.match(/[0-9.,]+/); return m ? parseFloat(m[0].replace(",", ".")) : 0; };

      if (lastWord === "Δώρο") {
        unit = "pieces"; 
        if (secondLastWord.includes("+")) secondLastWord.split("+").forEach(p => qty += extractNum(p)); 
        else qty = colJ.split(" + ").length; 
      } 
      else if (lastWord === "Πακέτο") { qty = 1; unit = "pieces"; } 
      else if (KEYWORDS_SPECIAL.includes(lastWord)) {
        qty = extractNum(secondLastWord);
        if (lastWord === "Δισκία") unit = "tablets";
        else if (lastWord === "Ταμπλέτες") unit = "tablets";
        else if (lastWord === "Κάψουλες") unit = "capsules";
        else if (lastWord === "Μεζούρες") unit = "washes";
        else if (lastWord === "Ρολά") unit = "rolls";
        else if (lastWord === "Φύλλα") unit = "sheets";
        else if (lastWord === "Φακελάκια") unit = "sachets";
      } 
      else if (/[0-9]/.test(lastWord)) {
        qty = extractNum(lastWord);
        if (lastWord.toLowerCase().includes("x")) {
           unit = "pieces"; isError = false;
        } else {
           const m = lastWord.match(/[a-zA-Zα-ωΑ-Ω]+/);
           let rawUnit = m ? m[0] : "";
           if (ALLOWED_INPUT_ENDINGS.includes(rawUnit)) {
              unit = OUTPUT_TRANSLATION[rawUnit] ? OUTPUT_TRANSLATION[rawUnit] : rawUnit;
           } else {
              unit = rawUnit; isError = true;
           }
        }
      } 
      else {
        const prevQty = extractNum(secondLastWord);
        if (prevQty > 0) { qty = prevQty; } else { qty = 1; }
        let rawUnit = lastWord;
        unit = rawUnit; isError = true; 
      }

      if (!isError) {
        if (unit === "g") { normWeight = qty / 1000; normUnit = "kg"; }
        else if (unit === "kg") { normWeight = qty; normUnit = "kg"; }
      }
    }
    
    results.weights.push([isError ? "" : normWeight]); results.wUnits.push([isError ? "" : normUnit]);
    results.nUnits.push([colJ ? 1 : ""]); results.cVals.push([colJ ? qty : ""]); results.cUnits.push([colJ ? unit : ""]);
    results.colorsUnit.push([isError ? "#ffc7ce" : null]);
  }

  ss.toast("Writing results...", "Step 3/3");
  const writeCol = (headerName, values, backgrounds) => {
    const colIndex = map[headerName] + 1; const rng = sheet.getRange(2, colIndex, values.length, 1);
    rng.setValues(values); 
    if (backgrounds) rng.setBackgrounds(backgrounds); 
    else rng.setBackground(null);
  };

  writeCol(HEADERS.OUT_BRAND_ID, results.brands, results.colorsBrand);
  writeCol(HEADERS.OUT_CAT_EL, results.catsEL, results.colorsCat);
  writeCol(HEADERS.OUT_CAT_EN, results.catsEN, results.colorsCat);
  writeCol(HEADERS.OUT_PROD_TYPE, results.pTypes);
  writeCol(HEADERS.OUT_VAT, results.vats);
  writeCol(HEADERS.OUT_TAGS, results.tags);
  writeCol(HEADERS.OUT_AGE_YESNO, results.ageYs);
  writeCol(HEADERS.OUT_AGE_LIMIT, results.ageLs);
  writeCol(HEADERS.OUT_TITLE_EN, results.titlesEN);
  writeCol(HEADERS.OUT_TITLE_EN_CY, results.titlesENCY);
  writeCol(HEADERS.OUT_TITLE_EL_CY, results.titlesELCY);
  
  writeCol(HEADERS.OUT_WEIGHT_VAL, results.weights);
  writeCol(HEADERS.OUT_WEIGHT_UNIT, results.wUnits);
  writeCol(HEADERS.OUT_NUM_UNITS, results.nUnits);
  writeCol(HEADERS.OUT_CONT_VAL, results.cVals);
  writeCol(HEADERS.OUT_CONT_UNIT, results.cUnits, results.colorsUnit);

  const titleColIndex = map[HEADERS.INPUT_TITLE_GR] + 1;
  const titleRange = sheet.getRange(2, titleColIndex, results.colorsUnit.length, 1);
  titleRange.setBackgrounds(results.colorsUnit); 
  
  const idColIndex = map[HEADERS.INPUT_CATEGORY] + 1;
  const idRange = sheet.getRange(2, idColIndex, results.colorsID.length, 1);
  idRange.setBackgrounds(results.colorsID);

  if (!fullMap) ss.toast("⚠️ Completed with Warnings (No Shop Type)", "Finished");
  else ss.toast("Done! 🚀", "Success");
}

/** --- FUNCTION: GET CONTENTS ONLY --- */
function get_cont() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const map = getColMap(sheet);
  if (!map) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  const rW = [], rWU = [], rN = [], rC = [], rCU = [];
  const colors = []; 

  const ALLOWED_INPUT_ENDINGS = ["τεμ", "g", "kg", "mg", "ml", "l", "m", "cm"];
  const OUTPUT_TRANSLATION = { "τεμ": "pieces" };
  const KEYWORDS_SPECIAL = ["Δισκία", "Ταμπλέτες", "Κάψουλες", "Μεζούρες", "Ρολά", "Φύλλα", "Φακελάκια"];

  data.forEach(row => {
    const titleRaw = (row[map[HEADERS.INPUT_TITLE_GR]] || "").toString().trim();
    const title = titleRaw.replace(/\s+-[0-9.,]+[%€].*$/i, "").trim();

    let qty = 0, unit = "", normWeight = "", normUnit = "";
    let isError = false; 
    
    if (!title) {
      rW.push([""]); rWU.push([""]); rN.push([""]); rC.push([""]); rCU.push([""]); colors.push([null]); 
      return;
    }

    const words = title.split(/\s+/);
    let lastWord = words[words.length - 1];
    let secondLastWord = words.length > 1 ? words[words.length - 2] : "";
    
    const extractNum = (s) => { const m = s.match(/[0-9.,]+/); return m ? parseFloat(m[0].replace(",", ".")) : 0; };

    if (lastWord === "Δώρο") {
      unit = "pieces";
      if (secondLastWord.includes("+")) { secondLastWord.split("+").forEach(p => qty += extractNum(p)); } 
      else { qty = title.split(" + ").length; }
    } 
    else if (lastWord === "Πακέτο") { qty = 1; unit = "pieces"; } 
    else if (KEYWORDS_SPECIAL.includes(lastWord)) {
      qty = extractNum(secondLastWord);
      if (lastWord === "Δισκία") unit = "tablets";
      else if (lastWord === "Ταμπλέτες") unit = "tablets";
      else if (lastWord === "Κάψουλες") unit = "capsules";
      else if (lastWord === "Μεζούρες") unit = "washes";
      else if (lastWord === "Ρολά") unit = "rolls";
      else if (lastWord === "Φύλλα") unit = "sheets";
      else if (lastWord === "Φακελάκια") unit = "sachets";
    } 
    else if (/[0-9]/.test(lastWord)) {
      qty = extractNum(lastWord);
      if (lastWord.toLowerCase().includes("x")) {
         unit = "pieces"; isError = false;
      } else {
         const m = lastWord.match(/[a-zA-Zα-ωΑ-Ω]+/);
         let rawUnit = m ? m[0] : ""; 
         if (ALLOWED_INPUT_ENDINGS.includes(rawUnit)) {
            if (OUTPUT_TRANSLATION[rawUnit]) { unit = OUTPUT_TRANSLATION[rawUnit]; } 
            else { unit = rawUnit; }
         } else {
            unit = rawUnit; isError = true;
         }
      }
    } 
    else {
      const prevQty = extractNum(secondLastWord);
      if (prevQty > 0) { qty = prevQty; } else { qty = 1; }
      let rawUnit = lastWord;
      unit = rawUnit; isError = true; 
    }

    if (!isError) {
       if (unit === "g") { normWeight = qty / 1000; normUnit = "kg"; }
       else if (unit === "kg") { normWeight = qty; normUnit = "kg"; }
    }

    rW.push([isError ? "" : normWeight]); 
    rWU.push([isError ? "" : normUnit]); 
    rN.push([1]); 
    rC.push([qty]); 
    rCU.push([unit]); 
    colors.push([isError ? "#ffc7ce" : null]);
  });

  const write = (header, values, backgrounds) => {
    const colIndex = map[header] + 1;
    const rng = sheet.getRange(2, colIndex, values.length, 1);
    rng.setValues(values);
    if (backgrounds) { rng.setBackgrounds(backgrounds); } 
    else { rng.setBackground(null); }
  };

  write(HEADERS.OUT_WEIGHT_VAL, rW);
  write(HEADERS.OUT_WEIGHT_UNIT, rWU);
  write(HEADERS.OUT_NUM_UNITS, rN);
  write(HEADERS.OUT_CONT_VAL, rC);
  write(HEADERS.OUT_CONT_UNIT, rCU, colors);

  const titleColIndex = map[HEADERS.INPUT_TITLE_GR] + 1;
  const titleRange = sheet.getRange(2, titleColIndex, colors.length, 1);
  titleRange.setBackgrounds(colors);
}

/** --- VALIDATION & LOOKUPS --- */

function validate_master_ids() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const masterList = getMasterCategoryList();
  if (!masterList || masterList.length === 0) { ss.toast("List empty!", "Error"); return; }
  const masterSet = new Set(masterList.map(id => id.toString().trim()));
  const dataRange = sheet.getRange(2, map[HEADERS.INPUT_CATEGORY] + 1, lastRow - 1, 1);
  const data = dataRange.getValues();
  const colors = data.map(row => (row[0] && !masterSet.has(row[0].toString().trim())) ? ["#ffc7ce"] : [null]);
  dataRange.setBackgrounds(colors);
  ss.toast("Validation complete.", "Done");
}

function get_categories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const shopType = sheet.getRange("A1").getValue();
  let fullMap = null;
  if (!shopType || shopType.toString().trim() === "") { ss.toast("⚠️ No shop type!", "Warning"); } 
  else { fullMap = getAllCategoryDataMap(shopType); if (!fullMap) return; }
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const resEL = [], resEN = [], colors = [];
  data.forEach(row => {
    const title = (row[map[HEADERS.INPUT_TITLE_GR]] || "").toString().trim(); 
    const key = (row[map[HEADERS.INPUT_CATEGORY]] || "").toString().trim().toLowerCase();
    let el = "", en = "", color = null;
    if (title || key) {
      color = "#ffc7ce"; // Default Red
      if (fullMap && key && fullMap.has(key)) { 
        const m = fullMap.get(key); 
        el = m.catEL; en = m.catEN; 
        if (el !== "" && el.toString().toUpperCase() !== "CHECK") { color = null; }
      }
    }
    resEL.push([el]); resEN.push([en]); colors.push([color]);
  });
  sheet.getRange(2, map[HEADERS.OUT_CAT_EL] + 1, resEL.length).setValues(resEL).setBackgrounds(colors);
  sheet.getRange(2, map[HEADERS.OUT_CAT_EN] + 1, resEN.length).setValues(resEN).setBackgrounds(colors);
}

function get_brandID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const brandMap = getBrandMap(); if (!brandMap) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const results = [], colors = [];
  data.forEach(row => {
    const text = (row[map[HEADERS.INPUT_TITLE_EN]] || "").toString().trim(); 
    let foundID = "", color = null;
    if (text) {
      const words = text.split(/\s+/); const maxWords = Math.min(words.length, 5);
      for (let k = maxWords; k > 0; k--) {
        const phrase = words.slice(0, k).join(" ").toLowerCase();
        if (brandMap.has(phrase)) { foundID = brandMap.get(phrase); break; }
      }
      if (!foundID) color = "#ffc7ce";
    }
    results.push([foundID]); colors.push([color]);
  });
  sheet.getRange(2, map[HEADERS.OUT_BRAND_ID] + 1, results.length).setValues(results).setBackgrounds(colors);
}

function get_productTitle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const rEn = [], rEnCy = [], rElCy = [];
  data.forEach(row => { const colI = row[map[HEADERS.INPUT_TITLE_EN]]; const colJ = row[map[HEADERS.INPUT_TITLE_GR]]; rEn.push([colI]); rEnCy.push([colI]); rElCy.push([colJ]); });
  const write = (h, val) => sheet.getRange(2, map[h] + 1, val.length).setValues(val);
  write(HEADERS.OUT_TITLE_EN, rEn); write(HEADERS.OUT_TITLE_EN_CY, rEnCy); write(HEADERS.OUT_TITLE_EL_CY, rElCy);
}

function get_product_type() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const shopType = sheet.getRange("A1").getValue();
  let fullMap = null;
  if (!shopType || shopType.toString().trim() === "") { ss.toast("⚠️ No shop type!", "Warning"); return; }
  fullMap = getAllCategoryDataMap(shopType); if (!fullMap) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const resPT = [], resTags = [], resAgeY = [], resAgeL = [];
  data.forEach(row => {
    const key = (row[map[HEADERS.INPUT_CATEGORY]] || "").toString().trim().toLowerCase();
    let pType = "", tags = "", ageY = "", ageL = "";
    if (key && fullMap.has(key)) { 
      const m = fullMap.get(key); 
      pType = m.pType; tags = m.tags; ageY = m.ageY; ageL = m.ageL;
    }
    resPT.push([pType]); resTags.push([tags]); resAgeY.push([ageY]); resAgeL.push([ageL]);
  });
  const write = (h, v) => sheet.getRange(2, map[h] + 1, v.length).setValues(v);
  write(HEADERS.OUT_PROD_TYPE, resPT); write(HEADERS.OUT_TAGS, resTags);
  write(HEADERS.OUT_AGE_YESNO, resAgeY); write(HEADERS.OUT_AGE_LIMIT, resAgeL);
  ss.toast("Extended data updated.", "Success");
}

function get_vat_rate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheets()[0];
  const map = getColMap(sheet); if (!map) return;
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return;
  const shopType = sheet.getRange("A1").getValue();
  let fullMap = null;
  if (!shopType || shopType.toString().trim() === "") { ss.toast("⚠️ No shop type!", "Warning"); return; }
  fullMap = getAllCategoryDataMap(shopType); if (!fullMap) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const res = [];
  data.forEach(row => {
    const key = (row[map[HEADERS.INPUT_CATEGORY]] || "").toString().trim().toLowerCase();
    let vat = "";
    if (key && fullMap.has(key)) { const m = fullMap.get(key); vat = m.vat; }
    res.push([vat]);
  });
  sheet.getRange(2, map[HEADERS.OUT_VAT] + 1, res.length, 1).setValues(res);
}

function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getMaxRows();
  sheet.getRange("A1").clearContent();
  restoreDropdown();
  const headersToClear = ["sku","price","category_gr","category_en","comments","barcode","title_gr","title_en","product_type","imageUrls","brand_id","vat_rate","weightValue","weightUnit","numberOfUnits","contentsValue","contentsUnit","category_id","tags","age_restricted_true_false","age_minimum"];
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndices = [];
  headersToClear.forEach(targetHeader => {
    const index = headerValues.findIndex(h => h.toString().trim() === targetHeader);
    if (index !== -1) colIndices.push(index + 1);
  });
  if (lastRow > 1 && colIndices.length > 0) {
    colIndices.forEach(col => { const range = sheet.getRange(2, col, lastRow - 1, 1); range.clearContent(); range.setBackground(null); });
  }
  SpreadsheetApp.getActive().toast("🗑️ All data cleared");
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet(); const map = getColMap(sheet); if (!map) return;
  const editedCol = e.range.getColumn(); const lastEditedCol = e.range.getLastColumn();
  const watchCols = [map[HEADERS.OUT_BRAND_ID] + 1, map[HEADERS.OUT_CAT_EL] + 1, map[HEADERS.OUT_CAT_EN] + 1];
  const isMatch = watchCols.some(c => editedCol <= c && lastEditedCol >= c);
  if (isMatch) e.range.setBackground(null);

  const idColIndex = map[HEADERS.INPUT_CATEGORY] + 1;
  if (editedCol === idColIndex && e.range.getRow() > 1) {
    const val = e.range.getValue().toString().trim();
    if (!val) { e.range.setBackground(null); } 
    else {
       const masterList = getMasterCategoryList();
       const masterSet = new Set(masterList.map(id => id.toString().trim()));
       e.range.setBackground(masterSet.has(val) ? null : "#ffc7ce");
    }
  }
}

function handleSearchEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const map = getColMap(sheet); if (!map) return; 
  if (map[HEADERS.INPUT_SEARCH_TERM] === undefined) return;
  const editedCol = e.range.getColumn();
  const searchColIdx = map[HEADERS.INPUT_SEARCH_TERM] + 1; 
  if (editedCol === searchColIdx && e.range.getRow() > 1) {
    const searchTerm = e.value; 
    const targetCell = sheet.getRange(e.range.getRow(), map[HEADERS.OUT_SEARCH_RESULT] + 1);
    if (!searchTerm) { targetCell.clearContent(); targetCell.clearDataValidations(); return; }
    const masterList = getMasterCategoryList();
    if (!masterList || masterList.length === 0) return;
    const searchParts = searchTerm.toString().toLowerCase().split(" ");
    const matches = masterList.filter(item => {
      if (!item) return false;
      const itemLower = item.toString().toLowerCase();
      return searchParts.every(part => itemLower.includes(part));
    });
    if (matches.length > 0) {
      const safeMatches = matches.slice(0, 500); 
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(safeMatches, true).setAllowInvalid(true).build();
      targetCell.setDataValidation(rule); targetCell.setValue("Select..."); 
    } else { targetCell.clearDataValidations(); targetCell.setValue("No matches found"); }
  }
}

function getMasterCategoryList() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("MASTER_CAT_LIST");
  if (cachedData) return JSON.parse(cachedData);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_MASTER_IDS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return [];
    const data = sheet.getRange(1, 1, lastRow, 1).getValues();
    const cleanList = data.map(r => r[0]).filter(String);
    try { cache.put("MASTER_CAT_LIST", JSON.stringify(cleanList), 21600); } catch (e) {}
    return cleanList;
  } catch (e) { return []; }
}

function getBrandMap() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_BRANDS);
    if (!sheet) { SpreadsheetApp.getUi().alert(`Error: Hidden sheet '${SHEET_NAME_BRANDS}' is missing.`); return null; }
    const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
    const map = new Map();
    data.forEach(r => { if (r[0]) map.set(r[0].toString().trim().toLowerCase(), r[1]); });
    return map;
  } catch (e) { return null; }
}

function getAllCategoryDataMap(shopType) {
  const typeClean = (shopType || "").toString().trim().toLowerCase();
  let targetSheetName = SHEET_NAME_RETAIL_MAP;
  if (typeClean === "pharmacy") targetSheetName = SHEET_NAME_HEALTH_MAP;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) { SpreadsheetApp.getUi().alert(`Error: Hidden sheet '${targetSheetName}' is missing.`); return null; }
    const lastCol = sheet.getLastColumn();
    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {}; headerValues.forEach((h, i) => { colMap[h.toString().trim()] = i; });
    const idxEL = 1; const idxEN = 2; 
    const idxPType = colMap[HEADERS.OUT_PROD_TYPE]; const idxVat = colMap[HEADERS.OUT_VAT];
    const idxTags = colMap[HEADERS.OUT_TAGS]; const idxAgeYes = colMap[HEADERS.OUT_AGE_YESNO]; const idxAgeLim = colMap[HEADERS.OUT_AGE_LIMIT];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    const map = new Map();
    data.forEach(r => {
      const key = r[0].toString().trim().toLowerCase();
      if (key) {
        map.set(key, {
          catEL: r[idxEL], catEN: r[idxEN],
          pType: (idxPType !== undefined) ? r[idxPType] : "",
          vat:   (idxVat !== undefined)   ? r[idxVat] : "",
          tags:  (idxTags !== undefined)  ? r[idxTags] : "",
          ageY:  (idxAgeYes !== undefined)? r[idxAgeYes] : "",
          ageL:  (idxAgeLim !== undefined)? r[idxAgeLim] : ""
        });
      }
    });
    return map;
  } catch (e) { console.error(e); return null; }
}

function getCategoryMap(shopType) {
  const full = getAllCategoryDataMap(shopType);
  if (!full) return null;
  const simpleMap = new Map();
  full.forEach((val, key) => { simpleMap.set(key, [val.catEL, val.catEN]); });
  return simpleMap;
}
