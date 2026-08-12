/**
 * Automatiškai nustato "-- Select Action --" E2 langelyje ir įrašo šiandienos datą į B3.
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[LIST]");
  if (sheet) {
    sheet.getRange("E2").setValue("-- Select Action --");
    
    // Įrašome šios dienos datą į B3 langelį (jei jis tuščias arba atidarius iš naujo)
    var dateCell = sheet.getRange("B3");
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    dateCell.setValue(today);
  }
}

function onEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();

  if (sheet.getName() === "[LIST]") {
    
    // ==========================================
    // 0. DROPDOWN MENU VALDYMAS (E2 LANGELIS)
    // ==========================================
    if (range.getA1Notation() === "E2") {
      var action = range.getValue().toString().trim();
      
      // PARINKTIS 1: Reset to "To pack"
      if (action.indexOf("Reset") !== -1) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 11) {
          var statusRange = sheet.getRange("C11:C" + lastRow);
          var values = statusRange.getValues();
          
          for (var i = 0; i < values.length; i++) {
            var itemVal = sheet.getRange(11 + i, 2).getValue().toString().trim();
            if (itemVal !== "" && itemVal !== "More...") {
              values[i][0] = "To pack";
            }
          }
          statusRange.setValues(values);
        }
        range.setValue("-- Select Action --");
      }
      
      // PARINKTIS 2: Add row / Add a new row
      else if (action.indexOf("Add") !== -1) {
        insertCustomItemRow(sheet);
        range.setValue("-- Select Action --");
      }
      return;
    }

    // ==========================================
    // 1. B8 LANGELIO APDOROJIMAS (Nuoroda arba Adresas)
    // ==========================================
    if (range.getA1Notation() === "B8") {
      var valB8 = range.getValue().toString().trim();
      
      if (valB8 === "") {
        range.setValue("paste link or address");
        range.setFontColor("#888888");
      } else if (valB8 !== "paste link or address") {
        range.setFontColor("#000000");
        
        if (valB8.includes("maps") || valB8.includes("goo.gl")) {
          var match = valB8.match(/place\/([^\/]+)/);
          if (match && match[1]) {
            var extractedPlace = decodeURIComponent(match[1].replace(/\+/g, ' '));
            range.setValue(extractedPlace);
          }
        }
      }
    }

    // ==========================================
    // 2. D STULPELIO (Notes/Qty) VALDYMAS
    // ==========================================
    if ((col === 4 || col === 5) && row >= 11) {
      var valD = sheet.getRange(row, 4).getValue().toString().trim();
      if (valD === "" || valD === undefined) {
        sheet.getRange(row, 4).setValue("Notes/Qty");
      }
    }

    // ==========================================
    // 3. AUTOMATINIS "To pack" ĮRAŠYMAS C STULPELIJE
    // ==========================================
    if (col === 2 && row >= 11) {
      var itemVal = range.getValue().toString().trim();
      var statusCell = sheet.getRange(row, 3);
      
      if (itemVal !== "" && itemVal !== "More..." && itemVal !== "Notes/Qty") {
        if (statusCell.getValue() === "") {
          statusCell.setValue("To pack");
        }
      } else if (itemVal === "" || itemVal === "More...") {
        statusCell.clearContent();
      }
    }

  }
}

/**
 * Įterpia naują suformatuotą eilutę PO PASIRINKTA EILUTE su visomis formulėmis ir būsena.
 */
function insertCustomItemRow(sheet) {
  var activeCell = sheet.getActiveCell();
  var rowIndex = activeCell ? activeCell.getRow() : sheet.getLastRow();

  // Apsauga: jei pasirinktas langelis yra viršutinėje zonoje (iki 10 eilutės), įterpiame po 11 eilutės arba lentelės gale
  if (rowIndex < 11) {
    rowIndex = sheet.getLastRow() < 11 ? 11 : sheet.getLastRow();
  }

  // 1. Įterpiame naują eilutę po pasirinktos eilutės
  sheet.insertRowAfter(rowIndex);
  var newRow = rowIndex + 1;

  // 2. Kopijuojame formatavimą iš viršutinės eilutės
  var sourceRange = sheet.getRange(rowIndex, 1, 1, 5);
  var targetRange = sheet.getRange(newRow, 1, 1, 5);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // 3. Sujungiame D ir E stulpelius (Merge D:E)
  sheet.getRange(newRow, 4, 1, 2).merge();

  // 4. A stulpelis: Įrašome išmaniosios numeracijos formulę
  var formula = '=IF(OR(ISBLANK(B' + newRow + '); B' + newRow + '="More..."; ISBLANK(C' + newRow + '); C' + newRow + '="Skip"); ""; COUNTIFS($B$11:B' + newRow + '; "<>"; $B$11:B' + newRow + '; "<>More..."; $C$11:C' + newRow + '; "<>"; $C$11:C' + newRow + '; "<>Skip"))';
  sheet.getRange(newRow, 1).setFormula(formula);

  // 5. B stulpelis: Išvalome pavadinimą
  sheet.getRange(newRow, 2).setValue('');

  // 6. C stulpelis: Kopijuojame taisykles ir įrašome pradinę "To pack" būseną
  var sourceStatusCell = sheet.getRange(rowIndex, 3);
  var targetStatusCell = sheet.getRange(newRow, 3);
  
  var rule = sourceStatusCell.getDataValidation();
  if (rule != null) {
    targetStatusCell.setDataValidation(rule);
  }
  
  targetStatusCell.setHorizontalAlignment("center");
  targetStatusCell.setValue('To pack');

  // 7. D:E stulpeliai: Įrašome pilką tekstą "Notes/Qty"
  var notesCell = sheet.getRange(newRow, 4);
  notesCell.setValue('Notes/Qty');
  notesCell.setFontColor('#8c8c8c');

  // 8. Išplečiame sąlyginį formatavimą
  extendConditionalFormatting(sheet, newRow);
}

/**
 * Atnaujina visas sąlyginio formatavimo taisykles lape.
 */
function extendConditionalFormatting(sheet, newRow) {
  var rules = sheet.getConditionalFormatRules();
  var updatedRules = [];

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var ranges = rule.getRanges();
    var newRanges = [];

    for (var j = 0; j < ranges.length; j++) {
      var range = ranges[j];
      if (range.getLastRow() >= newRow - 1) {
        newRanges.push(sheet.getRange(
          range.getRow(),
          range.getColumn(),
          range.getNumRows() + 1,
          range.getNumColumns()
        ));
      } else {
        newRanges.push(range);
      }
    }
    
    updatedRules.push(rule.copy().setRanges(newRanges).build());
  }

  sheet.setConditionalFormatRules(updatedRules);
}
