/**
 * Project: Niagara Association Of Photographic Art Competition Score Tracker
 * File: Code.gs
 * File: SettingsDialog.html
 * File: EmailMemberDialog.html
 * 
 * Author: Chris Empey (cempey@gmail.com) (cempey@NiagaraPhoto.org)
 * Copyright (c) 2026 Jane Doe
 * All Rights Reserved
 * 
 * History:
 * 2025-08-12: Initial creation (Chris Empey)
 * 2026-05-03: Conversion to AppsScript (Chris Empey)
 * 2025-08-17: Final Working Vesion 1 (Chris Empey)
 */



/**
 * // ============================================================
// CONFIGURATION
// ============================================================
*/
const CONFIG = {
  divMap: { B: 'Bronze', S: 'Silver', G: 'Gold', D: 'Diamond', P: 'Platinum' },
  catMap: { A: 'Action', C: 'Creative', J: 'Photo-Journalism', N: 'Nature', P: 'Pictorial', T: 'Portrait' },
  compMap: { C: 'Colour', M: 'Monochrome' },
  imgNumToMonth: {
    1: 'September', 2: 'September',
    3: 'October',   4: 'October',
    5: 'November',  6: 'November',
    7: 'December',  8: 'December',
    9: 'January',  10: 'January',
    11: 'February', 12: 'February'
  },
  monthOrder: ['September', 'October', 'November', 'December', 'January', 'February'],
  levelOrder: ['Bronze', 'Silver', 'Gold', 'Diamond', 'Platinum'],
  awardNames: ['Bronze', 'Silver', 'Gold', 'Diamond', 'Platinum'],
  levelColors: {
    Bronze:   '#CD7F32',
    Silver:   '#C0C0C0',
    Gold:     '#FFD700',
    Diamond:  '#B9F2FF',
    Platinum: '#E8D5FF'
  },
  minHonourScore: { Bronze: 10, Silver: 11, Gold: 12, Diamond: 13, Platinum: 14 }
};


/** 
// ============================================================
// COMPETITION SETTINGS: Comp (C/M), Format (Print/Digital), Year
// ============================================================
*/

function showSetupDialog() {
  openSettingsDialog('setup');
}

function showChangeSettingsDialog() {
  openSettingsDialog('change');
}

function openSettingsDialog(mode) {
  const existing = getCompetitionSettings() || {};
  const template = HtmlService.createTemplateFromFile('SettingsDialog');
  template.mode = mode;
  template.currentComp = existing.compType || 'C';
  template.currentFormat = existing.format || 'Print';
  template.currentYear = existing.startYear || '';
  const html = template.evaluate().setWidth(420).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    mode === 'setup' ? 'Competition Setup' : 'Change Competition Settings'
  );
}

// Called from SettingsDialog.html via google.script.run
function processSettingsForm(compType, format, startYearStr, mode) {
  compType = String(compType).trim().toUpperCase();
  format = String(format).trim();

  if (compType !== 'C' && compType !== 'M') throw new Error('Please select Colour or Monochrome.');
  if (format !== 'Print' && format !== 'Digital') throw new Error('Please select Print or Digital.');
  if (!/^\d{4}$/.test(String(startYearStr).trim())) throw new Error('Starting year must be exactly 4 digits.');

  const settings = {
    compType: compType,
    format: format,
    startYear: parseInt(startYearStr, 10)
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSettingsSheet(ss, settings);

  if (mode === 'setup') {
    createMembersSheet(ss);
    createHonourLevelsSheet(ss);
    createImportSheet(ss);
    createEntriesSheet(ss);
    applyCsvImportValidation();
    deleteDefaultSheet1(ss);
    Logger.log('Workbook setup complete: ' + buildCompetitionTitle(settings));
  } else {
    // Change mode: don't touch Members/Entries/CSV Import data.
    // Just refresh the Summary title if entries already exist.
    const membersSheet = ss.getSheetByName('Members');
    const entriesSheet = ss.getSheetByName('Entries');
    if (membersSheet && entriesSheet && entriesSheet.getLastRow() > 1) {
      const memberData = membersSheet.getDataRange().getValues();
      const memberLevelMap = {};
      for (let i = 1; i < memberData.length; i++) {
        const nameRaw = String(memberData[i][2]).trim();
        const level = memberData[i][3];
        if (nameRaw) memberLevelMap[nameRaw.toLowerCase()] = { fullName: nameRaw, level: level };
      }
      const allEntries = loadAllEntries(ss, memberLevelMap);
      if (allEntries.length > 0) {
        buildSummarySheet(ss, allEntries, memberLevelMap);
      }
    }
    Logger.log('Settings updated: ' + buildCompetitionTitle(settings));
  }
}

function createSettingsSheet(ss, settings) {
  PropertiesService.getDocumentProperties().setProperty('CompetitionSettings', JSON.stringify(settings));

  let sheet = ss.getSheetByName('Settings');
  if (!sheet) sheet = ss.insertSheet('Settings');
  sheet.clearContents();

  const compLabel = CONFIG.compMap[settings.compType] || settings.compType;
  const yearLabel = settings.startYear + '-' + (parseInt(settings.startYear, 10) + 1);
  const title = buildCompetitionTitle(settings);

  sheet.getRange('A1').setValue('Competition Settings').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setValue('Competition Title:').setFontWeight('bold');
  sheet.getRange('B3').setValue(title);
  sheet.getRange('A4').setValue('Colour/Monochrome:').setFontWeight('bold');
  sheet.getRange('B4').setValue(compLabel + ' (' + settings.compType + ')');
  sheet.getRange('A5').setValue('Format:').setFontWeight('bold');
  sheet.getRange('B5').setValue(settings.format);
  sheet.getRange('A6').setValue('Competition Years:').setFontWeight('bold');
  sheet.getRange('B6').setValue(yearLabel);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 320);
}

function buildCompetitionTitle(settings) {
  const compLabel = CONFIG.compMap[settings.compType] || settings.compType;
  const yearLabel = settings.startYear + '-' + (parseInt(settings.startYear, 10) + 1);
  return 'NAPA ' + compLabel + ' ' + settings.format + ' Competition ' + yearLabel;
}

function getCompetitionSettings() {
  const raw = PropertiesService.getDocumentProperties().getProperty('CompetitionSettings');
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { /* fall through to sheet fallback */ }
  }

  // Fallback: parse from Settings sheet (e.g., if the spreadsheet was copied to a new file)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return null;

  const compCell   = String(sheet.getRange('B4').getValue() || '');
  const formatCell = String(sheet.getRange('B5').getValue() || '').trim();
  const yearCell   = String(sheet.getRange('B6').getValue() || '');

  const compMatch = compCell.match(/\(([CM])\)/);
  const yearMatch = yearCell.match(/(\d{4})/);

  if (!compMatch || !formatCell || !yearMatch) return null;

  const settings = {
    compType: compMatch[1],
    format: formatCell,
    startYear: parseInt(yearMatch[1], 10)
  };
  PropertiesService.getDocumentProperties().setProperty('CompetitionSettings', JSON.stringify(settings));
  return settings;
}

function getCompetitionType() {
  const settings = getCompetitionSettings();
  return settings ? settings.compType : null;
}

function getCompetitionTitle() {
  const settings = getCompetitionSettings();
  return settings ? buildCompetitionTitle(settings) : 'Competition Summary';
}

/**
// ============================================================
// HELPER: Remove the default "Sheet1" if it's empty
// ============================================================
*/
function deleteDefaultSheet1(ss) {
  const sheet = ss.getSheetByName('Sheet1');
  if (!sheet) return;

  const isEmpty = sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;
  if (isEmpty) {
    // A spreadsheet must always have at least one sheet, so only delete
    // if there's at least one other sheet left after removing this one.
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
      Logger.log('Deleted empty default "Sheet1".');
    } else {
      Logger.log('"Sheet1" is empty but is the only sheet — not deleted.');
    }
  } else {
    Logger.log('"Sheet1" was not deleted because it contains data.');
  }
}

/**
 * // ============================================================
// MAINTENANCE: Members sheet
// ============================================================
*/
function createMembersSheet(ss) {
  let sheet = ss.getSheetByName('Members');
  if (!sheet) sheet = ss.insertSheet('Members');

  if (sheet.getLastRow() === 0) {
    const headers = ['First Name', 'Last Name', 'Full Name', 'Skill Level', 'Email'];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4A90D9')
      .setFontColor('#FFFFFF');
    [120, 120, 160, 110, 220].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
  }

  // Clear all validation before reapplying
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  const levelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.levelOrder)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('D2:D200').setDataValidation(levelRule);
}

function resetMembersValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Members');
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  const levelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.levelOrder)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('D2:D200').setDataValidation(levelRule);
  Logger.log('Members validation reset. Valid levels: ' + CONFIG.levelOrder.join(', '));
}


/**
// ============================================================
// MAINTENANCE: Honour Levels sheet
// ============================================================
*/
function createHonourLevelsSheet(ss) {
  let sheet = ss.getSheetByName('Honour Levels');
  if (!sheet) sheet = ss.insertSheet('Honour Levels');
  sheet.clearContents();

  const headers = ['Skill Level', 'Minimum Score for Honour Award', 'Honour Awards Available'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4A90D9')
    .setFontColor('#FFFFFF');

  const data = [
    ['Bronze',   10, 'Bronze, Silver, Gold, Diamond, Platinum'],
    ['Silver',   11, 'Silver, Gold, Diamond, Platinum'],
    ['Gold',     12, 'Gold, Diamond, Platinum'],
    ['Diamond',  13, 'Diamond, Platinum'],
    ['Platinum', 14, 'Platinum'],
  ];
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  [120, 220, 280].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);
}

/**
// ============================================================
// IMPORT: Sheet where CSV is pasted
//  Column order: Comp, Cat, CompNum, Div, ImgNum, Maker, Title, Score
// ============================================================
*/
function createImportSheet(ss) {
  let sheet = ss.getSheetByName('CSV Import');
  if (!sheet) sheet = ss.insertSheet('CSV Import');
  sheet.clearContents();

  const compType = getCompetitionType();
  const compNote = compType
    ? ' This workbook is set up for ' + CONFIG.compMap[compType] + ' (' + compType + ') — Comp column must match.'
    : '';

  // Instruction row
  sheet.getRange('A1')
    .setValue('Paste or enter tab-delimited data starting in row 2. Columns A–H must match the headers below, then run "Import CSV Data".' + compNote)
    .setFontWeight('bold')
    .setFontColor('#CC0000');

  // Header row for import columns
  const headers = ['Comp', 'Cat', 'CompNum', 'Div', 'ImgNum', 'Maker', 'Title', 'Score'];
  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#EEEEEE');
  sheet.setColumnWidth(1, 70);   // Comp
  sheet.setColumnWidth(2, 120);  // Cat
  sheet.setColumnWidth(3, 80);   // CompNum
  sheet.setColumnWidth(4, 80);   // Div
  sheet.setColumnWidth(5, 80);   // ImgNum
  sheet.setColumnWidth(6, 220);  // Maker
  sheet.setColumnWidth(7, 260);  // Title
  sheet.setColumnWidth(8, 70);   // Score
  // Freeze instruction row + header row
  sheet.setFrozenRows(2);

  // Add helpful validation / lookups
  applyCsvImportValidation();
}


/**
// ============================================================
// ENTRIES: Permanent store for all imported data
//  Column order: Comp, Category, CompNum, Division, ImgNum, Maker, Title, Score, Month
// ============================================================
*/
function createEntriesSheet(ss) {
  let sheet = ss.getSheetByName('Entries');
  if (!sheet) sheet = ss.insertSheet('Entries');
  sheet.clearContents();

  const headers = ['Comp', 'Category', 'CompNum', 'Division', 'ImgNum', 'Maker', 'Title', 'Score', 'Month'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4A90D9')
    .setFontColor('#FFFFFF');

  [70, 120, 80, 80, 70, 180, 260, 70, 100].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);
}


/**
// ============================================================
// IMPORT HELPER: Add new members (case-insensitive)
// ============================================================
*/
function updateMembersFromEntries(ss, newMembers, memberLevelMap) {
  const membersSheet = ss.getSheetByName('Members');
  if (!membersSheet) return;

  const newMemberRows = Object.values(newMembers);
  if (newMemberRows.length === 0) return;

  // Build lowercase set of existing names
  const currentData = membersSheet.getDataRange().getValues();
  const currentNames = new Set();
  for (let i = 1; i < currentData.length; i++) {
    const nameRaw = String(currentData[i][2]).trim();
    if (nameRaw) currentNames.add(nameRaw.toLowerCase());
  }

  const toInsert = newMemberRows.filter(m => !currentNames.has(m.fullName.toLowerCase()));
  if (toInsert.length === 0) return;

  const lastRow = membersSheet.getLastRow();
  const insertData = toInsert.map(m => [m.firstName, m.lastName, m.fullName, m.level, '']);
  membersSheet.getRange(lastRow + 1, 1, insertData.length, 5).setValues(insertData);

  // Update memberLevelMap with lowercase keys
  for (const m of toInsert) {
    const key = m.fullName.toLowerCase();
    memberLevelMap[key] = { fullName: m.fullName, level: m.level };
  }

  Logger.log(toInsert.length + ' new member(s) added: ' + toInsert.map(m => m.fullName).join(', '));
}

/**
// ============================================================
// OPTIONAL: Apply helpful validation to CSV Import
//  Column order: Comp(1), Cat(2), CompNum(3), Div(4), ImgNum(5), Maker(6), Title(7), Score(8)
// ============================================================
*/
function applyCsvImportValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('CSV Import');
  const membersSheet = ss.getSheetByName('Members');
  if (!importSheet || !membersSheet) return;
  const lastRow = importSheet.getMaxRows(); // apply broadly

  // Comp validation (A column)
  const compKeys = Object.keys(CONFIG.compMap); // ['C','M']
  const compRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(compKeys)
    .setAllowInvalid(true)
    .build();
  importSheet.getRange(3, 1, lastRow - 2, 1).setDataValidation(compRule);

  // Category validation (B column)
  const catKeys = Object.keys(CONFIG.catMap); // e.g., ['A','C','J','N','P','T']
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(catKeys)
    .setAllowInvalid(true)  // warn but don't block
    .build();
  importSheet.getRange(3, 2, lastRow - 2, 1).setDataValidation(catRule);

  // Division validation (D column)
  const divKeys = Object.keys(CONFIG.divMap); // e.g., ['B','S','G','D','P']
  const divRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(divKeys)
    .setAllowInvalid(true)
    .build();
  importSheet.getRange(3, 4, lastRow - 2, 1).setDataValidation(divRule);

  // Maker suggestions: Full Name from Members, but allow others (F column)
  const lastMemberRow = membersSheet.getLastRow();
  if (lastMemberRow > 1) {
    const fullNameRange = membersSheet.getRange(2, 3, lastMemberRow - 1, 1); // Members!C2:C
    const makerRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(fullNameRange)
      .setAllowInvalid(true)  // allow names not yet in Members
      .build();
    importSheet.getRange(3, 6, lastRow - 2, 1).setDataValidation(makerRule);
  }
}


/**
// ============================================================
// IMPORT: Parse CSV, confirm new members, store in Entries
//  Import row order: Comp, Cat, CompNum, Div, ImgNum, Maker, Title, Score
// ============================================================
*/
function importCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('CSV Import');
  const membersSheet = ss.getSheetByName('Members');
  let entriesSheet = ss.getSheetByName('Entries');
  const ui = SpreadsheetApp.getUi();

  if (!importSheet || !membersSheet) {
    ui.alert('Required sheets not found. Please run Setup first.');
    return;
  }
  if (!entriesSheet) {
    createEntriesSheet(ss);
    entriesSheet = ss.getSheetByName('Entries');
  }

  const compType = getCompetitionType();
  if (!compType) {
    ui.alert('This workbook has no Competition Type set. Please run Setup Workbook, or use "Change Competition Settings" from the menu, before importing.');
    return;
  }

  // Use the actual starting row of the data range so reported row numbers
  // always match the real sheet row (protects against row 1 being blank, etc.)
  const importRange = importSheet.getDataRange();
  const importStartRow = importRange.getRow();
  const raw = importRange.getValues();

  // ------------------------------------------------------------
  // Build member map with lowercase keys (needed for skill-level
  // mismatch checking below, so this now happens before validation)
  // ------------------------------------------------------------
  const memberData = membersSheet.getDataRange().getValues();
  const memberLevelMap = {};          // key -> { fullName, level }
  const existingMembers = new Set();  // key = name.toLowerCase()

  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    if (nameRaw) {
      const key = nameRaw.toLowerCase();
      memberLevelMap[key] = { fullName: nameRaw, level: level };
      existingMembers.add(key);
    }
  }

  // Build a lookup of (maker + ImgNum) pairs already stored in Entries,
  // so we can catch duplicate entry numbers per member.
  const existingEntryKeys = new Set();
  const entriesLastRow = entriesSheet.getLastRow();
  if (entriesLastRow > 1) {
    const entriesData = entriesSheet.getRange(2, 1, entriesLastRow - 1, entriesSheet.getLastColumn()).getValues();
    for (const row of entriesData) {
      const existingMaker = String(row[5]).trim().toLowerCase();  // Maker is Entries column F (index 5)
      const existingImgNum = parseInt(row[4]);                    // ImgNum is Entries column E (index 4)
      if (existingMaker && !isNaN(existingImgNum)) {
        existingEntryKeys.add(existingMaker + '||' + existingImgNum);
      }
    }
  }

  // ------------------------------------------------------------
  // PRE-VALIDATION PASS: Category, Division, ImgNum range,
  // duplicate entry numbers, and skill-level mismatches for
  // existing members must all be valid on every data row.
  // If anything fails, abort before writing anything.
  // ------------------------------------------------------------
  const catErrors = [];
  const divErrors = [];
  const imgNumRangeErrors = [];
  const duplicateErrors = [];
  const skillMismatchErrors = [];
  const seenInThisImport = new Set(); // maker||imgNum seen so far in this CSV paste

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row[0]) continue; // skip blank rows

    const sheetRow = importStartRow + i; // actual physical row on the sheet

    const firstCell = String(row[0]).trim().toUpperCase();
    if (firstCell === 'COMP') continue; // header row

    const cat    = String(row[1]).trim().toUpperCase();
    const div    = String(row[3]).trim().toUpperCase();
    const imgNum = parseInt(row[4]);
    const maker  = String(row[5]).trim();
    const makerKey = maker.toLowerCase();

    if (!CONFIG.catMap[cat]) {
      catErrors.push('Row ' + sheetRow + ': "' + row[1] + '" (' + (maker || 'no maker') + ')');
    }
    if (!CONFIG.divMap[div]) {
      divErrors.push('Row ' + sheetRow + ': "' + row[3] + '" (' + (maker || 'no maker') + ')');
    }
    if (isNaN(imgNum) || imgNum < 1 || imgNum > 12 || row[4] === '' || row[4] === null) {
      imgNumRangeErrors.push('Row ' + sheetRow + ': "' + row[4] + '" (' + (maker || 'no maker') + ')');
    } else if (maker) {
      const dupKey = makerKey + '||' + imgNum;
      if (existingEntryKeys.has(dupKey)) {
        duplicateErrors.push('Row ' + sheetRow + ': ' + maker + ' already has an entry #' + imgNum + ' in the Entries sheet');
      } else if (seenInThisImport.has(dupKey)) {
        duplicateErrors.push('Row ' + sheetRow + ': ' + maker + ' has entry #' + imgNum + ' listed more than once in this import');
      } else {
        seenInThisImport.add(dupKey);
      }
    }

    // Skill-level mismatch check: only applies to members who already exist.
    // New members will simply be added at the level shown in the CSV, so
    // there's nothing to compare them against yet.
    if (CONFIG.divMap[div] && maker && existingMembers.has(makerKey)) {
      const csvLevel = CONFIG.divMap[div];
      const existingLevel = memberLevelMap[makerKey].level;
      if (csvLevel !== existingLevel) {
        skillMismatchErrors.push(
          'Row ' + sheetRow + ': ' + maker + ' — CSV shows "' + csvLevel +
          '" but Members sheet has "' + existingLevel + '"'
        );
      }
    }
  }

  if (catErrors.length > 0 || divErrors.length > 0 || imgNumRangeErrors.length > 0 ||
      duplicateErrors.length > 0 || skillMismatchErrors.length > 0) {
    let msg = 'Import stopped — no data was written to Entries.\n\n';

    if (catErrors.length > 0) {
      msg += 'Invalid Category value(s) — must be one of A, C, J, N, P, T:\n';
      msg += catErrors.slice(0, 25).join('\n') + '\n';
      if (catErrors.length > 25) msg += '...and ' + (catErrors.length - 25) + ' more.\n';
      msg += '\n';
    }

    if (divErrors.length > 0) {
      msg += 'Invalid Division value(s) — must be one of B, S, G, D, P:\n';
      msg += divErrors.slice(0, 25).join('\n') + '\n';
      if (divErrors.length > 25) msg += '...and ' + (divErrors.length - 25) + ' more.\n';
      msg += '\n';
    }

    if (imgNumRangeErrors.length > 0) {
      msg += 'Invalid Entry Number value(s) — must be a whole number from 1 to 12:\n';
      msg += imgNumRangeErrors.slice(0, 25).join('\n') + '\n';
      if (imgNumRangeErrors.length > 25) msg += '...and ' + (imgNumRangeErrors.length - 25) + ' more.\n';
      msg += '\n';
    }

    if (duplicateErrors.length > 0) {
      msg += 'Duplicate Entry Number(s):\n';
      msg += duplicateErrors.slice(0, 25).join('\n') + '\n';
      if (duplicateErrors.length > 25) msg += '...and ' + (duplicateErrors.length - 25) + ' more.\n';
      msg += '\n';
    }

    if (skillMismatchErrors.length > 0) {
      msg += 'Skill Level Mismatch(es) — Division in CSV does not match Members sheet:\n';
      msg += skillMismatchErrors.slice(0, 25).join('\n') + '\n';
      if (skillMismatchErrors.length > 25) msg += '...and ' + (skillMismatchErrors.length - 25) + ' more.\n';
    }

    msg += '\nPlease correct these rows in CSV Import (or update the Members sheet) and run Import CSV Data again.';
    ui.alert('Invalid Data Found', msg, ui.ButtonSet.OK);
    return;
  }

  // ------------------------------------------------------------
  // Main parse pass
  // ------------------------------------------------------------
  const parsedRows = [];
  const entries = [];
  const newMembers = {};
  const mismatchedRows = [];

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row[0]) continue;

    const firstCell = String(row[0]).trim().toUpperCase();
    if (firstCell === 'COMP') continue;

    const rowComp = String(row[0]).trim().toUpperCase();
    const cat     = String(row[1]).trim().toUpperCase();
    const compNum = row[2];
    const div     = String(row[3]).trim().toUpperCase();
    const imgNum  = parseInt(row[4]);
    const maker   = String(row[5]).trim();
    const title   = String(row[6]).trim();
    const score   = parseInt(row[7]);

    // Cat, Div, ImgNum range/duplicates, and skill mismatches are already
    // guaranteed valid by the pre-validation pass above.
    if (!CONFIG.compMap[rowComp]) { Logger.log('Row ' + (importStartRow + i) + ': Unknown Comp value "' + rowComp + '" for ' + maker + ' — skipped'); continue; }
    if (isNaN(score) || score < 9 || score > 15) { Logger.log('Row ' + (importStartRow + i) + ': Invalid score "' + score + '" — skipped'); continue; }

    if (rowComp !== compType) {
      mismatchedRows.push({ row: importStartRow + i, maker: maker, comp: rowComp });
      continue;
    }

    const month = CONFIG.imgNumToMonth[imgNum];
    const level = CONFIG.divMap[div];
    const nameKey = maker.toLowerCase();

    if (!existingMembers.has(nameKey) && !newMembers[nameKey]) {
      const parts = maker.split(' ');
      const firstName = parts[0];
      const lastName  = parts.slice(1).join(' ');
      newMembers[nameKey] = { firstName, lastName, fullName: maker, level };
    }

    parsedRows.push([rowComp, cat, compNum, div, imgNum, maker, title, score, month]);

    entries.push({
      nameKey,
      category: CONFIG.catMap[cat],
      compNum,
      level,
      month,
      maker,
      title,
      score,
      comp: rowComp,
      bonus:  0,
      honour: ''
    });
  }

  if (mismatchedRows.length > 0) {
    let msg = 'This workbook is set up for the ' + CONFIG.compMap[compType] + ' competition (' + compType + ').\n\n';
    msg += mismatchedRows.length + ' row(s) had a different Comp value and were skipped:\n\n';
    mismatchedRows.slice(0, 20).forEach(r => {
      msg += 'Row ' + r.row + ': ' + r.maker + ' (Comp = "' + r.comp + '")\n';
    });
    if (mismatchedRows.length > 20) msg += '\n...and ' + (mismatchedRows.length - 20) + ' more.';
    ui.alert('Competition Type Mismatch', msg, ui.ButtonSet.OK);
  }

  if (entries.length === 0) {
    ui.alert('No valid entries found to import. Check the CSV format (including the Comp column) and try again.');
    return;
  }

  const newMemberKeys = Object.keys(newMembers);

  // Confirm new members BEFORE writing anything — listed alphabetically,
  // name and level separated by a dash.
  if (newMemberKeys.length > 0) {
    const sortedNewMembers = newMemberKeys
      .map(key => newMembers[key])
      .sort((a, b) => a.fullName.localeCompare(b.fullName));

    let preview = 'The following new members will be added:\n\n';
    sortedNewMembers.forEach(m => {
      preview += m.fullName + ' - ' + m.level + '\n';
    });
    preview += '\nClick OK to proceed with the import, or Cancel to stop so you can correct names/levels first.';

    const response = ui.alert('Confirm New Members', preview, ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) {
      ui.alert('Import cancelled. No data was written to Entries and no members were added.');
      return;
    }
  }

  const lastEntryRow = entriesSheet.getLastRow();
  entriesSheet.getRange(lastEntryRow + 1, 1, parsedRows.length, parsedRows[0].length)
    .setValues(parsedRows);

  const lastImportRow = importSheet.getLastRow();
  if (lastImportRow > 1) {
    importSheet.getRange(2, 1, lastImportRow - 1, importSheet.getLastColumn()).clearContent();
  }

  updateMembersFromEntries(ss, newMembers, memberLevelMap);

  const allEntries = loadAllEntries(ss, memberLevelMap);

  buildMemberSheets(ss, allEntries, memberLevelMap);
  buildSummarySheet(ss, allEntries, memberLevelMap);

  const msg = 'Import complete!\n' +
    entries.length + ' entries added.\n' +
    (mismatchedRows.length > 0 ? mismatchedRows.length + ' row(s) skipped due to Comp mismatch.\n' : '') +
    (newMemberKeys.length > 0
      ? newMemberKeys.length + ' new member(s) added to Members sheet.'
      : 'No new members.');
  ui.alert(msg);
}

/**
// ============================================================
// LOAD: Read and process all entries from Entries sheet
//  Entries row order: Comp, Category, CompNum, Division, ImgNum, Maker, Title, Score, Month
// ============================================================
*/
function loadAllEntries(ss, memberLevelMap) {
  const entriesSheet = ss.getSheetByName('Entries');
  if (!entriesSheet) return [];

  const raw = entriesSheet.getDataRange().getValues();
  const entries = [];

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row[1]) continue; // Category is the reliable "row has data" check now

    // Comp column (index 0) — tolerate older rows saved before this column existed
    const comp   = row[0] ? String(row[0]).trim().toUpperCase() : '';
    const cat    = String(row[1]).trim().toUpperCase();
    const div    = String(row[3]).trim().toUpperCase();
    const imgNum = parseInt(row[4]);
    const maker  = String(row[5]).trim();
    const title  = String(row[6]).trim();
    const score  = parseInt(row[7]);

    if (!CONFIG.catMap[cat] || !CONFIG.divMap[div] || !CONFIG.imgNumToMonth[imgNum]) continue;
    if (isNaN(score) || score < 9 || score > 15) continue;

    const nameKey = maker.toLowerCase();

    entries.push({
      nameKey,
      category: CONFIG.catMap[cat],
      level:    CONFIG.divMap[div],
      month:    CONFIG.imgNumToMonth[imgNum],
      maker,
      title,
      score,
      comp,
      bonus:  0,
      honour: ''
    });
  }

  // Bonus logic (per member+month, case-insensitive)
  const grouped = {};
  for (const e of entries) {
    const key = e.nameKey + '||' + e.month;
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(e);
  }
  for (const group of Object.values(grouped)) {
    const distinctCats = new Set(group.map(e => e.category));
    if (distinctCats.size > 1) group[0].bonus = 1;
  }

  // Honour awards using memberLevelMap (case-insensitive)
  for (const e of entries) {
    const mInfo = memberLevelMap[e.nameKey] || { fullName: e.maker, level: e.level };
    const memberLevel = mInfo.level;
    const minScore = CONFIG.minHonourScore[memberLevel] || 99;

    if (e.score < minScore || e.score < 10) { e.honour = ''; continue; }
    if (e.score >= 14)       e.honour = 'Platinum';
    else if (e.score === 13) e.honour = 'Diamond';
    else if (e.score === 12) e.honour = 'Gold';
    else if (e.score === 11) e.honour = 'Silver';
    else if (e.score === 10) e.honour = 'Bronze';
    else                     e.honour = '';
  }

  return entries;
}

/**
// ============================================================
// REPROCESS: After manual edits to Entries
// ============================================================
*/
function reprocessEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  const entriesSheet = ss.getSheetByName('Entries');
  const ui = SpreadsheetApp.getUi();

  if (!membersSheet || !entriesSheet) {
    ui.alert('Members or Entries sheet not found. Run Setup and Import CSV first.');
    return;
  }

  // Rebuild memberLevelMap (lowercase keys)
  const memberData = membersSheet.getDataRange().getValues();
  const memberLevelMap = {};
  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    if (nameRaw) memberLevelMap[nameRaw.toLowerCase()] = { fullName: nameRaw, level: level };
  }

  const allEntries = loadAllEntries(ss, memberLevelMap);
  if (allEntries.length === 0) {
    ui.alert('No valid entries found in Entries sheet.');
    return;
  }

  // Verify Comp values match this workbook's competition type
  const compType = getCompetitionType();
  if (compType) {
    const mismatches = allEntries.filter(e => e.comp && CONFIG.compMap[e.comp] && e.comp !== compType);
    if (mismatches.length > 0) {
      let msg = 'Warning: ' + mismatches.length + ' entr' + (mismatches.length === 1 ? 'y has' : 'ies have') +
        ' a Comp value that does not match this workbook\'s competition type (' + CONFIG.compMap[compType] + ').\n\n';
      mismatches.slice(0, 20).forEach(e => {
        msg += e.maker + ' — ' + e.title + ' (Comp = "' + e.comp + '")\n';
      });
      if (mismatches.length > 20) msg += '\n...and ' + (mismatches.length - 20) + ' more.';
      msg += '\nPlease review the Entries sheet. Reprocessing will continue.';
      ui.alert('Competition Type Mismatch', msg, ui.ButtonSet.OK);
    }
  }

  buildMemberSheets(ss, allEntries, memberLevelMap);
  buildSummarySheet(ss, allEntries, memberLevelMap);

  ui.alert('Reprocessing complete. Member sheets and Summary have been updated.');
}


// ============================================================
// BUILD: Individual member sheets
// ============================================================
function buildMemberSheets(ss, entries, memberLevelMap) {
  const entriesByMember = groupEntriesByMember(entries);
  const nameKeys = Object.keys(entriesByMember).sort();

  for (const nameKey of nameKeys) {
    const memberEntries = entriesByMember[nameKey];
    if (!memberEntries || memberEntries.length === 0) continue;

    const mInfo = memberLevelMap[nameKey] || { fullName: memberEntries[0].maker, level: memberEntries[0].level };
    const fullName = mInfo.fullName;
    const level    = mInfo.level;

    const sheetName = fullName.substring(0, 31);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    sheet.getRange('A1').setValue(fullName).setFontSize(14).setFontWeight('bold');
    sheet.getRange('B1').setValue('Skill Level: ' + level).setFontWeight('bold');

    const colHeaders = ['Month', 'Category', 'Title', 'Score', 'Honour Award', 'Bonus Point', 'Combined Score'];
    sheet.getRange(3, 1, 1, colHeaders.length)
      .setValues([colHeaders])
      .setFontWeight('bold')
      .setBackground('#4A90D9')
      .setFontColor('#FFFFFF');

    // Group this member's entries by month once, instead of filtering per month
    const entriesByMonth = {};
    for (const e of memberEntries) {
      if (!entriesByMonth[e.month]) entriesByMonth[e.month] = [];
      entriesByMonth[e.month].push(e);
    }

    let rowPtr = 4;
    let grandTotal = 0;
    let totalEntries = 0;
    const awardTotals = { Bronze: 0, Silver: 0, Gold: 0, Diamond: 0, Platinum: 0 };

    for (const month of CONFIG.monthOrder) {
      const monthEntries = entriesByMonth[month];
      if (!monthEntries || monthEntries.length === 0) continue;

      sheet.getRange(rowPtr, 1, 1, colHeaders.length)
        .merge()
        .setValue(month)
        .setBackground('#D9EAD3')
        .setFontWeight('bold');
      rowPtr++;

      let rawMonthScore = 0;
      const distinctCats = new Set(monthEntries.map(e => e.category));
      const monthBonus = distinctCats.size > 1 ? 1 : 0;

      for (const entry of monthEntries) {
        const combined = entry.score + entry.bonus;

        sheet.getRange(rowPtr, 1, 1, 7).setValues([[
          month,
          entry.category,
          entry.title,
          entry.score,
          entry.honour,
          entry.bonus > 0 ? 'Yes' : '',
          combined
        ]]);

        if (entry.honour) {
          const awardColors = {
            Bronze: '#CD7F32', Silver: '#C0C0C0', Gold: '#FFD700',
            Diamond: '#B9F2FF', Platinum: '#E8D5FF'
          };
          sheet.getRange(rowPtr, 5).setBackground(awardColors[entry.honour] || '#FFFFFF');
          awardTotals[entry.honour]++;
        }

        rawMonthScore += entry.score;
        totalEntries++;
        rowPtr++;
      }

      const monthTotal = rawMonthScore + monthBonus;
      sheet.getRange(rowPtr, 1, 1, 7)
        .setValues([[month + ' Total', '', '', '', '', monthBonus > 0 ? '+1 bonus' : '', monthTotal]])
        .setFontWeight('bold')
        .setBackground('#F3F3F3');
      rowPtr++;
      grandTotal += monthTotal;
    }

    rowPtr++;
    sheet.getRange(rowPtr, 1, 1, 7)
      .setValues([['SEASON TOTAL', '', '', '', '', '', grandTotal]])
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#4A90D9')
      .setFontColor('#FFFFFF');
    rowPtr++;

    sheet.getRange(rowPtr, 1).setValue('Total Entries').setFontWeight('bold');
    sheet.getRange(rowPtr, 7).setValue(totalEntries);
    rowPtr++;

    const totalAwards = Object.values(awardTotals).reduce((a, b) => a + b, 0);
    sheet.getRange(rowPtr, 1).setValue('Total Honour Awards').setFontWeight('bold');
    sheet.getRange(rowPtr, 7).setValue(totalAwards);
    rowPtr++;

    for (const award of CONFIG.awardNames) {
      if (awardTotals[award] > 0) {
        sheet.getRange(rowPtr, 2).setValue(award + ':');
        sheet.getRange(rowPtr, 7).setValue(awardTotals[award]);
        rowPtr++;
      }
    }

    [90, 140, 260, 60, 120, 100, 120].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(3);
  }
}

/**
// ============================================================
// BUILD: Summary sheet
// ============================================================
*/
function buildSummarySheet(ss, entries, memberLevelMap) {
  let sheet = ss.getSheetByName('Summary');
  if (!sheet) sheet = ss.insertSheet('Summary');
  sheet.clearContents();
  sheet.clearFormats();

  const title = getCompetitionTitle();

  sheet.getRange('A1').setValue(title)
    .setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Generated: ' + new Date().toLocaleDateString('en-CA'))
    .setFontColor('#666666');

  const headers = [
    'Skill Level', 'Member Name', 'Entries', 'Total Score', 'Total Awards',
    'Platinum', 'Diamond', 'Gold', 'Silver', 'Bronze',
    'Score 15', 'Score 14', 'Score 13', 'Score 12', 'Score 11', 'Score 10', 'Score 9'
  ];
  sheet.getRange(4, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4A90D9')
    .setFontColor('#FFFFFF');

  const stats = {}; // nameKey -> stats

  for (const e of entries) {
    const key = e.nameKey;
    const mInfo = memberLevelMap[key] || { fullName: e.maker, level: e.level };

    if (!stats[key]) {
      stats[key] = {
        fullName: mInfo.fullName,
        level:    mInfo.level,
        entries:  0,
        totalScore: 0,
        awards: { Bronze: 0, Silver: 0, Gold: 0, Diamond: 0, Platinum: 0 },
        scores: { 9: 0, 10: 0, 11: 0, 12: 0, 13: 0, 14: 0, 15: 0 }
      };
    }

    stats[key].entries++;
    stats[key].totalScore += e.score + e.bonus;
    if (e.honour) stats[key].awards[e.honour]++;
    if (stats[key].scores.hasOwnProperty(e.score)) stats[key].scores[e.score]++;
  }

  const rows = Object.entries(stats).map(([key, s]) => {
    const totalAwards = Object.values(s.awards).reduce((a, b) => a + b, 0);
    return { nameKey: key, name: s.fullName, ...s, totalAwards };
  });

  // Sort: level → total score desc → total awards desc → 15 → 14 → ... → 9
  rows.sort((a, b) => {
    const li = CONFIG.levelOrder.indexOf(a.level) - CONFIG.levelOrder.indexOf(b.level);
    if (li !== 0) return li;
    if (b.totalScore !== a.totalScore) return b.totalScore - a.totalScore;
    if (b.totalAwards !== a.totalAwards) return b.totalAwards - a.totalAwards;
    for (const s of [15, 14, 13, 12, 11, 10, 9]) {
      if (b.scores[s] !== a.scores[s]) return b.scores[s] - a.scores[s];
    }
    return 0;
  });

  let rowPtr = 5;
  let currentLevel = null;

  for (const r of rows) {
    if (r.level !== currentLevel) {
      currentLevel = r.level;
      sheet.getRange(rowPtr, 1, 1, headers.length)
        .merge()
        .setValue('── ' + currentLevel + ' Division ──')
        .setFontWeight('bold')
        .setBackground(CONFIG.levelColors[currentLevel] || '#EEEEEE')
        .setHorizontalAlignment('center');
      rowPtr++;
    }

    sheet.getRange(rowPtr, 1, 1, headers.length)
      .setValues([[
        r.level, r.name, r.entries, r.totalScore, r.totalAwards,
        r.awards.Platinum, r.awards.Diamond, r.awards.Gold, r.awards.Silver, r.awards.Bronze,
        r.scores[15], r.scores[14], r.scores[13], r.scores[12], r.scores[11], r.scores[10], r.scores[9]
      ]])
      .setBackground('#FFFFFF');
    rowPtr++;
  }

  [110, 180, 70, 90, 100, 80, 80, 70, 70, 70, 75, 75, 75, 75, 75, 75, 75]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(4);
}


/**
// ============================================================
// PRINTABLE SUMMARY (Columns A–E)
// ============================================================
*/
function buildPrintableSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    SpreadsheetApp.getUi().alert('Summary sheet not found. Please import data first.');
    return;
  }

  let printSheet = ss.getSheetByName('Print Summary');
  if (!printSheet) printSheet = ss.insertSheet('Print Summary');
  printSheet.clearContents();
  printSheet.clearFormats();

  const lastRow = summarySheet.getLastRow();
  if (lastRow < 4) {
    SpreadsheetApp.getUi().alert('No summary data found. Please import data first.');
    return;
  }

  const source = summarySheet.getRange(1, 1, lastRow, 5);
  const data   = source.getValues();

  printSheet.getRange(1, 1, lastRow, 5).setValues(data);

  // Copy main formatting bits row-by-row
  for (let r = 1; r <= lastRow; r++) {
    const src = summarySheet.getRange(r, 1, 1, 5);
    const dst = printSheet.getRange(r, 1, 1, 5);
    dst.setBackgrounds(src.getBackgrounds());
    dst.setFontWeights(src.getFontWeights());
    dst.setFontColors(src.getFontColors());
    dst.setFontSizes(src.getFontSizes());
    dst.setHorizontalAlignments(src.getHorizontalAlignments());
  }

  [110, 180, 70, 90, 100].forEach((w, i) => printSheet.setColumnWidth(i + 1, w));
  printSheet.setFrozenRows(4);
  printSheet.setHiddenGridlines(true);

  ss.setActiveSheet(printSheet);
  SpreadsheetApp.getUi().alert('Print Summary created. Use File → Print to print this sheet.');
}


// ============================================================
// EMAIL: Member report HTML (per member, full table)
//  memberEntries: entries already filtered to this one member
// ============================================================
function getMemberReportHtml(nameKey, memberEntries, memberLevelMap) {
  if (!memberEntries || memberEntries.length === 0) return '';

  const mInfo = memberLevelMap[nameKey] || { fullName: memberEntries[0].maker, level: memberEntries[0].level };
  const fullName = mInfo.fullName;
  const level    = mInfo.level;
  const levelColor = CONFIG.levelColors[level] || '#4A90D9';

  let html = '<div style="font-family: Arial, sans-serif; max-width: 800px;">';
  html += '<h2 style="background:' + levelColor + '; padding: 12px; border-radius: 4px; margin-bottom: 4px;">'
       + fullName + '</h2>';
  html += '<p style="margin-top:0; color:#555;">Skill Level: <strong>' + level + '</strong></p>';

  let grandTotal = 0;
  const awardTotals = { Bronze: 0, Silver: 0, Gold: 0, Diamond: 0, Platinum: 0 };

  for (const month of CONFIG.monthOrder) {
    const monthEntries = memberEntries.filter(e => e.month === month);
    if (monthEntries.length === 0) continue;

    const distinctCats = new Set(monthEntries.map(e => e.category));
    const monthBonus = distinctCats.size > 1 ? 1 : 0;
    const rawMonthScore = monthEntries.reduce((s, e) => s + e.score, 0);
    const monthTotal = rawMonthScore + monthBonus;
    grandTotal += monthTotal;

    html += '<h3 style="background:#D9EAD3; padding: 8px; border-radius: 4px; margin: 12px 0 0 0;">'
         + month + '</h3>';

    // Header row with all requested columns
    html += '<table style="width:100%; border-collapse:collapse; margin-bottom: 12px;">';
    html += '<thead><tr style="background:#4A90D9; color:#fff;">'
         + '<th style="padding:6px 8px; text-align:left;">Title</th>'
         + '<th style="padding:6px 8px; text-align:left;">Category</th>'
         + '<th style="padding:6px 8px; text-align:center;">Score</th>'
         + '<th style="padding:6px 8px; text-align:left;">Honour Award</th>'
         + '<th style="padding:6px 8px; text-align:center;">Bonus</th>'
         + '<th style="padding:6px 8px; text-align:center;">Combined</th>'
         + '</tr></thead><tbody>';

    for (const entry of monthEntries) {
      const awardColors = {
        Bronze: '#CD7F32', Silver: '#C0C0C0', Gold: '#FFD700',
        Diamond: '#B9F2FF', Platinum: '#E8D5FF'
      };
      const awardBg = entry.honour ? (awardColors[entry.honour] || '#FFFFFF') : '#FFFFFF';
      const honourText = entry.honour ? entry.honour : '&nbsp;';       // non‑breaking space if empty
      const bonusText  = entry.bonus > 0 ? '+1' : '&nbsp;';
      const combined   = entry.score + (entry.bonus || 0);

      html += '<tr style="border-bottom: 1px solid #ddd;">'
           + '<td style="padding:6px 8px;">' + escapeHtml(entry.title) + '</td>'
           + '<td style="padding:6px 8px;">' + escapeHtml(entry.category) + '</td>'
           + '<td style="padding:6px 8px; text-align:center;">' + entry.score + '</td>'
           + '<td style="padding:6px 8px; background:' + awardBg + ';">' + honourText + '</td>'
           + '<td style="padding:6px 8px; text-align:center;">' + bonusText + '</td>'
           + '<td style="padding:6px 8px; text-align:center;">' + combined + '</td>'
           + '</tr>';

      if (entry.honour && awardTotals.hasOwnProperty(entry.honour)) {
        awardTotals[entry.honour]++;
      }
    }

    // Monthly total row
    const bonusNote = monthBonus > 0 ? ' (includes +1 monthly bonus)' : '';
    html += '<tr style="background:#F3F3F3; font-weight:bold;">'
         + '<td colspan="5" style="padding:6px 8px; text-align:right;">Monthly Total' + bonusNote + '</td>'
         + '<td style="padding:6px 8px; text-align:center;">' + monthTotal + '</td>'
         + '</tr>';

    html += '</tbody></table>';
  }

  // Season totals
  const totalAwards = Object.values(awardTotals).reduce((a, b) => a + b, 0);
  html += '<table style="width:100%; border-collapse:collapse; margin-top:8px;">';
  html += '<tr style="background:#4A90D9; color:#fff; font-weight:bold;">'
       + '<td style="padding:8px;">Season Total Score</td>'
       + '<td style="padding:8px; text-align:center;">' + grandTotal + '</td></tr>';
  html += '<tr style="background:#F3F3F3;">'
       + '<td style="padding:8px;">Total Honour Awards</td>'
       + '<td style="padding:8px; text-align:center;">' + totalAwards + '</td></tr>';

  for (const award of CONFIG.awardNames) {
    if (awardTotals[award] > 0) {
      const awardColors = {
        Bronze: '#CD7F32', Silver: '#C0C0C0', Gold: '#FFD700',
        Diamond: '#B9F2FF', Platinum: '#E8D5FF'
      };
      html += '<tr style="background:' + (awardColors[award] || '#FFFFFF') + ';">'
           + '<td style="padding:6px 8px; padding-left:24px;">' + award + '</td>'
           + '<td style="padding:6px 8px; text-align:center;">' + awardTotals[award] + '</td></tr>';
    }
  }

  html += '</table></div>';
  return html;
}

// simple HTML escaping for titles/categories
function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


/**
// ============================================================
// EMAIL: Build HTML summary leaderboard for emails
// ============================================================
*/
function buildSummaryHtml(entries, memberLevelMap) {
  const stats = {};
  for (const e of entries) {
    const key = e.nameKey;
    const mInfo = memberLevelMap[key] || { fullName: e.maker, level: e.level };

    if (!stats[key]) {
      stats[key] = {
        fullName: mInfo.fullName,
        level:    mInfo.level,
        entries:  0,
        totalScore: 0,
        totalAwards: 0
      };
    }
    stats[key].entries++;
    stats[key].totalScore += e.score + e.bonus;
    if (e.honour) stats[key].totalAwards++;
  }

  const rows = Object.entries(stats).map(([key, s]) => ({ nameKey: key, ...s }));

  rows.sort((a, b) => {
    const li = CONFIG.levelOrder.indexOf(a.level) - CONFIG.levelOrder.indexOf(b.level);
    if (li !== 0) return li;
    if (b.totalScore !== a.totalScore) return b.totalScore - a.totalScore;
    return b.totalAwards - a.totalAwards;
  });

  let html = '<div style="font-family: Arial, sans-serif; max-width: 700px;">';
  html += '<h2 style="background:#4A90D9; color:#fff; padding:12px; border-radius:4px;">Season Leaderboard</h2>';
  html += '<table style="width:100%; border-collapse:collapse;">';
  html += '<thead><tr style="background:#4A90D9; color:#fff;">' +
    '<th style="padding:6px 8px; text-align:left;">Level</th>' +
    '<th style="padding:6px 8px; text-align:left;">Member</th>' +
    '<th style="padding:6px 8px; text-align:center;">Entries</th>' +
    '<th style="padding:6px 8px; text-align:center;">Score</th>' +
    '<th style="padding:6px 8px; text-align:center;">Awards</th>' +
    '</tr></thead><tbody>';

  let currentLevel = null;
  for (const r of rows) {
    if (r.level !== currentLevel) {
      currentLevel = r.level;
      const bg = CONFIG.levelColors[currentLevel] || '#EEEEEE';
      html += '<tr><td colspan="5" style="background:' + bg + '; font-weight:bold; padding:6px 8px; text-align:center;">── ' + currentLevel + ' Division ──</td></tr>';
    }
    html += '<tr style="border-bottom:1px solid #ddd;">' +
      '<td style="padding:6px 8px;">' + r.level + '</td>' +
      '<td style="padding:6px 8px;">' + r.fullName + '</td>' +
      '<td style="padding:6px 8px; text-align:center;">' + r.entries + '</td>' +
      '<td style="padding:6px 8px; text-align:center;">' + r.totalScore + '</td>' +
      '<td style="padding:6px 8px; text-align:center;">' + r.totalAwards + '</td>' +
      '</tr>';
  }

  html += '</tbody></table></div>';
  return html;
}


/**
// ============================================================
// EMAIL: Send to all members
// ============================================================
*/
function emailAllMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  if (!membersSheet) {
    SpreadsheetApp.getUi().alert('Members sheet not found. Please run Setup first.');
    return;
  }

  const memberData = membersSheet.getDataRange().getValues();
  const memberLevelMap = {};
  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    if (nameRaw) memberLevelMap[nameRaw.toLowerCase()] = { fullName: nameRaw, level };
  }

  const entries = loadAllEntries(ss, memberLevelMap);
  if (entries.length === 0) {
    SpreadsheetApp.getUi().alert('No entries found. Please import CSV data first.');
    return;
  }

  // Generated once — reused as-is for every member's email
  const summaryHtml = buildSummaryHtml(entries, memberLevelMap);
  const competitionTitle = getCompetitionTitle();

  // Group entries by member once, instead of filtering the full list per member
  const entriesByMember = groupEntriesByMember(entries);

  let sentCount = 0;
  let errorCount = 0;

  for (let i = 1; i < memberData.length; i++) {
    const firstName = memberData[i][0];
    const fullName  = String(memberData[i][2]).trim();
    const email     = memberData[i][4];

    if (!fullName || !email) continue;

    const key = fullName.toLowerCase();
    const memberEntries = entriesByMember[key];
    if (!memberEntries || memberEntries.length === 0) continue;

    const individualHtml = getMemberReportHtml(key, memberEntries, memberLevelMap);
    if (!individualHtml) continue;

    const body =
      '<div style="font-family:Arial,sans-serif;">' +
      '<p>Hi ' + firstName + ',</p>' +
      '<p>Please find your <strong>' + competitionTitle + '</strong> competition results below.</p>' +
      individualHtml +
      '<br><hr>' +
      summaryHtml +
      '</div>';

    try {
      GmailApp.sendEmail(email, competitionTitle + ' - Your Competition Results', '', {
        htmlBody: body,
        name: competitionTitle
      });
      sentCount++;
    } catch (e) {
      Logger.log('Failed to send to ' + fullName + ' (' + email + '): ' + e.message);
      errorCount++;
    }
  }

  SpreadsheetApp.getUi().alert('Emails sent: ' + sentCount + '\nFailed: ' + errorCount + '\nCheck Apps Script Logs for details.');
}



/**
// ============================================================
// HELPER: Group entries by member (nameKey) — computed once per call
// ============================================================
*/
function groupEntriesByMember(entries) {
  const grouped = {};
  for (const e of entries) {
    if (!grouped[e.nameKey]) grouped[e.nameKey] = [];
    grouped[e.nameKey].push(e);
  }
  return grouped;
} 
 



/**
// ============================================================
// EMAIL: Send to one member (picker dialog, no typing required)
// ============================================================
*/

function emailOneMember() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  if (!membersSheet) {
    SpreadsheetApp.getUi().alert('Members sheet not found. Please run Setup first.');
    return;
  }

  const memberData = membersSheet.getDataRange().getValues();
  const memberLevelMap = {};
  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    if (nameRaw) memberLevelMap[nameRaw.toLowerCase()] = { fullName: nameRaw, level: level };
  }

  const entries = loadAllEntries(ss, memberLevelMap);
  if (entries.length === 0) {
    SpreadsheetApp.getUi().alert('No entries found. Please import CSV data first.');
    return;
  }

  // Cache the processed entries + member map so "Send" doesn't have to
  // reload and reprocess the Entries sheet from scratch a second time.
  cacheEntriesForEmailPicker(entries, memberLevelMap);

  // Build a de-duplicated, sorted list of members who have at least one entry AND an email address
  const entryKeys = new Set(entries.map(e => e.nameKey));
  const eligible = [];

  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    const email   = String(memberData[i][4]).trim();
    if (!nameRaw || !email) continue;

    const key = nameRaw.toLowerCase();
    if (entryKeys.has(key)) {
      eligible.push({ key: key, fullName: nameRaw, level: level });
    }
  }

  eligible.sort((a, b) => a.fullName.localeCompare(b.fullName));

  const template = HtmlService.createTemplateFromFile('EmailMemberDialog');
  template.members = eligible;
  const html = template.evaluate().setWidth(400).setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(html, 'Email One Member');
}

// ============================================================
// HELPER: Cache processed entries + member map for the email picker
//  so sendEmailToSelectedMember() can reuse them instead of
//  reloading/reprocessing the Entries sheet from scratch.
// ============================================================
const EMAIL_PICKER_CACHE_KEY = 'emailPickerEntriesCache';
const EMAIL_PICKER_CACHE_SECONDS = 120; // picker dialog rarely stays open longer than this

function cacheEntriesForEmailPicker(entries, memberLevelMap) {
  try {
    const cache = CacheService.getScriptCache();
    const payload = JSON.stringify({ entries: entries, memberLevelMap: memberLevelMap });
    cache.put(EMAIL_PICKER_CACHE_KEY, payload, EMAIL_PICKER_CACHE_SECONDS);
  } catch (e) {
    // Cache put can fail if payload exceeds ~100KB (very large seasons).
    // Not fatal — sendEmailToSelectedMember() will fall back to reloading fresh.
    Logger.log('Could not cache entries for email picker: ' + e.message);
  }
}

function getCachedEntriesForEmailPicker() {
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(EMAIL_PICKER_CACHE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

// Called from EmailMemberDialog.html via google.script.run
function sendEmailToSelectedMember(nameKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  if (!membersSheet) throw new Error('Members sheet not found.');

  // Try the cache first (populated when the picker dialog was opened).
  // Falls back to a fresh reload if the cache expired or wasn't set.
  let cached = getCachedEntriesForEmailPicker();
  let entries, memberLevelMap;

  if (cached) {
    entries = cached.entries;
    memberLevelMap = cached.memberLevelMap;
  } else {
    const memberData = membersSheet.getDataRange().getValues();
    memberLevelMap = {};
    for (let i = 1; i < memberData.length; i++) {
      const nameRaw = String(memberData[i][2]).trim();
      const level   = memberData[i][3];
      if (nameRaw) memberLevelMap[nameRaw.toLowerCase()] = { fullName: nameRaw, level: level };
    }
    entries = loadAllEntries(ss, memberLevelMap);
  }

  // Email/first name always come fresh from Members, since email addresses
  // are sensitive to be relying on a 2-minute-old cache for.
  const memberData = membersSheet.getDataRange().getValues();
  let target = null;
  for (let i = 1; i < memberData.length; i++) {
    const nameRaw = String(memberData[i][2]).trim();
    const level   = memberData[i][3];
    const email   = memberData[i][4];
    const first   = memberData[i][0];
    if (!nameRaw) continue;

    const key = nameRaw.toLowerCase();
    if (key === nameKey) {
      target = { key, fullName: nameRaw, level, email, firstName: first || nameRaw.split(' ')[0] };
      break;
    }
  }

  if (!target || !target.email) {
    throw new Error('Member not found or has no email address.');
  }

  const memberEntries = entries.filter(e => e.nameKey === target.key);
  if (memberEntries.length === 0) {
    throw new Error('No entries found for ' + target.fullName + '.');
  }

  const individualHtml = getMemberReportHtml(target.key, memberEntries, memberLevelMap);
  const summaryHtml    = buildSummaryHtml(entries, memberLevelMap);
  const competitionTitle = getCompetitionTitle();

  const body =
    '<div style="font-family:Arial,sans-serif;">' +
    '<p>Hi ' + target.firstName + ',</p>' +
    '<p>Please find your <strong>' + competitionTitle + '</strong> competition results below.</p>' +
    individualHtml +
    '<br><hr>' +
    summaryHtml +
    '</div>';

  GmailApp.sendEmail(target.email, competitionTitle + ' - Your Competition Results', '', {
    htmlBody: body,
    name: competitionTitle
  });

  return 'Email sent to ' + target.fullName + ' at ' + target.email + '.';
}


/**
// ============================================================
// TESTING: Delete all data (for repeated test runs)
// ============================================================
*/
function resetAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset All Data',
    'This will permanently delete:\n' +
    '  • All rows in Members (except headers)\n' +
    '  • All rows in Entries (except headers)\n' +
    '  • All content in CSV Import (except instructions)\n' +
    '  • The Summary and Print Summary sheets\n' +
    '  • Every individual member results sheet\n' +
    '  • The saved Competition Type setting\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) {
    ui.alert('Reset cancelled. No data was changed.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheets that are structural/reserved and should never be deleted as "member sheets"
  const reservedSheets = new Set([
    'Members', 'Honour Levels', 'CSV Import', 'Entries',
    'Summary', 'Print Summary', 'Settings', 'Sheet1', 'Instructions'
  ]);

  // Clear Members (keep header row)
  const membersSheet = ss.getSheetByName('Members');
  if (membersSheet && membersSheet.getLastRow() > 1) {
    membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, membersSheet.getLastColumn()).clearContent();
  }

  // Clear Entries (keep header row)
  const entriesSheet = ss.getSheetByName('Entries');
  if (entriesSheet && entriesSheet.getLastRow() > 1) {
    entriesSheet.getRange(2, 1, entriesSheet.getLastRow() - 1, entriesSheet.getLastColumn()).clearContent();
  }

  // Clear CSV Import (keep instruction row + header row)
  const importSheet = ss.getSheetByName('CSV Import');
  if (importSheet && importSheet.getLastRow() > 2) {
    importSheet.getRange(3, 1, importSheet.getLastRow() - 2, importSheet.getLastColumn()).clearContent();
  }

  // Delete Summary and Print Summary sheets entirely
  ['Summary', 'Print Summary'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  // Delete every individual member sheet (anything not in the reserved list)
  let deletedMemberSheets = 0;
  ss.getSheets().forEach(s => {
    const name = s.getName();
    if (!reservedSheets.has(name)) {
      ss.deleteSheet(s);
      deletedMemberSheets++;
    }
  });

  // Clear the saved competition type so Setup/Import will prompt again
  PropertiesService.getDocumentProperties().deleteProperty('CompetitionSettings');
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet) ss.deleteSheet(settingsSheet);

  ui.alert(
    'Reset complete.\n\n' +
    'Deleted ' + deletedMemberSheets + ' member sheet(s), plus Summary/Print Summary/Settings.\n' +
    'Members, Entries, and CSV Import were cleared but left in place.\n\n' +
    'Run "Setup Workbook" again to re-select a Competition Type before importing.'
  );
}

/**
// ============================================================
// CREATE NEW COMPETITION FILE (copies this spreadsheet + its script)
// ============================================================
*/
function createNewCompetitionFile() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt(
    'Create New Competition File',
    'This will create a brand-new copy of this entire spreadsheet\n' +
    '(including this script, the menu, and all functions) in your Drive.\n\n' +
    'It will NOT include any of the current Members, Entries, or Summary data —\n' +
    'those will be cleared automatically in the new copy.\n\n' +
    'Enter a name for the new file (e.g. "NAPA Colour Print Competition 2027-2028"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const newName = response.getResponseText().trim();
  if (!newName) {
    ui.alert('No file name entered. Cancelled.');
    return;
  }

  // Copy the underlying Drive file (this duplicates the bound script too)
  const file = DriveApp.getFileById(ss.getId());
  const newFile = file.makeCopy(newName);
  const newSs = SpreadsheetApp.openById(newFile.getId());

  // Clean the new copy so it starts fresh (reuse the same logic as Reset All Data)
  cleanCopiedSpreadsheet(newSs);

  const url = newSs.getUrl();
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>New file created:</p>' +
      '<p><a href="' + url + '" target="_blank">' + newName + '</a></p>' +
      '<p>Open it, then run <b>🏆 Competition → 1. Setup Workbook</b> to select its Colour/Monochrome, Format, and Year.</p>')
    .setWidth(400)
    .setHeight(160);
  ui.showModalDialog(htmlOutput, 'New Competition File Ready');
}

/**
// ============================================================
// HELPER: Strip data/settings from a freshly-copied spreadsheet
// ============================================================
*/
function cleanCopiedSpreadsheet(newSs) {
  const reservedSheets = new Set([
    'Members', 'Honour Levels', 'CSV Import', 'Entries',
    'Summary', 'Print Summary', 'Settings', 'Sheet1'
  ]);

  // Clear Members (keep header row)
  const membersSheet = newSs.getSheetByName('Members');
  if (membersSheet && membersSheet.getLastRow() > 1) {
    membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, membersSheet.getLastColumn()).clearContent();
  }

  // Clear Entries (keep header row)
  const entriesSheet = newSs.getSheetByName('Entries');
  if (entriesSheet && entriesSheet.getLastRow() > 1) {
    entriesSheet.getRange(2, 1, entriesSheet.getLastRow() - 1, entriesSheet.getLastColumn()).clearContent();
  }

  // Clear CSV Import (keep instruction row + header row)
  const importSheet = newSs.getSheetByName('CSV Import');
  if (importSheet && importSheet.getLastRow() > 2) {
    importSheet.getRange(3, 1, importSheet.getLastRow() - 2, importSheet.getLastColumn()).clearContent();
  }

  // Delete Summary, Print Summary, Settings sheets
  ['Summary', 'Print Summary', 'Settings'].forEach(name => {
    const s = newSs.getSheetByName(name);
    if (s) newSs.deleteSheet(s);
  });

  // Delete any per-member results sheets carried over from the original
  newSs.getSheets().forEach(s => {
    if (!reservedSheets.has(s.getName())) {
      newSs.deleteSheet(s);
    }
  });

  // Note: Document Properties (like CompetitionSettings) do NOT copy over
  // to the new file automatically, so the new copy will correctly prompt
  // for settings the first time Setup Workbook is run.
}

// ============================================================
// INSTRUCTIONS: Build a formatted in-sheet user manual
// ============================================================
function showInstructions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Instructions');
  if (!sheet) {
    sheet = ss.insertSheet('Instructions');
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  sheet.setColumnWidth(1, 900);
  sheet.setHiddenGridlines(true);

  let row = 1;

  const writeTitle = (text) => {
    sheet.getRange(row, 1).setValue(text)
      .setFontSize(18).setFontWeight('bold').setFontColor('#FFFFFF');
    sheet.getRange(row, 1, 1, 1).setBackground('#4A90D9');
    sheet.setRowHeight(row, 34);
    row++;
  };

  const writeHeading = (text) => {
    row++; // blank spacer before each heading
    sheet.getRange(row, 1).setValue(text)
      .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF');
    sheet.getRange(row, 1).setBackground('#D9EAD3').setFontColor('#000000');
    sheet.setRowHeight(row, 26);
    row++;
  };

  const writeSubheading = (text) => {
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(11);
    row++;
  };

  const writeBody = (text) => {
    sheet.getRange(row, 1).setValue(text).setWrap(true).setFontSize(10);
    row++;
  };

  const writeBullet = (text) => {
    sheet.getRange(row, 1).setValue('• ' + text).setWrap(true).setFontSize(10);
    row++;
  };

  const writeBlank = () => { row++; };

  // ---------------- Title ----------------
  writeTitle('NAPA Competition Management System — Instruction Manual');
  writeBody('Generated: ' + new Date().toLocaleDateString('en-CA'));

  // ---------------- Overview ----------------
  writeHeading('1. Overview');
  writeBody('This system runs entirely inside this Google Sheet using Google Apps Script. It manages a photography club competition across six months (September–February), tracking member details, skill levels (Bronze, Silver, Gold, Diamond, Platinum), entries, scores, bonus points, Honour Awards, individual member results, a club-wide summary, and emailed reports.');
  writeBody('This workbook is configured under the 🏆 Competition menu → Change Competition Settings, combining Colour/Monochrome, Print/Digital, and a starting year into a title such as: NAPA Colour Print Competition 2026-2027.');
  writeBody('Because Colour, Monochrome, Print, and Digital competitions typically run independently, each competition should have its own spreadsheet file. See Section 10 for creating additional files.');

  // ---------------- Setup ----------------
  writeHeading('3. First-Time Setup');
  writeBody('From the 🏆 Competition menu, choose "1. Setup Workbook (first run only)". A dialog will ask you to confirm:');
  writeBullet('Colour or Monochrome');
  writeBullet('Print or Digital');
  writeBullet('Starting Year (enter exactly 4 digits, e.g. 2026)');
  writeBody('Click OK. This builds all required sheets: Settings, Members, Honour Levels, CSV Import, and Entries, and removes the default "Sheet1" tab if empty. You only need to run this once per file. To change these settings later without losing data, use "Change Competition Settings" instead.');

  // ---------------- Members ----------------
  writeHeading('4. Preparing Your Members List');
  writeBody('You do not need to manually enter members first — new names in a CSV import will be detected and you will be asked to confirm adding them. You can also manage members directly on the Members sheet:');
  writeBullet('Column A: First Name');
  writeBullet('Column B: Last Name');
  writeBullet('Column C: Full Name (must exactly match the Maker name used in CSV imports)');
  writeBullet('Column D: Skill Level (dropdown: Bronze, Silver, Gold, Diamond, Platinum)');
  writeBullet('Column E: Email Address');
  writeBody('If you correct a member\'s skill level mid-season, edit it here then run "3. Reprocess Entries" to recalculate their results.');

  // ---------------- CSV Import ----------------
  writeHeading('5. Preparing and Importing CSV Data');
  writeSubheading('5.1 CSV Column Order (paste starting in row 3 of CSV Import)');
  writeBullet('Column A — Comp: C (Colour) or M (Monochrome)');
  writeBullet('Column B — Cat: A, C, J, N, P, T (Action, Creative, Photo-Journalism, Nature, Pictorial, Portrait)');
  writeBullet('Column C — CompNum: competition/round number (informational)');
  writeBullet('Column D — Div: B, S, G, D, P (Bronze, Silver, Gold, Diamond, Platinum)');
  writeBullet('Column E — ImgNum: whole number 1–12 (maps to month)');
  writeBullet('Column F — Maker: member\'s full name');
  writeBullet('Column G — Title: entry title');
  writeBullet('Column H — Score: whole number 9–15');

  writeSubheading('5.2 Entry Number → Month Mapping');
  writeBullet('1, 2 = September');
  writeBullet('3, 4 = October');
  writeBullet('5, 6 = November');
  writeBullet('7, 8 = December');
  writeBullet('9, 10 = January');
  writeBullet('11, 12 = February');
  writeBody('Each member\'s entry number must be unique to them across the whole season.');

  writeSubheading('5.3 Running the Import');
  writeBody('Choose "2. Import CSV Data" from the menu. Every row is validated before anything is written:');
  writeBullet('Category must be one of A, C, J, N, P, T');
  writeBullet('Division must be one of B, S, G, D, P');
  writeBullet('Entry Number must be a whole number 1–12');
  writeBullet('No duplicate entry numbers per member (existing or within the same paste)');
  writeBullet('Skill level in the CSV must match the Members sheet for existing members');
  writeBullet('Comp column must match this workbook\'s Colour/Monochrome setting');
  writeBody('If any row fails, an alert lists every problem row with its actual sheet row number and nothing is written. Fix the flagged rows and re-import. Rows with a mismatched Comp value are skipped individually rather than blocking the whole import. New member names trigger a confirmation dialog (listed alphabetically as "Name - Level") before being added.');

  // ---------------- Results ----------------
  writeHeading('6. Reviewing Results');
  writeBody('Each member with entries gets their own tab showing every entry by month, scores, Honour Awards, bonus points, combined scores, and season totals. The Summary sheet lists all members grouped by Skill Level, sorted by Total Score, then Total Awards, then score tie-breakers (15s, then 14s, etc).');
  writeBody('If you manually edit the Entries sheet directly, run "3. Reprocess Entries (after manual edits)" to rebuild member sheets and the Summary from the corrected data.');

  // ---------------- Email ----------------
  writeHeading('7. Emailing Members');
  writeBody('"4. Email All Members" sends every eligible member (has entries + an email address) their individual results plus the season leaderboard in one email.');
  writeBody('"5. Email One Member (Test)" opens a picker dialog listing eligible members alphabetically with their skill level shown — no typing required, avoiding name-matching errors. Useful for previewing an email before sending to everyone.');

  // ---------------- Print ----------------
  writeHeading('8. Printable Summary');
  writeBody('"6. Build Printable Summary" creates or refreshes a "Print Summary" tab with a simplified version of the Summary sheet, ready for File → Print.');

  // ---------------- Settings ----------------
  writeHeading('9. Changing Competition Settings');
  writeBody('"Change Competition Settings" reopens the same dialog used during Setup, pre-filled with current values, and updates the Settings sheet and Summary title without touching Members, Entries, or CSV Import data.');

  // ---------------- New file ----------------
  writeHeading('10. Creating a New Competition File (New Season or Type)');
  writeBody('Because each competition/season should be its own file, use "🆕 Create New Competition File" to duplicate this entire spreadsheet (including the script) into a new file in your Drive. Any leftover data is automatically cleared from the copy. Open the new file and run Setup Workbook to configure its own settings.');

  // ---------------- Testing ----------------
  writeHeading('11. Testing / Resetting Data');
  writeBody('"⚠️ Reset All Data (Testing)" is destructive: it clears Members, Entries, and CSV Import data, deletes Summary/Print Summary/Settings sheets and every member sheet, and clears saved competition settings. A confirmation warning appears first. Run Setup Workbook again afterward before importing new data.');

  // ---------------- Menu reference ----------------
  writeHeading('12. Menu Reference');
  writeBullet('1. Setup Workbook (first run only) — initial setup, prompts for settings, builds all sheets');
  writeBullet('2. Import CSV Data — validates and imports data from CSV Import');
  writeBullet('3. Reprocess Entries (after manual edits) — rebuilds reports from current Entries data');
  writeBullet('4. Email All Members — sends results + leaderboard to every eligible member');
  writeBullet('5. Email One Member (Test) — picker dialog to email a single member');
  writeBullet('6. Build Printable Summary — creates/refreshes a simplified printable tab');
  writeBullet('Change Competition Settings — update settings without wiping data');
  writeBullet('Reset Members Validation — reapplies the Skill Level dropdown to Members');
  writeBullet('🆕 Create New Competition File — duplicates this file for a new season/competition');
  writeBullet('⚠️ Reset All Data (Testing) — wipes all data and settings');
  writeBullet('📖 View Instructions — shows this manual');

  // ---------------- Troubleshooting ----------------
  writeHeading('13. Troubleshooting');
  writeSubheading('"Required sheets not found. Please run Setup first."');
  writeBody('Setup Workbook hasn\'t been run yet, or a required sheet was deleted/renamed. Run "1. Setup Workbook".');
  writeBlank();
  writeSubheading('"This workbook has no Competition Type set."');
  writeBody('Settings weren\'t saved, or the Settings sheet was deleted. Run "Change Competition Settings" to reconfigure.');
  writeBlank();
  writeSubheading('Import stopped with "Invalid Data Found"');
  writeBody('One or more rows failed validation. The alert lists the exact row number and issue for each. Fix those rows in CSV Import and re-run the import — nothing is written until every row passes.');
  writeBlank();
  writeSubheading('A new member was added at the wrong skill level');
  writeBody('Correct their level in the Members sheet, then run "3. Reprocess Entries" to recalculate their Honour Awards.');
  writeBlank();
  writeSubheading('Emails aren\'t sending');
  writeBody('Check the member has a valid email address in Members. Check Extensions → Apps Script → Executions for error details.');
  writeBlank();
  writeSubheading('I need to fix something after emails have already gone out');
  writeBody('Correct the entry (via Reprocess or a fresh import), then use "Email One Member" to resend just that person\'s corrected results.');
  writeBlank();
  writeSubheading('The script asks for authorization again unexpectedly');
  writeBody('This can happen occasionally, especially on a new browser/device. Follow the prompts again — no data is affected.');

  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);
  sheet.activate();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Competition')
    .addItem('1. Setup Workbook (first run only)', 'showSetupDialog')
    .addSeparator()
    .addItem('2. Import CSV Data', 'importCSV')
    .addItem('3. Reprocess Entries (after manual edits)', 'reprocessEntries')
    .addSeparator()
    .addItem('4. Email All Members', 'emailAllMembers')
    .addItem('5. Email One Member (Test)', 'emailOneMember')
    .addSeparator()
    .addItem('6. Build Printable Summary', 'buildPrintableSummary')
    .addSeparator()
    .addItem('Change Competition Settings', 'showChangeSettingsDialog')
    .addItem('Reset Members Validation', 'resetMembersValidation')
    .addItem('🆕 Create New Competition File', 'createNewCompetitionFile')
    .addItem('⚠️ Reset All Data (Testing)', 'resetAllData')
    .addSeparator()
    .addItem('📖 View Instructions', 'showInstructions')
    .addToUi();
}
