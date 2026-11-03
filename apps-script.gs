/**
 * Feedbackgesprekken VT1: vinkjes voor de pagina op GitHub Pages.
 *
 * 1. Kies hieronder een docentencode (enkel voor Stijn en Joris).
 * 2. Kies bovenaan de functie "installeer" en klik op Uitvoeren (eenmalig).
 * 3. Implementeren > Nieuwe implementatie > Web-app,
 *    uitvoeren als: jezelf, toegang: iedereen. Kopieer de URL (eindigt op /exec)
 *    en plak die in config.js.
 */
const DOCENTENCODE = 'kies-hier-een-code';

const SHEET = 'Planning';
const HEADER = ['sleutel', 'datum', 'dagdeel', 'uur', 'docent', 'student', 'geweest', 'afgevinkt om'];

function installeer() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.getSheets()[0].setName(SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  const rows = PLANNING.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], false, '']);
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
  sh.getRange(2, 7, rows.length, 1).insertCheckboxes();
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('dd/mm hh:mm');
  sh.setFrozenRows(1);
  sh.hideColumns(1);
  sh.autoResizeColumns(2, 7);
}

// Wie in de Sheet zelf afvinkt, krijgt ook een tijdstip.
function onEdit(e) {
  const r = e.range;
  if (r.getSheet().getName() !== SHEET || r.getColumn() !== 7 || r.getRow() < 2) return;
  r.offset(0, 1).setValue(r.getValue() === true ? new Date() : '');
}

function doGet() {
  const v = sheet_().getDataRange().getValues();
  const checks = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][6] === true) checks.push({ key: String(v[i][0]), at: v[i][7] instanceof Date ? v[i][7].toISOString() : '' });
  }
  return json_({ checks });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'ongeldig' }); }
  if (b.code !== DOCENTENCODE) return json_({ ok: false, error: 'code' });
  if (b.check) return json_({ ok: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const i = keys.findIndex(r => String(r[0]) === b.key);
    if (i < 0) return json_({ ok: false, error: 'onbekend' });
    sh.getRange(i + 2, 7, 1, 2).setValues([[!!b.on, b.on ? new Date() : '']]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() { return SpreadsheetApp.getActive().getSheetByName(SHEET); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// sleutel, datum, dagdeel, uur, docent, student
const PLANNING = [
  ["2026-11-12_Drieghe-Lore", "2026-11-12", "vm", "08:30-08:45", "Stijn", "Drieghe Lore"],
  ["2026-11-12_Deliveyne-Ellen", "2026-11-12", "vm", "08:45-09:00", "Stijn", "Deliveyne Ellen"],
  ["2026-11-12_Van-Buggenhout-Iona", "2026-11-12", "vm", "09:00-09:15", "Stijn", "Van Buggenhout Iona"],
  ["2026-11-12_Nevejans-Amber", "2026-11-12", "vm", "09:15-09:30", "Stijn", "Nevejans Amber"],
  ["2026-11-12_Van-desteene-Maite", "2026-11-12", "vm", "09:30-09:45", "Stijn", "Van desteene Maïté"],
  ["2026-11-12_Afia-Boutaina", "2026-11-12", "vm", "10:00-10:15", "Stijn", "Afia Boutaina"],
  ["2026-11-12_Slingeneyer-Leeloo", "2026-11-12", "vm", "10:15-10:30", "Stijn", "Slingeneyer Leeloo"],
  ["2026-11-12_Christiaens-Milan", "2026-11-12", "vm", "10:30-10:45", "Stijn", "Christiaens Milan"],
  ["2026-11-12_Ojuri-Tobi", "2026-11-12", "vm", "10:45-11:00", "Stijn", "Ojuri Tobi"],
  ["2026-11-12_De-Smet-Max", "2026-11-12", "vm", "11:00-11:15", "Stijn", "De Smet Max"],
  ["2026-11-12_Dobbelaere-Maaike", "2026-11-12", "vm", "08:30-08:45", "Joris", "Dobbelaere Maaike"],
  ["2026-11-12_Cissoko-Djenab", "2026-11-12", "vm", "08:45-09:00", "Joris", "Cissoko Djenab"],
  ["2026-11-12_Monballiu-Farah", "2026-11-12", "vm", "09:00-09:15", "Joris", "Monballiu Farah"],
  ["2026-11-12_Segers-Yana", "2026-11-12", "vm", "09:15-09:30", "Joris", "Segers Yana"],
  ["2026-11-12_Braet-Lynn", "2026-11-12", "vm", "09:30-09:45", "Joris", "Braet Lynn"],
  ["2026-11-12_Pareit-Melana", "2026-11-12", "vm", "10:00-10:15", "Joris", "Pareit Melana"],
  ["2026-11-12_Mommerency-Esmee", "2026-11-12", "vm", "10:15-10:30", "Joris", "Mommerency Esmée"],
  ["2026-11-12_Vanthuyne-Jens", "2026-11-12", "vm", "10:30-10:45", "Joris", "Vanthuyne Jens"],
  ["2026-11-12_De-Staelen-Bo", "2026-11-12", "vm", "10:45-11:00", "Joris", "De Staelen Bo"],
  ["2026-11-12_Gashi-Angjela", "2026-11-12", "vm", "11:00-11:15", "Joris", "Gashi Angjela"],
  ["2026-11-12_Barbe-Zino", "2026-11-12", "nm", "12:30-12:45", "Stijn", "Barbé Zino"],
  ["2026-11-12_De-Laender-Timo", "2026-11-12", "nm", "12:45-13:00", "Stijn", "De Laender Timo"],
  ["2026-11-12_Pauwels-Senne", "2026-11-12", "nm", "13:00-13:15", "Stijn", "Pauwels Senne"],
  ["2026-11-12_Heyman-Angelo", "2026-11-12", "nm", "13:15-13:30", "Stijn", "Heyman Angelo"],
  ["2026-11-12_Temmerman-Jordi", "2026-11-12", "nm", "13:30-13:45", "Stijn", "Temmerman Jordi"],
  ["2026-11-12_Baes-Marie", "2026-11-12", "nm", "13:45-14:00", "Stijn", "Baes Marie"],
  ["2026-11-12_De-Ruyver-Eline", "2026-11-12", "vm", "11:15-11:30", "Stijn", "De Ruyver Eline"],
  ["2026-11-12_Weng-Amalia", "2026-11-12", "vm", "11:15-11:30", "Joris", "Weng Amalia"],
  ["2026-11-12_Lerooij-Stan", "2026-11-12", "nm", "15:00-15:15", "Joris", "Lerooij Stan"],
  ["2026-11-12_Peeters-Anna-Sofia-Anna", "2026-11-12", "nm", "15:15-15:30", "Joris", "Peeters Anna Sofia (Anna)"],
  ["2026-11-12_Maesschalck-Ilana", "2026-11-12", "nm", "12:30-12:45", "Joris", "Maesschalck Ilana"],
  ["2026-11-12_Van-Durme-Rhune", "2026-11-12", "nm", "12:45-13:00", "Joris", "Van Durme Rhune"],
  ["2026-11-12_Vasseur-Riet", "2026-11-12", "nm", "13:00-13:15", "Joris", "Vasseur Riet"],
  ["2026-11-12_De-Cock-Thibo", "2026-11-12", "nm", "13:15-13:30", "Joris", "De Cock Thibo"],
  ["2026-11-12_Ostyn-Ameline", "2026-11-12", "nm", "13:30-13:45", "Joris", "Ostyn Améline"],
  ["2026-11-12_Van-Beveren-Axelle", "2026-11-12", "nm", "14:00-14:15", "Joris", "Van Beveren Axelle"],
  ["2026-11-12_Beernaert-Leonardo", "2026-11-12", "nm", "14:15-14:30", "Joris", "Beernaert Leonardo"],
  ["2026-11-12_Pawlikowska-Magdalena", "2026-11-12", "nm", "14:30-14:45", "Joris", "Pawlikowska Magdalena"],
  ["2026-11-12_Verbeke-Thibo", "2026-11-12", "nm", "14:45-15:00", "Joris", "Verbeke Thibo"],
  ["2026-11-19_Verbeke-Thibo", "2026-11-19", "vm", "08:30-08:45", "Stijn", "Verbeke Thibo"],
  ["2026-11-19_Pawlikowska-Magdalena", "2026-11-19", "vm", "08:45-09:00", "Stijn", "Pawlikowska Magdalena"],
  ["2026-11-19_Maesschalck-Ilana", "2026-11-19", "vm", "09:00-09:15", "Stijn", "Maesschalck Ilana"],
  ["2026-11-19_Vasseur-Riet", "2026-11-19", "vm", "09:15-09:30", "Stijn", "Vasseur Riet"],
  ["2026-11-19_Van-Durme-Rhune", "2026-11-19", "vm", "09:30-09:45", "Stijn", "Van Durme Rhune"],
  ["2026-11-19_Van-Beveren-Axelle", "2026-11-19", "vm", "10:00-10:15", "Stijn", "Van Beveren Axelle"],
  ["2026-11-19_Ostyn-Ameline", "2026-11-19", "vm", "10:15-10:30", "Stijn", "Ostyn Améline"],
  ["2026-11-19_Beernaert-Leonardo", "2026-11-19", "vm", "10:30-10:45", "Stijn", "Beernaert Leonardo"],
  ["2026-11-19_De-Cock-Thibo", "2026-11-19", "vm", "10:45-11:00", "Stijn", "De Cock Thibo"],
  ["2026-11-19_Temmerman-Jordi", "2026-11-19", "vm", "08:30-08:45", "Joris", "Temmerman Jordi"],
  ["2026-11-19_Peeters-Anna-Sofia-Anna", "2026-11-19", "vm", "08:45-09:00", "Joris", "Peeters Anna Sofia (Anna)"],
  ["2026-11-19_Pauwels-Senne", "2026-11-19", "vm", "09:00-09:15", "Joris", "Pauwels Senne"],
  ["2026-11-19_De-Laender-Timo", "2026-11-19", "vm", "09:15-09:30", "Joris", "De Laender Timo"],
  ["2026-11-19_Baes-Marie", "2026-11-19", "vm", "09:30-09:45", "Joris", "Baes Marie"],
  ["2026-11-19_Lerooij-Stan", "2026-11-19", "vm", "10:00-10:15", "Joris", "Lerooij Stan"],
  ["2026-11-19_Weng-Amalia", "2026-11-19", "vm", "10:15-10:30", "Joris", "Weng Amalia"],
  ["2026-11-19_Barbe-Zino", "2026-11-19", "vm", "10:30-10:45", "Joris", "Barbé Zino"],
  ["2026-11-19_Heyman-Angelo", "2026-11-19", "vm", "10:45-11:00", "Joris", "Heyman Angelo"],
  ["2026-11-19_De-Ruyver-Eline", "2026-11-19", "vm", "11:00-11:15", "Joris", "De Ruyver Eline"],
  ["2026-11-19_Segers-Yana", "2026-11-19", "nm", "12:30-12:45", "Stijn", "Segers Yana"],
  ["2026-11-19_Dobbelaere-Maaike", "2026-11-19", "nm", "12:45-13:00", "Stijn", "Dobbelaere Maaike"],
  ["2026-11-19_Braet-Lynn", "2026-11-19", "nm", "13:00-13:15", "Stijn", "Braet Lynn"],
  ["2026-11-19_Vanthuyne-Jens", "2026-11-19", "nm", "13:15-13:30", "Stijn", "Vanthuyne Jens"],
  ["2026-11-19_Monballiu-Farah", "2026-11-19", "nm", "13:30-13:45", "Stijn", "Monballiu Farah"],
  ["2026-11-19_Gashi-Angjela", "2026-11-19", "nm", "14:00-14:15", "Stijn", "Gashi Angjela"],
  ["2026-11-19_De-Staelen-Bo", "2026-11-19", "nm", "14:15-14:30", "Stijn", "De Staelen Bo"],
  ["2026-11-19_Mommerency-Esmee", "2026-11-19", "nm", "14:30-14:45", "Stijn", "Mommerency Esmée"],
  ["2026-11-19_Pareit-Melana", "2026-11-19", "nm", "14:45-15:00", "Stijn", "Pareit Melana"],
  ["2026-11-19_Cissoko-Djenab", "2026-11-19", "nm", "15:00-15:15", "Stijn", "Cissoko Djenab"],
  ["2026-11-19_Van-desteene-Maite", "2026-11-19", "nm", "12:30-12:45", "Joris", "Van desteene Maïté"],
  ["2026-11-19_Christiaens-Milan", "2026-11-19", "nm", "12:45-13:00", "Joris", "Christiaens Milan"],
  ["2026-11-19_Deliveyne-Ellen", "2026-11-19", "nm", "13:00-13:15", "Joris", "Deliveyne Ellen"],
  ["2026-11-19_Afia-Boutaina", "2026-11-19", "nm", "13:15-13:30", "Joris", "Afia Boutaina"],
  ["2026-11-19_Van-Buggenhout-Iona", "2026-11-19", "nm", "13:30-13:45", "Joris", "Van Buggenhout Iona"],
  ["2026-11-19_De-Smet-Max", "2026-11-19", "nm", "14:00-14:15", "Joris", "De Smet Max"],
  ["2026-11-19_Nevejans-Amber", "2026-11-19", "nm", "14:15-14:30", "Joris", "Nevejans Amber"],
  ["2026-11-19_Drieghe-Lore", "2026-11-19", "nm", "14:30-14:45", "Joris", "Drieghe Lore"],
  ["2026-11-19_Slingeneyer-Leeloo", "2026-11-19", "nm", "14:45-15:00", "Joris", "Slingeneyer Leeloo"],
  ["2026-11-19_Ojuri-Tobi", "2026-11-19", "nm", "15:00-15:15", "Joris", "Ojuri Tobi"],
  ["2026-11-26_Slingeneyer-Leeloo", "2026-11-26", "vm", "08:30-08:45", "Stijn", "Slingeneyer Leeloo"],
  ["2026-11-26_Heyman-Angelo", "2026-11-26", "vm", "08:45-09:00", "Stijn", "Heyman Angelo"],
  ["2026-11-26_De-Laender-Timo", "2026-11-26", "vm", "09:00-09:15", "Stijn", "De Laender Timo"],
  ["2026-11-26_Baes-Marie", "2026-11-26", "vm", "09:15-09:30", "Stijn", "Baes Marie"],
  ["2026-11-26_Van-desteene-Maite", "2026-11-26", "vm", "09:30-09:45", "Stijn", "Van desteene Maïté"],
  ["2026-11-26_Van-Buggenhout-Iona", "2026-11-26", "vm", "10:00-10:15", "Stijn", "Van Buggenhout Iona"],
  ["2026-11-26_De-Ruyver-Eline", "2026-11-26", "vm", "10:15-10:30", "Stijn", "De Ruyver Eline"],
  ["2026-11-26_Afia-Boutaina", "2026-11-26", "vm", "10:30-10:45", "Stijn", "Afia Boutaina"],
  ["2026-11-26_Christiaens-Milan", "2026-11-26", "vm", "10:45-11:00", "Stijn", "Christiaens Milan"],
  ["2026-11-26_Lerooij-Stan", "2026-11-26", "vm", "11:00-11:15", "Stijn", "Lerooij Stan"],
  ["2026-11-26_Vanthuyne-Jens", "2026-11-26", "vm", "08:30-08:45", "Joris", "Vanthuyne Jens"],
  ["2026-11-26_Beernaert-Leonardo", "2026-11-26", "vm", "08:45-09:00", "Joris", "Beernaert Leonardo"],
  ["2026-11-26_Ostyn-Ameline", "2026-11-26", "vm", "09:00-09:15", "Joris", "Ostyn Améline"],
  ["2026-11-26_Gashi-Angjela", "2026-11-26", "vm", "09:15-09:30", "Joris", "Gashi Angjela"],
  ["2026-11-26_Van-Beveren-Axelle", "2026-11-26", "vm", "09:30-09:45", "Joris", "Van Beveren Axelle"],
  ["2026-11-26_Monballiu-Farah", "2026-11-26", "vm", "10:00-10:15", "Joris", "Monballiu Farah"],
  ["2026-11-26_Cissoko-Djenab", "2026-11-26", "vm", "10:15-10:30", "Joris", "Cissoko Djenab"],
  ["2026-11-26_Maesschalck-Ilana", "2026-11-26", "vm", "10:30-10:45", "Joris", "Maesschalck Ilana"],
  ["2026-11-26_Segers-Yana", "2026-11-26", "vm", "10:45-11:00", "Joris", "Segers Yana"],
  ["2026-11-26_Pauwels-Senne", "2026-11-26", "nm", "12:30-12:45", "Stijn", "Pauwels Senne"],
  ["2026-11-26_Peeters-Anna-Sofia-Anna", "2026-11-26", "nm", "12:45-13:00", "Stijn", "Peeters Anna Sofia (Anna)"],
  ["2026-11-26_Drieghe-Lore", "2026-11-26", "nm", "13:00-13:15", "Stijn", "Drieghe Lore"],
  ["2026-11-26_Nevejans-Amber", "2026-11-26", "nm", "13:15-13:30", "Stijn", "Nevejans Amber"],
  ["2026-11-26_Weng-Amalia", "2026-11-26", "nm", "13:30-13:45", "Stijn", "Weng Amalia"],
  ["2026-11-26_Temmerman-Jordi", "2026-11-26", "nm", "14:00-14:15", "Stijn", "Temmerman Jordi"],
  ["2026-11-26_Barbe-Zino", "2026-11-26", "nm", "14:15-14:30", "Stijn", "Barbé Zino"],
  ["2026-11-26_De-Smet-Max", "2026-11-26", "nm", "14:30-14:45", "Stijn", "De Smet Max"],
  ["2026-11-26_Ojuri-Tobi", "2026-11-26", "nm", "14:45-15:00", "Stijn", "Ojuri Tobi"],
  ["2026-11-26_Deliveyne-Ellen", "2026-11-26", "nm", "15:00-15:15", "Stijn", "Deliveyne Ellen"],
  ["2026-11-26_Mommerency-Esmee", "2026-11-26", "nm", "12:30-12:45", "Joris", "Mommerency Esmée"],
  ["2026-11-26_Pareit-Melana", "2026-11-26", "nm", "12:45-13:00", "Joris", "Pareit Melana"],
  ["2026-11-26_Vasseur-Riet", "2026-11-26", "nm", "13:00-13:15", "Joris", "Vasseur Riet"],
  ["2026-11-26_Verbeke-Thibo", "2026-11-26", "nm", "13:15-13:30", "Joris", "Verbeke Thibo"],
  ["2026-11-26_Van-Durme-Rhune", "2026-11-26", "nm", "13:30-13:45", "Joris", "Van Durme Rhune"],
  ["2026-11-26_De-Staelen-Bo", "2026-11-26", "nm", "14:00-14:15", "Joris", "De Staelen Bo"],
  ["2026-11-26_Pawlikowska-Magdalena", "2026-11-26", "nm", "14:15-14:30", "Joris", "Pawlikowska Magdalena"],
  ["2026-11-26_Dobbelaere-Maaike", "2026-11-26", "nm", "14:30-14:45", "Joris", "Dobbelaere Maaike"],
  ["2026-11-26_Braet-Lynn", "2026-11-26", "nm", "14:45-15:00", "Joris", "Braet Lynn"],
  ["2026-11-26_De-Cock-Thibo", "2026-11-26", "nm", "15:00-15:15", "Joris", "De Cock Thibo"],
  ["2026-12-03_De-Cock-Thibo", "2026-12-03", "vm", "08:30-08:45", "Stijn", "De Cock Thibo"],
  ["2026-12-03_Pareit-Melana", "2026-12-03", "vm", "08:45-09:00", "Stijn", "Pareit Melana"],
  ["2026-12-03_Mommerency-Esmee", "2026-12-03", "vm", "09:00-09:15", "Stijn", "Mommerency Esmée"],
  ["2026-12-03_Van-Durme-Rhune", "2026-12-03", "vm", "09:15-09:30", "Stijn", "Van Durme Rhune"],
  ["2026-12-03_Vasseur-Riet", "2026-12-03", "vm", "09:30-09:45", "Stijn", "Vasseur Riet"],
  ["2026-12-03_Verbeke-Thibo", "2026-12-03", "vm", "10:00-10:15", "Stijn", "Verbeke Thibo"],
  ["2026-12-03_Dobbelaere-Maaike", "2026-12-03", "vm", "10:15-10:30", "Stijn", "Dobbelaere Maaike"],
  ["2026-12-03_Braet-Lynn", "2026-12-03", "vm", "10:30-10:45", "Stijn", "Braet Lynn"],
  ["2026-12-03_De-Staelen-Bo", "2026-12-03", "vm", "10:45-11:00", "Stijn", "De Staelen Bo"],
  ["2026-12-03_Pawlikowska-Magdalena", "2026-12-03", "vm", "11:00-11:15", "Stijn", "Pawlikowska Magdalena"],
  ["2026-12-03_Nevejans-Amber", "2026-12-03", "vm", "08:30-08:45", "Joris", "Nevejans Amber"],
  ["2026-12-03_Weng-Amalia", "2026-12-03", "vm", "08:45-09:00", "Joris", "Weng Amalia"],
  ["2026-12-03_Barbe-Zino", "2026-12-03", "vm", "09:00-09:15", "Joris", "Barbé Zino"],
  ["2026-12-03_Pauwels-Senne", "2026-12-03", "vm", "09:15-09:30", "Joris", "Pauwels Senne"],
  ["2026-12-03_Ojuri-Tobi", "2026-12-03", "vm", "09:30-09:45", "Joris", "Ojuri Tobi"],
  ["2026-12-03_Peeters-Anna-Sofia-Anna", "2026-12-03", "vm", "10:00-10:15", "Joris", "Peeters Anna Sofia (Anna)"],
  ["2026-12-03_Drieghe-Lore", "2026-12-03", "vm", "10:15-10:30", "Joris", "Drieghe Lore"],
  ["2026-12-03_De-Smet-Max", "2026-12-03", "vm", "10:30-10:45", "Joris", "De Smet Max"],
  ["2026-12-03_Temmerman-Jordi", "2026-12-03", "vm", "10:45-11:00", "Joris", "Temmerman Jordi"],
  ["2026-12-03_Deliveyne-Ellen", "2026-12-03", "vm", "11:00-11:15", "Joris", "Deliveyne Ellen"],
  ["2026-12-03_Van-Beveren-Axelle", "2026-12-03", "nm", "12:30-12:45", "Stijn", "Van Beveren Axelle"],
  ["2026-12-03_Maesschalck-Ilana", "2026-12-03", "nm", "12:45-13:00", "Stijn", "Maesschalck Ilana"],
  ["2026-12-03_Ostyn-Ameline", "2026-12-03", "nm", "13:00-13:15", "Stijn", "Ostyn Améline"],
  ["2026-12-03_Beernaert-Leonardo", "2026-12-03", "nm", "13:15-13:30", "Stijn", "Beernaert Leonardo"],
  ["2026-12-03_Gashi-Angjela", "2026-12-03", "nm", "13:30-13:45", "Stijn", "Gashi Angjela"],
  ["2026-12-03_Vanthuyne-Jens", "2026-12-03", "nm", "14:00-14:15", "Stijn", "Vanthuyne Jens"],
  ["2026-12-03_Cissoko-Djenab", "2026-12-03", "nm", "14:15-14:30", "Stijn", "Cissoko Djenab"],
  ["2026-12-03_Monballiu-Farah", "2026-12-03", "nm", "14:30-14:45", "Stijn", "Monballiu Farah"],
  ["2026-12-03_Segers-Yana", "2026-12-03", "nm", "14:45-15:00", "Stijn", "Segers Yana"],
  ["2026-12-03_Heyman-Angelo", "2026-12-03", "nm", "12:30-12:45", "Joris", "Heyman Angelo"],
  ["2026-12-03_Lerooij-Stan", "2026-12-03", "nm", "12:45-13:00", "Joris", "Lerooij Stan"],
  ["2026-12-03_Slingeneyer-Leeloo", "2026-12-03", "nm", "13:00-13:15", "Joris", "Slingeneyer Leeloo"],
  ["2026-12-03_De-Laender-Timo", "2026-12-03", "nm", "13:15-13:30", "Joris", "De Laender Timo"],
  ["2026-12-03_Van-Buggenhout-Iona", "2026-12-03", "nm", "13:30-13:45", "Joris", "Van Buggenhout Iona"],
  ["2026-12-03_Christiaens-Milan", "2026-12-03", "nm", "14:00-14:15", "Joris", "Christiaens Milan"],
  ["2026-12-03_Baes-Marie", "2026-12-03", "nm", "14:15-14:30", "Joris", "Baes Marie"],
  ["2026-12-03_Afia-Boutaina", "2026-12-03", "nm", "14:30-14:45", "Joris", "Afia Boutaina"],
  ["2026-12-03_De-Ruyver-Eline", "2026-12-03", "nm", "14:45-15:00", "Joris", "De Ruyver Eline"],
  ["2026-12-03_Van-desteene-Maite", "2026-12-03", "nm", "15:00-15:15", "Joris", "Van desteene Maïté"],
  ["2026-12-10_Afia-Boutaina", "2026-12-10", "vm", "08:30-08:45", "Stijn", "Afia Boutaina"],
  ["2026-12-10_Baes-Marie", "2026-12-10", "vm", "08:45-09:00", "Stijn", "Baes Marie"],
  ["2026-12-10_Christiaens-Milan", "2026-12-10", "vm", "09:00-09:15", "Stijn", "Christiaens Milan"],
  ["2026-12-10_Deliveyne-Ellen", "2026-12-10", "vm", "09:15-09:30", "Stijn", "Deliveyne Ellen"],
  ["2026-12-10_De-Laender-Timo", "2026-12-10", "vm", "09:30-09:45", "Stijn", "De Laender Timo"],
  ["2026-12-10_Van-desteene-Maite", "2026-12-10", "vm", "10:00-10:15", "Stijn", "Van desteene Maïté"],
  ["2026-12-10_Slingeneyer-Leeloo", "2026-12-10", "vm", "10:15-10:30", "Stijn", "Slingeneyer Leeloo"],
  ["2026-12-10_Lerooij-Stan", "2026-12-10", "vm", "10:30-10:45", "Stijn", "Lerooij Stan"],
  ["2026-12-10_Peeters-Anna-Sofia-Anna", "2026-12-10", "vm", "10:45-11:00", "Stijn", "Peeters Anna Sofia (Anna)"],
  ["2026-12-10_Heyman-Angelo", "2026-12-10", "vm", "11:00-11:15", "Stijn", "Heyman Angelo"],
  ["2026-12-10_Maesschalck-Ilana", "2026-12-10", "vm", "08:30-08:45", "Joris", "Maesschalck Ilana"],
  ["2026-12-10_De-Cock-Thibo", "2026-12-10", "vm", "08:45-09:00", "Joris", "De Cock Thibo"],
  ["2026-12-10_Ostyn-Ameline", "2026-12-10", "vm", "09:00-09:15", "Joris", "Ostyn Améline"],
  ["2026-12-10_Gashi-Angjela", "2026-12-10", "vm", "09:15-09:30", "Joris", "Gashi Angjela"],
  ["2026-12-10_Van-Beveren-Axelle", "2026-12-10", "vm", "09:30-09:45", "Joris", "Van Beveren Axelle"],
  ["2026-12-10_Pawlikowska-Magdalena", "2026-12-10", "vm", "10:00-10:15", "Joris", "Pawlikowska Magdalena"],
  ["2026-12-10_Vasseur-Riet", "2026-12-10", "vm", "10:15-10:30", "Joris", "Vasseur Riet"],
  ["2026-12-10_Van-Durme-Rhune", "2026-12-10", "vm", "10:30-10:45", "Joris", "Van Durme Rhune"],
  ["2026-12-10_Verbeke-Thibo", "2026-12-10", "vm", "10:45-11:00", "Joris", "Verbeke Thibo"],
  ["2026-12-10_De-Ruyver-Eline", "2026-12-10", "nm", "12:30-12:45", "Stijn", "De Ruyver Eline"],
  ["2026-12-10_Nevejans-Amber", "2026-12-10", "nm", "12:45-13:00", "Stijn", "Nevejans Amber"],
  ["2026-12-10_Barbe-Zino", "2026-12-10", "nm", "13:00-13:15", "Stijn", "Barbé Zino"],
  ["2026-12-10_De-Smet-Max", "2026-12-10", "nm", "13:15-13:30", "Stijn", "De Smet Max"],
  ["2026-12-10_Temmerman-Jordi", "2026-12-10", "nm", "13:30-13:45", "Stijn", "Temmerman Jordi"],
  ["2026-12-10_Ojuri-Tobi", "2026-12-10", "nm", "14:00-14:15", "Stijn", "Ojuri Tobi"],
  ["2026-12-10_Van-Buggenhout-Iona", "2026-12-10", "nm", "14:15-14:30", "Stijn", "Van Buggenhout Iona"],
  ["2026-12-10_Weng-Amalia", "2026-12-10", "nm", "14:30-14:45", "Stijn", "Weng Amalia"],
  ["2026-12-10_Pauwels-Senne", "2026-12-10", "nm", "14:45-15:00", "Stijn", "Pauwels Senne"],
  ["2026-12-10_Drieghe-Lore", "2026-12-10", "nm", "15:00-15:15", "Stijn", "Drieghe Lore"],
  ["2026-12-10_Vanthuyne-Jens", "2026-12-10", "nm", "12:30-12:45", "Joris", "Vanthuyne Jens"],
  ["2026-12-10_Segers-Yana", "2026-12-10", "nm", "12:45-13:00", "Joris", "Segers Yana"],
  ["2026-12-10_Pareit-Melana", "2026-12-10", "nm", "13:00-13:15", "Joris", "Pareit Melana"],
  ["2026-12-10_Beernaert-Leonardo", "2026-12-10", "nm", "13:15-13:30", "Joris", "Beernaert Leonardo"],
  ["2026-12-10_Braet-Lynn", "2026-12-10", "nm", "13:30-13:45", "Joris", "Braet Lynn"],
  ["2026-12-10_Monballiu-Farah", "2026-12-10", "nm", "14:00-14:15", "Joris", "Monballiu Farah"],
  ["2026-12-10_Mommerency-Esmee", "2026-12-10", "nm", "14:15-14:30", "Joris", "Mommerency Esmée"],
  ["2026-12-10_Cissoko-Djenab", "2026-12-10", "nm", "14:30-14:45", "Joris", "Cissoko Djenab"],
  ["2026-12-10_De-Staelen-Bo", "2026-12-10", "nm", "14:45-15:00", "Joris", "De Staelen Bo"],
  ["2026-12-10_Dobbelaere-Maaike", "2026-12-10", "nm", "15:00-15:15", "Joris", "Dobbelaere Maaike"],
  ["2026-12-17_Dobbelaere-Maaike", "2026-12-17", "vm", "08:30-08:45", "Stijn", "Dobbelaere Maaike"],
  ["2026-12-17_Pareit-Melana", "2026-12-17", "vm", "08:45-09:00", "Stijn", "Pareit Melana"],
  ["2026-12-17_Monballiu-Farah", "2026-12-17", "vm", "09:00-09:15", "Stijn", "Monballiu Farah"],
  ["2026-12-17_De-Staelen-Bo", "2026-12-17", "vm", "09:15-09:30", "Stijn", "De Staelen Bo"],
  ["2026-12-17_Mommerency-Esmee", "2026-12-17", "vm", "09:30-09:45", "Stijn", "Mommerency Esmée"],
  ["2026-12-17_Braet-Lynn", "2026-12-17", "vm", "10:00-10:15", "Stijn", "Braet Lynn"],
  ["2026-12-17_Cissoko-Djenab", "2026-12-17", "vm", "10:15-10:30", "Stijn", "Cissoko Djenab"],
  ["2026-12-17_Beernaert-Leonardo", "2026-12-17", "vm", "10:30-10:45", "Stijn", "Beernaert Leonardo"],
  ["2026-12-17_Segers-Yana", "2026-12-17", "vm", "10:45-11:00", "Stijn", "Segers Yana"],
  ["2026-12-17_Vanthuyne-Jens", "2026-12-17", "vm", "11:00-11:15", "Stijn", "Vanthuyne Jens"],
  ["2026-12-17_Pauwels-Senne", "2026-12-17", "vm", "08:30-08:45", "Joris", "Pauwels Senne"],
  ["2026-12-17_Barbe-Zino", "2026-12-17", "vm", "08:45-09:00", "Joris", "Barbé Zino"],
  ["2026-12-17_Ojuri-Tobi", "2026-12-17", "vm", "09:00-09:15", "Joris", "Ojuri Tobi"],
  ["2026-12-17_Weng-Amalia", "2026-12-17", "vm", "09:15-09:30", "Joris", "Weng Amalia"],
  ["2026-12-17_Drieghe-Lore", "2026-12-17", "vm", "09:30-09:45", "Joris", "Drieghe Lore"],
  ["2026-12-17_Nevejans-Amber", "2026-12-17", "vm", "10:00-10:15", "Joris", "Nevejans Amber"],
  ["2026-12-17_Temmerman-Jordi", "2026-12-17", "vm", "10:15-10:30", "Joris", "Temmerman Jordi"],
  ["2026-12-17_Van-Buggenhout-Iona", "2026-12-17", "vm", "10:30-10:45", "Joris", "Van Buggenhout Iona"],
  ["2026-12-17_De-Ruyver-Eline", "2026-12-17", "vm", "10:45-11:00", "Joris", "De Ruyver Eline"],
  ["2026-12-17_De-Smet-Max", "2026-12-17", "vm", "11:00-11:15", "Joris", "De Smet Max"],
  ["2026-12-17_Pawlikowska-Magdalena", "2026-12-17", "nm", "12:30-12:45", "Stijn", "Pawlikowska Magdalena"],
  ["2026-12-17_Van-Durme-Rhune", "2026-12-17", "nm", "12:45-13:00", "Stijn", "Van Durme Rhune"],
  ["2026-12-17_De-Cock-Thibo", "2026-12-17", "nm", "13:00-13:15", "Stijn", "De Cock Thibo"],
  ["2026-12-17_Van-Beveren-Axelle", "2026-12-17", "nm", "13:15-13:30", "Stijn", "Van Beveren Axelle"],
  ["2026-12-17_Vasseur-Riet", "2026-12-17", "nm", "13:30-13:45", "Stijn", "Vasseur Riet"],
  ["2026-12-17_Verbeke-Thibo", "2026-12-17", "nm", "14:00-14:15", "Stijn", "Verbeke Thibo"],
  ["2026-12-17_Ostyn-Ameline", "2026-12-17", "nm", "14:15-14:30", "Stijn", "Ostyn Améline"],
  ["2026-12-17_Gashi-Angjela", "2026-12-17", "nm", "14:30-14:45", "Stijn", "Gashi Angjela"],
  ["2026-12-17_Maesschalck-Ilana", "2026-12-17", "nm", "14:45-15:00", "Stijn", "Maesschalck Ilana"],
  ["2026-12-17_Heyman-Angelo", "2026-12-17", "nm", "12:30-12:45", "Joris", "Heyman Angelo"],
  ["2026-12-17_Van-desteene-Maite", "2026-12-17", "nm", "12:45-13:00", "Joris", "Van desteene Maïté"],
  ["2026-12-17_Afia-Boutaina", "2026-12-17", "nm", "13:00-13:15", "Joris", "Afia Boutaina"],
  ["2026-12-17_Slingeneyer-Leeloo", "2026-12-17", "nm", "13:15-13:30", "Joris", "Slingeneyer Leeloo"],
  ["2026-12-17_Christiaens-Milan", "2026-12-17", "nm", "13:30-13:45", "Joris", "Christiaens Milan"],
  ["2026-12-17_De-Laender-Timo", "2026-12-17", "nm", "14:00-14:15", "Joris", "De Laender Timo"],
  ["2026-12-17_Peeters-Anna-Sofia-Anna", "2026-12-17", "nm", "14:15-14:30", "Joris", "Peeters Anna Sofia (Anna)"],
  ["2026-12-17_Baes-Marie", "2026-12-17", "nm", "14:30-14:45", "Joris", "Baes Marie"],
  ["2026-12-17_Lerooij-Stan", "2026-12-17", "nm", "14:45-15:00", "Joris", "Lerooij Stan"],
  ["2026-12-17_Deliveyne-Ellen", "2026-12-17", "nm", "15:00-15:15", "Joris", "Deliveyne Ellen"]
];
