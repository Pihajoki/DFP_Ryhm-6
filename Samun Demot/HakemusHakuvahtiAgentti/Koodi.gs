function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Oma Sovellus')
      .addItem('Avaa sivupaneeli', 'avaaSivupaneeli')
      .addItem('Avaa popup-ikkuna', 'avaaIkkuna')
      .addToUi();
}

function avaaSivupaneeli() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hakuvahti');
  SpreadsheetApp.getUi().showSidebar(html);
}

function avaaIkkuna() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(800)
      .setHeight(1050);
  SpreadsheetApp.getUi().showModalDialog(html, 'Hakuvahti');
}

// Alustaa taulukkorivin 2 ja otsikot tarvittaessa
function alustaTaulukkoTarvittaessa(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Parametrit (Sanat)', 'Sähköpostit', 'Linkki', 'Tila', 'AI Agentin löydökset']);
  }
}

// Hakee aiemmat tallennetut tiedot suoraan riviltä 2
function haeTiedotSheetsista() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  alustaTaulukkoTarvittaessa(sheet);
  
  if (sheet.getLastRow() < 2) return null;

  // Haetaan arvot riviltä 2 (Sarakkeet: Parametrit, Sähköpostit, Linkki)
  var values = sheet.getRange(2, 1, 1, 4).getValues()[0];
  
  return {
    parametrit: values[0] || '',
    sahkoposti: values[1] || '',
    linkki: values[2] || '',
    aiLoydokset: values[3] || ''
  };
}

// Tallentaa yksittäisen osion (Parametrit tai Sähköpostit) riville 2
function tallennaYksittainenTieto(tyyppi, arvo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  alustaTaulukkoTarvittaessa(sheet);

  if (tyyppi === 'parametri') {
    // Sarake A (Sarake 1) = Parametrit
    sheet.getRange(2, 1).setValue(arvo);
  } else if (tyyppi === 'sahkoposti') {
    // Sarake B (Sarake 2) = Sähköpostit
    sheet.getRange(2, 2).setValue(arvo);
  } else if (tyyppi === 'aiLoydokset') {
    sheet.getRange(2, 4).setValue(arvo); // Sarake D
  }
  return true;
}

// Tallentaa kaikki lomakkeen tiedot samalle riville (Rivi 2)
function prosessoiTiedot(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  alustaTaulukkoTarvittaessa(sheet);

  // Päivitetään kentät suoraan riville 2
  sheet.getRange(2, 1).setValue(data.parametrit);
  sheet.getRange(2, 2).setValue(data.sahkoposti);
  sheet.getRange(2, 3).setValue(data.linkki);
  sheet.getRange(2, 4).setValue(data.tila);

  return true;
}