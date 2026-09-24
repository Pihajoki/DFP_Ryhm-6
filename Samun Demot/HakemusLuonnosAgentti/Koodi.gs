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
      .setTitle('Hakemus luonnos AI-agentti');
  SpreadsheetApp.getUi().showSidebar(html);
}

function avaaIkkuna() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(800)
      .setHeight(1050);
  SpreadsheetApp.getUi().showModalDialog(html, 'Hakemus luonnos AI-agentti');
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

// Hakee tietyn kansion kaikkien tiedostojen nimet ja tekstit
function lueDriveKansionSisalto(folderUrlOrId) {
  // Irrotetaan ID linkistä jos syötetty koko URL
  var folderId = folderUrlOrId;
  if (folderUrlOrId.indexOf('/folders/') !== -1) {
    folderId = folderUrlOrId.split('/folders/')[1].split('?')[0];
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var tiedostotTiedot = [];

    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();

      // Luetaan vain tekstimuotoiset / Google Docs -tiedostot
      if (mimeType === MimeType.GOOGLE_DOCS) {
        var doc = DocumentApp.openById(file.getId());
        tiedostotTiedot.push({
          nimi: file.getName(),
          sisalto: doc.getBody().getText()
        });
      } else if (mimeType === MimeType.PLAIN_TEXT) {
        tiedostotTiedot.push({
          nimi: file.getName(),
          sisalto: file.getBlob().getDataAsString()
        });
      }
    }

    return tiedostotTiedot;
  } catch (e) {
    throw new Error("Kansion lukeminen epäonnistui: " + e.message);
  }
}