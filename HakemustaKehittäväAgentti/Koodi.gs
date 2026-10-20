/**
 * Luo Google Sheetsiin mukautetun ylävalikon sovelluksen avaimelle.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Kehitys Agenti')
      .addItem('Avaa sivupaneeli', 'avaaSivupaneeli')
      .addItem('Avaa popup-ikkuna', 'avaaIkkuna')
      .addToUi();
}

/**
 * Avaa sivupaneelin.
 */
function avaaSivupaneeli() {
  let html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hakemusta kehittävä Agentti');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Avaa popup-ikkunan (modal dialog).
 */
function avaaIkkuna() {
  let html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(1300)
      .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doGet(e) {

  const userProps = PropertiesService.getUserProperties();
  let sheetId = userProps.getProperty('KAYTTAJAN_SHEET_ID');

  if (!sheetId) {
    const uusiSS = SpreadsheetApp.create('Hakemusta kehittävä agentti.');
    sheetId = uusiSS.getId();
    
    // Tallennetaan ID UserProperties-muistiin
    userProps.setProperty('KAYTTAJAN_SHEET_ID', sheetId);
    
    // Nimetään ensimmäinen välilehti
    let sheet = uusiSS.getSheets()[0];
    sheet.setName('Keskusteluhistoria');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
  .setTitle('Hakemusta kehittävä Agentti')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

/**
 * PAINIKE 2: Lukee tallennetun aineiston Google Docista (docId) Sheetin 2. välilehdeltä
 * ja suorittaa analyysin käyttäjän valitsemalla mallilla ja syötteellä.
 * 
 * @param {Object} data - { viesti: "Käyttäjän kysymys", malli: "gemini-1.5-pro" }
 * @returns {string} Geminin vastaus
 */
function AgenttiKeskustelu(data) {
  try {
    // 1. Haetaan API-avain
    const props = PropertiesService.getScriptProperties();
    const GEMINI_API_KEY = props.getProperty("GEMINI_API_KEY");

    if (!GEMINI_API_KEY) {
      throw new Error("GEMINI_API_KEY puuttuu Script Propertiesista!");
    }

        // Etsitään välilehteä
    const ss = haeKayttajanSheet();
    let sheet1 = ss.getSheets()[0];

    if (typeof alustaTaulukkoHistoria === "function") {
      alustaTaulukkoHistoria(sheet1);
    }

    // Poimitaan UI:lta tulleet tiedot turvallisesti
    const kayttajanSyote = data && data.viesti ? data.viesti : "";
    Logger.log(kayttajanSyote);
    const valittuMalli = data && data.malli ? data.malli : "gemini-3.5-flash-lite";

    // 2. Luetaan uusin datarivi Sheetin 2. välilehdeltä (indeksi 1)
    const sheet = ss.getSheetByName("Parametrit");
    if (!sheet) {
      throw new Error("Virhe: Parametrit-välilehteä ei löytynyt! Tallenna ensin lähdeaineisto!");
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error("Ei tallennettuja aineistoja! Aja ensin taustadatan haku.");
    }

    const riviData = sheet.getRange(lastRow, 1, 1, 6).getValues()[0];
    const docId = riviData[4]; // Sarake E

    if (!docId) {
      throw new Error("Taulukosta ei löytynyt validia Google Doc ID:tä (Sarake E).");
    }

    // 3. Haetaan aiemmin tallennettu aineisto Google Docista
    Logger.log("Haetaan aineisto Google Docista ID:llä: " + docId);
    const doc = DocumentApp.openById(docId);
    const lahdeAineisto = doc.getBody().getText();

    // 4. Haetaan keskusteluhistoria.
    const keskusteluHistoria = haeAiempiHistoria(sheet1, kayttajanSyote);
    keskusteluHistoria.forEach(v => {
      Logger.log(v.role.toUpperCase() + ": " + v.parts[0].text);
    });


    // 5. Luodaan Payload (Tapa 1: systemInstruction + contents)
        const payload = {
          // A) Rooli, tehtävä ja tausta-aineisto erilliseen systemInstruction-kenttään
          systemInstruction: {
            parts: [
              {
                text: `Olet asiantunteva EURA 2021 -hankehakemusluonnoksen kehittäjä-agentti. Tehtäväsi on arvioida ja parantaa hakemuksen laatua huomioiden lähdeaineisto. Tärkeintä on vastata täsmällisesti käyttäjän antamaan syötteeseen.

    ---
    TAUSTA-AINEISTO (1. DRIVE-LÄHDEDATA, 2. HAETTAVAN HANKKEEN TIEDOT, 3. TÄMÄN HETKINEN HAKEMUSLUONNOS):

    ${lahdeAineisto || "Ei erillistä aineistoa"}
    ---

    Anna asiantuntevaa, konkreettista ja kehittävää palautetta käyttäjän pyyntöjen perusteella.`
              }
            ]
          },

          // B) Koko keskusteluhistoria + käyttäjän uusin viesti
          contents: keskusteluHistoria
        };

    // 6. Kutsutaan Geminiä UI:sta tulleella mallilla
    Logger.log("Kutsutaan Geminiä mallilla: " + valittuMalli);
    const vastausJson = kutsuGemini(payload, GEMINI_API_KEY, valittuMalli);
    const vastausTeksti = vastausJson.candidates[0].content.parts.find(p => p.text !== undefined).text;

    tallennaGeminiKeskustelu(kayttajanSyote, vastausJson, valittuMalli);
    return {
      success: true,
      vastaus: vastausTeksti,
      historia: haeKeskusteluhistoriaUI(sheet1) // Haetaan uusin historia 1. välilehdeltä
    };

  } catch (err) {
    Logger.log("Virhe AgenttiKeskustelu-funktiossa: " + err.toString());
    
    try {
      const ss = haeKayttajanSheet();
      const sheet = ss.getSheets()[1];
      sheet.getRange(sheet.getLastRow(), 8).setValue("Virhe: " + err.message);
    } catch(e) {}

    throw new Error("Kutsun suoritus epäonnistui: " + err.message);
  }
}

/**
 * PAINIKE 1: Haetaan taustadata, luodaan Google Doc Driven juureen 
 * ja tallennetaan parametrit + docId Taulukon 2. välilehdelle.
 * 
 * @param {Object} lomakeData - HTML UI:lta tulevat kentät
 */
function tallennaParametritJaAineisto(lomakeData) {

  const props = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();

  try {
    let ss = haeKayttajanSheet();
    let valilehdenNimi = "Parametrit";

    // Etsitään välilehteä nimen perusteella
    let sheet = ss.getSheetByName(valilehdenNimi);

    // Jos välilehteä ei ole olemassa, luodaan se
    if (!sheet) {
      sheet = ss.insertSheet(valilehdenNimi);
    }

    if (typeof alustaTaulukkoTarvittaessa === "function") {
      alustaTaulukkoTarvittaessa(sheet);
    }

    // Poimitaan arvot turvallisesti
    let linkki = lomakeData && lomakeData.linkki ? lomakeData.linkki : "";
    let lahdeUrl = lomakeData && lomakeData.lahdeUrl ? lomakeData.lahdeUrl : "";
    let hakemusTunnus = lomakeData && lomakeData.hakemusUrl ? lomakeData.hakemusUrl : "";

    Logger.log(linkki);
    Logger.log(lahdeUrl);
    Logger.log(hakemusTunnus);

    Logger.log("--- VAIHE 1: TAUSTADATAN HAKU KÄYNNISTETTY ---");
    Logger.log("Hakemuksen tunnus/nimi: " + hakemusTunnus);


    userProps.setProperty("STATUS", "📂 Luetaan lähdemateriaalia Drivesta...");
    // 1. Haetaan lähdeaineisto (Drive)
    let lahdeMateriaaliRaw;
    let lahdeTeksti = ""

    if (lahdeUrl && lahdeUrl.trim() !== "") {
      try {
        lahdeMateriaaliRaw = lueDriveKansionSisalto(lahdeUrl);
      } catch(e) {
        throw new Error("Lähdemateriaalin luku epäonnistui: " + e.message);
      }

      // Muotoillaan Drive-aineisto tekstiksi ja varmistetaan tekstinsiivous
      if (Array.isArray(lahdeMateriaaliRaw)) {
        lahdeTeksti = lahdeMateriaaliRaw.map(t => `--- TIEDOSTO: ${t.nimi} ---\n\n${t.sisalto}`).join("\n\n");
      } else {
        lahdeTeksti = lahdeMateriaaliRaw || "";
      }
    }

    userProps.setProperty("STATUS", "📝 Selataan läpi hankkeen sivustoa ja sen alalinkkejä...");
    // 2. Haetaan hakemuksen tiedot (Web/Jina)
    let hakemuksenTiedot;
    if (linkki && linkki.trim() !== "") {
      try {
        hakemuksenTiedot = haeTiedotSivulta(linkki, 8);
      } catch (e) {
        throw new Error("Hakemuksen tietojen hakeminen epäonnistui: " + e.message);
      }
    }


    userProps.setProperty("STATUS", "📂 Luetaan hakemuspohjaa...");
    //3. Haetaan täytetty hakemusluonnos
    let hakemusLuonnos;
    try {
    hakemusLuonnos = lueHakemus(hakemusTunnus);
    } catch (e) {
      throw new Error("HakemusLuonnoksen hakeminen epäonnistui: " + e.message);
    }

    userProps.setProperty("STATUS", "Tallennetaan...");
    // 4. Tallennetaan aineisto Google Docsiin Driven juureen ja otetaan ID talteen
    let tiedostonNimi = `Gemini_Aineisto`;
    
    // Luodaan dokumentti juureen
    let docUrl = tallennaAineistotGoogleDocsiin(lahdeTeksti, hakemuksenTiedot, hakemusLuonnos, tiedostonNimi);
    
    // Kaivetaan ID luodun dokumentin URL-osoitteesta
    let docId = docUrl.match(/[-\w]{25,}/)[0];
    Logger.log("Aineisto tallennettu Google Docsiin. Doc ID: " + docId);

    // 4. Tallennetaan parametrit ja luodun dokumentin ID Sheetin 2. välilehdelle
    let aikaleima = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    sheet.getRange(2, 1, 1, 6).setValues([[
      aikaleima,
      hakemusTunnus,
      lahdeUrl,
      linkki,
      docId,
      docUrl
    ]]);

    Logger.log("Parametrit ja Doc ID tallennettu taulukon 2. välilehdelle.");

    // 5. Palautetaan vastaus HTML UI:lle
    return {
      success: true,
      message: "Taustadata haettu ja tallennettu onnistuneesti!",
      docId: docId,
      docUrl: docUrl
    };

  } catch (err) {
    Logger.log("Virhe vaiheessa 1: " + err.toString());
    throw new Error("Virhe taustadatan haussa: " + err.message);
  }
}

/**
 * Pääfunktio: Lukee Google Drive -kansion tai yksittäisen tiedoston tiedot ja sisällön.
 */
function lueDriveKansionSisalto(url) {
  if (!url || url.trim() === "") {
    throw new Error("Drive-osoite tai ID puuttuu.");
  }

  let id = haeKansioIdUrlsta(url);
  if (!id) {
    Logger.log("Virheellinen Drive URL tai ID: " + url);
    throw new Error("Virheellinen Drive URL tai ID: " + url);
  }

  let tulokset = [];

  try {
    // 1. Koitetaan ensin hakea tiedostona
    let tiedosto = DriveApp.getFileById(id);
    let mimeType = tiedosto.getMimeType();

    // 2. Jos kyseessä on KANSIO
    if (mimeType === MimeType.FOLDER) {
      let kansio = DriveApp.getFolderById(id);
      let tiedostot = kansio.getFiles();

      while (tiedostot.hasNext()) {
        let f = tiedostot.next();
        tulokset.push({
          nimi: f.getName(),
          id: f.getId(),
          url: f.getUrl(),
          tyyppi: f.getMimeType(),
          sisalto: lueTiedostonTekstiTurvallisesti(f)
        });
      }
    } 
    // 3. Jos kyseessä on YKSITTÄINEN TIEDOSTO (Doc, PDF, Word, Txt jne.)
    else {
      tulokset.push({
        nimi: tiedosto.getName(),
        id: tiedosto.getId(),
        url: tiedosto.getUrl(),
        tyyppi: mimeType,
        sisalto: lueTiedostonTekstiTurvallisesti(tiedosto)
      });
    }

    // 4. Tarkistetaan että tuloksia todella löytyi
    if (tulokset.length === 0 && mimeType === MimeType.FOLDER) {
      Logger.log("Varoitus: Kansio oli tyhjä: " + url);
    }

    return tulokset;

  } catch (err) {
    Logger.log("Virhe kohteen (ID: " + id + ") käsittelyssä: " + err.toString());
    // PÄÄOIKAISU: Heitetään virhe edelleen pääfunktiolle, jotta ajo KESKEYTYY!
    throw new Error("Drive-kohteen lukeminen epäonnistui (ID: " + id + "): " + err.message);
  }
}

/**
 * Lukee yksittäisen tiedoston tekstin puhtaana, tiivistettynä tekstinä (ei HTML-muotoiluja tai turhia tyhjiä rivejä).
 */
function lueTiedostonTekstiTurvallisesti(tiedosto) {
  let mimeType = tiedosto.getMimeType();
  let raakaTeksti = "";

  try {
    // 1. Google Docs -tiedostot -> Luetaan PUHTAANA TEKSTINÄ
    if (mimeType === MimeType.GOOGLE_DOCS) {
      raakaTeksti = DocumentApp.openById(tiedosto.getId()).getBody().getText();
    } 
    // 2. Tavalliset tekstitiedostot (.txt, .csv, .json, .md)
    else if (mimeType === MimeType.PLAIN_TEXT || (mimeType.indexOf("text/") === 0 && mimeType !== "text/html")) {
      raakaTeksti = tiedosto.getBlob().getDataAsString("UTF-8");
    }
    // 3. HTML-tiedostot -> Riisutaan HTML-tägit pois
    else if (mimeType === "text/html") {
      let html = tiedosto.getBlob().getDataAsString("UTF-8");
      raakaTeksti = html.replace(/<[^>]*>/g, " ");
    }
    // 4. PDF- ja Word-tiedostot (.pdf, .doc, .docx)
    else if (
      mimeType === MimeType.PDF || 
      mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || 
      mimeType === "application/msword"
    ) {
      raakaTeksti = luePdfTaiWordTeksti(tiedosto);
    }
    // 5. UUSI: Google Sheets (.gsheet)
    else if (mimeType === MimeType.GOOGLE_SHEETS) {
      raakaTeksti = lueSheetsTeksti(tiedosto.getId());
    }
    // 6. UUSI: Google Slides (.gslides)
    else if (mimeType === MimeType.GOOGLE_SLIDES) {
      raakaTeksti = lueSlidesTeksti(tiedosto.getId());
    }
    // 7. UUSI: Excel (.xlsx) ja PowerPoint (.pptx) -> Muunnetaan Driven kautta teksti-OCR/Docs-muotoon
    else if (
      mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      mimeType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ) {
      raakaTeksti = lueMuunnettuDriveTeksti(tiedosto);
    }
    else {
      return "[Ei-tuettu tiedostomuoto: " + tiedosto.getName() + " (" + mimeType + ")]";
    }

    // Ajetaan siivous kaikelle onnistuneesti luetulle tekstille!
    return siivoaDriveTeksti(raakaTeksti);

  } catch (err) {
    Logger.log("Virhe tiedoston '" + tiedosto.getName() + "' lukemisessa: " + err.message);
    return "[Virhe tiedoston lukemisessa: " + err.message + "]";
  }
}

// Lukee Google Sheets -taulukon kaikki solut tekstiksi
function lueSheetsTeksti(id) {
  let ss = SpreadsheetApp.openById(id);
  let sheets = ss.getSheets();
  let tekstit = [];

  sheets.forEach(sheet => {
    let data = sheet.getDataRange().getValues();
    tekstit.push("--- Taulukko: " + sheet.getName() + " ---");
    data.forEach(rivi => {
      tekstit.push(rivi.join("\t")); // Erotetaan solut tabulaattorilla
    });
  });

  return tekstit.join("\n");
}

// Lukee Google Slides -esityksen kaikkien diojen tekstikentät
function lueSlidesTeksti(id) {
  let presentation = SlidesApp.openById(id);
  let slides = presentation.getSlides();
  let tekstit = [];

  slides.forEach((slide, index) => {
    tekstit.push("\n--- Dia " + (index + 1) + " ---");
    let pageElements = slide.getPageElements();
    pageElements.forEach(element => {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        let shape = element.asShape();
        if (shape.getText()) {
          let txt = shape.getText().asString().trim();
          if (txt) tekstit.push(txt);
        }
      }
    });
  });

  return tekstit.join("\n");
}

// Lukee .xlsx ja .pptx tiedostot kääntämällä ne väliaikaisesti Google Docs -muotoon
function lueMuunnettuDriveTeksti(tiedosto) {
  let resource = {
    title: "Temp_Doc_" + tiedosto.getName()
  };

  let tempDoc = Drive.Files.insert(resource, tiedosto.getBlob(), { convert: true });
  let teksti = "";

  if (tempDoc.mimeType === MimeType.GOOGLE_DOCS) {
    teksti = DocumentApp.openById(tempDoc.id).getBody().getText();
  } else if (tempDoc.mimeType === MimeType.GOOGLE_SHEETS) {
    teksti = lueSheetsTeksti(tempDoc.id);
  } else if (tempDoc.mimeType === MimeType.GOOGLE_SLIDES) {
    teksti = lueSlidesTeksti(tempDoc.id);
  }

  // Poistetaan väliaikainen tiedosto
  Drive.Files.remove(tempDoc.id);

  return teksti;
}

/**
 * Apufunktio: Poistaa Drive-teksteistä tyhjät rivit, turhat välilyönnit ja tiivistää peräkkäiset rivinvaihdot.
 */
function siivoaDriveTeksti(teksti) {
  if (!teksti) return "";
  
  return teksti
    // 1. Poistetaan rivit, joilla on vain pelkkiä välilyöntejä tai tabulaattoreita (estää "tyhjät" rivit)
    .replace(/^[ \t]+$/gm, '')
    // 2. Poistetaan samalla rivillä olevat tuplavälilyönnit
    .replace(/[ \t]{2,}/g, ' ')
    // 3. Tiivistetään 3 tai useampi peräkkäinen rivinvaihto enintään kahdeksi rivinvaihdoksi (\n\n)
    .replace(/(\r\n|\n|\r){3,}/g, '\n\n')
    .trim();
}

/**
 * Lukee PDF- ja Word-tiedostot muuntamalla ne väliaikaisesti Google Dociksi
 * ja palauttamalla sisällön puhtaana tekstinä.
 */
function luePdfTaiWordTeksti(tiedosto) {
  let tempFileId = null;
  let mimeType = tiedosto.getMimeType();
  let blob = tiedosto.getBlob();

  // Jos tiedosto on taustaltaan jo Google Doc
  if (blob.getContentType() === "application/vnd.google-apps.document") {
    try {
      return DocumentApp.openById(tiedosto.getId()).getBody().getText();
    } catch (e) {}
  }

  try {
    let resource = {
      title: "[TEMP] " + tiedosto.getName()
    };

    let tempFile = null;

    // YRITYS 1: Normaali konversio
    try {
      tempFile = Drive.Files.insert(resource, blob, { convert: true });
    } catch (err) {
      // YRITYS 2: OCR-tunnistus skannatuille PDF-tiedostoille
      if (mimeType === MimeType.PDF || mimeType === "application/pdf") {
        blob = blob.setContentType("application/pdf");
        tempFile = Drive.Files.insert(resource, blob, { convert: true, ocr: true, ocrLanguage: "fi" });
      } else {
        throw err;
      }
    }

    tempFileId = tempFile.id;

    // Luetaan väliaikaisesta Google Docista PELKKÄ TEKSTI (ei HTML-koodia)
    let teksti = DocumentApp.openById(tempFileId).getBody().getText();

    // Poistetaan väliaikainen tiedosto Drivesta
    DriveApp.getFileById(tempFileId).setTrashed(true);

    return teksti || "[Tiedosto oli tyhjä tai tekstiä ei pystytty tunnistamaan]";

  } catch (e) {
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (subErr) {}
    }
    Logger.log("Virhe PDF/Word-konversiossa (" + tiedosto.getName() + "): " + e.toString());
    return "[PDF/Word-tiedoston tekstin lukeminen epäonnistui: " + e.toString() + "]";
  }
}

/**
 * Irrottaa ID-merkkijonon Google Drive / Docs / Sheets / Slides -osoitteista
 * tai palauttaa syötteen sellaisenaan, jos se on jo valmis ID.
 */
function haeKansioIdUrlsta(url) {
  if (!url) return null;
  
  let syote = url.toString().trim();

  // 1. Jos syöte on jo pelkkä ID (ei sisällä vinoviivoja tai kysymysmerkkejä)
  if (!syote.includes('/') && !syote.includes('?') && syote.length >= 25) {
    return syote;
  }
  
  // 2. Haetaan osoitemuodoista /d/ID, /folders/ID tai /file/d/ID
  let match = syote.match(/\/(?:folders|d|file\/d)\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1]; // Palautetaan pelkkä ryhmä 1 (ilman /d/ tai /folders/ -osaa)
  }
  
  // 3. Haetaan hakuparametreista, esim. ?id=ID tai &id=ID
  let idMatch = syote.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) {
    return idMatch[1];
  }

  // 4. Varalla: napataan purettu vähintään 25 merkkiä pitkä merkkijono
  let fallbackMatch = syote.match(/([a-zA-Z0-9_-]{25,})/);
  return fallbackMatch ? fallbackMatch[1] : null;
}

/**
 * Varmistaa, että taulukon ensimmäisellä rivillä on otsikot ja hyvät muotoilut.
 */
function alustaTaulukkoTarvittaessa(sheet) {
  // Alustetaan vain jos taulukko on tyhjä
  if (sheet.getLastRow() === 0) {
    let otsikot = [["Aikaleima", "Hakemus luonnoksen linkki", "Drive linkki", "Hakemuksen URL", "docId", "docUrl"]];
    let otsikkoAlue = sheet.getRange(1, 1, 1, otsikot[0].length);

    // Lisätään teksti, lihavointi ja taustaväri
    otsikkoAlue.setValues(otsikot);
    otsikkoAlue.setFontWeight("bold");
    otsikkoAlue.setBackground("#f3f3f3");
    
    // Asetetaan leikkaus (CLIP) kaikille sarakkeille 1–6 otsikoista taulukon pohjaan asti.
    // Tämän ansiosta myös tulevat datarivit leikkaantuvat automaattisesti!
    let kokoAlue = sheet.getRange(1, 1, sheet.getMaxRows(), otsikot[0].length);
    kokoAlue.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Sovitetaan sarakkeiden leveydet automaattisesti tekstin mukaan
    sheet.setColumnWidth(1, 160); // Aikaleima
    sheet.setColumnWidth(2, 220); // Hakemus luonnoksen linkki
    sheet.setColumnWidth(3, 220); // Drive linkki
    sheet.setColumnWidth(4, 250); // Hakemuksen URL
    sheet.setColumnWidth(5, 180); // docId
    sheet.setColumnWidth(6, 300); // docUrl

    suojaaValilehti(sheet);
  }
}

/**
 * Apufunktio: Lukitsee välilehden niin, ettei siihen voi vahingossa kirjoittaa käyttöliittymästä.
 */
function suojaaValilehti(sheet) {
  // Tarkistetaan, onko välilehdellä jo suojaus päällä
  let suojaukset = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  
  if (suojaukset.length === 0) {
    let suojaus = sheet.protect().setDescription("Automatisoitu välilehti - Lukittu vahinkokirjoituksilta");
    
    // Poistetaan muilta käyttäjiltä muokkausoikeudet
    let me = Session.getEffectiveUser();
    suojaus.addEditor(me);
    suojaus.removeEditors(suojaus.getEditors());
    
    // Jos olet ainoa muokkaaja, asetetaan keltainen varoitus ("Warning-only"),
    // joka estää vahingossa kirjoittamisen kysymällä vahvistusta.
    if (suojaus.canDomainEdit()) {
      suojaus.setDomainEdit(false);
    }
    
    // Varmistetaan että skripti/käyttäjä saa keltaisen "Haluatko varmasti muokata?" -varoitusikkunan käyttöliittymässä
    suojaus.setWarningOnly(true); 
  }
}

/**
 * Varmistaa, että taulukon ensimmäisellä rivillä on otsikot ja hyvät muotoilut.
 */
function alustaTaulukkoHistoria(sheet) {
  // Alustetaan vain jos taulukko on tyhjä
  if (sheet.getLastRow() === 0) {
    let otsikot = [["Kellonaika","Käyttäjä", "Agentti", "Malli", "Hinta-arvio", "Syöte", "Cached", "Vastaus", "Ajattelu", "Yht.", "←Tokenit"]];
    let otsikkoAlue = sheet.getRange(1, 1, 1, otsikot[0].length);

    // Lisätään teksti, lihavointi ja taustaväri
    otsikkoAlue.setValues(otsikot);
    otsikkoAlue.setFontWeight("bold");
    otsikkoAlue.setBackground("#f3f3f3");

    // Asetetaan automaattinen tekstin rivitys sarakkeille B (Käyttäjä) ja C (Agentti)
    //getRange("B:C") kattaa sarakkeet 2 ja 3 koko taulukon korkeudelta
    sheet.getRange("B:C").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Sovitetaan sarakkeiden leveydet automaattisesti tekstin mukaan
    sheet.setColumnWidth(1, 160); // Kellonaika
    sheet.setColumnWidth(2, 350); // Käyttäjä
    sheet.setColumnWidth(3, 550); // Agentti
    sheet.setColumnWidth(4, 160); // Malli
    sheet.setColumnWidth(5, 100); //Hinta
    sheet.setColumnWidth(6, 70); //Syöte
    sheet.setColumnWidth(7, 70); //Cache
    sheet.setColumnWidth(8, 70); //Vastaus
    sheet.setColumnWidth(9, 70); //Ajattelu
    sheet.setColumnWidth(10, 70); //Yht

    suojaaValilehti(sheet);
  }
}

/**
 * Tarkistaa, onko tietyssä solussa arvoa.
 * @return {object} Tieto solun tilasta
 */
function tarkistaSolut() {
  const ss = haeKayttajanSheet(); // Tai .getSheetByName("TaulukonNimi")
  const sheet = ss.getSheetByName("Parametrit");

  const arvot = sheet.getRange(2, 2, 1, 3).getValues()[0];

  const onArvoa = (val) => val !== "" && val !== null && val !==undefined;
  
  return {
    B2: onArvoa(arvot[0]),
    C2: onArvoa(arvot[1]),
    D2: onArvoa(arvot[2])
  };
}



function kutsuGemini(payload, apiKey, valittuMalli) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${valittuMalli}:generateContent?key=${apiKey}`;
  Logger.log(`Yritetään mallia ${valittuMalli}`);

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode === 200) {
      const json = JSON.parse(responseText);
      const parts = json.candidates?.[0]?.content?.parts || [];
      
      // Etsitään ensimmäinen tekstiosa, jos responssissa on useampia part-olioita
      const textPart = parts.find(p => p.text !== undefined);
      
      if (!textPart || !textPart.text) {
        throw new Error(`Malli ${valittuMalli} palautti tyhjän vastauksen.`);
      }

      return json;
    }

    // Jos HTTP statuskoodi ei ollut 200, yritetään poimia virheviesti Geminin vastauksesta
    let virheViesti = responseText;
    try {
      const errJson = JSON.parse(responseText);
      if (errJson.error && errJson.error.message) {
        virheViesti = errJson.error.message;
      }
    } catch (e) {
      // Ei ollut JSON-muotoinen virhe, käytetään raakatekstiä
    }

    Logger.log(`⚠️ Virhe mallilla ${valittuMalli} (HTTP ${statusCode}): ${responseText}`);
    throw new Error(`Gemini API -virhe (${statusCode}): ${virheViesti}`);

  } catch (error) {
    Logger.log(`⚠️ Verkko- tai suoritusvirhe: ${error.toString()}`);
    // Heitetään virhe eteenpäin, jotta koko suoritus pysähtyy!
    throw new Error(`Gemini-kutsu epäonnistui: ${error.message}`);
  }
}


function haeTiedotSivulta(startUrl, maxSubLinks = 8) {
  let koottuAineisto = "";

  // Asetukset alalinkkien tekstille (riisutaan kuvat ja linkkikokoelmat)
  const jinaTextOptions = {
    muteHttpExceptions: true,
    headers: {
      "X-Return-Format": "text",
      "X-With-Images-Summary": "false",
      "X-With-Links-Summary": "false"
    }
  };

  try {
    // 1. Haetaan pääsivu (pidetään mukana myös valikot, jotta linkit löytyvät)
    Logger.log("Haetaan pääsivua: " + startUrl);
    const mainResponse = UrlFetchApp.fetch("https://r.jina.ai/" + startUrl, { muteHttpExceptions: true });

    if (mainResponse.getResponseCode() !== 200) {
      throw new Error(`Statuskoodi ${mainResponse.getResponseCode()}`);
    }

    const rawMainText = mainResponse.getContentText();

    if (!rawMainText || rawMainText.includes("Invalid URL") || rawMainText.startsWith("Error:")) {
      throw new Error("Sivuun ei päästä käsiksi.");
    }

    // Lisätään pääsivun teksti puhdistettuna
    koottuAineisto += `=== PÄÄSIVU: ${startUrl} ===\n\n${siivoaTeksti(rawMainText)}\n\n`;

    // 2. Poimitaan KAIKKI linkit ALKUPERÄISESTÄ raakatestistä (ennen siivousta)
    const urlRegex = /https?:\/\/[^\s\)\"\'>]+/g;
    const loydetytLinkit = rawMainText.match(urlRegex) || [];

    const hylattavatPaatteet = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.css', '.js', '.json', '.ico', '.zip'];

    const valitutLinkit = [...new Set(loydetytLinkit)]
      .filter(link => {
        const osat = link.toLowerCase().split('?')[0];
        const onkoMedia = hylattavatPaatteet.some(paate => osat.endsWith(paate));

        return link !== startUrl && 
               !onkoMedia &&
               !link.includes("facebook.com") && 
               !link.includes("twitter.com") &&
               !link.includes("linkedin.com") &&
               !link.includes("x.com");
      })
      .slice(0, maxSubLinks);

    Logger.log(`Löytyi ${loydetytLinkit.length} linkkiä raw-datasta. Valittiin ${valitutLinkit.length} alalinkkiä.`);

    // 3. Haetaan valitut alalinkit kevyinä teksteinä
    valitutLinkit.forEach((subUrl, index) => {
      try {
        const onkoPdf = subUrl.toLowerCase().split('?')[0].endsWith('.pdf');
        const tyyppiTeksti = onkoPdf ? "PDF" : "SIVU";

        Logger.log(`Haetaan ${tyyppiTeksti} (${index + 1}/${valitutLinkit.length}): ${subUrl}`);

        const subResponse = UrlFetchApp.fetch("https://r.jina.ai/" + subUrl, jinaTextOptions);
        let subText = siivoaTeksti(subResponse.getContentText());

        if (subText.length > 50) {
          koottuAineisto += `\n--- ALALINKKI ${index + 1} [${tyyppiTeksti}]: ${subUrl} ---\n\n${subText}\n\n`;
        }
      } catch (err) {
        Logger.log(`Virhe linkissä (${subUrl}): ` + err.toString());
      }
    });

  } catch (error) {
    Logger.log("Virhe: " + error.toString());
    throw new Error("Pääsivun haku epäonnistui: " + error.message);
  }

  return koottuAineisto;
}

/**
 * Puhdistaa tekstistä turhan syntaksin, Markdown-kuvat ja toistuvat rivinvaihdot.
 */
function siivoaTeksti(teksti) {
  if (!teksti) return "";
  
  return teksti
    .replace(/<[^>]*>/g, '')                  // Poistetaan HTML-tägit
    .replace(/!\[.*?\]\(.*?\)/g, '')          // Poistetaan kuvat
    .replace(/\[(.*?)\]\(.*?\)/g, '$1')       // Muutetaan Markdown-linkit pelkäksi tekstiksi
    .replace(/^#+\s+/gm, '')                  // Poistetaan otsikkosymbolit
    .replace(/(\r\n|\n|\r){3,}/g, '\n\n')     // Tiivistetään moninkertaiset tyhjät rivit
    .replace(/[ \t]{2,}/g, ' ')               // Tiivistetään tuplavälilyönnit
    .trim();
}

function tallennaAineistotGoogleDocsiin(aineisto1, aineisto2, aineisto3, tiedostonNimi = "Gemini_Prompt_Aineisto") {
  try {
    let rootFolder = DriveApp.getRootFolder();
    
    // 1. Etsitään ja poistetaan vanha saman niminen tiedosto, jos sellainen on jo olemassa
    let vanhatTiedostot = rootFolder.getFilesByName(tiedostonNimi);
    while (vanhatTiedostot.hasNext()) {
      let vanhaTiedosto = vanhatTiedostot.next();
      vanhaTiedosto.setTrashed(true); // Siirretään vanha roskakoriin
      Logger.log("Poistettiin vanha samanniminen tiedosto: " + tiedostonNimi);
    }

    // 2. Luodaan uusi dokumentti
    let doc = DocumentApp.create(tiedostonNimi);
    let body = doc.getBody();

    body.appendParagraph("=========================================");
    body.appendParagraph("OSA 1: DRIVE-TIEDOSTOJEN SISÄLTÖ");
    body.appendParagraph("=========================================").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(aineisto1 || "[Ei sisältöä]");

    body.appendPageBreak();

    body.appendParagraph("=========================================");
    body.appendParagraph("OSA 2: WEB-SIVUSTON SISÄLTÖ");
    body.appendParagraph("=========================================").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(aineisto2 || "[Ei sisältöä]");

    body.appendParagraph("=========================================");
    body.appendParagraph("OSA 3: ARVIOITAVA LUONNOS HAKEMUKSESTA");
    body.appendParagraph("=========================================").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(aineisto3 || "[Ei hakemuspohjaa]");

    doc.saveAndClose();

    Logger.log("Päivitetty uusi tiedosto Juureen: " + doc.getUrl());
    return doc.getUrl();

  } catch (err) {
    Logger.log("Virhe tallennuksessa: " + err.toString());
    throw new Error("Tallennus epäonnistui: " + err.message);
  }
}

function lueHakemus(pohjaDocUrl) {
  if (!pohjaDocUrl || pohjaDocUrl.trim() === "") {
    throw new Error("Hakemusluonnoksen Osoite/URL puuttuu.");
  }

  // Haetaan ID URL-osoitteesta
  const match = pohjaDocUrl.match(/[-\w]{25,}/);
  const pohjaDocId = match ? match[0] : null;

  if (!pohjaDocId) {
    throw new Error("Tiedosto-ID:tä ei pystytty erottamaan osoitteesta: " + pohjaDocUrl);
  }

  let tempFileId = null;

  try {
    let docIdToOpen = pohjaDocId;
    const tiedosto = DriveApp.getFileById(pohjaDocId);
    const mimeType = tiedosto.getMimeType();
    const blob = tiedosto.getBlob();

    // Jos kyseessä EI OLE valmiiksi Google Docs -tiedosto (vaan esim. PDF tai Word)
    const isGoogleDoc = mimeType === MimeType.GOOGLE_DOCS || blob.getContentType() === "application/vnd.google-apps.document";

    if (!isGoogleDoc) {
      let resource = { title: "[TEMP] " + tiedosto.getName() };
      let tempFile = null;

      // YRITYS 1: Normaali muunnos Word- ja PDF-tiedostoille
      try {
        tempFile = Drive.Files.insert(resource, blob, { convert: true });
      } catch (err) {
        // YRITYS 2: OCR-tunnistus skannatuille PDF-tiedostoille
        if (mimeType === MimeType.PDF || mimeType === "application/pdf") {
          let ocrBlob = blob.setContentType("application/pdf");
          tempFile = Drive.Files.insert(resource, ocrBlob, { convert: true, ocr: true, ocrLanguage: "fi" });
        } else {
          throw err;
        }
      }

      tempFileId = tempFile.id;
      docIdToOpen = tempFileId;
    }

    // Avataan Google Docs (joko alkuperäinen tai väliaikaisesti muunnettu)
    let doc = DocumentApp.openById(docIdToOpen);
    let body = doc.getBody();
    let teksti = "";

    // Käydään elementit läpi säilyttäen taulukkomuotoilut
    let numElements = body.getNumChildren();
    for (let i = 0; i < numElements; i++) {
      let element = body.getChild(i);
      let type = element.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        teksti += element.asParagraph().getText() + "\n";
      } 
      else if (type === DocumentApp.ElementType.TABLE) {
        let table = element.asTable();
        for (let r = 0; r < table.getNumRows(); r++) {
          let row = table.getRow(r);
          let rowText = [];
          for (let c = 0; c < row.getNumCells(); c++) {
            rowText.push(row.getCell(c).getText().trim());
          }
          teksti += "| " + rowText.join(" | ") + " |\n";
        }
        teksti += "\n";
      }
    }

    return teksti;

  } catch (e) {
    throw new Error("Virhe hakemuksen lukemisessa: " + e.message);
  } finally {
    // Poistetaan väliaikainen tiedosto AINA (onnistuipa lukeminen tai ei)
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (err) {}
    }
  }
}


/**
 * Tallentaa keskustelun tiedot Google Sheetsiin.
 */
function tallennaGeminiKeskustelu(kysymys, apiVastausJson, malli) {
  // 1. Valitaan aktiivisen taulukon ensimmäinen välilehti
  const ss = haeKayttajanSheet();
  const sheet = ss.getSheets()[0];
  
  // 2. Määritetään kellonaika
  const kellonaika = new Date();
  
  // 3. Poimitaan vastauksen teksti ja token-määrä API-vastauksesta
  const vastausTeksti = apiVastausJson.candidates[0].content.parts[0].text;
  
  // Poimitaan kokonaistokenit ja Cached tokenit(ja varmistetaan ettei koodi kaadu, jos tieto puuttuu)
  const cachedTokenit = apiVastausJson.usageMetadata?.cachedContentTokenCount || 0;
  const tokenit = apiVastausJson.usageMetadata ? apiVastausJson.usageMetadata.totalTokenCount : "Ei tietoa";
  const syotteenTokenit = apiVastausJson.usageMetadata ? apiVastausJson.usageMetadata.promptTokenCount : "Ei tietoa";
  const vastausTokenit = apiVastausJson.usageMetadata ? apiVastausJson.usageMetadata.candidatesTokenCount : "Ei tietoa";
  const ajatteluTokenit = apiVastausJson.usageMetadata ? apiVastausJson.usageMetadata.thoughtsTokenCount : "Ei tietoa";

  let kaikkiVastausTokenit = typeof vastausTokenit === "number" ? vastausTokenit : 0;

  // Jos ajattelutokenit on numeerinen tieto, lisätään se mukaan
  if (typeof ajatteluTokenit === "number") {
    kaikkiVastausTokenit += ajatteluTokenit;
  }

  // Lasketaan arvioitu hinta euroina
  const arvioEUR = laskeKustannusEUR(
    malli, 
    syotteenTokenit, 
    kaikkiVastausTokenit, 
    cachedTokenit
  );
  
  // 4. Lisätään rivi taulukkoon
  sheet.appendRow([
    kellonaika,
    kysymys,
    vastausTeksti,
    malli,      // Arvo tulee suoraan käyttöliittymästä
    arvioEUR,
    syotteenTokenit,
    cachedTokenit,
    vastausTokenit,
    ajatteluTokenit,
    tokenit
  ]);
}

/**
 * Haetaan aiemmat viestit Sheetsistä ja lisätään uusin viesti perään.
 */
function haeAiempiHistoria(sheet, uusinViesti) {
  const lastRow = sheet.getLastRow();
  const contents = [];

  // 1. Luetaan aiempi historia taulukosta (riviltä 2 alkaen)
  if (lastRow > 1) {
    const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();

    data.forEach(rivi => {
      const vanhaKayttaja = rivi[0];
      const vanhaAgentti = rivi[1];

      if (vanhaKayttaja) {
        contents.push({
          role: "user",
          parts: [{ text: String(vanhaKayttaja) }]
        });
      }

      if (vanhaAgentti) {
        contents.push({
          role: "model",
          parts: [{ text: String(vanhaAgentti) }]
        });
      }
    });
  }

  // 2. Lisätään uusin viesti muuttujasta taulukon loppuun
  if (uusinViesti) {
    contents.push({
      role: "user",
      parts: [{ text: String(uusinViesti) }]
    });
  }

  return contents;
}

function haeKeskusteluhistoriaUI(sheet) {
  // Jos sheet-oliota ei välitetä, haetaan aktiivinen taulukko
  if (!sheet) {
    const ss = haeKayttajanSheet();
    sheet = ss.getSheets()[0];
  }

  const lastRow = sheet.getLastRow();
  const historia = [];

  if (lastRow > 1) {
    // Haetaan riviltä 2 alkaen 2 saraketta (Käyttäjä ja Agentti)
    const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();

    data.forEach(rivi => {
      const kayttajaTeksti = rivi[0];
      const agenttiTeksti = rivi[1];

      if (kayttajaTeksti) {
        historia.push({ rooli: 'user', teksti: String(kayttajaTeksti) });
      }
      if (agenttiTeksti) {
        historia.push({ rooli: 'model', teksti: String(agenttiTeksti) });
      }
    });
  }

  return historia;
}

/**
 * Tyhjentää sekä keskusteluhistorian (välilehti 1) että parametrit (välilehti 2).
 * Säilyttää kummatkin otsikkorivit (rivi 1).
 */
function tyhjennaKaikkiPohjadata() {
  const ss = haeKayttajanSheet();
  const sheets = ss.getSheets();

  // 1. TYHJENNETÄÄN VÄLILEHTI 1 (Keskusteluhistoria)
  const sheet1 = sheets[0];
  const lastRow1 = sheet1.getLastRow();
  if (lastRow1 > 1) {
    sheet1.getRange(2, 1, lastRow1 - 1, sheet1.getLastColumn()).clearContent();
  }

  // 2. TYHJENNETÄÄN VÄLILEHTI 2 (Parametrit / Taustadata)
  if (sheets.length > 1) {
    const sheet2 = sheets[1];
    const lastRow2 = sheet2.getLastRow();
    if (lastRow2 > 1) {
      sheet2.getRange(2, 1, lastRow2 - 1, sheet2.getLastColumn()).clearContent();
    }
  } else {
    Logger.log("Varoitus: Toista välilehteä ei löytynyt!");
  }

  return "Keskusteluhistoria ja parametrit tyhjennetty onnistuneesti.";
}

/**
 * Palauttaa luodun Sheetsin URL-osoitteen käyttöliittymälle (Index.html),
 * jotta käyttäjä voi avata taulukon omalle välilehdelleen.
 */
function haeTaulukonUrl() {
  const ss = haeKayttajanSheet();
  return ss.getUrl();
}

function haeKayttajanSheet() {
  const userProps = PropertiesService.getUserProperties();
  const sheetId = userProps.getProperty('KAYTTAJAN_SHEET_ID');
  
  if (!sheetId) {
    throw new Error("Taulukkoa ei löytynyt. Avaa sovellus uudelleen.");
  }
  
  return SpreadsheetApp.openById(sheetId);
}

// Virallinen hinnasto per 1 000 000 tokenia ($ USD)
// Lähde: ai.google.dev/gemini-api/docs/pricing
const MODEL_PRICES = {
  "gemini-3.5-flash-lite": { prompt: 0.30, candidate: 2.50, cached: 0.03 },
  "gemini-3.6-flash":      { prompt: 1.50, candidate: 7.50, cached: 0.15 },
  "gemini-3.8-flash":      { prompt: 1.50, candidate: 7.50, cached: 0.15 }
};

// Viitekurssi Euroissa (esim. 1 USD = 0.92 EUR)
const USD_TO_EUR = 0.92;

function laskeKustannusEUR(valittuMalli, promptTokens, candidateTokens, cachedTokens, usdToEurRate = USD_TO_EUR) {
  // Haetaan valitun malli hinnat (oletuksena 3.6-flash)
  const hinnat = MODEL_PRICES[valittuMalli] || MODEL_PRICES["gemini-3.6-flash"];

  // VÄHENTÄMINEN: Lasketaan vain UUDET (ei-välimuistitetut) syötetokenit normihinnalla
  const uudetPromptTokens = Math.max(0, promptTokens - cachedTokens);

  // Lasketaan osioiden hinnat USD ($ per 1M tokenia)
  const promptHintaUSD    = (uudetPromptTokens * hinnat.prompt) / 1000000;
  const cachedHintaUSD    = (cachedTokens * hinnat.cached) / 1000000;
  const candidateHintaUSD = (candidateTokens * hinnat.candidate) / 1000000;

  const kokonaishintaUSD = promptHintaUSD + cachedHintaUSD + candidateHintaUSD;

  // Muunnetaan euroiksi
  const kokonaishintaEUR = kokonaishintaUSD * usdToEurRate;

  if (kokonaishintaEUR > 0 && kokonaishintaEUR < 0.00001) {
    return "< 0,00001 €";
  }

  return `${kokonaishintaEUR.toFixed(6).replace('.', ',')} €`;
}

// Funktio, jota selain kyselee säännöllisesti
function haeSuorituksenTila() {
  return PropertiesService.getUserProperties().getProperty("STATUS") || "Käynnistetään...";
}