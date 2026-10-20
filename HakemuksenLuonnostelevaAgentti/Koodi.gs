/**
 * Luo Google Sheetsiin mukautetun ylävalikon sovelluksen avaimelle.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Luonnos Agentti')
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
      .setTitle('Hakemuksen luonnosteleva agentti');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Avaa popup-ikkunan (modal dialog).
 */
function avaaIkkuna() {
  let html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setWidth(700)
      .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
  .setTitle('Hakemuksen luonnosteleva agentti')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

/**
 * Pääfunktio: Käsitellään käyttöliittymästä lähetetyt tiedot ja tallennetaan ne taulukkoon.
 * @param {Object} lomakeData - HTML-sivulta lähetetyt kenttien arvot
 * @return {Object} Tieto ajon onnistumisesta käyttöliittymälle
 */
function AgenttiLuoHakemuksen(lomakeData) {

  // Haetaan API-avain
  const props = PropertiesService.getScriptProperties();
  const GEMINI_API_KEY = props.getProperty("GEMINI_API_KEY");

  const userProps = PropertiesService.getUserProperties();

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0]; // Valitsee ensimmäisen välilehden

  alustaTaulukkoTarvittaessa(sheet);

  // Poimitaan arvot turvallisesti (tyhjä merkkijono, jos kenttä oli tyhjä)
  let ohjeet = lomakeData && lomakeData.ohjeet ? lomakeData.ohjeet : "";
  let linkki = lomakeData && lomakeData.linkki ? lomakeData.linkki : "";
  let kansioUrl = lomakeData && lomakeData.kansioUrl ? lomakeData.kansioUrl : "";
  let hakemusTunnus = lomakeData && lomakeData.hakemusUrl ? lomakeData.hakemusUrl : "";
  let malli = lomakeData && lomakeData.malli ? lomakeData.malli : "";

  Logger.log("--- UUSI AJO KÄYNNISTETTY ---");
  Logger.log("Ohjeet: " + ohjeet);
  Logger.log("Linkki: " + linkki);
  Logger.log("Kansio URL: " + kansioUrl);
  Logger.log("Hakemuksen tunnus: " + hakemusTunnus);
  Logger.log("Gemini malli: " + malli);

  userProps.setProperty("STATUS", "📂 Luetaan hakemuspohjaa...");

  let hakemusPohja = lueTiedostoIdlla(hakemusTunnus)
  Logger.log("Haettu teksti: " + hakemusPohja);

  // Jos kansion URL on syötetty mutta se on virheellinen
  if (kansioUrl && !haeKansioIdUrlsta(kansioUrl)) {
    throw new Error("Google Drive -kansion URL oli virheellinen!");
  }

  userProps.setProperty("STATUS", "📂 Luetaan lähdemateriaalia Drive kansiosta...");

  // Luetaan kansion sisältö vain, jos kansion osoite on annettu
  let kansionSisalto = [];
  if (kansioUrl) {
    kansionSisalto = lueDriveKansionSisalto(kansioUrl);
  }

  Logger.log("Löytyneitä tiedostoja: " + kansionSisalto.length);

  userProps.setProperty("STATUS", "📝 Selataan läpi hankkeen sivustoa ja sen alalinkkejä...");

  let hakemuksenTiedot;

    try {
      hakemuksenTiedot = haeTiedotSivulta(linkki, 8);
    } catch (e) {
      // Jos haeTiedotSivulta-funktio heittää jonkin oman virheen
      throw new Error("Hakemuksen tietojen hakeminen epäonnistui: " + e.message);
    }

    // Tarkistetaan, jäikö muuttuja tyhjäksi (null, undefined, tyhjä merkkijono jne.)
    if (!hakemuksenTiedot) {
      throw new Error("Hakemuksen tiedot jäivät tyhjiksi!");
    }

  Logger.log("Haettu sivusto: " + hakemuksenTiedot);

  // Kirjoitetaan tiedot taulukon riville 2
  sheet.getRange(2, 1).setValue(new Date()); // Aikaleima
  sheet.getRange(2, 2).setValue(ohjeet);     // Tarkentavat ohjeet
  sheet.getRange(2, 3).setValue(linkki);     // Hakemuksen URL
  sheet.getRange(2, 4).setValue(kansioUrl);  // Drive-kansion linkki
  sheet.getRange(2, 5).setValue(hakemusTunnus);  // Hakemus pohjan Drive tunnus
  sheet.getRange(2, 6).setValue(malli);  // Valittu gemini kielimalli


  let tilaTeksti = kansioUrl 
    ? kansionSisalto.length + " tiedostoa luettu" 
    : "Ei kansion osoitetta";

  sheet.getRange(2, 8).setValue(tilaTeksti); // Kansion tila

  // Varmistetaan, että tiedot päivittyvät taulukkoon välittömästi
  SpreadsheetApp.flush();

// 1. SYSTEM INSTRUCTION (Rooli + Laatuvaatimukset + Tiukat HTML-muotoilusäännöt)
const systemInstruction = {
  parts: [
    {
      text: `Olet kokenut EU-hankevalmistelija ja EURA 2021 -järjestelmän asiantuntija. Tehtäväsi on laatia eOppimiskeskukselle (Suomen eOppimiskeskus ry) erittäin laadukas, ammattimainen ja rahoituskriteerit täyttävä esitäytetty hankehakemus.

Sovelletut periaatteet asiasisällön tuottamisessa:
1. VAIKUTTAVUUS JA TARVE: Perustele hankkeen tarve selkeästi nojaten taustamateriaaleihin ja hakuilmoituksen painopisteisiin.
2. HYÖDYNNÄ TARKASTI LÄHTEITÄ: Synteesin tärkeysjärjestys on:
   - 1. Käyttäjän antamat täsmentävät ohjeet
   - 2. Haettavan hankkeen/rahoitushun kriteerit ja teemat (Verkkosivut)
   - 3. Organisaation ja hankeidean taustamateriaalit (Drive-kansio)
3. AMMATTIMAINEN EU-KIELI: Käytä selkeää, vakuuttavaa ja tulosorientoitunutta asiantuntijakieltä (esim. ESR+/EAKR-terminologiaa).
4. PUUTTUVA TIETO JA OLETUKSET: Jos tausta-aineistosta puuttuu tarkka tieto (esim. tietty euroalueen budjetti, Y-tunnus tai vastuuhenkilö), täytä kohta parhaan mahdollisen oletuksen mukaan ja merkitse selkeästi syötteeseen: [TÄYDENNÄ: puuttuva tieto / perustelu]. ÄLÄ käytä sanaa "arvaus".

---
TEKNISEN RENDERÖINNIN JA HTML-MUOTOILUN SÄÄNNÖT:
Palauta täytetty hakemus täysin PUHTAANA HTML-koodina ilman markdown-koodilohkoja (esim. ilman \`\`\`html -merkintöjä).

1. HTML-RAKENNE:
   - ÄLÄ lisää HTML-, HEAD- tai BODY-tageja. Palauta VAIN otsikot ja taulukot.
   - Kirjoita Ä- ja Ö-kirjaimet normaalisti. ÄLÄ käytä HTML-entiteettejä (&auml;).
   - Varmista, että jokainen avattu tagi (<td>, <tr>, <table>) suljetaan oikein.
   - ÄLÄ käytä inline-tyylejä (style="" tai class="").

2. OTSIKKO TYYLIT:
   - Pääotsikot: <h1>1 Otsikon nimi</h1>
   - Väliotsikot: <h2>1.1 Väliotsikon nimi</h2>

3. KENTTIEN JA TAULUKOIDEN RAKENNE:
   A) Yksittäiset vastauskentät:
   Luo aina 1-soluisena taulukkona. Kentän nimi ja vastaus SAMAN solun sisään.
   <table>
     <tr>
       <td>
         <p><b>Kentän nimi</b></p>
         <p>Kirjoita laadukas ja kattava vastaus tähän tekstiin...</p>
       </td>
     </tr>
   </table>

   B) Monisarakkeiset taulukot (esim. Toteuttajat, Budjetti):
   Taulukon yläpuolelle otsikko-p-tagi. Ensimmäiselle riville sarakkeiden lihavoidut otsikot.
   <p><b>Taulukon nimi</b></p>
   <table>
     <tr>
       <td><p><b>Sarake 1</b></p></td>
       <td><p><b>Sarake 2</b></p></td>
     </tr>
     <tr>
       <td><p>Sisältö 1</p></td>
       <td><p>Sisältö 2</p></td>
     </tr>
   </table>`
    }
  ]
};

// 2. USER PROMPT / CONTENTS (Dynamic Content)
const userPromptText = `Luo täytetty EURA 2021 -hankehakemus seuraavien tausta-aineistojen pohjalta:

---
1. KÄYTTÄJÄN OHJEET:
${ohjeet || "Ei erillisiä ohjeita."}

2. HAETTAVAN RAHOITUKSEN TIEDOT JA KRITEERIT (Verkkosivulta):
${hakemuksenTiedot || "Ei verkkosivutietoja saatavilla."}

3. LÄHDEMATERIAALI JA TAUSTATIEDOT (Drive-kansio):
${typeof kansionSisalto === 'object' ? JSON.stringify(kansionSisalto, null, 2) : (kansionSisalto || "Ei kansion tiedostoja.")}

4. TÄYTETTÄVÄ HAKEMUSPOHJA:
${hakemusPohja || "Ei erillistä hakemuspohjaa. Luo standardin EURA 2021 -hakemuksen mukainen rakenne taustamateriaalin pohjalta."}
---`;

// 3. KOKONAISUUTEENA LÄHETETTÄVÄ PAYLOAD
const payload = {
  systemInstruction: systemInstruction,
  contents: [
    {
      role: "user",
      parts: [
        {
          text: userPromptText
        }
      ]
    }
  ],
  generationConfig: {
    temperature: 0.2, // Matala temperature pitää mallin faktapohjaisena ja vähentää muotoiluvirheitä
    topP: 0.95,
    maxOutputTokens: 32000 // Varmistetaan, että pitkä hakemus ei katkea kesken
  }
};



  
  //const rawText = kutsuGeminiNeljasti(payload, GEMINI_API_KEY);
  //Logger.log("=== GEMININ VASTAUS ===");
  //Logger.log(rawText);

  userProps.setProperty("STATUS", "🤖 Lähetetään pyyntö Gemini-mallille...");

  const rawText = kutsuGemini(payload, GEMINI_API_KEY, malli);
  Logger.log("=== GEMININ VASTAUS ===");
  Logger.log(rawText);

  sheet.getRange(2, 7).setValue(rawText);
  SpreadsheetApp.flush();

  userProps.setProperty("STATUS", "📝 Kirjoitetaan hakemusluonnosta Driveen...");

  //haetaan drive kansion id ja päätetään luonnokselle nimi
  let driveID = haeKansioIdUrlsta(kansioUrl);
  let tiedostoNimi = luoTiedostonNimi();

  tallennaGeminivastaus(rawText, driveID, tiedostoNimi);

  return {
    onnistui: true,
    tiedostoja: kansionSisalto.length,
    vastaus: rawText,
    kansioUrl: kansioUrl
  };

}

/**
 * Hakee tietyn Google Drive -kansion kaikkien tiedostojen nimet, tyypit ja tekstisisällön.
 */
function lueDriveKansionSisalto(url) {
  let kansioId = haeKansioIdUrlsta(url);
  if (!kansioId) {
    Logger.log("Virheellinen Drive URL: " + url);
    return [];
  }

  try {
    let kansio = DriveApp.getFolderById(kansioId);
    let tiedostot = kansio.getFiles();
    let tulokset = [];

    while (tiedostot.hasNext()) {
      let tiedosto = tiedostot.next();
      let tekstiSisalto = lueTiedostonTekstiTurvallisesti(tiedosto);

      tulokset.push({
        nimi: tiedosto.getName(),
        id: tiedosto.getId(),
        url: tiedosto.getUrl(),
        tyyppi: tiedosto.getMimeType(),
        sisalto: tekstiSisalto
      });
    }

    Logger.log("Kansion tulokset: " + JSON.stringify(tulokset, null, 2));

    return tulokset;

  } catch (virhe) {
    Logger.log("Virhe kansion lukemisessa: " + virhe.toString());
    return [];
  }
}

function lueTiedostonTekstiTurvallisesti(tiedosto) {
  let mimeType = tiedosto.getMimeType();
  let fileId = tiedosto.getId();

  try {
    // 1. Google Docs (.gdoc)
    if (mimeType === MimeType.GOOGLE_DOCS || mimeType === "application/vnd.google-apps.document") {
      return DocumentApp.openById(fileId).getBody().getText();
    } 
    
    // 2. Google Sheets (.gsheet) - TÄRKEÄ KORJAUS
    if (mimeType === "application/vnd.google-apps.spreadsheet") {
      return lueSheetsTeksti(fileId);
    }

    // 3. Google Slides (.gslides) - TÄRKEÄ KORJAUS
    if (mimeType === "application/vnd.google-apps.presentation") {
      return lueSlidesTeksti(fileId);
    }

    // 4. Tavalliset tekstitiedostot (.txt, .csv, .json, .md)
    if (mimeType === MimeType.PLAIN_TEXT || (mimeType.indexOf("text/") === 0 && mimeType !== "text/html")) {
      return tiedosto.getBlob().getDataAsString("UTF-8");
    }

    // 5. HTML
    if (mimeType === "text/html") {
      let html = tiedosto.getBlob().getDataAsString("UTF-8");
      return html.replace(/<[^>]*>/g, " ");
    }

    // 6. PDF ja Word (.pdf, .doc, .docx)
    if (
      mimeType === MimeType.PDF || 
      mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || 
      mimeType === "application/msword"
    ) {
      return luePdfTaiWordTeksti(tiedosto);
    }

    // 7. Excel (.xlsx) ja PowerPoint (.pptx)
    if (
      mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      mimeType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ) {
      return lueMuunnettuDriveTeksti(tiedosto);
    }

    // Muut muodot (esim. kuvat .png, .jpg)
    return "[Ei-tuettu tiedostomuoto: " + tiedosto.getName() + " (" + mimeType + ")]";

  } catch (err) {
    Logger.log("Virhe tiedoston '" + tiedosto.getName() + "' lukemisessa: " + err.message);
    return "[Virhe tiedoston lukemisessa: " + err.message + "]";
  }
}

/**
 * Lukee Google Drive -tiedoston. 
 * Google Docsit luetaan HTML-muodossa muotoilujen säilyttämiseksi Geminille.
 * Tukee myös Sheets-, Slides-, Excel- ja PowerPoint-tiedostoja.
 */
function lueTiedostoIdlla(fileId) {
  try {
    const tiedosto = DriveApp.getFileById(fileId);
    const mimeType = tiedosto.getMimeType();

    // 1. Google Docs -tiedostot -> Luetaan HTML-muodossa muotoilujen säilyttämiseksi!
    if (mimeType === MimeType.GOOGLE_DOCS) {
      const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${fileId}&exportFormat=html`;
      const options = {
        headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
      };
      return UrlFetchApp.fetch(url, options).getContentText();
    }
    
    // 2. Tavalliset tekstitiedostot (.txt, .csv, .json, .md)
    if (mimeType === MimeType.PLAIN_TEXT || (mimeType.indexOf("text/") === 0 && mimeType !== "text/html")) {
      return tiedosto.getBlob().getDataAsString("UTF-8");
    }

    // 3. HTML-tiedostot (jos sellainen on ladattu suoraan Driveen)
    if (mimeType === "text/html") {
      let html = tiedosto.getBlob().getDataAsString("UTF-8");
      return html.replace(/<[^>]*>/g, " ");
    }

    // 4. PDF- ja Word-tiedostot (.pdf, .doc, .docx)
    if (
      mimeType === MimeType.PDF || 
      mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || 
      mimeType === "application/msword"
    ) {
      return luePdfTaiWordTeksti(tiedosto);
    }

    // 5. Google Sheets -taulukot (.gsheet)
    if (mimeType === MimeType.GOOGLE_SHEETS) {
      return lueSheetsTeksti(fileId);
    }

    // 6. Google Slides -esitykset (.gslides)
    if (mimeType === MimeType.GOOGLE_SLIDES) {
      return lueSlidesTeksti(fileId);
    }

    // 7. Excel (.xlsx) ja PowerPoint (.pptx)
    if (
      mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      mimeType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ) {
      return lueMuunnettuDriveTeksti(tiedosto);
    }

    // Muut muodot (esim. kuvat .png, .jpg)
    return "[Ei-tuettu tiedostomuoto: " + tiedosto.getName() + " (" + mimeType + ")]";

  } catch (err) {
    Logger.log("Virhe tiedoston lukemisessa: " + err.message);
    return "[Virhe tiedoston lukemisessa: " + err.message + "]";
  }
}

// ==========================================
// TARVITTAVAT APUFUNKTIOT
// ==========================================

// Lukee Google Sheets -taulukon tekstit
function lueSheetsTeksti(id) {
  let ss = SpreadsheetApp.openById(id);
  let sheets = ss.getSheets();
  let tekstit = [];

  sheets.forEach(sheet => {
    let data = sheet.getDataRange().getValues();
    tekstit.push("--- Taulukko: " + sheet.getName() + " ---");
    data.forEach(rivi => {
      tekstit.push(rivi.join("\t"));
    });
  });

  return tekstit.join("\n");
}

// Lukee Google Slides -esityksen tekstit dioittain
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

// Muuntaa Excel- (.xlsx) ja PowerPoint- (.pptx) tiedostot väliaikaisesti Drive API v2:lla
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

  // Poistetaan aina väliaikainen tiedosto
  Drive.Files.remove(tempDoc.id);

  return teksti;
}

/**
 * Apufunktio PDF- ja Word-tiedostojen sisällön lukemiseen HTML-muodossa (Drive API v2).
 * Säilyttää muotoilut (otsikot, lihavoinnit, taulukot ja listat) Geminille.
 * 
 * @param {GoogleAppsScript.Drive.File} tiedosto - Luettava Word- tai PDF-tiedosto.
 * @return {string} Puhdistettu HTML-teksti.
 */
function luePdfTaiWordTeksti(tiedosto) {
  let tempFileId = null;
  try {
    let blob = tiedosto.getBlob();
    let mimeType = tiedosto.getMimeType();

    let resource = {
      title: "[TEMP] " + tiedosto.getName()
    };

    let tempFile = null;

    // YRITYS 1: Normaali konversio ilman OCR-tunnistusta (ei kuluta OCR-kiintiötä)
    try {
      tempFile = Drive.Files.insert(resource, blob, { convert: true });
    } catch (err) {
      // YRITYS 2: Jos kyseessä oli skannattu PDF ja 1. yritys epäonnistui, yritetään OCR:ää
      if (mimeType === MimeType.PDF) {
        Logger.log("Sujuva muunnos epäonnistui, yritetään OCR-tunnistusta: " + tiedosto.getName());
        tempFile = Drive.Files.insert(resource, blob, { convert: true, ocr: true });
      } else {
        throw err;
      }
    }

    tempFileId = tempFile.id;

    // Odotetaan hetki tiedoston indeksöintiä ja muunnosta
    Utilities.sleep(1000);

    // Haetaan väliaikaisen Google Docin HTML-sisältö muotoilujen säilyttämiseksi
    const exportUrl = `https://docs.google.com/feeds/download/documents/export/Export?id=${tempFileId}&exportFormat=html`;
    const options = {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    };
    
    let htmlTeksti = UrlFetchApp.fetch(exportUrl, options).getContentText();

    // Poistetaan Googlen valtava <style>-koodi ja luokkamääritykset, jotka sotkevat Geminin
    htmlTeksti = htmlTeksti.replace(/<style[\s\S]*?<\/style>/gi, '');
    htmlTeksti = htmlTeksti.replace(/ class="[^"]*"/gi, '');
    htmlTeksti = htmlTeksti.replace(/ style="[^"]*"/gi, '');
    htmlTeksti = htmlTeksti.replace(/<\/?span[^>]*>/gi, '');

    // Poistetaan väliaikainen tiedosto Drivesta
    DriveApp.getFileById(tempFileId).setTrashed(true);

    return htmlTeksti.trim() || "[Tiedosto oli tyhjä tai tekstiä ei pystytty tunnistamaan]";

  } catch (e) {
    // Siivotaan väliaikaistiedosto virhetilanteessa
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (subErr) {}
    }
    Logger.log("Virhe PDF/Word-konversiossa (" + tiedosto.getName() + "): " + e.toString());
    
    // Jos kiintiö on täynnä, annetaan selkeä ilmoitus
    if (e.toString().indexOf("OCR") !== -1 || e.toString().indexOf("rate limit") !== -1) {
      return "[Ei voitu lukea: Googlen OCR-päiväkiintiö täynnä]";
    }
    
    return "[PDF/Word-tiedoston tekstin lukeminen epäonnistui: " + e.toString() + "]";
  }
}

/**
 * Varmistaa, että taulukon ensimmäisellä rivillä on otsikot.
 */
function alustaTaulukkoTarvittaessa(sheet) {

  sheet.clear();

  let otsikot = [["Aikaleima", "Ohjeet agentille", "Hakemuksen linkki", "Drive-kansio URL", "Hakemuspohjan ID", "Kielimalli", "Geminin vastaus", "Kansion tila"]];
  sheet.getRange(1, 1, 1, otsikot[0].length).setValues(otsikot);
  sheet.getRange(1, 1, 1, otsikot[0].length).setFontWeight("bold");
  
}

/**
 * Purkaa Drive-kansion URL-osoitteesta kansio-ID:n.
 */
function haeKansioIdUrlsta(url) {
  if (!url) return null;
  
  // Jos syöte on jo valmiiksi pelkkä ID
  if (url.indexOf("http") !== 0 && url.length > 20) {
    return url.trim();
  }

  // Etsitään kansio-ID osoitteesta
  let match = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  return (match && match[1]) ? match[1] : null;
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

      return textPart.text;
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

// 1. Tallennettaan pelkkä tiedosto Driveen (eikä kutsuta vielä AI-agenttia)
function tallennaTiedostoMuistiin(base64Data, tiedostonNimi, mimeType) {
  try {
    let decodedData = Utilities.base64Decode(base64Data);
    let blob = Utilities.newBlob(decodedData, mimeType, tiedostonNimi);
    
    // Tallennetaan Driveen
    let tiedosto = DriveApp.createFile(blob);
    
    // Palautetaan tiedoston ID selaimelle myöhempää käyttöä varten
    return {
      success: true,
      fileId: tiedosto.getId(),
      fileName: tiedosto.getName()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Hakee pääsivun sekä KAIKKI sen sisältämät linkit (sisäiset, ulkoiset ja PDF:ät).
 *
 * @param {string} startUrl - Pääsivun osoite.
 * @param {number} maxSubLinks - Kuinka monta linkkiä enintään haetaan (oletus: 5).
 * @return {string} Palauttaa kaikkien sivujen ja PDF-tiedostojen tekstit yhdessä muuttujassa.
 */
function haeTiedotSivulta(startUrl, maxSubLinks = 8) {
  let koottuAineisto = "";

  try {
    // 1. Haetaan pääsivu Jinalla
    Logger.log("Haetaan pääsivua: " + startUrl);
    const mainResponse = UrlFetchApp.fetch("https://r.jina.ai/" + startUrl, { muteHttpExceptions: true });

    // TARKISTUS 1: Tarkistetaan HTTP-statuskoodi (200 = Onnistui)
    const responseCode = mainResponse.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`Sivun haku epäonnistui statuskoodilla ${responseCode}. Tarkista osoite.`);
    }

    const mainText = mainResponse.getContentText();

    // TARKISTUS 2: Tarkistetaan Jina.ai:n palauttamat virhetekstit
    if (!mainText || mainText.includes("Invalid URL") || mainText.includes("Target URL returned") || mainText.startsWith("Error:")) {
      throw new Error("Syötetty URL-osoite on virheellinen tai sivuun ei päästä käsiksi.");
    }

    koottuAineisto += `=== PÄÄSIVU: ${startUrl} ===\n\n${mainText}\n\n`;

    // 2. Poimitaan sivulta kaikki URL-osoitteet
    const urlRegex = /https?:\/\/[^\s\)\"\'>]+/g;
    const loydetytLinkit = mainText.match(urlRegex) || [];

    // 3. Poistetaan mediatiedostot ja koodit, mutta JÄTETÄÄN .pdf MUKAAN!
    const hylattavatPaatteet = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.css', '.js', '.json', '.ico', '.zip'];

    const valitutLinkit = [...new Set(loydetytLinkit)]
      .filter(link => {
        const osat = link.toLowerCase().split('?')[0]; // Poistetaan query-parametrit ennen päätteen tarkistusta
        const onkoMediaTaiKoodi = hylattavatPaatteet.some(paate => osat.endsWith(paate));

        return link !== startUrl && 
               !onkoMediaTaiKoodi &&
               !link.includes("facebook.com") && 
               !link.includes("twitter.com") &&
               !link.includes("linkedin.com") &&
               !link.includes("x.com");
      })
      .slice(0, maxSubLinks); // Rajoitetaan määrää suoritusajan säästämiseksi

    Logger.log(`Löytyi ${loydetytLinkit.length} linkkiä. Haetaan seuraavat ${valitutLinkit.length} linkkiä (ml. PDF:ät):`, valitutLinkit);

    // 4. Haetaan valitut linkit ja PDF-tiedostot Jinalla
    valitutLinkit.forEach((subUrl, index) => {
      try {
        const onkoPdf = subUrl.toLowerCase().split('?')[0].endsWith('.pdf');
        const tyyppiTeksti = onkoPdf ? "PDF-TIEDOSTO" : "SIVU";

        Logger.log(`Haetaan ${tyyppiTeksti} (${index + 1}/${valitutLinkit.length}): ${subUrl}`);
        
        // Jina osaa muuntaa myös PDF-linkin tekstiksi!
        const subResponse = UrlFetchApp.fetch("https://r.jina.ai/" + subUrl, { muteHttpExceptions: true });
        const subText = subResponse.getContentText();

        koottuAineisto += `\n--- ALALINKKI ${index + 1} [${tyyppiTeksti}]: ${subUrl} ---\n\n${subText}\n\n`;
      } catch (err) {
        Logger.log(`Virhe haettaessa kohdetta (${subUrl}): ` + err.toString());
      }
    });

  } catch (error) {
    Logger.log("Virhe pääsivun haussa: " + error.toString());
    throw new Error("Pääsivun hakeminen epäonnistui: " + error.message);
  }

  return koottuAineisto;
}

/**
 * Pääfunktio, joka luo visuaalisen Google Docsin Geminin puhtaasta HTML-vastauksesta.
 * Soveltaa keskitettyä CSS-tyylitiedostoa ilman AI:n inline-tyylejä.
 * 
 * Vaatii: Drive API v2 (Services -> Drive API v2)
 */
function tallennaGeminivastaus(vastaus, folderId, docTitle) {
  if (!vastaus) {
    throw new Error("Virhe: Geminin vastaus oli tyhjä.");
  }

  // 1. Puhdistetaan Geminin vastaus mahdollisista koodilohkomerkeistä
  let puhdasHtml = vastaus
    .replace(/^```html/gi, '')
    .replace(/^```/gi, '')
    .replace(/```$/gi, '')
    .trim();

  // 2. Poistetaan Geminin mahdollisesti palauttamat ulkoiset html/body-tagit,
  //    jotta kääriydymme AINA meidän omaan täydelliseen <style>-runkoomme.
  puhdasHtml = puhdasHtml
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<\/?html[^>]*>/gi, '')
    .replace(/<head>[\s\S]*?<\/head>/gi, '')
    .replace(/<\/?body[^>]*>/gi, '')
    .trim();

  // 3. Poistetaan AI:n mahdolliset inline-style ja class -määritteet
  puhdasHtml = puhdasHtml
    .replace(/\s*style="[^"]*"/gi, '')
    .replace(/\s*class="[^"]*"/gi, '')
    .replace(/<\/?thead>/gi, '')
    .replace(/<\/?tbody>/gi, '');

  // 4. Kääritään AI:n tuottama rakenne omaan CSS-tyylitiedostoomme
  const kokoHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    /* Sivun perusasetukset */
    body { 
      font-family: Arial, sans-serif; 
      font-size: 10pt; 
      line-height: 1.3; 
      color: #333333; 
    }

    /* Pää- ja väliotsikot (esim. 1 Rahoittava viranomainen, 3.1 Hankkeen...) */
    h1 { 
      font-size: 15pt; 
      font-weight: bold; 
      color: #1a5276; /* EURA-tummansininen */
      margin-top: 16pt; 
      margin-bottom: 6pt; 
    }
    h2 { 
      font-size: 12pt; 
      font-weight: bold; 
      color: #1a5276; 
      margin-top: 12pt; 
      margin-bottom: 4pt; 
    }

    /* Taulukoiden muotoilu */
    table { 
      border-collapse: collapse; 
      width: 100%; 
      margin-top: 4pt; 
      margin-bottom: 8pt; 
    }

    td, th { 
      border: 1px solid #666666; 
      padding: 5px 8px; 
      vertical-align: top; 
      background-color: #ffffff !important; /* Pakotetaan valkoinen tausta (ei automaattisia harmaasävyjä) */
    }

    /* Taulukon sisällä olevat tekstikappaleet */
    td p {
      margin: 0 0 3pt 0;
      font-size: 10pt;
      color: #000000;
    }

    /* Solun sisällä oleva kentän otsikko (<b>) */
    td b, td strong {
      color: #1a5276; /* Tummansininen kentän otsikkoväri solussa */
      font-size: 9.5pt;
      display: block;
      margin-bottom: 2pt;
    }

    /* Kursivoidut ohjetekstit solussa (<i>) */
    td i, td em {
      color: #666666;
      font-size: 9.5pt;
    }

    /* Taulukon ulkopuoliset lihavoidut tekstikappaleet (esim. Ryhmähankkeen toteuttajat) */
    body > p b, body > p strong {
      color: #000000;
      font-size: 10pt;
    }
  </style>
</head>
<body>
  ${puhdasHtml}
</body>
</html>`;

  // 5. Luodaan väliaikainen .html-tiedosto haluttuun Drive-kansioon
  const targetFolder = DriveApp.getFolderById(folderId);
  const tempBlob = Utilities.newBlob(kokoHtml, 'text/html', docTitle + '.html');
  const tempFile = targetFolder.createFile(tempBlob);

  // 6. Muunnetaan HTML visuaaliseksi Google Dociksi Drive API v2 { convert: true } -toiminnolla
  const resource = {
    title: docTitle,
    mimeType: MimeType.GOOGLE_DOCS,
    parents: [{ id: folderId }]
  };

  const newDoc = Drive.Files.insert(resource, tempFile.getBlob(), { convert: true });

  // 7. Poistetaan väliaikainen HTML-tiedosto roskakoriin
  tempFile.setTrashed(true);

  const docUrl = "[https://docs.google.com/document/d/](https://docs.google.com/document/d/)" + newDoc.id;
  Logger.log('Tiedosto luotu ja muunnettu onnistuneesti! URL: ' + docUrl);
  return docUrl;
}

function luoTiedostonNimi() {
  const nyt = new Date();
  const aikavyöhyke = Session.getScriptTimeZone(); // Hakee projektin aikavyöhykkeen (esim. Europe/Helsinki)
  
  // Muotoillaan päivämäärä ja aika muotoon: VV-KK-PP_TT-MM-SS
  // Esimerkki: "hakemusluonnos_2026-09-30_19-52"
  const pvmAikaMuotoiltu = Utilities.formatDate(nyt, aikavyöhyke, "yyyy-MM-dd_HH-mm");
  
  const tiedostonNimi = `hakemusluonnos_${pvmAikaMuotoiltu}`;
  
  Logger.log(tiedostonNimi); 
  return tiedostonNimi;
}

// Funktio, jota selain kyselee säännöllisesti
function haeSuorituksenTila() {
  return PropertiesService.getUserProperties().getProperty("STATUS") || "Käynnistetään...";
}