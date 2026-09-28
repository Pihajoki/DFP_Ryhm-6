// =================================================================
// 1. KÄYTTÖLIITTYMÄ JA YLÄVALIKKO
// =================================================================

/**
 * Luo ylävalikon Google Sheetsiin, kun taulukko avataan
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Rahoitusautomaatti')
    .addItem('Avaa Asetukset (Pop-up)', 'avaaAsetuksetPopup')
    .addItem('Suorita haku nyt', 'ajaRahoitushakuKayttoliittymalla')
    .addToUi();
}

/**
 * Avaa HTML-pohjaisen Pop-up ikkunan
 */
function avaaAsetuksetPopup() {
  const html = HtmlService.createHtmlOutputFromFile('AsetuksetUI')
      .setWidth(500)
      .setHeight(550)
      .setTitle('⚙️ Rahoitusautomaatin Asetukset');
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ Rahoitusautomaatin Asetukset');
}

/**
 * Haetaan nykyiset asetukset Pop-up ikkunaa varten
 */
function haeAsetuksetPopupiin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const asetukset = lueAsetuksetSheetsista(ss);
  return {
    avainsanat: asetukset.avainsanat.join("\n"),
    spostit: asetukset.spostit.join("\n")
  };
}

/**
 * Tallentaa Pop-up ikkunasta lähetetyt uudet asetukset Sheetsiin
 */
function tallennaAsetuksetPopupista(uudetSanatTeksti, uudetSpostitTeksti) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Asetukset");
  if (!sheet) {
    sheet = ss.insertSheet("Asetukset");
  }
  
  sheet.clear();
  sheet.getRange("A1").setValue("Kiinnostavat Avainsanat").setFontWeight("bold").setBackground("#c9daf8");
  sheet.getRange("B1").setValue("Ilmoitus-Sähköpostit").setFontWeight("bold").setBackground("#d9ead3");

  const sanatRivit = uudetSanatTeksti.split("\n").map(s => [s.trim()]).filter(s => s[0] !== "");
  const spostitRivit = uudetSpostitTeksti.split("\n").map(s => [s.trim()]).filter(s => s[0] !== "");

  if (sanatRivit.length > 0) {
    sheet.getRange(2, 1, sanatRivit.length, 1).setValues(sanatRivit);
  }
  if (spostitRivit.length > 0) {
    sheet.getRange(2, 2, spostitRivit.length, 1).setValues(spostitRivit);
  }

  sheet.autoResizeColumns(1, 2);
  return "✅ Asetukset tallennettu!";
}

// =================================================================
// 2. HAKUAUTOMAATTI JA LOGIIKKA
// =================================================================

/**
 * PÄÄFUNKTIO: Lukee asetukset taulukosta, suorittaa haut ja lähettää ilmoitukset
 */
function ajaRahoitushakuKayttoliittymalla() {
  Logger.log("=== ALOITETAAN HAKU TARKIN SANARAJOIN ===");
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create("Rahoitusautomaatti_Dashboard");
  }

  const asetukset = lueAsetuksetSheetsista(ss);
  
  if (asetukset.avainsanat.length === 0) {
    Browser.msgBox("⚠️ Ei avainsanoja!", "Lisää ainakin yksi avainsana 'Asetukset'-välilehdelle tai Pop-up ikkunaan.", Browser.Buttons.OK);
    return;
  }

  Logger.log("Käytössä olevat avainsanat (" + asetukset.avainsanat.length + " kpl): " + asetukset.avainsanat.join(", "));

  let uudetHautYhteensa = [];

  // Haeavustuksia.fi
  let uudetHaeavustuksia = kasitteleUudetHaut(ss, "Haeavustuksia.fi", haeHaeavustuksiaData(asetukset.avainsanat));
  uudetHautYhteensa = uudetHautYhteensa.concat(uudetHaeavustuksia);

  // Rakennerahastot.fi
  let uudetRakennerahastot = kasitteleUudetHaut(ss, "Rakennerahastot", haeRakennerahastotData(asetukset.avainsanat));
  uudetHautYhteensa = uudetHautYhteensa.concat(uudetRakennerahastot);

  // SÄHKÖPOSTI-ILMOITUS
  if (uudetHautYhteensa.length > 0 && asetukset.spostit.length > 0) {
    lahetaIlmoitus(asetukset.spostit.join(","), uudetHautYhteensa);
    Logger.log("📧 Ilmoitus lähetetty osoitteisiin: " + asetukset.spostit.join(", "));
  } else if (uudetHautYhteensa.length === 0) {
    Logger.log("ℹ️ Ei uusia hankkeita tällä ajokerralla.");
  } else {
    Logger.log("⚠️ Uusia hankkeita löytyi, mutta sähköpostiosoitetta ei ollut määritelty Asetukset-sivulla.");
  }
}

/**
 * LUKEE ASETUKSET TAULUKOSTA (Luodaan oletuksilla, jos ei löydy)
 */
function lueAsetuksetSheetsista(ss) {
  let sheet = ss.getSheetByName("Asetukset");
  
  if (!sheet) {
    sheet = ss.insertSheet("Asetukset");
    
    sheet.getRange("A1").setValue("Kiinnostavat Avainsanat").setFontWeight("bold").setBackground("#c9daf8");
    sheet.getRange("B1").setValue("Ilmoitus-Sähköpostit").setFontWeight("bold").setBackground("#d9ead3");
    
    const asiakkaanSanat = [
      ["koulutus"], ["valmennus"], ["fasilitointi"], ["opetus"], ["ohjaus"],
      ["selvitys"], ["ennakointi"], ["tulevaisuus"], ["innovaatio"], ["uusi teknologia"],
      ["uudet teknologiat"], ["tekoäly"], [" ai "], ["robotiikka"], ["drooni"],
      ["droni"], ["dronet"], [" vr "], [" ar "], [" xr "], ["älylasit"],
      ["digitaaliset kaksoset"], ["digitaalinen kaksonen"], ["datan hyödyntäminen"],
      ["dataohjautuvuus"], ["analytiikka"], ["oppimisanalytiikka"], ["osaaminen"],
      ["osaamisen kehittäminen"], ["liiketoimintaosaaminen"], ["competence"],
      ["skills development"], ["digital competence"]
    ];

    const oletusSposti = [[Session.getActiveUser().getEmail()]];
    
    sheet.getRange(2, 1, asiakkaanSanat.length, 1).setValues(asiakkaanSanat);
    sheet.getRange(2, 2, 1, 1).setValues(oletusSposti);
    sheet.autoResizeColumns(1, 2);
  }

  const raakaAvainsanat = sheet.getRange("A2:A" + Math.max(sheet.getLastRow(), 2)).getValues();
  let avainsanat = [];
  raakaAvainsanat.forEach(r => {
    if (r[0] && String(r[0]).trim() !== "") {
      let val = String(r[0]).toLowerCase();
      if (val.trim().length <= 3 && !val.startsWith(" ") && !val.endsWith(" ")) {
        val = " " + val.trim() + " ";
      }
      avainsanat.push(val);
    }
  });

  const raakaSpostit = sheet.getRange("B2:B" + Math.max(sheet.getLastRow(), 2)).getValues();
  let spostit = [];
  raakaSpostit.forEach(r => {
    if (r[0] && String(r[0]).trim() !== "") spostit.push(String(r[0]).trim());
  });

  return { avainsanat: avainsanat, spostit: spostit };
}

/**
 * TARKISTAFUNKTIO: Sanarajat tunnistava suodatin
 */
function onkoKiinnostava(teksti, avainsanat) {
  if (!teksti) return null;
  
  const matalaTeksti = " " + String(teksti).toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, " ") + " ";

  for (let i = 0; i < avainsanat.length; i++) {
    const sana = avainsanat[i];
    
    if (sana.startsWith(" ") || sana.endsWith(" ")) {
      if (matalaTeksti.includes(sana)) {
        return sana.trim();
      }
    } else {
      if (matalaTeksti.includes(sana)) {
        return sana.trim();
      }
    }
  }

  return null;
}

/**
 * DUPLIKAATTISUOJA: Tallentaa vain uudet rivit
 */
function kasitteleUudetHaut(ss, sheetNimi, haetutHaut) {
  let sheet = ss.getSheetByName(sheetNimi);
  if (!sheet) {
    sheet = ss.insertSheet(sheetNimi);
    sheet.appendRow(["Koodi / ID", "Hakuilmoituksen Nimi", "Rahoittaja", "Hakuaika Päättyy", "Löytynyt Avainsana", "Suora Linkki", "Lisätty Taulukkoon"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#c9daf8");
  }

  const data = sheet.getDataRange().getValues();
  let aiemmatLinkit = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5]) aiemmatLinkit.add(String(data[i][5]).trim());
  }

  let uudetRivit = [];
  let uudetHautObj = [];

  haetutHaut.forEach(hanke => {
    if (!aiemmatLinkit.has(String(hanke.linkki).trim())) {
      const lisattyPvm = new Date().toLocaleDateString("fi-FI");
      uudetRivit.push([
        hanke.id || "EU-HAKU",
        hanke.nimi,
        hanke.rahoittaja,
        hanke.loppupvm || "-",
        hanke.osumaSyy.toUpperCase(),
        hanke.linkki,
        lisattyPvm
      ]);
      uudetHautObj.push(hanke);
    }
  });

  if (uudetRivit.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, uudetRivit.length, 7).setValues(uudetRivit);
    sheet.autoResizeColumns(1, 7);
    Logger.log(`✅ Lisättiin ${uudetRivit.length} uutta riviä sivuun '${sheetNimi}'.`);
  }

  return uudetHautObj;
}

// =================================================================
// 3. API- JA RAAPUTUSFUNKTIOT
// =================================================================

function haeHaeavustuksiaData(avainsanat) {
  const apiUrl = "https://www.haeavustuksia.fi/api/haku/list-items?Pagination.Page=1&Pagination.PageSize=500&Language=fi&ShowFuture=true&ShowOngoing=true&ShowEnded=false";
  let tulokset = [];
  try {
    const res = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText());
      (data.hakuilmoitukset || []).forEach(item => {
        let nimi = (typeof item.nimi === "object" && item.nimi !== null) ? (item.nimi.fi || "") : (item.nimi || "");
        let osuma = onkoKiinnostava(nimi, avainsanat);
        if (osuma) {
          tulokset.push({
            id: item.id || item.hakuId || "VALTIO",
            nimi: nimi,
            rahoittaja: item.jarjestajaNimi || "Valtio",
            loppupvm: item.hakuaikaLoppu || "-",
            linkki: item.id ? "https://www.haeavustuksia.fi/fi/haku/" + item.id : "https://www.haeavustuksia.fi",
            osumaSyy: osuma
          });
        }
      });
    }
  } catch (e) {}
  return tulokset;
}

function haeRakennerahastotData(avainsanat) {
  const url = "https://rakennerahastot.fi/hakuajat";
  let tulokset = [];
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0" } });
    if (res.getResponseCode() === 200) {
      const html = res.getContentText();
      const regex = /<a[^>]+href=["']([^"']+)["'][^>]*>(.*?)<\/a>/gi;
      let match;
      while ((match = regex.exec(html)) !== null) {
        const href = match[1];
        const teksti = match[2].replace(/<[^>]+>/g, '').trim();
        if (teksti.length > 10 && (href.includes("/haku") || href.includes("eura2021"))) {
          let osuma = onkoKiinnostava(teksti, avainsanat);
          if (osuma) {
            tulokset.push({
              id: "EU-HAKU",
              nimi: teksti,
              rahoittaja: "EAKR / ESR+ / JTF",
              loppupvm: "-",
              linkki: href.startsWith("http") ? href : "https://rakennerahastot.fi" + href,
              osumaSyy: osuma
            });
          }
        }
      }
    }
  } catch (e) {}
  return tulokset;
}

function lahetaIlmoitus(vastaanottajat, uudetHaut) {
  const aihe = `🚨 Uusia rahoitushakuja löytynyt (${uudetHaut.length} kpl)`;
  let viesti = `Moi!\n\nAutomaatio löysi ${uudetHaut.length} uutta hakua:\n\n`;
  uudetHaut.forEach((h, i) => {
    viesti += `${i + 1}. ${h.nimi}\n   - Avainsana: ${h.osumaSyy.toUpperCase()}\n   - Linkki: ${h.linkki}\n\n`;
  });
  MailApp.sendEmail(vastaanottajat, aihe, viesti);
}