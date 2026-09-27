const KIINNOSTAVAT_AVAINSANAT = [
  "tekoäly", "ai", "digi", "pk-yritys", "innovaatio", 
  "kehittäminen", "teknologia", "automaatio", "vihreä siirtymä"
];

function ajaUudetRahoitushaut() {
  Logger.log("=== TARKISTETAAN UUDET RAHOITUSHAUT ===");
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create("Kiinnostavat_Rahoitushaut_Automaatti");
  }

  let uudetHautYhteensa = [];

  // 1. Haeavustuksia.fi
  let uudetHaeavustuksia = kasitteleUudetHaut(ss, "Haeavustuksia.fi", haeHaeavustuksiaData());
  uudetHautYhteensa = uudetHautYhteensa.concat(uudetHaeavustuksia);

  // 2. Rakennerahastot.fi
  let uudetRakennerahastot = kasitteleUudetHaut(ss, "Rakennerahastot", haeRakennerahastotData());
  uudetHautYhteensa = uudetHautYhteensa.concat(uudetRakennerahastot);

  // 3. SÄHKÖPOSTI-ILMOITUS (Vain jos uusia löytyi!)
  if (uudetHautYhteensa.length > 0) {
    lahetaIlmoitusSahkopostiin(uudetHautYhteensa);
    Logger.log(`📧 Sähköposti lähetetty! Uusia hankkeita löytyi ${uudetHautYhteensa.length} kpl.`);
  } else {
    Logger.log("ℹ️ Ei uusia hankkeita tällä ajokerralla. Ei lähetetä sähköpostia.");
  }
}

/**
 * TARKISTAA DUPLIKAATIT JA TALLENTAA VAIN UUDET RIVIT
 */
function kasitteleUudetHaut(ss, sheetNimi, haetutHaut) {
  let sheet = ss.getSheetByName(sheetNimi);
  if (!sheet) {
    sheet = ss.insertSheet(sheetNimi);
    sheet.appendRow(["Koodi / ID", "Hakuilmoituksen Nimi", "Rahoittaja", "Hakuaika Päättyy", "Löytynyt Avainsana", "Suora Linkki", "Lisätty Taulukkoon"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#c9daf8");
  }

  // Luetaan olemassa olevat linkit/ID:t, jotta tiedetään mikä on jo tallennettu
  const data = sheet.getDataRange().getValues();
  let aiemmatLinkit = new Set();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][5]) aiemmatLinkit.add(String(data[i][5]).trim()); // Sarake F (Linkki)
  }

  let uudetRivit = [];
  let uudetHautObj = [];

  haetutHaut.forEach(hanke => {
    // Jos linkkiä ei löydy aiemmista -> kyseessä on UUSI hanke
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

  // Lisätään vain uudet rivit taulukon loppuun
  if (uudetRivit.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, uudetRivit.length, 7).setValues(uudetRivit);
    sheet.autoResizeColumns(1, 7);
    Logger.log(`✅ Taulukkoon '${sheetNimi}' lisättiin ${uudetRivit.length} uutta hanketta.`);
  }

  return uudetHautObj;
}

/**
 * LÄHETTÄÄ SÄHKÖPOSTIN VAIN UUSISTA HANKKEISTA
 */
function lahetaIlmoitusSahkopostiin(uudetHaut) {
  const kayttajanSposti = Session.getActiveUser().getEmail();
  const aihe = `🚨 Uusia rahoitushakuja löytynyt (${uudetHaut.length} kpl)`;
  
  let viesti = `Moi!\n\nAutomaatio löysi ${uudetHaut.length} uutta hakukriteereihisi sopivaa rahoitushakua:\n\n`;

  uudetHaut.forEach((h, index) => {
    viesti += `${index + 1}. ${h.nimi}\n`;
    viesti += `   - Rahoittaja: ${h.rahoittaja}\n`;
    viesti += `   - Avainsana: ${h.osumaSyy.toUpperCase()}\n`;
    viesti += `   - Linkki: ${h.linkki}\n\n`;
  });

  viesti += "Tiedot on päivitetty myös Google Sheets -taulukkoosi.\n\nTerveisin,\nRahoitusautomaatti";

  MailApp.sendEmail(kayttajanSposti, aihe, viesti);
}
// Voit syöttää tähän pilkulla eroteltuna kaikki vastaanottajat: (LISÄTTY TULEVAISUUTTA VARTEN)
// const vastaanottajat = "oma.osoite@gmail.com, ryhmalainen@gmail.com";

// MailApp.sendEmail(vastaanottajat, aihe, viesti);

/**
 * DATAN HAKUAPUFUNKTIOT
 */
function haeHaeavustuksiaData() {
  const apiUrl = "https://www.haeavustuksia.fi/api/haku/list-items?Pagination.Page=1&Pagination.PageSize=500&Language=fi&ShowFuture=true&ShowOngoing=true&ShowEnded=false";
  let tulokset = [];
  try {
    const res = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText());
      (data.hakuilmoitukset || []).forEach(item => {
        let nimi = (typeof item.nimi === "object" && item.nimi !== null) ? (item.nimi.fi || "") : (item.nimi || "");
        let osuma = onkoKiinnostava(nimi);
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

function haeRakennerahastotData() {
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
          let osuma = onkoKiinnostava(teksti);
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

function onkoKiinnostava(teksti) {
  if (!teksti) return null;
  const matalaTeksti = String(teksti).toLowerCase();
  for (let i = 0; i < KIINNOSTAVAT_AVAINSANAT.length; i++) {
    const sana = KIINNOSTAVAT_AVAINSANAT[i];
    if (matalaTeksti.includes(sana)) return sana;
  }
  return null;
}