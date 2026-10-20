// =================================================================
// 1. KÄYTTÖLIITTYMÄ JA YLÄVALIKKO
// =================================================================

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.createMenu('⚙️ Rahoitusautomaatti')
        .addItem('Avaa Hakemusvahti', 'avaaAsetuksetPopup')
        .addItem('Suorita haku nyt', 'ajaRahoitushakuKayttoliittymalla')
        .addToUi();
    }
  } catch (e) {}
}

function avaaAsetuksetPopup() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('AsetuksetUI')
        .setWidth(460)
        .setHeight(640)
        .setTitle('Automaattinen Hakemusvahti');
    SpreadsheetApp.getUi().showModalDialog(html, 'Automaattinen Hakemusvahti');
  } catch (e) {
    Logger.log("Pop-upia ei voitu avata: " + e.toString());
  }
}

// Alias vanhalle funktion nimelle
function avaaAsetuksetUI() {
  avaaAsetuksetPopup();
}

function haeKaikkiAsetuksetUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Asetukset");

  if (!sheet) {
    sheet = ss.insertSheet("Asetukset");
    alustaAsetusTaulukko(sheet);
  }

  return {
    tila: sheet.getRange("B1").getValue() !== "POIS",
    hakusanat: luePuhdasLista(sheet, "A", 2),
    spostit: luePuhdasLista(sheet, "C", 2),
    poissulkevat: luePuhdasLista(sheet, "E", 2),
    edellinenAika: sheet.getRange("G2").getValue() || "-"
  };
}

function tallennaAsetuksetJaAjaHaku(a) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Asetukset");
  if (!sheet) {
    sheet = ss.insertSheet("Asetukset");
  }

  // Tyhjennetään sarakkeet A, C ja E vanhoista arvoista
  sheet.getRange("A2:A100").clearContent();
  sheet.getRange("C2:C100").clearContent();
  sheet.getRange("E2:E100").clearContent();

  sheet.getRange("B1").setValue(a.tila ? "PÄÄLLÄ" : "POIS");

  kirjoitaLista(sheet, "A", 2, a.hakusanat || []);
  kirjoitaLista(sheet, "C", 2, a.spostit || []);
  kirjoitaLista(sheet, "E", 2, a.poissulkevat || []);

  const nyt = new Date();
  const aikaStr = nyt.toLocaleDateString("fi-FI") + " / klo " + nyt.toLocaleTimeString("fi-FI", {hour: '2-digit', minute:'2-digit'});
  sheet.getRange("G2").setValue(aikaStr);

  let tulosViesti = "Asetukset tallennettu!";
  if (a.tila) {
    tulosViesti = ajaRahoitushakuKayttoliittymalla();
  }

  return { viesti: tulosViesti, aika: aikaStr };
}

// =================================================================
// 2. PÄÄFUNKTIO JA HAKUAUTOMAATTI
// =================================================================

function ajaRahoitushakuKayttoliittymalla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const asetuksetUI = haeKaikkiAsetuksetUI();

  if (!asetuksetUI.tila) return "Haku kytketty POIS päältä.";
  if (!asetuksetUI.hakusanat || asetuksetUI.hakusanat.length === 0) return "⚠️ Ei hakusanoja valittuna.";

  // Haeavustuksia.fi -haku
  let rawHaeavustuksia = haeHaeavustuksiaData(asetuksetUI.hakusanat, asetuksetUI.poissulkevat);
  let uudetIlmoitettavat = kasitteleUudetHautOptionA(ss, "Haeavustuksia.fi", rawHaeavustuksia);

  // SÄHKÖPOSTI-ILMOITUS
  if (asetuksetUI.spostit.length > 0) {
    if (uudetIlmoitettavat.length > 0) {
      lahetaIlmoitus(asetuksetUI.spostit.join(","), uudetIlmoitettavat);
      return `✅ Taulukko päivitetty! Löytyi ${uudetIlmoitettavat.length} uutta hakua. Sähköposti lähetetty.`;
    } else {
      return "✅ Taulukko päivitetty (suodatukset ajantasalla). Ei uusia ilmoitettavia hakuja.";
    }
  } else {
    return "⚠️ Haku tehty, mutta sähköpostiosoitetta ei ole asetettu!";
  }
}

// =================================================================
// 3. HAKULOGIIKKA JA SYNKRONOINTI
// =================================================================

function onkoKiinnostava(teksti, hakusanat, poissulkevatSanat) {
  if (!teksti) return null;
  
  const rawTeksti = String(teksti).toLowerCase();
  const siivottuTeksti = " " + rawTeksti.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, " ") + " ";

  // 1. POISSULKEVAT SANAT: Jos löytyy, hylätään heti
  if (poissulkevatSanat && poissulkevatSanat.length > 0) {
    for (let i = 0; i < poissulkevatSanat.length; i++) {
      let eiSana = String(poissulkevatSanat[i]).toLowerCase().trim();
      if (!eiSana) continue;

      if (eiSana.length <= 3) {
        if (siivottuTeksti.includes(" " + eiSana + " ")) return null;
      } else {
        if (rawTeksti.includes(eiSana)) return null;
      }
    }
  }

  // 2. HYVÄKSYTTÄVÄT HAKUSANAT
  for (let i = 0; i < hakusanat.length; i++) {
    let sana = String(hakusanat[i]).toLowerCase().trim();
    if (!sana) continue;

    if (sana.length <= 3) {
      if (siivottuTeksti.includes(" " + sana + " ")) return sana;
    } else {
      if (rawTeksti.includes(sana)) return sana;
    }
  }

  return null;
}

function kasitteleUudetHautOptionA(ss, sheetNimi, haetutHaut) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let muistiSheet = ss.getSheetByName("Muisti");
  if (!muistiSheet) {
    muistiSheet = ss.insertSheet("Muisti");
    muistiSheet.appendRow(["Tunniste", "Nimi", "LähetettyPvm"]);
    try { muistiSheet.hideSheet(); } catch(e) {}
  }

  const vanhatMuistiRivit = muistiSheet.getDataRange().getValues();
  let lahetetytSet = new Set();
  
  for (let m = 1; m < vanhatMuistiRivit.length; m++) {
    const avain = String(vanhatMuistiRivit[m][0] || "").trim();
    if (avain) {
      lahetetytSet.add(avain);
    }
  }

  let sheet = ss.getSheetByName(sheetNimi);
  if (!sheet) {
    sheet = ss.insertSheet(sheetNimi);
  }

  sheet.clearContents();
  sheet.clearFormats();

  const otsikot = ["Koodi / ID", "Hakuilmoituksen Nimi", "Rahoittaja", "Hakuaika Päättyy", "Löytynyt Avainsana", "Suora Linkki", "Lisätty Taulukkoon"];
  let kaikkiRivit = [otsikot];
  let nähdytTunnisteet = new Set();
  const lisattyPvm = new Date().toLocaleDateString("fi-FI");

  let uudetSähköpostiin = [];
  let uudetMuistiRivit = [];

  if (haetutHaut && haetutHaut.length > 0) {
    haetutHaut.forEach(hanke => {
      const tunniste = luoUniikkiTunniste(hanke);
      
      if (!nähdytTunnisteet.has(tunniste)) {
        let linkkiKaava = hanke.linkki;
        if (hanke.linkki && hanke.linkki.startsWith("http")) {
          linkkiKaava = `=HYPERLINK("${hanke.linkki}"; "Avaa haku 🔗")`;
        }

        kaikkiRivit.push([
          hanke.id || "VALTIO",
          hanke.nimi,
          hanke.rahoittaja,
          hanke.loppupvm || "-",
          String(hanke.osumaSyy).toUpperCase(),
          linkkiKaava,
          lisattyPvm
        ]);
        nähdytTunnisteet.add(tunniste);

        if (!lahetetytSet.has(tunniste)) {
          uudetSähköpostiin.push(hanke);
          lahetetytSet.add(tunniste);
          uudetMuistiRivit.push([tunniste, hanke.nimi, lisattyPvm]);
        }
      }
    });
  }

  if (kaikkiRivit.length > 0) {
    sheet.getRange(1, 1, kaikkiRivit.length, 7).setValues(kaikkiRivit);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#c9daf8");
    try { sheet.autoResizeColumns(1, 7); } catch(e) {}
  }

  if (uudetMuistiRivit.length > 0) {
    muistiSheet.getRange(muistiSheet.getLastRow() + 1, 1, uudetMuistiRivit.length, 3).setValues(uudetMuistiRivit);
  }

  SpreadsheetApp.flush();

  return uudetSähköpostiin;
}

function luoUniikkiTunniste(hanke) {
  let L = String(hanke.linkki || "").toLowerCase().trim();
  let N = String(hanke.nimi || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  
  if (L.includes("/haku/") && !L.endsWith("/haku/")) {
    return L.split("/haku/")[1].replace(/[^a-z0-9]/g, "");
  }
  return N;
}

// =================================================================
// 4. API- HAKU JA PÄÄTTYMISPÄIVÄT
// =================================================================

function haeHaeavustuksiaData(hakusanat, poissulkevat) {
  let tulokset = [];
  let sivu = 1;
  let sivuKoko = 100;
  let jatkaHakua = true;

  try {
    while (jatkaHakua) {
      const apiUrl = `https://www.haeavustuksia.fi/api/haku/list-items?Pagination.Page=${sivu}&Pagination.PageSize=${sivuKoko}&Language=fi&ShowFuture=true&ShowOngoing=true&ShowEnded=false`;
      const res = UrlFetchApp.fetch(apiUrl, { 
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0" }
      });
      
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText());
        const ilmoitukset = data.hakuilmoitukset || data.items || [];
        
        if (!ilmoitukset || ilmoitukset.length === 0) {
          jatkaHakua = false;
          break;
        }

        ilmoitukset.forEach(item => {
          let nimiTeksti = "";
          if (item.nimi) {
            if (typeof item.nimi === "object") {
              nimiTeksti = item.nimi.fi || item.nimi.sv || item.nimi.en || "";
            } else {
              nimiTeksti = String(item.nimi);
            }
          }

          let osuma = onkoKiinnostava(nimiTeksti, hakusanat, poissulkevat);
          if (osuma) {
            const asianumero = item.hakuasianAsianumero || item.koodi || item.tunniste || item.id;
            let suoraLinkki = asianumero ? `https://www.haeavustuksia.fi/fi/haku/${asianumero}` : "https://www.haeavustuksia.fi";

            let jarjestaja = "Valtio";
            if (item.vastuullinenJarjestajaOrganisaatio) {
              if (typeof item.vastuullinenJarjestajaOrganisaatio === "object") {
                jarjestaja = item.vastuullinenJarjestajaOrganisaatio.fi || item.vastuullinenJarjestajaOrganisaatio.sv || "Valtio";
              } else {
                jarjestaja = String(item.vastuullinenJarjestajaOrganisaatio);
              }
            } else if (item.jarjestajaNimi) {
              jarjestaja = item.jarjestajaNimi;
            }

            let raakaPvm = item.hakuPaattyyDateTimeUtc || item.hakuaikaLoppu || item.paattyyPvm || "";
            let loppupvm = poimiPaattymisPaivamaara(raakaPvm, nimiTeksti);

            tulokset.push({
              id: asianumero || "VALTIO",
              nimi: nimiTeksti || "Nimetön haku",
              rahoittaja: jarjestaja,
              loppupvm: loppupvm,
              linkki: suoraLinkki,
              osumaSyy: osuma
            });
          }
        });

        if (ilmoitukset.length < sivuKoko || sivu >= 10) {
          jatkaHakua = false;
        } else {
          sivu++;
        }
      } else {
        jatkaHakua = false;
      }
    }
  } catch (e) {
    Logger.log("Virhe Haeavustuksia.fi haussa: " + e.toString());
  }

  return tulokset;
}

function poimiPaattymisPaivamaara(raakaPvm, teksti) {
  if (raakaPvm) {
    let d = new Date(String(raakaPvm));
    if (!isNaN(d.getTime())) {
      let pp = ("0" + d.getDate()).slice(-2);
      let kk = ("0" + (d.getMonth() + 1)).slice(-2);
      let vvvv = d.getFullYear();
      return `${pp}.${kk}.${vvvv}`;
    }
  }

  const t = String(teksti || "");
  const paattyyMatch = t.match(/(?:päättyy|paattyy|saakka|asti|viimeistään)\s*(?:klo\s*\d{1,2}[\.:]\d{2})?\s*(\d{1,2}\.\d{1,2}\.\d{4})/i);
  if (paattyyMatch && paattyyMatch[1]) {
    return paattyyMatch[1];
  }

  const kaikkiPvm = t.match(/(\d{1,2}\.\d{1,2}\.\d{4})/g);
  if (kaikkiPvm && kaikkiPvm.length > 0) {
    return kaikkiPvm[kaikkiPvm.length - 1];
  }

  return "-";
}

function lahetaIlmoitus(vastaanottajat, uudetHaut) {
  var aihe = "🚨 Uusia rahoitushakuja löytynyt (" + uudetHaut.length + " kpl)";
  var viesti = "Moi!\n\nAutomaattinen Hakemusvahti löysi " + uudetHaut.length + " uutta hakua:\n\n";

  uudetHaut.forEach(function(h, i) {
    viesti += (i + 1) + ". " + h.nimi + "\n";
    viesti += "   - Hakuaika päättyy: " + (h.loppupvm || "-") + "\n";
    viesti += "   - Avainsana: " + String(h.osumaSyy).toUpperCase() + "\n";
    viesti += "   - Linkki: " + h.linkki + "\n\n";
  });

  try {
    MailApp.sendEmail(vastaanottajat, aihe, viesti);
  } catch (e) {
    Logger.log("Virhe sähköpostin lähetyksessä: " + e.toString());
  }
}

// =================================================================
// 5. APUFUNKTIOT
// =================================================================

function alustaAsetusTaulukko(sheet) {
  sheet.clearContents();
  sheet.getRange("A1").setValue("Hakusanat").setFontWeight("bold");
  sheet.getRange("B1").setValue("PÄÄLLÄ");
  sheet.getRange("C1").setValue("Sähköpostit").setFontWeight("bold");
  sheet.getRange("E1").setValue("Poissulkevat sanat").setFontWeight("bold");

  kirjoitaLista(sheet, "A", 2, ["koulutus", "valmennus", "hanke", "kehittäminen"]);
  
  let sposti = "";
  try { 
    sposti = Session.getActiveUser().getEmail(); 
  } catch(e) {}
  
  if (sposti) {
    kirjoitaLista(sheet, "C", 2, [sposti]);
  }
}

function luePuhdasLista(sheet, sarakeKirjain, aloitusRivi) {
  const maxRivi = Math.max(sheet.getLastRow(), aloitusRivi);
  const arr = sheet.getRange(sarakeKirjain + aloitusRivi + ":" + sarakeKirjain + maxRivi).getValues();
  let tulos = [];
  arr.forEach(r => {
    if (r[0] && String(r[0]).trim() !== "") {
      tulos.push(String(r[0]).trim());
    }
  });
  return tulos;
}

function kirjoitaLista(sheet, sarakeKirjain, aloitusRivi, dataArray) {
  if (dataArray && dataArray.length > 0) {
    const rivit = dataArray.map(item => [item]);
    sheet.getRange(aloitusRivi, sarakeKirjaimestaNro(sarakeKirjain), rivit.length, 1).setValues(rivit);
  }
}

function sarakeKirjaimestaNro(letter) {
  return letter.charCodeAt(0) - 64;
}