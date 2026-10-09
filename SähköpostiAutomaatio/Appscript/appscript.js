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
    hakutyyppi: { rahoitushaku: true, avustukset: true },
    rahoittajat: { eu: true, oph: true, ely: true, muut: true },
    sivustot: { haeavustuksia: true, rakennerahastot: true },
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

  kirjoitaLista(sheet, "A", 2, a.hakusanat);
  kirjoitaLista(sheet, "C", 2, a.spostit);
  kirjoitaLista(sheet, "E", 2, a.poissulkevat);

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

  let uudetIlmoitettavat = [];

  // 1. Haeavustuksia.fi
  let rawHaeavustuksia = haeHaeavustuksiaData(asetuksetUI.hakusanat, asetuksetUI.poissulkevat);
  let uudetHaeavustuksia = kasitteleUudetHautOptionA(ss, "Haeavustuksia.fi", rawHaeavustuksia);
  uudetIlmoitettavat = uudetIlmoitettavat.concat(uudetHaeavustuksia);

  // 2. Rakennerahastot.fi
  let rawRakennerahastot = haeRakennerahastotData(asetuksetUI.hakusanat, asetuksetUI.poissulkevat);
  let uudetRakennerahastot = kasitteleUudetHautOptionA(ss, "Rakennerahastot", rawRakennerahastot);
  uudetIlmoitettavat = uudetIlmoitettavat.concat(uudetRakennerahastot);

  // 3. SÄHKÖPOSTI-ILMOITUS
  if (asetuksetUI.spostit.length > 0) {
    if (uudetIlmoitettavat.length > 0) {
      lahetaIlmoitus(asetuksetUI.spostit.join(","), uudetIlmoitettavat);
      return `✅ Taulukko päivitetty! Löytyi ${uudetIlmoitettavat.length} aivan uutta hakua. Sähköposti lähetetty.`;
    } else {
      return "✅ Taulukko päivitetty (suodatukset ajantasalla). Ei uusia ilmoitettavia hakuja.";
    }
  } else {
    return "⚠️ Haku tehty, mutta sähköpostiosoitetta ei ole asetettu!";
  }
}

// =================================================================
// 3. HAKULOGIIKKA JA SYNKRONOINTI (OPTION A)
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

/**
 * OPTIO A: TYHJENTÄÄ JA PÄIVITTÄÄ TAULUKON TÄYSIN AJAN TASALLE.
 * Käyttää 'Muisti'-välilehteä varmistamaan, ettei samaa hakua lähetetä sähköpostiin kahdesti.
 */
/**
 * OPTIO A: SYNKRONOI TAULUKON EIKÄ LÄHETÄ SÄHKÖPOSTIA SAMOISTA HAEISTA
 */
function kasitteleUudetHautOptionA(ss, sheetNimi, haetutHaut) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Varmistetaan 'Muisti'-välilehti
  let muistiSheet = ss.getSheetByName("Muisti");
  if (!muistiSheet) {
    muistiSheet = ss.insertSheet("Muisti");
    muistiSheet.appendRow(["Tunniste", "Nimi", "LähetettyPvm"]);
    try { muistiSheet.hideSheet(); } catch(e) {}
  }

  // Luetaan kaikki aiemmin sähköpostilla lähetetyt tunnisteet
  const vanhatMuistiRivit = muistiSheet.getDataRange().getValues();
  let lahetetytSet = new Set();
  
  for (let m = 1; m < vanhatMuistiRivit.length; m++) {
    const avain = String(vanhatMuistiRivit[m][0] || "").trim();
    if (avain) {
      lahetetytSet.add(avain);
    }
  }

  // 2. Valmistellaan kohdevälilehti (esim. Haeavustuksia.fi)
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
      // LUODAAN täysin puhdistettu uniikki tunniste (ilman erikoismerkkejä)
      const tunniste = luoUniikkiTunniste(hanke);
      
      if (!nähdytTunnisteet.has(tunniste)) {
        kaikkiRivit.push([
          hanke.id || "VALTIO/EU",
          hanke.nimi,
          hanke.rahoittaja,
          hanke.loppupvm || "-",
          String(hanke.osumaSyy).toUpperCase(),
          hanke.linkki,
          lisattyPvm
        ]);
        nähdytTunnisteet.add(tunniste);

        // Jos TÄTÄ TUNNISTETTA ei ole vielä koskaan lähetetty sähköpostitse:
        if (!lahetetytSet.has(tunniste)) {
          uudetSähköpostiin.push(hanke);
          lahetetytSet.add(tunniste);
          uudetMuistiRivit.push([tunniste, hanke.nimi, lisattyPvm]);
        }
      }
    });
  }

  // 3. Kirjoitetaan tuoreet tiedot taulukkoon
  if (kaikkiRivit.length > 0) {
    sheet.getRange(1, 1, kaikkiRivit.length, 7).setValues(kaikkiRivit);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#c9daf8");
    try { sheet.autoResizeColumns(1, 7); } catch(e) {}
  }

  // 4. Tallennetaan uudet tunnisteet Muisti-välilehdelle
  if (uudetMuistiRivit.length > 0) {
    muistiSheet.getRange(muistiSheet.getLastRow() + 1, 1, uudetMuistiRivit.length, 3).setValues(uudetMuistiRivit);
  }

  SpreadsheetApp.flush();

  return uudetSähköpostiin;
}

/**
 * APUFUNKTIO: Puhdistaa ja luo täysin varman tunnisteen hankkeelle.
 */
function luoUniikkiTunniste(hanke) {
  let L = String(hanke.linkki || "").toLowerCase().trim();
  let N = String(hanke.nimi || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  
  // Jos linkki sisältää suoran ID-polun /haku/va-..., käytetään sitä primary-tunnisteena
  if (L.includes("/haku/") && !L.endsWith("/haku/")) {
    return L.split("/haku/")[1].replace(/[^a-z0-9]/g, "");
  }
  
  // Muussa tapauksessa käytetään nimen puhdistettua versiota
  return N;
}

/**
 * APUFUNKTIO: Tyhjentää vanhan muistin, jos haluat aloittaa alusta
 */
function tyhjennaHakuMuisti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = ss.getSheetByName("Muisti");
  if (m) {
    m.clearContents();
    m.appendRow(["Tunniste", "Nimi", "LähetettyPvm"]);
    Logger.log("Muisti tyhjennetty!");
  }
}

// =================================================================
// 4. API- JA RAAPUTUSFUNKTIOT
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

            tulokset.push({
              id: asianumero || "VALTIO",
              nimi: nimiTeksti || "Nimetön haku",
              rahoittaja: jarjestaja,
              loppupvm: item.hakuPaattyyDateTimeUtc || item.hakuaikaLoppu || "-",
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

function haeRakennerahastotData(hakusanat, poissulkevat) {
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
        if (teksti.length > 10 && (href.includes("/haku") || href.includes("eura2021") || href.includes("rakennerahastot"))) {
          let osuma = onkoKiinnostava(teksti, hakusanat, poissulkevat);
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
  } catch (e) {
    Logger.log("Virhe Rakennerahastot.fi haussa: " + e.toString());
  }
  return tulokset;
}

function lahetaIlmoitus(vastaanottajat, uudetHaut) {
  const aihe = `🚨 Uusia rahoitushakuja löytynyt (${uudetHaut.length} kpl)`;
  let viesti = `Moi!\n\nAutomaattinen Hakemusvahti löysi ${uudetHaut.length} uutta hakua:\n\n`;
  uudetHaut.forEach((h, i) => {
    viesti += `${i + 1}. ${h.nimi}\n   - Avainsana: ${h.osumaSyy.toUpperCase()}\n   - Linkki: ${h.linkki}\n\n`;
  });
  
  try {
    MailApp.sendEmail(vastaanottajat, aihe, viesti);
  } catch(e) {
    Logger.log("Virhe sähköpostin lähetyksessä: " + e.toString());
  }
}

// =================================================================
// 5. APUFUNKTIOT
// =================================================================

function alustaAsetusTaulukko(sheet) {
  sheet.getRange("A1").setValue("Hakusanat").setFontWeight("bold");
  sheet.getRange("B1").setValue("PÄÄLLÄ");
  sheet.getRange("C1").setValue("Sähköpostit").setFontWeight("bold");
  sheet.getRange("E1").setValue("Poissulkevat sanat").setFontWeight("bold");
  
  kirjoitaLista(sheet, "A", 2, ["koulutus", "valmennus", "innovaatio"]);
  
  let sposti = "";
  try { sposti = Session.getActiveUser().getEmail(); } catch(e) {}
  if (sposti) kirjoitaLista(sheet, "C", 2, [sposti]);
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