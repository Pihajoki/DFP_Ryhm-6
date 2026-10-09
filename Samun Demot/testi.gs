// 1. TESTIFUNKTIO – Aja tämä yläpalkista (Suorita / Run)
function testaaPelkkaAlisivujenKerays() {
  Logger.log("=== ALOITETAAN JINA-ALISIVUJEN KERÄYSTESTI ===");
  
  const kohdeUrl = "https://eura2021.fi/hakuilmoitukset";
  keraaJaTulostaKaikkiAlisivutJinalla(kohdeUrl);
  
  Logger.log("\n=== KERÄYSTESTI SUORITETTU ===");
}

// 2. KERÄYS- JA TULOSTUSFUNKTIO
function keraaJaTulostaKaikkiAlisivutJinalla(paasivuUrl) {
  const jinaUrl = "https://r.jina.ai/" + paasivuUrl;

  try {
    // 1. Haetaan pääsivu
    const res = UrlFetchApp.fetch(jinaUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      Logger.log("Pääsivun haku epäonnistui statuskoodilla: " + res.getResponseCode());
      return;
    }

    const paasivuTeksti = res.getContentText() || "";

    // 2. Poimitaan kaikki alisivulinkit pääsivun Markdown-koodista
    const linkkiRegex = /\[([^\]]+)\]\((https?:\/\/[^\s\)]+)\)/g;
    let match;
    let kaikkialisivuUrlit = [];

    while ((match = linkkiRegex.exec(paasivuTeksti)) !== null) {
      const href = match[2];
      if (href.includes("/hakuilmoitus") || href.includes("/haku/")) {
        kaikkialisivuUrlit.push(href);
      }
    }

    // Karsitaan tuplat
    const uniikitUrlit = [...new Set(kaikkialisivuUrlit)];
    Logger.log("Löytyi yhteensä " + uniikitUrlit.length + " uniikkia alisivua.\n");

    // 3. Luetaan kukin alisivu vuorollaan ja tulostetaan teksti lokiin
    for (let i = 0; i < uniikitUrlit.length; i++) {
      const alisivuUrl = uniikitUrlit[i];

      Logger.log("\n==================================================");
      Logger.log("LADATAAN ALISIVU (" + (i + 1) + " / " + uniikitUrlit.length + "): " + alisivuUrl);
      Logger.log("==================================================");

      // Pidetään 1.5 sekunnin tauko Jina-ilmaisrajan noudattamiseksi
      Utilities.sleep(1500);

      try {
        const alisivuRes = UrlFetchApp.fetch("https://r.jina.ai/" + alisivuUrl, { muteHttpExceptions: true });
        
        if (alisivuRes.getResponseCode() === 200) {
          const alisivuTeksti = alisivuRes.getContentText() || "";

          if (alisivuTeksti.length > 0) {
            // Tulostetaan teksti lokiin pätkittäin (max 4000 merkkiä / rivi)
            tulostaKokoTekstiLokiin(alisivuTeksti);
          } else {
            Logger.log("-> Varoitus: Alisivun teksti oli tyhjä.");
          }
        } else {
          Logger.log("-> Virhe statuskoodilla: " + alisivuRes.getResponseCode());
        }
      } catch (err) {
        Logger.log("-> Virhe ladattaessa alisivua: " + err.toString());
      }
    }

  } catch (e) {
    Logger.log("Päävirhe: " + e.toString());
  }
}

// Apufunktio: Pätkii tekstin ettei Apps Script leikkaa lokia
function tulostaKokoTekstiLokiin(teksti) {
  const palaKoko = 4000;
  for (let i = 0; i < teksti.length; i += palaKoko) {
    Logger.log(teksti.substring(i, i + palaKoko));
  }
}
