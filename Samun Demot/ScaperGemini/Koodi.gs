function paivitaEuraHankkeetVarmallaLinkilla() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const MY_INTEREST_CRITERIA = "Uusiin tekoälyratkaisuihin, digitaaliseen siirtymään ja PK-yritysten kehittämiseen liittyvät hankkeet.";

  // Haetaan API-avaimet
  const props = PropertiesService.getScriptProperties();
  const SCRAPER_API_KEY = props.getProperty("SCRAPER_API_KEY");
  const GEMINI_API_KEY = props.getProperty("GEMINI_API_KEY");

  const targetUrl = "https://eura2021.fi/hakuilmoitukset/";
  const scraperUrl = `https://api.scraperapi.com/?api_key=${SCRAPER_API_KEY}&url=${encodeURIComponent(targetUrl)}&render=true`;

  Logger.log("Haetaan renderöity sivu ScraperAPI:lla...");
  const response = UrlFetchApp.fetch(scraperUrl, { "method": "get", "muteHttpExceptions": true });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`❌ ScraperAPI virhe: HTTP ${response.getResponseCode()}`);
    return;
  }
  
  const html = response.getContentText();

  // 1. Poimitaan kaikki otsikot ja linkit
  const titleRegex = /<td\s+[^>]*title=["']([^"']+)["']/gi;
  const linkRegex = /<a\s+[^>]*href=["'](\/hakuilmoitukset\/hakuilmoitus\/[a-f0-9-]+(?:\/)?)["'][^>]*>(.*?)<\/a>/gi;

  const titles = [];
  let tMatch;
  while ((tMatch = titleRegex.exec(html)) !== null) {
    if (tMatch[1] && tMatch[1].trim().length > 3) {
      titles.push(tMatch[1].trim());
    }
  }

  const links = [];
  let lMatch;
  while ((lMatch = linkRegex.exec(html)) !== null) {
    links.push({
      koodi: lMatch[2].replace(/<[^>]+>/g, "").trim(),
      href: "https://eura2021.fi" + lMatch[1]
    });
  }

  // TÄSSÄ MÄÄRITELLÄÄN loydetytHaut
  const loydetytHaut = links.map((item, index) => ({
    koodi: item.koodi,
    title: titles[index] || item.koodi,
    href: item.href
  }));

  Logger.log(`Löytyi ${loydetytHaut.length} hakuilmoitusta!`);

    // === TÄSSÄ LOKITETAAN GEMINILLE LÄHETETTÄVÄ DATA ===
  Logger.log("================ GEMINILLE LÄHETETTÄVÄ DATA ================");
  Logger.log(JSON.stringify(loydetytHaut, null, 2));
  Logger.log("============================================================");

  if (loydetytHaut.length === 0) return;

  // 2. Pyydetään Geminiltä arvio käyttäen virheensietofunktiota
  const prompt = `Olet asiantuntija. Arvioi seuraavat EURA 2021 -hakuilmoitukset kriteeriin: "${MY_INTEREST_CRITERIA}".

Ilmoitukset:
${JSON.stringify(loydetytHaut, null, 2)}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      response_mime_type: "application/json",
      temperature: 0.1,
      response_schema: {
        type: "OBJECT",
        properties: {
          analyysi: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                koodi: { type: "STRING" },
                title: { type: "STRING" },
                href: { type: "STRING" },
                sopiva: { type: "BOOLEAN" },
                perustelu: { type: "STRING" }
              },
              required: ["koodi", "title", "href", "sopiva", "perustelu"]
            }
          }
        }
      }
    }
  };

  const rawText = kutsuGeminiVarmasti(payload, GEMINI_API_KEY);

  if (rawText) {
    try {
      const tulos = JSON.parse(rawText);
      const lista = tulos.analyysi || [];
      let tallennettu = 0;

      lista.forEach(hanke => {
        if (hanke.sopiva) {
          sheet.appendRow([hanke.koodi, hanke.title, hanke.href, hanke.perustelu, new Date()]);
          Logger.log(`✅ TALLENNETTU: ${hanke.title}`);
          tallennettu++;
        } else {
          Logger.log(`❌ HYLÄTTY: ${hanke.title} (Syy: ${hanke.perustelu})`);
        }
      });

      Logger.log(`🎉 Valmis! Tallennettiin ${tallennettu} kiinnostavaa hanketta taulukkoon.`);
    } catch (e) {
      Logger.log("❌ JSON-parsimisvirhe: " + e.toString());
    }
  } else {
    Logger.log("❌ Epäonnistui: Vastausta ei saatu ruuhkan tai virheen vuoksi.");
  }
}

/**
 * Apufunktio: Yrittää kutsua päämallia uudelleen ruuhkan sattuessa.
 * Jos se epäonnistuu, vaihtaa Lite-varamalliin.
 */
function kutsuGeminiVarmasti(payload, apiKey) {
  const mallit = ["gemini-3.5-flash", "gemini-3.5-flash-lite"]
  const MAX_RETRIES = 2;

  for (const malli of mallit) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${malli}:generateContent?key=${apiKey}`;
    
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      Logger.log(`Yritetään mallia ${malli} (Yritys ${attempt}/${MAX_RETRIES})...`);

      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (statusCode === 200) {
        const json = JSON.parse(responseText);
        return json.candidates?.[0]?.content?.parts?.[0]?.text || null;
      }

      Logger.log(`⚠️ Virhe mallilla ${malli} (HTTP ${statusCode}): ${responseText}`);

      if (statusCode === 503 || statusCode === 429) {
        const odotusaika = attempt * 2000;
        Logger.log(`Ruuhkaa havaittu. Odotetaan ${odotusaika / 1000} sekuntia...`);
        Utilities.sleep(odotusaika);
      } else {
        break;
      }
    }
  }

  return null;
}