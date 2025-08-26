/**
 * MarketDataService (v8.8)
 * -------------------------------------------------------------
 * Librería de datos de mercado para Google Sheets (Apps Script).
 * Devuelve cotizaciones en UNA fila 1×6 con el formato:
 *   { Nombre ; Ticker ; Precio ; Divisa ; Fuente ; FechaISO }
 *
 * Fuentes soportadas:
 *  - YAHOO (API pública de quote)
 *  - COINGECKO (coins/markets)
 *  - INVESTING (HTML scraping robusto)
 *      · Identificador preferido = ISIN si está disponible (fondos/ETFs/bonos)
 *      · Búsqueda por ISIN con ranking por sección (funds > etfs > bonds > equities …)
 *      · Detección de DIVISA contextual al precio (evita falsos USD)
 *  - GOOGLEFINANCE (página pública de Google Finance)
 *
 * Enrutadores:
 *  - resolveQuote(source, identifier, currency)
 *  - resolveQuoteByIsin(isin, hint, strictFunds)
 *
 * Notas:
 *  - CacheService para reducir llamadas (60 s–6 h).
 *  - Uso responsable de fuentes públicas; evita abusos.
 *  - Config ES: en celdas usa ; como separador de argumentos.
 */

/** ===================== UTILIDADES PRIVADAS ===================== **/

/** Cache: lectura segura. */
function _getCache(key) {
  try {
    const c = CacheService.getScriptCache();
    const v = c.get(key);
    return v !== null ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

/** Cache: escritura segura (capado a 6h). */
function _setCache(key, value, seconds) {
  try {
    const c = CacheService.getScriptCache();
    c.put(key, JSON.stringify(value), Math.max(1, Math.min(seconds, 21600)));
  } catch (e) {}
}

/** Coerción numérica tolerante. */
function _toNumber(x) {
  if (x === null || x === undefined || x === "") return "";
  const n = Number(x);
  return isNaN(n) ? "" : n;
}

/** Parser numérico europeo (punto miles, coma decimal). */
function _parseEuropeanNumber(s) {
  if (!s || typeof s !== "string") return "";
  s = s.trim().replace(/[^\d,.\-]/g, "");
  if (s.indexOf(",") > -1 && s.indexOf(".") > -1) s = s.replace(/\./g, "").replace(",", ".");
  else if (s.indexOf(",") > -1) s = s.replace(",", ".");
  const n = Number(s);
  return isNaN(n) ? "" : n;
}

/** Primer número plausible en HTML (fallback). */
function _firstNumberLike(html) {
  const m = html && html.match(/[-+]?\d{1,3}([.,]\d{3})*([.,]\d+)?/);
  return m ? _parseEuropeanNumber(m[0]) : "";
}

/** Timestamp ISO (UTC). */
function _isoNow() { return new Date().toISOString(); }

/** Normaliza ISIN. */
function _normalizeIsin(isin) { return String(isin || "").toUpperCase().replace(/\s+/g, ""); }

/** ¿Se parece a un ISIN? Devuelve el ISIN o "". */
function _looksLikeIsin(s) {
  const m = String(s || "").toUpperCase().match(/\b([A-Z]{2}[A-Z0-9]{9}\d)\b/);
  return m ? m[1] : "";
}

/** ===================== YAHOO FINANCE ===================== **/

/**
 * Quote de Yahoo Finance.
 * @param {string} symbol p.ej. "AAPL", "IWDA.AS"
 * @return {Object[][]} {Nombre;Ticker;Precio;Divisa;Fuente;FechaISO}
 */
function getYahooQuote(symbol) {
  if (!symbol) return [["","","","","YAHOO",_isoNow()]];
  const key = "yfq:" + symbol;
  const cached = _getCache(key);
  if (cached) return cached;

  const url = "https://query1.finance.yahoo.com/v7/finance/quote?symbols=" + encodeURIComponent(symbol);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) return [["","","","","YAHOO",_isoNow()]];

  const data = JSON.parse(res.getContentText());
  const ok = data && data.quoteResponse && data.quoteResponse.result && data.quoteResponse.result.length;
  if (!ok) return [["","","","","YAHOO",_isoNow()]];

  const q = data.quoteResponse.result[0] || {};
  const name = q.longName || q.shortName || "";
  const ticker = q.symbol || String(symbol);
  const price = _toNumber(q.regularMarketPrice || q.postMarketPrice || q.preMarketPrice);
  const currency = (q.currency || "").toUpperCase() || "";
  const row = [[name, ticker, price, currency, "YAHOO", _isoNow()]];
  _setCache(key, row, 90);
  return row;
}

/** Búsqueda en Yahoo por texto/ISIN (autocomplete). */
function _yahooSearch(query) {
  const url = "https://query2.finance.yahoo.com/v1/finance/search?quotesCount=10&newsCount=0&q=" + encodeURIComponent(query);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) return null;
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

/** ===================== COINGECKO ===================== **/

/**
 * Quote de CoinGecko (solo cripto).
 * @param {string} coinId     ej. "bitcoin"
 * @param {string} vsCurrency ej. "eur"
 */
function getCryptoQuote(coinId, vsCurrency) {
  coinId = (coinId || "bitcoin").toLowerCase();
  vsCurrency = (vsCurrency || "eur").toLowerCase();

  const key = "cgq:" + coinId + ":" + vsCurrency;
  const cached = _getCache(key);
  if (cached) return cached;

  const url = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=" +
              encodeURIComponent(vsCurrency) + "&ids=" + encodeURIComponent(coinId);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) return [["","","","","COINGECKO",_isoNow()]];

  const arr = JSON.parse(res.getContentText());
  const it = (arr && arr.length) ? arr[0] : null;
  if (!it) return [["","","","","COINGECKO",_isoNow()]];

  const name = it.name || "";
  const ticker = (it.symbol || "").toUpperCase();
  const price = _toNumber(it.current_price);
  const currency = vsCurrency.toUpperCase();
  const row = [[name, ticker, price, currency, "COINGECKO", _isoNow()]];
  _setCache(key, row, 60);
  return row;
}

/** ===================== INVESTING (HTML) ===================== **/

/**
 * Normaliza identificadores de Investing:
 *  - URL absoluta → se usa tal cual
 *  - ISIN → NO forzar /funds/ aquí (se resolverá mediante búsqueda en getInvestingQuote)
 *  - Slug que empieza por / o por sección conocida → prefijar dominio ES
 *  - Último recurso → prefijar dominio ES
 */
function _normalizeInvestingUrl(id) {
  let s = String(id || "").trim();

  if (/^https?:\/\//i.test(s)) return s;      // URL absoluta
  if (_looksLikeIsin(s)) return s;            // ISIN “pelado”: se resolverá arriba
  if (s.startsWith("/")) return "https://es.investing.com" + s;

  if (/^(funds|etfs|equities|indices|currencies|crypto|bonds|certificates)\//i.test(s)) {
    return "https://es.investing.com/" + s;
  }
  return "https://es.investing.com/" + s;
}

/** Descarga HTML de Investing con headers “browser‑like”. */
function _fetchInvestingHtml(url) {
  return UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"
    }
  });
}

/** Extractores de ficha Investing (nombre / ticker / ISIN / precio / moneda). */
function _extractInvestingName(html) {
  let m = html.match(/<h1[^>]*>\s*([^<][^<]+)\s*<\/h1>/i);
  if (m && m[1]) return m[1].trim();
  m = html.match(/"name"\s*:\s*"([^"]{3,100})"/i);
  if (m && m[1]) return m[1].trim();
  m = html.match(/<title[^>]*>\s*([^<][^<]+)\s*<\/title>/i);
  if (m && m[1]) return m[1].replace(/\s*\|\s*Investing.*$/i,"").trim();
  return "";
}

function _extractInvestingTicker(html) {
  let m = html.match(/"symbol"\s*:\s*"([A-Za-z0-9\-.]{1,20})"/i);
  if (m && m[1]) return m[1].toUpperCase();
  m = html.match(/data-test=["']symbol["'][^>]*>\s*([^<\s][^<]{1,20})</i);
  if (m && m[1]) return m[1].trim().toUpperCase();
  return "";
}

function _extractInvestingISIN(html) {
  const m = html.match(/\b([A-Z]{2}[A-Z0-9]{9}\d)\b/);
  return m ? m[1] : "";
}

/** Localiza el “ancla” del precio para crear una ventana de contexto fiable. */
function _findPriceAnchorIndex(html) {
  const regs = [
    /data-test=["']instrument-price-last["'][^>]*>/i,
    /data-test=["']last-price-value["'][^>]*>/i,
    /id=["']last_last["'][^>]*>/i,
    /class=["'][^"']*last-price[^"']*["'][^>]*>/i
  ];
  for (let i = 0; i < regs.length; i++) {
    const m = regs[i].exec(html);
    if (m) return m.index;
  }
  return -1;
}

function _extractInvestingPrice(html) {
  const patterns = [
    /data-test=["']instrument-price-last["'][^>]*>\s*([^<\s][^<]*)</i,
    /data-test=["']last-price-value["'][^>]*>\s*([^<\s][^<]*)</i,
    /<span[^>]*id=["']last_last["'][^>]*>\s*([^<\s][^<]*)</i,
    /"last_price"\s*:\s*"([^"]+)"/i,
    /"last_last"\s*:\s*"([^"]+)"/i,
    /class=["']last-price[^"']*["'][^>]*>\s*<[^>]*>\s*([^<\s][^<]*)</i,
    /<span[^>]*class=["'][^"']*(?:text-5xl|text-2xl)[^"']*["'][^>]*>\s*([^<\s][^<]*)</i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = html.match(patterns[i]);
    if (m && m[1]) {
      const n = _parseEuropeanNumber(m[1]);
      if (n !== "") return n;
    }
  }
  const fb = _firstNumberLike(html);
  return fb === "" ? "" : fb;
}

/**
 * Detección de divisa “precio‑céntrica” para evitar falsos positivos:
 *  1) Intenta campos estructurados (priceCurrency / currency / quoted_currency / microdatos)
 *  2) Ventana alrededor del precio: códigos ISO o símbolos cercanos
 *  3) Etiquetas “Divisa/Moneda/Currency” en la cabecera de ficha
 *  4) Si no hay señal fiable, devuelve "" (mejor vacío que incorrecto)
 */
function _extractInvestingCurrency(html) {
  let m = html.match(/"priceCurrency"\s*:\s*"([A-Z]{3})"/i);
  if (m && m[1]) return m[1].toUpperCase();
  m = html.match(/"currency"\s*:\s*"([A-Z]{3})"/i);
  if (m && m[1]) return m[1].toUpperCase();
  m = html.match(/"quoted_?currency"\s*:\s*"([A-Z]{3})"/i);
  if (m && m[1]) return m[1].toUpperCase();
  m = html.match(/itemprop=["']priceCurrency["'][^>]*content=["']([A-Z]{3})["']/i);
  if (m && m[1]) return m[1].toUpperCase();

  const idx = _findPriceAnchorIndex(html);
  if (idx >= 0) {
    const win = html.substring(Math.max(0, idx - 600), Math.min(html.length, idx + 900));
    let mc = win.match(/\b(EUR|USD|GBP|JPY|CHF|AUD|CAD|CNY|SEK|NOK|DKK)\b/i);
    if (mc && mc[1]) return mc[1].toUpperCase();
    if (/[€]/.test(win)) return "EUR";
    if (/\$/.test(win))  return "USD";
    if (/£/.test(win))   return "GBP";
    if (/¥/.test(win))   return "JPY";
  }

  const headWin = html.substring(0, Math.min(html.length, 30000));
  let reKV = new RegExp("(Divisa|Moneda|Currency)[^A-Z]{0,40}\\b(EUR|USD|GBP|JPY|CHF|AUD|CAD|CNY|SEK|NOK|DKK)\\b", "i");
  m = headWin.match(reKV);
  if (m && m[2]) return m[2].toUpperCase();
  reKV = new RegExp("\\b(EUR|USD|GBP|JPY|CHF|AUD|CAD|CNY|SEK|NOK|DKK)\\b[^A-Z]{0,40}(Divisa|Moneda|Currency)", "i");
  m = headWin.match(reKV);
  if (m && m[1]) return m[1].toUpperCase();

  return "";
}

/** Extrae ISIN de una URL de ficha (funds/etfs/equities/bonds/certificates). */
function _extractIsinFromInvestingUrl(url) {
  const m = String(url || "").toUpperCase().match(/\/(funds|etfs|equities|bonds|certificates)\/([A-Z]{2}[A-Z0-9]{9}\d)\b/);
  return m ? m[2] : "";
}

/** ===================== BÚSQUEDA EN INVESTING POR ISIN ===================== **/

/** Página de resultados de búsqueda de Investing para una query. */
function _investingSearchHtml(query) {
  const url = "https://es.investing.com/search/?q=" + encodeURIComponent(query);
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"
    }
  });
  if (res.getResponseCode() !== 200) return "";
  return res.getContentText();
}

/**
 * Selección del mejor enlace por ISIN:
 *  - Extrae enlaces con sección conocida
 *  - Puntuación por sección (funds>etfs>bonds>equities>indices>currencies>crypto>certificates)
 *  - Bonus si el ISIN aparece en el contexto del enlace y si el hint encaja
 *  - Penaliza listados
 */
function _investingPickBestLinkByIsin(html, isin, hint) {
  if (!html) return "";

  const sectRe = /(\/(funds|etfs|bonds|equities|indices|currencies|crypto|certificates)\/[^"#?]+)"/ig;
  const candidates = [];
  let m;
  while ((m = sectRe.exec(html)) !== null) {
    const href = m[1];
    const section = m[2].toLowerCase();
    candidates.push({ href, section, idx: m.index });
    if (candidates.length > 50) break;
  }
  if (!candidates.length) return "";

  const hintTokens = String(hint || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s\-]/g, " ")
    .split(/\s+/).filter(Boolean);

  const SECTION_PRIORITY = {
    "funds": 90, "etfs": 80, "bonds": 70, "equities": 60,
    "indices": 40, "currencies": 30, "crypto": 20, "certificates": 50
  };

  function score(c) {
    let s = SECTION_PRIORITY[c.section] || 10;

    const start = Math.max(0, c.idx - 220);
    const end   = Math.min(html.length, c.idx + 220);
    const ctx = html.substring(start, end).toLowerCase();

    if (isin && ctx.indexOf(isin.toLowerCase()) !== -1) s += 50;

    if (hintTokens.length) {
      const slug = c.href.toLowerCase().replace(/^\/(funds|etfs|bonds|equities|indices|currencies|crypto|certificates)\//, "");
      const slugTokens = slug.split(/[-\/]+/).filter(Boolean);
      for (const t of hintTokens) {
        if (t.length >= 3 && (ctx.indexOf(t) !== -1 || slugTokens.some(st => st.indexOf(t) !== -1))) s += 6;
      }
    }

    if (/\/(world-|most-|top-)/i.test(c.href)) s -= 10;
    return s;
  }

  candidates.sort((a, b) => score(b) - score(a));
  return candidates[0] ? ("https://es.investing.com" + candidates[0].href) : "";
}

/** ===================== QUOTE DE INVESTING ===================== **/

/**
 * Obtener cotización desde Investing:
 *  - Si identifier es URL → usarla directamente
 *  - Si identifier parece ISIN → buscar mejor ficha por sección (ranking) y usarla
 *  - Si identifier es slug/texto → normalizar URL o buscar por texto
 *  Identificador preferido (Ticker): ISIN (HTML o URL) > symbol
 */
function getInvestingQuote(identifier) {
  if (!identifier) return [["","","","","INVESTING",_isoNow()]];

  const raw = String(identifier).trim();
  const possibleIsin = _looksLikeIsin(raw);

  let url = "";
  if (/^https?:\/\//i.test(raw)) {
    url = raw;
  } else if (possibleIsin) {
    const cacheKeyFind = "inv_find:" + possibleIsin;
    const cachedFind = _getCache(cacheKeyFind);
    if (cachedFind) {
      url = cachedFind;
    } else {
      const htmlSearch = _investingSearchHtml(possibleIsin);
      url = _investingPickBestLinkByIsin(htmlSearch, possibleIsin, "");
      if (url) _setCache(cacheKeyFind, url, 600);
    }
    if (!url) {
      // Último recurso para algunos catálogos: ficha “funds/ISIN”
      url = "https://es.investing.com/funds/" + possibleIsin;
    }
  } else {
    const norm = _normalizeInvestingUrl(raw);
    if (/^https?:\/\//i.test(norm)) {
      url = norm;
    } else {
      const htmlSearch = _investingSearchHtml(raw);
      url = _investingPickBestLinkByIsin(htmlSearch, "", raw) || "";
      if (!url) url = "https://es.investing.com/" + raw;
    }
  }

  const key = "invq:" + url;
  const cached = _getCache(key);
  if (cached) return cached;

  let res = _fetchInvestingHtml(url);
  let html = (res.getResponseCode() === 200) ? res.getContentText() : "";

  if (!html) {
    const mobile = url.replace("://es.", "://m.").replace("://www.", "://m.");
    res = _fetchInvestingHtml(mobile);
    html = (res.getResponseCode() === 200) ? res.getContentText() : "";
  }
  if (!html) {
    const global = url.replace("://es.", "://www.");
    res = _fetchInvestingHtml(global);
    html = (res.getResponseCode() === 200) ? res.getContentText() : "";
  }

  if (!html) return [["","","","","INVESTING",_isoNow()]];

  const name = _extractInvestingName(html);
  const price = _extractInvestingPrice(html);
  const currency = _extractInvestingCurrency(html);

  // Ticker preferido: ISIN (HTML o URL) > symbol
  let ticker = "";
  const isinHtml = _extractInvestingISIN(html);
  const isinUrl  = _extractIsinFromInvestingUrl(url);
  if (isinHtml)       ticker = isinHtml;
  else if (isinUrl)   ticker = isinUrl;
  else                ticker = _extractInvestingTicker(html) || "";

  const row = [[name, ticker, price, currency, "INVESTING", _isoNow()]];
  _setCache(key, row, 60);
  return row;
}

/** ===================== GOOGLE FINANCE (HTML) ===================== **/

function _normalizeGoogleSymbol(sym) { return String(sym || "").trim(); }

function _currencyFromSymbols_(html) {
  if (!html) return "";
  if (/[€]/.test(html)) return "EUR";
  if (/\$/.test(html)) return "USD";
  if (/£/.test(html)) return "GBP";
  return "";
}

/** Quote desde Google Finance (scraping de página pública). */
function getGoogleFinanceQuote(symbol) {
  var sym = _normalizeGoogleSymbol(symbol);
  if (!sym) return [["","","","","GOOGLEFINANCE", _isoNow()]];

  var key = "gfq:" + sym;
  var cached = _getCache(key);
  if (cached) return cached;

  var attempts = [sym];
  if (sym.indexOf(":") > -1) {
    var parts = sym.split(":");
    if (parts.length === 2) attempts.push(parts[1] + ":" + parts[0]);
  }

  for (var i = 0; i < attempts.length; i++) {
    var s = attempts[i];
    var url = "https://www.google.com/finance/quote/" + encodeURIComponent(s) + "?hl=en";
    try {
      var res = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
          "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"
        }
      });
      if (res.getResponseCode() !== 200) continue;

      var html = res.getContentText();

      var name = "";
      var m = html.match(/"companyName"\s*:\s*"([^"]{3,120})"/i);
      if (m && m[1]) name = m[1].trim();
      if (!name) { m = html.match(/"name"\s*:\s*"([^"]{3,120})"/i); if (m && m[1]) name = m[1].trim(); }
      if (!name) { m = html.match(/<title[^>]*>\s*([^<][^<]+?)\s*[–—-]\s*Google Finance\s*<\/title>/i); if (m && m[1]) name = m[1].trim(); }

      var price = "";
      m = html.match(/"price"\s*:\s*"([0-9.,]+)"/i);
      if (m && m[1]) price = _parseEuropeanNumber(m[1]);
      if (price === "") { m = html.match(/"regularMarketPrice"\s*:\s*([0-9.]+)/i); if (m && m[1]) price = _toNumber(m[1]); }
      if (price === "") { m = html.match(/data-last-price=["']([0-9.,]+)["']/i); if (m && m[1]) price = _parseEuropeanNumber(m[1]); }
      if (price === "") price = _firstNumberLike(html);

      var currency = "";
      m = html.match(/"currency"\s*:\s*"([A-Z]{3})"/i);
      if (m && m[1]) currency = m[1].toUpperCase();
      if (!currency) { m = html.match(/"priceCurrency"\s*:\s*"([A-Z]{3})"/i); if (m && m[1]) currency = m[1].toUpperCase(); }
      if (!currency) { m = html.match(/itemprop=["']priceCurrency["'][^>]*content=["']([A-Z]{3})["']/i); if (m && m[1]) currency = m[1].toUpperCase(); }
      if (!currency) currency = _currencyFromSymbols_(html);

      var ticker = "";
      m = html.match(/"symbol"\s*:\s*"([A-Za-z0-9:.\-]{2,40})"/i);
      if (m && m[1]) ticker = m[1].toUpperCase();
      if (!ticker) ticker = s.toUpperCase();

      var row = [[name, ticker, price, currency, "GOOGLEFINANCE", _isoNow()]];
      if (name || (price !== "")) {
        _setCache(key, row, 60);
        return row;
      }
    } catch (e) { /* siguiente intento */ }
  }

  return [["","","","","GOOGLEFINANCE", _isoNow()]];
}

/** ===================== ENRUTADORES ===================== **/

/**
 * Enrutador por fuente.
 * @param {string} source "YAHOO" | "COINGECKO" | "INVESTING" | "GOOGLEFINANCE" | "GOOGLE"
 * @param {string} identifier ticker / id / url / isin según la fuente
 * @param {string} currency  solo aplica a COINGECKO (p. ej. "EUR")
 */
function resolveQuote(source, identifier, currency) {
  source = String(source || "").toUpperCase();
  identifier = String(identifier || "");
  currency = String(currency || "EUR").toUpperCase();

  if (source === "YAHOO")              return getYahooQuote(identifier);
  if (source === "COINGECKO")          return getCryptoQuote(identifier, currency);
  if (source === "INVESTING")          return getInvestingQuote(identifier);
  if (source === "GOOGLEFINANCE" || source === "GOOGLE")
                                       return getGoogleFinanceQuote(identifier);

  return [["","","","","", _isoNow()]];
}

/**
 * Búsqueda por ISIN:
 *  - strictFunds = true → consulta Investing directamente (prioriza “funds”)
 *  - strictFunds = false/omitido → intenta Yahoo y, si no es usable, Investing con ranking
 *  - Si "hint" es una URL de Investing, se usa directamente (atajo para mapeos específicos)
 *
 * @param {string}  isin        Código ISIN (ej.: "FR00140081Y1")
 * @param {string}  hint        (opcional) Texto de pista o URL canónica de Investing
 * @param {boolean} strictFunds (opcional) TRUE para forzar Investing (ideal fondos UCITS)
 */
function resolveQuoteByIsin(isin, hint, strictFunds) {
  const code = _normalizeIsin(isin);
  const hintStr = String(hint || "").trim();
  const key  = "isinq:" + code + ":" + hintStr.toLowerCase().replace(/\s+/g,"-") + ":" + String(strictFunds||false);
  const cached = _getCache(key);
  if (cached) return cached;

  // Si la pista es una URL de Investing, úsala directamente
  if (/^https?:\/\/[^ ]*investing\.com/i.test(hintStr)) {
    const rowDirect = getInvestingQuote(hintStr);
    _setCache(key, rowDirect, 300);
    return rowDirect;
  }

  if (strictFunds === true) {
    const rowStrict = _resolveIsinViaInvesting_(code, hintStr);
    _setCache(key, rowStrict, 300);
    return rowStrict;
  }

  // 1) Yahoo (acciones/ETF) — solo aceptar si trae precio+divisa
  try {
    const y = _yahooSearch(code);
    if (y && y.quotes && y.quotes.length) {
      const hit = y.quotes.find(q => {
        const t = (q.quoteType || "").toUpperCase();
        return ["FUTURE","OPTION","INDEX","CURRENCY"].indexOf(t) === -1;
      }) || y.quotes[0];

      if (hit && hit.symbol) {
        const rowY = getYahooQuote(hit.symbol);
        if (rowY && rowY[0] && rowY[0][2] !== "" && rowY[0][3] !== "") {
          _setCache(key, rowY, 300);
          return rowY;
        }
      }
    }
  } catch (e) { /* continuar con Investing */ }

  // 2) Investing con ranking por sección
  const row = _resolveIsinViaInvesting_(code, hintStr);
  _setCache(key, row, 300);
  return row;
}

/**
 * Resolución auxiliar por ISIN usando Investing con ranking por sección.
 * Si no hay match en la búsqueda y el hint es una URL de Investing, se usa esa URL.
 * Último recurso: intentar /funds/ISIN (no siempre existirá).
 */
function _resolveIsinViaInvesting_(code, hint) {
  try {
    const html = _investingSearchHtml(code);
    const bestUrl = _investingPickBestLinkByIsin(html, code, hint || "");
    if (bestUrl) return getInvestingQuote(bestUrl);

    if (hint && /^https?:\/\/[^ ]*investing\.com/i.test(hint)) {
      return getInvestingQuote(hint);
    }
    return getInvestingQuote("https://es.investing.com/funds/" + code);
  } catch (e) {
    if (hint && /^https?:\/\/[^ ]*investing\.com/i.test(hint)) {
      try { return getInvestingQuote(hint); } catch (e2) {}
    }
    return [["","","","","",_isoNow()]];
  }
}
