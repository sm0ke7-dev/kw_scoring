/**
 * GenerateKeywords.gs
 * Keyword generation module for creating keyword variations from templates, niches, and locations
 */

function GENERATE_KEYWORDS(templates, niche, niche_placeholder, locations, location_placeholder) {
    const toStr = v => (v === null || v === undefined) ? "" : String(v).trim();
  
    function normalize1D(arg) {
      if (!Array.isArray(arg)) return [toStr(arg)].filter(Boolean);
      const out = [];
      if (Array.isArray(arg[0])) {
        for (let r = 0; r < arg.length; r++) {
          for (let c = 0; c < arg[r].length; c++) {
            const v = toStr(arg[r][c]);
            if (v) out.push(v);
          }
        }
      } else {
        for (let i = 0; i < arg.length; i++) {
          const v = toStr(arg[i]);
          if (v) out.push(v);
        }
      }
      return out;
    }
  
    function normalizeNiche2Cols(arg) {
      if (!Array.isArray(arg)) {
        const v = toStr(arg);
        if (!v) return [];
        throw new Error("The 'niche' parameter must be a 2-column range: [service, core_keyword].");
      }
      const rows = Array.isArray(arg[0]) ? arg : [arg];
      const out = [];
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!Array.isArray(row) || row.length < 2) continue;
        const service = toStr(row[0]);
        const core   = toStr(row[1]);
        if (service && core) out.push([service, core]);
      }
      return out;
    }
  
    function replaceAllLiteral(str, find, replace) {
      if (!find) return str;
      return String(str).split(find).join(replace);
    }
  
    const tplList   = normalize1D(templates);
    const locList   = normalize1D(locations);
    const nicheRows = normalizeNiche2Cols(niche);
  
    const nichePh = toStr(niche_placeholder);
    const locPh   = toStr(location_placeholder);
  
    if (!tplList.length)   throw new Error("No templates provided.");
    if (!nicheRows.length) throw new Error("No niche rows provided. Expect two columns: service, core_keyword.");
    if (!locList.length)   throw new Error("No locations provided.");
    if (!nichePh)          throw new Error("niche_placeholder is empty.");
    if (!locPh)            throw new Error("location_placeholder is empty.");
  
    const out = [];
    // out.push(["service", "location", "core_keyword", "keyword"]); // optional header
  
    // Group by Location → Service → Template
    for (let j = 0; j < locList.length; j++) {
      const location = locList[j];
      for (let i = 0; i < nicheRows.length; i++) {
        const [service, core_keyword] = nicheRows[i];
        for (let t = 0; t < tplList.length; t++) {
          const template = tplList[t];
          let keyword = template;
          keyword = replaceAllLiteral(keyword, nichePh, core_keyword);
          keyword = replaceAllLiteral(keyword, locPh, location);
          keyword = keyword.toLowerCase(); // 👈 lowercase final keyword
          out.push([service, location, core_keyword, keyword]);
        }
      }
    }
  
    return out;
  }

