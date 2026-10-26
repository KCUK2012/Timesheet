/*
 * Outlook Timesheet: sign-in and lists (clients, internal codes, projects).
 * Uses Microsoft's nested app authentication (NAA) through MSAL.js, so staff are
 * signed in silently with the Microsoft 365 account they already use in Outlook.
 * Reads the lists from SharePoint (an .xlsx workbook, a CSV file or a SharePoint list) through
 * Microsoft Graph with delegated permission: staff can only read what their own SharePoint access allows.
 */
(function (root) {
  "use strict";

  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["https://graph.microsoft.com/Sites.Read.All"];
  var pcaPromise = null;
  var stage = "not started";
  function setStage(t) { stage = t; }

  function getApp(cfg) {
    if (!pcaPromise) {
      var base = location.href.replace(/[^/]*$/, "");
      pcaPromise = root.msal.createNestablePublicClientApplication({
        auth: {
          clientId: cfg.clientId,
          authority: "https://login.microsoftonline.com/" + cfg.tenantId,
          redirectUri: base + "redirect.html"
        },
        cache: { cacheLocation: "sessionStorage" }  // fresh token cache for each Outlook session
      });
    }
    return pcaPromise;
  }

  function describe(e) {
    if (!e) return "unknown error";
    return [e.errorCode, e.message].filter(Boolean).join(": ");
  }

  async function getToken(cfg, forceRefresh) {
    var app;
    setStage("Step 1 (start sign-in)");
    try { app = await getApp(cfg); }
    catch (e) { throw new Error("Step 1 (start sign-in) failed. " + describe(e)); }
    var request = { scopes: SCOPES, forceRefresh: !!forceRefresh };
    setStage("Step 2 (sign-in)");
    try {
      var silent = await app.acquireTokenSilent(request);
      return silent.accessToken;
    } catch (e1) {
      try {
        // Consent needed or session expired: ask the user once.
        var interactive = await app.acquireTokenPopup(request);
        return interactive.accessToken;
      } catch (e2) {
        throw new Error("Step 2 (sign-in) failed. Silent: " + describe(e1) + " | Pop-up: " + describe(e2));
      }
    }
  }

  async function graphGet(url, token, step) {
    var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!res.ok) {
      var detail = "";
      try { var body = await res.json(); detail = body && body.error ? (body.error.code + ": " + body.error.message) : ""; } catch (x) {}
      var hint = (res.status === 403 || res.status === 404)
        ? " You may not have access, or the site/file path in config.js is wrong."
        : "";
      var err = new Error(step + " failed (HTTP " + res.status + ")." + hint + (detail ? " Details: " + detail : ""));
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  function splitAliases(text) {
    return String(text || "")
      .split(/[;\n]/)
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s.length > 0; });
  }

  /** Minimal CSV parser: handles quoted fields, commas and line breaks in quotes. */
  function parseCsv(text) {
    text = String(text || "").replace(/^\uFEFF/, "");
    var rows = [], row = [], field = "", inQuotes = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
        } else { field += ch; }
      } else if (ch === '"') { inQuotes = true; }
      else if (ch === ",") { row.push(field); field = ""; }
      else if (ch === "\n" || ch === "\r") {
        if (ch === "\r" && text[i + 1] === "\n") i++;
        row.push(field); rows.push(row); row = []; field = "";
      } else { field += ch; }
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (c) { return c.trim() !== ""; }); });
  }

  function findColumn(headers, preferred, fallbacks) {
    var wanted = [preferred].concat(fallbacks).map(function (h) { return String(h).toLowerCase(); });
    for (var w = 0; w < wanted.length; w++) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i].trim().toLowerCase() === wanted[w]) return i;
      }
    }
    return -1;
  }

  function clientsFromCsv(text, cfg) {
    var rows = parseCsv(text);
    if (!rows.length) throw new Error("The client file in SharePoint is empty.");
    var headers = rows[0];
    var nameIdx = findColumn(headers, cfg.nameColumn, ["Title", "Client", "Client name", "Name"]);
    var aliasIdx = findColumn(headers, cfg.aliasColumn, ["Aliases", "Alias", "Recognised", "Short forms"]);
    if (nameIdx < 0) {
      throw new Error('The client file needs a column headed "' + cfg.nameColumn + '" (or "Title").');
    }
    var clients = [];
    rows.slice(1).forEach(function (r) {
      var name = String(r[nameIdx] || "").trim();
      if (name) clients.push({ name: name, aliases: aliasIdx >= 0 ? splitAliases(r[aliasIdx]) : [] });
    });
    return clients;
  }

  async function getSite(cfg, token) {
    var sitePath = (cfg.sitePath || "").trim();
    setStage("Step 3 (find SharePoint site)");
    var siteUrl = GRAPH + "/sites/" + cfg.sharePointHost + (sitePath ? ":" + encodeURI(sitePath) : "");
    return graphGet(siteUrl, token, "Step 3 (find SharePoint site)");
  }

  /** Reads a CSV file from the site's default document library. */
  async function loadFromFile(cfg, token, site) {
    var path = "/" + String(cfg.filePath || "").replace(/^\/+/, "");
    var encoded = path.split("/").map(function (seg) {
      return encodeURIComponent(seg).replace(/'/g, "%27");
    }).join("/");
    // Graph's /content endpoint redirects, which browsers block for authorised
    // cross-origin requests, so ask for the short-lived download URL instead.
    setStage("Step 4 (find list file)");
    // Request the item without $select: SharePoint omits the download link when it is selected explicitly.
    var item = await graphGet(GRAPH + "/sites/" + site.id + "/drive/root:" + encoded, token, "Step 4 (find list file)");
    var dl = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"];
    setStage("Step 5 (download list file)");
    var res;
    if (dl) {
      // Pre-authorised, short-lived link: no Authorization header needed.
      res = await fetch(dl);
    } else {
      // Fallback: ask Graph for the file content by item ID.
      try {
        res = await fetch(GRAPH + "/sites/" + site.id + "/drive/items/" + item.id + "/content",
          { headers: { Authorization: "Bearer " + token } });
      } catch (e) {
        throw new Error("Step 5 (download list file) failed: no download link returned, and direct download was blocked. Fields received: " +
          Object.keys(item).join(", "));
      }
    }
    if (!res.ok) throw new Error("Step 5 (download list file) failed (HTTP " + res.status + ")." +
      (dl ? "" : " No download link was returned; fields received: " + Object.keys(item).join(", ")));
    if (/\.xlsx$/i.test(path)) {
      setStage("Step 6 (read workbook)");
      return listsFromWorkbook(await root.KCXlsx.read(await res.arrayBuffer()), cfg);
    }
    return { clients: clientsFromCsv(await res.text(), cfg), internalCodes: null, projects: null };
  }

  /** Finds a tab by name, ignoring case and spaces at either end. */
  function findSheet(book, name) {
    var want = String(name).trim().toLowerCase();
    for (var k in book) if (k.trim().toLowerCase() === want) return book[k];
    return null;
  }
  function dropBlankRows(rows) {
    return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ""; }); });
  }

  /** Turns the workbook's tabs into the add-in's lists. */
  function listsFromWorkbook(book, cfg) {
    var tabs = cfg.sheets || {};
    var clientRows = findSheet(book, tabs.clients || "Clients");
    if (!clientRows) throw new Error('The workbook has no "' + (tabs.clients || "Clients") + '" tab.');
    clientRows = dropBlankRows(clientRows);
    if (!clientRows.length) throw new Error("The Clients tab is empty.");
    var h = clientRows[0];
    var nameIdx = findColumn(h, cfg.nameColumn, ["Title", "Client", "Client name", "Name"]);
    var aliasIdx = findColumn(h, cfg.aliasColumn, ["Aliases", "Alias"]);
    if (nameIdx < 0) throw new Error('The Clients tab needs a column headed "' + cfg.nameColumn + '".');
    var clients = [];
    clientRows.slice(1).forEach(function (r) {
      var name = String(r[nameIdx] || "").trim();
      if (name) clients.push({ name: name, aliases: aliasIdx >= 0 ? splitAliases(r[aliasIdx]) : [] });
    });

    var internalCodes = null;
    var codeRows = findSheet(book, tabs.internalCodes || "Internal codes");
    if (codeRows) {
      codeRows = dropBlankRows(codeRows);
      var ch = codeRows[0] || [];
      var codeIdx = findColumn(ch, "Code", ["Internal code", "Name"]);
      var hintIdx = findColumn(ch, "Hint", ["Help", "Description"]);
      var nftIdx = findColumn(ch, "Not for timesheet", ["NFT"]);
      if (codeIdx >= 0) {
        internalCodes = [];
        codeRows.slice(1).forEach(function (r) {
          var code = String(r[codeIdx] || "").trim();
          if (code) internalCodes.push({
            label: code, subject: code,
            hint: hintIdx >= 0 ? String(r[hintIdx] || "").trim() : "",
            nft: nftIdx >= 0 && /^(y|yes|true|1)$/i.test(String(r[nftIdx] || "").trim())
          });
        });
        if (!internalCodes.length) internalCodes = null;
      }
    }

    var projects = null;
    var projRows = findSheet(book, tabs.projects || "Projects");
    if (projRows) {
      projRows = dropBlankRows(projRows);
      var ph = projRows[0] || [];
      var areaIdx = findColumn(ph, "Area", []);
      var projIdx = findColumn(ph, "Project", ["Project name", "Name"]);
      if (projIdx >= 0) {
        projects = [];
        projRows.slice(1).forEach(function (r) {
          var pn = String(r[projIdx] || "").trim();
          if (pn) projects.push({ area: areaIdx >= 0 ? String(r[areaIdx] || "").trim() : "", name: pn });
        });
        if (!projects.length) projects = null;
      }
    }
    return { clients: clients, internalCodes: internalCodes, projects: projects };
  }

  /** Reads a SharePoint list. */
  async function loadFromList(cfg, token, site) {
    var listQuery = GRAPH + "/sites/" + site.id + "/lists?$select=id,displayName&$filter=displayName eq '" +
      String(cfg.listName).replace(/'/g, "''") + "'";
    var lists = await graphGet(listQuery, token, "Step 4 (find list)");
    if (!lists.value || !lists.value.length) {
      throw new Error('The SharePoint list "' + cfg.listName + '" was not found.');
    }
    var url = GRAPH + "/sites/" + site.id + "/lists/" + lists.value[0].id +
      "/items?$top=500&$expand=fields($select=" + cfg.nameColumn + "," + cfg.aliasColumn + ")";
    var clients = [];
    while (url) {
      var page = await graphGet(url, token, "Step 5 (read list)");
      (page.value || []).forEach(function (item) {
        var f = item.fields || {};
        var name = String(f[cfg.nameColumn] || "").trim();
        if (name) clients.push({ name: name, aliases: splitAliases(f[cfg.aliasColumn]) });
      });
      url = page["@odata.nextLink"] || null;
    }
    return clients;
  }

  /**
   * Returns { clients: [{ name, aliases[] }], internalCodes: [...] or null, projects: [{ area, name }] or null }.
   * internalCodes and projects are null when the source does not hold them (CSV file or list).
   */
  async function loadListsWith(cfg, token) {
    var site = await getSite(cfg, token);
    if (cfg.source === "list") {
      return { clients: await loadFromList(cfg, token, site), internalCodes: null, projects: null };
    }
    return loadFromFile(cfg, token, site);
  }

  /** Reads the start and expiry times inside a sign-in token, to explain a 401. */
  function tokenTimes(token) {
    try {
      var part = String(token).split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
      while (part.length % 4) part += "=";
      var claims = JSON.parse(atob(part));
      function iso(sec) { return sec ? new Date(sec * 1000).toISOString().replace(".000", "") : "?"; }
      return "Sign-in valid from " + iso(claims.nbf) + " to " + iso(claims.exp) +
        " (UTC). This computer's time now: " + new Date().toISOString().replace(/\.\d+Z$/, "Z") + " (UTC).";
    } catch (x) { return "Sign-in times could not be read."; }
  }

  async function loadLists(cfg) {
    var token = await getToken(cfg);
    try {
      return await loadListsWith(cfg, token);
    } catch (e) {
      if (!e || e.status !== 401) throw e;
    }
    // 401: the sign-in was rejected as expired. Ask for a brand-new one and try again.
    token = await getToken(cfg, true);
    try {
      return await loadListsWith(cfg, token);
    } catch (e2) {
      if (!e2 || e2.status !== 401) throw e2;
    }
    // Still rejected: try an interactive sign-in once, then report the token times.
    try {
      var app = await getApp(cfg);
      var popup = await app.acquireTokenPopup({ scopes: SCOPES, prompt: "select_account" });
      token = popup.accessToken;
      return await loadListsWith(cfg, token);
    } catch (e3) {
      throw new Error("Microsoft keeps rejecting the sign-in as expired. " + tokenTimes(token) +
        " If the two times do not overlap, check the computer's date, time and time zone. " +
        (e3 && e3.message ? "Last error: " + e3.message : ""));
    }
  }

  root.KCData = {
    getStage: function () { return stage; },
    loadLists: loadLists,
    listsFromWorkbook: listsFromWorkbook,
    splitAliases: splitAliases, parseCsv: parseCsv, clientsFromCsv: clientsFromCsv
  };
})(window);
