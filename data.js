/*
 * KC Timesheet: sign-in and client list.
 * Uses Microsoft's nested app authentication (NAA) through MSAL.js, so staff are
 * signed in silently with the Microsoft 365 account they already use in Outlook.
 * Reads the client list (a CSV file or a list) from SharePoint through Microsoft Graph with delegated
 * permission: staff can only read what their own SharePoint access allows.
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
        cache: { cacheLocation: "localStorage" }
      });
    }
    return pcaPromise;
  }

  function describe(e) {
    if (!e) return "unknown error";
    return [e.errorCode, e.message].filter(Boolean).join(": ");
  }

  async function getToken(cfg) {
    var app;
    setStage("Step 1 (start sign-in)");
    try { app = await getApp(cfg); }
    catch (e) { throw new Error("Step 1 (start sign-in) failed. " + describe(e)); }
    var request = { scopes: SCOPES };
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
      throw new Error(step + " failed (HTTP " + res.status + ")." + hint + (detail ? " Details: " + detail : ""));
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
    setStage("Step 4 (find client file)");
    // Request the item without $select: SharePoint omits the download link when it is selected explicitly.
    var item = await graphGet(GRAPH + "/sites/" + site.id + "/drive/root:" + encoded, token, "Step 4 (find client file)");
    var dl = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"];
    setStage("Step 5 (download client file)");
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
        throw new Error("Step 5 (download client file) failed: no download link returned, and direct download was blocked. Fields received: " +
          Object.keys(item).join(", "));
      }
    }
    if (!res.ok) throw new Error("Step 5 (download client file) failed (HTTP " + res.status + ")." +
      (dl ? "" : " No download link was returned; fields received: " + Object.keys(item).join(", ")));
    return clientsFromCsv(await res.text(), cfg);
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

  /** Returns [{ name, aliases[] }] from SharePoint (CSV file or list, per config). */
  async function loadClients(cfg) {
    var token = await getToken(cfg);
    var site = await getSite(cfg, token);
    return cfg.source === "list" ? loadFromList(cfg, token, site) : loadFromFile(cfg, token, site);
  }

  root.KCData = { getStage: function () { return stage; }, loadClients: loadClients, splitAliases: splitAliases, parseCsv: parseCsv, clientsFromCsv: clientsFromCsv };
})(window);
