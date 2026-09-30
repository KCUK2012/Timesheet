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

  async function getToken(cfg) {
    var app = await getApp(cfg);
    var request = { scopes: SCOPES };
    try {
      var silent = await app.acquireTokenSilent(request);
      return silent.accessToken;
    } catch (e) {
      // Consent needed or session expired: ask the user once.
      var interactive = await app.acquireTokenPopup(request);
      return interactive.accessToken;
    }
  }

  async function graphGet(url, token) {
    var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (res.status === 403 || res.status === 404) {
      throw new Error("You do not have access to the client list in SharePoint, or it could not be found. Please contact the owner of the client list.");
    }
    if (!res.ok) throw new Error("SharePoint returned an error (HTTP " + res.status + ").");
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
    var siteUrl = GRAPH + "/sites/" + cfg.sharePointHost + (sitePath ? ":" + encodeURI(sitePath) : "");
    return graphGet(siteUrl, token);
  }

  /** Reads a CSV file from the site's default document library. */
  async function loadFromFile(cfg, token, site) {
    var path = "/" + String(cfg.filePath || "").replace(/^\/+/, "");
    var encoded = path.split("/").map(function (seg) {
      return encodeURIComponent(seg).replace(/'/g, "%27");
    }).join("/");
    // Graph's /content endpoint redirects, which browsers block for authorised
    // cross-origin requests, so ask for the short-lived download URL instead.
    var item = await graphGet(GRAPH + "/sites/" + site.id + "/drive/root:" + encoded +
      "?$select=id,name,@microsoft.graph.downloadUrl", token);
    var dl = item["@microsoft.graph.downloadUrl"];
    if (!dl) throw new Error("The client file could not be downloaded from SharePoint.");
    var res = await fetch(dl);
    if (!res.ok) throw new Error("The client file could not be downloaded (HTTP " + res.status + ").");
    return clientsFromCsv(await res.text(), cfg);
  }

  /** Reads a SharePoint list. */
  async function loadFromList(cfg, token, site) {
    var listQuery = GRAPH + "/sites/" + site.id + "/lists?$select=id,displayName&$filter=displayName eq '" +
      String(cfg.listName).replace(/'/g, "''") + "'";
    var lists = await graphGet(listQuery, token);
    if (!lists.value || !lists.value.length) {
      throw new Error('The SharePoint list "' + cfg.listName + '" was not found.');
    }
    var url = GRAPH + "/sites/" + site.id + "/lists/" + lists.value[0].id +
      "/items?$top=500&$expand=fields($select=" + cfg.nameColumn + "," + cfg.aliasColumn + ")";
    var clients = [];
    while (url) {
      var page = await graphGet(url, token);
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

  root.KCData = { loadClients: loadClients, splitAliases: splitAliases, parseCsv: parseCsv, clientsFromCsv: clientsFromCsv };
})(window);
