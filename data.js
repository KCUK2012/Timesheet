/*
 * KC Timesheet: sign-in and client list.
 * Uses Microsoft's nested app authentication (NAA) through MSAL.js, so staff are
 * signed in silently with the Microsoft 365 account they already use in Outlook.
 * Reads the client list from SharePoint through Microsoft Graph with delegated
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
      throw new Error("You do not have access to the client list in SharePoint, or it could not be found. Please contact the list owner.");
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

  /** Returns [{ name, aliases[] }] from the SharePoint list. */
  async function loadClients(cfg) {
    var token = await getToken(cfg);

    var sitePath = (cfg.sitePath || "").trim();
    var siteUrl = GRAPH + "/sites/" + cfg.sharePointHost + (sitePath ? ":" + encodeURI(sitePath) : "");
    var site = await graphGet(siteUrl, token);

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

  root.KCData = { loadClients: loadClients, splitAliases: splitAliases };
})(window);
