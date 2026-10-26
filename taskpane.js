/* Outlook Timesheet: task pane behaviour. Requires office.js, config.js, xlsx-reader.js, data.js and subject.js. */
(function () {
  "use strict";

  var VERSION = "1.3.2";
  var S = window.KCSubject;
  var config = null;
  var sep = S.DEFAULT_SEP;
  var el = {};
  var started = false;

  // ---------- Office.js promise wrappers ----------
  function officeGet(prop) {
    return new Promise(function (resolve, reject) {
      prop.getAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r.error);
      });
    });
  }
  function officeSet(prop, value) {
    return new Promise(function (resolve, reject) {
      prop.setAsync(value, function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(r.error);
      });
    });
  }

  // ---------- UI helpers ----------
  function $(id) { return document.getElementById(id); }
  function currentType() {
    return document.querySelector('input[name="type"]:checked').value;
  }
  function setStatus(text, kind) {
    el.status.textContent = text || "";
    el.status.className = "status" + (kind ? " " + kind : "");
  }
  function setWarnings(list) {
    el.warnings.innerHTML = "";
    (list || []).forEach(function (w) {
      var d = document.createElement("div");
      d.textContent = w;
      el.warnings.appendChild(d);
    });
  }
  function option(value, text) {
    var o = document.createElement("option");
    o.value = value;
    o.textContent = text;
    return o;
  }
  function sameText(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }

  // ---------- Lists ----------
  // Recent clients: kept only in this browser profile on this computer (localStorage).
  var RECENT_KEY = "outlookTimesheet.recentClients";
  var RECENT_MAX = 5;
  function getRecent() {
    try {
      var v = JSON.parse(window.localStorage.getItem(RECENT_KEY) || "[]");
      return Array.isArray(v) ? v.filter(function (x) { return typeof x === "string"; }) : [];
    } catch (e) { return []; }
  }
  function rememberClient(name) {
    if (!name) return;
    try {
      var list = getRecent().filter(function (n) { return n !== name; });
      list.unshift(name);
      window.localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, RECENT_MAX)));
    } catch (e) { /* storage unavailable: carry on without it */ }
  }

  function clientLabel(c) {
    return c.aliases && c.aliases.length ? c.name + "  (" + c.aliases.join(", ") + ")" : c.name;
  }

  function populateClients(filter) {
    var selected = el.clientSelect.value;
    el.clientSelect.innerHTML = "";
    var matching = config.clients
      .slice()
      .sort(function (a, b) { return a.name.localeCompare(b.name, "en-GB"); })
      .filter(function (c) { return S.matchesClient(c, filter); });
    var byName = {};
    matching.forEach(function (c) { byName[c.name] = c; });
    var recent = getRecent().filter(function (n) { return byName[n]; });
    var target = el.clientSelect;
    if (recent.length) {
      var g1 = document.createElement("optgroup");
      g1.label = "Recent";
      recent.forEach(function (n) { g1.appendChild(option(n, clientLabel(byName[n]))); });
      el.clientSelect.appendChild(g1);
      target = document.createElement("optgroup");
      target.label = "All clients";
      el.clientSelect.appendChild(target);
    }
    matching.forEach(function (c) { target.appendChild(option(c.name, clientLabel(c))); });
    if (selected) el.clientSelect.value = selected;
    if (!el.clientSelect.value && matching.length === 1) el.clientSelect.value = matching[0].name;
  }

  function otherCfg() { return window.KC_CONFIG.otherCode || { label: "Option not available" }; }
  function isOtherName(name) { return sameText(name, otherCfg().label); }

  function populateCodes() {
    var selected = el.codeSelect.value;
    el.codeSelect.innerHTML = "";
    el.codeSelect.appendChild(option("", "Choose a code"));
    config.internalCodes.forEach(function (c) {
      if (!isOtherName(c.subject)) el.codeSelect.appendChild(option(c.subject, c.label));
    });
    el.codeSelect.appendChild(option(otherCfg().label, otherCfg().label));
    el.codeSelect.value = selected;
    if (el.codeSelect.value !== selected) el.codeSelect.value = "";
  }

  function workTypeList() { return window.KC_CONFIG.workTypes || ["Subscription Service", "Project Work"]; }

  function populateWorkTypes() {
    var current = currentWorkType();
    el.workTypes.innerHTML = "";
    workTypeList().forEach(function (w) {
      var lab = document.createElement("label");
      var inp = document.createElement("input");
      inp.type = "radio"; inp.name = "workType"; inp.value = w;
      if (w === current) inp.checked = true;
      inp.addEventListener("change", refresh);
      lab.appendChild(inp);
      lab.appendChild(document.createTextNode(" " + w));
      el.workTypes.appendChild(lab);
    });
  }

  function currentWorkType() {
    var r = document.querySelector('input[name="workType"]:checked');
    return r ? r.value : "";
  }
  function setWorkType(value) {
    document.querySelectorAll('input[name="workType"]').forEach(function (r) { r.checked = sameText(r.value, value); });
  }
  function findWorkType(name) {
    var aliases = window.KC_CONFIG.workTypeAliases || {};
    for (var k in aliases) if (sameText(k, name)) name = aliases[k];
    return workTypeList().filter(function (w) { return sameText(w, name); })[0] || null;
  }

  function findCode(name) {
    var hit = config.internalCodes.filter(function (c) { return sameText(c.subject, name); })[0];
    return hit ? hit.subject : null;
  }

  function currentCode() {
    var v = el.codeSelect.value;
    if (isOtherName(v)) return { subject: v, hint: (config.internalCodes.filter(function (c) { return isOtherName(c.subject); })[0] || {}).hint || "" };
    return config.internalCodes.filter(function (c) { return c.subject === v; })[0] || null;
  }

  function currentParts() {
    var type = currentType();
    return {
      type: type,
      target: type === "Internal" ? el.codeSelect.value : el.clientSelect.value,
      workType: type === "External" ? currentWorkType() : "",
      activityOptional: activityIsOptional(),
      activity: el.activity.value,
      flags: { ACTUAL: el.flagActual.checked, NFT: type === "Internal" && el.flagNft.checked }
    };
  }

  // Codes marked "Not for Timesheet" in the workbook do not need an activity.
  function activityIsOptional() {
    if (currentType() !== "Internal") return false;
    var code = currentCode();
    return !!(code && code.nft);
  }

  function missingText() {
    var missing = [];
    if (currentType() === "Internal") {
      var nft = el.flagNft.checked;
      if (!el.codeSelect.value && !nft) missing.push("a code");
      if (!el.activity.value.trim()) missing.push("an activity");
      return "Still needed: " + missing.join(" and ") + ".";
    }
    if (!el.clientSelect.value) missing.push("a client");
    if (!currentWorkType()) missing.push("a work type (" + workTypeList().join(" or ") + ")");
    if (!el.activity.value.trim()) missing.push("an activity");
    return "Still needed: " + missing.join(", ") + ".";
  }

  // Codes marked "Not for timesheet" in the workbook tick NFT automatically.
  var autoNft = false;
  function applyAutoNft() {
    var code = currentCode();
    var forced = currentType() === "Internal" && !!(code && code.nft);
    if (forced) {
      el.flagNft.checked = true;
      autoNft = true;
    } else if (autoNft) {
      el.flagNft.checked = false;
      autoNft = false;
    }
    el.flagNft.disabled = forced;
  }

  function refresh() {
    if (!config) return;
    var internal = currentType() === "Internal";
    applyAutoNft();
    if (!internal) el.flagNft.checked = false;
    var nft = internal && el.flagNft.checked;
    el.actualLabel.hidden = nft;
    el.activityOptional.hidden = !activityIsOptional();
    el.clientBlock.hidden = internal;
    el.codeBlock.hidden = !internal;
    el.nftLabel.hidden = !internal;
    el.activityHint.hidden = internal;
    el.activity.placeholder = internal
      ? (isOtherName(el.codeSelect.value) ? "Describe the internal work" : "e.g. expenses, team meeting")
      : "e.g. annual DPO review, drafting";
    var code = currentCode();
    el.codeHint.textContent = internal && code && code.hint ? code.hint : "";

    var subject = S.build(currentParts(), sep);
    el.preview.textContent = subject || missingText();
    el.apply.disabled = !subject;
  }

  // ---------- Reading and writing the appointment ----------
  async function prefillFromSubject() {
    var subject = await officeGet(Office.context.mailbox.item.subject);
    var parsed = S.parse(subject);
    if (!parsed) return;
    var segs = parsed.segments;
    var warnings = [];

    document.querySelector('input[name="type"][value="' + parsed.type + '"]').checked = true;
    el.flagActual.checked = parsed.flags.ACTUAL;
    el.flagNft.checked = parsed.type === "Internal" && parsed.flags.NFT;

    if (parsed.type === "External") {
      var clientText = segs[0] || "";
      var name = S.resolveClient(config.clients, clientText);
      if (name) el.clientSelect.value = name;
      else if (clientText) warnings.push('"' + clientText + '" is not on the client list. Choose the client from the list, or ask for it to be added to the Clients tab.');
      var wt = null;
      if (!parsed.legacy) {
        wt = parsed.workType ? findWorkType(parsed.workType) : null;
        el.activity.value = segs.slice(1).join(sep);
      } else {
        // Older format: External – Client – [work type –] Activity
        wt = segs.length >= 3 ? findWorkType(segs[1]) : null;
        el.activity.value = segs.slice(wt ? 2 : 1).join(sep);
      }
      if (wt) setWorkType(wt);
      else warnings.push("Please choose " + workTypeList().join(" or ") + " for this entry.");
    } else {
      var codeName = !segs.length ? null : (isOtherName(segs[0]) ? otherCfg().label : findCode(segs[0]));
      if (codeName) {
        el.codeSelect.value = codeName;
        el.activity.value = segs.slice(1).join(sep);
      } else if (parsed.flags.NFT) {
        el.activity.value = segs.join(sep);
      } else {
        el.activity.value = segs.slice(1).join(sep);
        if (segs[0]) warnings.push('"' + segs[0] + '" is not one of the current internal codes. Please pick one from the list.');
      }
    }
    setWarnings(warnings);
  }

  async function checkTiming() {
    var item = Office.context.mailbox.item;
    try {
      var start = await officeGet(item.start);
      var end = await officeGet(item.end);
      var allDay = false;
      // isAllDayEvent is only available from requirement set Mailbox 1.14.
      if (Office.context.requirements.isSetSupported("Mailbox", "1.14") && item.isAllDayEvent) {
        allDay = await officeGet(item.isAllDayEvent);
      }
      return S.timingWarnings(start, end, allDay);
    } catch (e) {
      return [];
    }
  }

  async function apply() {
    var subject = S.build(currentParts(), sep);
    if (!subject) return;
    el.apply.disabled = true;
    setStatus("Updating subject...");
    try {
      await officeSet(Office.context.mailbox.item.subject, subject);
      if (currentType() === "External") { rememberClient(el.clientSelect.value); populateClients(el.clientFilter.value); }
      setWarnings(await checkTiming());
      setStatus("Subject updated. Save or send the entry to keep the change.", "ok");
    } catch (e) {
      setStatus("Could not update the subject: " + (e && e.message ? e.message : "unknown error"), "error");
    } finally {
      el.apply.disabled = false;
    }
  }

  // ---------- Loading the lists ----------
  function withTimeout(promise, ms) {
    var timer;
    return Promise.race([
      promise,
      new Promise(function (_, reject) {
        timer = setTimeout(function () {
          reject(new Error("Timed out after " + Math.round(ms / 1000) + " seconds at " + window.KCData.getStage() +
            ". If this is sign-in, check the Entra redirect addresses and admin consent."));
        }, ms);
      })
    ]).finally(function () { clearTimeout(timer); });
  }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  async function loadConfig() {
    var cfg = window.KC_CONFIG;
    if (!cfg || /FILL IN/.test(String(cfg.tenantId) + cfg.clientId + cfg.sitePath + cfg.filePath + cfg.listName)) {
      throw new Error("The add-in has not been set up yet: config.js still needs the tenant, app and SharePoint details.");
    }
    var lists = null, lastError = null;
    var attempts = cfg.loadAttempts || 2;
    for (var i = 1; i <= attempts && !lists; i++) {
      setStatus(i === 1 ? "Loading lists from SharePoint..." : "Trying again to load the lists (attempt " + i + " of " + attempts + ")...");
      try { lists = await withTimeout(window.KCData.loadLists(cfg), 45000); }
      catch (e) { lastError = e; if (i < attempts) await wait(2000); }
    }
    if (!lists) throw lastError || new Error("The lists could not be loaded.");
    if (!lists.clients.length) throw new Error("The client list is empty.");
    setStatus("");
    return {
      separator: cfg.separator,
      clients: lists.clients,
      internalCodes: lists.internalCodes || cfg.internalCodes
    };
  }

  async function reloadLists() {
    el.reload.disabled = true;
    el.apply.disabled = true;
    try {
      config = await loadConfig();
      if (config.separator) sep = config.separator;
      populateClients(el.clientFilter.value);
      populateCodes();
      populateWorkTypes();
      if (!started) {
        started = true;
        try { await prefillFromSubject(); } catch (e) { /* new or untitled item: nothing to prefill */ }
        if (!el.warnings.children.length) setWarnings(await checkTiming());
      } else {
        setStatus("Lists reloaded.", "ok");
      }
      refresh();
    } catch (e) {
      setStatus((e && e.message ? e.message : "The lists could not be loaded.") +
        " Click Reload lists to try again.", "error");
      el.preview.textContent = "The lists have not loaded.";
    } finally {
      el.reload.disabled = false;
    }
  }

  // ---------- Start-up ----------
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;

    var ver = document.getElementById("version");
    if (ver) {
      var naa = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1") ? "supported" : "not supported";
      ver.textContent = "Version " + VERSION + " | Outlook silent sign-in: " + naa;
    }

    ["clientBlock", "codeBlock", "clientFilter", "clientSelect", "workTypes", "codeSelect",
     "codeHint", "activityHint", "actualLabel", "activityOptional", "activity", "flagActual", "flagNft", "nftLabel", "preview",
     "apply", "warnings", "status", "reload"]
      .forEach(function (id) { el[id] = $(id); });

    document.querySelectorAll('input[name="type"]').forEach(function (r) {
      r.addEventListener("change", refresh);
    });
    el.clientFilter.addEventListener("input", function () { if (config) populateClients(el.clientFilter.value); refresh(); });
    el.clientSelect.addEventListener("change", refresh);
    el.codeSelect.addEventListener("change", refresh);
    el.activity.addEventListener("input", refresh);
    el.activity.addEventListener("keydown", function (e) { if (e.key === "Enter" && !el.apply.disabled) apply(); });
    el.flagActual.addEventListener("change", refresh);
    el.flagNft.addEventListener("change", refresh);
    el.apply.addEventListener("click", apply);
    el.reload.addEventListener("click", reloadLists);

    reloadLists();
  });
})();
