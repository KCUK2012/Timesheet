/* KC Timesheet: task pane behaviour. Requires office.js and subject.js. */
(function () {
  "use strict";

  var VERSION = "1.1.2";
  var S = window.KCSubject;
  var config = null;
  var sep = S.DEFAULT_SEP;
  var el = {};

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

  function populateClients(filter) {
    var selected = el.clientSelect.value;
    el.clientSelect.innerHTML = "";
    config.clients
      .slice()
      .sort(function (a, b) { return a.name.localeCompare(b.name, "en-GB"); })
      .filter(function (c) { return S.matchesClient(c, filter); })
      .forEach(function (c) {
        var o = document.createElement("option");
        o.value = c.name;
        o.textContent = c.aliases && c.aliases.length ? c.name + "  (" + c.aliases.join(", ") + ")" : c.name;
        if (c.name === selected) o.selected = true;
        el.clientSelect.appendChild(o);
      });
    if (!el.clientSelect.value && el.clientSelect.options.length === 1) {
      el.clientSelect.options[0].selected = true;
    }
  }

  function populateCodes() {
    el.codeSelect.innerHTML = "";
    var blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "Choose a code";
    el.codeSelect.appendChild(blank);
    config.internalCodes.forEach(function (c) {
      var o = document.createElement("option");
      o.value = c.subject;
      o.textContent = c.label;
      el.codeSelect.appendChild(o);
    });
  }

  function currentCode() {
    var v = el.codeSelect.value;
    return config.internalCodes.filter(function (c) { return c.subject === v; })[0] || null;
  }

  function renderQuick() {
    el.quick.innerHTML = "";
    var key = currentType() === "Internal" ? el.codeSelect.value : el.clientSelect.value;
    var list = (config.quickActivities && config.quickActivities[key]) || [];
    list.forEach(function (text) {
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = text;
      b.addEventListener("click", function () { el.activity.value = text; refresh(); });
      el.quick.appendChild(b);
    });
  }

  function currentParts() {
    var type = currentType();
    return {
      type: type,
      target: type === "Internal" ? el.codeSelect.value : el.clientSelect.value,
      activity: el.activity.value,
      flags: { ACTUAL: el.flagActual.checked, NFT: el.flagNft.checked }
    };
  }

  function refresh() {
    var internal = currentType() === "Internal";
    el.clientBlock.hidden = internal;
    el.codeBlock.hidden = !internal;
    var code = currentCode();
    el.codeHint.textContent = internal && code && code.hint ? code.hint : "";
    renderQuick();

    var subject = S.build(currentParts(), sep);
    el.preview.textContent = subject || "Choose a client or code and add an activity.";
    el.apply.disabled = !subject;
  }

  // ---------- Reading and writing the appointment ----------
  async function prefillFromSubject() {
    var subject = await officeGet(Office.context.mailbox.item.subject);
    var parsed = S.parse(subject);
    if (!parsed) return;

    document.querySelector('input[name="type"][value="' + parsed.type + '"]').checked = true;
    if (parsed.type === "External") {
      var name = S.resolveClient(config.clients, parsed.target);
      if (name) el.clientSelect.value = name;
      else setWarnings(['"' + parsed.target + '" is not on the client list. Choose the client from the list, or ask for it to be added to the crib sheet.']);
    } else {
      el.codeSelect.value = parsed.target;
      if (!el.codeSelect.value) {
        setWarnings(['"' + parsed.target + '" is not one of the agreed internal codes. Please pick one from the list.']);
      }
    }
    el.activity.value = parsed.activity;
    el.flagActual.checked = parsed.flags.ACTUAL;
    el.flagNft.checked = parsed.flags.NFT;
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
      setWarnings(await checkTiming());
      setStatus("Subject updated. Save or send the entry to keep the change.", "ok");
    } catch (e) {
      setStatus("Could not update the subject: " + (e && e.message ? e.message : "unknown error"), "error");
    } finally {
      el.apply.disabled = false;
    }
  }

  async function loadConfig() {
    var cfg = window.KC_CONFIG;
    if (!cfg || /FILL IN/.test(String(cfg.tenantId) + cfg.clientId + cfg.sitePath + cfg.filePath + cfg.listName)) {
      throw new Error("The add-in has not been set up yet: config.js still needs the tenant, app and SharePoint details.");
    }
    setStatus("Loading client list from SharePoint...");
    var clients = await Promise.race([
      window.KCData.loadClients(cfg),
      new Promise(function (_, reject) {
        setTimeout(function () {
          reject(new Error("Timed out after 45 seconds at " + window.KCData.getStage() +
            ". If this is sign-in, check the Entra redirect addresses and admin consent."));
        }, 45000);
      })
    ]);
    setStatus("");
    return {
      separator: cfg.separator,
      clients: clients,
      internalCodes: cfg.internalCodes,
      quickActivities: cfg.quickActivities
    };
  }

  // ---------- Start-up ----------
  Office.onReady(async function (info) {
    if (info.host !== Office.HostType.Outlook) return;

    var ver = document.getElementById("version");
    if (ver) {
      var naa = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1") ? "supported" : "not supported";
      ver.textContent = "Version " + VERSION + " | Outlook silent sign-in: " + naa;
    }

    ["clientBlock", "codeBlock", "clientFilter", "clientSelect", "codeSelect", "codeHint",
     "activity", "quick", "flagActual", "flagNft", "preview", "apply", "warnings", "status"]
      .forEach(function (id) { el[id] = $(id); });

    try {
      config = await loadConfig();
      if (config.separator) sep = config.separator;
    } catch (e) {
      setStatus(e && e.message ? e.message : "The client list could not be loaded.", "error");
      return;
    }

    populateClients("");
    populateCodes();

    document.querySelectorAll('input[name="type"]').forEach(function (r) {
      r.addEventListener("change", refresh);
    });
    el.clientFilter.addEventListener("input", function () { populateClients(el.clientFilter.value); refresh(); });
    el.clientSelect.addEventListener("change", refresh);
    el.codeSelect.addEventListener("change", refresh);
    el.activity.addEventListener("input", refresh);
    el.activity.addEventListener("keydown", function (e) { if (e.key === "Enter" && !el.apply.disabled) apply(); });
    el.flagActual.addEventListener("change", refresh);
    el.flagNft.addEventListener("change", refresh);
    el.apply.addEventListener("click", apply);

    try { await prefillFromSubject(); } catch (e) { /* new or untitled item: nothing to prefill */ }
    refresh();

    var existing = el.warnings.children.length ? [] : await checkTiming();
    if (existing.length) setWarnings(existing);
  });
})();
