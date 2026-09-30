/*
 * KC Timesheet: subject-line logic (no Office dependency, so it can be unit tested).
 * Format:  External – Client – Activity
 *          Internal – Code – Activity
 * Optional flags ACTUAL and NFT are appended at the end of the subject.
 */
(function (root) {
  "use strict";

  var DEFAULT_SEP = " – "; // space, en dash, space
  var FLAGS = ["ACTUAL", "NFT"];

  function clean(text) {
    return String(text || "").replace(/\s+/g, " ").trim();
  }

  /** Build the subject from its parts. Returns "" if a required part is missing. */
  function build(parts, sep) {
    sep = sep || DEFAULT_SEP;
    var type = parts.type === "Internal" ? "Internal" : "External";
    var target = clean(parts.target);
    var activity = clean(parts.activity);
    if (!target || !activity) return "";
    var subject = [type, target, activity].join(sep);
    FLAGS.forEach(function (flag) {
      if (parts.flags && parts.flags[flag]) subject += " " + flag;
    });
    return subject;
  }

  /**
   * Parse an existing subject so the panel can be pre-filled.
   * Accepts an en dash, em dash or hyphen as the separator.
   * Returns null if the subject is not in the timesheet format.
   */
  function parse(subject) {
    var s = clean(subject);
    var flags = { ACTUAL: false, NFT: false };
    FLAGS.forEach(function (flag) {
      var re = new RegExp("\\s+" + flag + "$");
      // Flags may appear in either order at the end.
      if (re.test(s)) { flags[flag] = true; s = s.replace(re, ""); }
    });
    FLAGS.forEach(function (flag) {
      var re = new RegExp("\\s+" + flag + "$");
      if (re.test(s)) { flags[flag] = true; s = s.replace(re, ""); }
    });
    var m = s.match(/^(External|Internal)\s+[–—-]\s+(.+?)\s+[–—-]\s+(.+)$/i);
    if (!m) return null;
    return {
      type: m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase(),
      target: m[2],
      activity: m[3],
      flags: flags
    };
  }

  /** Warnings based on the timesheet rules for timing. Dates are JS Date objects. */
  function timingWarnings(start, end, isAllDay) {
    var warnings = [];
    if (isAllDay) {
      warnings.push("This is an all-day entry, so it will be ignored as work. Untick All day and set start and end times.");
    }
    if (start && end) {
      var hours = (end - start) / 36e5;
      if (hours > 10) {
        warnings.push("This entry is longer than ten hours, so it will be read as a diary marker, not work. Split or shorten it.");
      }
      if (hours <= 0) {
        warnings.push("The end time is not after the start time.");
      }
    }
    return warnings;
  }

  /** Case-insensitive match against a client's name and aliases. */
  function matchesClient(client, query) {
    var q = clean(query).toLowerCase();
    if (!q) return true;
    var terms = [client.name].concat(client.aliases || []);
    return terms.some(function (t) { return String(t).toLowerCase().indexOf(q) !== -1; });
  }

  /** Resolve a typed name or alias to the recognised client name, or null. */
  function resolveClient(clients, text) {
    var q = clean(text).toLowerCase();
    for (var i = 0; i < clients.length; i++) {
      var c = clients[i];
      var terms = [c.name].concat(c.aliases || []);
      for (var j = 0; j < terms.length; j++) {
        if (String(terms[j]).toLowerCase() === q) return c.name;
      }
    }
    return null;
  }

  var api = {
    DEFAULT_SEP: DEFAULT_SEP,
    build: build,
    parse: parse,
    timingWarnings: timingWarnings,
    matchesClient: matchesClient,
    resolveClient: resolveClient
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.KCSubject = api;
})(typeof window !== "undefined" ? window : this);
