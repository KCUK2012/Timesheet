/*
 * Outlook Timesheet: subject-line logic (no Office dependency, so it can be unit tested).
 * Format (from version 1.3):
 *   External:  Client – Activity (External, Subscription Service)   or (External, Project Work)
 *   Internal:  Code – Activity (Internal)
 *   ACTUAL is added inside the bracket: (External, Project Work, ACTUAL), (Internal, ACTUAL)
 *   Not for timesheet (Internal only): NFT – Code – Activity. The code is optional; the activity is not.
 * Older subjects ("External – Client – Activity ACTUAL", "Internal – Code – Activity NFT") are still read.
 */
(function (root) {
  "use strict";

  var DEFAULT_SEP = " – "; // space, en dash, space
  var FLAGS = ["ACTUAL", "NFT"];
  var SPLIT = /\s+[–—-]\s+/;

  function clean(text) {
    return String(text || "").replace(/\s+/g, " ").trim();
  }

  /** Build the subject from its parts. Returns "" if a required part is missing. */
  function build(parts, sep) {
    sep = sep || DEFAULT_SEP;
    var type = parts.type === "Internal" ? "Internal" : "External";
    var flags = parts.flags || {};
    var nft = type === "Internal" && !!flags.NFT;
    var target = clean(parts.target);
    var workType = clean(parts.workType);
    var activity = clean(parts.activity);
    if (!activity) return "";
    if (nft) return ["NFT", target, activity].filter(Boolean).join(sep);
    var tags;
    if (type === "External") {
      if (!target || !workType) return "";
      tags = ["External", workType];
    } else {
      if (!target) return "";
      tags = ["Internal"];
    }
    if (flags.ACTUAL) tags.push("ACTUAL");
    return target + sep + activity + " (" + tags.join(", ") + ")";
  }

  function splitSegs(text) {
    text = clean(text);
    return text ? text.split(SPLIT) : [];
  }

  /**
   * Parse an existing subject so the panel can be pre-filled.
   * Returns null if the subject is not in a timesheet format, otherwise
   * { type, segments, workType, flags, legacy }:
   *   External segments: [client, activity...]  (legacy: [client, work type?, activity...])
   *   Internal segments: [code, activity...]    (NFT: [code?, activity...])
   */
  function parse(subject) {
    var s = clean(subject);
    var flags = { ACTUAL: false, NFT: false };

    // Not for timesheet: "NFT – ...".
    var n = s.match(/^NFT(?:\s+[–—-]\s+(.*))?$/i);
    if (n) {
      flags.NFT = true;
      return { type: "Internal", segments: splitSegs(n[1]), workType: null, flags: flags, legacy: false };
    }

    // Tags in a bracket at the very end, starting with External or Internal.
    var b = s.match(/^(.*\S)\s*\(([^()]*)\)$/);
    if (b) {
      var tags = b[2].split(",").map(clean).filter(Boolean);
      var first = (tags[0] || "").toLowerCase();
      if (first === "external" || first === "internal") {
        var workType = null;
        tags.slice(1).forEach(function (t) {
          var up = t.toUpperCase();
          if (FLAGS.indexOf(up) !== -1) flags[up] = true;
          else if (!workType) workType = t;
        });
        return {
          type: first === "external" ? "External" : "Internal",
          segments: splitSegs(b[1]), workType: workType, flags: flags, legacy: false
        };
      }
    }

    // Older format: "External – ..." / "Internal – ..." with flags at the end.
    for (var k = 0; k < 2; k++) {
      FLAGS.forEach(function (flag) {
        var re = new RegExp("\\s+" + flag + "$");
        if (re.test(s)) { flags[flag] = true; s = s.replace(re, ""); }
      });
    }
    var m = s.match(/^(External|Internal)\b(.*)$/i);
    if (!m) return null;
    var rest = m[2].trim();
    var segs = [];
    if (rest) {
      if (!/^[–—-]\s+/.test(rest)) return null;
      segs = splitSegs(rest.replace(/^[–—-]\s+/, ""));
    }
    var type = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
    if (type === "External") flags.NFT = false;
    return { type: type, segments: segs, workType: null, flags: flags, legacy: true };
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
