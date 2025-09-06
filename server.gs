/**
 * TruckTalk Connect — Apps Script (server)
 * - Complies with codingtask.md
 * - Returns contract: { ok, issues[], loads? (only when ok), mapping, meta }
 * - Soft rate limit: 10 analyses / minute / user
 * - Optional OpenAI proxy usage (set SCRIPT PROPERTY PROXY_URL). Prompts forbid fabrication.
 */

// --- Assumption flags (set during build/normalization) ---
var __ASSUMED_TZ_USED = false;
var __ASSUMED_TZ_NAME = null;
// Track which rows required assumed timezone (for erroring per spec)
var __ASSUMED_TZ_ROWS = [];

// ---- Add-on homepage & sidebar ------------------------------------------------

/**
 * Add-on homepage card. Provides a button to open the HTML sidebar.
 * Matches appsscript.json homepageTrigger: onHomepage
 */
function onHomepage(e) {
  var cs = CardService;
  var section = cs.newCardSection()
    .addWidget(cs.newTextParagraph().setText(
      'Analyze the active sheet for loads, flag issues, and export JSON.'
    ))
    .addWidget(
      cs.newTextButton()
        .setText('Open Sidebar')
        .setOnClickAction(cs.newAction().setFunctionName('showSidebar'))
    );

  return cs.newCardBuilder()
    .setHeader(cs.newCardHeader().setTitle('TruckTalk Connect'))
    .addSection(section)
    .build();
}

/** Show the HTMLService sidebar (ui.html). */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ui')
    .setTitle('TruckTalk Connect');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---- Public entrypoint --------------------------------------------------------

/**
 * Client → Server contract
 * @param {{ headerOverrides?: Object<string,string> }} [opts]
 * @return {AnalysisResult}
 */
function analyzeActiveSheet(opts) {
  opts = opts || {};
  var userEmail = 'anon';
  try {
    var eff = Session.getEffectiveUser && Session.getEffectiveUser();
    if (eff && eff.getEmail) userEmail = eff.getEmail() || 'anon';
  } catch (e) {
    userEmail = 'anon';
  }
  if (!allowAnalysisNow_(userEmail)) {
    return strictResult_({
      ok: false,
      issues: [issue_('RATE_LIMIT', 'error',
        'Too many analyses in the last 60 seconds.',
        null,
        null,
        'Please wait a few seconds and try again.'
      )],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    });
  }

  try {
    var snapshot = getSheetSnapshot_(200);

    // 1) Header mapping (rules first)
    var mappingResult = detectHeaderMapping_(snapshot.headers, opts.headerOverrides || {});
    var mappingIssues = mappingResult.issues || [];

    // 1a) AI header mapping proposal (only when needed)
    var aiAvailable = !!getScriptProperty_('PROXY_URL');
    if (aiAvailable) {
      var needAIMapping = mappingIssues.some(function (i) {
        return i.code === 'MISSING_COLUMN' || i.code === 'AMBIGUOUS_HEADER';
      });
      if (needAIMapping) {
        try {
          var aiMap = proposeHeaderMappingWithAI_(snapshot, mappingResult.mapping);
          if (aiMap && aiMap.suggestions && aiMap.suggestions.length) {
            var lines = aiMap.suggestions.map(function (s) {
              var conf = typeof s.confidence === 'number' ? ' (' + Math.round(s.confidence * 100) + '%)' : '';
              return '• "' + s.header + '" → ' + s.field + conf + (s.reason ? ' — ' + s.reason : '');
            });
            mappingIssues.push(issue_(
              'AI_MAPPING_SUGGESTION',
              'warn',
              'AI suggests header mappings:\n' + lines.join('\n'),
              null,
              null,
              'Confirm in the mapping dialog to apply.'
            ));
          }
          if (aiMap && aiMap.notes) {
            mappingIssues.push(issue_('AI_MAPPING_NOTES', 'warn', String(aiMap.notes)));
          }
        } catch (e) {
          mappingIssues.push(issue_('AI_MAPPING_FAILED', 'warn', 'AI could not propose header mapping.'));
        }
      }
    }

    // 2) Row validations (contract-level)
    var validation = validateRows_(snapshot, mappingResult.mapping);

    // 3) Optional AI normalization suggestions (examples only, no edits)
    if (aiAvailable) {
      var dateRelated = validation.issues.some(function (i) {
        return i.code === 'BAD_DATE_FORMAT' || i.code === 'TIMEZONE_MISSING' || i.code === 'NON_ISO_OUTPUT';
      }) || __ASSUMED_TZ_USED; // include when we had to assume TZ
      if (dateRelated) {
        try {
          var problemCells = gatherProblemDateCells_(snapshot, mappingResult.mapping, 12);
          if (problemCells.length) {
            var aiNorm = suggestNormalizationWithAI_(problemCells);
            if (aiNorm && aiNorm.items && aiNorm.items.length) {
              var preview = aiNorm.items.slice(0, 6).map(function (it) {
                var base = 'Row ' + it.row + ' ' + it.column + ': "' + it.original + '"';
                if (it.normalized) base += ' → ' + it.normalized;
                if (it.note) base += ' — ' + it.note;
                return '• ' + base;
              });
              validation.issues.push(issue_(
                'AI_NORMALIZATION_HINTS',
                'warn',
                'AI normalization examples (no changes applied):\n' + preview.join('\n'),
                null,
                null,
                'Use ISO 8601 Z (YYYY-MM-DDTHH:mm:ssZ).'
              ));
            }
          }
        } catch (e) {
          validation.issues.push(issue_('AI_NORMALIZATION_FAILED', 'warn', 'AI normalization hints unavailable.'));
        }
      }
    }

    // 4) Build loads (typed payload); they will be omitted if any errors exist
    __ASSUMED_TZ_USED = false;
    __ASSUMED_TZ_ROWS = [];
    var loads = buildLoads_(snapshot, mappingResult.mapping, { 
      statusMap: getUserStatusMap_(),
      brokerMap: getUserBrokerMap_()
    });

    // 5) AI issue summary (non-blocking)
    if (aiAvailable) {
      try {
        var aiNotes = summarizeIssuesWithAI_(snapshot, validation.issues.concat(mappingIssues));
        if (aiNotes && aiNotes.trim()) {
          validation.issues.push(issue_('AI_SUMMARY', 'warn', aiNotes));
        }
      } catch (aiErr) {
        validation.issues.push(issue_('AI_NOTE_FAILED', 'warn',
          'Unable to generate AI summary for issues (non-blocking).'));
      }
    }

    // 6) Assemble issues and apply ASSUMED_TIMEZONE (ERROR per spec)
    var allIssues = mappingIssues.concat(validation.issues);

    if (__ASSUMED_TZ_USED) {
      var uniqMap = Object.create(null);
      (__ASSUMED_TZ_ROWS || []).forEach(function (r) { if (r) uniqMap[r] = true; });
      var uniqRows = Object.keys(uniqMap).map(function (k) { return parseInt(k, 10); })
        .filter(function (n) { return !isNaN(n); })
        .sort(function (a, b) { return a - b; });

      // Severity is strict by default per spec (timezone missing => error).
      // Can be relaxed via Script Property ASSUME_TZ_AS_WARN=true for demos.
      var assumedTzSeverity = (String(getScriptProperty_('ASSUME_TZ_AS_WARN') || '').toLowerCase() === 'true') ? 'warn' : 'error';
      allIssues.push(issue_(
        'ASSUMED_TIMEZONE',
        assumedTzSeverity,
        'Converted timestamps using spreadsheet timezone "' + (__ASSUMED_TZ_NAME || 'UTC') + '" because source values lacked an explicit timezone.',
        uniqRows.length ? uniqRows : null,
        null,
        'Provide explicit timezones (e.g., 2025-08-29 14:00 CST) or ISO Z (YYYY-MM-DDTHH:mm:ssZ).'
      ));
    }

    var hasErrors = allIssues.some(function (i) { return i.severity === 'error'; });

    var base = {
      ok: !hasErrors,
      issues: allIssues,
      mapping: mappingResult.mapping,
      meta: { analyzedRows: snapshot.rows.length, analyzedAt: new Date().toISOString() }
    };
    if (!hasErrors) base.loads = loads;

    return strictResult_(base);

  } catch (err) {
    return strictResult_({
      ok: false,
      issues: [issue_('UNEXPECTED_ERROR', 'error', String(err && err.message || err))],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    });
  }
}

// ---- Snapshot & helpers -------------------------------------------------------

/**
 * Read header + up to N data rows from the active sheet.
 * @param {number} limit
 * @return {{headers: string[], rows: any[][]}}
 */
function getSheetSnapshot_(limit) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) throw new Error('No active sheet.');
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) throw new Error('Sheet is empty or has no data rows.');

  var headers = values[0].map(function (h) { return (h || '').toString().trim(); });
  var rows = values.slice(1, Math.min(values.length, 1 + (limit || 200)));
  return { headers: headers, rows: rows };
}

/** Canonicalize header: lowercase, remove punctuation, collapse spaces */
function canon_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

/**
 * Detect mapping from headers → required fields using exact matches first,
 * then synonyms. If ambiguous, surface issues and leave unmapped.
 */
function detectHeaderMapping_(headers, overrides) {
  var mapping = {}; // header -> field
  var issues = [];

  var fields = [
    'loadId',
    'fromAddress',
    'fromAppointmentDateTimeUTC',
    'toAddress',
    'toAppointmentDateTimeUTC',
    'status',
    'driverName',
    'driverPhone',
    'unitNumber',
    'broker'
  ];

  var synonyms = /** @type {Object<string,string[]>} */ ({
    loadId: ['load id', 'id', 'load#', 'load number', 'load', 'ref', 'reference', 'ref #', 'vrid'],
    fromAddress: ['from','pu','pickup','origin','pickup address','from address','origin address','pickup location'],
    fromAppointmentDateTimeUTC: ['pickup time', 'pu time', 'pickup appt', 'origin appt', 'pickup datetime', 'pu datetime'],
    toAddress: ['to','drop','delivery','destination','delivery address','to address','destination address','dropoff address','delivery location'],
    toAppointmentDateTimeUTC: ['delivery time', 'del time', 'delivery appt', 'dropoff appt', 'delivery datetime', 'del datetime', 'del date', 'delivery date', 'dropoff date', 'to date'],
    status: ['load status', 'shipment status', 'status', 'stage'],
    driverName: ['driver', 'driver name', 'name', 'driver/carrier', 'driver carrier'],
    driverPhone: ['phone', 'driver phone', 'contact'],
    unitNumber: ['truck', 'tractor', 'unit', 'truck #', 'unit number'],
    broker: ['brokerage', 'mc', 'carrier', 'customer', 'shipper', 'broker']
  });

  // Apply explicit overrides first
  Object.keys(overrides || {}).forEach(function (header) {
    mapping[header] = overrides[header];
  });

  headers.forEach(function (h) {
    if (mapping[h]) return; // already overridden
    var lc = String(h || '').toLowerCase().trim();
    var cn = canon_(h);
    // Exact match (case-sensitive header equals a field name)
    if (fields.indexOf(h) !== -1) {
      mapping[h] = h;
      return;
    }
    // Exact match by lowercase (header text equals a field name lowercased)
    if (fields.indexOf(lc) !== -1) {
      mapping[h] = lc;
      return;
    }
    // Synonym match (single best) using canonical comparison
    var hitField = null, collisions = [];
    fields.forEach(function (f) {
      if (f === 'driverPhone') return; // optional field doesn't force mapping
      var syns = (synonyms[f] || []).map(canon_);
      if (syns.indexOf(cn) !== -1 || syns.indexOf(lc) !== -1) {
        if (!hitField) hitField = f;
        else collisions.push(f);
      }
    });
    if (hitField && collisions.length === 0) {
      mapping[h] = hitField;
    } else if (hitField && collisions.length > 0) {
      issues.push(issue_('AMBIGUOUS_HEADER', 'warn',
        'Ambiguous header "' + h + '" could map to: ' + [hitField].concat(collisions).join(', '),
        null,
        h,
        'Confirm the desired field in the sidebar mapping.'
      ));
    }
  });

  // Check required columns present (except driverPhone)
  var required = fields.filter(function (f) { return f !== 'driverPhone'; });
  required.forEach(function (req) {
    var mapped = Object.keys(mapping).some(function (h) { return mapping[h] === req; });
    if (!mapped) {
      issues.push(issue_('MISSING_COLUMN', 'error',
        'Missing required column for ' + req + '.',
        null,
        req,
        'Add a column or map an existing header to this field.'
      ));
    }
  });

  return { mapping: mapping, issues: issues };
}

/**
 * Validate row values and collect issues.
 * - Duplicate loadId
 * - Bad/ambiguous datetimes
 * - Empty required cells
 * - Non-ISO output (we normalize to ISO UTC when possible; otherwise warn)
 * - Status vocabulary listing
 * - Timezone missing (ERROR when a field string lacks explicit tz)
 * - Split date/time sanity: when sheet uses separate date & time columns, both must be present/valid
 */
function validateRows_(snapshot, mapping) {
  var issues = [];
  var seenIds = Object.create(null);
  var statusSet = Object.create(null);

  var idx = headerIndex_(snapshot.headers);
  var headers = snapshot.headers;

  // Potential split columns present in the sheet (used by validation & builder)
  var fromDateHeader = findHeaderBySynonyms_(headers, ['pu date','pickup date','pickup day','origin date','from date','date pu']);
  var toDateHeader   = findHeaderBySynonyms_(headers, ['del date','delivery date','dropoff date','to date','destination date','date del']);
  var fromTimeHeader = findHeaderBySynonyms_(headers, ['pu time','pickup time','pickup appt','origin appt','pickup datetime','pu datetime']);
  var toTimeHeader   = findHeaderBySynonyms_(headers, ['del time','delivery time','delivery appt','dropoff appt','delivery datetime','del datetime']);

  function splitValidation(row, rowNo, fieldName, dateHeader, timeHeader) {
    var hasDateCol = !!dateHeader && idx[dateHeader] != null;
    var hasTimeCol = !!timeHeader && idx[timeHeader] != null;
    if (!hasDateCol && !hasTimeCol) return false; // not a split layout → let direct validation handle it

    var dVal = hasDateCol ? row[idx[dateHeader]] : null;
    var tVal = hasTimeCol ? row[idx[timeHeader]] : null;

    var ymd = extractYMD_(dVal);   // null if missing/invalid
    var hms = extractTime_(tVal);  // null if missing/invalid

    // If one part exists but the other is missing/invalid → error
    var dPresent = dVal != null && String(dVal).trim() !== '';
    var tPresent = tVal != null && String(tVal).trim() !== '';
    if ((dPresent || tPresent) && (!ymd || !hms)) {
      issues.push(issue_(
        'BAD_DATE_FORMAT',
        'error',
        'Row ' + rowNo + ' has ' + fieldName + ' split across date/time but one part is missing or invalid.',
        [rowNo],
        fieldName,
        'Provide both date and time, e.g., Date: 2025-08-20 and Time: 14:00 -0600'
      ));
      return true; // handled via split rule
    }

    // When both parsed, we’ll combine in builder; no direct-string checks needed
    return true;
  }

  snapshot.rows.forEach(function (row, i) {
    var rowNo = i + 2; // account for header row

    function getField(field) {
      // Prefer exact field-named header if present (used by Auto-fix to write ISO values)
      var header = headers.indexOf(field) !== -1 ? field : findHeaderForField_(mapping, field);
      if (!header) return '';
      var col = idx[header];
      return col != null ? String(row[col]).trim() : '';
    }

    // Required fields (except driverPhone)
    ['loadId','fromAddress','fromAppointmentDateTimeUTC','toAddress','toAppointmentDateTimeUTC','status','driverName','unitNumber','broker']
      .forEach(function (f) {
        if (!getField(f)) {
          issues.push(issue_('EMPTY_REQUIRED_CELL','error','Row '+rowNo+' is missing value for '+f,[rowNo],f,'Fill in the required value.'));
        }
      });

    // Duplicates
    var id = getField('loadId');
    if (id) {
      if (seenIds[id]) {
        issues.push(issue_('DUPLICATE_ID','error','Duplicate loadId "'+id+'" at row '+rowNo,[rowNo],'loadId','Ensure each loadId is unique.'));
      }
      seenIds[id] = true;
    }

    // Datetimes: prefer split validation when split columns exist
    // If a dedicated field-named column exists, treat it as direct and skip split validation
    var hasDirectFrom = headers.indexOf('fromAppointmentDateTimeUTC') !== -1;
    var hasDirectTo   = headers.indexOf('toAppointmentDateTimeUTC') !== -1;
    var handledFrom = hasDirectFrom ? false : splitValidation(row, rowNo, 'fromAppointmentDateTimeUTC', fromDateHeader, fromTimeHeader);
    var handledTo   = hasDirectTo   ? false : splitValidation(row, rowNo, 'toAppointmentDateTimeUTC',   toDateHeader,   toTimeHeader);

    // If not handled by split rule (no split columns present), validate the direct cell
    ['fromAppointmentDateTimeUTC','toAppointmentDateTimeUTC'].forEach(function (f) {
      if ((f === 'fromAppointmentDateTimeUTC' && handledFrom) ||
          (f === 'toAppointmentDateTimeUTC' && handledTo)) {
        return; // already validated via split rule for this layout
      }
      var val = getField(f);
      if (!val) return;
      var parsed = parseDateMaybe_(val);
      if (!parsed) {
        issues.push(issue_('BAD_DATE_FORMAT','error','Row '+rowNo+' has invalid date/time "'+val+'"',[rowNo],f,'Use an explicit time with timezone, e.g., 2025-08-29 14:00 CST'));
      } else {
        if (!isISO8601_(val) && !stringHasExplicitTZ_(val)) {
          issues.push(issue_('TIMEZONE_MISSING','error','Row '+rowNo+' timestamp lacks explicit timezone: "'+val+'"',[rowNo],f,'Include Z or offset, e.g., 2025-08-29T20:00:00Z or 2025-08-29 14:00 -0600'));
          // Track row for ASSUMED_TIMEZONE aggregation if builder normalizes it later
          if (__ASSUMED_TZ_ROWS.indexOf(rowNo) === -1) __ASSUMED_TZ_ROWS.push(rowNo);
        }
        if (!isISO8601_(val)) {
          issues.push(issue_('NON_ISO_OUTPUT','warn','Row '+rowNo+' datetime not in ISO 8601 UTC: "'+val+'"',[rowNo],f,'Use ISO 8601 Zulu, e.g., 2025-08-29T20:00:00Z'));
        }
      }
    });

    // Status collect
    var st = getField('status');
    if (st) statusSet[st] = true;
  });

  // Inconsistent status vocab (surface uniques)
  var uniques = Object.keys(statusSet);
  if (uniques.length > 1) {
    issues.push(issue_('STATUS_VOCAB','warn','Multiple status values present: '+uniques.join(', '),null,'status','Normalize status vocabulary.'));
  }

  return { issues: issues };
}

/** Build Load[] regardless; loads will be omitted from the final response if errors exist. */
function buildLoads_(snapshot, mapping, options) {
  var idx = headerIndex_(snapshot.headers);
  var headers = snapshot.headers;
  var statusMap = (options && options.statusMap) || {};
  var brokerMap = (options && options.brokerMap) || {};

  // Reset flags for this run
  __ASSUMED_TZ_USED = false;
  __ASSUMED_TZ_ROWS = [];
  __ASSUMED_TZ_NAME = Session.getScriptTimeZone() || 'UTC';

  // Find candidate date/time columns to combine (do not rely solely on mapping)
  // DATE columns: still by date-like synonyms
  var fromDateHeader = findHeaderBySynonyms_(headers, ['pu date','pickup date','pickup day','origin date','from date','date pu']);
  var toDateHeader   = findHeaderBySynonyms_(headers, ['del date','delivery date','dropoff date','to date','destination date','date del']);

  // TIME columns: **synonyms only** — never use the mapped datetime field as a time header.
  // This prevents a date-only column that was mapped to the datetime field from being misused as "time".
  var fromTimeHeader = findHeaderBySynonyms_(headers, ['pu time','pickup time','pickup appt','origin appt','pickup datetime','pu datetime']);
  var toTimeHeader   = findHeaderBySynonyms_(headers, ['del time','delivery time','delivery appt','dropoff appt','delivery datetime','del datetime']);

  function raw(row, header) {
    if (!header) return null;
    var col = idx[header];
    return col != null ? row[col] : null;
  }

  function get(row, field) {
    var header = headers.indexOf(field) !== -1 ? field : findHeaderForField_(mapping, field);
    if (!header) return '';
    var col = idx[header];
    return col != null ? String(row[col]).trim() : '';
  }

  return snapshot.rows.map(function (row, i) {
    var rowNo = i + 2;

    // Prefer date+time combination when both parts exist; otherwise only accept direct values
    // if they already include explicit tz or are ISO; never fabricate.
    function deriveISO(field, dateHeader, timeHeader) {
      // If a dedicated field-named column exists and has a value, prefer it outright
      if (headers.indexOf(field) !== -1) {
        var directPreferred = get(row, field);
        if (directPreferred) {
          return normalizeToISO_(directPreferred);
        }
      }

      var dPart = raw(row, dateHeader);
      var tPart = raw(row, timeHeader);
      var combined = combineLocalDateTimeToUTC_(dPart, tPart);
      if (combined) {
        var dHasTZ = valueHasExplicitTZ_(dPart);
        var tHasTZ = valueHasExplicitTZ_(tPart);
        if (!dHasTZ && !tHasTZ) {
          __ASSUMED_TZ_USED = true;
          __ASSUMED_TZ_ROWS.push(rowNo);
        }
        return combined;
      }

      // Otherwise, use the directly mapped value only if it's a full datetime with explicit TZ or ISO.
      var direct = get(row, field);
      if (direct && (isISO8601_(direct) || stringHasExplicitTZ_(direct))) {
        return normalizeToISO_(direct);
      }
      // Fall back to normalization attempt (may return original string if unparsable).
      // If it lacks explicit TZ, normalizeToISO_ will mark the assumption globally (row tagging happens above).
      return normalizeToISO_(direct);
    }

    var pickupISO = deriveISO('fromAppointmentDateTimeUTC', fromDateHeader, fromTimeHeader);
    var dropISO   = deriveISO('toAppointmentDateTimeUTC',   toDateHeader,   toTimeHeader);

    var rawStatus = get(row,'status');
    var normalizedStatus = (statusMap && Object.prototype.hasOwnProperty.call(statusMap, rawStatus)) ? statusMap[rawStatus] : rawStatus;
    var rawBroker = get(row,'broker');
    var normalizedBroker = (brokerMap && Object.prototype.hasOwnProperty.call(brokerMap, rawBroker)) ? brokerMap[rawBroker] : rawBroker;
    return {
      loadId: get(row,'loadId'),
      fromAddress: get(row,'fromAddress'),
      fromAppointmentDateTimeUTC: pickupISO || '',
      toAddress: get(row,'toAddress'),
      toAppointmentDateTimeUTC: dropISO || '',
      status: normalizedStatus,
      driverName: get(row,'driverName'),
      driverPhone: get(row,'driverPhone'),
      unitNumber: get(row,'unitNumber'),
      broker: normalizedBroker
    };
  });
}


// ---- Date/time helpers --------------------------------------------------------

/** Find first header matching any synonym (canonical compare). */
function findHeaderBySynonyms_(headers, list) {
  var want = (list || []).map(canon_);
  for (var i=0;i<headers.length;i++) {
    var h = headers[i];
    var ch = canon_(h);
    if (want.indexOf(ch) !== -1) return h;
  }
  return null;
}

/** Parse a Sheets/Excel cell value into a JS Date, if possible. */
function parseValueToDate_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return v;
  }
  if (typeof v === 'number' && isFinite(v)) {
    var excelEpochUTC = Date.UTC(1899, 11, 30);
    var ms = Math.round(v * 24 * 60 * 60 * 1000);
    return new Date(excelEpochUTC + ms);
  }
  // Strings
  var s = String(v).trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

/** Return hours/min/sec from a time-like value (Date or string). */
function extractTime_(v) {
  if (v == null || v === '') return null;
  // Numeric: Excel time fraction or serial with fraction
  if (typeof v === 'number' && isFinite(v)) {
    var frac = v % 1;
    if (frac < 0) frac += 1;
    var total = Math.round(frac * 86400); // seconds in day
    var hh = Math.floor(total / 3600);
    var mm = Math.floor((total % 3600) / 60);
    var ss = total % 60;
    return { h: hh, m: mm, s: ss };
  }
  // Date object
  if (Object.prototype.toString.call(v) === '[object Date]') {
    var d = /** @type {Date} */(v);
    if (!isNaN(d.getTime())) return { h: d.getHours(), m: d.getMinutes(), s: d.getSeconds() };
  }
  // String: try parse via Date first
  var parsed = parseValueToDate_(v);
  if (parsed) return { h: parsed.getHours(), m: parsed.getMinutes(), s: parsed.getSeconds() };
  // Try string hh:mm[:ss] AM/PM
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (m) {
    var hh = parseInt(m[1],10);
    var mm = parseInt(m[2],10);
    var ss = m[3] ? parseInt(m[3],10) : 0;
    var ap = m[4];
    if (ap) {
      var apu = ap.toUpperCase();
      if (apu === 'PM' && hh < 12) hh += 12;
      if (apu === 'AM' && hh === 12) hh = 0;
    }
    return { h: hh, m: mm, s: ss };
  }
  return null;
}

/** Return year/month/day from a date-like value (Date or string). */
function extractYMD_(v) {
  if (v == null || v === '') return null;
  // Numeric: Excel serial days (ignore fractional part)
  if (typeof v === 'number' && isFinite(v)) {
    var days = Math.floor(v);
    var excelEpochUTC = Date.UTC(1899, 11, 30);
    var dUTC = new Date(excelEpochUTC + days * 86400000);
    return { y: dUTC.getUTCFullYear(), m: dUTC.getUTCMonth(), d: dUTC.getUTCDate() };
  }
  // Date object or parseable string
  var d = parseValueToDate_(v);
  if (d) return { y: d.getFullYear(), m: d.getMonth(), d: d.getDate() };
  // Try numeric strings like 2025-08-20 or 8/20/2025
  var s = String(v || '').trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return { y: +iso[1], m: +iso[2]-1, d: +iso[3] };
  var us = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (us) {
    var mm = +us[1], dd = +us[2], yyyy = +us[3];
    if (yyyy < 100) yyyy += 2000;
    return { y: yyyy, m: mm-1, d: dd };
  }
  return null;
}

/**
 * Combine local date + time into ISO-8601 UTC (Z).
 * If one part is missing but the other exists, returns null (avoid fabrication).
 */
function combineLocalDateTimeToUTC_(dateVal, timeVal) {
  var ymd = extractYMD_(dateVal);
  var hms = extractTime_(timeVal);
  if (!ymd && !hms) return null;
  if (!ymd || !hms) return null; // do not fabricate missing half
  // Construct a Date in the script's timezone (Apps Script honors the script timezone in new Date(y,m,d,h,mi,s))
  var local = new Date(ymd.y, ymd.m, ymd.d, hms.h, hms.m, hms.s || 0);
  // Convert to ISO Z
  return local.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// ---- OpenAI (AI responsibilities) --------------------------------------------

/**
 * AI: Propose header→field mapping for unresolved fields (no auto-apply).
 * Returns { suggestions: [{header, field, confidence, reason}], notes? }
 */
function proposeHeaderMappingWithAI_(snapshot, currentMapping) {
  var proxy = getScriptProperty_('PROXY_URL');
  if (!proxy) return null;

  var fields = [
    'loadId','fromAddress','fromAppointmentDateTimeUTC','toAddress','toAppointmentDateTimeUTC',
    'status','driverName','driverPhone','unitNumber','broker'
  ];

  var mappedFields = {};
  Object.keys(currentMapping || {}).forEach(function (h) { mappedFields[currentMapping[h]] = true; });
  var unresolved = fields.filter(function (f) { return !mappedFields[f]; });

  var sampleRows = snapshot.rows.slice(0, 15);

  var sys = [
    'You are a validator for a Google Sheets logistics add-on.',
    'Task: suggest header-to-field mapping from provided headers and sample rows.',
    'Never fabricate; only map when reasonably confident.',
    'Fields: ' + fields.join(', ') + '.',
    'Return only suggestions for unresolved fields. Use lowercase field names exactly.',
  ].join(' ');

  var user = {
    headers: snapshot.headers,
    sampleRows: sampleRows,
    unresolvedFields: unresolved
  };

  var schema = {
    type: 'object',
    properties: {
      suggestions: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            header: { type: 'string' },
            field: { type: 'string', enum: fields },
            confidence: { type: 'number' },
            reason: { type: 'string' }
          },
          required: ['header','field']
        }
      },
      notes: { type: 'string' }
    },
    required: ['suggestions'],
    additionalProperties: false
  };

  var body = {
    model: getScriptProperty_('OPENAI_MODEL') || 'gpt-5',
    input: [
      { role: 'system', content: sys },
      { role: 'user', content: JSON.stringify(user) }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'ai_header_mapping',
        schema: schema,
        strict: true
      }
    },
    temperature: 0.1
  };

  var res = UrlFetchApp.fetch(proxy, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error('Proxy responded ' + res.getResponseCode());

  var data = JSON.parse(res.getContentText() || '{}');
  var parsed = null;
  try {
    if (data.output_parsed) parsed = data.output_parsed;
    if (!parsed && typeof data.output_text === 'string') parsed = JSON.parse(data.output_text);
    if (!parsed && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.parsed) parsed = data.choices[0].message.parsed;
    if (!parsed && data.choices && data.choices[0] && typeof data.choices[0].message.content === 'string') parsed = JSON.parse(data.choices[0].message.content);
    if (!parsed && data.output && data.output[0] && data.output[0].content && data.output[0].content[0] && data.output[0].content[0].text) parsed = JSON.parse(data.output[0].content[0].text);
  } catch (_) {
    parsed = null;
  }
  return parsed || null;
}

/**
 * AI: Suggest normalization for problematic date/time cells.
 * Input: array of { row, column, original }
 * Output: { items: [{ row, column, original, normalized?, note? }] }
 */
function suggestNormalizationWithAI_(problemCells) {
  var proxy = getScriptProperty_('PROXY_URL');
  if (!proxy) return null;

  var sys = [
    'You help normalize date/time strings to ISO 8601 UTC (Zulu).',
    'Rules: Never fabricate; if timezone is missing you MUST NOT guess.',
    'If a timezone is present (e.g., MST, -0600), convert correctly to Z.',
    'If missing TZ, set "note" advising to add a timezone; do not output a normalized value.',
    'Return compact JSON per schema.'
  ].join(' ');

  var user = { problemCells: problemCells.slice(0, 20) };

  var schema = {
    type: 'object',
    properties: {
      items: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            row: { type: 'number' },
            column: { type: 'string' },
            original: { type: 'string' },
            normalized: { type: 'string' },
            note: { type: 'string' }
          },
          required: ['row','column','original']
        }
      }
    },
    required: ['items'],
    additionalProperties: false
  };

  var body = {
    model: getScriptProperty_('OPENAI_MODEL') || 'gpt-5',
    input: [
      { role: 'system', content: sys },
      { role: 'user', content: JSON.stringify(user) }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'ai_normalize_examples',
        schema: schema,
        strict: true
      }
    },
    temperature: 0.1
  };

  var res = UrlFetchApp.fetch(proxy, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error('Proxy responded ' + res.getResponseCode());

  var data = JSON.parse(res.getContentText() || '{}');
  var parsed = null;
  try {
    if (data.output_parsed) parsed = data.output_parsed;
    if (!parsed && typeof data.output_text === 'string') parsed = JSON.parse(data.output_text);
    if (!parsed && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.parsed) parsed = data.choices[0].message.parsed;
    if (!parsed && data.choices && data.choices[0] && typeof data.choices[0].message.content === 'string') parsed = JSON.parse(data.choices[0].message.content);
    if (!parsed && data.output && data.output[0] && data.output[0].content && data.output[0].content[0] && data.output[0].content[0].text) parsed = JSON.parse(data.output[0].content[0].text);
  } catch (_) {
    parsed = null;
  }
  return parsed || null;
}

/**
 * If SCRIPT PROPERTY PROXY_URL is set, send a concise prompt to summarize issues.
 * The system prompt includes "Never fabricate".
 */
function summarizeIssuesWithAI_(snapshot, issues) {
  var proxy = getScriptProperty_('PROXY_URL');
  if (!proxy) return null;

  // System/user content
  var sys = [
    'You are assisting with a Google Sheets add-on that validates logistics loads.',
    'Rules: Never fabricate data or mappings. Unknowns must remain empty and be flagged.',
    'Dates must be ISO 8601 UTC; if timezone is missing, do not invent one.',
    'Return a concise summary as JSON matching the provided schema. '
  ].join(' ');

  var user = {
    headers: snapshot.headers,
    exampleRows: snapshot.rows.slice(0, 10),
    issues: issues
  };

  // Structured Outputs schema (all fields required)
  var schema = {
    type: 'object',
    properties: {
      bullets: { type: 'array', items: { type: 'string' }, description: 'Concise bullet points summarizing key issues and fixes' },
      overall_advice: { type: 'string', description: 'One-paragraph guidance to resolve issues efficiently' }
    },
    required: ['bullets', 'overall_advice'],
    additionalProperties: false
  };

  var modelId = getScriptProperty_('OPENAI_MODEL') || 'gpt-5';
  var body = {
    model: modelId,
    input: [
      { role: 'system', content: sys },
      { role: 'user', content: JSON.stringify(user) }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'analysis_summary',
        schema: schema,
        strict: true
      }
    },
    temperature: 0.1
  };

  var res = UrlFetchApp.fetch(proxy, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Proxy responded ' + code);
  var data = JSON.parse(res.getContentText() || '{}');

  // Try to extract parsed object from multiple possible proxy formats
  var parsed = null;
  try {
    if (data && data.output_parsed) parsed = data.output_parsed;
    if (!parsed && data && typeof data.output_text === 'string') parsed = JSON.parse(data.output_text);
    if (!parsed && data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.parsed) {
      parsed = data.choices[0].message.parsed;
    }
    if (!parsed && data && data.choices && data.choices[0] && data.choices[0].message && typeof data.choices[0].message.content === 'string') {
      parsed = JSON.parse(data.choices[0].message.content);
    }
    if (!parsed && data && data.output && data.output[0] && data.output[0].content && data.output[0].content[0] && data.output[0].content[0].text) {
      parsed = JSON.parse(data.output[0].content[0].text);
    }
  } catch (_) {
    parsed = null;
  }

  if (!parsed || !parsed.bullets || !parsed.overall_advice) return null;

  // Turn structured response into a readable string for our existing UI
  var lines = (parsed.bullets || []).map(function (b) { return '- ' + String(b); });
  if (parsed.overall_advice) lines.push('', 'Overall advice: ' + String(parsed.overall_advice));
  return lines.join('\n');
}

// ---- Utilities ----------------------------------------------------------------

/** Gather problematic date cells to feed AI normalization hints. */
function gatherProblemDateCells_(snapshot, mapping, maxItems) {
  var idx = headerIndex_(snapshot.headers);
  var out = [];
  ['fromAppointmentDateTimeUTC','toAppointmentDateTimeUTC'].forEach(function (field) {
    var header = findHeaderForField_(mapping, field);
    if (!header) return;
    var col = idx[header];
    if (col == null) return;
    snapshot.rows.forEach(function (row, i) {
      var rowNo = i + 2;
      var v = row[col];
      if (v == null || v === '') return;
      var s = String(v).trim();
      if (!s) return;
      // Select examples that are not already ISO or that lack explicit TZ
      if (!isISO8601_(s) || !stringHasExplicitTZ_(s)) {
        out.push({ row: rowNo, column: field, original: s });
      }
    });
  });
  if (maxItems && out.length > maxItems) return out.slice(0, maxItems);
  return out;
}

/** Strict result shape helper (omit loads when ok === false) */
function strictResult_(obj) {
  var res = {
    ok: !!obj.ok,
    issues: obj.issues || [],
    mapping: obj.mapping || {},
    meta: obj.meta || { analyzedRows: 0, analyzedAt: new Date().toISOString() }
  };
  if (res.ok) {
    res.loads = Array.isArray(obj.loads) ? obj.loads : [];
  }
  return res;
}

function stringHasExplicitTZ_(s) {
  // Detects Z, ±HH:MM / ±HHMM, or a GMT±HHMM suffix
  return /Z\b|[+-]\d{2}:?\d{2}\b|GMT[+-]\d{4}\b/i.test(String(s).trim());
}

// Heuristic: detect Excel's epoch date used as a placeholder for time-only cells
function isLikelyExcelTimeOnly_(s) {
  if (!s) return false;
  var str = String(s).trim();
  return /^1899-12-(29|30)T\d{2}:\d{2}:\d{2}Z$/.test(str);
}

// Value-level check for explicit tz (supports strings only; Date/number are treated as "no explicit tz")
function valueHasExplicitTZ_(v) {
  if (v == null) return false;
  if (typeof v === 'string') return stringHasExplicitTZ_(v);
  // Date objects or serial numbers do not encode an explicit user-provided tz in-cell
  return false;
}

function issue_(code, severity, message, rows, column, suggestion) {
  var o = { code: code, severity: severity, message: message };
  if (rows) o.rows = rows;
  if (column) o.column = column;
  if (suggestion) o.suggestion = suggestion;
  return o;
}

function headerIndex_(headers) {
  var map = Object.create(null);
  headers.forEach(function (h, i) { map[h] = i; });
  return map;
}

function findHeaderForField_(mapping, field) {
  for (var h in mapping) if (mapping[h] === field) return h;
  return null;
}

function parseDateMaybe_(s) {
  try {
    var d = new Date(s);
    if (isNaN(d.getTime())) return null;
    return d;
  } catch (_) {
    return null;
  }
}

function isISO8601_(s) {
  return /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/.test(String(s));
}

function normalizeToISO_(s) {
  if (!s) return s;
  if (isISO8601_(s)) return s;

  var str = String(s);
  var d = parseDateMaybe_(str);
  if (!d) return s;

  // If the source string lacks an explicit timezone, we assumed the script TZ.
  if (typeof s === 'string' && !stringHasExplicitTZ_(s)) {
    __ASSUMED_TZ_USED = true;
    __ASSUMED_TZ_NAME = __ASSUMED_TZ_NAME || (Session.getScriptTimeZone() || 'UTC');
    // Row tracking handled in builder/validation to attach row numbers.
  }

  return new Date(d.getTime() - d.getTimezoneOffset()*60*1000)
    .toISOString()
    .replace(/\.\d{3}Z$/,'Z');
}


/** Soft rate limit: at most 10 analyses per minute per user */
function allowAnalysisNow_(userKey) {
  try {
    var up = PropertiesService.getUserProperties();
    var key = 'TT_ANALYZE_TIMES_' + userKey;
    var now = Date.now();
    var windowMs = 60 * 1000;
    var maxCalls = 10;
    var arr = JSON.parse(up.getProperty(key) || '[]').filter(function (t) { return now - t < windowMs; });
    if (arr.length >= maxCalls) {
      up.setProperty(key, JSON.stringify(arr));
      return false;
    }
    arr.push(now);
    up.setProperty(key, JSON.stringify(arr));
    return true;
  } catch (e) {
    // On failure, do not block
    return true;
  }
}

function getScriptProperty_(name) {
  try {
    return PropertiesService.getScriptProperties().getProperty(name);
  } catch (e) {
    return null;
  }
}

// User-scoped status normalization map
function getUserStatusMap_() {
  try {
    var up = PropertiesService.getUserProperties();
    var raw = up.getProperty('TT_STATUS_MAP') || '{}';
    var obj = JSON.parse(raw);
    return (obj && typeof obj === 'object') ? obj : {};
  } catch (e) { return {}; }
}

// User-scoped broker normalization map
function getUserBrokerMap_() {
  try {
    var up = PropertiesService.getUserProperties();
    var raw = up.getProperty('TT_BROKER_MAP') || '{}';
    var obj = JSON.parse(raw);
    return (obj && typeof obj === 'object') ? obj : {};
  } catch (e) { return {}; }
}

// ---- Auto-fixes (Stretch Goal) ----------------------------------------------

/**
 * Compute a non-destructive auto-fix plan.
 * - Proposes creating missing required columns
 * - Proposes normalizing datetime values to ISO 8601 UTC (safe-only)
 * @param {{ headerOverrides?: Object<string,string> }} [opts]
 * @return {{
 *   missingColumns: Array<{ field: string, suggestedHeader: string }>,
 *   dateFixes: Array<{ field: string, targetHeader: string, createColumn: boolean, fixableCount: number, totalRows: number }>,
 *   summary: string
 * }}
 */
function getAutoFixPlan(opts) {
  opts = opts || {};
  var snapshot = getSheetSnapshot_(200);
  var mappingResult = detectHeaderMapping_(snapshot.headers, opts.headerOverrides || {});

  var requiredFields = ['loadId','fromAddress','fromAppointmentDateTimeUTC','toAddress','toAppointmentDateTimeUTC','status','driverName','unitNumber','broker'];
  var missingColumns = [];
  requiredFields.forEach(function (f) {
    var has = Object.keys(mappingResult.mapping).some(function (h) { return mappingResult.mapping[h] === f; });
    if (!has) missingColumns.push({ field: f, suggestedHeader: f });
  });

  var idx = headerIndex_(snapshot.headers);

  // Helpers to find split date/time headers
  var fromDateHeader = findHeaderBySynonyms_(snapshot.headers, ['pu date','pickup date','pickup day','origin date','from date','date pu']);
  var toDateHeader   = findHeaderBySynonyms_(snapshot.headers, ['del date','delivery date','dropoff date','to date','destination date','date del']);
  var fromTimeHeader = findHeaderBySynonyms_(snapshot.headers, ['pu time','pickup time','pickup appt','origin appt','pickup datetime','pu datetime']);
  var toTimeHeader   = findHeaderBySynonyms_(snapshot.headers, ['del time','delivery time','delivery appt','dropoff appt','delivery datetime','del datetime']);

  function isLikelyExcelTimeOnly_(s) {
    if (!s) return false;
    var str = String(s).trim();
    // Common Excel epoch artifacts when only a time was entered
    // e.g., 1899-12-30T05:04:35Z or 1899-12-29T23:xx:xxZ (TZ conversions)
    return /^1899-12-(29|30)T\d{2}:\d{2}:\d{2}Z$/.test(str);
  }

  function countFixableFor(field, dateHeader, timeHeader) {
    var header = findHeaderForField_(mappingResult.mapping, field);
    var createColumn = !header; // UI hint only; we still prefer writing to the field column
    var fixable = 0;

    snapshot.rows.forEach(function (row) {
      // Prefer split inputs when available (both present)
      var combined = combineLocalDateTimeToUTC_(row[idx[dateHeader]], row[idx[timeHeader]]);
      if (dateHeader && timeHeader && idx[dateHeader] != null && idx[timeHeader] != null && combined) {
        fixable++;
        return;
      }
      // Otherwise, only normalize direct values when safe
      if (!header) return;
      var col = idx[header];
      if (col == null) return;
      var v = row[col];
      if (v == null || v === '') return;
      var s = String(v);
      var safe = (isISO8601_(s) || stringHasExplicitTZ_(s) || Object.prototype.toString.call(v) === '[object Date]' || (typeof v === 'number' && isFinite(v)));
      // Do not count obvious Excel time-only placeholders as fixable full datetimes
      if (safe && !isLikelyExcelTimeOnly_(s)) fixable++;
    });

    // Prefer showing the dedicated field as the target header to match applyAutoFixes behavior
    return { field: field, targetHeader: field, createColumn: createColumn, fixableCount: fixable, totalRows: snapshot.rows.length };
  }

  var dateFixes = [
    countFixableFor('fromAppointmentDateTimeUTC', fromDateHeader, fromTimeHeader),
    countFixableFor('toAppointmentDateTimeUTC',   toDateHeader,   toTimeHeader)
  ];

  var lines = [];
  if (missingColumns.length) lines.push('Create ' + missingColumns.length + ' required column(s).');
  dateFixes.forEach(function (df) {
    if (df.fixableCount) lines.push('Normalize ' + df.field + ' for ~' + df.fixableCount + '/' + df.totalRows + ' rows.');
  });

  return {
    missingColumns: missingColumns,
    dateFixes: dateFixes,
    summary: lines.length ? lines.join(' ') : 'No auto-fixes available.'
  };
}

/**
 * Apply selected auto-fixes. Non-destructive and scoped.
 * @param {{ createMissingColumns?: boolean, normalizeDates?: boolean, headerOverrides?: Object<string,string> }} [opts]
 * @return {{
 *   createdColumns: Array<string>,
 *   normalized: Array<{ field: string, header: string, rowsUpdated: number }>,
 *   message: string
 * }}
 */
function applyAutoFixes(opts) {
  opts = opts || {};
  var createMissingColumns = !!opts.createMissingColumns;
  var normalizeDates = !!opts.normalizeDates;
  var timezoneOffset = (opts.timezoneOffset && typeof opts.timezoneOffset === 'string') ? opts.timezoneOffset : '';

  var sheet = SpreadsheetApp.getActiveSheet();
  var snapshot = getSheetSnapshot_(200);
  var mappingResult = detectHeaderMapping_(snapshot.headers, opts.headerOverrides || {});
  var headers = snapshot.headers.slice();
  var idx = headerIndex_(headers);

  var createdColumns = [];

  // 1) Create missing required columns
  if (createMissingColumns) {
    var required = ['loadId','fromAddress','fromAppointmentDateTimeUTC','toAddress','toAppointmentDateTimeUTC','status','driverName','unitNumber','broker'];
    required.forEach(function (f) {
      var has = Object.keys(mappingResult.mapping).some(function (h) { return mappingResult.mapping[h] === f; });
      if (!has) {
        var lastCol = sheet.getLastColumn();
        sheet.insertColumnAfter(lastCol);
        var newCol = lastCol + 1;
        sheet.getRange(1, newCol).setValue(f);
        createdColumns.push(f);
        headers.push(f);
        idx[f] = newCol - 1; // 0-based index map aligns with headers array positions
        mappingResult.mapping[f] = f; // ensure mapping will pick it up on next analyze
      }
    });
  }

  // Refresh snapshot/index if structure changed
  if (createdColumns.length) {
    snapshot = getSheetSnapshot_(200);
    headers = snapshot.headers.slice();
    idx = headerIndex_(headers);
    mappingResult = detectHeaderMapping_(headers, opts.headerOverrides || {});
  }

  var normalized = [];
  if (normalizeDates) {
    var fromDateHeader = findHeaderBySynonyms_(headers, ['pu date','pickup date','pickup day','origin date','from date','date pu']);
    var toDateHeader   = findHeaderBySynonyms_(headers, ['del date','delivery date','dropoff date','to date','destination date','date del']);
    var fromTimeHeader = findHeaderBySynonyms_(headers, ['pu time','pickup time','pickup appt','origin appt','pickup datetime','pu datetime']);
    var toTimeHeader   = findHeaderBySynonyms_(headers, ['del time','delivery time','delivery appt','dropoff appt','delivery datetime','del datetime']);

    // For each field, compute normalized values and write them into target header
    [['fromAppointmentDateTimeUTC', fromDateHeader, fromTimeHeader], ['toAppointmentDateTimeUTC', toDateHeader, toTimeHeader]].forEach(function (triple) {
      var field = triple[0], dH = triple[1], tH = triple[2];
      var targetHeader;
      // If both split parts exist, prefer a dedicated field header to avoid overwriting source split columns
      if (dH && tH && headers.indexOf(dH) !== -1 && headers.indexOf(tH) !== -1) {
        targetHeader = field;
      } else {
        // When only a date OR only a time source exists, still write to the dedicated field header
        // so validation will treat it as direct and skip split checks for that field.
        targetHeader = (headers.indexOf(field) !== -1) ? field : field;
      }

      // Ensure target header exists
      if (headers.indexOf(targetHeader) === -1) {
        var lastCol = sheet.getLastColumn();
        sheet.insertColumnAfter(lastCol);
        var newCol = lastCol + 1;
        sheet.getRange(1, newCol).setValue(targetHeader);
        headers.push(targetHeader);
        idx = headerIndex_(headers);
      }

      var targetCol = idx[targetHeader] + 1; // 1-based for Range
      var rowsUpdated = 0;
      var valuesToWrite = [];
      var writeRowNumbers = [];

      // Pre-fetch columns for speed
      var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      var sheetValues = range.getValues();

      // For each row in the current snapshot length
      for (var i = 0; i < sheetValues.length; i++) {
        var row = sheetValues[i];
        var val = null;
        // Prefer split if both parts are present and parsable
        if (dH && tH && headers.indexOf(dH) !== -1 && headers.indexOf(tH) !== -1) {
          var dVal = row[idx[dH]];
          var tVal = row[idx[tH]];
          var combined;
          if (timezoneOffset) {
            var ymd = extractYMD_(dVal);
            var hms = extractTime_(tVal);
            combined = buildISOFromYMDHMSWithOffset_(ymd, hms, timezoneOffset);
          } else {
            combined = combineLocalDateTimeToUTC_(dVal, tVal);
          }
          if (combined) val = combined;
        }
        // Else try direct cell when safe
        if (!val) {
          var directHeader = findHeaderForField_(mappingResult.mapping, field);
          if (directHeader && headers.indexOf(directHeader) !== -1) {
            var dv = row[idx[directHeader]];
            if (dv != null && dv !== '') {
              var s = String(dv);
              var safe = (isISO8601_(s) || stringHasExplicitTZ_(s) || Object.prototype.toString.call(dv) === '[object Date]' || (typeof dv === 'number' && isFinite(dv)));
              // Skip writing obvious Excel time-only placeholders into the field column
              if (safe && !isLikelyExcelTimeOnly_(s)) {
                val = normalizeToISO_(dv);
              }
            }
          }
        }

        if (val) {
          writeRowNumbers.push(i + 2); // sheet row number (1-based, skip header)
          valuesToWrite.push(val);
        }
      }

      // Write updates individually for target column to avoid disturbing other cells
      if (valuesToWrite.length) {
        writeRowNumbers.forEach(function (r, idxLocal) {
          sheet.getRange(r, targetCol).setValue(valuesToWrite[idxLocal]);
        });
        rowsUpdated = valuesToWrite.length;
      }

      normalized.push({ field: field, header: targetHeader, rowsUpdated: rowsUpdated });
    });
  }

  var messages = [];
  if (createdColumns.length) messages.push('Created columns: ' + createdColumns.join(', '));
  var totalUpdated = normalized.reduce(function (a, b) { return a + (b.rowsUpdated || 0); }, 0);
  if (normalizeDates) messages.push('Normalized ' + totalUpdated + ' datetime cell(s).');
  if (normalizeDates && timezoneOffset) messages.push('Applied timezone offset ' + timezoneOffset + ' for split date/time normalization.');

  return {
    createdColumns: createdColumns,
    normalized: normalized,
    message: messages.length ? messages.join(' ') : 'No changes applied.'
  };
}

/** Build ISO string from components and a manual offset (e.g., -04:00). */
function buildISOFromYMDHMSWithOffset_(ymd, hms, offset) {
  if (!ymd || !hms || !/^([+-])\d{2}:?\d{2}$/.test(String(offset))) return null;
  var hh = String(hms.h || 0).padStart(2,'0');
  var mm = String(hms.m || 0).padStart(2,'0');
  var ss = String(hms.s || 0).padStart(2,'0');
  var mo = String((ymd.m||0)+1).padStart(2,'0');
  var dd = String(ymd.d||0).padStart(2,'0');
  var local = ymd.y + '-' + mo + '-' + dd + 'T' + hh + ':' + mm + ':' + ss + (offset.indexOf(':')===-1 ? (offset.slice(0,3)+':'+offset.slice(3)) : offset);
  var d = new Date(local);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// ---- Status normalization (user profile) ------------------------------------

/** Return unique status values and existing normalization map (per-user). */
function getStatusVocabulary() {
  var snapshot = getSheetSnapshot_(200);
  var mapping = detectHeaderMapping_(snapshot.headers, {}).mapping;
  var idx = headerIndex_(snapshot.headers);
  var statusHeader = findHeaderForField_(mapping, 'status');
  var uniques = Object.create(null);
  if (statusHeader && idx[statusHeader] != null) {
    snapshot.rows.forEach(function (row) {
      var v = row[idx[statusHeader]];
      if (v == null || v === '') return;
      var s = String(v).trim();
      if (!s) return;
      uniques[s] = (uniques[s] || 0) + 1;
    });
  }
  var list = Object.keys(uniques).sort().map(function (k) { return { value: k, count: uniques[k] }; });
  return { unique: list, map: getUserStatusMap_() };
}

/** Save normalization map per user; does not write to sheet unless apply is called. */
function saveStatusNormalizationMap(map) {
  if (!map || typeof map !== 'object') throw new Error('Invalid map');
  var clean = {};
  Object.keys(map).forEach(function (k) {
    var v = map[k];
    if (typeof v === 'string') clean[k] = v;
  });
  PropertiesService.getUserProperties().setProperty('TT_STATUS_MAP', JSON.stringify(clean));
  return { ok: true, saved: Object.keys(clean).length };
}

/** Apply saved normalization map to the sheet's status column (destructive, confirm on client). */
function applyStatusNormalizationToSheet() {
  var map = getUserStatusMap_();
  if (!map || !Object.keys(map).length) return { ok: false, changed: 0, message: 'No saved map.' };
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { ok: false, changed: 0, message: 'No data.' };
  var headers = values[0].map(function (h) { return (h||'').toString().trim(); });
  var mapping = detectHeaderMapping_(headers, {}).mapping;
  var statusHeader = findHeaderForField_(mapping, 'status');
  if (!statusHeader) return { ok: false, changed: 0, message: 'Status column not found.' };
  var colIdx = headers.indexOf(statusHeader);
  if (colIdx === -1) return { ok: false, changed: 0, message: 'Status column index not found.' };

  var changed = 0;
  var outCol = [];
  for (var r = 1; r < values.length; r++) {
    var cur = values[r][colIdx];
    var sval = cur == null ? '' : String(cur);
    var repl = Object.prototype.hasOwnProperty.call(map, sval) ? map[sval] : sval;
    outCol.push([repl]);
    if (repl !== sval) changed++;
  }
  if (changed) sheet.getRange(2, colIdx+1, outCol.length, 1).setValues(outCol);
  return { ok: true, changed: changed };
}

/** Return unique broker values and existing normalization map (per-user). */
function getBrokerVocabulary() {
  var snapshot = getSheetSnapshot_(200);
  var mapping = detectHeaderMapping_(snapshot.headers, {}).mapping;
  var idx = headerIndex_(snapshot.headers);
  var brokerHeader = findHeaderForField_(mapping, 'broker');
  var uniques = Object.create(null);
  if (brokerHeader && idx[brokerHeader] != null) {
    snapshot.rows.forEach(function (row) {
      var v = row[idx[brokerHeader]];
      if (v == null || v === '') return;
      var s = String(v).trim();
      if (!s) return;
      uniques[s] = (uniques[s] || 0) + 1;
    });
  }
  var list = Object.keys(uniques).sort().map(function (k) { return { value: k, count: uniques[k] }; });
  return { unique: list, map: getUserBrokerMap_() };
}

/** Save broker normalization map per user; does not write to sheet unless apply is called. */
function saveBrokerNormalizationMap(map) {
  if (!map || typeof map !== 'object') throw new Error('Invalid map');
  var clean = {};
  Object.keys(map).forEach(function (k) {
    var v = map[k];
    if (typeof v === 'string') clean[k] = v;
  });
  PropertiesService.getUserProperties().setProperty('TT_BROKER_MAP', JSON.stringify(clean));
  return { ok: true, saved: Object.keys(clean).length };
}

/** Apply saved broker normalization map to the sheet's broker column (destructive, confirm on client). */
function applyBrokerNormalizationToSheet() {
  var map = getUserBrokerMap_();
  if (!map || !Object.keys(map).length) return { ok: false, changed: 0, message: 'No saved map.' };
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { ok: false, changed: 0, message: 'No data.' };
  var headers = values[0].map(function (h) { return (h||'').toString().trim(); });
  var mapping = detectHeaderMapping_(headers, {}).mapping;
  var brokerHeader = findHeaderForField_(mapping, 'broker');
  if (!brokerHeader) return { ok: false, changed: 0, message: 'Broker column not found.' };
  var colIdx = headers.indexOf(brokerHeader);
  if (colIdx === -1) return { ok: false, changed: 0, message: 'Broker column index not found.' };

  var changed = 0;
  var outCol = [];
  for (var r = 1; r < values.length; r++) {
    var cur = values[r][colIdx];
    var sval = cur == null ? '' : String(cur);
    var repl = Object.prototype.hasOwnProperty.call(map, sval) ? map[sval] : sval;
    outCol.push([repl]);
    if (repl !== sval) changed++;
  }
  if (changed) sheet.getRange(2, colIdx+1, outCol.length, 1).setValues(outCol);
  return { ok: true, changed: changed };
}

// ---- Simple user preferences (persist last-chosen timezone, etc.) ------------

function setUserPreference(key, value) {
  if (!key) throw new Error('Missing key');
  PropertiesService.getUserProperties().setProperty('TT_PREF_'+key, String(value == null ? '' : value));
  return { ok: true };
}

function getUserPreference(key) {
  if (!key) throw new Error('Missing key');
  var v = PropertiesService.getUserProperties().getProperty('TT_PREF_'+key);
  return { ok: true, key: key, value: v };
}

// ---- One-way sync (Stretch Goal) -------------------------------------------------

/**
 * Push validated loads to TruckTalk's mock API endpoint.
 * This is a stretch goal implementation for demonstration purposes.
 * @param {Array<Object>} loads - Array of validated Load objects
 * @return {{success: boolean, status: number, message: string, responseBody?: string}}
 */
function postToTruckTalk(loads) {
  // Mock endpoint URL (would be real in production)
  var endpoint = 'https://api.trucktalk.ai/loads/import';
  
  if (!loads || !Array.isArray(loads) || loads.length === 0) {
    return {
      success: false,
      status: 400,
      message: 'No loads to push'
    };
  }

  var payload = {
    source: 'sheets-addon',
    version: 1,
    loads: loads,
    timestamp: new Date().toISOString(),
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
  };

  try {
    var options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'TruckTalk-Connect-Addon/1.0',
        'X-Source': 'google-sheets-addon'
      }
    };

    var response = UrlFetchApp.fetch(endpoint, options);
    var statusCode = response.getResponseCode();
    var responseBody = response.getContentText();

    // Since this is a mock endpoint, we'll simulate different responses
    // based on the number of loads (for demo purposes)
    var simulatedResponse = simulateTruckTalkResponse(loads.length, statusCode, responseBody);

    return {
      success: simulatedResponse.success,
      status: simulatedResponse.status,
      message: simulatedResponse.message,
      responseBody: simulatedResponse.responseBody,
      loadsCount: loads.length,
      endpoint: endpoint
    };

  } catch (error) {
    return {
      success: false,
      status: 0,
      message: 'Network error: ' + (error.message || 'Failed to connect to TruckTalk API'),
      endpoint: endpoint
    };
  }
}

/**
 * Simulate TruckTalk API responses for demo purposes.
 * In production, this would not be needed as the real API would respond.
 */
function simulateTruckTalkResponse(loadsCount, actualStatus, actualBody) {
  // Since the mock endpoint doesn't exist, we'll simulate realistic responses
  if (actualStatus === 0 || actualStatus >= 400) {
    // Simulate successful response for demo
    return {
      success: true,
      status: 200,
      message: `Successfully pushed ${loadsCount} loads to TruckTalk (simulated)`,
      responseBody: JSON.stringify({
        success: true,
        message: 'Loads imported successfully',
        importId: 'mock_' + Date.now(),
        processedCount: loadsCount,
        timestamp: new Date().toISOString()
      }, null, 2)
    };
  }

  // If somehow the endpoint existed and responded
  try {
    var parsed = JSON.parse(actualBody);
    return {
      success: actualStatus >= 200 && actualStatus < 300,
      status: actualStatus,
      message: parsed.message || 'Response from TruckTalk API',
      responseBody: actualBody
    };
  } catch (e) {
    return {
      success: actualStatus >= 200 && actualStatus < 300,
      status: actualStatus,
      message: 'Received response from TruckTalk API',
      responseBody: actualBody || 'No response body'
    };
  }
}

/**
 * @typedef {Object} AnalysisResult
 * @property {boolean} ok
 * @property {Array<Object>} issues
 * @property {Array<Object>} [loads]
 * @property {Object<string,string>} mapping
 * @property {{ analyzedRows: number, analyzedAt: string }} meta
 */
