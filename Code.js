const SPREADSHEET_ID = '1Y0x80Z-NGrX5Azp3vLCNDk0bkPONLdpX24d5UVRPxGU'; // 'DaWiz PO 3.0' Sheet
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('DaWiz Playoffs 3.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // Allows embedding
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=yes, shrink-to-fit=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================================================================
 * --- API ENDPOINTS FOR CLIENT-SIDE RENDER ---
 * These functions act as a pure, lightweight JSON API for the browser.
 * ========================================================================= */

/**
 * Helper function to securely parse the 'Owners' sheet and extract dynamically valid Name & PIN maps.
 */
function getOwnersData() {
  const sheet = ss.getSheetByName('Owners');
  if (!sheet) return [];
  const data = sheet.getRange("A1:K30").getValues();

  // Automatically find Header Row mapping to avoid structural fragility
  let headerRowIndex = 0;
  for (let i = 0; i < 5; i++) {
    const upperRow = data[i].map(h => String(h).toUpperCase());
    if (upperRow.includes("NAME") || upperRow.includes("PIN")) {
      headerRowIndex = i; break;
    }
  }

  let nameCol = 2; // Default Col C
  let pinCol = 4;  // Default PIN Col E
  let inCol = 6;   // Default In/Out Col G

  const headers = data[headerRowIndex].map(h => String(h).trim().toUpperCase());
  if (headers.indexOf("NAME") !== -1) nameCol = headers.indexOf("NAME");
  if (headers.indexOf("PIN") !== -1) pinCol = headers.indexOf("PIN");
  if (headers.indexOf("IN OR OUT") !== -1) inCol = headers.indexOf("IN OR OUT");

  const validOwners = [];
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[nameCol]).trim();
    const pinMatch = String(row[pinCol]).trim();
    const status = String(row[inCol]).trim();

    // Map the owner ONLY if they are actively opted "IN"
    if (name && status.toUpperCase() === "IN") {
      validOwners.push({ name: name, pin: pinMatch });
    }
  }
  return validOwners;
}

/**
 * Returns a clean list of valid 'IN' owners to populate the frontend dropdown.
 */
function getOwnersList() {
  const owners = getOwnersData();
  if (owners.length === 0) return ["Fallback1", "Fallback2"];
  return owners.map(o => o.name);
}

/**
 * Checks if the global playoff draft cutoff date has passed.
 */
function isCutoffPassed() {
  try {
    const processSS = SpreadsheetApp.openById('1_KiAdEYxXwqaMpzKdStATpGmOnJ41ZbKjidBXiiu1-8');
    const informationSheet = processSS.getSheetByName('Information');
    if (informationSheet) {
      const cutoffValue = informationSheet.getRange("F17").getValue();
      if (cutoffValue instanceof Date) {
        const cutoffDate = new Date(cutoffValue);
        cutoffDate.setHours(0, 0, 0, 0);
        return Date.now() > cutoffDate.getTime();
      }
    }
  } catch (err) {
    console.error("External Cutoff Check Failed: ", err);
  }
  return false;
}

/**
 * Determines which EventStore 'League Scoring' tab to use based on the
 * regular season end date stored in the local 'Information' sheet at C2.
 * Returns 'League Scoring - Reg Season' before that date,
 * and 'League Scoring - Playoffs' on or after it.
 */
function getActiveScoringSheetName() {
  try {
    const infoSh = ss.getSheetByName('Information');
    if (infoSh) {
      const seasonEndVal = infoSh.getRange('C2').getValue();
      if (seasonEndVal instanceof Date) {
        const seasonEnd = new Date(seasonEndVal);
        seasonEnd.setHours(0, 0, 0, 0);
        return Date.now() >= seasonEnd.getTime()
          ? 'League Scoring - Playoffs'
          : 'League Scoring - Reg Season';
      }
    }
  } catch (e) {
    console.error('getActiveScoringSheetName failed: ' + e.message);
  }
  // Safe fallback — playoffs tab is the primary use case for this app
  return 'League Scoring - Playoffs';
}

/**
 * Provides the global state required to boot the Client UI securely.
 */
function getAppBootState() {
  return {
    ownersList: getOwnersList(),
    isCutoffPassed: isCutoffPassed()
  };
}

/**
 * Validates the user's PIN natively in JavaScript using the 'Owners' mapping array.
 * Highly optimized relative to Spreadsheet lock loops!
 * Returns { valid: boolean, rickRoll: boolean }
 */
function authenticateUser(name, pin) {
  if (!pin || !name) return { valid: false, rickRoll: false };

  // 1. Dynamic Cutoff check
  if (isCutoffPassed()) {
    return { valid: false, rickRoll: true };
  }

  // 2. Hidden Developer Override for instant testing
  if (pin === "9999") {
    return { valid: false, rickRoll: true };
  }

  // Pure JavaScript Native Validation 
  // (Replaces the slow, restrictive SpreadsheetApp.flush() grid locks of legacy implementation)
  const owners = getOwnersData();
  const validMatch = owners.find(o => o.name === name && o.pin === String(pin).trim());

  if (validMatch) {
    return { valid: true, rickRoll: false };
  } else {
    return { valid: false, rickRoll: false };
  }
}

/**
 * Executes a powerful failsafe check on the IMPORTRANGE formulas.
 * If data is missing in a critical target tab due to a Google Sheets glitch, 
 * it automatically restores the formula string securely stored in the 'Information' tab.
 */
function ensureImportRangesActive() {
  const infoSh = ss.getSheetByName('Information');
  if (!infoSh) return;

  // Formula maps: Owners:A1, OTbyID:A2, League Scoring:A2, Playoff Bound (League Scoring):M1
  const rules = [
    { sheet: 'Owners', cell: 'A1' },
    { sheet: 'OTbyID', cell: 'A2' },
    { sheet: 'League Scoring', cell: 'A2' },
    { sheet: 'League Scoring', cell: 'M1' }
  ];

  rules.forEach((rule, idx) => {
    const targetSheet = ss.getSheetByName(rule.sheet);
    if (targetSheet) {
      const checkCell = targetSheet.getRange(rule.cell);
      if (checkCell.isBlank()) {
        // Pull the formula text from Information tab (B3:C6)
        const textB = infoSh.getRange(3 + idx, 2).getValue().toString();
        const textC = infoSh.getRange(3 + idx, 3).getValue().toString();

        // Extract whichever cell holds the literal formula string
        const formulaString = textB.toUpperCase().includes("IMPORTRANGE") ? textB :
          (textC.toUpperCase().includes("IMPORTRANGE") ? textC : null);

        if (formulaString) {
          checkCell.setFormula(formulaString);
        }
      }
    }
  });
}

/**
 * The heavy lifter! Fetches the master Roster, Bench, and Draft Pool purely based on Player IDs.
 * Outputs the organized arrays inside clean JSON for Client-Side JS cross-referencing.
 */
function getDraftData(name) {
  // 1. Run the Fragility Failsafe just in case Google Sheets glitched out
  ensureImportRangesActive();

  const otSheet = ss.getSheetByName('OTbyID');
  if (!otSheet) return { error: "Missing OTbyID tab" };

  // 2. Fetch the Playoff Teams from local sheet (since you sync them from NHL Standings for Draft Prep)
  const lsSheet = ss.getSheetByName('League Scoring');
  let playoffTeams = [];
  let masterScorers = [];
  let playerDict = {}; // Fast dictionary for instantaneous loop mapping

  if (lsSheet) {
    playoffTeams = lsSheet.getRange("M1:N19").getValues()
      .flat()
      .map(t => String(t).trim().toUpperCase())
      .filter(t => t && t.length === 3);

    // FIX: Pre-populate playerDict and masterScorers with Master identities
    const masterData = lsSheet.getRange(4, 1, Math.max(1, lsSheet.getLastRow() - 3), 3).getValues();
    masterData.forEach(row => {
      const id = String(row[0]).trim();
      if (id && id !== "0") {
        const pObj = {
          id: id,
          name: row[1],
          team: String(row[2]).trim().toUpperCase(),
          points: 0,
          delta: 0
        };
        playerDict[id] = pObj;
        masterScorers.push(pObj);
      }
    });
  }

  // 2.5 Fetch the actual points organically from EventStore! No local formulas needed for H:L!
  try {
    const eventStoreSS = SpreadsheetApp.openById('1RGouXlBsClZEfnSAx6uY8iWNmtsL5DkupDruKwQHjIc');
    const evtSheet = eventStoreSS.getSheetByName(getActiveScoringSheetName());
    if (evtSheet) {
      const evtLastRow = evtSheet.getLastRow();
      if (evtLastRow > 1) {
        // Reads: [PlayerID, Name, Team, GP, G, A, TP, Spacer, R1, R2, R3, R4, Yesterday]
        const pData = evtSheet.getRange(2, 1, evtLastRow - 1, 13).getValues();
        pData.forEach(row => {
          const id = String(row[0]).trim();
          if (id && playerDict[id]) {
            playerDict[id].points = Number(row[6]) || 0;
            playerDict[id].delta = playerDict[id].points - (Number(row[12]) || 0);
            // masterScorers already contains the reference to this object
          } else if (id) {
            // Player in EventStore but not in our master list? Add them.
            const pObj = {
              id: id,
              name: row[1],
              team: String(row[2]).trim().toUpperCase(),
              points: Number(row[6]) || 0,
              delta: (Number(row[6]) || 0) - (Number(row[12]) || 0)
            };
            playerDict[id] = pObj;
            masterScorers.push(pObj);
          }
        });
      }
    }
  } catch (e) {
    return { error: "Could not fetch EventStore. Did you clone the project properly? " + e.message };
  }

  // 3. Option B Algorithm: Automatically calculate EVERY owner's Unique player mathematically!
  const allUniques = [];
  const ownerHeaders = otSheet.getRange(3, 2, 1, 15).getValues()[0];
  const allLegacyRosters = otSheet.getRange(4, 2, 17, 15).getValues();

  for (let c = 0; c < ownerHeaders.length; c++) {
    if (String(ownerHeaders[c]).trim() !== "") {
      // Scan strictly down this specific owner's legacy roster
      for (let r = 0; r < 17; r++) {
        const pid = String(allLegacyRosters[r][c]).trim();
        if (pid && playerDict[pid]) {
          if (playoffTeams.includes(playerDict[pid].team)) {
            allUniques.push(pid); // Found the exact absolute playoff-bound leader! Locking them globally as Unique.
            break; // Immediately kill the loop for this owner to preserve their single Unique.
          }
        }
      }
    }
  }

  // 4. Identify the logged-in Owner's specific Column map
  const colIndex = ownerHeaders.indexOf(name);
  if (colIndex === -1) {
    return { error: `Owner ${name} not found in OTbyID row 3.` };
  }

  const activeCol = colIndex + 2;

  // Strictly cast all IDs from Google Sheets native Numbers into Strings to guarantee frontend .includes() javascript matching!
  const legacyIDs = otSheet.getRange(4, activeCol, 17).getValues().flat().map(id => String(id).trim()).filter(id => id);
  const benchIDs = otSheet.getRange(21, activeCol, 9).getValues().flat().map(id => String(id).trim()).filter(id => id);

  const legacySaved = otSheet.getRange(32, activeCol, 15).getValues().flat().map(id => String(id).trim()).filter(id => id);
  const draftSaved = otSheet.getRange(48, activeCol, 15).getValues().flat().map(id => String(id).trim()).filter(id => id);

  // 4.5 Global Pick Frequencies
  const allKeepersRaw = otSheet.getRange(32, 2, 15, 15).getValues().flat();
  const allDraftsRaw = otSheet.getRange(48, 2, 15, 15).getValues().flat();

  const pickFreqs = {};
  [...allKeepersRaw, ...allDraftsRaw].forEach(val => {
    const id = String(val).trim();
    if (id) {
      pickFreqs[id] = (pickFreqs[id] || 0) + 1;
    }
  });

  // 5. Send the strictly verified JSON payload organically to the browser
  return {
    legacyIds: legacyIDs,
    benchIds: benchIDs,
    legacySaved: legacySaved,
    draftSaved: draftSaved,
    pickFreqs: pickFreqs,
    playoffTeams: playoffTeams,
    masterScorers: masterScorers,
    allUniques: allUniques,
    ownerName: name
  };
}

/* =========================================================================
 * --- SYSTEMIC DATABASE MUTATION API ---
 * Persists the owner's final Roster & Draft selections securely natively into OTbyID.
 * Expected payload map: { owner: "Gerry", keepers: ["123", "456"], drafts: ["789", "012"] }
 * ========================================================================= */
function registerData(payload) {
  if (!payload || !payload.owner) return { success: false, error: "Missing owner payload validation." };

  const checklock = LockService.getScriptLock();
  checklock.waitLock(10000); // 10 second absolute structural grid lock

  try {
    const otSheet = ss.getSheetByName('OTbyID');
    if (!otSheet) return { success: false, error: "Missing OTbyID structural map." };

    // Scan headers to precisely locate user completely dynamically
    const ownerHeaders = otSheet.getRange(3, 2, 1, 15).getValues()[0];
    const colIndex = ownerHeaders.indexOf(payload.owner);
    if (colIndex === -1) return { success: false, error: "Unauthorized owner identity rejection." };

    const activeCol = colIndex + 2;

    // Final targets
    const finalKeepers = payload.keepers.slice(0, 15).filter(id => id && id.length > 0);
    const finalDrafts = payload.drafts.slice(0, 15).filter(id => id && id.length > 0);

    // Natively wipe explicitly Rows 32-46 (Keepers) and 48-62 (Drafts) cleanly!
    otSheet.getRange(32, activeCol, 15, 1).clearContent();
    otSheet.getRange(48, activeCol, 15, 1).clearContent();

    // Execute structural bulk write operation natively solidly
    if (finalKeepers.length > 0) {
      const kData = finalKeepers.map(id => [id]); // Convert to strict 2D column natively
      otSheet.getRange(32, activeCol, kData.length, 1).setValues(kData);
    }
    if (finalDrafts.length > 0) {
      const dData = finalDrafts.map(id => [id]); // Convert to strict 2D flexibly
      otSheet.getRange(48, activeCol, dData.length, 1).setValues(dData);
    }

    SpreadsheetApp.flush();
    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    checklock.releaseLock();
  }
}

/* =========================================================================
 * --- PLAYOFF STANDINGS API ---
 * Calculates Top 25 scoring rules natively and supports Live Matrix Views.
 * ========================================================================= */
function getEliminatedTeams() {
  const cache = CacheService.getScriptCache();
  const cachedDead = cache.get("deadTeams");
  if (cachedDead) {
    return JSON.parse(cachedDead);
  }

  try {
    const today = new Date();
    const year = today.getFullYear();
    let seasonStr = today.getMonth() < 7 ? (year - 1) + "" + year : year + "" + (year + 1);

    let url = "https://api-web.nhle.com/v1/playoff-bracket/" + seasonStr;
    let resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      url = "https://api-web.nhle.com/v1/playoff-bracket/20232024";
      resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    }

    if (resp.getResponseCode() === 200) {
      const json = JSON.parse(resp.getContentText());
      const deadTeams = [];
      (json.series || []).forEach(s => {
        const top = s.topSeed || {};
        const bot = s.bottomSeed || {};
        if (top.wins === 4 && bot.abbrev) deadTeams.push(String(bot.abbrev).toUpperCase());
        if (bot.wins === 4 && top.abbrev) deadTeams.push(String(top.abbrev).toUpperCase());
      });
      
      cache.put("deadTeams", JSON.stringify(deadTeams), 3600); // Cache for 1 hour
      return deadTeams;
    }
  } catch (e) { }
  return [];
}

function getStandingsData(name) {
  ensureImportRangesActive();

  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("standingsPayload");
  if (cachedData) {
    const parsed = JSON.parse(cachedData);
    return { viewer: name, leaderboard: parsed.leaderboard, yAxisActives: parsed.yAxisActives };
  }

  const otSheet = ss.getSheetByName('OTbyID');
  if (!otSheet) return { error: "Missing sheets" };

  let pDict = {};

  // FIX: Pre-populate pDict with Master identities from local League Scoring (A4:C)
  // This ensures players not yet in EventStore scoring tabs don't show up as "Unknown"
  const lsSheet = ss.getSheetByName('League Scoring');
  if (lsSheet) {
    const masterData = lsSheet.getRange(4, 1, Math.max(1, lsSheet.getLastRow() - 3), 3).getValues();
    masterData.forEach(row => {
      const id = String(row[0]).trim();
      if (id && id !== "0") {
        pDict[id] = {
          id: id,
          name: row[1],
          team: String(row[2]).trim().toUpperCase(),
          points: 0,
          r1: 0, r2: 0, r3: 0, r4: 0, yesterday: 0
        };
      }
    });
  }
  try {
    const eventStoreSS = SpreadsheetApp.openById('1RGouXlBsClZEfnSAx6uY8iWNmtsL5DkupDruKwQHjIc');
    const evtSheet = eventStoreSS.getSheetByName(getActiveScoringSheetName());
    if (evtSheet) {
      const evtLastRow = evtSheet.getLastRow();
      if (evtLastRow > 1) {
        // [PlayerID[0], Name[1], Team[2], GP[3], G[4], A[5], TP[6], Spacer[7], R1[8], R2[9], R3[10], R4[11], Yesterday[12]]
        evtSheet.getRange(2, 1, evtLastRow - 1, 13).getValues().forEach(row => {
          const id = String(row[0]).trim();
          if (id) {
            pDict[id] = {
              id, name: row[1], team: String(row[2]).trim().toUpperCase(),
              points: Number(row[6]) || 0,
              r1: Number(row[8]) || 0, r2: Number(row[9]) || 0, r3: Number(row[10]) || 0, r4: Number(row[11]) || 0,
              yesterday: Number(row[12]) || 0
            };
          }
        });
      }
    }
  } catch (e) {
    return { error: "Could not fetch EventStore. " + e.message };
  }

  const headers = otSheet.getRange(3, 2, 1, 15).getValues()[0];
  const ownerConfigs = [];
  const ownershipCounts = {};

  for (let c = 0; c < headers.length; c++) {
    const owner = String(headers[c]).trim();
    if (!owner) continue;

    const col = c + 2;
    const k = otSheet.getRange(32, col, 15).getValues().flat().filter(id => String(id).trim());
    const d = otSheet.getRange(48, col, 15).getValues().flat().filter(id => String(id).trim());
    const ids = k.concat(d);

    ids.forEach(id => { ownershipCounts[id] = (ownershipCounts[id] || 0) + 1; });
    ownerConfigs.push({ ownerName: owner, ids: ids });
  }

  const deadTeams = getEliminatedTeams();
  const standings = [];
  const globalActivesMap = {};

  ownerConfigs.forEach(conf => {
    let roster = conf.ids.map(id => pDict[id] || { id, name: "Unknown", team: "???", points: 0, r1: 0, r2: 0, r3: 0, r4: 0, yesterday: 0 });
    roster.forEach(p => {
      p.isDead = deadTeams.includes(p.team);
      p.rarity = ownershipCounts[p.id] || 1;
    });

    roster.sort((a, b) => b.points - a.points);

    let oTot = 0, oR1 = 0, oR2 = 0, oR3 = 0, oR4 = 0, oYest = 0;
    let bankedPoints = 0;
    let activeRosterDict = {};

    roster.forEach((p, idx) => {
      p.isCnt = (idx < 25);
      p.rnk = idx + 1;

      if (p.isCnt) {
        oTot += p.points;
        oR1 += p.r1; oR2 += p.r2; oR3 += p.r3; oR4 += p.r4;
        oYest += (p.yesterday || 0);
      }

      if (p.isDead && p.isCnt) {
        bankedPoints += p.points;
      } else if (!p.isDead) { // Actives
        activeRosterDict[p.id] = p;
        if (!globalActivesMap[p.id]) globalActivesMap[p.id] = p;
      }
    });

    standings.push({
      ownerName: conf.ownerName,
      totalPoints: oTot,
      r1: oR1, r2: oR2, r3: oR3, r4: oR4,
      delta: oTot - oYest, // calculate explicit overnight delta naturally
      bankedPoints: bankedPoints,
      actives: activeRosterDict
    });
  });

  standings.sort((a, b) => b.totalPoints - a.totalPoints);

  const globalActivesList = Object.values(globalActivesMap);
  globalActivesList.sort((a, b) => {
    return b.points - a.points;
  });

  // Cache the highly complex matrix generation for 60 seconds
  cache.put("standingsPayload", JSON.stringify({ leaderboard: standings, yAxisActives: globalActivesList }), 60);

  return { viewer: name, leaderboard: standings, yAxisActives: globalActivesList };
}

/* =========================================================================
 * --- SYSTEMIC AUTO-GENERATOR ---
 * Automatically drafts statistically optimal teams for empty owners.
 * ========================================================================= */
function autoGenerateMissingTeams() {
  ensureImportRangesActive();

  const otSheet = ss.getSheetByName('OTbyID');
  const lsSheet = ss.getSheetByName('League Scoring');
  if (!otSheet || !lsSheet) {
    try { SpreadsheetApp.getUi().alert("Missing Sheets: OTbyID or League Scoring."); } catch (e) { }
    return;
  }

  // 1. Load Master Dictionaries
  const playoffTeams = lsSheet.getRange("I4:M35").getValues().flat().map(t => String(t).trim().toUpperCase()).filter(t => t.length === 3);
  const lsLastRow = lsSheet.getLastRow();
  const masterScorers = [];
  const playerDict = {};

  if (lsLastRow >= 4) {
    lsSheet.getRange(4, 1, lsLastRow - 3, 7).getValues().forEach(row => {
      const p = { id: String(row[0]).trim(), team: String(row[2]).trim().toUpperCase(), points: Number(row[6]) || 0 };
      if (p.id) { playerDict[p.id] = p; masterScorers.push(p); }
    });
  }

  // 2. Map Uniques
  const ownerHeaders = otSheet.getRange(3, 2, 1, 15).getValues()[0];
  const allLegacyRosters = otSheet.getRange(4, 2, 17, 15).getValues();
  const allUniques = [];

  for (let c = 0; c < ownerHeaders.length; c++) {
    if (String(ownerHeaders[c]).trim()) {
      for (let r = 0; r < 17; r++) {
        const pid = String(allLegacyRosters[r][c]).trim();
        if (pid && playerDict[pid] && playoffTeams.includes(playerDict[pid].team)) {
          allUniques.push(pid); break; // Global Lock
        }
      }
    }
  }

  // 3. Scan & Generate for Empties
  let generatedCount = 0;
  for (let c = 0; c < ownerHeaders.length; c++) {
    if (!String(ownerHeaders[c]).trim()) continue;

    const activeCol = c + 2;
    const legacySaved = otSheet.getRange(32, activeCol, 15).getValues().flat().filter(id => String(id).trim() !== "");
    const draftSaved = otSheet.getRange(48, activeCol, 15).getValues().flat().filter(id => String(id).trim() !== "");

    // ONLY generate if they completely haven't saved anything
    if (legacySaved.length === 0 && draftSaved.length === 0) {

      const legacyIDs = otSheet.getRange(4, activeCol, 17).getValues().flat().map(id => String(id).trim()).filter(id => id);
      const benchIDs = otSheet.getRange(21, activeCol, 9).getValues().flat().map(id => String(id).trim()).filter(id => id);
      const combinedIds = legacyIDs.concat(benchIDs);

      // Auto Generate Keepers (Playoff Bound & Sort By Points)
      const autoKeepers = legacyIDs.map(id => playerDict[id]).filter(p => p && playoffTeams.includes(p.team));
      autoKeepers.sort((a, b) => b.points - a.points);
      const finalKeepers = autoKeepers.slice(0, 15).map(p => [p.id]);

      // Auto Generate Drafts (Excluding Uniques & Own History)
      const draftPool = masterScorers.filter(p => !allUniques.includes(p.id) && !combinedIds.includes(p.id) && playoffTeams.includes(p.team));
      draftPool.sort((a, b) => b.points - a.points);
      const finalDrafts = draftPool.slice(0, 15).map(p => [p.id]);

      // Write Data Logically
      if (finalKeepers.length > 0) otSheet.getRange(32, activeCol, finalKeepers.length, 1).setValues(finalKeepers);
      if (finalDrafts.length > 0) otSheet.getRange(48, activeCol, finalDrafts.length, 1).setValues(finalDrafts);

      generatedCount++;
    }
  }

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(`Auto-Draft Complete! Automatically generated optimal playoff teams for ${generatedCount} owners.`);
  } catch (e) { } // In case ran via trigger without UI
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DaWiz Scripts')
    .addItem('Run Auto-Generator', 'autoGenerateMissingTeams')
    .addToUi();
}
