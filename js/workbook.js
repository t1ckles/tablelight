(function (root) {
  const XLSX = root.XLSX;
  const TL = root.Tablelight = root.Tablelight || {};

  const RESERVED = new Set(["ReadMe", "_Campaign", "_Index", "_Lore", "_Party", "_SessionLog"]);
  const FIELD_ORDER = [
    "Name",
    "Race",
    "RaceCustom",
    "Class",
    "ClassCustom",
    "Role",
    "RoleCustom",
    "ClassRole",
    "Alignment",
    "Concept",
    "History",
    "Location",
    "Status",
    "Conditions",
    "Attitude",
    "AttitudeValue",
    "AC",
    "HP",
    "HPMax",
    "STR",
    "DEX",
    "CON",
    "INT",
    "WIS",
    "CHA",
    "KeyStats",
    "Voice",
    "Motivation",
  ];

  const KEY_MAP = {
    name: "Name",
    race: "Race",
    species: "Race",
    racecustom: "RaceCustom",
    "race custom": "RaceCustom",
    class: "Class",
    classcustom: "ClassCustom",
    "class custom": "ClassCustom",
    classrole: "ClassRole",
    "class/role": "ClassRole",
    role: "Role",
    rolecustom: "RoleCustom",
    "role custom": "RoleCustom",
    background: "Role",
    alignment: "Alignment",
    history: "History",
    conditions: "Conditions",
    str: "STR",
    strength: "STR",
    dex: "DEX",
    dexterity: "DEX",
    con: "CON",
    constitution: "CON",
    int: "INT",
    intelligence: "INT",
    wis: "WIS",
    wisdom: "WIS",
    cha: "CHA",
    charisma: "CHA",
    concept: "Concept",
    location: "Location",
    status: "Status",
    "health/condition": "Status",
    health: "Status",
    condition: "Status",
    attitude: "Attitude",
    attitudevalue: "AttitudeValue",
    "attitude value": "AttitudeValue",
    ac: "AC",
    armorclass: "AC",
    "armor class": "AC",
    hp: "HP",
    hitpoints: "HP",
    "hit points": "HP",
    hpmax: "HPMax",
    "hp max": "HPMax",
    maxhp: "HPMax",
    "max hp": "HPMax",
    keystats: "KeyStats",
    "key stats": "KeyStats",
    stats: "KeyStats",
    voice: "Voice",
    "voice/tone": "Voice",
    motivation: "Motivation",
    "key motivation": "Motivation",
  };

  function str(value) {
    if (value === undefined || value === null) return "";
    return String(value);
  }

  function sheetFrom(aoa, widths) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (widths) ws["!cols"] = widths.map((wch) => ({ wch }));
    return ws;
  }

  function replaceSheet(wb, name, ws) {
    if (wb.Sheets[name]) wb.Sheets[name] = ws;
    else XLSX.utils.book_append_sheet(wb, ws, name);
  }

  function rowsOf(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  }

  function canonKey(label) {
    const raw = str(label).trim();
    if (!raw) return "";
    const mapped = KEY_MAP[raw.toLowerCase()];
    return mapped || raw;
  }

  function toSheetName(name) {
    let s = str(name)
      .trim()
      .replace(/[\\/?*[\]:]/g, "")
      .replace(/[^A-Za-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    if (!s) s = "NPC";
    return s.slice(0, 31);
  }

  function uniqueSheetName(wb, base) {
    let name = toSheetName(base);
    if (RESERVED.has(name) || name.startsWith("_")) name = "NPC_" + name.replace(/^_+/, "");
    name = name.slice(0, 31);
    if (!wb.Sheets[name]) return name;
    let n = 2;
    while (true) {
      const suffix = "_" + n;
      const candidate = (name.slice(0, 31 - suffix.length) + suffix).slice(0, 31);
      if (!wb.Sheets[candidate]) return candidate;
      n += 1;
    }
  }

  function isNpcSheet(name) {
    return name && !RESERVED.has(name) && !String(name).startsWith("_");
  }

  function defaultCampaign() {
    return {
      CampaignName: "Untitled campaign",
      Ruleset: "D&D 2024",
      PartyLevel: "1",
      PartySize: "4",
      SessionNumber: "1",
      InGameDate: "",
      InGameTime: "Morning",
      Boundaries: "Heroic fantasy. No sexual content. Keep violence direct but not graphic.",
      HouseRules: "",
    };
  }

  function blankFields(partial) {
    const fields = {};
    FIELD_ORDER.forEach((key) => {
      fields[key] = "";
    });
    fields.Attitude = "Neutral";
    fields.AttitudeValue = "0";
    Object.keys(partial || {}).forEach((key) => {
      const canon = FIELD_ORDER.includes(key) ? key : canonKey(key);
      if (!canon) return;
      fields[canon] = str(partial[key]);
    });
    if (root.Tablelight.rules) root.Tablelight.rules.prepareFields(fields);
    return fields;
  }

  function readCampaign(wb) {
    const campaign = defaultCampaign();
    rowsOf(wb, "_Campaign").forEach((row) => {
      const key = str(row[0]).trim();
      if (!key || key === "Key") return;
      if (Object.prototype.hasOwnProperty.call(campaign, key)) campaign[key] = str(row[1]);
    });
    campaign.Ruleset = "D&D 2024";
    return campaign;
  }

  function readLore(wb) {
    const rows = rowsOf(wb, "_Lore");
    const lore = [];
    rows.slice(1).forEach((row) => {
      if (!str(row[0]) && !str(row[1]) && !str(row[3])) return;
      lore.push({
        id: str(row[0]) || "L" + (lore.length + 1),
        title: str(row[1]),
        tags: str(row[2]),
        body: str(row[3]),
        updated: str(row[4]),
      });
    });
    return lore;
  }

  function readParty(wb) {
    const rows = rowsOf(wb, "_Party");
    const party = [];
    rows.slice(1).forEach((row) => {
      if (!str(row[0]) && !str(row[1])) return;
      party.push({
        name: str(row[0]),
        className: str(row[1]),
        level: str(row[2]),
        ac: str(row[3]),
        hp: str(row[4]),
        notes: str(row[5]),
      });
    });
    return party;
  }

  function readLog(wb) {
    const rows = rowsOf(wb, "_SessionLog");
    const log = [];
    rows.slice(1).forEach((row) => {
      if (!str(row[0]) && !str(row[5])) return;
      log.push({
        timestamp: str(row[0]),
        session: str(row[1]),
        inGameDate: str(row[2]),
        agent: str(row[3]),
        query: str(row[4]),
        response: str(row[5]),
      });
    });
    return log;
  }

  function readNpcSheet(wb, name) {
    const rows = rowsOf(wb, name);
    const fields = blankFields();
    const extras = [];
    const interactions = [];
    const equipment = [];
    let mode = "fields";
    rows.forEach((row) => {
      const a = str(row[0]).trim();
      const b = str(row[1]);
      if (mode === "fields") {
        if (!a) return;
        if (a.toUpperCase() === "EQUIPMENT") {
          mode = "equip-header";
          return;
        }
        if (a.toUpperCase() === "INTERACTIONS") {
          mode = "header";
          return;
        }
        const key = canonKey(a);
        if (FIELD_ORDER.includes(key)) fields[key] = b;
        else extras.push([a, b]);
        return;
      }
      if (mode === "equip-header") {
        mode = "equipment";
        return;
      }
      if (mode === "equipment") {
        if (a.toUpperCase() === "INTERACTIONS") {
          mode = "header";
          return;
        }
        if (!a && !b) return;
        equipment.push({ item: a, qty: str(row[1]) || "1", notes: str(row[2]) });
        return;
      }
      if (mode === "header") {
        mode = "interactions";
        return;
      }
      if (!a && !str(row[1]) && !str(row[2]) && !str(row[3])) return;
      interactions.push({
        date: str(row[0]),
        session: str(row[1]),
        type: str(row[2]) || "Note",
        entry: str(row[3]),
      });
    });
    if (!fields.Name) fields.Name = name.replace(/_/g, " ");
    if (root.Tablelight.rules) root.Tablelight.rules.prepareFields(fields);
    return { sheet: name, fields, extras, interactions, equipment };
  }

  function readNpcs(wb) {
    return wb.SheetNames.filter(isNpcSheet).map((name) => readNpcSheet(wb, name));
  }

  function npcAoa(npc) {
    const aoa = FIELD_ORDER.map((key) => [key, str(npc.fields[key])]);
    (npc.extras || []).forEach((pair) => aoa.push([pair[0], str(pair[1])]));
    aoa.push([]);
    aoa.push(["EQUIPMENT"]);
    aoa.push(["Item", "Qty", "Notes"]);
    (npc.equipment || []).forEach((item) => {
      aoa.push([str(item.item), str(item.qty || "1"), str(item.notes)]);
    });
    aoa.push([]);
    aoa.push(["INTERACTIONS"]);
    aoa.push(["Date", "Session", "Type", "Entry"]);
    (npc.interactions || []).forEach((item) => {
      aoa.push([str(item.date), str(item.session), str(item.type), str(item.entry)]);
    });
    return aoa;
  }

  function writeNpcSheet(wb, npc) {
    const ws = sheetFrom(npcAoa(npc), [18, 56, 16, 80]);
    if (wb.Sheets[npc.sheet]) {
      wb.Sheets[npc.sheet] = ws;
      return;
    }
    const indexAt = wb.SheetNames.indexOf("_Index");
    XLSX.utils.book_append_sheet(wb, ws, npc.sheet);
    if (indexAt >= 0) {
      const last = wb.SheetNames.pop();
      wb.SheetNames.splice(indexAt, 0, last);
    }
  }

  function rebuildIndex(wb) {
    const aoa = [["Sheet", "Name", "Race", "Class", "Role", "Alignment", "Location", "Status", "AC", "HP", "HPMax"]];
    readNpcs(wb).forEach((npc) => {
      const f = npc.fields;
      const rules = root.Tablelight.rules;
      const species = rules ? rules.shown(f.Race, f.RaceCustom) : f.Race;
      const role = rules ? rules.shown(f.Role, f.RoleCustom) : f.Role;
      aoa.push([npc.sheet, f.Name, species, rules ? rules.shown(f.Class, f.ClassCustom) : f.Class, role, f.Alignment, f.Location, f.Status, f.AC, f.HP, f.HPMax]);
    });
    replaceSheet(wb, "_Index", sheetFrom(aoa, [28, 24, 16, 22, 28, 16, 14, 16, 8, 8, 8]));
  }

  function saveCampaign(wb, campaign) {
    const base = defaultCampaign();
    const merged = Object.assign(base, campaign || {});
    const aoa = [["Key", "Value"]];
    Object.keys(base).forEach((key) => aoa.push([key, str(merged[key])]));
    replaceSheet(wb, "_Campaign", sheetFrom(aoa, [20, 88]));
    if (wb.Props) wb.Props.Title = merged.CampaignName || "Tablelight campaign";
  }

  function saveLore(wb, lore) {
    const aoa = [["Id", "Title", "Tags", "Body", "Updated"]];
    (lore || []).forEach((entry, index) => {
      aoa.push([
        str(entry.id) || "L" + (index + 1),
        str(entry.title),
        str(entry.tags),
        str(entry.body),
        str(entry.updated),
      ]);
    });
    replaceSheet(wb, "_Lore", sheetFrom(aoa, [10, 28, 24, 88, 22]));
  }

  function saveParty(wb, party) {
    const aoa = [["Name", "Class", "Level", "AC", "HP", "Notes"]];
    (party || []).forEach((member) => {
      aoa.push([str(member.name), str(member.className), str(member.level), str(member.ac), str(member.hp), str(member.notes)]);
    });
    replaceSheet(wb, "_Party", sheetFrom(aoa, [22, 18, 10, 8, 8, 40]));
  }

  function saveLog(wb, log) {
    const aoa = [["Timestamp", "Session", "InGameDate", "Agent", "Query", "Response"]];
    (log || []).forEach((entry) => {
      aoa.push([str(entry.timestamp), str(entry.session), str(entry.inGameDate), str(entry.agent), str(entry.query), str(entry.response)]);
    });
    replaceSheet(wb, "_SessionLog", sheetFrom(aoa, [22, 12, 18, 24, 40, 80]));
  }

  function readMeSheet() {
    const lines = [
      "Tablelight campaign workbook",
      "",
      "Each NPC has one sheet tab. The tab name is the NPC name with spaces and punctuation turned into underscores, such as Mara_Quill.",
      "NPC sheets are a key-value list in columns A and B, then an INTERACTIONS table.",
      "Do not rename the sheets ReadMe, _Campaign, _Index, _Lore, _Party, or _SessionLog.",
      "_Index is rebuilt by Tablelight so the assistant can find NPCs quickly.",
      "World notes live on _Lore. The party roster lives on _Party. Session setup lives on _Campaign.",
      "You can edit this file in Excel while Tablelight is closed, then open it again from the site.",
    ];
    return sheetFrom(lines.map((line) => [line]), [110]);
  }

  function ensureSystem(wb) {
    if (!wb.Sheets.ReadMe) replaceSheet(wb, "ReadMe", readMeSheet());
    if (!wb.Sheets._Campaign) saveCampaign(wb, defaultCampaign());
    if (!wb.Sheets._Lore) saveLore(wb, []);
    if (!wb.Sheets._Party) saveParty(wb, []);
    if (!wb.Sheets._SessionLog) saveLog(wb, []);
    if (!wb.Sheets._Index) rebuildIndex(wb);
    if (!wb.Props) wb.Props = {};
  }

  function createWorkbook(campaign, lore, party, npcs) {
    const wb = XLSX.utils.book_new();
    wb.Props = { Title: (campaign && campaign.CampaignName) || "Tablelight campaign" };
    replaceSheet(wb, "ReadMe", readMeSheet());
    saveCampaign(wb, campaign || defaultCampaign());
    saveLore(wb, lore || []);
    saveParty(wb, party || []);
    saveLog(wb, []);
    (npcs || []).forEach((npc) => writeNpcSheet(wb, npc));
    rebuildIndex(wb);
    return wb;
  }

  function sampleNpc(sheet, fields, interactions, equipment) {
    return {
      sheet,
      fields: blankFields(fields),
      extras: [],
      interactions: interactions || [],
      equipment: equipment || [],
    };
  }

  function createSample() {
    const data = root.Tablelight.drownedBell();
    const npcs = data.npcs.map((npc) => sampleNpc(npc.sheet, npc.fields, npc.interactions, npc.equipment));
    return createWorkbook(data.campaign, data.lore, data.party, npcs);
  }

  function createBlank() {
    return createWorkbook(defaultCampaign(), [], [], []);
  }

  function read(wb) {
    ensureSystem(wb);
    return {
      campaign: readCampaign(wb),
      lore: readLore(wb),
      party: readParty(wb),
      npcs: readNpcs(wb),
      log: readLog(wb),
    };
  }

  function createNpc(wb, fields) {
    const seeded = blankFields(fields);
    const sheet = uniqueSheetName(wb, seeded.Name || "NPC");
    const npc = { sheet, fields: seeded, extras: [], interactions: [], equipment: [] };
    if (!npc.fields.Name) npc.fields.Name = sheet.replace(/_/g, " ");
    writeNpcSheet(wb, npc);
    rebuildIndex(wb);
    return npc;
  }

  function writeNpc(wb, npc) {
    if (!npc || !wb.Sheets[npc.sheet]) throw new Error("That NPC sheet is not in the workbook.");
    writeNpcSheet(wb, npc);
    rebuildIndex(wb);
    return readNpcSheet(wb, npc.sheet);
  }

  function deleteNpc(wb, sheet) {
    const index = wb.SheetNames.indexOf(sheet);
    if (index < 0 || !isNpcSheet(sheet)) throw new Error("That sheet cannot be removed.");
    wb.SheetNames.splice(index, 1);
    delete wb.Sheets[sheet];
    rebuildIndex(wb);
  }

  function renameNpcSheet(wb, sheet, displayName) {
    const npc = readNpcSheet(wb, sheet);
    const next = uniqueSheetName(wb, displayName || npc.fields.Name);
    if (next === sheet) return npc;
    npc.sheet = next;
    writeNpcSheet(wb, npc);
    const index = wb.SheetNames.indexOf(sheet);
    wb.SheetNames.splice(index, 1);
    delete wb.Sheets[sheet];
    rebuildIndex(wb);
    return readNpcSheet(wb, next);
  }

  function appendLog(wb, entry) {
    const log = readLog(wb);
    log.push(entry);
    saveLog(wb, log);
    return log;
  }

  function nextLoreId(lore) {
    let max = 0;
    (lore || []).forEach((entry) => {
      const match = /^L(\d+)$/.exec(str(entry.id));
      if (match) max = Math.max(max, parseInt(match[1], 10));
    });
    return "L" + (max + 1);
  }

  function fromArrayBuffer(buffer) {
    const wb = XLSX.read(buffer, { type: "array" });
    ensureSystem(wb);
    return wb;
  }

  function toArrayBuffer(wb) {
    return XLSX.write(wb, { bookType: "xlsx", type: "array" });
  }

  TL.workbook = {
    FIELD_ORDER,
    RESERVED,
    toSheetName,
    defaultCampaign,
    blankFields,
    createBlank,
    createSample,
    read,
    saveCampaign,
    saveLore,
    saveParty,
    createNpc,
    writeNpc,
    deleteNpc,
    renameNpcSheet,
    appendLog,
    nextLoreId,
    fromArrayBuffer,
    toArrayBuffer,
    isNpcSheet,
  };
})(typeof window !== "undefined" ? window : globalThis);
