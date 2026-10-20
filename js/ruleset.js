(function (root) {
  const TL = root.Tablelight = root.Tablelight || {};

  const SPECIES = ["Aasimar", "Dragonborn", "Dwarf", "Elf", "Gnome", "Goliath", "Halfling", "Human", "Orc", "Tiefling", "Other"];
  const CLASSES = ["None", "Barbarian", "Bard", "Cleric", "Druid", "Fighter", "Monk", "Paladin", "Ranger", "Rogue", "Sorcerer", "Warlock", "Wizard", "Other"];
  const ROLES = ["None", "Acolyte", "Artisan", "Charlatan", "Criminal", "Entertainer", "Farmer", "Guard", "Guide", "Hermit", "Merchant", "Noble", "Sage", "Sailor", "Scribe", "Soldier", "Wayfarer", "Other"];
  const ALIGNMENTS = ["Lawful Good", "Neutral Good", "Chaotic Good", "Lawful Neutral", "Neutral", "Chaotic Neutral", "Lawful Evil", "Neutral Evil", "Chaotic Evil", "Unaligned"];
  const CONDITIONS = ["Blinded", "Charmed", "Deafened", "Exhaustion", "Frightened", "Grappled", "Incapacitated", "Invisible", "Paralyzed", "Petrified", "Poisoned", "Prone", "Restrained", "Stunned", "Unconscious"];
  const ABILITIES = ["STR", "DEX", "CON", "INT", "WIS", "CHA"];
  const TIMES = ["Dawn", "Morning", "Midday", "Afternoon", "Dusk", "Evening", "Night", "Midnight"];
  const WORLD_KINDS = [
    { id: "all", label: "All" },
    { id: "settlement", label: "Settlements", tags: ["settlement", "village", "city", "town"] },
    { id: "nation", label: "Nations", tags: ["nation"] },
    { id: "faction", label: "Factions", tags: ["faction"] },
    { id: "location", label: "Places", tags: ["location"] },
    { id: "quest", label: "Quests", tags: ["quest"] },
    { id: "monster", label: "Monsters", tags: ["monster"] },
    { id: "world", label: "Lore", tags: ["world", "lore", "item", "rules"] },
  ];

  const GIVEN = {
    Aasimar: ["Sera", "Ilan", "Maera", "Oren", "Haleth", "Nima"],
    Dragonborn: ["Bask", "Korv", "Ska", "Thrum", "Vexa", "Azr"],
    Dwarf: ["Bram", "Hild", "Torv", "Dura", "Kaz", "Nala"],
    Elf: ["Leth", "Syl", "Vael", "Maelis", "Ori", "Sira"],
    Gnome: ["Pim", "Lark", "Tilli", "Nob", "Zep", "Mira"],
    Goliath: ["Kav", "Ula", "Thom", "Ruk", "Sela", "Bor"],
    Halfling: ["Ness", "Hob", "Tilly", "Bramble", "Posy", "Wil"],
    Human: ["Edda", "Calder", "Maret", "Col", "Greta", "Owen", "Ida", "Harun", "Lisle", "Perrin"],
    Orc: ["Grath", "Mog", "Ska", "Urg", "Hesh", "Drel"],
    Tiefling: ["Vess", "Cinder", "Nyx", "Rook", "Ash", "Lior"],
    Other: ["Ari", "Ren", "Sol", "Keel", "Nova", "Quill"],
  };
  const SURNAMES = ["Voss", "Pye", "Fen", "Hale", "Holt", "Keel", "Marrow", "Quinn", "Ash", "Dune", "Brask", "Wynn", "Corbin", "Iven"];

  function pick(list) {
    return list[Math.floor(Math.random() * list.length)];
  }

  function randomName(species) {
    const given = GIVEN[species] || GIVEN.Human;
    return pick(given) + " " + pick(SURNAMES);
  }

  function abilityMod(score) {
    const n = parseInt(score, 10);
    if (Number.isNaN(n)) return "";
    const mod = Math.floor((n - 10) / 2);
    return (mod >= 0 ? "+" : "") + mod;
  }

  function healthOf(hp, max) {
    const current = parseInt(hp, 10);
    const total = parseInt(max, 10);
    if (Number.isNaN(current) || Number.isNaN(total) || total <= 0) return { label: "Unset", percent: "" };
    const percent = Math.max(0, Math.round((current / total) * 100));
    let label = "Critical";
    if (current <= 0) label = "Down";
    else if (percent >= 100) label = "Healthy";
    else if (percent >= 75) label = "Bruised";
    else if (percent >= 50) label = "Injured";
    else if (percent >= 25) label = "Bloodied";
    return { label, percent: String(percent) };
  }

  function shown(choice, custom) {
    if (!choice || choice === "None") return "";
    if (choice === "Other") return String(custom || "").trim();
    return choice;
  }

  function pullScore(text, ability) {
    const match = String(text || "").match(new RegExp("\\b" + ability + "\\s*(\\d+)", "i"));
    return match ? match[1] : "";
  }

  function prepareFields(fields) {
    ABILITIES.forEach((ability) => {
      if (!fields[ability]) {
        const found = pullScore(fields.KeyStats, ability);
        if (found) fields[ability] = found;
      }
    });
    const classUnset = !fields.Class || fields.Class === "None";
    const roleUnset = !fields.Role || fields.Role === "None";
    if (classUnset && roleUnset && fields.ClassRole) {
      if (CLASSES.indexOf(fields.ClassRole) >= 0) fields.Class = fields.ClassRole;
      else {
        fields.Class = "None";
        fields.Role = "Other";
        fields.RoleCustom = fields.ClassRole;
      }
    }
    if (!fields.Class) fields.Class = "None";
    if (!fields.Role) fields.Role = "None";
    if (!fields.Alignment) fields.Alignment = "Unaligned";
    if (fields.Race && SPECIES.indexOf(fields.Race) < 0) {
      fields.RaceCustom = fields.Race;
      fields.Race = "Other";
    }
    const condition = CONDITIONS.find((name) => name.toLowerCase() === String(fields.Status || "").toLowerCase());
    if (condition) {
      const have = String(fields.Conditions || "").split(",").map((part) => part.trim()).filter(Boolean);
      if (have.indexOf(condition) < 0) have.push(condition);
      fields.Conditions = have.join(", ");
    }
    fields.Status = healthOf(fields.HP, fields.HPMax).label;
    const cls = shown(fields.Class, fields.ClassCustom);
    const role = shown(fields.Role, fields.RoleCustom);
    const parts = [cls, role].filter(Boolean);
    if (parts.length) fields.ClassRole = parts.join(" · ");
    return fields;
  }

  TL.rules = {
    SPECIES,
    CLASSES,
    ROLES,
    ALIGNMENTS,
    CONDITIONS,
    ABILITIES,
    TIMES,
    WORLD_KINDS,
    randomName,
    abilityMod,
    healthOf,
    shown,
    prepareFields,
  };
})(typeof window !== "undefined" ? window : globalThis);
