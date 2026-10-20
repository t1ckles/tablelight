(function (root) {
  const TL = root.Tablelight = root.Tablelight || {};

  const AGENTS = {
    rules: {
      id: "rules",
      label: "Rules Specialist",
      temperature: 0.2,
    },
    narrative: {
      id: "narrative",
      label: "Narrative Generator",
      temperature: 0.7,
    },
    combat: {
      id: "combat",
      label: "Combat Builder",
      temperature: 0.4,
    },
    lore: {
      id: "lore",
      label: "World & Lore Manager",
      temperature: 0.2,
    },
    npc: {
      id: "npc",
      label: "NPC & State Manager",
      temperature: 0.2,
    },
  };

  const RULES_RE = /\b(rules?|spells?|saving throw|save dc|advantage|disadvantage|conditions?|grappled|grapple|concentration|bonus action|reaction|cover|resistance|immunity|attack roll|skill check|ability check|how does|mechanics|action economy|opportunity attack)\b/i;
  const NARRATIVE_RE = /\b(describe|description|read-?aloud|dialogue|what would .+ say|says|say|scene|atmosphere|sensory|narrat|look like|looks like|voice)\b/i;
  const COMBAT_RE = /\b(encounters?|initiative|tactics|combat|monsters?|balance|deadly|hard encounter|terrain|morale|battlefield)\b/i;
  const LORE_RE = /\b(lore|history|histories|factions?|region|legend|rumou?r|agenda|goals?|who rules|world)\b/i;
  const WRITE_RE = /\b((?:took|takes?|taking|suffers?|suffered|loses?|lost|hit for)\s+(?:\d+|all)|dealt\s+\d+|\d+\s+(?:points?\s+of\s+)?(?:[a-z]+\s+)?damage|(?:heal(?:s|ed|ing)?|recover(?:s|ed|ing)?|restor(?:e|es|ed|ing)|regains?|regained)(?:\s+\d+|.{0,40}\b(?:all|remaining|the rest)\b)|fully healed|full heal|back to full|set\s+hp|hp\s+(?:to|is|now)|moved to|now (?:at|in)|location\s+is|gave\b|given\b|received\b|died\b|killed\b|befriended)\b/i;
const ATTITUDE_WRITE_RE = /\battitude\s+(?:is|now|to|toward|towards|:)/i;
const LOG_WRITE_RE = /^(?:update|record|log)\b/i;
  const STOP = new Set(["what", "does", "this", "that", "with", "from", "your", "about", "where", "when", "who", "have", "they", "them", "their", "would", "could", "should", "into", "onto", "the", "and", "for", "are", "was", "were", "how", "why", "can", "you", "please", "main", "active"]);

  function normalize(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[_'’.-]/g, " ")
      .replace(/[^a-z0-9\s]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

function matchNpcs(npcs, text, options) {
  const query = normalize(text);
  const hits = [];
  const seen = new Set();
  function add(npc) {
    if (!npc || seen.has(npc.sheet)) return;
    seen.add(npc.sheet);
    hits.push(npc);
  }
  (npcs || []).forEach((npc) => {
    const name = normalize(npc.fields.Name);
    const sheet = normalize(npc.sheet).replace(/_/g, " ");
    if (name && query.includes(name)) add(npc);
    else if (sheet && query.includes(sheet)) add(npc);
  });
  (npcs || []).forEach((npc) => {
    const parts = normalize(npc.fields.Name).split(" ").filter((part) => part.length > 3);
    parts.forEach((part) => {
      if (!new RegExp("\\b" + part + "\\b").test(query)) return;
      const owners = (npcs || []).filter((other) => normalize(other.fields.Name).split(" ").includes(part));
      if (owners.length === 1) add(owners[0]);
    });
  });
  if (options && options.includeLocation) {
    (npcs || []).forEach((npc) => {
      const location = String(npc.fields.Location || "").trim();
      if (location.length < 8) return;
      if (text.toLowerCase().includes(location.toLowerCase())) add(npc);
    });
  }
  return hits;
}

  function searchLore(lore, text) {
    const words = normalize(text).split(" ").filter((word) => word.length > 3 && !STOP.has(word));
    if (!words.length) return [];
    const ranked = (lore || [])
      .map((entry) => {
        const hay = normalize(entry.title + " " + entry.tags + " " + entry.body);
        const score = words.reduce((sum, word) => sum + (hay.includes(word) ? 1 : 0), 0);
        return { entry, score };
      })
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score);
    if (!ranked.length) return [];
    const top = ranked[0].score;
    const chosen = top >= 2 ? ranked.filter((item) => item.score === top) : ranked.slice(0, 3);
    return chosen.slice(0, 6).map((item) => item.entry);
  }

  function grabLine(body, label) {
    const match = String(body || "").match(new RegExp(label + "\\s*:\\s*([^\\n]+)", "i"));
    return match ? match[1].trim() : "";
  }

  function loreFacts(entry) {
    return {
      status: grabLine(entry.body, "Current status/location") || entry.title,
      relationship: grabLine(entry.body, "Relationship to party") || "Not recorded in the workbook",
      goals: grabLine(entry.body, "Active goals") || "Not recorded in the workbook",
    };
  }

  function parseNumber(value) {
    const match = String(value || "").match(/-?\d+/);
    return match ? parseInt(match[0], 10) : NaN;
  }

  function parseHealing(text) {
    return String(text || "").match(/\b(?:heal(?:s|ed|ing)?|recover(?:s|ed|ing)?|restor(?:e|es|ed|ing)|cure[sd]?|regains?|regained)\s+(\d+)/i);
  }

  function healsAll(text) {
    const source = String(text || "");
    if (parseHealing(source)) return false;
    return /\b(?:heal(?:s|ed|ing)?|recover(?:s|ed|ing)?|restor(?:e|es|ed|ing)|cure[sd]?|regains?|regained)\b.{0,40}\b(?:all|remaining|the rest)\b/i.test(source)
      || /\b(?:fully healed|full heal|back to full|at full health)\b/i.test(source);
  }

  function takesAll(text) {
    const source = String(text || "");
    if (parseHealing(source) || healsAll(source)) return false;
    return /\b(?:took|takes?|taking|suffers?|suffered|loses?|lost)\b.{0,40}\ball\b/i.test(source);
  }

  function parseDamage(text) {
    if (parseHealing(text) || healsAll(text)) return null;
    const verb = String(text || "").match(/\b(?:took|takes?|taking|suffers?|suffered|loses?|lost|hit for)\s+(\d+)\s+(?:points?\s+of\s+)?(?:[a-z]+\s+)?(?:damage|hit points|hp)\b(?:\s+from\s+([^.?!]+))?/i);
    if (verb) return verb;
    if (/\b(?:deal|deals|dealt|dealing)\b/i.test(text)) return null;
    return String(text || "").match(/\b(\d+)\s+(?:points?\s+of\s+)?(?:[a-z]+\s+)?damage\b(?:\s+from\s+([^.?!]+))?/i);
  }

  function proposeUpdates(text, npcs, campaign) {
    return (npcs || []).map((npc) => {
      const fields = {};
      let type = "Note";
      let entry = text.trim();
      const healing = parseHealing(text);
      const fullHeal = healsAll(text);
      const damage = healing || fullHeal ? null : parseDamage(text);
      const allDamage = takesAll(text);
      const hpSet = text.match(/\bhp\s+(?:to|is|now|:)\s*(\d+)/i);
      const acSet = text.match(/\bac\s+(?:to|is|now|:)\s*(\d+)/i);
      const moved = text.match(/(?:moved to|now (?:at|in)|location(?:\s+is|:))\s+([^.?!]+)/i);
      const attitude = text.match(/\battitude(?:\s+\w+){0,6}?\s+(hostile|unfriendly|neutral|friendly|allied)\b/i);
      const status = text.match(/\bstatus(?:\s+is|:)\s+([^.?!]+)/i);
      const explicit = text.match(/\b(?:log|note|record)\s*:\s*(.+)$/i);

      if (fullHeal) {
        const current = parseNumber(npc.fields.HP);
        const max = parseNumber(npc.fields.HPMax);
        if (!Number.isNaN(max)) {
          fields.HP = String(max);
          fields.Status = (root.Tablelight.rules ? root.Tablelight.rules.healthOf(max, max).label : "Healthy");
        }
        const gained = (!Number.isNaN(max) && !Number.isNaN(current)) ? Math.max(0, max - current) : null;
        entry = gained ? "Healed all damage (" + gained + " HP)." : "Healed all damage.";
        type = "Healing";
      } else if (healing) {
        const amount = parseInt(healing[1], 10);
        const current = parseNumber(npc.fields.HP);
        const max = parseNumber(npc.fields.HPMax);
        if (!Number.isNaN(current)) {
          let next = current + amount;
          if (!Number.isNaN(max)) next = Math.min(max, next);
          fields.HP = String(next);
          fields.Status = (root.Tablelight.rules ? root.Tablelight.rules.healthOf(next, npc.fields.HPMax).label : "Injured");
        }
        entry = "Healed " + amount + " HP.";
        type = "Healing";
      } else if (allDamage) {
        const current = parseNumber(npc.fields.HP);
        fields.HP = "0";
        fields.Status = (root.Tablelight.rules ? root.Tablelight.rules.healthOf(0, npc.fields.HPMax).label : "Down");
        entry = !Number.isNaN(current) ? "Took all remaining damage (" + current + " HP)." : "Took all remaining damage.";
        type = "Damage";
      } else if (damage) {
        const amount = parseInt(damage[1], 10);
        const current = parseNumber(npc.fields.HP);
        if (!Number.isNaN(current)) {
          const next = Math.max(0, current - amount);
          fields.HP = String(next);
          fields.Status = (root.Tablelight.rules ? root.Tablelight.rules.healthOf(next, npc.fields.HPMax).label : (next === 0 ? "Down" : "Injured"));
        }
        const source = damage[2] ? " from " + damage[2].trim() : "";
        entry = "Took " + amount + " damage" + source + ".";
        type = "Damage";
      }

      if (hpSet && !damage && !healing) {
        fields.HP = String(parseInt(hpSet[1], 10));
        entry = explicit ? explicit[1].trim() : "HP set to " + fields.HP + ".";
        type = "Note";
      }
      if (acSet) fields.AC = String(parseInt(acSet[1], 10));
      if (moved) {
        fields.Location = moved[1].trim();
        if (!damage && !healing) {
          entry = "Now at " + fields.Location + ".";
          type = "Travel";
        }
      }
      if (attitude) {
        const label = attitude[1].toLowerCase();
        fields.Attitude = label.charAt(0).toUpperCase() + label.slice(1);
        const scale = { hostile: "-2", unfriendly: "-1", neutral: "0", friendly: "1", allied: "2" };
        fields.AttitudeValue = scale[label];
        if (!damage && !healing && !moved) {
          entry = "Attitude toward the party is now " + fields.Attitude + ".";
          type = "Attitude";
        }
      }
      if (status && root.Tablelight.rules) {
        const word = status[1].trim();
        const condition = root.Tablelight.rules.CONDITIONS.find((name) => name.toLowerCase() === word.toLowerCase());
        if (condition) fields.Conditions = condition;
      }
      if (explicit && !damage && !healing) {
        entry = explicit[1].trim();
        type = "Note";
      }

      return {
        sheet: npc.sheet,
        name: npc.fields.Name,
        before: Object.assign({}, npc.fields),
        fields,
        interaction: {
          date: (campaign && campaign.InGameDate) || "",
          session: (campaign && campaign.SessionNumber) || "",
          type,
          entry,
        },
      };
    }).filter((proposal) => Object.keys(proposal.fields).length || proposal.interaction.entry);
  }

  function parseNewNpc(text) {
    const match = text.trim().match(/^(?:new|add|create)\s+npc\s*[:\-]\s*(.+)$/i);
    if (!match) return null;
    const bits = match[1].split(",").map((part) => part.trim()).filter(Boolean);
    return {
      Name: bits[0] || "",
      Race: bits[1] || "",
      ClassRole: bits[2] || "",
      Location: bits[3] || "",
      Concept: bits.slice(4).join(", "),
      Status: "Healthy",
      Attitude: "Neutral",
      AttitudeValue: "0",
    };
  }

function route(text, ctx) {
  const named = matchNpcs(ctx.npcs, text);
  const loreHits = searchLore(ctx.lore, text);
  const draft = parseNewNpc(text);
  const scene = NARRATIVE_RE.test(text);
  const write = !draft && !scene && named.length > 0 && !RULES_RE.test(text) && (WRITE_RE.test(text) || ATTITUDE_WRITE_RE.test(text) || LOG_WRITE_RE.test(text));
  const npcs = scene ? matchNpcs(ctx.npcs, text, { includeLocation: true }) : named;
    const intents = [];

    if (draft) {
      return {
        npcs: [],
        loreHits: [],
        intents: [],
        write: false,
        draft,
        proposals: [],
        dispatch: "NPC & State Manager · new sheet",
      };
    }

    if (write) {
      return {
        npcs,
        loreHits,
        intents: ["npc"],
        write: true,
        draft: null,
        proposals: proposeUpdates(text, npcs, ctx.campaign),
        dispatch: "NPC & State Manager · update " + npcs.map((npc) => npc.fields.Name).join(", "),
      };
    }

    if (NARRATIVE_RE.test(text)) intents.push("narrative");
    if (COMBAT_RE.test(text)) intents.push("combat");
    if (RULES_RE.test(text)) intents.push("rules");
    const loreAsk = LORE_RE.test(text) || (loreHits.length && /\b(who|what|where|why|status|working)\b/i.test(text));
    if (loreAsk) intents.push("lore");

    const possessive = npcs.length && /\b(ac|hp|stats?|status|where|who|attitude|want|sheet|condition)\b/i.test(text);
    if (possessive && !/\b(rules?|spells?|how does|conditions?|grappled|advantage)\b/i.test(text)) {
      const rulesAt = intents.indexOf("rules");
      if (rulesAt >= 0) intents.splice(rulesAt, 1);
    }

    const npcQuestion = npcs.length && (possessive || intents.length === 0 || intents.includes("narrative") || intents.includes("combat"));
    if (npcQuestion && !intents.includes("npc")) intents.unshift("npc");
    if (npcs.length && intents.includes("narrative") && !intents.includes("npc")) intents.unshift("npc");

    const labels = [];
    if (intents.includes("npc")) {
      labels.push("NPC & State Manager · read " + npcs.map((npc) => npc.fields.Name).join(", "));
    }
    intents.forEach((id) => {
      if (id !== "npc") labels.push(AGENTS[id].label);
    });

    return {
      npcs,
      loreHits,
      intents,
      write: false,
      draft: null,
      proposals: [],
      dispatch: labels.join(" → ") || "No specialist matched",
    };
  }

  function campaignBlock(ctx) {
    const c = ctx.campaign || {};
    const party = (ctx.party || []).map((member) => {
      return [member.name, member.className, "level " + member.level, "AC " + member.ac, "HP " + member.hp, member.notes].filter(Boolean).join(" · ");
    });
    return [
      "CAMPAIGN CONTEXT",
      "Name: " + (c.CampaignName || "Untitled"),
      "Ruleset: " + (c.Ruleset || "Unspecified"),
      "Party level: " + (c.PartyLevel || "?"),
      "Party size: " + (c.PartySize || "?"),
      "Session: " + (c.SessionNumber || "?"),
      "In-game date: " + (c.InGameDate || "unspecified"),
      "Boundaries: " + (c.Boundaries || "None recorded."),
      "House rules: " + (c.HouseRules || "None recorded."),
      "",
      "PARTY ROSTER",
      party.length ? party.join("\n") : "No adventurers recorded.",
    ].join("\n");
  }

  function npcBlock(npc) {
    const f = npc.fields;
    const recent = (npc.interactions || []).slice(-3);
    return [
      "Name & Concept: " + [f.Name, f.Race, f.ClassRole].filter(Boolean).join(", ") + (f.Concept ? " — " + f.Concept : ""),
      "Current Location & Status: " + [f.Location || "Unknown location", f.Status || "Unknown condition"].join(", "),
      "Attitude towards Party: " + [f.Attitude || "Neutral", f.AttitudeValue].filter((part) => part !== undefined && part !== "").join(" "),
      "Key Stats/AC/HP: AC " + (f.AC || "—") + " · HP " + (f.HP || "—") + (f.HPMax ? "/" + f.HPMax : "") + (f.KeyStats ? " · " + f.KeyStats : ""),
      "Voice / Tone: " + (f.Voice || "Not recorded"),
      "Key Motivation: " + (f.Motivation || "Not recorded"),
      "Recent Interactions:",
      recent.length ? recent.map((item) => "- Session " + item.session + " / " + item.date + " / " + item.type + ": " + item.entry).join("\n") : "- None recorded",
    ].join("\n");
  }

  function loreBlock(entries) {
    if (!entries || !entries.length) return "KNOWLEDGE SOURCES\nNone matched in the campaign workbook.";
    return "KNOWLEDGE SOURCES\n" + entries.map((entry) => {
      return "### " + entry.title + " [" + entry.tags + "]\n" + entry.body;
    }).join("\n\n");
  }

  const PROMPTS = {
    rules: [
      "You are a TTRPG Rules Engine. Provide exact, concise rule references, mechanics, and calculations.",
      "State the official rule immediately, followed by a 1-sentence mechanical summary.",
      "Structure the response as:",
      "**Rule/Feature Name**: [Name]",
      "**Page / Source Reference**: [Source if known. If you are not sure of the page, write Page unknown — confirm in the book. Never invent a page number.]",
      "**Mechanical Effect**: [Exact mechanics, DC, saving throws, damage types, in your own words. Do not quote long copyrighted rules text.]",
      "**DM Ruling Suggestion**: [Quick guidance for edge cases or common house rules]",
      "Never generate narrative prose or NPC dialogue.",
      "If a house rule in the campaign context touches the question, mention it inside DM Ruling Suggestion.",
      "Obey the campaign boundaries. If the request breaks them, refuse in one sentence.",
    ].join("\n"),
    narrative: [
      "You are the Narrative & Scene Assistant. Draft evocative, sensory-rich descriptions for locations, atmospheric events, and NPC dialogue.",
      "Keep scene descriptions under 4 sentences.",
      "Include at least 2 distinct senses in location descriptions.",
      "For NPC dialogue, provide:",
      "**Voice / Tone Note**: [from the sheet when one was provided]",
      "**Key Motivation**: [what they want right now]",
      "**Sample Line**: \"[Direct dialogue sentence]\"",
      "Do not roll dice or interpret mechanics.",
      "Use the NPC state and campaign notes when they are provided. You may add sensory detail that does not contradict them. Do not invent faction facts, secret loyalties, or stats.",
      "Obey the campaign boundaries. If the request breaks them, refuse in one sentence.",
    ].join("\n"),
    combat: [
      "You are the Combat & Encounter Specialist. Balance battles, design tactical terrain, and outline monster behaviors.",
      "When generating encounters, provide:",
      "**Encounter Summary**: [Difficulty rating relative to party level]",
      "**Tactics**: [1-2 sentences on how creatures behave, priority targets, or morale triggers]",
      "**Environmental Features**: [1-2 hazards or dynamic battlefield elements]",
      "Structure stat adjustments as bullet points.",
      "Keep tactical summaries short enough to run at the table.",
      "Use the party roster and any NPC stat lines provided. Do not invent a page citation.",
      "Obey the campaign boundaries. If the request breaks them, refuse in one sentence.",
    ].join("\n"),
    lore: [
      "You are the World & Lore Manager. Store and retrieve persistent world data, faction agendas, and location histories.",
      "Ground every answer strictly in the knowledge sources provided. If the sources do not say it, say it is not in the campaign notes.",
      "For NPC or faction queries, output:",
      "**Current Status/Location**: [Where they are now]",
      "**Relationship to Party**: [Friendly / Hostile / Neutral + Reason]",
      "**Active Goals**: [What they are working on behind the scenes]",
      "Flag contradictions with the notes in a final line that begins **Contradiction**:.",
      "Do not invent history to fill gaps.",
      "Obey the campaign boundaries. If the request breaks them, refuse in one sentence.",
    ].join("\n"),
    npcWrite: [
      "You are the NPC & State Manager. Convert the dungeon master's sentence into a sheet update.",
      "Return only JSON with this shape:",
      '{"sheet":"Sheet_Name","fieldUpdates":{"HP":"12","Status":"Injured","Location":null,"Attitude":null,"AttitudeValue":null,"AC":null},"interaction":{"type":"Damage","entry":"short log line"}}',
      "Use null for any field that does not change.",
      "Allowed field keys: Name, Race, ClassRole, Concept, Location, Status, Attitude, AttitudeValue, AC, HP, HPMax, KeyStats, Voice, Motivation.",
      "Attitude must be Hostile, Unfriendly, Neutral, Friendly, or Allied when set.",
      "Do not invent stats that were not implied by the sentence or the current sheet.",
    ].join("\n"),
  };

  function userPrompt(agentId, text, ctx, plan) {
    const parts = [campaignBlock(ctx), ""];
    if (plan.npcs.length && (agentId === "narrative" || agentId === "combat" || agentId === "npc")) {
      parts.push("NPC STATE");
      plan.npcs.forEach((npc) => {
        parts.push(npcBlock(npc));
        parts.push("");
      });
    }
    if (agentId === "lore" || agentId === "narrative" || agentId === "combat") {
      parts.push(loreBlock(agentId === "lore" ? (plan.loreHits.length ? plan.loreHits : ctx.lore) : plan.loreHits));
      parts.push("");
    }
    parts.push("DUNGEON MASTER REQUEST");
    parts.push(text);
    return parts.join("\n");
  }

  function hasModel(settings) {
    return Boolean(settings && settings.apiKey && settings.baseUrl && settings.model);
  }

  async function callModel(settings, options) {
    if (!hasModel(settings)) {
      const error = new Error("NO_KEY");
      error.code = "NO_KEY";
      throw error;
    }
    const base = String(settings.baseUrl || "").replace(/\/$/, "");
    async function attempt(useJson) {
      const headers = {
        Authorization: "Bearer " + settings.apiKey,
        "Content-Type": "application/json",
      };
      if (/openrouter\.ai/i.test(base)) {
        headers["HTTP-Referer"] = (root.location && root.location.href) || "http://localhost";
        headers["X-Title"] = "Tablelight";
      }
      const body = {
        model: settings.model,
        temperature: options.temperature,
        messages: [
          { role: "system", content: options.system },
          { role: "user", content: options.user },
        ],
      };
      if (options.maxTokens) body.max_tokens = options.maxTokens;
      if (useJson) body.response_format = { type: "json_object" };
      const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
      const timer = controller ? setTimeout(() => controller.abort(), 45000) : null;
      try {
        const response = await fetch(base + "/chat/completions", {
          method: "POST",
          headers,
          body: JSON.stringify(body),
          signal: controller ? controller.signal : undefined,
        });
        const raw = await response.text();
        if (!response.ok) {
          const error = new Error(raw.slice(0, 400) || ("Model request failed (" + response.status + ")"));
          error.status = response.status;
          throw error;
        }
        const data = JSON.parse(raw);
        const content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
        if (!content) throw new Error("The model returned an empty answer.");
        return content;
      } catch (error) {
        if (error && error.name === "AbortError") throw new Error("The model took too long to answer.");
        if (error instanceof TypeError) throw new Error("The browser could not reach the model endpoint. Check the URL. For Ollama, start it and allow this site's origin.");
        throw error;
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    try {
      return await attempt(Boolean(options.json));
    } catch (error) {
      if (options.json && /response_format|json_object|unsupported/i.test(error.message || "")) {
        return attempt(false);
      }
      if (options.maxTokens && /max_tokens|max_completion_tokens/i.test(error.message || "")) {
        const retry = Object.assign({}, options, { maxTokens: 0 });
        return callModel(settings, retry);
      }
      if (error.status === 401 || error.status === 403) {
        throw new Error("The model rejected the API key.");
      }
      throw error;
    }
  }

  function parseModelJson(text) {
    const trimmed = String(text || "").trim().replace(/^```(?:json)?\s*/i, "").replace(/```$/, "");
    const start = trimmed.indexOf("{");
    const end = trimmed.lastIndexOf("}");
    if (start < 0 || end < start) throw new Error("The model did not return a sheet update.");
    return JSON.parse(trimmed.slice(start, end + 1));
  }

  function modelProposal(parsed, npcs, campaign) {
    const sheet = parsed.sheet;
    const npc = (npcs || []).find((item) => item.sheet === sheet) || (npcs || []).find((item) => normalize(item.fields.Name) === normalize(parsed.name));
    if (!npc) throw new Error("The model named an NPC that is not in the workbook.");
    const allowed = new Set(TL.workbook.FIELD_ORDER);
    const fields = {};
    const updates = parsed.fieldUpdates || parsed.fields || {};
    Object.keys(updates).forEach((key) => {
      if (!allowed.has(key)) return;
      if (updates[key] === null || updates[key] === undefined || updates[key] === "") return;
      fields[key] = String(updates[key]).slice(0, 500);
    });
    const interaction = parsed.interaction || {};
    return {
      sheet: npc.sheet,
      name: npc.fields.Name,
      before: Object.assign({}, npc.fields),
      fields,
      interaction: {
        date: (campaign && campaign.InGameDate) || "",
        session: (campaign && campaign.SessionNumber) || "",
        type: String(interaction.type || "Note").slice(0, 40),
        entry: String(interaction.entry || "").slice(0, 500),
      },
    };
  }

  function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

  function wordTokens(text) {
    const words = normalize(text).split(" ").filter((word) => word.length > 2 && !STOP.has(word));
    const stems = [];
    words.forEach((word) => {
      if (word.length > 4 && word.endsWith("s")) stems.push(word.slice(0, -1));
    });
    return words.concat(stems);
  }

  function kindFromTags(tags) {
    const padded = " " + normalize(tags) + " ";
    if (padded.includes(" monster ")) return "monster";
    if (padded.includes(" location ") || padded.includes(" region ")) return "location";
    if (padded.includes(" faction ")) return "faction";
    if (padded.includes(" quest ")) return "quest";
    if (padded.includes(" rules ")) return "rules";
    if (padded.includes(" item ")) return "item";
    return "note";
  }

  function library(ctx) {
    const docs = [];
    (ctx.npcs || []).forEach((npc) => {
      const fields = npc.fields || {};
      docs.push({
        kind: "npc",
        title: fields.Name || npc.sheet,
        sheet: npc.sheet,
        tags: [fields.Race, fields.RaceCustom, fields.Class, fields.ClassCustom, fields.Role, fields.RoleCustom, fields.ClassRole, fields.Location, "person", "npc"].filter(Boolean).join(" "),
        text: [fields.Name, fields.Race, fields.RaceCustom, fields.Class, fields.ClassCustom, fields.Role, fields.RoleCustom, fields.ClassRole, fields.Concept, fields.History, fields.Location, fields.Voice, fields.Motivation, fields.KeyStats].filter(Boolean).join("\n"),
        npc,
      });
    });
    (ctx.lore || []).forEach((entry) => {
      docs.push({
        kind: kindFromTags(entry.tags),
        title: entry.title,
        sheet: "",
        tags: entry.tags || "",
        text: [entry.title, entry.tags, entry.body].filter(Boolean).join("\n"),
        entry,
      });
    });
    return docs;
  }

  function scoreDoc(doc, query, words) {
    const title = normalize(doc.title);
    const q = normalize(query);
    let score = 0;
    if (title && q.includes(title)) score += 16;
    title.split(" ").filter((word) => word.length > 2 && !STOP.has(word)).forEach((word) => {
      const stem = word.length > 4 && word.endsWith("s") ? word.slice(0, -1) : word;
      if (words.includes(word) || words.includes(stem)) score += 5;
    });
    const tagWords = normalize(doc.tags).split(" ").filter(Boolean);
    words.forEach((word) => {
      if (tagWords.includes(word)) score += 6;
      else if (normalize(doc.text).includes(word)) score += 1;
    });
    return score;
  }

  function sectionMap(body) {
    const out = [];
    let name = "";
    let lines = [];
    function flush() {
      if (!name) return;
      out.push({ name: name.toLowerCase(), text: lines.join("\n").trim() });
    }
    String(body || "").split("\n").forEach((line) => {
      const match = /^([A-Za-z][A-Za-z '&/-]{0,40}):\s*(.*)$/.exec(line);
      if (match) {
        flush();
        name = match[1];
        lines = match[2] ? [match[2]] : [];
      } else if (name) {
        lines.push(line);
      }
    });
    flush();
    return out;
  }

  function sectionText(body, label) {
    const found = sectionMap(body).find((section) => section.name === label.toLowerCase());
    return found ? found.text : "";
  }

  function formatSections(body, labels) {
    const blocks = [];
    labels.forEach((label) => {
      const value = sectionText(body, label);
      if (!value) return;
      const prose = label === "Read-aloud" || label === "Senses" || label === "Summary";
      const block = label === "Stat block" || label === "Traits" || label === "Actions" || label === "Reactions";
      if (prose) {
        value.split("\n").forEach((line) => {
          if (line.trim()) blocks.push(line.trim());
        });
        return;
      }
      if (block) {
        blocks.push(label);
        value.split("\n").forEach((line) => {
          if (line.trim()) blocks.push(line.trim());
        });
        return;
      }
      blocks.push("**" + label + "**: " + value.replace(/\s*\n\s*/g, " "));
    });
    return blocks.join("\n");
  }

  function oneLiner(doc) {
    if (doc.kind === "npc") {
      const fields = doc.npc.fields;
      const rules = root.Tablelight.rules;
      const job = rules ? (rules.shown(fields.Role, fields.RoleCustom) || rules.shown(fields.Class, fields.ClassCustom) || fields.ClassRole) : fields.ClassRole;
      return [job, fields.Location].filter(Boolean).join(", ");
    }
    const summary = sectionText(doc.entry.body, "Summary") || sectionText(doc.entry.body, "Hook") || sectionText(doc.entry.body, "Read-aloud") || sectionText(doc.entry.body, "Active goals");
    const sentence = summary.split(". ")[0];
    if (!sentence) return "";
    return sentence.endsWith(".") ? sentence : sentence + ".";
  }

  function namesADocument(query, docs) {
    const q = normalize(query);
    return docs.some((doc) => {
      const title = normalize(doc.title);
      if (!title) return false;
      if (q.includes(title)) return true;
      const words = title.split(" ").filter((word) => word.length > 3 && !STOP.has(word));
      const hits = words.filter((word) => q.includes(word));
      return hits.length >= 2 || (words.length === 1 && hits.length === 1 && words[0].length > 5);
    });
  }

  function listKind(query, docs) {
    const q = normalize(query);
    if (namesADocument(query, docs)) return "";
    if (/\bquests?\b|\bhooks?\b|\bjobs?\b/.test(q)) return "quest";
    if (/\b(people|npcs|everyone|who lives|who is here|whos here)\b/.test(q)) return "npc";
    if (/\bfactions?\b/.test(q)) return "faction";
    if (/\blocations?\b|\bplaces?\b/.test(q)) return "location";
    if (/\bmonsters?\b|\bcreatures?\b|\bstat blocks\b/.test(q)) return "monster";
    return "";
  }

  const SECTION_ORDER = {
    monster: ["Summary", "Stat block", "Traits", "Actions", "Reactions", "Equipment", "Tactics"],
    location: ["Read-aloud", "Senses", "For the DM", "People", "Encounters", "Checks"],
    faction: ["Summary", "Current status/location", "Relationship to party", "Active goals", "For the DM"],
    quest: ["Summary", "Hook", "Steps", "Reward", "Current status/location", "Relationship to party", "Active goals", "For the DM"],
    item: ["Summary", "Current status/location", "Relationship to party", "Active goals", "For the DM"],
    rules: ["Summary", "For the DM", "Checks"],
    note: ["Summary", "For the DM"],
  };

  function personKind(fields) {
    const rules = root.Tablelight.rules;
    const species = rules ? rules.shown(fields.Race, fields.RaceCustom) : (fields.Race === "Other" ? fields.RaceCustom : fields.Race);
    const job = rules ? (rules.shown(fields.Role, fields.RoleCustom) || rules.shown(fields.Class, fields.ClassCustom)) : fields.ClassRole;
    return [species, job].filter(Boolean).join(" ").toLowerCase();
  }

  function composeNpc(npc) {
    const fields = npc.fields;
    const rules = root.Tablelight.rules;
    const parts = String(fields.Concept || "").split(/\n?Secret:\s*/i);
    const lines = [];
    const kind = personKind(fields);
    const article = kind && /^[aeiou]/i.test(kind) ? "an " : "a ";
    lines.push(fields.Name + " is " + (kind ? article + kind : "recorded") + (fields.Location ? ", at " + fields.Location : "") + ".");
    if (parts[0].trim()) lines.push(parts[0].trim());
    const health = rules ? rules.healthOf(fields.HP, fields.HPMax) : { label: fields.Status || "Unset", percent: "" };
    lines.push("They are " + health.label + (health.percent !== "" ? " (" + health.percent + "% of their hit points)" : "") + ", and " + (fields.Attitude || "Neutral") + " toward the party (" + (fields.AttitudeValue || "0") + ").");
    if (fields.Alignment) lines.push("**Alignment**: " + fields.Alignment);
    if (fields.Conditions) lines.push("**Conditions**: " + fields.Conditions);
    if (fields.Motivation) lines.push(fields.Motivation);
    if (fields.Voice) lines.push("**Voice**: " + fields.Voice);
    if (parts[1]) lines.push("**For the DM**: " + parts[1].trim());
    if (fields.History) lines.push("**History**: " + fields.History);
    lines.push("**AC**: " + (fields.AC || "—"));
    lines.push("**HP**: " + (fields.HP || "—") + (fields.HPMax ? "/" + fields.HPMax : ""));
    if (rules) {
      const scores = rules.ABILITIES.map((ability) => {
        const score = fields[ability];
        if (!score) return "";
        return ability + " " + score + " (" + rules.abilityMod(score) + ")";
      }).filter(Boolean);
      if (scores.length) lines.push("**Abilities**: " + scores.join(", "));
    }
    if (fields.KeyStats) lines.push("**Combat**: " + fields.KeyStats);
    if ((npc.equipment || []).length) {
      lines.push("**Equipment**:");
      npc.equipment.forEach((item) => lines.push("- " + item.item + (item.notes ? " — " + item.notes : "")));
    }
    const recent = (npc.interactions || []).slice(-3);
    if (recent.length) {
      lines.push("**Recent**:");
      recent.forEach((item) => lines.push("- " + item.entry));
    }
    return lines.join("\n");
  }

  function residents(ctx, title) {
    const words = normalize(title).split(" ").filter((word) => word.length > 4);
    return (ctx.npcs || []).filter((npc) => {
      const loc = normalize(npc.fields.Location || "");
      return words.some((word) => loc.includes(word));
    }).slice(0, 6).map((npc) => {
      return npc.fields.Name + " (" + [npc.fields.ClassRole, npc.fields.Location].filter(Boolean).join(", ") + ")";
    });
  }

  function composeDoc(doc, ctx) {
    if (doc.kind === "npc") return composeNpc(doc.npc);
    const order = SECTION_ORDER[doc.kind] || SECTION_ORDER.note;
    let body = formatSections(doc.entry.body, order);
    if (doc.kind === "location") {
      const names = residents(ctx, doc.title);
      if (names.length) body += "\n**On the sheets**: " + names.join("; ");
    }
    if (doc.kind === "rules" && ctx.campaign && ctx.campaign.HouseRules) {
      body += "\n**House rules**: " + ctx.campaign.HouseRules;
    }
    return body;
  }

  function inventory(ctx) {
    const counts = {};
    library(ctx).forEach((doc) => {
      counts[doc.kind] = (counts[doc.kind] || 0) + 1;
    });
    const labels = { npc: "people", location: "places", faction: "factions", quest: "quests", monster: "monsters", item: "items" };
    return Object.keys(labels).filter((key) => counts[key]).map((key) => counts[key] + " " + labels[key]).join(", ");
  }

  function partyReply(ctx) {
    const members = ctx.party || [];
    const lines = ["The roster on the campaign sheet."];
    if (!members.length) lines.push("No adventurers are written down.");
    members.forEach((member) => {
      lines.push("- **" + (member.name || "Unnamed") + "**: " + [member.className, member.level ? "level " + member.level : "", member.ac ? "AC " + member.ac : "", member.hp ? "HP " + member.hp : "", member.notes].filter(Boolean).join(", "));
    });
    if (ctx.campaign) {
      lines.push("**Party level**: " + (ctx.campaign.PartyLevel || "—"));
      lines.push("**Party size**: " + (ctx.campaign.PartySize || "—"));
    }
    return { title: "The party", body: lines.join("\n"), sheet: "" };
  }

  function catalog(kind, docs, query, ctx) {
    const words = wordTokens(query);
    let items = docs.filter((doc) => doc.kind === kind);
    const places = docs.filter((doc) => doc.kind === "location").map((doc) => ({ doc, score: scoreDoc(doc, query, words) })).filter((item) => item.score >= 5).sort((a, b) => b.score - a.score);
    if (places.length && (kind === "monster" || kind === "npc")) {
      const placeWords = normalize(places[0].doc.title).split(" ").filter((word) => word.length > 3 && words.includes(word));
      placeWords.sort((a, b) => b.length - a.length);
      const key = placeWords[0];
      if (key) {
        const narrowed = items.filter((doc) => normalize(doc.tags + " " + doc.text).includes(key));
        if (narrowed.length) items = narrowed;
      }
    }
    const titles = { quest: "Quests on file", monster: "Monsters on file", faction: "Factions on file", npc: "People on file", location: "Places on file" };
    if (!items.length) {
      return { title: titles[kind] || "On file", body: "Nothing of that kind is stored. On file: " + (inventory(ctx) || "no notes") + ".", sheet: "" };
    }
    return { title: titles[kind] || "On file", body: items.map((doc) => "- **" + doc.title + "**: " + oneLiner(doc)).join("\n"), sheet: "" };
  }

  function encounterReply(docs) {
    const lines = ["The notes do not build a new fight. These are the encounters already written down."];
    docs.filter((doc) => doc.kind === "location").forEach((doc) => {
      const text = sectionText(doc.entry.body, "Encounters");
      if (text) lines.push("**" + doc.title + "**: " + text.replace(/\s*\n\s*/g, " "));
    });
    return { title: "Encounters on file", body: lines.join("\n"), sheet: "" };
  }

  function rankedDocs(query, ctx) {
    const words = wordTokens(query);
    return library(ctx).map((doc) => ({ doc, score: scoreDoc(doc, query, words) })).filter((item) => item.score >= 4).sort((a, b) => b.score - a.score);
  }

  function composeReply(text, ctx) {
    const q = normalize(text);
    const docs = library(ctx);
    const kind = listKind(text, docs);
    if (kind) return catalog(kind, docs, text, ctx);
    if (/\b(party|roster|adventurers)\b/.test(q) && !/\b(encounter|fight|combat)\b/.test(q)) return partyReply(ctx);
    if (/\b(encounter|fight|combat)\b/.test(q) && !/\b(stat|block)\b/.test(q)) {
      const ranked = rankedDocs(text, ctx);
      const place = ranked.find((item) => item.doc.kind === "location");
      if (place && place.score >= 8) return { title: place.doc.title, body: composeDoc(place.doc, ctx), sheet: "" };
      return encounterReply(docs);
    }
    const ranked = rankedDocs(text, ctx);
    if (!ranked.length) {
      return {
        title: "Not in the notes",
        body: "Nothing stored matches that. On file: " + (inventory(ctx) || "no notes") + ". Ask for a person, a place, a faction, a quest, or a monster by name.",
        sheet: "",
      };
    }
    const focused = /\b(describe|who|where|stat)\b/.test(q);
    const chosen = ranked.filter((item, index) => index === 0 || (!focused && item.score >= ranked[0].score - 2)).slice(0, 2);
    return {
      title: chosen.length === 1 ? chosen[0].doc.title : chosen.map((item) => item.doc.title).join(" · "),
      body: chosen.map((item) => composeDoc(item.doc, ctx)).join("\n\n"),
      sheet: chosen.length === 1 ? (chosen[0].doc.sheet || "") : "",
    };
  }

  function isPlay(text) {
    const raw = String(text || "").trim();
    if (/^(who|what|where|when|why|how|describe|stat|list)\b/i.test(raw) && !/^what happens\b/i.test(raw)) return false;
    if (/\bstat block\b/i.test(raw)) return false;
    return /\b(we|i|let's|lets|the party)\b/i.test(raw) && /\b(enter|go|head|walk|approach|arrive|search|investigate|inspect|talk|speak|sneak|hide|wait|rest|ring|open|climb|cross|attack|fight|strike|try|look)\b/i.test(raw);
  }

  function playSeed(text, ctx) {
    const clock = ctx.campaign ? String(ctx.campaign.InGameDate || "") + " " + String(ctx.campaign.InGameTime || "") : "";
    const hp = (ctx.npcs || []).map((npc) => (npc.fields && npc.fields.HP) || "").join(",");
    let hash = 2166136261;
    const source = String(text || "") + "|" + clock + "|" + hp;
    for (let i = 0; i < source.length; i++) {
      hash ^= source.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0) % 20 + 1;
  }

  function playMode(text) {
    const q = normalize(text);
    if (/\b(attack|fight|strike|charge)\b/.test(q)) return "fight";
    if (/\b(talk|speak)\b/.test(q)) return "talk";
    if (/\b(search|investigate|inspect|sneak)\b/.test(q) || /\b(investigation|religion|survival|perception|stealth|insight|persuasion|athletics|arcana)\b/.test(q)) return "check";
    return "scene";
  }

  function bestOfKind(text, ctx, kind, floor) {
    const hit = rankedDocs(text, ctx).find((item) => item.doc.kind === kind && item.score >= floor);
    return hit ? hit.doc : null;
  }

  function splitSentences(text) {
    return String(text || "").split(/(?<=\.)\s+/).map((sentence) => sentence.trim()).filter(Boolean);
  }

  function parseChecks(text) {
    const found = [];
    splitSentences(text).forEach((sentence) => {
      let match = sentence.match(/\b([A-Za-z]+)\s+DC\s+(\d+)/i);
      if (match) {
        found.push({ skill: match[1].toLowerCase(), dc: parseInt(match[2], 10), success: sentence });
        return;
      }
      match = sentence.match(/\bDC\s+(\d+)\s+([A-Za-z]+)/i);
      if (match) found.push({ skill: match[2].toLowerCase(), dc: parseInt(match[1], 10), success: sentence });
    });
    return found;
  }

  function encounterLines(body, action) {
    const actionWords = wordTokens(action);
    return splitSentences(sectionText(body, "Encounters")).filter((sentence) => {
      const conditional = /\bif\b|\bonly if\b|\bwhen\b|\bunless\b/i.test(sentence);
      if (!conditional) return true;
      const words = wordTokens(sentence);
      return words.some((word) => word.length > 4 && actionWords.includes(word));
    });
  }

  function publicPeople(place) {
    return sectionText(place.entry.body, "People");
  }

  function clockLine(ctx) {
    if (!ctx.campaign) return "";
    return [ctx.campaign.InGameTime, ctx.campaign.InGameDate].filter(Boolean).join(" · ");
  }

  function tryVerbs(skills) {
    const verbs = {
      investigation: "investigate",
      religion: "ask what this place means",
      survival: "search for a way through",
      perception: "look closer",
      stealth: "move quietly",
      insight: "watch who is here",
      persuasion: "talk your way on",
      athletics: "climb",
      arcana: "study the working",
    };
    const lines = [];
    skills.forEach((skill) => {
      const verb = verbs[skill] || skill;
      if (lines.indexOf(verb) < 0) lines.push(verb);
    });
    return lines.slice(0, 3);
  }

  function proseResult(sentence) {
    let text = String(sentence || "")
      .replace(/\b[A-Za-z]+\s+DC\s+\d+\s+/i, "")
      .replace(/\ba?\s*DC\s+\d+\s+[A-Za-z]+\s+check\s+(with it,?\s+)?/i, "")
      .replace(/^(shows|knows|tells|notices|finds|spots|catches|reads|recalls)\s+(that\s+)?/i, "");
    text = text.trim().replace(/^[,:\s]+/, "");
    if (!text) return sentence;
    text = text.charAt(0).toUpperCase() + text.slice(1);
    if (!/[.!?]$/.test(text)) text += ".";
    return text;
  }

  function beat(kind, text) {
    const value = String(text || "").trim();
    if (!value) return "";
    return "[[" + kind + "]]" + value;
  }

  function playScene(place, text, ctx) {
    const lines = [];
    lines.push(beat("place", sectionText(place.entry.body, "Read-aloud")));
    lines.push(beat("sense", sectionText(place.entry.body, "Senses")));
    const here = encounterLines(place.entry.body, text);
    if (here.length) lines.push(beat("act", here.join(" ")));
    lines.push(beat("act", publicPeople(place)));
    const verbs = tryVerbs(parseChecks(sectionText(place.entry.body, "Checks")).map((check) => check.skill));
    if (verbs.length) lines.push(beat("next", "You could " + verbs.join(", or ") + "."));
    if (/\b(rest|wait)\b/i.test(text)) lines.push(beat("next", "You could rest. Nothing heals until you say who, and by how much."));
    return { title: place.title, body: lines.filter(Boolean).join("\n\n"), status: clockLine(ctx), sheet: "" };
  }

  function playCheck(place, text, ctx) {
    const checks = parseChecks(sectionText(place.entry.body, "Checks"));
    const q = normalize(text);
    const stems = { investigate: "investigation", search: "investigation", inspect: "investigation", religion: "religion", survive: "survival", survival: "survival", sneak: "stealth", hide: "stealth", perceive: "perception", perception: "perception", persuade: "persuasion", persuasion: "persuasion", insight: "insight", athletics: "athletics", stealth: "stealth", arcana: "arcana" };
    const asked = Object.keys(stems).filter((word) => q.includes(word)).map((word) => stems[word]);
    const matches = checks.filter((check) => asked.indexOf(check.skill) >= 0 || q.includes(check.skill));
    if (!matches.length) {
      const scene = playScene(place, text, ctx);
      scene.body += "\n\n" + beat("next", "Nothing you try has a result written here.");
      return scene;
    }
    const words = wordTokens(text);
    matches.sort((a, b) => {
      const overlap = (sentence) => wordTokens(sentence).filter((word) => words.includes(word)).length;
      return overlap(b.success) - overlap(a.success);
    });
    const chosen = matches[0];
    const bonus = text.match(/([+-]\s*\d+)/);
    const mod = bonus ? parseInt(bonus[1].replace(/\s+/g, ""), 10) : 0;
    const die = playSeed(text, ctx);
    const total = die + mod;
    const success = total >= chosen.dc;
    const lines = [];
    const signed = mod > 0 ? " + " + mod : mod < 0 ? " - " + Math.abs(mod) : "";
    lines.push(beat("roll", "You roll. d20 " + die + signed + " = " + total + " vs DC " + chosen.dc + "."));
    if (!bonus) lines.push(beat("next", "A bare roll. Name a modifier if you have one, such as " + chosen.skill + " +4."));
    lines.push(beat(success ? "place" : "act", success ? proseResult(chosen.success) : "Nothing comes of it. The place stays as you first saw it."));
    return { title: place.title, body: lines.join("\n\n"), status: clockLine(ctx), sheet: "" };
  }

  function playTalk(npc, ctx) {
    const fields = npc.fields;
    const concept = String(fields.Concept || "").split(/\n?Secret:\s*/i)[0].trim();
    const manner = { Hostile: "They want you gone.", Unfriendly: "They are not glad to see you.", Neutral: "They wait.", Friendly: "They seem glad you came.", Allied: "They treat you as one of their own." };
    const quote = String(fields.Voice || "").match(/"([^"]+)"/);
    const lines = [];
    lines.push(beat("act", (fields.Name || "They") + (fields.Location ? " is at " + fields.Location : " is here") + "."));
    lines.push(beat("emote", manner[fields.Attitude] || manner.Neutral));
    if (concept) lines.push(beat("place", splitSentences(concept)[0]));
    if (quote) lines.push(beat("say", "\"" + quote[1] + "\""));
    lines.push(beat("next", "They will not hand you a secret."));
    return { title: fields.Name, body: lines.filter(Boolean).join("\n\n"), status: clockLine(ctx), sheet: npc.sheet };
  }

  function playFight(monster, place, ctx) {
    const lines = [];
    const summary = sectionText(monster.entry.body, "Summary");
    const actions = splitSentences(sectionText(monster.entry.body, "Actions"));
    if (place) lines.push(beat("place", "At " + place.title + "."));
    if (summary) lines.push(beat("place", summary));
    if (actions.length) lines.push(beat("act", actions[0]));
    lines.push(beat("next", "Say who is hit, and for how much."));
    return { title: monster.title, body: lines.filter(Boolean).join("\n\n"), status: clockLine(ctx), sheet: "" };
  }

  function playTurn(text, ctx) {
    if (!isPlay(text)) return null;
    const mode = playMode(text);
    const place = bestOfKind(text, ctx, "location", 8);
    if (mode === "talk") {
      const people = matchNpcs(ctx.npcs, text);
      if (people.length) return playTalk(people[0], ctx);
    }
    if (mode === "fight") {
      const monster = bestOfKind(text, ctx, "monster", 8);
      if (monster) return playFight(monster, place, ctx);
    }
    if (!place) {
      const places = library(ctx).filter((doc) => doc.kind === "location").map((doc) => doc.title);
      return {
        title: "You can't go there",
        body: places.length ? "Places in this campaign: " + places.join(", ") + "." : "This campaign has no places written down.",
        status: clockLine(ctx),
        sheet: "",
      };
    }
    if (mode === "check") return playCheck(place, text, ctx);
    return playScene(place, text, ctx);
  }

  function preview(text, ctx) {
    const plan = route(text, ctx);
    if (plan.draft) return "New person on a sheet";
    if (plan.write) return "Update " + plan.npcs.map((npc) => npc.fields.Name).join(", ");
    const played = playTurn(text, ctx);
    if (played) return played.title;
    return composeReply(text, ctx).title;
  }

  async function ask(text, ctx, settings) {
    const plan = route(text, ctx);
    if (plan.draft) return [{ kind: "draft-npc", fields: plan.draft }];

    if (plan.write) {
      let proposals = plan.proposals;
      const messages = [];
      const parsedFields = proposals.some((proposal) => Object.keys(proposal.fields).length);
      if (!parsedFields && hasModel(settings)) {
        try {
          const raw = await callModel(settings, {
            system: PROMPTS.npcWrite,
            user: userPrompt("npc", text, ctx, plan),
            temperature: 0.1,
            json: true,
            maxTokens: 500,
          });
          proposals = [modelProposal(parseModelJson(raw), plan.npcs, ctx.campaign)];
        } catch (error) {
          if (error.code !== "NO_KEY") messages.push({ kind: "note", text: "Using your sentence as the log line. " + error.message });
        }
      }
      proposals.forEach((proposal) => {
        if (!Object.keys(proposal.fields).length && !(proposal.interaction && proposal.interaction.entry)) return;
        messages.push({ kind: "proposal", proposal });
      });
      if (!messages.length) messages.push({ kind: "note", text: "That did not change a sheet." });
      return messages;
    }

    const played = playTurn(text, ctx);
    if (played) return [{ kind: "reply", title: played.title, kicker: "At the table", body: played.body, sheet: played.sheet || "", status: played.status || "" }];

    const reply = composeReply(text, ctx);
    return [{ kind: "reply", title: reply.title, kicker: "From the notes", body: reply.body, sheet: reply.sheet || "" }];
  }

  TL.agents = {
    AGENTS,
    normalize,
    matchNpcs,
    searchLore,
    loreFacts,
    proposeUpdates,
    parseDamage,
    parseNewNpc,
    route,
    preview,
    composeReply,
    playTurn,
    ask,
    npcBlock,
    hasModel,
    ping(settings) {
      return callModel(settings, {
        system: "Reply with the single word ready.",
        user: "ping",
        temperature: 0,
        maxTokens: 16,
      });
    },
  };
})(typeof window !== "undefined" ? window : globalThis);
