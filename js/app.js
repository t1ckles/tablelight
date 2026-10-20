(function () {
  const TL = window.Tablelight;
  const canPick = window.isSecureContext && typeof window.showOpenFilePicker === "function";
  const PRESETS = [
    { id: "openai", label: "OpenAI", baseUrl: "https://api.openai.com/v1", model: "gpt-4o-mini" },
    { id: "openrouter", label: "OpenRouter", baseUrl: "https://openrouter.ai/api/v1", model: "openai/gpt-4o-mini" },
    { id: "groq", label: "Groq", baseUrl: "https://api.groq.com/openai/v1", model: "llama-3.3-70b-versatile" },
    { id: "ollama", label: "Ollama on this computer", baseUrl: "http://localhost:11434/v1", model: "llama3.1" },
    { id: "custom", label: "Custom OpenAI-compatible", baseUrl: "", model: "" },
  ];
  const ATTITUDES = ["Hostile", "Unfriendly", "Neutral", "Friendly", "Allied"];
  const INTERACTION_TYPES = ["Met", "Dialogue", "Combat", "Damage", "Healing", "Item", "Attitude", "Travel", "Note"];
  const LABELS = {
    Name: "Name",
    Race: "Race",
    ClassRole: "Class / role",
    Concept: "Concept",
    Location: "Location",
    Status: "Health / condition",
    Attitude: "Attitude",
    AttitudeValue: "Attitude value",
    AC: "AC",
    HP: "HP",
    HPMax: "HP max",
    KeyStats: "Key stats",
    Voice: "Voice / tone",
    Motivation: "Key motivation",
  };

  const state = {
    view: "table",
    wb: null,
    fileHandle: null,
    fileName: "campaign.xlsx",
    dirty: false,
    needsReconnect: false,
    campaign: null,
    lore: [],
    party: [],
    npcs: [],
    log: [],
    messages: [],
    selectedSheet: "",
    selectedLore: "",
    npcQuery: "",
    npcTab: "sheet",
    worldKind: "all",
    busy: false,
    forceStick: false,
    settings: loadSettings(),
  };

  let seq = 1;
  let mirrorTimer = 0;

  function uid() {
    seq += 1;
    return "m" + seq;
  }

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    const a = attrs || {};
    let pendingValue;
    Object.keys(a).forEach((key) => {
      if (a[key] === false || a[key] === null || a[key] === undefined) return;
      if (key === "class") node.className = a[key];
      else if (key === "value") pendingValue = a[key];
      else if (a[key] === true) node.setAttribute(key, "");
      else node.setAttribute(key, a[key]);
    });
    const list = children === undefined || children === null ? [] : Array.isArray(children) ? children : [children];
    const append = (child) => {
      if (child === null || child === undefined || child === false) return;
      if (Array.isArray(child)) {
        child.forEach(append);
        return;
      }
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    };
    list.forEach(append);
    if (pendingValue !== undefined && "value" in node) node.value = pendingValue;
    return node;
  }

  function loadSettings() {
    const fallback = { preset: "openai", baseUrl: "https://api.openai.com/v1", model: "gpt-4o-mini", apiKey: "" };
    try {
      return Object.assign(fallback, JSON.parse(localStorage.getItem("tablelight-settings") || "{}"));
    } catch (error) {
      return fallback;
    }
  }

  function saveSettings() {
    localStorage.setItem("tablelight-settings", JSON.stringify(state.settings));
  }

  function toast(text) {
    const node = el("div", { class: "toast" }, text);
    document.getElementById("toasts").append(node);
    setTimeout(() => node.remove(), 3200);
  }

  function announce(text) {
    const live = document.getElementById("live");
    live.textContent = "";
    setTimeout(() => { live.textContent = text; }, 30);
  }

  function idb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open("tablelight", 1);
      request.onupgradeneeded = () => request.result.createObjectStore("kv");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function idbGet(key) {
    return idb().then((db) => new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readonly");
      const request = tx.objectStore("kv").get(key);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    }));
  }

  function idbSet(key, value) {
    return idb().then((db) => new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    }));
  }

  function idbDel(key) {
    return idb().then((db) => new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    }));
  }

  function hydrate() {
    const data = TL.workbook.read(state.wb);
    state.campaign = data.campaign;
    state.lore = data.lore;
    state.party = data.party;
    state.npcs = data.npcs;
    state.log = data.log;
  }

  function selectedNpc() {
    return state.npcs.find((npc) => npc.sheet === state.selectedSheet) || null;
  }

  function context() {
    return { campaign: state.campaign, lore: state.lore, party: state.party, npcs: state.npcs };
  }

  function flushNow() {
    clearTimeout(mirrorTimer);
    if (!state.wb || !state.campaign) return;
    TL.workbook.saveCampaign(state.wb, state.campaign);
    TL.workbook.saveLore(state.wb, state.lore);
    TL.workbook.saveParty(state.wb, state.party);
    const npc = selectedNpc();
    if (npc && state.wb.Sheets[npc.sheet]) TL.workbook.writeNpc(state.wb, npc);
  }

  async function mirrorOnly() {
    const bytes = TL.workbook.toArrayBuffer(state.wb);
    await idbSet("snapshot", { dirty: state.dirty || !state.fileHandle, name: state.fileName, bytes });
  }

  async function writeHandle() {
    const bytes = TL.workbook.toArrayBuffer(state.wb);
    const writable = await state.fileHandle.createWritable();
    await writable.write(new Blob([bytes]));
    await writable.close();
    state.dirty = false;
    state.fileName = state.fileHandle.name || state.fileName;
    await mirrorOnly();
    paintSaveBadge();
  }

  function scheduleMirror() {
    state.dirty = true;
    paintSaveBadge();
    clearTimeout(mirrorTimer);
    mirrorTimer = setTimeout(async () => {
      try {
        flushNow();
        if (state.fileHandle && !state.needsReconnect) await writeHandle();
        else await mirrorOnly();
      } catch (error) {
        state.needsReconnect = Boolean(state.fileHandle);
        paintSaveBadge();
        toast(error.message || "Could not save the workbook.");
      }
    }, 400);
  }

  function paintSaveBadge() {
    const node = document.getElementById("save-state");
    if (!node) return;
    if (!window.isSecureContext) node.textContent = "Open through a local server so saving can stick";
    else if (state.needsReconnect) node.textContent = "Reconnect to keep saving this file";
    else if (state.fileHandle && state.dirty) node.textContent = "Unsaved";
    else if (state.fileHandle) node.textContent = "Saved";
    else node.textContent = "In this browser";
  }

  function workbookHasWork() {
    return state.npcs.length > 0 || state.lore.length > 0 || state.messages.length > 0 || (state.campaign && state.campaign.CampaignName !== "Untitled campaign");
  }

  function replaceWorkbook(wb, name, handle) {
    flushNow();
    state.wb = wb;
    state.fileName = name || "campaign.xlsx";
    state.fileHandle = handle || null;
    state.needsReconnect = false;
    state.dirty = !handle;
    state.messages = [];
    state.selectedSheet = "";
    state.selectedLore = "";
    hydrate();
    saveMessages();
    scheduleMirror();
    state.view = "table";
    state.forceStick = true;
    render();
  }

  async function rememberHandle(handle) {
    state.fileHandle = handle;
    state.needsReconnect = false;
    if (handle) await idbSet("fileHandle", handle);
    else await idbDel("fileHandle");
  }

  async function openWithPicker() {
    const [handle] = await window.showOpenFilePicker({
      types: [excelType()],
    });
    const file = await handle.getFile();
    const wb = TL.workbook.fromArrayBuffer(await file.arrayBuffer());
    replaceWorkbook(wb, file.name, handle);
    await rememberHandle(handle);
    state.dirty = false;
    await mirrorOnly();
    paintSaveBadge();
    toast("Opened " + file.name);
  }

  function excelType() {
    return {
      description: "Excel workbook",
      accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
    };
  }

  async function saveWorkbook() {
    flushNow();
    try {
      if (state.fileHandle && !state.needsReconnect) {
        await writeHandle();
        toast("Saved to " + state.fileName);
        return;
      }
      if (canPick) {
        const handle = await window.showSaveFilePicker({ suggestedName: state.fileName || "tablelight-campaign.xlsx", types: [excelType()] });
        await rememberHandle(handle);
        state.fileName = handle.name || state.fileName;
        await writeHandle();
        render();
        toast("Saved to " + state.fileName);
        return;
      }
      downloadWorkbook();
    } catch (error) {
      if (error && error.name === "AbortError") return;
      toast(error.message || "Could not save the workbook.");
    }
  }

  function downloadWorkbook() {
    flushNow();
    const bytes = TL.workbook.toArrayBuffer(state.wb);
    const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const link = el("a", { href: url, download: state.fileName || "tablelight-campaign.xlsx" });
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    toast("Downloaded " + (state.fileName || "the workbook"));
  }

  async function reconnect() {
    if (!state.fileHandle) return;
    const perm = await state.fileHandle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") {
      toast("The browser still needs permission for that workbook.");
      return;
    }
    state.needsReconnect = false;
    try {
      if (state.dirty) await writeHandle();
      else {
        const file = await state.fileHandle.getFile();
        state.wb = TL.workbook.fromArrayBuffer(await file.arrayBuffer());
        state.fileName = file.name;
        hydrate();
        await mirrorOnly();
      }
      render();
      toast("Reconnected to " + state.fileName);
    } catch (error) {
      toast(error.message || "Could not reconnect.");
    }
  }

  function askForNewWorkbook(kind) {
    const start = () => {
      const wb = kind === "sample" ? TL.workbook.createSample() : TL.workbook.createBlank();
      replaceWorkbook(wb, kind === "sample" ? "drowned-bell.xlsx" : "campaign.xlsx", null);
      rememberHandle(null);
      toast(kind === "sample" ? "The Drowned Bell is ready" : "Blank campaign is ready");
    };
    if (!workbookHasWork()) {
      start();
      return;
    }
    openModal("Start a new workbook?", el("p", {}, "The campaign on screen will be replaced. Save first if you still need this Excel file."), [
      { label: "Cancel" },
      { label: kind === "sample" ? "Load The Drowned Bell" : "Start blank", primary: true, onClick: start },
    ]);
  }

  function openModal(title, body, actions) {
    const modal = document.getElementById("modal");
    modal.hidden = false;
    const card = el("div", { class: "modal-card", role: "dialog" }, [el("h2", {}, title), body]);
    const row = el("div", { class: "modal-actions" });
    (actions || [{ label: "Close" }]).forEach((action) => {
      const button = el("button", { class: action.primary ? "btn primary" : "btn", type: "button" }, action.label);
      button.addEventListener("click", () => {
        const collected = action.collect ? action.collect() : undefined;
        const captured = action.capture && document.getElementById(action.capture) ? document.getElementById(action.capture).value : undefined;
        closeModal();
        if (action.onClick) action.onClick(action.collect ? collected : captured);
      });
      row.append(button);
    });
    card.append(row);
    modal.replaceChildren(card);
  }

  function closeModal() {
    const modal = document.getElementById("modal");
    modal.hidden = true;
    modal.replaceChildren();
  }

  function render() {
    const main = document.getElementById("main");
    const shouldStick = state.forceStick || (main && main.scrollHeight - main.scrollTop - main.clientHeight < 140);
    renderNav();
    renderTopbar();
    renderMain();
    paintHint();
    if (shouldStick && state.view === "table") document.getElementById("main").scrollTop = document.getElementById("main").scrollHeight;
    state.forceStick = false;
  }

  function renderNav() {
    const nav = document.getElementById("nav");
    const items = [["table", "Table"], ["npcs", "NPCs"], ["lore", "World"], ["session", "Session"], ["settings", "Settings"]];
    nav.replaceChildren(
      el("div", { class: "brand" }, [el("span", { "aria-hidden": "true" }, "◆"), el("span", {}, "Tablelight")]),
      el("p", { class: "brand-sub" }, "Live DM screen")
    );
    items.forEach((item) => {
      const button = el("button", { class: "nav-btn" + (state.view === item[0] ? " active" : ""), type: "button", "data-action": "view", "data-view": item[0] }, [
        el("span", {}, item[1]),
        item[0] === "npcs" ? el("span", {}, String(state.npcs.length)) : null,
      ]);
      if (state.view === item[0]) button.setAttribute("aria-current", "page");
      nav.append(button);
    });
    nav.append(el("p", { class: "nav-note" }, "NPC sheets stay in your Excel file."));
  }

  function renderTopbar() {
    const connected = TL.agents.hasModel(state.settings);
    document.getElementById("topbar").replaceChildren(
      el("div", { class: "file-block" }, [
        el("div", { class: "file-name" }, state.fileName || "campaign.xlsx"),
        el("div", { class: "file-meta", id: "save-state" }, ""),
      ]),
      el("div", { class: "top-actions" }, [
        el("span", { class: connected ? "check" : "file-meta" }, connected ? "Model connected" : "Workbook only"),
        state.needsReconnect ? el("button", { class: "btn primary", type: "button", "data-action": "reconnect" }, "Reconnect") : null,
        el("button", { class: "btn", type: "button", "data-action": "new-workbook" }, "New"),
        el("button", { class: "btn", type: "button", "data-action": "open-workbook" }, "Open"),
        el("button", { class: "btn", type: "button", "data-action": "save-workbook" }, "Save"),
        state.view === "table" && state.messages.length ? el("button", { class: "btn", type: "button", "data-action": "clear-table" }, "Clear") : null,
      ])
    );
    paintSaveBadge();
  }

  function renderMain() {
    const main = document.getElementById("main");
    if (state.view === "npcs") main.replaceChildren(renderNpcs());
    else if (state.view === "lore") main.replaceChildren(renderLore());
    else if (state.view === "session") main.replaceChildren(renderSession());
    else if (state.view === "settings") main.replaceChildren(renderSettings());
    else main.replaceChildren(renderTable());
  }

  function renderTable() {
    const thread = el("div", { class: "thread", id: "thread" });
    if (!state.messages.length) {
      thread.append(
        el("div", { class: "page-head" }, el("div", {}, [
          el("h1", {}, state.campaign.CampaignName || "The screen is clear"),
          el("p", { class: "help" }, "Type the way you would ask a binder. The answer is read back from the notes."),
        ])),
        el("div", { class: "chips" }, [
          state.campaign.CampaignName === "The Drowned Bell" ? null : el("button", { class: "chip", type: "button", "data-action": "use-sample" }, "Load The Drowned Bell"),
          el("button", { class: "chip", type: "button", "data-action": "use-blank" }, "Start a blank campaign"),
        ]),
        el("div", { class: "chips" }, examples().map((text) => el("button", { class: "chip", type: "button", "data-action": "fill", "data-text": text }, text)))
      );
    }
    state.messages.forEach((message) => thread.append(renderMessage(message)));
    if (state.busy) thread.append(el("div", { class: "pending" }, "Checking the notes…"));
    return thread;
  }

  function examples() {
    if (state.campaign && state.campaign.CampaignName === "The Drowned Bell") {
      return [
        "Who is the harbormaster?",
        "Describe the Bell Tower.",
        "What does the Pale Choir want?",
        "Stat block for the Bell Warden.",
        "What quests are open?",
        "We enter the Bell Tower.",
        "Edda Voss took 6 damage.",
      ];
    }
    const name = (state.npcs[0] && state.npcs[0].fields.Name) || "a person on a sheet";
    const place = (state.lore[0] && state.lore[0].title) || "a place in the notes";
    return [
      "Where is " + name + ", and what do they want?",
      "Describe " + place + ".",
      "What quests are open?",
      "What monsters are written down?",
      name + " took 5 damage.",
    ];
  }

  function renderMessage(message) {
    if (message.kind === "user") return el("div", { class: "bubble" }, message.text);
    if (message.kind === "reply") return renderReply(message);
    if (message.kind === "dispatch") {
      return el("div", { class: "dispatch" }, [el("strong", {}, "Orchestrator"), document.createTextNode(message.text)]);
    }
    if (message.kind === "note") return el("div", { class: "note" }, message.text);
    if (message.kind === "npc") return renderNpcCard(message);
    if (message.kind === "lore") return renderLoreCard(message);
    if (message.kind === "proposal") return renderProposal(message);
    if (message.kind === "draft-npc") return renderDraft(message);
    if (message.kind === "party") return renderPartyCard(message);
    if (message.kind === "agent") return renderAgentCard(message);
    return el("div", { class: "note" }, "Unsupported answer.");
  }

  function cardHead(kind, label, aside) {
    return el("div", { class: "card-head" }, [
      el("div", { class: "card-kicker k-" + kind }, label),
      aside ? el("div", { class: "muted-paper" }, aside) : null,
    ]);
  }

  function cardFoot(message, extra) {
    return el("div", { class: "card-foot" }, [
      extra || null,
      el("div", { class: "row" }, [
        el("button", { class: "btn", type: "button", "data-action": "copy-message", "data-id": message.id }, "Copy"),
        message.filed
          ? el("button", { class: "btn", type: "button", disabled: true }, "Filed")
          : el("button", { class: "btn", type: "button", "data-action": "file-message", "data-id": message.id }, "File to session log"),
      ]),
    ]);
  }

  function addFact(parent, label, value) {
    if (!value) return;
    parent.append(el("div", { class: "fact" }, [el("div", { class: "fact-k" }, label), el("div", {}, value)]));
  }

  function renderReply(message) {
    const playing = message.kicker === "At the table";
    const card = el("article", { class: playing ? "room" : "card" });
    const open = message.sheet ? el("button", { class: "btn", type: "button", "data-action": "open-npc", "data-sheet": message.sheet }, "Open sheet") : null;
    if (playing) {
      card.append(el("div", { class: "room-status" }, message.status || "At the table"));
      card.append(el("h2", {}, message.title || "Here"));
      renderRoom(card, message.body);
      card.append(cardFoot(message, open));
      return card;
    }
    const aside = message.sheet ? "Sheet " + message.sheet : "";
    card.append(cardHead("lore", message.kicker || "From the notes", aside));
    card.append(el("h2", {}, message.title || "From the notes"));
    renderRich(card, message.body);
    card.append(cardFoot(message, open));
    return card;
  }

  function fieldLabel(key) {
    if (key === "Race" && state.campaign && /2024/.test(state.campaign.Ruleset || "")) return "Species";
    return LABELS[key] || key;
  }

  function renderNpcCard(message) {
    const npc = message.npc;
    const fields = npc.fields;
    const card = el("article", { class: "card" });
    card.append(cardHead("npc", "On a sheet", "Sheet " + npc.sheet));
    card.append(el("h2", {}, fields.Name || npc.sheet));
    card.append(el("p", { class: "lede" }, [fields.Race, fields.ClassRole].filter(Boolean).join(" · ")));
    if (fields.Concept) card.append(el("p", {}, fields.Concept));
    const stats = el("div", { class: "stats" });
    [
      ["AC", fields.AC || "—"],
      ["HP", (fields.HP || "—") + (fields.HPMax ? "/" + fields.HPMax : "")],
      ["Attitude", [fields.Attitude || "—", fields.AttitudeValue].filter(Boolean).join(" ")],
      ["Status", fields.Status || "—"],
    ].forEach((pair) => stats.append(el("div", { class: "stat" }, [el("b", {}, pair[1]), el("span", {}, pair[0])])));
    card.append(stats);
    const facts = el("div");
    addFact(facts, "Current Location & Status", [fields.Location || "Unknown", fields.Status || "Unknown"].join(" · "));
    addFact(facts, "Attitude towards Party", [fields.Attitude || "Neutral", fields.AttitudeValue].filter((part) => part !== "").join(" · "));
    addFact(facts, "Key Stats/AC/HP", fields.KeyStats);
    addFact(facts, "Voice / Tone", fields.Voice);
    addFact(facts, "Key Motivation", fields.Motivation);
    card.append(facts);
    const recent = (npc.interactions || []).slice(-3);
    card.append(el("h3", {}, "Recent interactions"));
    if (!recent.length) card.append(el("p", { class: "muted-paper" }, "None recorded."));
    recent.forEach((item) => card.append(el("p", {}, "Session " + item.session + " · " + item.date + " · " + item.type + " — " + item.entry)));
    card.append(cardFoot(message, el("button", { class: "btn", type: "button", "data-action": "open-npc", "data-sheet": npc.sheet }, "Open sheet")));
    return card;
  }

  function renderLoreCard(message) {
    const card = el("article", { class: "card" });
    card.append(cardHead("lore", "World & Lore Manager", "From the workbook"));
    message.entries.forEach((entry) => {
      const facts = TL.agents.loreFacts(entry);
      card.append(el("h2", {}, entry.title));
      const wrap = el("div");
      addFact(wrap, "Current Status/Location", facts.status);
      addFact(wrap, "Relationship to Party", facts.relationship);
      addFact(wrap, "Active Goals", facts.goals);
      card.append(wrap);
      const extra = String(entry.body || "").split("\n").map((line) => line.trim()).filter((line) => line && !/^(current status\/location|relationship to party|active goals)\s*:/i.test(line));
      if (extra.length) card.append(el("p", { class: "muted-paper" }, extra.join(" ")));
    });
    card.append(cardFoot(message));
    return card;
  }

  function renderPartyCard(message) {
    const card = el("article", { class: "card" });
    card.append(cardHead("combat", "Combat Builder", "Party on file"));
    card.append(el("p", {}, "Level " + (message.level || "?") + " · size " + (message.size || "?")));
    if (!message.party.length) card.append(el("p", { class: "muted-paper" }, "No adventurers on the roster yet. Add them in Session."));
    message.party.forEach((member) => {
      card.append(el("p", {}, [member.name, member.className, "AC " + member.ac, "HP " + member.hp].filter(Boolean).join(" · ")));
    });
    card.append(cardFoot(message));
    return card;
  }

  function renderAgentCard(message) {
    const card = el("article", { class: "card" });
    card.append(cardHead(message.agent, message.title, message.agent === "rules" ? "Confirm in your book" : ""));
    renderRich(card, message.body);
    card.append(cardFoot(message));
    return card;
  }

  function renderRoom(container, text) {
    String(text || "").split(/\n\n/).forEach((block) => {
      const match = block.trim().match(/^\[\[(place|sense|act|emote|say|roll|next)\]\]([\s\S]*)$/);
      const kind = match ? match[1] : "place";
      const body = (match ? match[2] : block).trim();
      if (!body) return;
      const line = el("p", { class: "room-" + kind });
      appendInline(line, body);
      container.append(line);
    });
  }

  function renderRich(container, text) {
    let list = null;
    String(text || "").split(/\n/).forEach((raw) => {
      const line = raw.trim();
      if (!line) {
        list = null;
        return;
      }
      const labeled = line.match(/^\*\*(.+?)\*\*:?\s*(.*)$/);
      if (labeled) {
        list = null;
        const value = el("div");
        appendInline(value, labeled[2]);
        container.append(el("div", { class: "fact" }, [el("div", { class: "fact-k" }, labeled[1].replace(/:$/, "")), value]));
        return;
      }
      const bullet = line.match(/^[-*•]\s+(.*)$/);
      if (bullet) {
        if (!list) {
          list = el("ul", { class: "rich-list" });
          container.append(list);
        }
        const item = el("li");
        appendInline(item, bullet[1]);
        list.append(item);
        return;
      }
      list = null;
      const paragraph = el("p");
      appendInline(paragraph, line.replace(/^#+\s*/, ""));
      container.append(paragraph);
    });
  }

  function appendInline(node, text) {
    const re = /\*\*(.+?)\*\*/g;
    let last = 0;
    let match;
    while ((match = re.exec(text))) {
      node.append(document.createTextNode(text.slice(last, match.index)));
      node.append(el("strong", {}, match[1]));
      last = match.index + match[0].length;
    }
    node.append(document.createTextNode(text.slice(last)));
  }

  function fillDamageFields(proposal) {
    if (!proposal || !proposal.fields || proposal.fields.HP || !proposal.before) return;
    const damage = TL.agents.parseDamage(proposal.interaction && proposal.interaction.entry);
    if (!damage) return;
    const current = parseInt(proposal.before.HP, 10);
    if (Number.isNaN(current)) return;
    const next = Math.max(0, current - parseInt(damage[1], 10));
    proposal.fields.HP = String(next);
    proposal.fields.Status = TL.rules.healthOf(next, proposal.before.HPMax).label;
    if (!proposal.interaction.type || proposal.interaction.type === "Note") proposal.interaction.type = "Damage";
  }

  function renderProposal(message) {
    if (!message.applied) fillDamageFields(message.proposal);
    const proposal = message.proposal;
    const card = el("article", { class: "card" });
    card.append(cardHead("npc", "Sheet update", message.applied ? "Applied" : "Proposed update"));
    card.append(el("h2", {}, proposal.name));
    if (message.applied) {
      card.append(el("p", {}, "The sheet has this update."));
      card.append(el("div", { class: "card-foot" }, el("button", { class: "btn", type: "button", "data-action": "undo-proposal", "data-id": message.id }, "Undo")));
      return card;
    }
    const keys = Object.keys(proposal.fields);
    if (!keys.length) card.append(el("p", { class: "muted-paper" }, "No stat changes. This files an interaction only."));
    keys.forEach((key) => {
      card.append(el("div", { class: "diff" }, [
        el("div", {}, fieldLabel(key)),
        el("div", { class: "muted-paper" }, String(proposal.before[key] || "—")),
        el("input", { value: proposal.fields[key], "data-proposal-id": message.id, "data-proposal-field": key }),
      ]));
    });
    card.append(el("label", {}, ["Log line", el("textarea", { "data-proposal-id": message.id, "data-proposal-field": "entry" }, proposal.interaction.entry)]));
    const type = el("select", { "data-proposal-id": message.id, "data-proposal-field": "type" });
    INTERACTION_TYPES.forEach((name) => type.append(el("option", { value: name }, name)));
    type.value = proposal.interaction.type || "Note";
    card.append(el("label", {}, ["Type", type]));
    card.append(el("div", { class: "row" }, [
      el("button", { class: "btn primary", type: "button", "data-action": "apply-proposal", "data-id": message.id }, "Apply to sheet"),
      el("button", { class: "btn", type: "button", "data-action": "dismiss-proposal", "data-id": message.id }, "Dismiss"),
    ]));
    return card;
  }

  function renderDraft(message) {
    const card = el("article", { class: "card" });
    card.append(cardHead("npc", "NPC & State Manager", "New sheet"));
    card.append(el("h2", {}, message.fields.Name || "Unnamed NPC"));
    card.append(el("p", {}, [message.fields.Race, message.fields.ClassRole, message.fields.Location].filter(Boolean).join(" · ")));
    card.append(el("button", { class: "btn primary", type: "button", "data-action": "create-draft", "data-id": message.id }, "Create sheet"));
    return card;
  }

  function renderNpcs() {
    if (!selectedNpc() && state.npcs.length) {
      const first = state.npcs.slice().sort((a, b) => String(a.fields.Name).localeCompare(String(b.fields.Name)))[0];
      state.selectedSheet = first.sheet;
    }
    const page = el("section");
    page.append(el("div", { class: "page-head" }, [
      el("div", {}, [el("h1", {}, "NPCs"), el("p", { class: "help" }, "One Excel sheet per person. The tab name drops spaces and punctuation.")]),
      el("button", { class: "btn primary", type: "button", "data-action": "new-npc" }, "New NPC"),
    ]));
    const listHost = el("div", { class: "list", id: "npc-list" });
    renderNpcList(listHost);
    const list = el("div", { class: "pane" }, [
      el("label", {}, ["Find", el("input", { id: "npc-search", value: state.npcQuery, placeholder: "Name, place, role" })]),
      listHost,
    ]);
    page.append(el("div", { class: "split" }, [list, el("div", { id: "npc-detail" }, renderNpcEditor())]));
    return page;
  }

  function renderNpcList(host) {
    const target = host || document.getElementById("npc-list");
    if (!target) return;
    const query = state.npcQuery.trim().toLowerCase();
    const items = state.npcs.filter((npc) => {
      const hay = [npc.fields.Name, npc.fields.Race, npc.fields.ClassRole, npc.fields.Location].join(" ").toLowerCase();
      return !query || hay.includes(query);
    }).sort((a, b) => String(a.fields.Name).localeCompare(String(b.fields.Name)));
    target.replaceChildren();
    if (!items.length) {
      target.append(el("p", { class: "empty" }, query ? "No NPC matches that." : "No NPC sheets yet."));
      return;
    }
    items.forEach((npc) => {
      const button = el("button", { class: "list-btn" + (npc.sheet === state.selectedSheet ? " active" : ""), type: "button", "data-action": "select-npc", "data-sheet": npc.sheet }, [
        el("div", {}, npc.fields.Name || npc.sheet),
        el("small", {}, [npc.fields.Location, npc.fields.Attitude].filter(Boolean).join(" · ")),
      ]);
      target.append(button);
    });
  }

  function choiceSelect(field, options, current) {
    const select = el("select", { "data-npc-field": field });
    if (!current) select.append(el("option", { value: "", selected: true }, "Choose"));
    options.forEach((value) => select.append(el("option", { value, selected: value === current ? true : null }, value)));
    if (current && options.indexOf(current) < 0) select.append(el("option", { value: current, selected: true }, current));
    return select;
  }

  function customField(choice, field, custom) {
    return el("label", { hidden: choice === "Other" ? null : true }, [
      "Other",
      el("input", { "data-npc-field": field, value: custom || "", placeholder: "Type it in" }),
    ]);
  }

  function abilityGrid(fields) {
    return el("div", { class: "abilities wide" }, TL.rules.ABILITIES.map((ability) => {
      return el("label", { class: "ability" }, [
        ability,
        el("input", { "data-npc-field": ability, type: "number", min: "1", max: "30", inputmode: "numeric", value: fields[ability] || "" }),
        el("span", { class: "mod", "data-mod": ability }, TL.rules.abilityMod(fields[ability]) || "—"),
      ]);
    }));
  }

  function healthBadge(fields) {
    const health = TL.rules.healthOf(fields.HP, fields.HPMax);
    return el("div", { id: "health-readout", class: "health" }, health.percent ? health.label + " · " + health.percent + "%" : "Set hit points to calculate health");
  }

  function conditionGrid(fields) {
    const have = String(fields.Conditions || "").split(",").map((part) => part.trim()).filter(Boolean);
    return el("div", { class: "condition-grid" }, TL.rules.CONDITIONS.map((name) => {
      return el("label", {}, [
        el("input", { type: "checkbox", "data-condition": name, value: name, checked: have.indexOf(name) >= 0 ? true : null }),
        name,
      ]);
    }));
  }

  function renderNpcEditor() {
    const npc = selectedNpc();
    if (!npc) return el("div", { class: "pane empty" }, "Select an NPC, or create one.");
    const fields = npc.fields;
    const desired = TL.workbook.toSheetName(fields.Name);
    const tabs = el("div", { class: "sheet-tabs" }, ["sheet", "history", "gear"].map((tab) => {
      const label = tab === "sheet" ? "Sheet" : tab === "history" ? "History" : "Gear";
      return el("button", { class: "chip" + (state.npcTab === tab ? " active" : ""), type: "button", "data-action": "npc-tab", "data-tab": tab }, label);
    }));
    let body;
    if (state.npcTab === "history") body = renderHistory(npc);
    else if (state.npcTab === "gear") body = renderGear(npc);
    else body = renderIdentity(fields);
    return el("div", { class: "pane" }, [
      el("div", { class: "page-head" }, [
        el("div", {}, [
          el("h2", {}, fields.Name || "NPC"),
          el("p", { class: "help", id: "sheet-hint" }, desired === npc.sheet ? "Sheet tab " + npc.sheet : "Sheet tab " + npc.sheet + ". Match tab to rename it " + desired + "."),
        ]),
        el("div", { class: "row" }, [
          desired === npc.sheet ? null : el("button", { class: "btn", type: "button", "data-action": "rename-sheet" }, "Match tab to name"),
          el("button", { class: "btn danger", type: "button", "data-action": "delete-npc" }, "Delete"),
        ]),
      ]),
      tabs,
      body,
    ]);
  }

  function renderIdentity(fields) {
    const grid = el("div", { class: "form-grid", id: "npc-form" });
    grid.append(el("div", { class: "name-row wide" }, [
      el("label", {}, ["Name", el("input", { "data-npc-field": "Name", value: fields.Name })]),
      el("button", { class: "btn", type: "button", "data-action": "suggest-sheet-name" }, "Suggest name"),
    ]));
    grid.append(el("label", {}, ["Species", choiceSelect("Race", TL.rules.SPECIES, fields.Race || "")]));
    grid.append(customField(fields.Race, "RaceCustom", fields.RaceCustom));
    grid.append(el("label", {}, ["Class", choiceSelect("Class", TL.rules.CLASSES, fields.Class || "None")]));
    grid.append(customField(fields.Class, "ClassCustom", fields.ClassCustom));
    grid.append(el("label", {}, ["Role", choiceSelect("Role", TL.rules.ROLES, fields.Role || "None")]));
    grid.append(customField(fields.Role, "RoleCustom", fields.RoleCustom));
    grid.append(el("label", {}, ["Alignment", choiceSelect("Alignment", TL.rules.ALIGNMENTS, fields.Alignment || "Unaligned")]));
    grid.append(el("label", {}, ["Location", el("input", { "data-npc-field": "Location", value: fields.Location })]));
    const attitude = el("select", { "data-npc-field": "Attitude" });
    ATTITUDES.forEach((value) => attitude.append(el("option", { value, selected: value === fields.Attitude ? true : null }, value)));
    grid.append(el("label", {}, ["Attitude", attitude]));
    grid.append(el("label", {}, ["AC", el("input", { "data-npc-field": "AC", value: fields.AC })]));
    grid.append(el("label", {}, ["HP", el("input", { "data-npc-field": "HP", type: "number", min: "0", value: fields.HP })]));
    grid.append(el("label", {}, ["HP max", el("input", { "data-npc-field": "HPMax", type: "number", min: "0", value: fields.HPMax })]));
    grid.append(el("label", {}, ["Health", healthBadge(fields)]));
    grid.append(abilityGrid(fields));
    grid.append(el("p", { class: "help wide" }, "Health is the share of hit points left. It changes when HP changes. A 10 is +0, a 12 is +1, an 8 is -1."));
    grid.append(el("div", { class: "hp-row wide" }, [
      el("span", { class: "help" }, "Quick HP"),
      ["-5", "-1", "+1", "+5"].map((label) => el("button", { class: "btn", type: "button", "data-action": "hp", "data-delta": label }, label)),
    ]));
    grid.append(el("div", { class: "wide" }, [el("p", { class: "help" }, "Conditions"), conditionGrid(fields)]));
    grid.append(el("label", { class: "wide" }, ["Concept", el("textarea", { "data-npc-field": "Concept" }, fields.Concept)]));
    grid.append(el("label", { class: "wide" }, ["Voice / tone", el("textarea", { "data-npc-field": "Voice" }, fields.Voice)]));
    grid.append(el("label", { class: "wide" }, ["Key motivation", el("textarea", { "data-npc-field": "Motivation" }, fields.Motivation)]));
    grid.append(el("label", { class: "wide" }, ["Combat notes", el("textarea", { "data-npc-field": "KeyStats" }, fields.KeyStats)]));
    return grid;
  }

  function renderHistory(npc) {
    const interactions = el("div", { class: "interactions", id: "interaction-list" });
    (npc.interactions || []).forEach((item, index) => {
      interactions.append(el("div", { class: "interaction" }, [
        el("div", {}, "S" + item.session + " · " + item.type),
        el("div", {}, item.entry),
        el("button", { class: "btn", type: "button", "data-action": "delete-interaction", "data-index": String(index) }, "Remove"),
      ]));
    });
    return el("div", {}, [
      el("label", {}, ["Backstory", el("textarea", { "data-npc-field": "History", rows: "8" }, npc.fields.History)]),
      el("h3", {}, "Recorded history"),
      el("p", { class: "help" }, "Each line is stamped with the in-game date and the session number."),
      interactions.childNodes.length ? interactions : el("p", { class: "empty" }, "Nothing has been recorded yet."),
      el("div", { class: "form-grid" }, [
        el("label", {}, ["Type", interactionTypeSelect()]),
        el("label", { class: "wide" }, ["What happened", el("textarea", { id: "interaction-entry", placeholder: "Stamped with the session date on the Session page" })]),
      ]),
      el("button", { class: "btn primary", type: "button", "data-action": "log-interaction" }, "Log to history"),
    ]);
  }

  function renderGear(npc) {
    const list = el("div", { class: "interactions" });
    (npc.equipment || []).forEach((item, index) => {
      list.append(el("div", { class: "interaction" }, [
        el("div", {}, item.qty ? item.qty + "× " + item.item : item.item),
        el("div", {}, item.notes),
        el("button", { class: "btn", type: "button", "data-action": "remove-gear", "data-index": String(index) }, "Remove"),
      ]));
    });
    return el("div", {}, [
      el("p", { class: "help" }, "Gear this person is actually carrying. Use 2024 equipment names."),
      list.childNodes.length ? list : el("p", { class: "empty" }, "No gear listed."),
      el("div", { class: "gear-row" }, [
        el("label", {}, ["Item", el("input", { id: "gear-item", placeholder: "Dagger" })]),
        el("label", {}, ["Qty", el("input", { id: "gear-qty", value: "1" })]),
        el("label", {}, ["Notes", el("input", { id: "gear-notes", placeholder: "Where it is, or why they have it" })]),
        el("button", { class: "btn primary", type: "button", "data-action": "add-gear" }, "Add"),
      ]),
    ]);
  }

  function interactionTypeSelect() {
    const select = el("select", { id: "interaction-type" });
    INTERACTION_TYPES.forEach((name) => select.append(el("option", { value: name }, name)));
    return select;
  }

  function worldMatch(entry) {
    const kind = TL.rules.WORLD_KINDS.find((item) => item.id === state.worldKind);
    if (!kind || kind.id === "all") return true;
    const tags = String(entry.tags || "").toLowerCase().split(/[^a-z]+/);
    return kind.tags.some((tag) => tags.indexOf(tag) >= 0);
  }

  function renderLore() {
    const visible = state.lore.filter(worldMatch);
    if (!visible.some((item) => item.id === state.selectedLore) && visible[0]) state.selectedLore = visible[0].id;
    const entry = state.lore.find((item) => item.id === state.selectedLore);
    const chips = el("div", { class: "kind-row" }, TL.rules.WORLD_KINDS.map((kind) => {
      return el("button", { class: "chip" + (state.worldKind === kind.id ? " active" : ""), type: "button", "data-action": "world-kind", "data-kind": kind.id }, kind.label);
    }));
    const list = el("div", { class: "list" });
    visible.forEach((item) => {
      list.append(el("button", { class: "list-btn" + (item.id === state.selectedLore ? " active" : ""), type: "button", "data-action": "select-lore", "data-id": item.id }, [
        el("div", {}, item.title || "Untitled note"),
        el("small", {}, item.tags),
      ]));
    });
    if (!visible.length) list.append(el("p", { class: "empty" }, "Nothing filed under this heading yet."));
    const editor = entry && worldMatch(entry) ? el("div", { class: "pane", id: "lore-form" }, [
      el("label", {}, ["Title", el("input", { "data-lore-field": "title", value: entry.title })]),
      el("label", {}, ["Tags", el("input", { "data-lore-field": "tags", value: entry.tags, placeholder: "faction, town, nation" })]),
      el("label", {}, ["Note", el("textarea", { "data-lore-field": "body", rows: "12" }, entry.body)]),
      el("p", { class: "help" }, "Factions, settlements, nations, places, quests, and monsters all live in this workbook. A monster's Equipment: line is its gear."),
      el("button", { class: "btn danger", type: "button", "data-action": "delete-lore" }, "Delete note"),
    ]) : el("div", { class: "pane empty" }, "Pick a heading, or add a note.");
    return el("section", {}, [
      el("div", { class: "page-head" }, [
        el("div", {}, [el("h1", {}, "World"), el("p", { class: "help" }, "Factions, villages, cities, nations, places, and the rest of the setting.")]),
        el("button", { class: "btn primary", type: "button", "data-action": "new-lore" }, "New note"),
      ]),
      chips,
      el("div", { class: "split" }, [el("div", { class: "pane" }, list), editor]),
    ]);
  }

  function renderSession() {
    const campaign = state.campaign;
    const grid = el("div", { class: "form-grid", id: "session-form" }, [
      el("label", { class: "wide" }, ["Campaign", el("input", { "data-campaign": "CampaignName", value: campaign.CampaignName })]),
      el("div", { class: "ruleset-card" }, [
        el("span", {}, "Ruleset"),
        el("strong", {}, "D&D 2024"),
        el("p", { class: "help" }, "Species, classes, backgrounds, alignments, and conditions on the NPC sheets are the 2024 lists."),
      ]),
      el("label", {}, ["Party level", el("input", { "data-campaign": "PartyLevel", value: campaign.PartyLevel })]),
      el("label", {}, ["Party size", el("input", { "data-campaign": "PartySize", value: campaign.PartySize })]),
      el("label", {}, ["Session number", el("input", { "data-campaign": "SessionNumber", value: campaign.SessionNumber })]),
      el("label", { class: "wide" }, ["In-game date", el("input", { "data-campaign": "InGameDate", value: campaign.InGameDate, placeholder: "12th of the Grey Tide" })]),
      el("div", { class: "wide" }, [
        el("span", {}, "In-game clock"),
        el("div", { class: "clock-times" }, TL.rules.TIMES.map((time) => {
          return el("button", { class: "chip" + (campaign.InGameTime === time ? " active" : ""), type: "button", "data-action": "set-time", "data-time": time }, time);
        })),
      ]),
      el("label", { class: "wide" }, ["Content boundaries", el("textarea", { "data-campaign": "Boundaries" }, campaign.Boundaries)]),
      el("label", { class: "wide" }, ["House rules", el("textarea", { "data-campaign": "HouseRules" }, campaign.HouseRules)]),
    ]);
    const party = el("div", { id: "party-list" });
    state.party.forEach((member, index) => party.append(partyRow(member, index)));
    const log = el("div", { class: "list" });
    state.log.slice(-5).reverse().forEach((entry) => {
      log.append(el("div", { class: "interaction" }, [
        el("div", {}, entry.agent),
        el("div", {}, entry.query),
      ]));
    });
    return el("section", {}, [
      el("div", { class: "page-head" }, [
        el("div", {}, [el("h1", {}, "Session"), el("p", { class: "help" }, "Boundaries, house rules, and the roster sit on the campaign sheet.")]),
      ]),
      el("div", { class: "pane" }, [grid]),
      el("div", { class: "page-head" }, [
        el("h2", {}, "Party"),
        el("button", { class: "btn", type: "button", "data-action": "add-party" }, "Add adventurer"),
      ]),
      party.childNodes.length ? party : el("p", { class: "empty" }, "No adventurers yet."),
      el("h2", {}, "Recent session log"),
      log.childNodes.length ? log : el("p", { class: "empty" }, "Each answer on the table is written here, on the _SessionLog sheet."),
    ]);
  }

  function partyRow(member, index) {
    const fields = [["name", "Name"], ["className", "Class"], ["level", "Level"], ["ac", "AC"], ["hp", "HP"], ["notes", "Notes"]];
    return el("div", { class: "party-row" }, fields.map((field) => {
      return el("label", {}, [field[1], el("input", { "data-party-index": String(index), "data-party-field": field[0], value: member[field[0]] })]);
    }).concat(el("button", { class: "btn", type: "button", "data-action": "remove-party", "data-index": String(index) }, "Remove")));
  }

  function renderSettings() {
    const preset = el("select", { id: "preset", "data-setting": "preset" });
    PRESETS.forEach((item) => preset.append(el("option", { value: item.id, selected: item.id === state.settings.preset ? true : null }, item.label)));
    return el("section", {}, [
      el("div", { class: "page-head" }, [
        el("div", {}, [
          el("h1", {}, "Settings"),
          el("p", { class: "help" }, "The campaign answers from the workbook. No model is required. A key stays in this browser and is not used to invent the adventure."),
        ]),
      ]),
      el("div", { class: "pane form-grid" }, [
        el("label", {}, ["Provider", preset]),
        el("label", { class: "wide" }, ["Endpoint", el("input", { "data-setting": "baseUrl", value: state.settings.baseUrl, spellcheck: "false" })]),
        el("label", {}, ["Model", el("input", { "data-setting": "model", value: state.settings.model, spellcheck: "false" })]),
        el("label", {}, ["API key", el("input", { "data-setting": "apiKey", value: state.settings.apiKey, type: "password", autocomplete: "off", spellcheck: "false" })]),
        el("div", { class: "row wide" }, [
          el("button", { class: "btn primary", type: "button", "data-action": "test-model" }, "Test connection"),
          el("button", { class: "btn", type: "button", "data-action": "clear-key" }, "Clear key"),
        ]),
        el("p", { id: "model-result", class: "help wide" }, ""),
      ]),
      el("div", { class: "pane" }, [
        el("h2", {}, "How an answer is made"),
        el("p", {}, "You type. The screen finds the people, places, factions, quests, and monsters in the workbook and reads them back."),
        el("p", {}, "Damage, healing, attitude, location, and a line that starts with log: become a proposed sheet update. Nothing is written until you press Apply."),
        el("p", {}, "If the notes do not contain a fact, the reply says so. Page numbers and spell text are not invented."),
        el("p", { class: "help" }, "Chrome and Edge can keep writing to the same .xlsx. Other browsers download a copy when you save."),
      ]),
    ]);
  }

  function paintHint() {
    const hint = document.getElementById("composer-hint");
    const text = document.getElementById("ask-input").value.trim();
    if (!text || !state.campaign) {
      hint.textContent = "Enter sends. Shift+Enter adds a line.";
      return;
    }
    hint.textContent = TL.agents.preview(text, context());
  }

  function messagePlain(message) {
    if (!message) return "";
    if (message.kind === "reply") {
      const body = String(message.body || "").replace(/\[\[(?:place|sense|act|emote|say|roll|next)\]\]/g, "");
      return (message.title ? message.title + "\n" : "") + body;
    }
    if (message.kind === "npc") return TL.agents.npcBlock(message.npc);
    if (message.kind === "agent") return message.title + "\n" + message.body;
    if (message.kind === "lore") {
      return message.entries.map((entry) => {
        const facts = TL.agents.loreFacts(entry);
        return entry.title + "\nStatus: " + facts.status + "\nRelationship: " + facts.relationship + "\nGoals: " + facts.goals;
      }).join("\n\n");
    }
    if (message.kind === "proposal") {
      const lines = Object.keys(message.proposal.fields).map((key) => key + ": " + message.proposal.fields[key]);
      return message.proposal.name + "\n" + lines.join("\n") + "\n" + message.proposal.interaction.entry;
    }
    return message.text || message.title || "";
  }

  function lastQuery() {
    const users = state.messages.filter((message) => message.kind === "user");
    return users.length ? users[users.length - 1].text : "";
  }

  function logLabel(message) {
    if (message.kind === "proposal" && message.proposal) return message.proposal.name || "Sheet";
    return message.title || message.kicker || message.kind;
  }

  function fileMessage(message) {
    const kinds = { reply: 1, proposal: 1, lore: 1, npc: 1, party: 1, agent: 1, "draft-npc": 1 };
    if (!message || message.filed || !state.wb || !kinds[message.kind]) return false;
    const date = [state.campaign.InGameDate, state.campaign.InGameTime].filter(Boolean).join(", ");
    state.log = TL.workbook.appendLog(state.wb, {
      timestamp: new Date().toLocaleString(),
      session: state.campaign.SessionNumber,
      inGameDate: date,
      agent: logLabel(message),
      query: lastQuery(),
      response: messagePlain(message).slice(0, 4000),
    });
    message.filed = true;
    scheduleMirror();
    return true;
  }

  function findMessage(id) {
    const matches = state.messages.filter((message) => message.id === id);
    return matches.length ? matches[matches.length - 1] : null;
  }

  async function submitAsk() {
    const input = document.getElementById("ask-input");
    const text = input.value.trim();
    if (!text || state.busy) return;
    flushNow();
    input.value = "";
    state.messages.push({ id: uid(), kind: "user", text, at: Date.now() });
    state.view = "table";
    state.busy = true;
    state.forceStick = true;
    render();
    document.getElementById("ask-send").disabled = true;
    try {
      const produced = await TL.agents.ask(text, context(), state.settings);
      produced.forEach((message) => {
        const stored = Object.assign({ id: uid(), at: Date.now() }, message);
        state.messages.push(stored);
        fileMessage(stored);
      });
      const dispatch = produced.find((message) => message.kind === "dispatch");
      if (dispatch) announce(dispatch.text);
    } catch (error) {
      state.messages.push({ id: uid(), kind: "note", text: error.message || "The table could not answer.", at: Date.now() });
    } finally {
      state.busy = false;
      state.messages = state.messages.slice(-40);
      saveMessages();
      state.forceStick = true;
      render();
      document.getElementById("ask-send").disabled = false;
      document.getElementById("ask-input").focus();
    }
  }

  function saveMessages() {
    try {
      sessionStorage.setItem("tablelight-messages", JSON.stringify(state.messages));
    } catch (error) {
      /* The table can keep going if the browser refuses the log. */
    }
  }

  function restoreMessages() {
    try {
      const saved = JSON.parse(sessionStorage.getItem("tablelight-messages") || "[]");
      if (Array.isArray(saved)) {
        state.messages = saved;
        saved.forEach((message) => {
          const n = parseInt(String(message && message.id || "").replace(/^m/, ""), 10);
          if (!Number.isNaN(n)) seq = Math.max(seq, n);
        });
      }
    } catch (error) {
      state.messages = [];
    }
  }

  function applyProposal(id) {
    const message = findMessage(id);
    if (message && message.proposal) fillDamageFields(message.proposal);
    const npc = message && state.npcs.find((item) => item.sheet === message.proposal.sheet);
    if (!npc) {
      toast("That NPC sheet is no longer in the workbook.");
      return;
    }
    message.undo = JSON.parse(JSON.stringify(npc));
    Object.keys(message.proposal.fields).forEach((key) => { npc.fields[key] = message.proposal.fields[key]; });
    if (message.proposal.interaction && message.proposal.interaction.entry.trim()) npc.interactions.push(message.proposal.interaction);
    TL.workbook.writeNpc(state.wb, npc);
    message.applied = true;
    saveMessages();
    scheduleMirror();
    render();
    toast("Updated " + npc.fields.Name);
  }

  function undoProposal(id) {
    const message = findMessage(id);
    if (!message || !message.undo) return;
    TL.workbook.writeNpc(state.wb, message.undo);
    const index = state.npcs.findIndex((npc) => npc.sheet === message.undo.sheet);
    if (index >= 0) state.npcs[index] = message.undo;
    message.applied = false;
    const restored = message.undo;
    message.undo = null;
    saveMessages();
    scheduleMirror();
    render();
    toast("Restored " + restored.fields.Name);
  }

  function bumpHp(delta) {
    const npc = selectedNpc();
    const current = parseInt(npc && npc.fields.HP, 10);
    if (!npc || Number.isNaN(current)) {
      toast("Set a numeric HP on the sheet first.");
      return;
    }
    const max = parseInt(npc.fields.HPMax, 10);
    let next = current + delta;
    if (!Number.isNaN(max)) next = Math.min(max, next);
    next = Math.max(0, next);
    npc.fields.HP = String(next);
    npc.fields.Status = TL.rules.healthOf(next, npc.fields.HPMax).label;
    npc.interactions.push({
      date: state.campaign.InGameDate,
      session: state.campaign.SessionNumber,
      type: delta < 0 ? "Damage" : "Healing",
      entry: (delta < 0 ? "Lost " : "Recovered ") + Math.abs(delta) + " HP (" + current + " → " + next + ").",
    });
    TL.workbook.writeNpc(state.wb, npc);
    scheduleMirror();
    render();
  }

  function createNamedNpc(name, extra) {
    const npc = TL.workbook.createNpc(state.wb, Object.assign({ Name: name }, extra || {}));
    hydrate();
    state.selectedSheet = npc.sheet;
    state.npcQuery = "";
    state.view = "npcs";
    scheduleMirror();
    render();
    toast("Created sheet " + npc.sheet);
  }

  async function onAction(button) {
    const action = button.dataset.action;
    if (action === "view") {
      flushNow();
      state.view = button.dataset.view;
      if (state.view === "table") state.forceStick = true;
      render();
      return;
    }
    if (action === "fill") {
      const input = document.getElementById("ask-input");
      input.value = button.dataset.text;
      input.focus();
      paintHint();
      return;
    }
    if (action === "use-sample") return askForNewWorkbook("sample");
    if (action === "use-blank") return askForNewWorkbook("blank");
    if (action === "new-workbook") {
      openModal("New workbook", el("p", {}, "A blank file, or the adventure The Drowned Bell."), [
        { label: "Cancel" },
        { label: "Blank", onClick: () => askForNewWorkbook("blank") },
        { label: "The Drowned Bell", primary: true, onClick: () => askForNewWorkbook("sample") },
      ]);
      return;
    }
    if (action === "open-workbook") {
      try {
        if (canPick) await openWithPicker();
        else document.getElementById("file-input").click();
      } catch (error) {
        if (error && error.name !== "AbortError") toast(error.message || "Could not open that file.");
      }
      return;
    }
    if (action === "save-workbook") return saveWorkbook();
    if (action === "reconnect") return reconnect();
    if (action === "open-npc" || action === "select-npc") {
      flushNow();
      state.selectedSheet = button.dataset.sheet;
      state.view = "npcs";
      render();
      return;
    }
    if (action === "apply-proposal") return applyProposal(button.dataset.id);
    if (action === "undo-proposal") return undoProposal(button.dataset.id);
    if (action === "clear-table") {
      state.messages = [];
      saveMessages();
      render();
      toast("Cleared the table. Sheets are unchanged.");
      return;
    }
    if (action === "dismiss-proposal") {
      state.messages = state.messages.filter((message) => message.id !== button.dataset.id);
      saveMessages();
      render();
      return;
    }
    if (action === "copy-message") {
      const text = messagePlain(findMessage(button.dataset.id));
      try {
        await navigator.clipboard.writeText(text);
        toast("Copied");
      } catch (error) {
        toast("The browser blocked copying.");
      }
      return;
    }
    if (action === "file-message") {
      const message = findMessage(button.dataset.id);
      if (!message || message.filed) {
        toast("Already on the session log.");
        return;
      }
      fileMessage(message);
      saveMessages();
      render();
      toast("Filed on the session log sheet");
      return;
    }
    if (action === "create-draft") {
      const message = findMessage(button.dataset.id);
      if (!message.fields.Name) return toast("That NPC needs a name.");
      createNamedNpc(message.fields.Name);
      const created = selectedNpc();
      if (created) {
        Object.keys(message.fields).forEach((key) => { if (message.fields[key]) created.fields[key] = message.fields[key]; });
        TL.workbook.writeNpc(state.wb, created);
        render();
      }
      return;
    }
    if (action === "new-npc") {
      const race = el("select", { id: "new-race" });
      race.append(el("option", { value: "" }, "Choose"));
      TL.rules.SPECIES.forEach((name) => race.append(el("option", { value: name }, name)));
      const cls = el("select", { id: "new-class" });
      TL.rules.CLASSES.forEach((name) => cls.append(el("option", { value: name }, name)));
      const role = el("select", { id: "new-role" });
      TL.rules.ROLES.forEach((name) => role.append(el("option", { value: name }, name)));
      const alignment = el("select", { id: "new-alignment" });
      TL.rules.ALIGNMENTS.forEach((name) => alignment.append(el("option", { value: name, selected: name === "Unaligned" ? true : null }, name)));
      const form = el("div", { class: "form-grid" }, [
        el("div", { class: "name-row wide" }, [
          el("label", {}, ["Name", el("input", { id: "new-npc-name", placeholder: "A name from the notes, or a suggested one" })]),
          el("button", { class: "btn", type: "button", "data-action": "suggest-name" }, "Suggest name"),
        ]),
        el("label", {}, ["Species", race]),
        el("label", { id: "new-race-wrap", hidden: true }, ["Other species", el("input", { id: "new-race-custom", placeholder: "Type it in" })]),
        el("label", {}, ["Class", cls]),
        el("label", { id: "new-class-wrap", hidden: true }, ["Other class", el("input", { id: "new-class-custom", placeholder: "Type it in" })]),
        el("label", {}, ["Role", role]),
        el("label", { id: "new-role-wrap", hidden: true }, ["Other role", el("input", { id: "new-role-custom", placeholder: "Type it in" })]),
        el("label", {}, ["Alignment", alignment]),
      ]);
      const reveal = (select, wrapId) => {
        select.addEventListener("change", () => {
          document.getElementById(wrapId).hidden = select.value !== "Other";
        });
      };
      reveal(race, "new-race-wrap");
      reveal(cls, "new-class-wrap");
      reveal(role, "new-role-wrap");
      openModal("New NPC", form, [
        { label: "Cancel" },
        { label: "Create sheet", primary: true, collect: () => ({
          Name: document.getElementById("new-npc-name").value.trim(),
          Race: document.getElementById("new-race").value,
          RaceCustom: document.getElementById("new-race-custom").value.trim(),
          Class: document.getElementById("new-class").value,
          ClassCustom: document.getElementById("new-class-custom").value.trim(),
          Role: document.getElementById("new-role").value,
          RoleCustom: document.getElementById("new-role-custom").value.trim(),
          Alignment: document.getElementById("new-alignment").value,
        }), onClick: (fields) => {
          if (!fields.Name) return toast("The NPC needs a name.");
          createNamedNpc(fields.Name, fields);
        } },
      ]);
      setTimeout(() => document.getElementById("new-npc-name").focus(), 0);
      return;
    }
    if (action === "delete-npc") {
      const npc = selectedNpc();
      if (!npc) return;
      openModal("Delete " + (npc.fields.Name || npc.sheet) + "?", el("p", {}, "This removes their sheet from the workbook."), [
        { label: "Cancel" },
        { label: "Delete sheet", primary: true, onClick: () => {
          TL.workbook.deleteNpc(state.wb, npc.sheet);
          state.selectedSheet = "";
          hydrate();
          scheduleMirror();
          render();
        } },
      ]);
      return;
    }
    if (action === "rename-sheet") {
      const npc = selectedNpc();
      const renamed = TL.workbook.renameNpcSheet(state.wb, npc.sheet, npc.fields.Name);
      hydrate();
      state.selectedSheet = renamed.sheet;
      scheduleMirror();
      render();
      return;
    }
    if (action === "hp") return bumpHp(parseInt(button.dataset.delta, 10));
    if (action === "log-interaction") {
      const npc = selectedNpc();
      const entry = document.getElementById("interaction-entry").value.trim();
      if (!entry) return toast("Write what happened first.");
      npc.interactions.push({
        date: state.campaign.InGameDate,
        session: state.campaign.SessionNumber,
        type: document.getElementById("interaction-type").value,
        entry,
      });
      TL.workbook.writeNpc(state.wb, npc);
      scheduleMirror();
      render();
      return;
    }
    if (action === "delete-interaction") {
      const npc = selectedNpc();
      npc.interactions.splice(parseInt(button.dataset.index, 10), 1);
      TL.workbook.writeNpc(state.wb, npc);
      scheduleMirror();
      render();
      return;
    }
    if (action === "npc-tab") {
      state.npcTab = button.dataset.tab || "sheet";
      render();
      return;
    }
    if (action === "world-kind") {
      state.worldKind = button.dataset.kind || "all";
      render();
      return;
    }
    if (action === "set-time") {
      state.campaign.InGameTime = button.dataset.time;
      scheduleMirror();
      render();
      return;
    }
    if (action === "suggest-name") {
      const species = document.getElementById("new-race");
      const input = document.getElementById("new-npc-name");
      if (input) input.value = TL.rules.randomName(species ? species.value : "");
      return;
    }
    if (action === "suggest-sheet-name") {
      const npc = selectedNpc();
      if (!npc) return;
      npc.fields.Name = TL.rules.randomName(npc.fields.Race);
      TL.workbook.writeNpc(state.wb, npc);
      hydrate();
      state.selectedSheet = npc.sheet;
      scheduleMirror();
      render();
      return;
    }
    if (action === "add-gear") {
      const npc = selectedNpc();
      const item = document.getElementById("gear-item").value.trim();
      if (!item) return toast("Name the item first.");
      npc.equipment = npc.equipment || [];
      npc.equipment.push({
        item,
        qty: document.getElementById("gear-qty").value.trim() || "1",
        notes: document.getElementById("gear-notes").value.trim(),
      });
      TL.workbook.writeNpc(state.wb, npc);
      hydrate();
      state.selectedSheet = npc.sheet;
      scheduleMirror();
      render();
      return;
    }
    if (action === "remove-gear") {
      const npc = selectedNpc();
      npc.equipment.splice(parseInt(button.dataset.index, 10), 1);
      TL.workbook.writeNpc(state.wb, npc);
      hydrate();
      state.selectedSheet = npc.sheet;
      scheduleMirror();
      render();
      return;
    }
    if (action === "new-lore") {
      const kind = TL.rules.WORLD_KINDS.find((item) => item.id === state.worldKind);
      const tag = kind && kind.tags ? kind.tags[0] : "world";
      const entry = {
        id: TL.workbook.nextLoreId(state.lore),
        title: "New note",
        tags: state.worldKind === "all" ? "world" : tag,
        body: "Summary: \nCurrent status/location: \nRelationship to party: \nActive goals: \nFor the DM: \n",
        updated: "Session " + state.campaign.SessionNumber,
      };
      state.lore.push(entry);
      state.selectedLore = entry.id;
      TL.workbook.saveLore(state.wb, state.lore);
      scheduleMirror();
      render();
      return;
    }
    if (action === "select-lore") {
      flushNow();
      state.selectedLore = button.dataset.id;
      render();
      return;
    }
    if (action === "delete-lore") {
      state.lore = state.lore.filter((entry) => entry.id !== state.selectedLore);
      state.selectedLore = state.lore[0] ? state.lore[0].id : "";
      TL.workbook.saveLore(state.wb, state.lore);
      scheduleMirror();
      render();
      return;
    }
    if (action === "add-party") {
      state.party.push({ name: "", className: "", level: state.campaign.PartyLevel, ac: "", hp: "", notes: "" });
      scheduleMirror();
      render();
      return;
    }
    if (action === "remove-party") {
      state.party.splice(parseInt(button.dataset.index, 10), 1);
      TL.workbook.saveParty(state.wb, state.party);
      scheduleMirror();
      render();
      return;
    }
    if (action === "test-model") {
      const result = document.getElementById("model-result");
      result.textContent = "Contacting the model…";
      try {
        const text = await TL.agents.ping(state.settings);
        result.textContent = "Connected. The model said: " + String(text).trim().slice(0, 80);
      } catch (error) {
        result.textContent = error.code === "NO_KEY" || error.message === "NO_KEY" ? "Add an API key first." : error.message;
      }
      return;
    }
    if (action === "clear-key") {
      state.settings.apiKey = "";
      saveSettings();
      render();
      toast("API key cleared");
    }
  }

  function onInput(event) {
    const target = event.target;
    if (target.dataset.campaign) {
      state.campaign[target.dataset.campaign] = target.value;
      scheduleMirror();
    }
    if (target.dataset.condition) {
      const npc = selectedNpc();
      if (!npc) return;
      npc.fields.Conditions = [...document.querySelectorAll("[data-condition]:checked")].map((box) => box.value).join(", ");
      scheduleMirror();
      return;
    }
    if (target.dataset.npcField) {
      const npc = selectedNpc();
      if (!npc) return;
      const key = target.dataset.npcField;
      npc.fields[key] = target.value;
      if (key === "HP" || key === "HPMax") {
        const health = TL.rules.healthOf(npc.fields.HP, npc.fields.HPMax);
        npc.fields.Status = health.label;
        const badge = document.getElementById("health-readout");
        if (badge) badge.textContent = health.percent ? health.label + " · " + health.percent + "%" : "Set hit points to calculate health";
      }
      if (TL.rules.ABILITIES.indexOf(key) >= 0) {
        const mod = document.querySelector('[data-mod="' + key + '"]');
        if (mod) mod.textContent = TL.rules.abilityMod(target.value) || "—";
      }
      if (key === "Race" || key === "Class" || key === "Role" || key === "RaceCustom" || key === "ClassCustom" || key === "RoleCustom") {
        const jobClass = TL.rules.shown(npc.fields.Class, npc.fields.ClassCustom);
        const jobRole = TL.rules.shown(npc.fields.Role, npc.fields.RoleCustom);
        npc.fields.ClassRole = [jobClass, jobRole].filter(Boolean).join(" · ");
        if (key === "Race" || key === "Class" || key === "Role") {
          scheduleMirror();
          render();
          return;
        }
      }
      const hint = document.getElementById("sheet-hint");
      if (hint) {
        const desired = TL.workbook.toSheetName(npc.fields.Name);
        hint.textContent = desired === npc.sheet ? "Sheet tab " + npc.sheet : "Sheet tab " + npc.sheet + ". Match tab to rename it " + desired + ".";
      }
      scheduleMirror();
    }
    if (target.dataset.loreField) {
      const entry = state.lore.find((item) => item.id === state.selectedLore);
      if (!entry) return;
      entry[target.dataset.loreField] = target.value;
      entry.updated = "Session " + state.campaign.SessionNumber;
      scheduleMirror();
    }
    if (target.dataset.partyField) {
      const member = state.party[parseInt(target.dataset.partyIndex, 10)];
      if (!member) return;
      member[target.dataset.partyField] = target.value;
      scheduleMirror();
    }
    if (target.dataset.setting) {
      state.settings[target.dataset.setting] = target.value;
      if (target.dataset.setting === "preset") {
        const chosen = PRESETS.find((item) => item.id === target.value);
        if (chosen && chosen.id !== "custom") {
          state.settings.baseUrl = chosen.baseUrl;
          state.settings.model = chosen.model;
        }
        saveSettings();
        render();
        return;
      }
      saveSettings();
      paintSaveBadge();
      renderTopbar();
    }
    if (target.dataset.proposalField) {
      const message = findMessage(target.dataset.proposalId);
      if (!message) return;
      if (target.dataset.proposalField === "entry") message.proposal.interaction.entry = target.value;
      else if (target.dataset.proposalField === "type") message.proposal.interaction.type = target.value;
      else message.proposal.fields[target.dataset.proposalField] = target.value;
      saveMessages();
    }
    if (target.id === "npc-search") {
      state.npcQuery = target.value;
      renderNpcList();
    }
  }

  async function boot() {
    if (!window.XLSX || !TL || !TL.workbook) {
      document.getElementById("main").textContent = "Tablelight could not load its libraries. Serve this folder from a local web server.";
      return;
    }
    restoreMessages();
    state.wb = TL.workbook.createBlank();
    try {
      const snapshot = await idbGet("snapshot");
      if (snapshot && snapshot.bytes) {
        state.wb = TL.workbook.fromArrayBuffer(snapshot.bytes);
        state.fileName = snapshot.name || state.fileName;
        state.dirty = Boolean(snapshot.dirty);
      }
      const handle = await idbGet("fileHandle");
      if (handle) {
        state.fileHandle = handle;
        state.fileName = handle.name || state.fileName;
        let perm = "prompt";
        if (handle.queryPermission) perm = await handle.queryPermission({ mode: "readwrite" });
        if (perm === "granted") {
          state.needsReconnect = false;
          if (state.dirty) await writeHandle();
          else {
            const file = await handle.getFile();
            state.wb = TL.workbook.fromArrayBuffer(await file.arrayBuffer());
            state.fileName = file.name;
          }
        } else state.needsReconnect = true;
      }
    } catch (error) {
      state.needsReconnect = Boolean(state.fileHandle);
    }
    hydrate();
    render();
    document.getElementById("ask-input").focus();
  }

  function init() {
    document.body.addEventListener("click", (event) => {
      const button = event.target.closest("[data-action]");
      if (button) Promise.resolve(onAction(button)).catch((error) => toast(error.message || "Something went wrong."));
      if (event.target.id === "modal") closeModal();
    });
    document.body.addEventListener("input", onInput);
    document.getElementById("composer").addEventListener("submit", (event) => {
      event.preventDefault();
      submitAsk();
    });
    document.getElementById("ask-input").addEventListener("keydown", (event) => {
      if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        submitAsk();
      }
    });
    document.getElementById("ask-input").addEventListener("input", paintHint);
    document.getElementById("file-input").addEventListener("change", async (event) => {
      const file = event.target.files && event.target.files[0];
      event.target.value = "";
      if (!file) return;
      const wb = TL.workbook.fromArrayBuffer(await file.arrayBuffer());
      replaceWorkbook(wb, file.name, null);
      await rememberHandle(null);
      toast("Opened " + file.name + " in this browser");
    });
    window.addEventListener("dragover", (event) => event.preventDefault());
    window.addEventListener("drop", async (event) => {
      event.preventDefault();
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (!file || !/\.xlsx$/i.test(file.name)) return;
      const wb = TL.workbook.fromArrayBuffer(await file.arrayBuffer());
      replaceWorkbook(wb, file.name, null);
      await rememberHandle(null);
      toast("Opened " + file.name);
    });
    boot();
  }

  init();
})();
