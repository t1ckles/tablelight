const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
global.XLSX = require(path.join(root, "vendor", "xlsx.full.min.js"));
eval(fs.readFileSync(path.join(root, "js", "campaign.js"), "utf8"));
eval(fs.readFileSync(path.join(root, "js", "ruleset.js"), "utf8"));
eval(fs.readFileSync(path.join(root, "js", "workbook.js"), "utf8"));
eval(fs.readFileSync(path.join(root, "js", "agents.js"), "utf8"));

const wbApi = global.Tablelight.workbook;
const agents = global.Tablelight.agents;
let failed = 0;

function assert(condition, message) {
  if (!condition) {
    failed += 1;
    console.error("FAIL " + message);
  }
}

assert(wbApi.toSheetName("Gundren Rockseeker") === "Gundren_Rockseeker", "sheet names use underscores");
assert(wbApi.toSheetName("Vess (Kade)!") === "Vess_Kade", "punctuation collapses");

const wb = wbApi.createSample();
const first = wbApi.read(wb);
assert(first.campaign.CampaignName === "The Drowned Bell", "campaign name");
assert(first.campaign.Ruleset === "D&D 2024", "2024 ruleset");
assert(first.npcs.length === 8, "eight people");
assert(first.npcs.map((npc) => npc.sheet).join(",") === "Edda_Voss,Brother_Calder,Ness_Quinn,Maelor_Pye,Sister_Wynn,Ivo_Brask,Old_Pett,Pilot_Sarn", "sheet names");
const edda = first.npcs.find((npc) => npc.sheet === "Edda_Voss");
assert(edda.fields.HP === "32" && edda.fields.Location.indexOf("Tide Office") >= 0, "Edda fields");
assert(edda.interactions.length === 1, "Edda interactions");
assert(first.lore.some((entry) => entry.title === "Pale Choir"), "faction stored");
assert(first.lore.some((entry) => entry.title === "Silence the Bell"), "quest stored");
assert(first.lore.some((entry) => entry.title === "Bell Warden"), "monster stored");
assert(!first.npcs.some((npc) => npc.sheet.startsWith("_") || npc.sheet === "ReadMe"), "system sheets are hidden");
assert(!JSON.stringify(first.lore).match(/page\s+\d+/i), "no invented page numbers");

const ctx = wbApi.read(wb);
const noKey = { apiKey: "", baseUrl: "", model: "" };

const damage = agents.route("Edda Voss took 6 damage.", ctx);
assert(damage.write && damage.npcs.length === 1 && damage.npcs[0].sheet === "Edda_Voss", "damage stays on Edda");
assert(damage.proposals[0].fields.HP === "26", "HP reduced from 32");
assert(damage.proposals[0].fields.Status === "Bruised", "26 of 32 hit points is bruised");
const takes = agents.route("Edda Voss takes 2 damage.", ctx);
assert(takes.write && takes.proposals[0].fields.HP === "30" && takes.proposals[0].interaction.type === "Damage", "takes 2 damage proposes 30 HP");
assert(takes.proposals[0].fields.Status === "Bruised", "30 of 32 hit points is bruised");
const eddaNow = ctx.npcs.find((npc) => npc.sheet === "Edda_Voss");
eddaNow.fields.HP = "30";
const heal = agents.route("Edda Voss heals 2 damage.", ctx);
assert(heal.write && heal.proposals[0].interaction.type === "Healing" && heal.proposals[0].interaction.entry === "Healed 2 HP.", "heals is healing");
assert(heal.proposals[0].fields.HP === "32" && heal.proposals[0].fields.Status === "Healthy", "heals 2 damage restores hit points");
const allHeal = agents.route("Edda heals all damage.", ctx);
assert(allHeal.write && allHeal.proposals[0].interaction.type === "Healing", "heals all damage is a sheet update");
assert(allHeal.proposals[0].fields.HP === "32" && allHeal.proposals[0].fields.Status === "Healthy", "heals all damage restores to maximum");
assert(allHeal.proposals[0].interaction.entry === "Healed all damage (2 HP).", "all-damage heal names the amount");
const allHurt = agents.route("Edda takes all damage.", ctx);
assert(allHurt.write && allHurt.proposals[0].fields.HP === "0" && allHurt.proposals[0].fields.Status === "Down", "takes all damage drops to zero");
eddaNow.fields.HP = "32";
assert(global.Tablelight.rules.abilityMod(12) === "+1" && global.Tablelight.rules.abilityMod(8) === "-1" && global.Tablelight.rules.abilityMod(10) === "+0", "ability modifiers");
assert(edda.equipment.some((item) => item.item === "Dagger"), "Edda carries a dagger");
assert(edda.fields.STR === "14" && edda.fields.Alignment === "Neutral", "Edda scores and alignment");

const attitude = agents.route("What is Brother Calder's attitude?", ctx);
assert(!attitude.write, "asking attitude is not an update");

const draft = agents.route("new npc: Sera Vale, elf, bard, the docks", ctx);
assert(draft.draft && draft.draft.Name === "Sera Vale" && draft.draft.Location === "the docks", "new npc draft");

const live = wbApi.read(wb);
const eddaLive = live.npcs.find((npc) => npc.sheet === "Edda_Voss");
eddaLive.fields.HP = "26";
eddaLive.fields.Status = "Injured";
eddaLive.interactions.push({ date: "12th of the Grey Tide", session: "1", type: "Damage", entry: "Took 6 damage." });
wbApi.writeNpc(wb, eddaLive);
const again = wbApi.read(wbApi.fromArrayBuffer(wbApi.toArrayBuffer(wb)));
const edda2 = again.npcs.find((npc) => npc.sheet === "Edda_Voss");
assert(edda2.fields.HP === "26" && edda2.fields.Status === "Bruised", "damage round-trip keeps the calculated health");
assert(edda2.interactions.length === 2, "interaction appended");
assert(again.lore.find((entry) => entry.title === "Saltwick").body.indexOf("Read-aloud:") >= 0, "lore newlines survive");
assert(wb.SheetNames.indexOf("Edda_Voss") < wb.SheetNames.indexOf("_Index"), "NPC tabs stay before the index");

const created = wbApi.createNpc(wb, { Name: "Gundren Rockseeker", Race: "Dwarf" });
assert(created.sheet === "Gundren_Rockseeker", "created tab name");
const duplicate = wbApi.createNpc(wb, { Name: "Gundren Rockseeker" });
assert(duplicate.sheet !== created.sheet, "duplicate tabs are unique");

function bodyOf(messages) {
  const reply = messages.find((message) => message.kind === "reply");
  return reply ? reply.title + "\n" + reply.body : "";
}

Promise.all([
  agents.ask("Who is the harbormaster?", ctx, noKey),
  agents.ask("Describe the Bell Tower.", ctx, noKey),
  agents.ask("What does the Pale Choir want?", ctx, noKey),
  agents.ask("Stat block for the Bell Warden.", ctx, noKey),
  agents.ask("What quests are open?", ctx, noKey),
  agents.ask("What monsters are in the bell tower?", ctx, noKey),
  agents.ask("Where is the clapper?", ctx, noKey),
  agents.ask("How does grappling work?", ctx, noKey),
  agents.ask("Who is Gundren?", ctx, noKey),
  agents.ask("Edda Voss took 6 damage.", ctx, noKey),
]).then((results) => {
  const [harbor, tower, choir, warden, quests, monsters, clapper, rules, missing, hit] = results.map(bodyOf);
  assert(results[0].length === 1 && results[0][0].kind === "reply", "recall is one reply");
  assert(results[0][0].sheet === "Edda_Voss", "harbormaster opens Edda's sheet");
  assert(harbor.indexOf("Tide Office") >= 0 && harbor.indexOf("Edda Voss") >= 0, "harbormaster is Edda");
  assert(!/model is connected|specialist/i.test(harbor), "recall does not ask for a model");
  assert(tower.indexOf("empty hook") >= 0 && tower.indexOf("Wet bronze") >= 0, "tower read-aloud and senses");
  assert(choir.indexOf("stands") >= 0 || choir.indexOf("stand") >= 0, "choir goal");
  assert(warden.indexOf("AC 16") >= 0 && warden.indexOf("CR 5") >= 0 && warden.indexOf("Bell-hammer") >= 0, "warden stat block");
  assert(quests.indexOf("Silence the Bell") >= 0 && quests.indexOf("The Missing Pilots") >= 0 && quests.indexOf("The Red Ledger") >= 0, "open quests");
  assert(monsters.indexOf("Bell Warden") >= 0 && monsters.indexOf("Drowned Chorister") >= 0, "tower monsters");
  assert(monsters.indexOf("Salt Wretch") < 0, "wharf monster stays off the tower list");
  assert(clapper.indexOf("Throat") >= 0, "clapper is in the Throat");
  assert(rules.indexOf("2024 Player's Handbook") >= 0 && !/page\s+\d+/i.test(rules), "rules point at the book without a page");
  assert(missing.indexOf("Not in the notes") >= 0, "unknown name is refused");
  assert(results[9].some((message) => message.kind === "proposal" && message.proposal.fields.HP === "26"), "damage proposes a sheet update");
  assert(!results[9].some((message) => message.kind === "dispatch" || message.kind === "agent"), "an update is not an agent");
  if (failed) {
    console.error(failed + " failed");
    process.exit(1);
  }
  console.log("roundtrip ok");
}).catch((error) => {
  console.error(error);
  process.exit(1);
});
