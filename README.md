# Tablelight

A live dungeon-master screen you can open on your computer or publish with GitHub Pages. You type a question. Tablelight reads the answer back from the campaign workbook on your machine: people, places, factions, quests, and monsters. It does not invent a fact that is not written down.

The workbook is not uploaded with the site. No model is required. Load The Drowned Bell from the empty screen, or open your own `.xlsx`. A model key in Settings is unused for these answers and stays in this browser.

## Run it locally

From this folder:

```powershell
py -m http.server 8080
```

Open http://localhost:8080 in Chrome or Edge. Those browsers can keep saving into the same `.xlsx`. Other browsers can open a workbook and download a copy when you press Save.

## Put it on GitHub Pages

1. Create a repository and push this folder.
2. In the repository, open Settings → Pages.
3. Deploy from the `main` branch, folder `/ (root)`.
4. Open the Pages URL in Chrome or Edge.

Use a project site (`https://you.github.io/tablelight/`) or a user site. Relative links are already set for both. Do not commit your campaign `.xlsx`. It is ignored on purpose.

## How a question is answered

One typed line, one reply, taken from the notes.

- A person comes back from that person's sheet: where they are, what they want, voice, AC, HP, and the last few interactions.
- A place comes back as the read-aloud, the senses, and the DM note stored on that place.
- A faction, quest, or item comes back with the status, relationship, and goals written on it.
- A monster comes back as the stat block stored in the notes. Those blocks are original and shaped for the 2024 rules. Spell text and page numbers are not copied, and they are not invented.
- "What quests are open?" and "what monsters are in the tower?" list the matching notes.
- Damage, healing, attitude, location, and a line that starts with `log:` become a proposed sheet update. Nothing is written until you press Apply.

The line under the ask box names the note it is about to open.

The Drowned Bell is a level 5 adventure for four adventurers, set in the port of Saltwick. Load it, then ask who the harbormaster is.

## Workbook layout

- `Edda_Voss` and other NPC tabs: column A is the field name, column B is the value, then an `INTERACTIONS` table.
- Tab names are the NPC name with spaces and punctuation turned into underscores.
- `_Campaign` holds the ruleset, party level, session, boundaries, and house rules.
- `_Lore` holds world notes. Start status, relationship, and goals lines with those labels so they can be read back.
- `_Party` is the adventurer roster.
- `_Index` is rebuilt automatically.
- `_SessionLog` stores answers you file from a card.

You can edit the file in Excel while the site is closed, then open it again.
