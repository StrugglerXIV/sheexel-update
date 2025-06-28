# Sheexcel-Updated

**Excel in your Sheets** – A Foundry VTT module that lets you map cells from a Google Sheet directly into your actor sheet, then roll checks, saves, attacks and spells with live values.

---

## 📖 Overview

Sheexcel-Updated embeds a Google Sheet in your character’s sheet sidebar and lets you:

- **Map** individual cells to “keywords” (e.g. “Athletics”, “Perception”, “Fireball Damage”).
- **Group** them into four subtabs: **Checks**, **Saves**, **Attacks**, **Spells**.
- **Roll** straight from your sheet via clickable buttons, with configurable Advantage/Normal/Disadvantage.
- **Import/Export** your entire cell-to-keyword mapping as JSON, so you can share presets.
- **Batch-fetch** dozens of cells in one API call to minimize rate limits.
- **Customize Attacks** with extra fields:  
  • Name cell • Damage formula cell • Crit range cell  
  – Crit damage is calculated by doubling positive die results and positive modifiers only.
- **Situational Bonuses** via a prompt, anywhere you roll.

---

## 🚀 Features

1. **Live Embedding**  
   Display your Google Sheet (with or without menus) right inside Foundry.

2. **Dynamic References**  
   Define any number of “references” by cell address, sheet name and keyword.  
   └ Mapped values appear in both the References tab and in the Main view.

3. **Four Subtabs**  
   - **Checks** – Skill checks  
   - **Saves** – Saving throws  
   - **Attacks** – Attack rolls (with name, damage & crit range)  
   - **Spells** – Spell DCs or other spell‐based values

4. **Click-to-Roll**  
   Buttons on each keyword roll a d20 (with mod), optionally with advantage/disadvantage.

5. **Critical Damage**  
   Attack rolls detect crits (≥ your crit range) and auto-roll & post doubled damage.

6. **JSON Import/Export**  
   Map hundreds of references in one go by uploading a JSON file.

7. **Search Bar**  
   Quickly filter the visible keywords in the Main tab.

8. **Modular Helpers**  
   Organized into ES modules under `helpers/` for data preparation, JSON import, batch fetching, rolling, situational prompts, etc.

---

## 📦 Installation

1. Clone or download this repository to your Foundry **modules** folder:  
FoundryVTT/
└─ modules/
└─ sheexcel_updated/
├─ module.json
├─ scripts/
├─ helpers/
└─ templates/
2. In Foundry’s **Manage Modules**, enable **Sheexcel-Updated**.
3. Open any Actor of type **Character**, **NPC**, **Creature** or **Vehicle** and switch its sheet to “Sheexcel”.

---

## ⚙️ Configuration

1. In the **Settings** tab of your actor sheet, paste a Google Sheet URL and click **Update Sheet**.  
2. Grant the sheet at least “Viewer” access to your API key (you can supply your own in `batchFetcher.js`).

---

## 🎲 Usage

1. **Add References**  
- Go to **References** → click **Add Reference**  
- Enter:  
  - **Cell** (e.g. `D17`)  
  - **Keyword** (e.g. `Athletics`)  
  - **Sheet** (e.g. `Core`)  
  - **Type**: checks/saves/attacks/spells  

2. **Save References**  
- Click **Save References** to batch-fetch all values.  
- Values populate in the **Main** → **Checks** (or whichever subtab).

3. **Roll**  
- Toggle Advantage/Normal/Disadvantage at top of **Main**.  
- Click any keyword button to roll.  
- For **Attacks**, you’ll get a crit check, plus an immediate damage roll.

4. **Import/Export JSON**  
- Click **Import JSON…** in **References** to bulk-load mappings.  
- Paste back your exported JSON to replicate your setup elsewhere.

---

## 🗂️ File Structure

sheexcel_updated/
├ module.json
├ scripts/
│ └ sheexcel.js # Main sheet class & hooks
├ helpers/
│ ├ prepareData.js # Merge Foundry’s data + flags
│ ├ importer.js # JSON import/export
│ ├ batchFetcher.js # Google Sheets batch API calls
│ ├ roller.js # d20 + damage + situational bonus logic
│ └ situational.js # Dialog prompt for extra bonuses
├ templates/
│ ├ sheet-template.html # Handlebars main template
│ └ partials/ # Tab partials (main-tab.hbs, references-tab.hbs, settings-tab.hbs)
└ styles/
└ sheexcel.css

---

## 🤝 Contributing

- Feel free to open issues or PRs to fix bugs, improve performance, or add new features!  
- Keep helpers small and focused.  
- Maintain Foundry v11+ compatibility.

---

## 📜 License

MIT © Struggler 
