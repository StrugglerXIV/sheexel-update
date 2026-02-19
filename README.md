# Sheexcel Updated (Foundry VTT)

Sheexcel Updated links Google Sheets to Foundry VTT actor sheets. It reads data from a Google Sheet and renders checks, saves, attacks, spells, gear, armor, and stats in a custom sheet. Optional write-back lets you edit HP/Vitality and some stats from Foundry.

## Features

- Read data from a Google Sheet into a custom Foundry actor sheet
- Checks/saves/attacks/spells/gears auto-scan from the Core sheet and other tabs
- HP and Vitality orbs with live values
- Armor and Stats summary blocks
- Optional write-back to Google Sheets for HP/Vitality and selected stats

## Installation (GitHub)

1. Download or clone this repository.
2. Copy the folder to your Foundry data directory:
   - `FoundryVTT/Data/modules/sheexcel_updated`
3. Enable the module in Foundry: **Setup** -> **Add-on Modules** -> **Sheexcel Updated**.

## Required Setup

### 1) Google Sheets API Key (Read)

This module reads from Google Sheets using an API key.

1. Create a project in Google Cloud Console.
2. Enable **Google Sheets API**.
3. Create an **API key**.
4. In Foundry: **Module Settings** -> **Sheexcel Updated** -> paste the API key.

### 2) Sheet URL

Open the actor sheet and enter the Google Sheet URL in the Configuration tab, then click **Update Sheet**.

## Optional: Write-back to Google Sheets (Edit HP/Vitality/Stats)

To write back, you must use OAuth (Google does not allow writes with API key alone).

1. Create an OAuth Client ID (Web application).
2. Add your Foundry URL to **Authorized JavaScript origins**.
3. In Foundry: **Module Settings** -> **Google OAuth Client ID** -> paste it.
4. Click the HP or Vitality orb (or editable stat rows) and approve the consent prompt.

Notes:
- This uses Google Identity Services and the `spreadsheets` scope.
- If the OAuth app is in Testing, add your Google account as a test user.

## Sheet Layout Expectations

The module scans the Core sheet and looks for specific headers and label/value patterns.

### Core sheet

- **Checks**: Scans a rectangle near Athletics -> Languages and groups adjacent rows as subchecks.
- **Saves**: Scans the Core sheet for Save-style labels.
- **Attacks**: Scans attack blocks that contain `Range`, `Accuracy`, `Critical`, `Damage`, `Special`.
- **HP/Vitality**: Finds the `Health`/`Vitality` row and reads `Base` and `Tot` columns.
- **Armor**: Finds the `Armor` header and reads rows below it. Values are embedded in the same cell (e.g., `4 Slash`).
- **Stats**: Finds the `Stats` header and reads rows below it through `Exhaustion`.
  - Base and Tot are detected in the header row (or the row below).
  - Health and Vitality are skipped here (they live in the orbs).

If your sheet has a different layout, adjust the labels to match the expected headers, or update the scan ranges in code.

## Using the Sheet

- **Update Sheet**: pulls sheet metadata and stores sheet info.
- **Update Skills/Saves/Attacks/Spells/Gears**: scans the sheet and rebuilds references.
- **Update All**: runs all scans.
- **Armor/Stats Refresh**: click the refresh icon on each card to re-read values.
- **HP/Vitality edit**: click the orb, enter a value or `+/-` adjustment.

## Notes on Caching

- Armor and Stats values are cached on the actor once loaded. They only update when you click Refresh.
- Checks/attacks/spells/gears update when you click their respective Update buttons.

## Troubleshooting

- **429 Too Many Requests**: reduce refresh frequency and use manual refresh buttons.
- **OAuth access_denied**: add your account as a test user or publish the OAuth app.
- **Values missing**: verify the Core sheet headers (`Stats`, `Armor`, `Base`, `Tot`) and label text.

## License

See [module.json](module.json) for module metadata. Add a license file if you plan to distribute publicly.
