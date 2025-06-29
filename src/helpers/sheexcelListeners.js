import { importJsonHandler, exportJsonHandler } from "./importer.js";
import { handleRoll } from "./roller.js";
import { onSearch } from "./mainSearch.js";
import { nestCheck, moveCheckToRoot } from "./dragDrop.js";
import { onSaveReferences } from "./buttons/saveButton.js";

export function attachSheexcelListeners(sheet, html) {
    html.find('.sheexcel-profile-pic').click(ev => {
        const current = sheet.actor.img;
        new FilePicker({
            type: "image",
            current,
            callback: path => sheet.actor.update({ img: path })
        }).render(true);
    });

    html.find('.sheexcel-check-entry[draggable=true]')
        .on('dragstart', ev => {
            ev.stopPropagation(); // <--- Add this line
            const payload = {
                checkId: ev.currentTarget.dataset.checkId,
                subType: 'sheexcel-check'
            };
            console.log("Payload for dragstart:", payload);
            ev.originalEvent.dataTransfer.setData('text/plain', JSON.stringify(payload));
        })
        .on('dragover', ev => ev.preventDefault())
        .on('drop', ev => {
            ev.preventDefault();
            let payload;
            try {
                payload = JSON.parse(ev.originalEvent.dataTransfer.getData('text/plain'));
            } catch { return; }
            if (!payload?.checkId) return;
            const draggedId = payload.checkId;
            const targetId = ev.currentTarget.dataset.checkId;
            nestCheck(sheet, draggedId, targetId);
        });

    html.find('.sheexcel-check-dropzone')
        .on('drop', ev => {
            ev.preventDefault();
            let payload;
            try {
                payload = JSON.parse(ev.originalEvent.dataTransfer.getData('text/plain'));
            } catch { return; }
            if (!payload?.checkId) return;
            moveCheckToRoot(sheet, payload.checkId);
        });

    html.find('.sheexcel-collapse-toggle').click(ev => {
        const entry = $(ev.currentTarget).closest('.sheexcel-check-entry');
        entry.find('.sheexcel-subchecks').toggle();
        $(ev.currentTarget).text(
            entry.find('.sheexcel-subchecks').is(':visible') ? '[–]' : '[+]'
        );
    });

    html.find(".sheexcel-sheet-toggle").on("click", sheet._onToggleSidebar.bind(sheet));
    html.find(".sheexcel-sheet-main, .sheexcel-sheet-references, .sheexcel-sheet-settings")
        .on("click", sheet._onToggleTab.bind(sheet));
    html.find(".sheexcel-setting-update-sheet").on("click", sheet._onUpdateSheet.bind(sheet));
    html.find('.sheexcel-search').on('input', onSearch.bind(sheet));
    html.on("click", ".sheexcel-reference-add-button", sheet._onAddReference.bind(sheet));
    html.on("click", ".sheexcel-reference-remove-button", sheet._onRemoveReference.bind(sheet));
    html.on("click", ".sheexcel-reference-remove-save", sheet._onFetchAndUpdateCellValueByIndex.bind(sheet));
    html.on("click", ".sheexcel-reference-save-button", (event) => onSaveReferences(sheet, event));
    html.find(".sheexcel-import-json").on("click", () => html.find("#sheexcel-json-file").click());
    html.find("#sheexcel-json-file").on("change", e => importJsonHandler(e, sheet).then(() => sheet.render(false)));
    html.find('.sheexcel-export-json').on('click', () => exportJsonHandler(sheet));
    html.find(".sheexcel-main-subtab-nav a.item").on("click", sheet._onToggleSubtab.bind(sheet));
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-roll", e => handleRoll(e, sheet));
    html.find(".sheexcel-main-subtab-nav a.item.active").click();
    html.on("change", "select.sheexcel-reference-input[data-type='refType']", sheet._onRefTypeChange.bind(sheet));

    // Sidebar state, roll mode, damage mode
    const c = sheet.actor.getFlag("sheexcel_updated", "sidebarCollapsed");
    html.find('.sheexcel-sidebar').toggleClass('collapsed', !!c);
    const saved = game.settings.get("sheexcel", "rollMode") || "norm";
    html.find(`input[name="roll-mode"][value="${saved}"]`).prop("checked", true);
    html.find('input[name="roll-mode"]').on('change', (event) => {
        game.settings.set("sheexcel", "rollMode", event.target.value);
    });
    const savedModes = game.settings.get("sheexcel", "damageModes") || {};
    html.find('.sheexcel-damage-mode').each(function () {
        const idx = $(this).data('index');
        const mode = savedModes[idx] || "normal";
        $(this).val(mode);
    });
    html.find('.sheexcel-damage-mode').on('change', function (event) {
        const idx = $(this).data('index');
        const value = $(this).val();
        const modes = game.settings.get("sheexcel", "damageModes") || {};
        modes[idx] = value;
        game.settings.set("sheexcel", "damageModes", modes);
    });
}