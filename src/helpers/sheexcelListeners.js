import { importJsonHandler, exportJsonHandler } from "./importer.js";
import { handleRoll } from "./roller.js";
import { onSearch } from "./mainSearch.js";
import { nestCheck, moveCheckToRoot, reorderCheck } from './dragDrop.js';
import { onSaveReferences } from "./buttons/saveButton.js";
import { CSS_CLASSES, SETTINGS } from "./constants.js";
import { loadingManager } from "./loadingManager.js";
import { handleInlineRollInteraction } from "./inlineRolls.js";

export function attachSheexcelListeners(sheet, html) {
    html.find('.sheexcel-profile-pic').on('click', ev => {
        ev.preventDefault();
        ev.stopPropagation();

        const current = sheet.actor.img;
        new Dialog({
            title: "Profile Image",
            content: "<p>Open the image browser or view the current image?</p>",
            buttons: {
                browse: {
                    label: "Image Browser",
                    callback: () => {
                        new FilePicker({
                            type: "image",
                            current,
                            callback: path => sheet.actor.update({ img: path })
                        }).render(true);
                    }
                },
                view: {
                    label: "View Image",
                    callback: () => {
                        const src = current || "";
                        const content = src
                            ? `<img src="${src}" alt="${sheet.actor.name}" style="max-width:100%;height:auto;display:block;margin:0 auto;" />`
                            : "<p>No image set.</p>";
                        new Dialog({
                            title: sheet.actor.name,
                            content,
                            buttons: {
                                close: { label: "Close" }
                            }
                        }).render(true);
                    }
                },
                cancel: { label: "Cancel" }
            },
            default: "browse"
        }).render(true);
    });

    html.find('.sheexcel-armor-refresh').on('click', ev => {
        ev.preventDefault();
        ev.stopPropagation();
        sheet._updateArmorBlock(html, { force: true }).catch(() => {});
    });

    html.find('.sheexcel-stats-refresh').on('click', ev => {
        ev.preventDefault();
        ev.stopPropagation();
        sheet._updateStatsBlock(html, { force: true }).catch(() => {});
    });

    const nameWrap = html.find('.sheexcel-actor-name-wrap');
    const nameLabel = html.find('.sheexcel-actor-name-label');
    const nameInput = html.find('.sheexcel-actor-name-input');

    nameLabel.on('click', () => {
        nameWrap.addClass('is-editing');
        nameInput.val(sheet.actor.name || nameLabel.text());
        nameInput.trigger('focus');
        nameInput.trigger('select');
    });

    nameInput.on('keydown', ev => {
        if (ev.key === 'Enter') {
            ev.preventDefault();
            nameInput.trigger('blur');
        }
        if (ev.key === 'Escape') {
            ev.preventDefault();
            nameInput.val(sheet.actor.name || nameLabel.text());
            nameWrap.removeClass('is-editing');
        }
    });

    nameInput.on('blur', async () => {
        const value = String(nameInput.val() || '').trim();
        if (value && value !== sheet.actor.name) {
            await sheet.actor.update({ name: value });
            nameLabel.text(value);
        }
        nameWrap.removeClass('is-editing');
    });

    // Enhanced drag and drop for both main checks and subchecks
    html.find('.sheexcel-check-entry[draggable=true], .sheexcel-check-entry-sub[draggable=true]')
        .on('dragstart', ev => {
            ev.stopPropagation();
            try {
                const checkId = ev.currentTarget.dataset.checkId;
                if (!checkId) {
                    console.error("❌ Sheexcel | Missing checkId on draggable element");
                    ev.preventDefault();
                    return;
                }
                
                const payload = {
                    checkId: checkId,
                    subType: 'sheexcel-check',
                    isSubcheck: ev.currentTarget.classList.contains('sheexcel-check-entry-sub')
                };
                console.log("Drag payload:", payload);
                ev.originalEvent.dataTransfer.setData('text/plain', JSON.stringify(payload));
                ev.originalEvent.dataTransfer.effectAllowed = 'move';
                
                // Visual feedback
                ev.currentTarget.style.opacity = '0.5';
            } catch (error) {
                console.error("❌ Sheexcel | Drag start failed:", error);
                ev.preventDefault();
            }
        })
        .on('dragend', ev => {
            // Reset visual feedback and clean up
            ev.currentTarget.style.opacity = '';
            html.find('.sheexcel-drop-indicator').remove();
            html.find('.sheexcel-check-entry').each(function() {
                this.style.borderColor = '';
                this.style.backgroundColor = '';
            });
        })
    // Drop targets for nesting and reordering
    let lastDragTarget = null;
    let lastDragPosition = null;
    
    // Add basic drop detection
    html.on('drop', function(ev) {
        ev.preventDefault();  // Prevent browser default behavior
    });
    
    // Handle drops on the grid container
    html.find('.sheexcel-checks-grid')
        .on('dragover', ev => {
            ev.preventDefault();
            ev.originalEvent.dataTransfer.dropEffect = 'move';
            
            // Find the closest check entry
            const checkEntry = $(ev.target).closest('.sheexcel-check-entry');
            if (checkEntry.length) {
                const rect = checkEntry[0].getBoundingClientRect();
                const midY = rect.top + rect.height / 2;
                const isTopHalf = ev.originalEvent.clientY < midY;
                const currentTarget = checkEntry[0];
                const currentPosition = isTopHalf ? 'before' : 'after';
                
                // Only update indicators if target or position changed
                if (lastDragTarget !== currentTarget || lastDragPosition !== currentPosition) {
                    // Clean up previous indicators and styling
                    html.find('.sheexcel-drop-indicator').remove();
                    html.find('.sheexcel-check-entry').each(function() {
                        this.style.borderColor = '';
                        this.style.backgroundColor = '';
                    });
                    
                    // Determine if this is a reorder or nest operation
                    const targetHasChildren = currentTarget.dataset.hasChildren === 'true';
                    const isReorderZone = isTopHalf || !targetHasChildren;
                    
                    if (isReorderZone) {
                        // Show reorder indicator - make sure it doesn't block drops
                        const indicator = $('<div class="sheexcel-drop-indicator" style="pointer-events: none;"></div>');
                        if (isTopHalf) {
                            $(currentTarget).before(indicator);
                        } else {
                            $(currentTarget).after(indicator);
                        }
                    } else {
                        // Show nest indicator
                        currentTarget.style.borderColor = '#c7a86d';
                        currentTarget.style.backgroundColor = 'rgba(191, 160, 90, 0.1)';
                    }
                    
                    lastDragTarget = currentTarget;
                    lastDragPosition = currentPosition;
                }
            }
        })
        .on('drop', async (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            
            let targetElement = null;
            
            // If drop lands on grid container, find the check entry by coordinates
            if (ev.target.classList.contains('sheexcel-checks-grid')) {
                const allChecks = $(ev.target).find('.sheexcel-check-entry');
                const dropX = ev.originalEvent.clientX;
                const dropY = ev.originalEvent.clientY;
                
                allChecks.each(function(index) {
                    const rect = this.getBoundingClientRect();
                    if (dropX >= rect.left && dropX <= rect.right && 
                        dropY >= rect.top && dropY <= rect.bottom) {
                        targetElement = this;
                        return false; // Break the loop
                    }
                });
                
                if (!targetElement && lastDragTarget && lastDragTarget.dataset?.checkId) {
                    targetElement = lastDragTarget;
                }
            } else {
                // Try closest for other elements
                const checkEntry = $(ev.target).closest('.sheexcel-check-entry');
                if (checkEntry.length) {
                    targetElement = checkEntry[0];
                }
            }
if (!targetElement || !targetElement.dataset?.checkId) {
                return;
            }
            
            // Reset visual feedback immediately
            targetElement.style.borderColor = '';
            targetElement.style.backgroundColor = '';
            html.find('.sheexcel-drop-indicator').remove();
            lastDragTarget = null;
            lastDragPosition = null;
            
            let payload;
            try {
                const data = ev.originalEvent.dataTransfer.getData('text/plain');
                payload = JSON.parse(data);
            } catch (error) { 
                console.warn("❌ Sheexcel | Invalid drag data:", error);
                return; 
            }
            
            if (!payload?.checkId) {
                console.warn("❌ Sheexcel | No checkId in drag payload");
                return;
            }
            
            const draggedId = payload.checkId;
            const targetId = targetElement.dataset.checkId;
            
            if (draggedId === targetId) {
                console.warn("❌ Sheexcel | Cannot drop item on itself");
                return;
            }
            
            // Determine operation type based on drop position
            const rect = targetElement.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            const isTopHalf = ev.originalEvent.clientY < midY;
            const targetHasChildren = targetElement.dataset.hasChildren === 'true';
            const isReorderOperation = isTopHalf || !targetHasChildren;
            
            try {
                if (isReorderOperation) {
                    // Reorder operation
                    const position = isTopHalf ? 'before' : 'after';
                    await reorderCheck(sheet, draggedId, targetId, position);
                    ui.notifications.info("Check reordered successfully");
                } else {
                    // Nest operation
                    await nestCheck(sheet, draggedId, targetId);
                    ui.notifications.info("Check nested successfully");
                }
            } catch (error) {
                console.error("❌ Sheexcel | Drop operation failed:", error);
                ui.notifications.error("Failed to move check: " + error.message);
            }
        });

    html.find('.sheexcel-check-dropzone')
        .on('dragover', ev => {
            ev.preventDefault();
            // Visual feedback for dropzone
            $(ev.currentTarget).addClass('active-dropzone');
        })
        .on('dragleave', ev => {
            // Reset visual feedback
            $(ev.currentTarget).removeClass('active-dropzone');
        })
        .on('drop', async (ev) => {
            ev.preventDefault();
            // Reset visual feedback
            $(ev.currentTarget).removeClass('active-dropzone');
            
            let payload;
            try {
                const data = ev.originalEvent.dataTransfer.getData('text/plain');
                payload = JSON.parse(data);
            } catch (error) { 
                console.warn("❌ Sheexcel | Invalid drag data:", error);
                return; 
            }
            
            if (!payload?.checkId) {
                console.warn("❌ Sheexcel | No checkId in drag payload");
                return;
            }
            
            try {
                await moveCheckToRoot(sheet, payload.checkId);
                ui.notifications.info("Check moved to root level");
            } catch (error) {
                console.error("❌ Sheexcel | Move to root failed:", error);
                ui.notifications.error("Failed to move check to root: " + error.message);
            }
        });

        // Clear search functionality
    html.find('.sheexcel-search-clear').on('click', function() {
        const searchInput = html.find('.sheexcel-search');
        searchInput.val('').trigger('input');
        $(this).hide();
    });
    
    // Show/hide clear button based on search input
    html.find('.sheexcel-search').on('input', function() {
        const clearBtn = html.find('.sheexcel-search-clear');
        if ($(this).val().trim()) {
            clearBtn.show();
        } else {
            clearBtn.hide();
        }
    });
    
    // Initialize subchecks as collapsed and set proper icons
    html.find('.sheexcel-subchecks').hide();
    html.find('.sheexcel-collapse-icon').text('▶');
    
    // Enhanced collapse toggle with smooth animation
    html.find('.sheexcel-collapse-toggle').on('click', function(ev) {
        ev.preventDefault();
        const $toggle = $(this);
        const $entry = $toggle.closest('.sheexcel-check-entry');
        const $subchecks = $entry.find('.sheexcel-subchecks');
        const $icon = $toggle.find('.sheexcel-collapse-icon');
        
        if ($subchecks.is(':visible')) {
            $subchecks.slideUp(200);
            $icon.text('▶');
            $toggle.attr('title', 'Show subchecks');
        } else {
            $subchecks.slideDown(200);
            $icon.text('▼');
            $toggle.attr('title', 'Hide subchecks');
        }
    });

    html.find(".sheexcel-sheet-toggle").on("click", sheet._onToggleSidebar.bind(sheet));
    html.find(".sheexcel-sheet-main, .sheexcel-sheet-references, .sheexcel-sheet-settings, .sheexcel-sheet-sheet")
        .on("click", sheet._onToggleTab.bind(sheet));
    html.find(".sheexcel-setting-update-sheet").on("click", sheet._onUpdateSheet.bind(sheet));
    html.find('.sheexcel-search').on('input', onSearch.bind(sheet));
    html.find('.sheexcel-references-search-input').on('input', function(ev) {
        const searchTerm = ev.target.value.toLowerCase();
        html.find('.sheexcel-reference-card').each(function() {
            const card = $(this);
            const name = card.find('.sheexcel-reference-name').text().toLowerCase();
            const cell = card.find('.sheexcel-cell-input').val().toLowerCase();
            const sheet = card.find('.sheexcel-sheet-select').val().toLowerCase();
            const type = card.find('.sheexcel-type-select').val().toLowerCase();
            
            const matches = name.includes(searchTerm) || 
                           cell.includes(searchTerm) || 
                           sheet.includes(searchTerm) || 
                           type.includes(searchTerm);
            
            card.toggleClass('sheexcel-search-match', matches);
            card.toggle(matches);
        });

        html.find('.sheexcel-reference-type-group').each(function() {
            const group = $(this);
            const matchingCards = group.find('.sheexcel-reference-card.sheexcel-search-match').length;
            group.toggle(matchingCards > 0);
        });
    });
    html.on("click", ".sheexcel-reference-add-button", sheet._onAddReference.bind(sheet));
    html.on("click", ".sheexcel-bulk-attacks-button", sheet._onBulkAddAttacks.bind(sheet));
    html.on("click", ".sheexcel-bulk-undo-button", sheet._onUndoBulkAddAttacks.bind(sheet));
    html.on("click", ".sheexcel-clear-checks-button", sheet._onClearChecks.bind(sheet));
    html.on("click", ".sheexcel-clear-saves-button", sheet._onClearSaves.bind(sheet));
    html.on("click", ".sheexcel-clear-attacks-button", sheet._onClearAttacks.bind(sheet));
    html.on("click", ".sheexcel-clear-spells-button", sheet._onClearSpells.bind(sheet));
    html.on("click", ".sheexcel-clear-abilities-button", sheet._onClearAbilities.bind(sheet));
    html.on("click", ".sheexcel-clear-gears-button", sheet._onClearGears.bind(sheet));
    html.on("click", ".sheexcel-reference-remove-button", sheet._onRemoveReference.bind(sheet));
    html.on("click", ".sheexcel-reference-remove-save", function(event) {
        event.preventDefault();
        const index = parseInt($(this).data('index'));
        sheet._onFetchAndUpdateCellValueByIndex(index);
    });
    html.on("click", ".sheexcel-reference-save-button", (event) => onSaveReferences(sheet, event));
    
    // Collapsible reference card toggle
    html.on("click", ".sheexcel-card-header[data-action='toggle-collapse']", function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        const card = $(this).closest('.sheexcel-reference-card');
        card.toggleClass('collapsed');
    });

    // Collapsible reference type group toggle
    html.on("click", ".sheexcel-reference-group-header[data-action='toggle-group-collapse']", function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        const group = $(this).closest('.sheexcel-reference-type-group');
        group.toggleClass('collapsed');
    });
    html.find(".sheexcel-import-json").on("click", () => html.find("#sheexcel-json-file").click());
    html.find("#sheexcel-json-file").on("change", e => importJsonHandler(e, sheet).then(() => sheet.render(false)));
    html.find('.sheexcel-export-json').on('click', () => exportJsonHandler(sheet));
    html.find(".sheexcel-main-subtab-nav a.item").on("click", sheet._onToggleSubtab.bind(sheet));
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-skills-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateSkillsFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-saves-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateSavesFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-gears-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateGearsFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-attacks-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateAttacksFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-roll", e => handleRoll(e, sheet));
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-spells-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateSpellsFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-abilities-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateAbilitiesFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-update-rest-button", (ev) => {
        ev.preventDefault();
        sheet._onUpdateRestFromSheet(ev);
    });
    html.find(".sheexcel-update-all-button").on("click", (ev) => {
        ev.preventDefault();
        sheet._onUpdateAllFromSheet(ev);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-spell-toggle", function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        const card = $(this).closest(".sheexcel-spell-card");
        card.toggleClass("collapsed");
        const isCollapsed = card.hasClass("collapsed");
        $(this).attr("title", isCollapsed ? "Expand" : "Collapse");
        $(this).find(".sheexcel-spell-toggle-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-gear-toggle", function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        const card = $(this).closest(".sheexcel-gear-card");
        card.toggleClass("collapsed");
        const isCollapsed = card.hasClass("collapsed");
        $(this).attr("title", isCollapsed ? "Expand" : "Collapse");
        $(this).find(".sheexcel-gear-toggle-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-gear-type-header", function(ev) {
        ev.preventDefault();
        const group = $(this).closest(".sheexcel-gear-type-group");
        group.toggleClass("collapsed");
        const isCollapsed = group.hasClass("collapsed");
        $(this).find(".sheexcel-gear-type-collapse-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-gears-toggle-all-groups", function(ev) {
        ev.preventDefault();
        const button = $(this);
        const container = button.closest(".sheexcel-main-subtab-content");
        const groups = container.find(".sheexcel-gear-type-group");
        const state = button.data("state") || "expanded";
        const collapse = state === "expanded";
        groups.toggleClass("collapsed", collapse);
        groups.find(".sheexcel-gear-type-collapse-icon").text(collapse ? "▶" : "▼");
        button.data("state", collapse ? "collapsed" : "expanded");
        button.text(collapse ? "Expand All Groups" : "Collapse All Groups");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-gears-toggle-all-cards", function(ev) {
        ev.preventDefault();
        const button = $(this);
        const container = button.closest(".sheexcel-main-subtab-content");
        const cards = container.find(".sheexcel-gear-card");
        const state = button.data("state") || "expanded";
        const collapse = state === "expanded";
        cards.toggleClass("collapsed", collapse);
        cards.find(".sheexcel-gear-toggle").attr("title", collapse ? "Expand" : "Collapse");
        cards.find(".sheexcel-gear-toggle-icon").text(collapse ? "▶" : "▼");
        button.data("state", collapse ? "collapsed" : "expanded");
        button.text(collapse ? "Expand All Items" : "Collapse All Items");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-spell-circle-header", function(ev) {
        ev.preventDefault();
        const group = $(this).closest(".sheexcel-spell-circle-group");
        group.toggleClass("collapsed");
        const isCollapsed = group.hasClass("collapsed");
        $(this).find(".sheexcel-spell-circle-toggle-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-spells-toggle-groups", function(ev) {
        ev.preventDefault();
        const button = $(this);
        const container = button.closest(".sheexcel-main-subtab-content");
        const groups = container.find(".sheexcel-spell-circle-group");
        const state = button.data("state") || "expanded";
        const collapse = state === "expanded";
        groups.toggleClass("collapsed", collapse);
        groups.find(".sheexcel-spell-circle-toggle-icon").text(collapse ? "▶" : "▼");
        button.data("state", collapse ? "collapsed" : "expanded");
        button.text(collapse ? "Expand All Circles" : "Collapse All Circles");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-spells-toggle-all", function(ev) {
        ev.preventDefault();
        const button = $(this);
        const container = button.closest(".sheexcel-main-subtab-content");
        const cards = container.find(".sheexcel-spell-card");
        const state = button.data("state") || "expanded";
        const collapse = state === "expanded";
        cards.toggleClass("collapsed", collapse);
        cards.find(".sheexcel-spell-toggle").attr("title", collapse ? "Expand" : "Collapse");
        cards.find(".sheexcel-spell-toggle-icon").text(collapse ? "▶" : "▼");
        button.data("state", collapse ? "collapsed" : "expanded");
        button.text(collapse ? "Expand All Spells" : "Collapse All Spells");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-abilities-toggle-all", function(ev) {
        ev.preventDefault();
        const button = $(this);
        const container = button.closest(".sheexcel-main-subtab-content");
        const groups = container.find(".sheexcel-ability-group");
        const cards = container.find(".sheexcel-ability-card");
        const state = button.data("state") || "expanded";
        const collapse = state === "expanded";
        groups.toggleClass("collapsed", collapse);
        groups.find(".sheexcel-spell-circle-toggle-icon").text(collapse ? "▶" : "▼");
        cards.toggleClass("collapsed", collapse);
        cards.find(".sheexcel-ability-toggle").attr("title", collapse ? "Expand" : "Collapse");
        cards.find(".sheexcel-spell-toggle-icon").text(collapse ? "▶" : "▼");
        button.data("state", collapse ? "collapsed" : "expanded");
        button.text(collapse ? "Expand All Abilities" : "Collapse All Abilities");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-spell-card", (e) => {
        if ($(e.target).closest(".sheexcel-spell-toggle").length) return;
        const index = Number($(e.currentTarget).data("index"));
        if (!Number.isInteger(index)) return;
        if ($(e.currentTarget).hasClass("sheexcel-ability-card")) {
            sheet._onPostAbilityToChat(index);
            return;
        }
        sheet._onPostSpellToChat(index);
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-rest-section-header", function(ev) {
        ev.preventDefault();
        const group = $(this).closest(".sheexcel-rest-section");
        group.toggleClass("collapsed");
        const isCollapsed = group.hasClass("collapsed");
        $(this).find(".sheexcel-rest-section-toggle-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click", ".sheexcel-rest-toggle", function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        const card = $(this).closest(".sheexcel-rest-card");
        card.toggleClass("collapsed");
        const isCollapsed = card.hasClass("collapsed");
        $(this).attr("title", isCollapsed ? "Expand" : "Collapse");
        $(this).find(".sheexcel-rest-toggle-icon").text(isCollapsed ? "▶" : "▼");
    });
    html.find(".sheexcel-main-subtab-content").on("click keydown", ".sheexcel-rest-line-post, .sheexcel-rest-card-post", function(ev) {
        if (ev.type === "keydown" && !["Enter", " ", "Spacebar"].includes(ev.key)) return;
        if ($(ev.target).closest(".sheexcel-inline-roll").length) return;
        ev.preventDefault();
        ev.stopPropagation();
        const entryIndex = Number($(this).data("entryIndex"));
        const detailIndex = $(this).data("detailIndex");
        const lineType = String($(this).data("lineType") || "detail");
        if (!Number.isInteger(entryIndex)) return;
        sheet._onPostRestLineToChat(
            entryIndex,
            detailIndex == null ? null : Number(detailIndex),
            lineType
        );
    });
    html.find(".sheexcel-main-subtab-content").on("click keydown", ".sheexcel-inline-roll", (ev) => {
        handleInlineRollInteraction(
            ev,
            () => sheet.actor,
            (target) => {
                const restTitle = target.closest(".sheexcel-rest-card")?.querySelector(".sheexcel-rest-title")?.textContent?.trim();
                if (restTitle) return restTitle;

                const spellOrAbility = target.closest(".sheexcel-spell-card")?.querySelector(".sheexcel-spell-name")?.textContent?.trim();
                if (spellOrAbility) return spellOrAbility;

                const gear = target.closest(".sheexcel-gear-card")?.querySelector(".sheexcel-gear-name")?.textContent?.trim();
                if (gear) return gear;

                return "Sheet";
            }
        );
    });
    html.find(".sheexcel-main-subtab-nav a.item.active").click();
    html.on("change", "select.sheexcel-reference-input[data-type='refType']", sheet._onRefTypeChange.bind(sheet));

    // Sidebar state, roll mode, damage mode restoration
    const c = sheet.actor.getFlag("sheexcel_updated", "sidebarCollapsed");
    html.find('.sheexcel-sidebar').toggleClass('collapsed', !!c);
    
    // Restore roll mode from settings with error handling
    try {
        const saved = game.settings.get("sheexcel", SETTINGS.ROLL_MODE) || "norm";
        html.find(`input[name="roll-mode"][value="${saved}"]`).prop("checked", true);
    } catch (error) {
        console.warn("❌ Sheexcel | Failed to restore roll mode:", error);
    }
    
    html.find('input[name="roll-mode"]').on('change', (event) => {
        try {
            game.settings.set("sheexcel", SETTINGS.ROLL_MODE, event.target.value);
        } catch (error) {
            console.error("❌ Sheexcel | Failed to save roll mode:", error);
        }
    });
    
    // Damage mode handling with error recovery
    try {
        const savedModes = game.settings.get("sheexcel", SETTINGS.DAMAGE_MODES) || {};
        html.find('.sheexcel-damage-mode').each(function () {
            const idx = $(this).data('index');
            const mode = savedModes[idx] || "normal";
            $(this).val(mode);
        });
    } catch (error) {
        console.warn("❌ Sheexcel | Failed to restore damage modes:", error);
    }
    
    html.find('.sheexcel-damage-mode').on('change', function (event) {
        try {
            const idx = $(this).data('index');
            const value = $(this).val();
            const modes = game.settings.get("sheexcel", SETTINGS.DAMAGE_MODES) || {};
            modes[idx] = value;
            game.settings.set("sheexcel", SETTINGS.DAMAGE_MODES, modes);
        } catch (error) {
            console.error("❌ Sheexcel | Failed to save damage mode:", error);
        }
    });

    // --- Iframe focus scroll-lock ---
    // When the Google Sheet iframe has focus, the browser auto-scrolls ancestor
    // containers any time the iframe navigates internally (cell confirm, Enter, etc.).
    // Fix: while the iframe has focus, run a rAF loop that continuously pins every
    // scrollable ancestor at its position from when focus entered.
    // The lock is released the moment focus returns to the Foundry window.
    const iframe = html.find('.sheexcel-google-sheet')[0];
    const embedWrapper = html.find('.sheexcel-sheet-google-wrapper')[0];
    const resizer = html.find('.sheexcel-sheet-resizer')[0];

    if (embedWrapper) {
        const MIN_EMBED_HEIGHT = 360;
        const savedEmbedHeight = Number(sheet.actor.getFlag(MODULE_NAME, FLAGS.SHEET_EMBED_HEIGHT));

        if (Number.isFinite(savedEmbedHeight) && savedEmbedHeight > 0) {
            embedWrapper.style.height = `${Math.max(MIN_EMBED_HEIGHT, Math.round(savedEmbedHeight))}px`;
        }

        if (resizer) {
            let startY = 0;
            let startHeight = 0;
            let resizing = false;

            const onMouseMove = (ev) => {
                if (!resizing) return;
                const delta = ev.clientY - startY;
                const next = Math.max(MIN_EMBED_HEIGHT, Math.round(startHeight + delta));
                embedWrapper.style.height = `${next}px`;
            };

            const onMouseUp = async () => {
                if (!resizing) return;
                resizing = false;
                $(document).off('mousemove.sheexcelResize', onMouseMove);
                $(document).off('mouseup.sheexcelResize', onMouseUp);

                const finalHeight = Math.max(MIN_EMBED_HEIGHT, Math.round(embedWrapper.getBoundingClientRect().height));
                try {
                    await sheet.actor.setFlag(MODULE_NAME, FLAGS.SHEET_EMBED_HEIGHT, finalHeight);
                } catch (error) {
                    console.error('❌ Sheexcel | Failed to persist sheet embed height:', error);
                }
            };

            $(resizer).on('mousedown', (ev) => {
                ev.preventDefault();
                startY = ev.clientY;
                startHeight = embedWrapper.getBoundingClientRect().height;
                resizing = true;
                $(document).on('mousemove.sheexcelResize', onMouseMove);
                $(document).on('mouseup.sheexcelResize', onMouseUp);
            });

            html.one('remove', () => {
                $(document).off('mousemove.sheexcelResize', onMouseMove);
                $(document).off('mouseup.sheexcelResize', onMouseUp);
            });
        }
    }

    if (iframe) {
        let lockLoop = null;
        let locked = null;

        const isSheetTabActive = () => html.find('.sheexcel-sidebar-tab.sheet').hasClass('active');

        const getScrollableAncestors = (el) => {
            const ancestors = [];
            let node = el.parentElement;
            while (node) {
                const ov = window.getComputedStyle(node).overflowY;
                if (ov === 'auto' || ov === 'scroll' || ov === 'overlay') {
                    ancestors.push(node);
                }
                node = node.parentElement;
            }
            return ancestors;
        };

        const stopLock = () => {
            if (lockLoop) { cancelAnimationFrame(lockLoop); lockLoop = null; }
            locked = null;
        };

        const runLoop = () => {
            if (!locked) return;
            // Only keep the lock while the iframe is focused on the active Sheet tab.
            if (document.activeElement !== iframe || !isSheetTabActive()) {
                stopLock();
                return;
            }
            for (const { el, top } of locked) el.scrollTop = top;
            lockLoop = requestAnimationFrame(runLoop);
        };

        const onBlur = () => {
            if (document.activeElement !== iframe || !isSheetTabActive()) return;
            const ancestors = getScrollableAncestors(iframe);
            locked = ancestors.map(el => ({ el, top: el.scrollTop }));
            if (lockLoop) cancelAnimationFrame(lockLoop);
            lockLoop = requestAnimationFrame(runLoop);
        };

        const onFocus = () => stopLock();

        window.addEventListener('blur', onBlur);
        window.addEventListener('focus', onFocus);

        // Defensive stop: if user interacts anywhere outside the iframe, release lock.
        const onUserInteraction = (ev) => {
            if ($(ev.target).closest('.sheexcel-google-sheet').length) return;
            stopLock();
        };
        html.on('mousedown wheel keydown touchstart', onUserInteraction);

        // Clean up when the sheet re-renders so listeners don't stack
        html.one('remove', () => {
            window.removeEventListener('blur', onBlur);
            window.removeEventListener('focus', onFocus);
            html.off('mousedown wheel keydown touchstart', onUserInteraction);
            stopLock();
        });
    }
}