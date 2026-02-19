// helpers/mainSearch.js
export function onSearch(event) {
  const query = (event.currentTarget.value || "").trim().toLowerCase();
  const active = this.element.find(".sheexcel-main-subtab-content:visible").first();
  if (!active.length) return;

  const tab = active.data("tab");

  if (tab === "checks") {
    const entries = active.find(".sheexcel-check-entry");
    let matchedCount = 0;

    entries.each((_, entry) => {
      const row = $(entry);
      const text = row.text().toLowerCase();
      const matches = !query || text.includes(query);
      row.toggle(matches);
      if (matches) matchedCount += 1;
    });

    active.find(".sheexcel-checks-grid").toggle(matchedCount > 0);
    active.find(".sheexcel-check-dropzone").toggle(!query || matchedCount > 0);
    active.find(".sheexcel-empty-checks").toggle(entries.length === 0 || (!!query && matchedCount === 0));
    return;
  }

  if (tab === "saves") {
    active.find(".save-button-margin").each((_, entry) => {
      const row = $(entry);
      row.toggle(!query || row.text().toLowerCase().includes(query));
    });
    return;
  }

  if (tab === "attacks") {
    active.find(".attack-entry").each((_, entry) => {
      const row = $(entry);
      row.toggle(!query || row.text().toLowerCase().includes(query));
    });
    return;
  }

  if (tab === "spells") {
    const cards = active.find(".sheexcel-spell-card");

    cards.each((_, entry) => {
      const card = $(entry);
      const matches = !query || card.text().toLowerCase().includes(query);
      card.toggle(matches);
    });

    active.find(".sheexcel-spell-circle-group").each((_, entry) => {
      const group = $(entry);
      const hasMatches = group.find(".sheexcel-spell-card:visible").length > 0;
      group.toggle(hasMatches || !query);
    });
    return;
  }

  if (tab === "gears") {
    const cards = active.find(".sheexcel-gear-card");

    cards.each((_, entry) => {
      const card = $(entry);
      const matches = !query || card.text().toLowerCase().includes(query);
      card.toggle(matches);
    });

    active.find(".sheexcel-gear-type-group").each((_, entry) => {
      const group = $(entry);
      const hasMatches = group.find(".sheexcel-gear-card:visible").length > 0;
      group.toggle(hasMatches || !query);
    });
  }
}
