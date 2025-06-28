// helpers/mainSearch.js
export function onSearch(event) {
  const query = event.currentTarget.value.toLowerCase();
  const active = this.element.find(".sheexcel-main-subtab-content:visible");

  // For each entry (check, attack, save, spell), check if it matches the query
  active.children("div").each((i, entry) => {
    const txt = $(entry).text().toLowerCase();
    $(entry).toggle(txt.includes(query));
  });
}
