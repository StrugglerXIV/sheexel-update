// helpers/mainSearch.js
export function onSearch(event) {
  const query   = event.currentTarget.value.toLowerCase();
  const active  = this.element.find(".sheexcel-main-subtab-content:visible");
  active.find("button.sheexcel-roll").each((i,btn) => {
    const txt = $(btn).text().toLowerCase();
    $(btn).closest("div").toggle(txt.includes(query));
  });
}
