// helpers/mainSearch.js
const SEARCH_HIGHLIGHT_CLASS = "sheexcel-search-highlight";

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function clearSearchHighlights(root) {
  if (!root) return;

  const existingMarks = $(root).find(`mark.${SEARCH_HIGHLIGHT_CLASS}`);

  existingMarks.each((_, mark) => {
      const textNode = document.createTextNode(mark.textContent || "");
      mark.replaceWith(textNode);
    });

  // Merge adjacent text nodes after unwrapping <mark> tags so multi-char queries
  // (e.g. "vi") can be found inside a single text node.
  root.normalize();
}

function buildSearchTerms(query) {
  return String(query || "")
    .split(",")
    .map((term) => term.trim().toLowerCase())
    .filter(Boolean);
}

function highlightSearchTermsIn(root, queries) {
  if (!root || !Array.isArray(queries) || !queries.length) return;

  // Ensure previously modified nodes are normalized before scanning.
  root.normalize();

  const pattern = new RegExp(
    queries
      .slice()
      .sort((a, b) => b.length - a.length)
      .map(escapeRegExp)
      .join("|"),
    "gi"
  );

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !node.nodeValue.trim()) {
        return NodeFilter.FILTER_REJECT;
      }

      const parent = node.parentElement;
      if (!parent) {
        return NodeFilter.FILTER_REJECT;
      }

      if (parent.closest(`mark.${SEARCH_HIGHLIGHT_CLASS}`)) {
        return NodeFilter.FILTER_REJECT;
      }

      if (["SCRIPT", "STYLE", "TEXTAREA", "INPUT", "BUTTON"].includes(parent.tagName)) {
        return NodeFilter.FILTER_REJECT;
      }

      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const textNodes = [];
  let currentNode = walker.nextNode();
  while (currentNode) {
    textNodes.push(currentNode);
    currentNode = walker.nextNode();
  }

  textNodes.forEach((textNode) => {
    const source = textNode.nodeValue;
    pattern.lastIndex = 0;
    if (!pattern.test(source)) {
      return;
    }

    pattern.lastIndex = 0;
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;

    source.replace(pattern, (match, offset) => {
      if (offset > lastIndex) {
        fragment.appendChild(document.createTextNode(source.slice(lastIndex, offset)));
      }

      const mark = document.createElement("mark");
      mark.className = SEARCH_HIGHLIGHT_CLASS;
      mark.textContent = match;
      fragment.appendChild(mark);
      lastIndex = offset + match.length;
      return match;
    });

    if (lastIndex < source.length) {
      fragment.appendChild(document.createTextNode(source.slice(lastIndex)));
    }

    textNode.parentNode.replaceChild(fragment, textNode);
  });
}

function groupHasDisplayedMatches(group, selector) {
  return group.find(selector).filter((__, element) => {
    return $(element).css("display") !== "none";
  }).length > 0;
}

function matchesAllTerms(text, searchTerms) {
  if (!searchTerms.length) return true;
  return searchTerms.every((term) => text.includes(term));
}

export function onSearch(event) {
  const query = (event.currentTarget.value || "").trim().toLowerCase();
  const searchTerms = buildSearchTerms(query);
  const activeSidebarTab = this.element.find(".sheexcel-sidebar-tab.active").first();
  const activeMainSubtab = activeSidebarTab.data("tab") === "main"
    ? this.element.find(".sheexcel-main-subtab-content:visible").first()
    : $();
  const active = activeMainSubtab.length ? activeMainSubtab : activeSidebarTab;
  if (!active.length) return;

  const tab = active.data("tab");
  clearSearchHighlights(active[0]);

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
      const titleNode = card.find(".sheexcel-spell-name").get(0);
      if (titleNode) {
        highlightSearchTermsIn(titleNode, searchTerms);
      }

      card.find(".sheexcel-spell-tag, .sheexcel-spell-stats, .sheexcel-spell-section-title, .sheexcel-spell-section-text").each((__, node) => {
        highlightSearchTermsIn(node, searchTerms);
      });

      const text = card.text().toLowerCase();
      const matches = matchesAllTerms(text, searchTerms);
      card.toggle(matches);
    });

    active.find(".sheexcel-spell-circle-group").each((_, entry) => {
      const group = $(entry);
      const hasMatches = groupHasDisplayedMatches(group, ".sheexcel-spell-card");
      group.toggle(hasMatches || !query);
    });
    return;
  }

  if (tab === "gears") {
    const cards = active.find(".sheexcel-gear-card");

    cards.each((_, entry) => {
      const card = $(entry);
      const titleNode = card.find(".sheexcel-gear-name").get(0);
      if (titleNode) {
        highlightSearchTermsIn(titleNode, searchTerms);
      }

      card.find(".sheexcel-gear-label, .sheexcel-gear-attribute, .sheexcel-gear-section-title, .sheexcel-gear-section-text, .sheexcel-gear-abilities-list").each((__, node) => {
        highlightSearchTermsIn(node, searchTerms);
      });

      const text = card.text().toLowerCase();
      const matches = matchesAllTerms(text, searchTerms);
      card.toggle(matches);
    });

    active.find(".sheexcel-gear-type-group").each((_, entry) => {
      const group = $(entry);
      const hasMatches = groupHasDisplayedMatches(group, ".sheexcel-gear-card");
      group.toggle(hasMatches || !query);
    });
    return;
  }

  if (tab === "abilities") {
    const cards = active.find(".sheexcel-ability-card");
    const groups = active.find(".sheexcel-ability-group");

    cards.each((_, entry) => {
      const card = $(entry);
      const name = card.find(".sheexcel-spell-name").text().toLowerCase();
      const effect = card.find(".sheexcel-spell-section-text").text().toLowerCase();
      const fullText = card.text().toLowerCase();
      const matches = matchesAllTerms(fullText, searchTerms);

      const titleNode = card.find(".sheexcel-spell-name").get(0);
      if (titleNode) {
        highlightSearchTermsIn(titleNode, searchTerms);
      }

      card.find(".sheexcel-spell-section-text").each((__, sectionNode) => {
        highlightSearchTermsIn(sectionNode, searchTerms);
      });

      card.toggle(matches);
    });

    groups.each((_, entry) => {
      const group = $(entry);
      const hasMatches = groupHasDisplayedMatches(group, ".sheexcel-ability-card");

      group.toggle(hasMatches || !query);
    });
    return;
  }

  if (tab === "rest") {
    const cards = active.find(".sheexcel-rest-card");

    cards.each((_, entry) => {
      const card = $(entry);
      const titleNode = card.find(".sheexcel-rest-title").get(0);
      if (titleNode) {
        highlightSearchTermsIn(titleNode, searchTerms);
      }

      card.find(".sheexcel-rest-summary, .sheexcel-rest-detail-line").each((__, node) => {
        highlightSearchTermsIn(node, searchTerms);
      });

      const text = card.text().toLowerCase();
      const matches = matchesAllTerms(text, searchTerms);
      card.toggle(matches);
    });

    active.find(".sheexcel-rest-section").each((_, entry) => {
      const group = $(entry);
      const hasMatches = groupHasDisplayedMatches(group, ".sheexcel-rest-card");
      group.toggle(hasMatches || !query);
    });
  }
}
