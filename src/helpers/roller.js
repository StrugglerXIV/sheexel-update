// helpers/roller.js
export async function handleRoll(event, sheet) {
  event.preventDefault();
  const $btn    = $(event.currentTarget);
  const mod     = Number($btn.data("value")) || 0;
  const keyword = $btn.text().trim();
  // 1️⃣ Read the chosen mode
  const mode = sheet.element.find('input[name="roll-mode"]:checked').val();
  // 2️⃣ Build the d20 formula
  let formula;
  if (mode === "adv") {
    formula = `2d20kh1 + ${mod}`;
  } else if (mode === "dis") {
    formula = `2d20kl1 + ${mod}`;
  } else {
    formula = `1d20 + ${mod}`;
  }
  const flavor = `${keyword} (${mod >= 0 ? "+" : ""}${mod})`;

  // 3️⃣ Construct and evaluate the roll asynchronously
  const roll = await new Roll(formula).evaluate({async: true});

  // 4️⃣ Post to chat with proper speaker and flavor
  roll.toMessage({
    speaker: ChatMessage.getSpeaker({ actor: sheet.actor }),
    flavor
  });
}
