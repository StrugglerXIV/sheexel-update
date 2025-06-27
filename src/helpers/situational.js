// Controls the situational bonuses.
export function promptBonus(keyword) {
  return new Promise(resolve => {
    new Dialog({
      title: `${keyword} — Situational Bonus?`,
      content: `
        <p>Bonus formula (e.g. +3, 1d4, 2d6+1):</p>
        <input type="text" name="bonus" value=""/>
      `,
      buttons: {
        roll: {
          label: "Roll",
          callback: html => {
            const raw = html.find("input[name='bonus']").val().trim();
            resolve(raw || "0");
          }
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null)
        }
      },
      default: "roll"
    }).render(true);
  });
}

export function promptDmgBonus(keyword) {
  return new Promise(resolve => {
    new Dialog({
      title: `${keyword} — DMG Bonus?`,
      content: `
        <p>Bonus formula (e.g. +3, 1d4, 2d6+1):</p>
        <input type="text" name="bonus" value=""/>
      `,
      buttons: {
        roll: {
          label: "Roll",
          callback: html => {
            const raw = html.find("input[name='bonus']").val().trim();
            resolve(raw || "0");
          }
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null)
        }
      },
      default: "roll"
    }).render(true);
  });
}


