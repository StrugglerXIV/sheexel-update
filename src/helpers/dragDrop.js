/**
 * Move a check as a subcheck of another check.
 * @param {SheexcelActorSheet} sheet - The sheet instance (usually 'this')
 * @param {string} draggedId - The ID of the dragged check
 * @param {string} targetId - The ID of the target check
 */
export async function nestCheck(sheet, draggedId, targetId) {
  let checks = foundry.utils.deepClone(sheet.actor.getFlag("sheexcel_updated", "cellReferences") || []);
  if (!checks.length) return;

  function removeCheck(checksArr, id) {
    for (let i = 0; i < checksArr.length; i++) {
      if (checksArr[i].id === id) {
        return checksArr.splice(i, 1)[0];
      }
      if (checksArr[i].subchecks) {
        const found = removeCheck(checksArr[i].subchecks, id);
        if (found) return found;
      }
    }
    return null;
  }
  function findCheck(checksArr, id) {
    for (let check of checksArr) {
      if (check.id === id) return check;
      if (check.subchecks) {
        const found = findCheck(check.subchecks, id);
        if (found) return found;
      }
    }
    return null;
  }

  const draggedCheck = removeCheck(checks, draggedId);
  if (!draggedCheck) return;
  const targetCheck = findCheck(checks, targetId);
  if (!targetCheck) return;
  targetCheck.subchecks = targetCheck.subchecks || [];
  targetCheck.subchecks.push(draggedCheck);

  await sheet.actor.setFlag("sheexcel_updated", "cellReferences", checks);
  sheet.render(false);
}

/**
 * Move a check to the root level.
 * @param {SheexcelActorSheet} sheet - The sheet instance (usually 'this')
 * @param {string} checkId - The ID of the check to move
 */
export async function moveCheckToRoot(sheet, checkId) {
  let checks = foundry.utils.deepClone(sheet.actor.getFlag("sheexcel_updated", "cellReferences") || []);
  if (!checks.length) return;
  let movedCheck = null;

  function removeCheck(checksArr, id) {
    for (let i = 0; i < checksArr.length; i++) {
      if (checksArr[i].id === id) {
        movedCheck = checksArr.splice(i, 1)[0];
        return true; // Stop searching after removal
      }
      if (Array.isArray(checksArr[i].subchecks) && checksArr[i].subchecks.length) {
        if (removeCheck(checksArr[i].subchecks, id)) return true;
      }
    }
    return false;
  }

  removeCheck(checks, checkId);

  if (!movedCheck) {
    console.warn("No check found to move for id:", checkId);
    return;
  }

  // Add it to the root array
  checks.push(movedCheck);

  // Save and re-render
  await sheet.actor.setFlag("sheexcel_updated", "cellReferences", checks);
  sheet.render(false);
}