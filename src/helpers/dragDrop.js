/**
 * Reorder checks by inserting the dragged check at a specific position.
 * @param {SheexcelActorSheet} sheet - The sheet instance
 * @param {string} draggedId - The ID of the dragged check
 * @param {string} targetId - The ID of the target check
 * @param {string} position - Either 'before' or 'after'
 */
export async function reorderCheck(sheet, draggedId, targetId, position = 'after') {
  let checks = foundry.utils.deepClone(sheet.actor.getFlag("sheexcel_updated", "cellReferences") || []);
  if (!checks.length) {
    console.warn("❌ Sheexcel | No checks available for reordering");
    return;
  }

  // Find and remove the dragged check (only from root level)
  const draggedIndex = checks.findIndex(check => check.id === draggedId);
  if (draggedIndex === -1) {
    console.warn(`❌ Sheexcel | Dragged check ${draggedId} not found at root level`);
    return;
  }

  // Find target index
  const targetIndex = checks.findIndex(check => check.id === targetId);
  if (targetIndex === -1) {
    console.warn(`❌ Sheexcel | Target check ${targetId} not found at root level`);
    return;
  }

  // Remove the dragged check
  const [draggedCheck] = checks.splice(draggedIndex, 1);
  
  // Calculate new insertion index (adjust for removed item)
  let newIndex = targetIndex;
  if (draggedIndex < targetIndex) {
    newIndex--; // Adjust because we removed an item before the target
  }
  
  if (position === 'before') {
    checks.splice(newIndex, 0, draggedCheck);
  } else {
    checks.splice(newIndex + 1, 0, draggedCheck);
  }

  await sheet.actor.setFlag("sheexcel_updated", "cellReferences", checks);
  sheet.render(false);
}

/**
 * Move a check as a subcheck of another check.
 * @param {SheexcelActorSheet} sheet - The sheet instance (usually 'this')
 * @param {string} draggedId - The ID of the dragged check
 * @param {string} targetId - The ID of the target check
 */
export async function nestCheck(sheet, draggedId, targetId) {
  let checks = foundry.utils.deepClone(sheet.actor.getFlag("sheexcel_updated", "cellReferences") || []);
  if (!checks.length) {
    console.warn("❌ Sheexcel | No checks available for nesting");
    return;
  }

  // Recursive function to remove a check and return it
  function removeCheck(checksArr, id) {
    for (let i = 0; i < checksArr.length; i++) {
      if (checksArr[i].id === id) {
        return checksArr.splice(i, 1)[0];
      }
      if (Array.isArray(checksArr[i].subchecks)) {
        const found = removeCheck(checksArr[i].subchecks, id);
        if (found) return found;
      }
    }
    return null;
  }
  
  // Recursive function to find a check
  function findCheck(checksArr, id) {
    for (let check of checksArr) {
      if (check.id === id) return check;
      if (Array.isArray(check.subchecks)) {
        const found = findCheck(check.subchecks, id);
        if (found) return found;
      }
    }
    return null;
  }

  const draggedCheck = removeCheck(checks, draggedId);
  if (!draggedCheck) {
    console.warn(`❌ Sheexcel | Could not find check with ID: ${draggedId}`);
    return;
  }
  
  const targetCheck = findCheck(checks, targetId);
  if (!targetCheck) {
    console.warn(`❌ Sheexcel | Could not find target check with ID: ${targetId}`);
    return;
  }
  
  // Initialize subchecks array if it doesn't exist
  targetCheck.subchecks = targetCheck.subchecks || [];
  targetCheck.subchecks.push(draggedCheck);
  
  // Sort subchecks alphabetically by keyword
  targetCheck.subchecks.sort((a, b) => 
    (a.keyword || '').localeCompare(b.keyword || '', undefined, { sensitivity: 'base' })
  );

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
  if (!checks.length) {
    console.warn("❌ Sheexcel | No checks available to move");
    return;
  }
  
  // Recursive function to remove a check and return it
  function removeCheck(checksArr, id) {
    for (let i = 0; i < checksArr.length; i++) {
      if (checksArr[i].id === id) {
        return checksArr.splice(i, 1)[0];
      }
      if (Array.isArray(checksArr[i].subchecks)) {
        const found = removeCheck(checksArr[i].subchecks, id);
        if (found) return found;
      }
    }
    return null;
  }

  const movedCheck = removeCheck(checks, checkId);

  if (!movedCheck) {
    console.warn(`❌ Sheexcel | No check found to move for id: ${checkId}`);
    return;
  }

  // Add it to the root array
  checks.push(movedCheck);

  // Save and re-render
  await sheet.actor.setFlag("sheexcel_updated", "cellReferences", checks);
  sheet.render(false);
}