export function wrapSheexcelChatFlavor(content, variant = "roll") {
  const safeVariant = String(variant || "roll")
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, "-")
    .replace(/^-+|-+$/g, "") || "roll";

  return `<div class="sheexcel-chat-flavor sheexcel-chat-flavor-${safeVariant}">${content}</div>`;
}