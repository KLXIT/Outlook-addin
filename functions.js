function checkWrikeId(event) {
  Office.context.mailbox.item.subject.getAsync((result) => {
    const subject = result.value || "";
    if (!subject.match(/\bWrike\b/i)) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
        type: "informationalMessage",
        message: "⚠️ Please include a Wrike Task ID before sending.",
        icon: "icon16",
        persistent: false
      });
    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
        type: "informationalMessage",
        message: "✅ Wrike Task ID found in subject.",
        icon: "icon16",
        persistent: false
      });
    }
  });
  event.completed();
}
