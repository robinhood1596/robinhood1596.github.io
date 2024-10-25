Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Check if Office.context.mailbox.item is defined
  if (Office.context.mailbox && Office.context.mailbox.item) {
    // Show a notification message.
    Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);
  } else {
    console.error("Office.context.mailbox.item is undefined.");
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);