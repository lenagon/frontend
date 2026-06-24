/* eslint-disable no-undef */

// Показывает информационное уведомление прямо в открытом письме,
// чтобы было видно, что кнопка нажата. Реальная логика отправки в ERP
// сюда пока не добавлена - сейчас это только визуальный макет интерфейса.
function showPlaceholder(label) {
  const text =
    'Кнопка "' + label + '" пока не подключена к ERP - это только макет интерфейса.';

  try {
    if (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "sendToErpPlaceholder",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: text,
          icon: "icon16",
          persistent: false,
        }
      );
    }
  } catch (e) {
    // На случай, если notificationMessages недоступен в текущем контексте
    console.log(text, e);
  }
}

function onSendToErp(event) {

    const item = Office.context.mailbox.item;

    let subject = item.subject || "";

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        "sendToErpPlaceholder",
        {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Subject: " + subject,
            persistent: false
        }
    );

    event.completed();
}

function onSendMessage(event) {
  showPlaceholder("Send2ERP (Message)");
  event.completed();
}

function onSendAttachments(event) {
  showPlaceholder("Send2ERP (Attachments)");
  event.completed();
}

function onSendAllAttachments(event) {
  showPlaceholder("Send2ERP (All Attachments)");
  event.completed();
}

Office.onReady();

// Современный способ связать имя функции из манифеста с обработчиком.
Office.actions.associate("onSendToErp", onSendToErp);
Office.actions.associate("onSendMessage", onSendMessage);
Office.actions.associate("onSendAttachments", onSendAttachments);
Office.actions.associate("onSendAllAttachments", onSendAllAttachments);
