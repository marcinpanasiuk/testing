Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.to.getAsync(function (result) {
        console.log(`OnNewMessageCompose, recipients count: ${result.value.length}`);
        event.completed();
    });
}

/**
 * Handles the OnMessageRecipientsChanged event.
 */
function onMessageRecipientsChangedHandler(event) {
    Office.context.mailbox.item.to.getAsync(function (result) {
        console.log(`OnMessageRecipientsChanged, recipients count: ${result.value.length}`);
        var signature = "<strong style='font-size: 25px;'> David Johnson </strong>";
        Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);