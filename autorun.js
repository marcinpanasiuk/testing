Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {
    console.log("OnNewMessageCompose");
}

/**
 * Handles the OnMessageRecipientsChanged event.
 */
function onMessageRecipientsChangedHandler(event) {
    console.log("OnMessageRecipientsChanged");

    var signature = "<strong style='font-size: 25px;'> David Johnson </strong>";
    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);