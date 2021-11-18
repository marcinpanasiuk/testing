Office.initialize = function (reason) { };

/**
 * Handles the OnMessageRecipientsChanged event.
 */
function onMessageRecipientsChangedHandler(event) {

    var signature = "<strong style='font-size: 25px;'> David Johnson </strong>";
    console.log(`Setting signature to "${signature}".`);
    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
}

Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);