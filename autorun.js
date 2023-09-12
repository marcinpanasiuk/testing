Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.to.getAsync(function (result) {
        console.log(`OnNewMessageCompose, recipients count: ${result.value.length}`);
        Office.context.mailbox.getUserIdentityTokenAsync(function(result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              var msg = `Token retrieval failed with message: ${result.error.message}`;
            } else {
              var msg = result.value;
            }
            var signature = `<strong style='font-size: 25px;'> ${msg} </strong>`;
            Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
          });
    });
}

/**
 * Handles the OnMessageRecipientsChanged event.
 */
function onMessageRecipientsChangedHandler(event) {
    Office.context.mailbox.item.to.getAsync(function (result) {
        console.log(`OnMessageRecipientsChanged, recipients count: ${result.value.length}`);

        // we simulate siganture downloading and rendering delay
        setTimeout(function () {
            var signature = "<strong style='font-size: 25px;'> TESTING OK </strong>";
            Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
        }, 2000);
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
