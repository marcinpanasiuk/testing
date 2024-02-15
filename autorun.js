Office.initialize = function () { };

function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onNewMessageCompose fired", function () {

        try {
            Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
                type: "insightMessage",
                actions: [{ actionText: 'Show pane', actionType: 'showTaskPane', commandId: 'paneButton', contextData: '' }],
                message: "Hello!",
                icon: 'icon32'
            }, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    let message = result.error.name + ": " + result.error.message;
                    console.error(message);
                    Office.context.mailbox.item.body.prependAsync("failed: " + message, function () {
                        event.completed();
                    });
                }
                event.completed();
            });
        }
        catch (error) {
            let message = result.error.name + ": " + result.error.message;
            console.error(message);
            Office.context.mailbox.item.body.prependAsync("exception: " + message, function () {
                event.completed();
            });
        }
    });
}

function onMessageComposeHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onMessageCompose fired", function () {
        event.completed();
    });
}

function onMessageRecipientsChangedHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onNewMessageCompose fired", function () {
        event.completed();
    });
}

function OnMessageFromChangedHandler(event) {
    Office.context.mailbox.item.body.prependAsync("OnMessageFromChanged fired", function () {
        event.completed();
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("OnMessageFromChangedHandler", OnMessageFromChangedHandler);