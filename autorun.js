Office.initialize = function () { };

function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onNewMessageCompose fired\n\n", function () {

        try {
            Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
                type: "insightMessage",
                actions: [{ actionText: 'Show pane', actionType: 'showTaskPane', commandId: 'paneButton', contextData: '' }],
                message: "Hello!",
                icon: 'icon32'
            }, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    let message = result.error.name + ": " + result.error.message;
                    console.log(message);
                    Office.context.mailbox.item.body.prependAsync("failed: " + message + "\n\n", function () {
                        event.completed();
                    });
                }
                event.completed();
            });
        }
        catch (error) {
            console.log(error);
            let message = result.error.name + ": " + result.error.message;
            console.log(message);
            Office.context.mailbox.item.body.prependAsync("exception: " + message + "\n\n", function () {
                event.completed();
            });
        }
    });
}

function onMessageComposeHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onMessageCompose fired\n\n", function () {
        event.completed();
    });
}

function onMessageRecipientsChangedHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onMessageRecipientsChanged fired\n\n", function () {
        event.completed();
    });
}

function OnMessageFromChangedHandler(event) {
    Office.context.mailbox.item.body.prependAsync("OnMessageFromChanged fired\n\n", function () {
        event.completed();
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("OnMessageFromChangedHandler", OnMessageFromChangedHandler);