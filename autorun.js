Office.initialize = function () { };

function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.body.prependAsync("onNewMessageCompose fired", function () {
        Office.context.mailbox.item.notificationMessages.replaceAsync("addin-message", {
            type: "insightMessage",
            actions: [{ actionText: 'Open add-in pane', actionType: 'showTaskPane', commandId: 'paneButton', contextData: '' }],
            message: "Hello!",
            icon: 'icon32'
        }, callback)
        event.completed();
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