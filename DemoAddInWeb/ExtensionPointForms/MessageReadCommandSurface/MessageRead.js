(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            loadProps();
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
        });
    };

    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;

        $('#ewsRequest').text(getSubjectEWSRequest(item.itemId));
        $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        $('#itemClass').text(item.itemClass);
        $('#itemId').text(item.itemId);
        $('#itemType').text(item.itemType);

        $('#message-props').hide();
        $('#appointment-props').hide();


        if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            loadAppointmentProps(item);
        }
        else {
            loadMessageProps(item);
        }
    }

    function loadMessageProps(item) {
        $('#message-props').show();

        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#conversationId').text(item.conversationId);
        $('#from').html(buildEmailAddressString(item.from));
        $('#internetMessageId').text(item.internetMessageId);
        $('#normalizedSubject').text(item.normalizedSubject);
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));

        // The following function doesn't work in some API sets (e.g. Exchange 2013)
        if (item.body !== undefined) {
            item.body.getAsync('text', function (async) { $('#body').html(async.value); $('#bodylength').html(async.value.length); });
        }
    }


    // Load properties from an Appointment object
    function loadAppointmentProps(item) {
        $('#appointment-props').show();

        $('#appt-attachments').html(buildAttachmentsString(item.attachments));
        $('#end').text(item.end.toLocaleString());
        $('#location').text(item.location);
        $('#appt-normalizedSubject').text(item.normalizedSubject);
        $('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
        $('#organizer').html(buildEmailAddressString(item.organizer));
        $('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));
        $('#resources').html(buildEmailAddressesString(item.resources));
        $('#start').text(item.start.toLocaleString());
        $('#appt-subject').text(item.subject);
    }
})();

function log(logInfo, asNewLine) {
    // Add the provided log text to the log window
    if (asNewLine === undefined) { asNewLine = true; }
    if (asNewLine) {
        if ($('#log').val() === "") {
            // Don't prepend a new line if there is nothing in the log
            $('#log').text($('#log').val() + logInfo);
        }
        else {
            $('#log').text($('#log').val() + "\n" + logInfo);
        }
    }
    else {
        $('#log').text($('#log').val() + logInfo);
    }
}

function displayMessageForm() {
    if (!$("#specificItemId").val()) { $("#specificItemId").val(Office.context.mailbox.item.itemId); }
    displayMessageFormItemId($("#specificItemId").val());
}

function displayMessageFormItemId(itemId) {
    Office.context.mailbox.displayMessageForm(itemId);
}

function getRESTToken() {
    log("Requesting REST token");
    Office.context.mailbox.getCallbackTokenAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var token = result.value;
            $("#RESTTokenId").val(token);
            log(" - success", false);
        } else {
            log(" - FAILED", false);
            $("#RESTTokenId").val("Error: " + result.error.code);
            log("Error: " + result.error.code);
        }
    });
}

function getAccessToken() {
    log("Requesting access token");
    Office.context.auth.getAccessTokenAsync(getAccessTokenCallback);
}

function getAccessTokenCallback(asyncResult) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var token = result.value;
        $("#accessTokenId").val(token);
        log(" - success", false);
    } else {
        log(" - FAILED", false);
        $("#accessTokenId").val("Error: " + result.error.code);
        log("Error: " + result.error.code);
    }

}

function getBody() {
    var coercionType = Office.CoercionType.Html;
    if ($("#specificItemId").val() === "Text")
        coercionType = Office.CoercionType.Text;
    log("Requesting item body");
    Office.context.mailbox.item.body.getAsync(coercionType, function (result) {
        if (result.status === "succeeded") {
            var body = result.value;
            $("#bodyId").val(body);
            log(" - complete");
        } else {
            log(" - FAILED", false);
            $("#bodyId").val("Error: " + result.error.code);
            log("Error: " + result.error.code);
        }
    });
}

function toggleVisibility(toggleDivId) {
    // Toggles the visibility of the specified element, depending upon button state

    var elementToToggle = document.getElementById(toggleDivId);
    if (elementToToggle.style.display === "none") {
        elementToToggle.style.display = "block";
    } else {
        elementToToggle.style.display = "none";
    }
}

function removeAttachmentsButton() {
    //removeAttachments(removeAttachmentsCallback);
    var item = Office.context.mailbox.item;
    if (item.attachments.length > 0) {
        // Remove the attachments
        for (i = 0; i < item.attachments.length; i++) {
            var attachment = item.attachments[i];
            Office.context.mailbox.item.removeAttachmentAsync(
                attachment.id,
                { asyncContext: null },
                function (asyncResult) {
                    log(attachment.name + " removed");
                }
            );
        }
    }
    else {
        log("No attachments found");
    }
}

function removeAttachmentsRESTButton() {
    removeAttachments(removeAttachmentsCallback);
}

function removeAttachments(callback) {
    var mailItemId = $('#itemId').val();
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (asyncResult) {
        if (asyncResult.status === "succeeded") {
            var getAttachmentsUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0) + '/attachments';
            $.ajax({
                url: getAttachmentsUrl,
                contentType: 'application/json',
                type: 'get',
                headers: { 'Authorization': 'Bearer ' + asyncResult.value }
            }).done(function (attachments) {
                var attachmentsList = [];
                var itemsProcessed = 0;
                attachments.value.forEach(function (attachment) {
                    if (attachment.Name !== "Mailplus.lqa") {
                        attachmentsList.push({
                            "Id": attachment.Id,
                            "Name": attachment.Name
                        });
                    }
                });
                if (attachmentsList.length > 0) {
                    attachmentsList.forEach(function (attachment) {
                        removeAttachment(attachment.Id, function (result) {
                            if (result === false) {
                                callback(false);
                            }
                            itemsProcessed++;
                            if (itemsProcessed === attachmentsList.length) {
                                callback(true);
                            }
                        });
                    });

                } else {
                    callback(true);
                }
            }).fail(function (error) {
                callback(false);
            });
        }
        else {
            callback(false);
        }
    });
}

function removeAttachmentsCallback(asyncResult) {
    if (asyncResult === true) {
        log("Attachment removed");
    } else {
        log("FAILED to remove attachment");
    }
}

function removeAttachment(attachmentId, callback) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (asyncResult) {
        if (asyncResult.status === "succeeded") {
            var messageId = Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
            var removeAttachmentsUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + messageId + '/attachments/' + attachmentId;
            $.ajax({
                url: removeAttachmentsUrl,
                contentType: 'application/json',
                type: 'DELETE',
                headers: { 'Authorization': 'Bearer ' + asyncResult.value }
            }).done(function (result) {
                callback(true);
            }).fail(function (error) {
                log(error.statusText);
                callback(false);
            });
        }
    });
}

function getSubjectEWSRequest(id) {
    // Return a GetItem operation request for the subject of the specified item.
    var request =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '        <t:AdditionalProperties>' +
        '            <t:FieldURI FieldURI="item:Subject"/>' +
        '        </t:AdditionalProperties>' +
        '      </ItemShape>' +
        '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';

    return request;
}

function sendEWSRequest() {
    log('Reading EWS request');
    var requestXml = $("#ewsRequest").val();
    log(' - complete', false);
    if (requestXml.length > 10) {
        log('Sending EWS request, length is ' + requestXml.length);
        result = Office.context.mailbox.makeEwsRequestAsync(requestXml, sendEWSRequestCallback);
        if (result === null) {
            log(' - FAILED', false);
        } else {
            log(' - complete', false);
        }
    }
    else {
        log('Invalid request');
    }
}

function sendEWSRequestCallback(asyncResult) {
    var result = asyncResult.value;
    var context = asyncResult.asyncContext;

    log('EWS Response received (' + asyncResult.status + ')');
    if (asyncResult.status === "succeeded") {
        $("#ewsResponse").text(result);
    } else {
        $("#ewsResponse").text("Error: " + asyncResult.error.message);
        log("Error: " + asyncResult.error.message);
    }
}

function saveAsync() {
    log("Saving item");
    Office.context.mailbox.item.saveAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            log(" - complete", false);
            var token = result.value;
            $("#itemId").val(result.value);
        } else {
            log(" - FAILED", false);
            log("Error: " + result.error.code);
        }
    });
}

// Helper function for displaying notifications
function showNotification(header, content) {
    log(header + ": " + content);
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    var element = document.querySelector('.ms-MessageBanner');
    messageBanner = new fabric.MessageBanner(element);
    messageBanner.showBanner();
}

function displayDialogAsync() {
    //Office.context.ui.displayDialogAsync('https://www.google.com', { height: 100, width: 100 }, dialogCallback );
    Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog/Dialog.html", { height: 50, width: 50 }, dialogCallback);
    ////IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps. 
    //Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html",
    //    { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

var dialog;
function dialogCallback(asyncResult) {
    if (asyncResult.status === "failed") {
        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Error", "Domain is not trusted");
                break;
            case 12005:
                showNotification("Error", "HTTPS is required");
                break;
            case 12007:
                showNotification("Error", "A dialog is already opened.");
                break;
            default:
                showNotification("Error", asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;

        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}