(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $('#attFileAttachmentAsyncBtn').on('click', function () {
            addFileAttachmentAsyncCall();
        });
        $(document).ready(function () {
            //var element = document.querySelector('.ms-MessageBanner');
            //messageBanner = new fabric.MessageBanner(element);
            //messageBanner.hideBanner();
			//loadProps();

			// The following function doesn't work in some API sets (e.g. Exchange 2013)
			if (Office.context.mailbox.addHandlerAsync !== undefined) {
				// Set up ItemChanged event
				Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
			}
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

    function itemChanged(eventArgs) {
        // Update UI based on the new current item
        loadProps();
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

    function loadMessageProps(item)
    {
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


    //// Helper function for displaying notifications
    //function showNotification(header, content) {
    //    $("#notificationHeader").text(header);
    //    $("#notificationBody").text(content);
    //    messageBanner.showBanner();
    //    messageBanner.toggleExpansion();
    //}
 })();

function displayMessageForm() {
    if (!$("#specificItemId").val()) { $("#specificItemId").val(Office.context.mailbox.item.itemId); }
    displayMessageFormItemId($("#specificItemId").val());
}

function displayDialogAsync() {
    //Office.context.ui.displayDialogAsync('https://www.google.com', { height: 100, width: 100 }, dialogCallback );
    Office.context.ui.displayDialogAsync(window.location.origin + "/MessageRead/Dialog.html", { height: 50, width: 50 }, dialogCallback);
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

function messageHandler(arg) {
    dialog.close();
    showNotification(arg.message);
}


function eventHandler(arg) {

    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.
    switch (arg.error) {
        case 12002:
            showNotification("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            showNotification("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            showNotification("Dialog closed by user");
            break;
        default:
            showNotification("Undefined error in dialog window");
            break;
    }
}

function displayMessageFormItemId(itemId) {
    Office.context.mailbox.displayMessageForm(itemId);
}

function getCallbackToken() {
    Office.context.mailbox.getCallbackTokenAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var token = result.value;
            $("#callbackTokenId").val(token);
        } else {
            $("#callbackTokenId").val("Error: " + result.error.code);
        }
    });
}

function cb(asyncResult) {
    var token = asyncResult.value;
    $("#callbackTokenId").val(token);
}

function getAccessToken() {
    Office.context.auth.getAccessTokenAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var token = result.value;
            $("#accessTokenId").val(token);
        } else {
            $("#accessTokenId").val("Error: " + result.error.code);
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
	$("#ewsResponse").text('Reading request');
	var requestXml = $("#ewsRequest").val();
	$("#ewsResponse").text('Request read');
	if (requestXml.length > 10) {
		$("#ewsResponse").text('Sending request, length is ' + requestXml.length);
		result = Office.context.mailbox.makeEwsRequestAsync(requestXml, sendEWSRequestCallback);
		if (result === null) {
			$("#ewsResponse").text('Failed to send request');
		} else {
			$("#ewsResponse").text('Request sent');
		}
	}
	else {
		$("#ewsResponse").text('Invalid request');
	}
}

function sendEWSRequestCallback(asyncResult) {
	var result = asyncResult.value;
	var context = asyncResult.asyncContext;

	$("#ewsResponse").val('Response received (' + asyncResult.status + ')');
	if (asyncResult.status === "succeeded") {
		$("#ewsResponse").text(result);
	} else {
		$("#ewsResponse").text("Error: " + asyncResult.error.message);
	}}

function saveAsync() {
    Office.context.mailbox.item.saveAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var token = result.value;
            $("#itemId").val(result.value);
        } else {
            $("#itemId").val("Error: " + result.error.code);
        }
    });
}

// Helper function for displaying notifications
function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    //var element = document.querySelector('.ms-MessageBanner');
    //messageBanner = new fabric.MessageBanner(element);
    //messageBanner.showBanner();
 }

function addFileAttachmentAsyncCall() {
    if ($("#fileAttachUrl").val() && $("#fileAttachName").val()) {
        WriteToLog("Adding file attachment with name \"" + $("#fileAttachName").val() + "\" from URL \"" + $("#fileAttachUrl").val() + "\"");
        Office.cast.item.toMessageCompose(Office.context.mailbox.item).addFileAttachmentAsync(
            $("#fileAttachUrl").val(),
            $("#fileAttachName").val(),
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    WriteToLog(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    var attachmentID = asyncResult.value;
                    WriteToLog('ID of added attachment: ' + attachmentID);
                }
            });
    }
}

function WriteToLog(message) {
    $("#logFileld").val($("#logFileld").val() + '\r\n'  + message);

}