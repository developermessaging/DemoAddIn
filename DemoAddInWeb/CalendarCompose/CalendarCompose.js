(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      //var element = document.querySelector('.ms-MessageBanner');
      ////messageBanner = new fabric.MessageBanner(element);
      //messageBanner.hideBanner();
      //loadProps();
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

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

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
  }



  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();

function requiredAttendeesSetAsync() {
    var recipients = [
        {
            "displayName": $("#requiredAttendeesNameId").val(),
            "emailAddress": $("#requiredAttendeesEmailId").val()
        }
    ];

    Office.context.mailbox.item.requiredAttendees.setAsync(recipients, function (result) {
        if (result.error) {
            showMessage(result.error);
        } else {
            var msg = "success!";
            showMessage(msg);
        }
    });
}


function requiredAttendeesAddAsync() {
    var recipients = [
        {
            "displayName": $("#requiredAttendeesNameId").val(),
            "emailAddress": $("#requiredAttendeesEmailId").val()
        }
    ];

    Office.context.mailbox.item.requiredAttendees.addAsync(recipients, function (result) {
        if (result.error) {
            showMessage(result.error);
        } else {
            var msg = "success!";
            showMessage(msg);
        }
    });
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

function saveCustomProperty() {

}
