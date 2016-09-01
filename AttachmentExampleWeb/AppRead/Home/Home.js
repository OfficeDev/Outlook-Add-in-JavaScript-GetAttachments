/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />
var xhr;
var serviceRequest;

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            initApp();
        });
    };

    function initApp() {
        $("#footer").hide();

        if (Office.context.mailbox.item.attachments == undefined) {
            var testButton = document.getElementById("testButton");
            testButton.onclick = "";
            showToast("Not supported", "Attachments are not supported by your Exchange server.");
        } else if (Office.context.mailbox.item.attachments.length == 0) {
            var testButton = document.getElementById("testButton");
            testButton.onclick = "";
            showToast("No attachments", "There are no attachments on this item.");
        } else {

            // Initalize a context object for the app.
            //   Set the fields that are used on the request
            //   object to default values.
            serviceRequest = new Object();
            serviceRequest.attachmentToken = "";
            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            serviceRequest.attachments = new Array();
        }
    };

})();

function testAttachments() {
    Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
};

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
        serviceRequest.attachmentToken = asyncResult.value;
        makeServiceRequest();
    }
    else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}

function makeServiceRequest() {
    var attachment;
    xhr = new XMLHttpRequest();

    // Update the URL to point to your service location.
    xhr.open("POST", "https://localhost:44320/api/AttachmentService", true);

    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xhr.onreadystatechange = requestReadyStateChange;

    // Translate the attachment details into a form easily understood by WCF.
    for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment !== undefined) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
    }

    // Send the request. The response is handled in the 
    // requestReadyStateChange function.
    xhr.send(JSON.stringify(serviceRequest));
};


// Handles the response from the JSON web service.
function requestReadyStateChange() {
    if (xhr.readyState == 4) {
        if (xhr.status == 200) {
            var response = JSON.parse(xhr.responseText);
            if (!response.isError) {
                // The response indicates that the server recognized
                // the client identity and processed the request.
                // Show the response.
                var names = "<h2>Attachments processed: " + response.attachmentsProcessed + "</h2>";

                for (i = 0; i < response.attachmentNames.length; i++) {
                    names += response.attachmentNames[i] + "<br />";
                }
                document.getElementById("names").innerHTML = names;
            } else {
                showToast("Runtime error", response.message);
            }
        } else {
            if (xhr.status == 404) {
                showToast("Service not found", "The app server could not be found.");
            } else {
                showToast("Unknown error", "There was an unexpected error: " + xhr.status + " -- " + xhr.statusText);
            }
        }
    }
};

// Shows the service response.
function showResponse(response) {
    showToast("Service Response", "Attachments processed: " + response.attachmentsProcessed);
}

// Displays a message for 10 seconds.
function showToast(title, message) {

    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;

    $("#footer").show("slow");

    window.setTimeout(function () { $("#footer").hide("slow") }, 10000);
};

// *********************************************************
//
// Outlook-Add-in-Javascript-GetAttachments, https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************