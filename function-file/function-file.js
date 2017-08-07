// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

Office.initialize = function () {
}

function getMailContents(event){
  var itemId = Office.context.mailbox.item.itemId.substring(0, 50);
  var subject = Office.context.mailbox.item.subject;
  var from = Office.context.mailbox.item.from.emailAddress;
  
  var to; 
  Office.context.mailbox.item.to.getAsync('text', function(async){
    if (async.status !== Office.AsyncResultStatus.Succeeded){
      to = "Cannot get email address from 'To'.";  
    }
    else{
      to = async.value;  
    }
  });
  
  var createdTime = Office.context.mailbox.item.dateTimeCreated;
  
  var body;
  Office.context.mailbox.item.body.getAsync('text', function (async){
    if (async.status !== Office.AsyncResultStatus.Succeeded){
      body = "Cannot get email address from 'Body'.";  
    }
    else{
      body = async.value;  
    }
  });
  
  var bodyHTML;
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(async){
    if (async.status !== Office.AsyncResultStatus.Succeeded){
      bodyHTML = "Cannot get email address from 'Body' in HTML format.";  
    }
    else{
      bodyHTML = async.value.trim();  
    }
  });
  
  var tmp = "";
  var contents = tmp.concat("Subject: ", subject, "\r\n",
                           "From: ", from, "\r\n",
                           "To: ", to, "\r\n",
                           "Created Time: ", createdTime, "\r\n", "\r\n",
                           "Body in text plain:\r\n", body, "\r\n\r\n",
                           "Body in HTML:\r\n", bodyHTML);
  
  download(contents,"email_" + itemId + ".txt");
  event.completed();
}

// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

// Adds text into the body of the item, then reports the results
// to the info bar.
function addTextToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        statusUpdate(icon, "\"" + text + "\" inserted successfully.");
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
          type: "errorMessage",
          message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
        });
      }
      event.completed();
    });
}

function addDefaultMsgToBody(event) {
  addTextToBody("Inserted by the Add-in Command Demo add-in.", "blue-icon-16", event);
}

function addMsg1ToBody(event) {
  addTextToBody("Hello World!", "red-icon-16", event);
}

function addMsg2ToBody(event) {
  addTextToBody("Add-in commands are cool!", "red-icon-16", event);
}

function addMsg3ToBody(event) {
  addTextToBody("Visit https://developer.microsoft.com/en-us/outlook/ today for all of your add-in development needs.", "red-icon-16", event);
}

// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
  var subject = Office.context.mailbox.item.subject;
  
  Office.context.mailbox.item.notificationMessages.addAsync("subject", {
    type: "informationalMessage",
    icon: "blue-icon-16",
    message: "Subject: " + subject,
    persistent: false
  });
  
  download(subject, 'subject.txt', 'text/plain');
  
  event.completed();
}

// Gets the item class of the item and displays it in the info bar.
function getItemClass(event) {
  var itemClass = Office.context.mailbox.item.itemClass;
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemClass", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item Class: " + itemClass,
    persistent: false
  });
  
  event.completed();
}

// Gets the date and time when the item was created and displays it in the info bar.
function getDateTimeCreated(event) {
  var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
  
  Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Created: " + dateTimeCreated.toLocaleString(),
    persistent: false
  });
  
  event.completed();
}

// Gets the ID of the item and displays it in the info bar.
function getItemID(event) {
  // Limited to 150 characters max in the info bar, so 
  // only grab the first 50 characters of the ID
  var itemID = Office.context.mailbox.item.itemId.substring(0, 50);
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemID", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item ID -> " + itemID,
    persistent: false
  });
  
  event.completed();
}

function getSubjectRequest(id) {
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

function sendRequest(event) {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
  
  event.completed();
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
  download(result, 'result.txt', 'text/plain');
  download(context, 'context.txt', 'text/plain');
}

// Function to download data to a file
function download(data, filename) {
    var file = new Blob([data], {type: "text/plain;charset=utf-8"});
    if (window.navigator.msSaveOrOpenBlob) // IE10+
        window.navigator.msSaveOrOpenBlob(file, filename);
    else { // Others
        var a = document.createElement("a"),
                url = URL.createObjectURL(file);
        a.setAttribute("href",url);
        a.setAttribute("download",filename);
        document.body.appendChild(a);
        a.click();
        setTimeout(function() {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);  
        }, 0); 
    }
}
