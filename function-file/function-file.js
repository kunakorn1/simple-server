
Office.initialize = function () {
}

var itemId;
var subject;
var from;
var fromName;
var bodyHTML = "";
var createdTime;

function getMailContents(){
  
  itemId = Office.context.mailbox.item.itemId.substring(0, 50);
  subject = Office.context.mailbox.item.subject;
  from = Office.context.mailbox.item.from.emailAddress;
  fromName = Office.context.mailbox.item.from.displayName;
  createdTime = Office.context.mailbox.item.dateTimeCreated;
  
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(asyncResult){
    bodyHTML = asyncResult.value;  
    showMessage("msgBody", "We got the body contents in HTML format!!");
    var tmp = "";
    var contents = tmp.concat("<!doctype html>",
                              "<html>", "\r\n",
                              "<head>", "<br>\r\n",
                              "<meta charset='utf-8'>"
                              "<title>", "Mail: ", itemId, "</title>", "\r\n",
                              "</head>", "\r\n",
                              "<body>", "\r\n",
                           "<div><font face='Calibri, sans-serif' color='#000000' style='font-size:11pt'><b>Subject: </b>", subject, "<br>", "\r\n",
                              "<b>Sent: </b>", createdTime, "<br>", "\r\n",
                              "<b>From: </b>", fromName, "&lt " , from , " &gt", "<br>", "\r\n",
                              "<b>To: xxx@xxx.xxx</b>", "<br>", "\r\n",
                              "</div><br><br>", "\r\n",
                             "</body>", "\r\n",
                             "</html>");
  
    download(contents,"email_" + itemId + ".html");
  });
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

function showMessage(key, msg){
  Office.context.mailbox.item.notificationMessages.addAsync(key, {
    type: "informationalMessage",
    message : msg,
    icon : "icon32",
    persistent: false
  });
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


