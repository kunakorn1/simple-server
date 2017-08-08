
Office.initialize = function () {
}

var itemId;
var subject;
var from;
var fromName;
var arrayOfToRecipients;
var mailTo = "";
var mailCC = "";
var mailBcc = "";
var bodyHTML = "";
var createdTime;

function getMailContents(){
  
  itemId = Office.context.mailbox.item.itemId.substring(0, 50);
  subject = Office.context.mailbox.item.subject;
  from = Office.context.mailbox.item.from.emailAddress;
  fromName = Office.context.mailbox.item.from.displayName;
  createdTime = Office.context.mailbox.item.dateTimeCreated;
  
   arrayOfToRecipients = Office.context.mailbox.item.to;
  for(i=0;i<arrayOfToRecipients.length;i++)
  {
    mailTo = mailTo + arrayOfToRecipients[i].displayName + " (" +  arrayOfToRecipients[i].emailAddress + ") ";
  }
  
  arrayOfToRecipients = Office.context.mailbox.item.cc;
  for(i=0;i<arrayOfToRecipients.length;i++)
  {
    mailCC = mailCC + arrayOfToRecipients[i].displayName + " (" +  arrayOfToRecipients[i].emailAddress + ") ";
  }
  
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(asyncResult){
    bodyHTML = asyncResult.value;  
    showMessage("msgBody", "We got the body contents in HTML format!!");
    var tmp = "";
    var contents = tmp.concat("<!doctype html>",
                              "<html>", "\r\n",
                              "<head>", "<br>\r\n",
                              "<meta charset='utf-8'> <br>\r\n",
                              "<title>", "Mail: ", itemId, "</title>", "<br>\r\n",
                              "</head>", "<br>\r\n",
                              "<body>", "<br>\r\n",
                           "<div><font face='Calibri, sans-serif' color='#000000' style='font-size:11pt'><b>Subject: </b>", subject, "<br>\r\n",
                              "<b>Sent: </b>", createdTime, "<br>\r\n",
                              "<b>From: </b>", fromName, "&lt " , from , " &gt", "<br>\r\n");
    if(mailTo !== ""){
      contents = contents.concat("<b>To: </b>", mailTo, "<br>\r\n");
    }
    if(mailCC !== ""){
      contents = contents.concat("<b>CC: </b>", mailCC, "<br>\r\n");
    }
     contents = contents.concat("</div><br><br>\r\n",
                              bodyHTML,
                             "</body>", "<br>\r\n",
                             "</html>");
  
    download(contents,"email_" + itemId + ".html");
    getAttachmentFiles();
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

function getAttachmentFiles(){
  var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
      var _att = _Item.attachments[i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += _att.name;
      outputString += "<BR>ID: " + _att.id;
      outputString += "<BR>contentType: " + _att.contentType;
      outputString += "<BR>size: " + _att.size;
      outputString += "<BR>attachmentType: " + _att.attachmentType;
      outputString += "<BR>isInline: " + _att.isInline;
    }
  
  download(outputString, "AttachmentFiles.txt");
  }
}


