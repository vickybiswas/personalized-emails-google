### **Send Personalized Emails with Google Sheets and Docs**

Easily send customized emails to multiple people using this guide. Follow these steps:

---

### **1. Create a Google Sheet**
- Open Google Sheets and create a new sheet.
- Add **three columns** with these headers:  
  - `Email`  
  - `Name`  
  - `Note`  

#### **Example:**
| Email                   | Name   | Note                                          |
|-------------------------|--------|-----------------------------------------------|
| email.example.com       | Pankaj | I have known you since childhood and need your time. |
| anotheremail@example.com | John   | We met in Boston and have not caught up since. |

---

### **2. Create a Google Doc**
- Open Google Docs and create a new document.
- Write your email content using placeholders `{name}` and `{note}`.  

#### **Example:**
```text
Hi {name},

{note}

I wanted to invite you to my birthday.

Me
```

### **3. Add App Script in Google Sheets**
1. Go to **Extensions** >> **Apps Script** in the menu bar.  
2. Paste the following code into the script editor:

```javascript
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var emailTemplate = getEmailTemplate();

  for (var i = 1; i < data.length; i++) {
    var email = data[i][0];
    var name = data[i][1];
    var note = data[i][2];
    var personalizedEmail = emailTemplate.replace("{name}", name).replace("{note}", note);
    sendEmail(email, personalizedEmail);
  }
}

function getEmailTemplate() {
  var docId = "1w5ij-XXXXXXXXXXXXXXX-MytE_pB_XXXXXXXXXXXXXX"; // Replace with your Doc ID
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+docId+"&exportFormat=html";
  var param = 
        {
          method      : "get",
          headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
          muteHttpExceptions:true,
        };
  var html = UrlFetchApp.fetch(url,param).getContentText();

  return html;
}

function sendEmail(email, message) {
  GmailApp.sendEmail(email, "Subject of the email", "", {
    htmlBody: message
  });
}

function main() {
  sendEmails();
}
```

### **4. Run the Script**
1. Save the script.
2. From the toolbar in the Apps Script editor, select the function `main` from the dropdown menu.
3. Click **Run**.  
4. Grant any required permissions.


---



### Tips
1. If your email has wide margins change Page Margins from google Docs and set them to 0
2. [Will add more as questions are asked or anyone encounters an issue]