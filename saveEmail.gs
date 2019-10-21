//Author: MH Charles-Etuk
//Date: 10/15/2019
//Description: Google Script program to save emails with certain labels applied as PDFs in a specified Google Drive folder

function save_Gmail_as_PDF()
{
  var label = GmailApp.getUserLabelByName("Save As PDF"); //find label of emails to save
  
  if (label == null) //if label doesn't exist create it
    GmailApp.createLabel('Save As PDF');
  
  else
  {
    var threads = label.getThreads();  //get all the emails
    
    for (var i = 0; i < threads.length; i++) //loop through each email thread in the set of emails
    {  
      var messages = threads[i].getMessages(); //store all the messages in this message array 
      var message = messages[0]; //store each message in this message variable
      var body    = message.getBody(); //save text
      var subject = message.getSubject(); //save text
      var attachments  = message.getAttachments(); //save attachments if any
      
      for (var j = 1; j < messages.length; j++) //loop through the messages
      {
        body += messages[j].getBody(); //append new text to the body
        var temp_attach = messages[j].getAttachments(); //store attachment in this variable
        
        if (temp_attach.length > 0) //if the email has an attachment in it, add it to the attachments
        {
          for (var k = 0; k < temp_attach.length; k++)
            attachments.push(temp_attach[k]);
        }
      }
      
      // Create an HTML File from the Message Body
      
      var emails = DriveApp.getFolderById('1GV5nIwTNK9fE2OAxvO4_ue7YyYExyeb5');
      var bodydochtml = emails.createFile(subject+'.html', body, "text/html");
      var bodyId = bodydochtml.getId();
 
      // Convert the HTML to PDF
      var bodydocpdf = bodydochtml.getAs('application/pdf');
      
      // Check if there are attachments and handle them
      if (attachments.length > 0)
      {
        var new_folder = emails.createFolder(subject);
        var folder_id = new_folder.getId();
        var att_folder = DriveApp.getFolderById(folder_id);
        for (var j = 0; j < attachments.length; j++) 
        {
          att_folder.createFile(attachments[j]);
          Utilities.sleep(1000);
        }
        att_folder.createFile(bodydocpdf);
      }
      
      // Otherwise just save 
      else
        emails.createFile(bodydocpdf);

      label.removeFromThread(threads[i]);
    }
  }  
}
