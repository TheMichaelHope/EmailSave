//Author: MH Charles-Etuk
//Date: 10/15/2019
//Description: Google Script functions to save emails with certain labels applied as PDFs in a specified Google Drive folder



function save_Gmail_as_PDF()
{
  var label = GmailApp.getUserLabelByName("Save As PDF");
  
  if(label == null)
  {
    GmailApp.createLabel('Save As PDF');
  }
  
  else
  {
    var threads = label.getThreads();  
    
    for(var i = 0; i < threads.length; i++) 
    {  
      var messages = threads[i].getMessages();  
      var message = messages[0];
      var body    = message.getBody();
      var subject = message.getSubject();
      var attachments  = message.getAttachments();
      
      for(var j = 1; j<messages.length; j++)
      {
        body += messages[j].getBody();
        var temp_attach = messages[j].getAttachments();
        
        if(temp_attach.length > 0)
        {
          for(var k =0; k < temp_attach.length; k++)
            attachments.push(temp_attach[k]);
        }
      }
      
      // Create an HTML File from the Message Body
      
      
      var emails = DriveApp.getFolderById('folder_id');
      var bodydochtml = emails.createFile(subject+'.html', body, "text/html");
      var bodyId=bodydochtml.getId();
 
      // Convert the HTML to PDF
      var bodydocpdf = bodydochtml.getAs('application/pdf');
      
      if(attachments.length > 0)
      {
        emails.createFolder(subject);
        var folder = emails.getFolder(subject);
        for(var j = 0; j < attachments.length; j++) 
        {
          folder.createFile(attachments[j]);
          Utilities.sleep(1000);
        }
        folder.createFile(bodydocpdf);
      }
      
      else
        emails.createFile(bodydocpdf);

      label.removeFromThread(threads[i]);
    }
  }  
}
