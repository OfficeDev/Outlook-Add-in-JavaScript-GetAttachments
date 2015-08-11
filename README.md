# Outlook add-in: Get attachments from an Exchange server

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary
This sample shows you how to get attachments from an Exchange mailbox.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - A computer running Exchange 2013 with at least one email account, or an Office 365 account. You can sign up for [an Office 365 Developer subscription](http://aka.ms/o365-android-connect-signup) and get an Office 365 account through it.
  - Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

- AttachmentExampleManifest.xml: The manifest file for the Outlook add-in.
- AppRead\Home\Home.html: The HTML user interface for the mail add-in for Outlook.
- AppRead\Home\Home.js: The JavaScript file that handles sending the attachment information to the remote Attachment service included with this sample.

The AttachmentService project defines a REST service by using the WCF API. The project contains the following files:

- Controllers\AttachmentServiceController.cs: The service object that provides the business logic for the sample service.
- Models\ServiceRequest: The object that represents a web request. The contents of the object are created from a JSON request object sent from your mail add-in.
- Models\Attachment.cs: The utility object that helps deserialize the JSON object that is sent by the mail add-in.
- Models\AttachmentDetails.cs: The object that represents the details of each attachment. It provides a .NET Framework object that matches the mail add-in's `AttachmentDetails` object.
- Models\ServiceResponse: The object that represents a response from the web service. The contents of the object are serialized to a JSON object when they are sent back to the mail add-in.
- Web.config: Binds the sample service to the web server endpoint.



<a name="codedescription"></a>
##Description of the code

This sample shows you how to retrieve attachments from a web service that supports your mail add-in. For example, you can create a service that uploads photos to a sharing site, or a service that stores documents into a repository. The service gets the attachments directly from the Exchange server, and doesn't require the client to perform extra processing to get the attachment and then send it along to the service.

The sample has two parts. The first part, the mail app, runs in the email client. The mail add-in is shown whenever a message or an appointment is the active item. When you select the **Test attachments** button, the mail add-in sends details about the attachment to the web service that processes the request. The service uses the following steps to process attachments:

- Sends a [GetAttachment](http://msdn.microsoft.com/en-us/library/aa494316(v=exchg.150).aspx) operation request to the Exchange server that hosts the mailbox. The server responds by sending the attachment to the service. In this sample, the service simply writes the XML from the server to trace output.
- Returns the number of attachments processed to the mail app.



<a name="build"></a>
## Build and debug ##
**Note**: The mail add-in will be activated on any email message in the user's Inbox that has one or more attachments. You can make it easier to test the add-in by sending one or more email messages to your test account before you run the sample add-in.

1. Open the solution in Visual Studio. Press F5 to build and deploy the sample add-in.
2. Connect to an Exchange account by providing the email address and password for an Exchange 2013 server.
3. Allow the server to configure the mail account.
4. Log on to the email account by entering the account name and password. 
5. Select a message in the Inbox.
6. Wait for the add-in bar to appear over the message.
7. In the add-in bar, click **AttachmentExample**.
8. When the mail add-in appears, click the **TestAttachments** button to send a request to the Exchange server.
9. The server will respond with the number of attachments processed for the item. This should equal the number of attachments that the item contains.

<a name="troubleshooting"></a>
##Troubleshooting
The following are common errors that can occur when you use Outlook Web App to test a mail add-in for Outlook:

- The add-in bar does not appear when a message is selected. If this occurs, restart the application by selecting **Debug – Stop Debugging** in the Visual Studio window, then press F5 to rebuild and deploy the add-in. 
- Changes to the JavaScript code may not be picked up when you deploy and run the add-in. If the changes are not picked up, clear the cache on the web browser by selecting **Tools – Internet options** and clicking the **Delete…** button. Delete the temporary Internet files and then restart the add-in. 

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Web API: The Official Microsoft ASP.NET Site](http://www.asp.net/web-api)
- [How to: Get attachments from an Exchange server](http://msdn.microsoft.com/en-us/library/dn148008.aspx)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
