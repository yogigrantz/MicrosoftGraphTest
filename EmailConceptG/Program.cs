using EmailConceptG;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Kiota.Abstractions.Authentication;
using MimeKit.Utils;

// This is a sample program for Sending Email, Receiving Emails, and Deleting Emails with Microsoft.Graph 5.6.0


string emailUsername = "tester@somedomainOffice365.com"; 

// Instantiating Auth provider
IAuthenticationProvider cap = new ConsoleAppAuthenticationProvider(new OAuth2DTO()
{
    ClientId = "**********",
    TenantId = "**********",
    UserName = emailUsername,
    Password = "**********"
}, @"C:\Temp\GraphCache.txt", new[] { "Mail.ReadWrite", "Mail.Send", "SMTP.Send" });

// Pass auth provider to the new instance of graphClient
GraphServiceClient graphClient = new GraphServiceClient(cap);

GraphSendEmailDTO se = new GraphSendEmailDTO()
{
    To = new string[] { "tester1000@mailinator.com" },
    Subject = "This is a test email",
    Body = "This is a <b>test email</b> with attachment and linked resources <br /><img src='thankyou.gif' />",
    IsHtml = true,
    AttachmentFileNames = new string[]
    {
        "vw.jpg", "Car.jpg"
    },
    LinkedResourceFileNames = new string[] { "thankyou.gif"}
};

var body = new SendMailPostRequestBody()
{
    Message = new Message()
    {
        Subject = se.Subject,
        Body = new ItemBody()
        {
            Content = se.Body,
            ContentType = se.IsHtml ? BodyType.Html : BodyType.Text
        },
        ToRecipients = new List<Recipient>()
        {
            new Recipient()
            {
                 EmailAddress = new EmailAddress()
                 {
                      Address = se.To[0]
                 }
            }
        }
    },
    SaveToSentItems = true
};

if (se.LinkedResourceFileNames != null)
{
    foreach (string a in se.LinkedResourceFileNames)
    {
        if (System.IO.File.Exists(a))
        {
            FileInfo fi = new FileInfo(a);
            FileAttachment fileAttachment = new FileAttachment()
            {
                Name = fi.Name,
                ContentBytes = System.IO.File.ReadAllBytes(a)
            };

            if (MimeTypes.MimeTypesDictionary.TryGetValue(fi.Extension, out string mimetype))
                fileAttachment.ContentType = mimetype;
            else
                fileAttachment.ContentType = "image/jpeg";
            fileAttachment.Size = (int)fi.Length;
            fileAttachment.ContentId = MimeUtils.GenerateMessageId();

            body.Message.Body.Content = body.Message.Body.Content.Replace(a, $"cid:{fileAttachment.ContentId}");

            if (body.Message.Attachments == null)
                body.Message.Attachments = new List<Attachment>();

            body.Message.Attachments.Add(fileAttachment);
        }
    }
}
if (se.AttachmentFileNames != null)
{
    foreach (string a in se.AttachmentFileNames)
    {
        if (System.IO.File.Exists(a))
        {
            FileInfo fi = new FileInfo(a);
            FileAttachment fileAttachment = new FileAttachment()
            {
                Name = fi.Name,
                ContentBytes = System.IO.File.ReadAllBytes(a)
            };

            if (MimeTypes.MimeTypesDictionary.TryGetValue(fi.Extension, out string mimetype))
                fileAttachment.ContentType = mimetype;
            else
                fileAttachment.ContentType = "text/plain";
            fileAttachment.Size = (int)fi.Length;

            if (body.Message.Attachments == null)
                body.Message.Attachments = new List<Attachment>();

            body.Message.Attachments.Add(fileAttachment);

        }
    }

}
// This is to send the email.
// At this point the auth provider will be called and the authentication / authorization will take place
// Authentication if it has been 20 minutes since the last authentication, authorization if it is withing 20 minutes since the last authentication
try
{
    await graphClient.Users[emailUsername].SendMail.PostAsync(body);
    Console.WriteLine($"email to {body.Message.ToRecipients[0].EmailAddress.Address} has been sent");
}
catch (Exception ex1)
{
    Console.WriteLine(ex1.Message);
}
//-----------------------------------------------------------------------------------------

// This is to receive and delete emails
int maxNbrOfEmails = 1000;
string[] scopes = new[] { "Mail.ReadWrite" };

await DeleteEmails(graphClient, maxNbrOfEmails, "INBOX");

await DeleteEmails(graphClient, maxNbrOfEmails, "DeletedItems");

await DeleteEmails(graphClient, maxNbrOfEmails, "SentItems");

static async Task DeleteEmails(GraphServiceClient graphClient, int maxNbrOfEmails, string mailbox)
{
    MessageCollectionResponse? messages = await graphClient.Me.MailFolders[mailbox].Messages.GetAsync(x => { x.QueryParameters.Top = maxNbrOfEmails; });

    if (messages?.Value.Count > 0)
    {
        int i = 0;
        foreach (Message m in messages.Value)
        {
            i++;
            Console.WriteLine($"Deleting [{mailbox}] {m.ReceivedDateTime} - {m.Sender.EmailAddress.Address} - {m.Subject} ");
            
            try
            {
                await graphClient.Me.MailFolders[mailbox].Messages[m.Id].DeleteAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting email from [{mailbox}] #{i} {m.Id} {m.ReceivedDateTime} {m.Subject} {ex.Message}");
            }
        }
    }
}
