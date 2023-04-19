
using EmailConceptG;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Kiota.Abstractions.Authentication;

// This is a sample program for Sending Email, Receiving Emails, and Deleting Emails with Microsoft.Graph 5.5.0


string emailUsername = "office365user@domain.onmicrosoft.com"; 

// Instantiating Auth provider
IAuthenticationProvider cap = new ConsoleAppAuthenticationProvider(new OAuth2DTO()
{
    ClientId = "*****",
    TenantId = "*****",
    UserName = emailUsername,
    Password = "*****"
}, @"C:\Temp\GraphCache.txt", new[] { "Mail.ReadWrite", "Mail.Send", "SMTP.Send" });

// Pass auth provider to the new instance of graphClient
GraphServiceClient graphClient = new GraphServiceClient(cap);

// Uncomment these to test Email Sending functionality -------
//
//var body = new SendMailPostRequestBody()
//{
//    Message = new Message()
//    {
//        Subject = "Test email",
//        Body = new ItemBody()
//        {
//            Content = "This is just a <b>Test</b> for sending email with Microsoft.Graph",
//            ContentType = BodyType.Html
//        },
//        ToRecipients = new List<Recipient>()
//        {
//            new Recipient()
//            {
//                 EmailAddress = new EmailAddress()
//                 {
//                      Address = "tester1000@mailinator.com"
//                 }
//            }
//        }
//    },
//    SaveToSentItems = true
//};

// This is to send the email.
// At this point the auth provider will be called and the authentication / authorization will take place
// Authentication if it has been 20 minutes since the last authentication, authorization if it is withing 20 minutes since the last authentication
//await graphClient.Users[emailUsername].SendMail.PostAsync(body);
//Console.WriteLine($"Email sent to {testEmail} ");

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
            Console.WriteLine($"Deleting {m.ReceivedDateTime} - {m.Sender.EmailAddress.Address} - {m.Subject} ");
            
            try
            {
                await graphClient.Me.MailFolders[mailbox].Messages[m.Id].DeleteAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting email #{i} {m.Id} {m.ReceivedDateTime} {m.Subject} {ex.Message}");
            }
        }
    }
}
