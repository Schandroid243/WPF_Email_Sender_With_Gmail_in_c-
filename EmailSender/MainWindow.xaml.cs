
using Google.Apis.Auth;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using MailKit.Security;
using MimeKit;
using MimeKit.Encodings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EmailSender
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string accessToken;
        private string username;
        private GmailService service;
        private UserCredential credentials;
        private string userId;
        public MainWindow()
        {
            InitializeComponent();
            GetGmailCredentitals();
            if (SenderEmail.Text == null)
            {
                Send.IsEnabled = false;
            }
            else
            {
                Send.IsEnabled = true;
            }
        }

        private async void DisconnectAccount()
        {
            try
            {
                if (!string.IsNullOrEmpty(accessToken))   
                {
                    await credentials.RevokeTokenAsync(CancellationToken.None);//This code line is used to cancel the access token then you'll restart the app it will request you to connect again
                    var jwtPayload = GoogleJsonWebSignature.ValidateAsync(credentials.Token.IdToken).Result;//this one has no use here Lol just to test something but don't mind ! 😂
                }
            }
            catch (Exception ex)   
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private async void GetGmailCredentitals()
        {
           try 
            {
                credentials = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = "Here enter your cliend id from the google console",
                    ClientSecret = "Here enter your secret code which will be used as the email account password"
                },
                new[] { "email", "profile", "http://mail.google.com/" },
                "user",
                CancellationToken.None
                );

                await credentials.RefreshTokenAsync(CancellationToken.None);

                var jwtPayload = GoogleJsonWebSignature.ValidateAsync(credentials.Token.IdToken).Result;
                
                
                SenderEmail.Text = jwtPayload.Email;
                username = jwtPayload.GivenName;
                accessToken = credentials.Token.AccessToken.ToString();

                userId = credentials.UserId.ToString();

                service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credentials,
                    ApplicationName = "Here enter the App name you've created in the google console"
                });
            } 
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(username, SenderEmail.Text));
                message.To.Add(new MailboxAddress("Dear", DestinationEmail.Text));
                message.Subject = "Celebration";

                var multipart = new Multipart("mixed");

                var body = new TextPart("plain")
                {
                    Text = DestinationBody.Text
                };

                multipart.Add(body);

                var pdfAttachment = new MimeKit.MimePart("application", "pdf")//Change the MimePart according to the attachement you want to send
                {
                    Content = new MimeKit.MimeContent(File.OpenRead(@"Here enter the path of your pdf file"), ContentEncoding.Default),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = "brochureAurtech.pdf"
                };

                multipart.Add(pdfAttachment);

                message.Body = multipart;

                var messageText = message.ToString();
                byte[] messageBytes = Encoding.UTF8.GetBytes(messageText);
                string base64Message = Convert.ToBase64String(messageBytes);

                string urlSafeBase64Message = base64Message.Replace('+', '-').Replace('/', '_');

                var email = service.Users.Messages.Send(new Message
                {
                    Raw = urlSafeBase64Message
                }, "me").Execute();

                MessageBox.Show($"Message sent !");
            }
            catch (Exception ex)   
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void Send_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            DisconnectAccount();
        }
    }
}
