using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.Net;
using System.Net.Sockets;
using MailKit.Security;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace IRAutomation
{
    class GetInboxMessages
    {


              
        private string user;
        private string password;
        private string emailhost = "outlook.office365.com";
        private int sslport = 993;

        public GetInboxMessages(string mailboxuser, string mailboxpassword) {

            this.user = mailboxuser;
            this.password = mailboxpassword;
        }

        public List<string> fetch(bool p)
        {
            List<string> htmlbody = new List<string>();
            bool proxy = p;

            if (proxy)
            {
                


                using (var client = new ImapClient())
                {

                    client.SslProtocols = System.Security.Authentication.SslProtocols.Tls11;

                    

                    client.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => {
                        if (sslPolicyErrors == SslPolicyErrors.None)
                            return true;
                                               
                        if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
                        {
                            if (chain != null && chain.ChainStatus != null)
                            {
                                foreach (var status in chain.ChainStatus)
                                {
                                    if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
                                    {
                                         
                                        continue;
                                    }
                                    else if (status.Status != X509ChainStatusFlags.NoError)
                                    {
                                        
                                        return true;
                                    }
                                }
                            }

                            
                            return true;
                        }

                        return true;
                    };






                    //client.Connect(emailhost, 993, true);
                    client.Connect("127.0.0.1",2022,true);
                    

                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    client.Authenticate(user, password);
                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);
                    var items = inbox.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.Flags);
                    foreach (var mailItem in items)
                    {
                        var message = inbox.GetMessage(mailItem.UniqueId);
                        if (message.From.ToString() == "\"Proofpoint TAP\" <tap-notifications@proofpoint.com>")
                        {
                            htmlbody.Add(message.HtmlBody);
                        }

                        var folder = client.GetFolder("Archive");  //impliment after testing
                        inbox.MoveTo(mailItem.UniqueId, folder);
                    }


                    client.Disconnect(true);
                    return htmlbody;

                }
            }

            else
            {

                using (var client = new ImapClient())
                {

                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    client.Connect(emailhost, sslport, true);
                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    client.Authenticate(user, password);
                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);
                    var items = inbox.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.Flags);
                    foreach (var mailItem in items)
                    {
                        var message = inbox.GetMessage(mailItem.UniqueId);
                        if (message.From.ToString() == "\"Proofpoint TAP\" <tap-notifications@proofpoint.com>")
                        {
                            htmlbody.Add(message.HtmlBody);
                        }

                        var folder = client.GetFolder("Archive");  
                        inbox.MoveTo(mailItem.UniqueId, folder);
                    }


                    client.Disconnect(true);
                    return htmlbody;

                }
            }



        } 
        }
}
