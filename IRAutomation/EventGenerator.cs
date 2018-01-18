using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Newtonsoft.Json;

namespace IRAutomation
{
    class EventGenerator
    {

        protected string mailuser;
        protected string mailpass;
        protected string snuser;
        protected string snpass;

       

        public EventGenerator(string _mailuser, string _mailpass, string _snuser, string _snpass)
        {
            this.mailuser = _mailuser;
            this.mailpass = _mailpass;
            this.snuser = _snuser;
            this.snpass = _snpass;

        }

        public void generateevents(bool p)
        {

            ProofpointEvent proofpointevent;
            List<ProofpointEvent> eventlist = new List<ProofpointEvent>();
            List<string> htmlbody = new List<string>();
            
            GetInboxMessages inbox = new GetInboxMessages(mailuser, mailpass);
            htmlbody = inbox.fetch(p);
            foreach (string s in htmlbody)
            {
                proofpointevent = new ProofpointEvent();
                HtmlDocument htmldoc = new HtmlDocument();
                htmldoc.LoadHtml(s);
                var rows = htmldoc.DocumentNode.SelectNodes("//table//tr");
                if (rows != null)
                {
                    foreach (var row in rows)
                    {
                        if (row.InnerText.Replace(System.Environment.NewLine, "") != string.Empty)
                        {
                            var rowtext = row.InnerText.Trim().Replace(System.Environment.NewLine, ",").Replace("&#8212;", ",").Split(',');
                            string rowdata = rowtext[0].Replace(" ", "").ToLower();
                            string rowdata2 = rowtext[1].Replace("&lt;", "").Replace("&gt;", "");
                            if (rowdata == "attachmentsha256")
                            {
                                proofpointevent.attachmentsha256 = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "recipient" | rowdata == "recipients")
                            {
                                proofpointevent.recipients = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "url")
                            {
                                proofpointevent.url = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "category")
                            {
                                proofpointevent.category = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "condemnationtime")
                            {
                                proofpointevent.condemnationtime = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "timedelivered")
                            {
                                proofpointevent.timedelivered = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "subject")
                            {
                                proofpointevent.subject = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "sender")
                            {
                                proofpointevent.sender = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "headerfrom")
                            {
                                proofpointevent.headerfrom = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "headerreplyto")
                            {
                                proofpointevent.headerreplyto = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "message-id")
                            {
                                proofpointevent.messageid = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "senderip")
                            {
                                proofpointevent.senderip = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "messagesize")
                            {
                                proofpointevent.messagesize = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "time")
                            {
                                proofpointevent.clicktime = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "sourceip")
                            {
                                proofpointevent.sourceip = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "useragent")
                            {
                                proofpointevent.useragent = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata == "proofpoint")
                            {
                                proofpointevent.eventtitle = rowdata2;
                                rowdata = string.Empty;
                            }
                            if (rowdata != string.Empty)
                            {

                                Console.WriteLine(rowdata);
                            }

                        }
                    }//end foreach
                }//if rows
                eventlist.Add(proofpointevent);
            }

            foreach (ProofpointEvent obj in eventlist)
            {
                string _correlation_id;
                string _short_description;
                string _description = "";
                string _affected_user;
                string _category = "";

                if (obj.eventtitle == "URL Defense")
                {
                    _category = "Spear Phishing";
                    _description = "Tap Alert Details" + "\nAffected User " + obj.recipients + "\nClickTime " + obj.clicktime + "\nPhishing URL " + obj.url;
                }
                if (obj.eventtitle == "Attachment Defense")
                {
                    _category = "Malicious code activity";
                    _description = "Tap Alert Details" + "\nAffected User " + obj.recipients + "\nSubject " + obj.subject + "\nHash Value " + obj.attachmentsha256 + "\nSender " + obj.sender + "\nSender IP " + obj.senderip;
                }

                           
                _correlation_id = obj.condemnationtime.Replace("-", "").Replace(":", "").Replace("Z", "");
                _short_description = obj.eventtitle + " " + obj.category;
                _affected_user = obj.recipients;


                Console.WriteLine("   " + obj.recipients);


                var result = snow.CreateIncidentServiceNow(snuser, snpass, _correlation_id, _short_description, _description, _affected_user, _category);

                string json = JsonConvert.SerializeObject(new
                {
                    text = "The follwoing Proofpoint incident was generated from a TAP alert  "+result+"  "+_description
                });



                TeamsPost post = new TeamsPost(json);

            }
                                   
        }
    }
}

