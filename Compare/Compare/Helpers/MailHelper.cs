﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;

namespace Compare.Helpers
{
    class MailHelper
    {
        internal static int SendMails()
        {
            var counter = 0;
            var sendFrom = ConfigurationManager.AppSettings["SendMailFrom"];
            var mailSubject = ConfigurationManager.AppSettings["MailSubject"];
            var host = ConfigurationManager.AppSettings["smtpHost"];
            var port = Int32.Parse(ConfigurationManager.AppSettings["smtpPort"]);
            var user = ConfigurationManager.AppSettings["smtpUser"];
            var password = ConfigurationManager.AppSettings["smtpPassword"];

            var alreadySentMails = new List<string>();

            if (Globals.Production)
            {
                foreach (var order in Globals.ExcelOrderListReady)
                {
                    if (alreadySentMails.Contains(order.Email) || string.IsNullOrEmpty(order.Email))
                    {
                        continue;
                    }

                    var mail = new MailMessage(sendFrom, order.Email);
                    mail.Subject = mailSubject;
                    mail.Body = CreateBody(order);
                    var attachement = new Attachment(order.ImageFile);
                    mail.Attachments.Add(attachement);

                    try
                    {
                        using (var smtp = new SmtpClient(host, port))
                        {
                            smtp.UseDefaultCredentials = false;
                            smtp.Credentials = new NetworkCredential(user, password);
                            smtp.EnableSsl = true;

                            smtp.Send(mail);
                        }
                        alreadySentMails.Add(order.Email);

                        counter++;
                        order.Status = Enums.MailSentStatus.MailSend;
                    }
                    catch (Exception e)
                    {
                        order.Status = Enums.MailSentStatus.MailNotSent;
                    }

                    Thread.Sleep(5000);
                }
            }

            return counter;
        }

        private static string CreateBody(Models.Order order)
        {
            if (string.IsNullOrEmpty(Globals.SelectedMailBody) || !File.Exists(Globals.SelectedMailBody))
                throw new Exception($"Filen (mailbody) {Globals.SelectedMailBody} eksisterer ikke!");

            var result = File.ReadAllText(Globals.SelectedMailBody);

            result = result.Replace("{newline}", Environment.NewLine);
            result = result.Replace("{schoolUrl}", Globals.SchoolUrl);
            result = result.Replace("{chosenImage}", order.ChosenImage.ToString());
            result = result.Replace("{firstName}", order.Firstname);

            return result;
        }
    }
}
