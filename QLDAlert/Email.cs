using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace QLDAlert {
    class Email {
        const string INPUT_DATE_FORMAT = "M/d/yyyy";

        private Log log;

        public Email(Log log) {
            this.log = log;
        }

        /// <summary>
        /// Load mail content from file.
        /// </summary>
        /// <param name="mailPath"></param>
        /// <returns></returns>
        public string writeEmail(string mailPath) {
            string emailContent = "";

            if (!File.Exists(mailPath)) {
                log.appendLog("Cannot find the mail file.", -1);
            }

            try {
                StreamReader reader = new StreamReader(File.OpenRead(mailPath));

                while (!reader.EndOfStream) {
                    emailContent += reader.ReadLine();
                }

                reader.Close();
            } catch (Exception ex) {
                log.appendLog("Failed to open mail file.\n" + ex.ToString(), -1);
            }

            return emailContent;
        }

        /// <summary>
        /// Get recipient list from file and send alert email.
        /// </summary>
        /// <param name="mailListPath"></param>
        /// <param name="mailPath"></param>
        /// <param name="reportPath"></param>
        // void sendAlertMail(string mailListPath, string mailPath, string reportPath) {
        //    string ewsUrl = "CANTEEN_EXCHANGE_SERVER";
        //    string userName = "EMAIL_USERNAME";
        //    string password = "EMAIL_PASSWORD";
        //    string title = "Weekly Update " + DateTime.Today.ToString(INPUT_DATE_FORMAT);

        //    if (!File.Exists(mailListPath)) {
        //        log.appendLog("Can't find mail list.", -1);
        //        exitProgram();
        //    }

        //    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        //    service.Url = new Uri(ewsUrl);
        //    service.UseDefaultCredentials = true;
        //    service.Credentials = new WebCredentials(userName, password);

        //    EmailMessage message = new EmailMessage(service);
        //    message.Subject = title;
        //    message.Body = writeEmail(mailPath);

        //    try {
        //        log.appendLog("Getting recipient list.", 1);

        //        StreamReader reader = new StreamReader(File.OpenRead(mailListPath));

        //        while (!reader.EndOfStream) {
        //            message.CcRecipients.Add(reader.ReadLine().Trim());
        //        }
        //        reader.Close();
        //    } catch (Exception ex) {
        //        log.appendLog("Failed reading mail list.\n" + ex.ToString(), -1);
        //        exitProgram();
        //    }

        //    message.Attachments.AddFileAttachment(reportPath);

        //    log.appendLog("Sending Email.", 1);
        //    Console.WriteLine("This may take a while. Please wait.....");

        //    try {
        //        message.SendAndSaveCopy();
        //    } catch (Exception ex) {
        //        log.appendLog("Failed to send the mail.\n" + ex.ToString(), -1);
        //        exitProgram();
        //    }

        //    log.appendLog("Email sent.", 1);
        //}
        
        public void sendAlertMail(string mailPath, string reportPath) {
            string from = "EMAIL_FROM";
            string to = "EMAIL_TO";

            log.appendLog("Preparing email....", 1);

            MailMessage mail = new MailMessage();
            mail.To.Add(to);
            mail.From = new MailAddress(from, "CanTeen QLD Weekly Alert", System.Text.Encoding.UTF8);
            mail.Subject = "Weekly Update " + DateTime.Today.ToString(INPUT_DATE_FORMAT);
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.Body = writeEmail(mailPath);
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;
            mail.Attachments.Add(new Attachment(reportPath));

            SmtpClient client = new SmtpClient();
            //Add the Creddentials- use your own email id and password

            client.Credentials = new NetworkCredential(from, "EMAIL");

            client.Port = 587; // Gmail works on this port
            client.Host = "ctn-exch-01.canteen.org.au";
            client.EnableSsl = false; //Gmail works on Server Secured Layer
            try {
                log.appendLog("Sending email. This might take a few seconds....", 1);
                client.Send(mail);
                mail.Attachments.Dispose();
                log.appendLog("Email sent successfully.", 1);
            } catch (Exception ex) {
                Exception ex2 = ex;
                string errorMessage = "Eorro Message\n";
                while (ex2 != null) {
                    errorMessage += ex2.ToString();
                    ex2 = ex2.InnerException;
                }
                log.appendLog(errorMessage, -1);
            } // end try 
        }      
    }
}
