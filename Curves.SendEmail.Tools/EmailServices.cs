using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web;
using System.Web.Hosting;

namespace Curves.SendEmail.Tools
{
    public class EmailServices
    {
        public async static Task SendMailAsync(string sendMail, string email, string subject, string message)
        {
            try
            {
                var _email = sendMail;
                var _eacc = ConfigurationManager.AppSettings["emailAccount"];
                var _epass = ConfigurationManager.AppSettings["emailPassword"];
                var _dispName = "Frank Email Service";
                MailMessage myMessage = new MailMessage();

                //Add receiver's email by constrocture parameter
                myMessage.To.Add(email);
                //Sender's email
                myMessage.From = new MailAddress(_email, _dispName);
                //Subject by constrocture parameter
                myMessage.Subject = subject;
                //Body by constrocture parameter
                myMessage.Body = message;
                myMessage.IsBodyHtml = true;
                //夾帶檔案
                //Attachment attachment;
                //attachment = new Attachment(HostingEnvironment.MapPath("~/Content/手開發票範例檔(韻智股份有限公司).pdf"));
                //myMessage.Attachments.Add(attachment);

                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.EnableSsl = true;
                    smtp.Host = "smtp.sendgrid.net";
                    smtp.Port = 587;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(_eacc, _epass);
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.SendCompleted += (s, e) => { smtp.Dispose(); };
                    await smtp.SendMailAsync(myMessage);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}