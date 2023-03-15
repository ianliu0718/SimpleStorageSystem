using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using 簡易倉儲系統.Properties;

namespace 簡易倉儲系統.EssentialTool
{
    internal class SendMail
    {
        private String SendMailAcc;
        private String SendMailPwd;
        private String _errMsg;
        public String ErrorMessage { get { return _errMsg; } }

        public SendMail(string _SendMailAcc, string _SendMailPwd)
        {
            SendMailAcc = _SendMailAcc;
            SendMailPwd = _SendMailPwd;
        }

        public Boolean Send(String Addressee, String Subject, String Body)
        {
            MailMessage mail = new MailMessage();
            NetworkCredential cred = new NetworkCredential(SendMailAcc, SendMailPwd);
            try
            {
                if (!SendMailAcc.Contains("@gmail.com"))
                {
                    _errMsg = "請使用 GMail 信箱寄信！";
                    return true;
                }

                Body = Body.Replace("\r\n", "<br>");

                //收件者
                mail.To.Add(Addressee);

                mail.Subject = Subject;
                //寄件者
                mail.From = new System.Net.Mail.MailAddress(SendMailAcc);
                mail.IsBodyHtml = true;
                mail.Body = Body;
                //設定SMTP
                SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                smtp.UseDefaultCredentials = true;
                smtp.EnableSsl = true;
                smtp.Credentials = cred;
                smtp.Port = 587;
                //送出Mail
                smtp.Send(mail);
                return false;
            }
            catch (Exception ex)
            {
                _errMsg = ex.ToString();
                return true;
            }
        }

        public static string encode(String strData)
        {
            try { return System.Convert.ToBase64String(System.Text.UTF8Encoding.UTF8.GetBytes(strData)); }
            catch { return ""; }
        }

        public static string decode(String strData)
        {
            try { return System.Text.UTF8Encoding.UTF8.GetString(System.Convert.FromBase64String(strData)); }
            catch { return ""; }
        }
    }
}
