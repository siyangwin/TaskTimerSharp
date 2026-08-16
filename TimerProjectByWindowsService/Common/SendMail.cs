using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;

namespace TimerProjectByWindowsService.Common
{
    public class SendMail
    {
        public static readonly string MailHost = ConfigurationManager.AppSettings["MailHost"];                  //指定发送邮件的服务器地址或IP，如smtp.163.com
        public static readonly int Port = ParsePort(ConfigurationManager.AppSettings["Port"]);                  //指定发送邮件端口 
        public static readonly string MailAddress = ConfigurationManager.AppSettings["MailAddress"];            //发件人邮箱地址
        public static readonly string MailDisplayName = ConfigurationManager.AppSettings["MailDisplayName"];    //发件人邮箱用户名
        public static readonly string MailPassWord = ConfigurationManager.AppSettings["MailPassWord"];          //发件人邮箱密码
        public static readonly bool SSL = ParseBool(ConfigurationManager.AppSettings["SSL"], "SSL");            //指定是否需要SSL加密

        private static int ParsePort(string value)
        {
            int port;
            if (!int.TryParse(value, out port))
            {
                throw new ConfigurationErrorsException("app.config中appSettings的Port必须为数字，当前值：" + (value ?? "(未配置)"));
            }
            return port;
        }

        private static bool ParseBool(string value, string keyName)
        {
            bool result;
            if (!bool.TryParse(value, out result))
            {
                throw new ConfigurationErrorsException("app.config中appSettings的" + keyName + "必须为true/false，当前值：" + (value ?? "(未配置)"));
            }
            return result;
        }

        /// <summary>
        /// 解析收件人：按;或,分割，去空白、去空项、去掉不含@的非法地址、去重。
        /// 原版对空地址/无@地址会抛异常并导致服务崩溃。
        /// </summary>
        public static string[] ParseRecipients(string mailTo)
        {
            if (string.IsNullOrWhiteSpace(mailTo)) return new string[0];

            return mailTo.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0 && s.Contains("@"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="MailTo">收件人</param>
        /// <param name="file">附件集合</param>
        public string SendmailFile(string MailTo, string[] file, string Subject, string Body, MailPriority mailPriority = MailPriority.Normal)
        {
            if (string.IsNullOrWhiteSpace(MailHost) || string.IsNullOrWhiteSpace(MailAddress))
            {
                return "邮件服务器/发件人未配置(appSettings: MailHost、MailAddress)，无法发送邮件";
            }

            string[] recipients = ParseRecipients(MailTo);
            if (recipients.Length == 0)
            {
                return "收件人为空或全部非法，无法发送邮件";
            }

            string res = "";
            try
            {
                using (SmtpClient smtpClient = new SmtpClient())
                using (MailMessage mailMessage = new MailMessage())
                {
                    smtpClient.Host = MailHost;
                    smtpClient.EnableSsl = SSL;
                    smtpClient.Port = Port;//指定发送邮件端口 
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new System.Net.NetworkCredential(MailAddress, MailPassWord);
                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    mailMessage.BodyEncoding = Encoding.UTF8;
                    mailMessage.IsBodyHtml = true;//是否为html格式 
                    mailMessage.Priority = mailPriority;//发送邮件的优先等级 
                    mailMessage.From = new MailAddress(MailAddress, MailDisplayName);//发件人和显示发件人名称

                    for (int i = 0; i < recipients.Length; i++)
                    {
                        string displayName = recipients[i].Substring(0, recipients[i].IndexOf("@"));
                        mailMessage.To.Add(new MailAddress(recipients[i], displayName));
                    }
                    mailMessage.Subject = Subject;//邮件主题 
                    //添加邮件附件，可发送多个文件
                    if (file != null)
                    {
                        foreach (var filename in file)
                        {
                            if (!string.IsNullOrWhiteSpace(filename))
                            {
                                mailMessage.Attachments.Add(new Attachment(filename, MediaTypeNames.Application.Octet));
                            }
                        }
                    }
                    mailMessage.Body = Body;//邮件内容

                    smtpClient.Send(mailMessage);
                    res = "成功";
                }
            }
            catch (Exception ex)
            {
                res = "邮箱异常！" + ex.Message;
            }
            return res;
        }
    }
}
