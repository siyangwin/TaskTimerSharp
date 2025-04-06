using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;

namespace TimerProjectByWindowsService.Common
{
    public class SendMail
    {
        public static readonly string MailHost = ConfigurationSettings.AppSettings["MailHost"];                  //指定发送邮件的服务器地址或IP，如smtp.163.com
        public static readonly int Port = Convert.ToInt32(ConfigurationSettings.AppSettings["Port"]);            //指定发送邮件端口 
        public static readonly string MailAddress = ConfigurationSettings.AppSettings["MailAddress"];            //发件人邮箱地址
        public static readonly string MailDisplayName = ConfigurationSettings.AppSettings["MailDisplayName"];    //发件人邮箱用户名
        public static readonly string MailPassWord = ConfigurationSettings.AppSettings["MailPassWord"];          //发件人邮箱密码
        public static readonly bool SSL = Convert.ToBoolean(ConfigurationSettings.AppSettings["SSL"]);          //指定是否需要SSL加密

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="MailTo">收件人</param>
        /// <param name="file">附件集合</param>
        public string SendmailFile(string MailTo, string[] file, string Subject, string Body, MailPriority mailPriority= MailPriority.Normal)
        {
            string res = "";
            SmtpClient smtpClient = new SmtpClient();
            MailMessage mailMessage = new MailMessage();
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

            //分割数组，并去除重复
            string[] MailToList = MailTo.Split(';').Distinct().ToArray();
            for (int i = 0; i < MailToList.Length; i++)
            {
                mailMessage.To.Add(new MailAddress(MailToList[i], MailToList[i].Substring(0, MailToList[i].IndexOf("@"))));  //收件人和收件人显示姓名
            }
            //mailMessage.To.Add(MailTo);//收件人
            mailMessage.Subject = Subject;//邮件主题 
            mailMessage.Attachments.Clear();
            //添加邮件附件，可发送多个文件
            foreach (var filename in file)
            {
                mailMessage.Attachments.Add(new Attachment(filename, MediaTypeNames.Application.Octet));
            }
            mailMessage.Body = Body;//邮件内容

            try
            {
                smtpClient.Send(mailMessage);
                res = "成功";
            }
            catch (Exception ex)
            {
                res = "邮箱异常！" + ex.Message;
                //throw new Exception("邮箱异常！" + ex.Message);
            }
            return res;
        }

    }
}
