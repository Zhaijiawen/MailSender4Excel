using MailSender4Excel.DataModel;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace MailSender4Excel.Util
{
	public class MailSenderUtil
	{
		static MailSenderUtil()
		{
			ServicePointManager.ServerCertificateValidationCallback = delegate
(object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
{ return true; };
		}

		private static Dictionary<string, SmtpClient> smtpClientDictField;
		private static Dictionary<string, SmtpClient> SmtpClientDict
		{
			get
			{
				if (smtpClientDictField == null)
				{
					smtpClientDictField = new Dictionary<string, SmtpClient>();
				}
				return smtpClientDictField;
			}
			set
			{
				smtpClientDictField = value;
			}
		}
		private static MailMessage BuildMessage(MailDataModel mailData)
		{
			MailAddress mailAddress = new MailAddress(mailData.From);
			MailMessage mailMessage = new MailMessage();
			mailMessage.To.Add(mailData.To);
			mailMessage.From = mailAddress;
			mailMessage.Subject = mailData.Subject;
			mailMessage.SubjectEncoding = Encoding.UTF8;
			mailMessage.Body = mailData.Body;
			mailMessage.BodyEncoding = Encoding.UTF8;
			mailMessage.Priority = MailPriority.High;
			mailMessage.IsBodyHtml = true;
			return mailMessage;
		}

		private static SmtpClient GetSmtpClient(MailDataModel mailData)
		{
			SmtpClient smtpClient;
			if (SmtpClientDict.ContainsKey(mailData.From))
			{
				smtpClient = SmtpClientDict[mailData.From];
			}
			else
			{
				smtpClient = new SmtpClient
				{
					Credentials = new NetworkCredential(mailData.From, mailData.Password),
					Host = mailData.SMTPAddress,
					EnableSsl = mailData.EnableSsl,
					Port = mailData.Port
				};
				SmtpClientDict[mailData.From] = smtpClient;
			}
			return smtpClient;
		}

		public static void Send(MailDataModel mailData)
		{
			MailMessage mailMessage = BuildMessage(mailData);
			var smtpClient = GetSmtpClient(mailData);
			smtpClient.Send(mailMessage);
		}
	}
}
