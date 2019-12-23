namespace MailSender4Excel.DataModel
{
	public class MailDataModel
	{
		/// <summary>
		/// 发送者
		/// </summary>
		public string From { get; set; }

		/// <summary>
		/// 收件人
		/// </summary>
		public string To { get; set; }

		/// <summary>
		/// 标题
		/// </summary>
		public string Subject { get; set; }

		/// <summary>
		/// 正文
		/// </summary>
		public string Body { get; set; }

		/// <summary>
		/// 发件人密码
		/// </summary>
		public string Password { get; set; }

		/// <summary>
		/// SMTP邮件服务器
		/// </summary>
		public string SMTPAddress { get; set; }

		/// <summary>
		/// 正文是否是html格式
		/// </summary>
		public bool IsbodyHtml { get; set; }

		/// <summary>
		/// 邮件端口
		/// </summary>
		public int Port
		{
			get;
			set;
		}

		/// <summary>
		/// 启用ssl
		/// </summary>
		public bool EnableSsl
		{
			get;
			set;
		}
	}
}
