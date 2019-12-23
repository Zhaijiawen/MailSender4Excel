namespace MailSender4Excel.ConfigModel
{
	/// <summary>
	/// 邮箱配置
	/// </summary>
	public class MailConfigModel
	{
		/// <summary>
		/// 目标邮件地址（公式）
		/// </summary>
		public string MailTo
		{
			get;
			set;
		}
		/// <summary>
		/// 发送邮件的地址
		/// </summary>
		public string MailAddress
		{
			get;
			set;
		}
		/// <summary>
		/// 密码
		/// </summary>
		public string MailPassword
		{
			get;
			set;
		}
		/// <summary>
		/// 主题
		/// </summary>
		public string MailSubject
		{
			get;
			set;
		}
		/// <summary>
		/// smtp地址
		/// </summary>
		public string SMTPAddress
		{
			get;
			set;
		}

		/// <summary>
		/// 发送邮件端口
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
