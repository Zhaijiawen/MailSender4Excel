using System.ComponentModel;

namespace MailSender4Excel.ViewModel
{
	public class MainWindowViewModel : INotifyPropertyChanged
	{
		public event PropertyChangedEventHandler PropertyChanged;

		private string textField;
		/// <summary>
		/// 输出文本
		/// </summary>
		public string Text
		{
			get
			{
				return textField;
			}
			set
			{
				if (textField != value)
				{
					textField = value;
					PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
				}
			}
		}

		private string configFilePathField;
		/// <summary>
		/// 配置文件路径
		/// </summary>
		public string ConfigFilePath
		{
			get
			{
				return configFilePathField;
			}
			set
			{
				if (configFilePathField != value)
				{
					configFilePathField = value;
					PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ConfigFilePath)));
				}
			}
		}

		private string testMailAddressField;
		/// <summary>
		/// 测试邮件地址
		/// </summary>
		public string TestMailAddress
		{
			get
			{
				return testMailAddressField;
			}
			set
			{
				if (testMailAddressField != value)
				{
					testMailAddressField = value;
					PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TestMailAddress)));
				}
			}
		}
	}
}
