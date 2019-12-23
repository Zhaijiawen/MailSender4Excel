using System.Collections.Generic;

namespace MailSender4Excel.ConfigModel
{
	/// <summary>
	/// 主配置
	/// </summary>
	public class MainConfigModel
	{
		/// <summary>
		/// 文件地址
		/// </summary>
		public string FilePath
		{
			get;
			set;
		}
		/// <summary>
		/// 成功标识符
		/// </summary>
		public string SuccessSimple
		{
			get;
			set;
		}
		/// <summary>
		/// 标识符位置
		/// </summary>
		public string SuccessSimpleLocation
		{
			get;
			set;
		}
		private IList<string> sheetNamesField;
		/// <summary>
		/// sheet页名称
		/// </summary>
		public IList<string> SheetNames
		{
			get
			{
				if (sheetNamesField == null)
				{
					sheetNamesField = new List<string>();
				}
				return sheetNamesField;
			}
			set
			{
				sheetNamesField = value;
			}
		}
		/// <summary>
		/// 主sheet页
		/// </summary>
		public string MainSheetName
		{
			get;
			set;
		}
		/// <summary>
		/// 模板路径
		/// </summary>
		public string TemplatePath
		{
			get;
			set;
		}
		/// <summary>
		/// 邮件主题参数个数
		/// </summary>
		public int BodyParamCount
		{
			get;
			set;
		}
		/// <summary>
		/// 邮件主题参数个数
		/// </summary>
		public int SubjectParamCount
		{
			get;
			set;
		}
		/// <summary>
		/// 邮件目标地址参数个数
		/// </summary>
		public int MailToParamCount
		{
			get;
			set;
		}
	}
}
