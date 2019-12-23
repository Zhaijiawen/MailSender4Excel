namespace MailSender4Excel.ConfigModel
{
	/// <summary>
	/// sheet页配置
	/// </summary>
	public class SheetConfigModel
	{
		/// <summary>
		/// 起始行
		/// </summary>
		public int StartingLine
		{
			get;
			set;
		}
		/// <summary>
		/// 结束行
		/// </summary>
		public int EndLine
		{
			get;
			set;
		}
		/// <summary>
		/// 唯一标识
		/// </summary>
		public string UniquelyIdentifiesLine
		{
			get;
			set;
		}
	}
}
