using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace MailSender4Excel.Util
{
	public class IniFileReadUtil
	{
		#region 系统api
		[DllImport("kernel32", CharSet = CharSet.Auto)]
		private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

		[DllImport("kernel32", CharSet = CharSet.Auto)]
		private static extern long GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
		#endregion

		#region 读ini文件
		public static string ReadIniData(string section, string key, string noText, string iniFilePath)
		{
			if (File.Exists(iniFilePath))
			{
				StringBuilder temp = new StringBuilder(1024);
				GetPrivateProfileString(section, key, noText, temp, 1024, iniFilePath);
				return temp.ToString();
			}
			else
			{
				return string.Empty;
			}
		}
		#endregion

		#region 写ini文件
		public static bool WriteIniData(string section, string key, string value, string iniFilePath)
		{
			if (File.Exists(iniFilePath))
			{
				long OpStation = WritePrivateProfileString(section, key, value, iniFilePath);
				if (OpStation == 0)
				{
					return false;
				}
				else
				{
					return true;
				}
			}
			else
			{
				return false;
			}
		}
		#endregion
	}
}
