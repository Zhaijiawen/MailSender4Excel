using MailSender4Excel.ConfigModel;
using MailSender4Excel.DataModel;
using MailSender4Excel.Util;
using MailSender4Excel.ViewModel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace MailSender4Excel
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : System.Windows.Window
	{
		/// <summary>
		/// 视图模型
		/// </summary>
		public MainWindowViewModel ViewModel
		{
			get
			{
				return DataContext as MainWindowViewModel;
			}
			set
			{
				DataContext = value;
			}
		}
		/// <summary>
		/// 主要配置
		/// </summary>
		private MainConfigModel MainConfigModel
		{
			get;
			set;
		}
		/// <summary>
		/// 邮件配置
		/// </summary>
		private MailConfigModel MailConfigModel
		{
			get;
			set;
		}

		private Dictionary<string, SheetConfigModel> sheetConfigModelsField;
		/// <summary>
		/// sheet配置
		/// </summary>
		private Dictionary<string, SheetConfigModel> SheetConfigModels
		{
			get
			{
				if (sheetConfigModelsField == null)
				{
					sheetConfigModelsField = new Dictionary<string, SheetConfigModel>();
				}
				return sheetConfigModelsField;
			}
			set
			{
				sheetConfigModelsField = value;
			}
		}

		private IList<string> bodyParamsField;
		/// <summary>
		/// 邮件体参数
		/// </summary>
		private IList<string> BodyParams
		{
			get
			{
				if (bodyParamsField == null)
				{
					bodyParamsField = new List<string>();
				}
				return bodyParamsField;
			}
			set
			{
				bodyParamsField = value;
			}
		}

		private IList<string> subjectParamsField;
		/// <summary>
		/// 主题参数
		/// </summary>
		private IList<string> SubjectParams
		{
			get
			{
				if (subjectParamsField == null)
				{
					subjectParamsField = new List<string>();
				}
				return subjectParamsField;
			}
			set
			{
				subjectParamsField = value;
			}
		}

		private IList<string> mailToParamsField;
		/// <summary>
		/// 收件人参数
		/// </summary>
		private IList<string> MailToParams
		{
			get
			{
				if (mailToParamsField == null)
				{
					mailToParamsField = new List<string>();
				}
				return mailToParamsField;
			}
			set
			{
				mailToParamsField = value;
			}
		}

		private Dictionary<string, object[,]> paramsDictField;
		/// <summary>
		/// 参数值集合
		/// </summary>
		private Dictionary<string, object[,]> ParamsDict
		{
			get
			{
				if (paramsDictField == null)
				{
					paramsDictField = new Dictionary<string, object[,]>();
				}
				return paramsDictField;
			}
			set
			{
				paramsDictField = value;
			}
		}

		private StringBuilder outPutTextField;
		/// <summary>
		/// 输出文本
		/// </summary>
		private StringBuilder OutPutText
		{
			get
			{
				if (outPutTextField == null)
				{
					outPutTextField = new StringBuilder();
				}
				return outPutTextField;
			}
			set
			{
				outPutTextField = value;
			}
		}

		/// <summary>
		/// 工作簿
		/// </summary>
		public OfficeExcel.Workbook Workbook
		{
			get;
			set;
		}

		private Dictionary<string, OfficeExcel.Worksheet> sheetName2ExcelSheetField;
		/// <summary>
		/// sheet字典 key-sheetName value-sheet
		/// </summary>
		public Dictionary<string, OfficeExcel.Worksheet> SheetName2ExcelSheet
		{
			get
			{
				if (sheetName2ExcelSheetField == null)
				{
					sheetName2ExcelSheetField = new Dictionary<string, Worksheet>();
				}
				return sheetName2ExcelSheetField;
			}
			set
			{
				sheetName2ExcelSheetField = value;
			}
		}

		public MainWindow()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			ViewModel = new MainWindowViewModel();
		}

		private void ButtonStart_Click(object sender, RoutedEventArgs e)
		{
			BuildAndSendMails();
		}

		/// <summary>
		/// 加载配置文件
		/// </summary>
		/// <returns></returns>
		private bool LoadConfig()
		{
			string configPath = ViewModel.ConfigFilePath;
			if (File.Exists(configPath))
			{
				try
				{
					MainConfigModel = new MainConfigModel
					{
						FilePath = IniFileReadUtil.ReadIniData("Main", "FilePath", null, configPath),
						SuccessSimple = IniFileReadUtil.ReadIniData("Main", "SuccessSimple", null, configPath),
						SuccessSimpleLocation = IniFileReadUtil.ReadIniData("Main", "SuccessSimpleLocation", null, configPath),
						SheetNames = IniFileReadUtil.ReadIniData("Main", "SheetNames", null, configPath).Split(',').ToList(),
						MainSheetName = IniFileReadUtil.ReadIniData("Main", "MainSheetName", null, configPath),
						TemplatePath = IniFileReadUtil.ReadIniData("Main", "TemplatePath", null, configPath),
						BodyParamCount = int.Parse(IniFileReadUtil.ReadIniData("Main", "BodyParamCount", null, configPath)),
						SubjectParamCount = int.Parse(IniFileReadUtil.ReadIniData("Main", "SubjectParamCount", null, configPath)),
						MailToParamCount = int.Parse(IniFileReadUtil.ReadIniData("Main", "MailToParamCount", null, configPath)),
					};
					MailConfigModel = new MailConfigModel
					{
						MailTo = IniFileReadUtil.ReadIniData("Mail", "MailTo", null, configPath),
						MailAddress = IniFileReadUtil.ReadIniData("Mail", "MailAddress", null, configPath),
						MailPassword = IniFileReadUtil.ReadIniData("Mail", "MailPassword", null, configPath),
						MailSubject = IniFileReadUtil.ReadIniData("Mail", "MailSubject", null, configPath),
						SMTPAddress = IniFileReadUtil.ReadIniData("Mail", "SMTPAddress", null, configPath),
						Port = int.Parse(IniFileReadUtil.ReadIniData("Mail", "Port", null, configPath)),
						EnableSsl = bool.Parse(IniFileReadUtil.ReadIniData("Mail", "EnableSsl", null, configPath)),
					};
					foreach (var sheetName in MainConfigModel.SheetNames)
					{
						SheetConfigModel sheetConfigModel = new SheetConfigModel
						{
							StartingLine = int.Parse(IniFileReadUtil.ReadIniData(sheetName, "StartingLine", null, configPath)),
							EndLine = int.Parse(IniFileReadUtil.ReadIniData(sheetName, "EndLine", null, configPath)),
							UniquelyIdentifiesLine = IniFileReadUtil.ReadIniData(sheetName, "UniquelyIdentifiesLine", null, configPath),
						};
						SheetConfigModels[sheetName] = sheetConfigModel;
					}
					for (int i = 0; i < MainConfigModel.BodyParamCount; i++)
					{
						BodyParams.Add(IniFileReadUtil.ReadIniData("BodyParams", i.ToString(), null, configPath));
					}
					for (int i = 0; i < MainConfigModel.SubjectParamCount; i++)
					{
						SubjectParams.Add(IniFileReadUtil.ReadIniData("SubjectParams", i.ToString(), null, configPath));
					}
					for (int i = 0; i < MainConfigModel.MailToParamCount; i++)
					{
						MailToParams.Add(IniFileReadUtil.ReadIniData("MailToParams", i.ToString(), null, configPath));
					}
					//加载工作簿相关内容
					Workbook = LoadWorkbook(MainConfigModel.FilePath);
					foreach (var sheetName in MainConfigModel.SheetNames)
					{
						SheetName2ExcelSheet[sheetName] = LoadWorksheet(Workbook, sheetName);
					}
					PushMessage("读取成功。\n");
					return true;
				}
				catch (Exception ex)
				{
					PushMessage("读取失败。\n");
					PushMessage(ex.Message + "\n");
					return false;
				}
			}
			else
			{
				PushMessage("找不到配置文件！\n");
				return false;
			}

		}
		/// <summary>
		/// 加载工作簿
		/// </summary>
		/// <param name="filePath"></param>
		/// <returns></returns>
		private OfficeExcel.Workbook LoadWorkbook(string filePath)
		{
			OfficeExcel.Application application = new OfficeExcel.Application();
			OfficeExcel.Workbook workbook = application.Workbooks.Open(filePath);
			return workbook;
		}
		/// <summary>
		/// 加载工作表
		/// </summary>
		/// <param name="workbook"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		private OfficeExcel.Worksheet LoadWorksheet(OfficeExcel.Workbook workbook, string sheetName)
		{
			OfficeExcel.Worksheet worksheet = workbook.Worksheets[sheetName];
			return worksheet;
		}

		/// <summary>
		/// 构建邮件
		/// </summary>
		/// <returns></returns>
		private void BuildAndSendMails()
		{
			if (Workbook != null)
			{
				try
				{
					//主sheet配置
					SheetConfigModel mainSheetConfig = SheetConfigModels[MainConfigModel.MainSheetName];
					//模板字符串				
					string bodyHtmlString = File.ReadAllText(MainConfigModel.TemplatePath);
					//失败个数
					int faildCount = 0;
					//邮件总数
					int mailCount = mainSheetConfig.EndLine - mainSheetConfig.StartingLine + 1;
					//成功的标识
					object[,] successSimple = ReadRangeValues(MainConfigModel.MainSheetName + "." + MainConfigModel.SuccessSimpleLocation);
					//回写的标识地址
					IList<string> mailDataSuccessLocation = new List<string>();

					for (int i = mainSheetConfig.StartingLine; i <= mainSheetConfig.EndLine; i++)
					{
						PushMessage("正在处理(" + (i - mainSheetConfig.StartingLine + 1) + "/" + mailCount + ")\n");
						try
						{
							if (!(MainConfigModel.SuccessSimple == (successSimple[i - mainSheetConfig.StartingLine + 1, 1] == null ? string.Empty : successSimple[i - mainSheetConfig.StartingLine + 1, 1].ToString())))
							{
								MailDataModel mailDataModel = new MailDataModel
								{
									From = MailConfigModel.MailAddress,
									Password = MailConfigModel.MailPassword,
									SMTPAddress = MailConfigModel.SMTPAddress,
									IsbodyHtml = true,
									To = FormatString(CalcParamValues(MailToParams, i), MailConfigModel.MailTo),
									Subject = FormatString(CalcParamValues(SubjectParams, i), MailConfigModel.MailSubject),
									Body = FormatString(CalcParamValues(BodyParams, i), bodyHtmlString),
									Port = MailConfigModel.Port,
									EnableSsl = MailConfigModel.EnableSsl,
								};
								MailSenderUtil.Send(mailDataModel);
								mailDataSuccessLocation.Add(MainConfigModel.SuccessSimpleLocation + i);
							}
						}
						catch (Exception ex)
						{
							PushMessage(ex.Message + "\n");
							Interlocked.Increment(ref faildCount);
							continue;
						}
					}
					//回写成功标志
					foreach (var range in GetMergeRanges(LoadWorksheet(Workbook, MainConfigModel.MainSheetName), mailDataSuccessLocation))
					{
						range.Value2 = MainConfigModel.SuccessSimple;
					}

					PushMessage("处理完毕。" + "\n");
					PushMessage("总数量" + (mailCount) + "\n");
					PushMessage("失败数量" + (faildCount) + "\n");
				}
				catch (Exception ex)
				{
					PushMessage("出现错误:" + ex.Message + "\n");
				}
			}
		}

		/// <summary>
		/// 计算参数
		/// </summary>
		/// <param name="paramKeys"></param>
		/// <param name="rowIndex"></param>
		/// <returns></returns>
		private IList<string> CalcParamValues(IList<string> paramKeys, int rowIndex)
		{
			SheetConfigModel mainSheetConfig = SheetConfigModels[MainConfigModel.MainSheetName];
			IList<string> result = new List<string>();
			foreach (var paramKey in paramKeys)
			{
				object[,] paramArray = ReadRangeValues(paramKey);
				string[] vs = paramKey.Split('.');

				if (vs.Length == 3)
				{
					result.Add(GetValue(1, paramArray));
				}
				else
				{
					if (vs[0] == MainConfigModel.MainSheetName)
					{
						string value = paramArray[rowIndex - mainSheetConfig.StartingLine + 1, 1] == null ? string.Empty : paramArray[rowIndex - mainSheetConfig.StartingLine + 1, 1].ToString();
						result.Add(value);
					}
					else
					{
						object[,] mainSheetUniquelyParam = ReadRangeValues(MainConfigModel.MainSheetName + "." + mainSheetConfig.UniquelyIdentifiesLine);
						string mainSheetUniquelyValue = mainSheetUniquelyParam[rowIndex - mainSheetConfig.StartingLine + 1, 1] == null ? string.Empty : mainSheetUniquelyParam[rowIndex - mainSheetConfig.StartingLine + 1, 1].ToString();
						object[,] targetSheetUniquelyParam = ReadRangeValues(vs[0] + "." + SheetConfigModels[vs[0]].UniquelyIdentifiesLine);
						bool hasUniquelyValue = false;
						int targetSheetRowIndex = 0;
						for (int i = 1; i < targetSheetUniquelyParam.Length + 1; i++)
						{
							string uniquelyValue = targetSheetUniquelyParam[i, 1] == null ? string.Empty : targetSheetUniquelyParam[i, 1].ToString();
							if (uniquelyValue == mainSheetUniquelyValue)
							{
								hasUniquelyValue = true;
								targetSheetRowIndex = i;
								break;
							}
						}
						if (hasUniquelyValue)
						{
							object[,] targetSheetParam = ReadRangeValues(vs[0] + "." + vs[1]);
							string value = GetValue(targetSheetRowIndex, targetSheetParam);
							result.Add(value);
						}
						else
						{
							throw new ApplicationException("无法找到" + mainSheetUniquelyValue + "的唯一标识!");
						}
					}
				}
			}
			return result;
		}

		private string GetValue(int index, object[,] valueArray)
		{
			return valueArray[index, 1] == null ? string.Empty : valueArray[index, 1].ToString();
		}

		private object[,] ReadRangeValues(string paramKey)
		{
			string[] vs = paramKey.Split('.');
			object[,] paramArray;
			if (ParamsDict.ContainsKey(paramKey))
			{
				paramArray = ParamsDict[paramKey];
			}
			else
			{
				SheetConfigModel sheetConfigModel = SheetConfigModels[vs[0]];
				if (vs.Length == 2)
				{
					var array = SheetName2ExcelSheet[vs[0]].Range[vs[1] + sheetConfigModel.StartingLine + ":" + vs[1] + sheetConfigModel.EndLine].Value[10];
					if (array as object[,] != null)
					{
						paramArray = array;
					}
					else
					{
						int[] arrParam = { 1, 1 };
						paramArray = Array.CreateInstance(typeof(object),
										   arrParam,
										   arrParam) as object[,];
						paramArray[1, 1] = array;
					}

				}
				else
				{
					int[] arrParam = { 1, 1 };
					paramArray = Array.CreateInstance(typeof(object),
									   arrParam,
									   arrParam) as object[,];
					paramArray[1, 1] = SheetName2ExcelSheet[vs[0]].Range[vs[1] + vs[2]].Value[10];
				}
				ParamsDict[paramKey] = paramArray;
			}
			return paramArray;
		}

		private string FormatString(IList<string> paramList, string targetString)
		{
			StringBuilder stringBuilder = new StringBuilder(targetString);
			if (paramList.Any())
			{
				for (int i = 0; i < paramList.Count; i++)
				{
					stringBuilder.Replace("{" + i.ToString() + "}", paramList[i]);
				}
				return stringBuilder.ToString();
			}
			else
			{
				return targetString;
			}
		}

		/// <summary>
		/// 将信息输出到textbox中
		/// </summary>
		/// <param name="message"></param>
		private void PushMessage(string message)
		{
			OutPutText.Append(message);
			ViewModel.Text = OutPutText.ToString();
			DoEvents();
		}

		private void ButtonSelectConfigFile_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog fileDialog = new OpenFileDialog
			{
				InitialDirectory = Environment.CurrentDirectory,
				Filter = "ini 配置文件|*.ini",
				Multiselect = false,
			};
			if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				ViewModel.ConfigFilePath = fileDialog.FileName;
			};
			PushMessage("加载配置文件中。\n");
			if (Workbook != null)
			{
				Workbook.Save();
				Workbook.Close();
				SheetName2ExcelSheet.Clear();
				SheetConfigModels.Clear();
				BodyParams.Clear();
				SubjectParams.Clear();
				MailToParams.Clear();
				OutPutText.Clear();
				ParamsDict.Clear();
				buttonStart.Click -= ButtonStart_Click;
				buttonSendTestMail.Click -= ButtonSendTestMail_Click;
			}
			bool loadSuccess = LoadConfig();
			if (loadSuccess)
			{
				buttonStart.Click += ButtonStart_Click;
				buttonSendTestMail.Click += ButtonSendTestMail_Click;
			}
		}

		private void ButtonSendTestMail_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrEmpty(ViewModel.TestMailAddress))
			{
				System.Windows.MessageBox.Show("先输入测试地址！");
			}
			else
			{
				try
				{
					SheetConfigModel mainSheetConfig = SheetConfigModels[MainConfigModel.MainSheetName];
					MailDataModel mailDataModel = new MailDataModel
					{
						From = MailConfigModel.MailAddress,
						Password = MailConfigModel.MailPassword,
						SMTPAddress = MailConfigModel.SMTPAddress,
						IsbodyHtml = true,
						To = ViewModel.TestMailAddress,
						Subject = FormatString(CalcParamValues(SubjectParams, mainSheetConfig.StartingLine), MailConfigModel.MailSubject),
						Body = FormatString(CalcParamValues(BodyParams, mainSheetConfig.StartingLine), File.ReadAllText(MainConfigModel.TemplatePath)),
						Port = MailConfigModel.Port,
						EnableSsl = MailConfigModel.EnableSsl,
					};
					MailSenderUtil.Send(mailDataModel);
					System.Windows.MessageBox.Show("发送成功！");
				}
				catch (Exception ex)
				{
					System.Windows.MessageBox.Show(ex.Message);
				}
			}
		}

		private void Window_Closing(object sender, CancelEventArgs e)
		{
			Workbook?.Save();
			Workbook?.Close();
		}

		public void DoEvents()
		{
			DispatcherFrame frame = new DispatcherFrame();
			Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
				new DispatcherOperationCallback(delegate (object f)
				{
					((DispatcherFrame)f).Continue = false;

					return null;
				}
					), frame);
			Dispatcher.PushFrame(frame);
		}

		public static IList<Range> GetMergeRanges(Worksheet sheet, IList<string> addresList)
		{
			List<Range> rangeList = new List<Range>();
			if (addresList == null || !addresList.Any())
			{
				return rangeList;
			}

			StringBuilder sb = new StringBuilder();
			var app = sheet.Application;
			OfficeExcel.Range rg = sheet.Range[addresList[0]];

			void GetRange()
			{
				var str = sb.ToString().TrimEnd(',');
				var tmpRange = sheet.Range[str];
				try
				{
					rg = app.Union(rg, tmpRange);
				}
				catch
				{
					rangeList.Add(rg);
					rg = app.Union(tmpRange, tmpRange);
				}
				sb.Clear();
			}

			// 地址的最大长度不能超过256，否则Office会识别出错
			var maxLength = 240;

			while (addresList.Any())
			{
				if (sb.Length < maxLength)
				{
					sb.AppendFormat("{0},", addresList[0]);
					addresList.RemoveAt(0);
				}
				else
				{
					GetRange();
				}
			}

			if (sb.Length > 1)
			{
				GetRange();
			}

			if (!rangeList.Any())
			{
				rangeList.Add(rg);
			}

			return rangeList;
		}
	}
}
