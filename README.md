# MailSender4Excel
根据excel文件和html模板批量发送邮件。支持附件。
# SheetNames说明
此项表示Excel数据的抽取包含了哪些sheet。多个sheet名称用逗号分隔。配置后要对应添加[sheetName]的配置。此配置包含对应sheet的起始行，结束行和唯一标识。
# UniquelyIdentifiesLine唯一标识说明
多个sheet时，需要对应的标识来找到另一个sheet中与当前项对应的数据。
# 参数表示说明
参数表示均为{index}来表示。第一个参数{0}，第二个参数{1},以此类推
# 参数值说明
参数值可以表示为 sheet名称.列名 或 sheet名称.列名.行索引。其中第一种参数取固定单元格。第二种参数按照当前执行的行来获取参数。
参数分3种。配完对应的数量后。要在对应参数下面配置对应的参数值。
