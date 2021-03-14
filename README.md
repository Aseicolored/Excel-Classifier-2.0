# Excel-Classifier-2.0

The Gen 2 Excel Classidier ! Quicker ! Easier !

环境要求:
1. 操作系统为Windows 7及以上;
2. 已安装python 3.7及以上;
3. 已安装xlwings模块.

用途:
通过列Excel工作簿的字符串(或其他类型的数据)将另一个更为完全的Excel中所对应的数据提取出来.

使用方法:
1. 在 "Source/Path: " 输入数据来源的工作簿;
2. 在 "Output/Path: " 输入将要写入的工作簿;
3. 在 "Worksheet/Source: " 输入数据来源工作簿要被读取的工作表;
4. 在 "Worksheet/Source: " 输入将要写入的工作簿要被操作的工作表;
5. 在 "Source/Ground: " 输入用于作为提取数据判断依据的列(数据来源工作表的列与写入工作表的列必须相同);
6. 在 "Column/Left: " 输入将要被读取和写入的最左列;
7. 在 "Column/Right: " 输入将要被读取和写入的最右列;
8. 在 "Source/Row/Begin: " 输入在被读取工作表中将要被读取的首行;
9. 在 "Source/Row/End: " 输入在被读取工作表中将要被读取的尾行;
10. 在 "Output/Row/Begin: " 输入在被写入工作表中将要被读取的首行;
11. 在 "Output/Row/End: " 输入在被写入工作表中将要被读取的尾行;
12. 开始运行!

注意:
1. Excel工作簿路径可以为相对路径或者绝对路径;
2. 没有过多的保护机制,使用时请多注意.
