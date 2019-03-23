# 《C#控制台编程》实验报告
#### 学院：软件学院  班级：2017级软工1班  学号：3017218062   姓名：刘书裴
#### 日期：  2019  年 3 月 14 日
## 一、功能概述
###### 1、命令行输入一个”字符串”，则输出对应二维码。
###### 2、命令行输入”-fqrcode.txt”，则将qrcode.txt中每一行字符串转化为二维码图片。
###### 3、命令行输入”-mqrcode”，则将mysql的mydata库qrcode表中每一行字符串转化为二维码图片。
###### 4、命令行输入”-eqrcode.xlsx”，则将qrcode.xlsx中每一行字符串转化为二维码图片。
###### 5、命令行输入过多、过长或无输入，则输出对应错误。

## 二、项目特色
###### 1、生成的二维码加入了背景和logo，显得更加华丽。
###### 2、考虑所有参数情况和文件情况：
######## 1)加入参数个数限制和长度限制，违反限制会输出相应错误。
######## 2)加入文件类型判断和存在判断，不存在即新建文件。
###### 3、加入try-catch异常语句，防止意外错误的发生。
###### 4、将大部分功能集成为多个成员函数，避免修改麻烦。

## 三、代码总量
![](https://github.com/cxdzb/Lab1/blob/master/Images/1.png?raw=true)

## 四、工作时间
![](https://github.com/cxdzb/Lab1/blob/master/Images/time.png?raw=true)

总共8hours 30minutes左右。

## 五、知识点总结图（Concept MAP）
![](https://github.com/cxdzb/Lab1/blob/master/Images/2.png?raw=true)

## 六、结论
#### 实验过程：
![](https://github.com/cxdzb/Lab1/blob/master/Images/3.png?raw=true)

###### 1、StringToQrCode和PrintQrCode
QrCode通过new QrEncoder(ErrorCorrectionLevel.M).Encode(str)创建二维码对象，内部包含一个方阵（Matrix），可通过遍历打印出来。
###### 2、DumpPng
QrCode需要用GraphicsRenderer进行渲染，再通过render.draw(graph,qrcode)存取数据，读取并打印图片需要先创建Bitmap位图再使用位图创建Graphics，graph.FillRectangle()可以填充一个有色矩形，graph.DrawImage可以将一张图绘制在另一张图上，取焦点时可以使用new Point，最后可以使用Sava保存图片。
###### 3、ReadTxt
使用FileStream文件流读取文件，设置参数为FileMode.OpenOrCreate和FileAccess.Read表示存在文件时读取，不存在文件时新建后读取。使用StreamReader可通过ReadLine()逐行读取，读取后使用Add()将每一行添加到List中。
###### 4、ReadMysql
创建MySqlConnection连接，使用Open()开启连接，将mysql命令传入MySqlCommand，然后用ExecuteReader()执行并保存结果，不断使用Read()读取，最后Close()。
###### 5、ReadExcel
类似于txt，也是用FileStream读取。本次是用NPOI处理Excel，先使用XSSFWorkbook和HSSFWorkbook处理xls与xlsx两种格式，再取GetSheetAt()，使用GetRow().GetCell()取值。
###### 6、Main
使用多次判断来实现多种参数的输入，使用try-catch语句避免意外错误的发生而中止程序。新建一个对象用以实现所有功能。
#### 实验结果：
###### 1、输入“myqrcode”
![](https://github.com/cxdzb/Lab1/blob/master/Images/4.png?raw=true)

###### 2、输入“myqrcode test”
![](https://github.com/cxdzb/Lab1/blob/master/Images/5.png?raw=true)

###### 3、输入“myqrcode -f”
![](https://github.com/cxdzb/Lab1/blob/master/Images/6.png?raw=true)

###### 4、输入“myqrcode -m”
![](https://github.com/cxdzb/Lab1/blob/master/Images/7.png?raw=true)

###### 5、输入“myqrcode -e”
![](https://github.com/cxdzb/Lab1/blob/master/Images/8.png?raw=true)

###### 6、输入“myqrcode -fresource\qrcode.txt”
![](https://github.com/cxdzb/Lab1/blob/master/Images/9.png?raw=true)

###### 7、输入“myqrcode -mqrcode”
![](https://github.com/cxdzb/Lab1/blob/master/Images/10.png?raw=true)

###### 8、输入“myqrcode -eresource\qrcode.xlsx”
![](https://github.com/cxdzb/Lab1/blob/master/Images/11.png?raw=true)

###### 9、输入“myqrcode 123qwe123qw123qwe123qwe123we123qwe123qwe123qwe123qwe123q”
![](https://github.com/cxdzb/Lab1/blob/master/Images/12.png?raw=true)

###### 10、输入“myqrcode arg1 arg2 arg3”
![](https://github.com/cxdzb/Lab1/blob/master/Images/13.png?raw=true)
