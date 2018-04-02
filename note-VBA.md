# vba学习笔记

- MsgBox 弹窗	
- b = Application.WorksheetFunction.CountA(Range("a:e")) 统计A到E列非空单元格
- Ubound(arr)最大索引号 Lbound(arr) 最小索引号
	- 返回一个 Long 型数据，其值为指定的数组维可用的最大下标。**
	- dimension 可选的；Variant (Long)。指定返回哪一维的上界。1 表示第一维，2 表示第二维，如此等等。如果省略 dimension，就认为是 1。
	- ,其中arr为数组 二维以上数组，则需确定维数，第几维，如Ubound(arr，2)
- join（arr，“@”） 将数组里面的没一个元素用@连接起来变成一个字符串

- a = Range("B2:B4").Value 
	- 直接从单元格导入数组，必然是二维数组。哪怕是你只导入一行 。系统也默认是一个一行多列的二维数组，这是系统设定的。牢记它。
	- 从单元格导入的数组的维度下标都是从1 开始的。 arr = Range ("A1: F1")，arr(1,1) 是A1 的值, arr(1,2) 是B1 以此类推，在这种情况下不会出现arr(0,0)

- application.workbooks("book1").worksheets("sheets1").range("A2")

- ctrl+j

- 判断语句

	- if 判断 then XXX elseif 判断 then XXX else XXX end if 
	- select case 表达式 case 判断 XXX case 判断 XXX .... XXX end select
- 循环语句
	- for i 1 to 100 step 1 XXX next i(可选)
	- do while 逻辑表达式 XXX [exit do](可选) XXX loop 逻辑表达式
	- do XXX [exit do](可选) XXX loop while 逻辑表达式
	- do until 逻辑表达式 XXX [exit do](可选) XXX loop 逻辑表达式为false执行循环体
	- for each 元素变量 in 集合或者元素名称 XXX next【元素变量】
	- With语句使您可以对指定的对象执行一系列语句，而不必重新限定对象的名称。例如，要更改大量的单个对象的不同属性，将属性分配语句内使用控制结构中，指的一次，而每个属性赋值引用它的对象。下面的示例阐释了使用With语句将值分配给同一对象的多个属性。
``` 
	Sub with_try ()
	With MyObject
		.Height = 100 ‘备注：Same as MyObject.Height = 100. 
		.Caption = "Hello World" ’备注：Same as MyObject.Caption = "Hello World". ）
		With .Font ‘备注：套嵌使用
			.Color = Red ' 备注：Same as MyObject.Font.Color = Red.
			.Bold = True ' 备注：Same as MyObject.Font.Bold = True.
		End With
	End With 	
	End Sub
```
- 自定义函数
	-	例子：自定义函数计算自选区域内黄色单元格的个数
- 定义易失行函数
	- 在自定义函数开头加入 Application.Volatile (True)
	- 用于将用户自定义函数标记为易失性函数，无论何时在工作表的任意单元格中进行计算时，易失性函数都必须重新进行计算。非易失性函数只在输入变量改变时才重新计算，若不用于计算工作表单元格的用户自定义函数中，则此方法无效。
- Application.DisplayAlerts 属性
	- 默认值为 True 。将此属性设置为 False 可在宏运行时禁止显示提示和警告消息；当出现需要用户应答的消息时，Microsoft Excel 将选择默认应答。即决定是否显示警告信息。
	- 例如该过程删除与活动工作表不同名的工作表，如果不加入Application.DisplayAlerts =false 则会提示告警
``` 
	Public Sub delsht()
  	Dim sheet As Worksheet
 	Application.DisplayAlerts = False 加入则不提示直接删除
 	For Each sheet In Worksheets
 		If sheet.Name <> ActiveSheet.Name Then
 			sheet.Delete
 		End If
	Next
	Application.DisplayAlerts = true 需要同时在后面重新设置回来
	End Sub
```
- Application.WorksheetFunction 可用来调用部分工作表函数
	- arr = Application.WorksheetFunction.CountIf(Range("A1:B10"), ">100")
	- 如果VBA有相同的函数，则不能引用工作表的函数，要用VBA的函数，如使用len函数计算，要写成len("abcd"),不能写成Application.WorksheetFunction.len("abcd")
- 打开文件workbooks.open ("路径加文件名excel"，可选14个参数)
- 保存文件thisworkbook.save ActiveWorkbook.Save
- 另存为，并打开新文件 ThisWorkbook.SaveAs "备份.xlsm"
- 另存为，保留原来文件不打开新文件 ThisWorkbook.SaveCopyAs "备份.xlsm"
- 关闭工作簿 表达式 . Close( SaveChanges, Filename, RouteWorkbook )
- Workbooks.Close关闭所有
- ThisWorkbook.Close关闭程序所在工作簿
- ActiveWorkbook.Close 关闭活跃工作簿
- Workbooks("BOOK1.XLS").Close SaveChanges:=False 关闭工作簿并放弃所有对此工作簿的更改。
- 第几张工作表
	- Worksheets(3).Range("A3") = 111 工作簿按顺序从左到右为第三张，索引号为3
	- Sheet9.Range("A3") = 222 代码名称为Sheet9,可直接引用
	- Worksheets.Item(3).Range("A3") = 444 同第一条，item为索引
	- Worksheets("表5").Range("A3") = 333 Sheet9的标签名为表5
	- 以上4条等价，其中由于Sheet9标签名称为表5，但是位于工作簿第三位，所以都是等价
	- MsgBox ActiveSheet.CodeName 可显示当前活动窗口代码名称
- 添加工作表 表达式 . Add( Before, After, Count, Type ) 
	- 表达式 :一个代表 Sheets 对象的变量。
	- Worksheets.Add before:=Worksheets(Worksheets.Count) 新建工作表插入到活动工作簿的最后一张工作表之前。
	- Worksheets.Add before:=Worksheets(Worksheets.Count), Count:=3 多个参数，新建3张工作表插入到活动工作簿的最后一张工作表之前。
	- Worksheets.Add(before:=Worksheets(Worksheets.Count)).Name = "新建" 可以在添加工作表的同时修改名字，注意需要加括号
	- 修改工作表名称 Worksheets.Item(7).Name = "工作表"
- 激活工作表
	- Worksheets(1).Activate，Worksheets(3).Select ，当隐藏工作表时，2条都无法选择
	- worksheets.select 可以选中所有工作表，但使用activate不行
- 参数后面加 :=
- 复制工作表 表达式 . Copy( Before, After ) 
	- 表达式 :一个代表 Sheets 对象的变量。
	- Worksheets("Sheet1").Copy after:=Worksheets("Sheet3")
	- 如果既不指定 Before 也不指定 After，则 Microsoft Excel 将新建一个工作簿，其中包含复制的工作表。
- 移动工作表 表达式 . Move( Before, After ) 与复制类似
	- Worksheets("Sheet1").Move after:=Worksheets("Sheet3")
- 隐藏or显示工作表 
	- Worksheet.Visible 属性 返回或设置一个XlSheetVisibility值，决定对象是否可见。
	- xlSheetHidden 对应 0 xlSheetVisible 对应 1/-1 xlSheetVeryHidden 对应 2
	- Worksheets(1).Visible = False Worksheets(1).Visible = 0 Worksheets(1).Visible = xlSheetHidden 隐藏工作表，等同于格式菜单里隐藏表达式
	- Worksheets(2).Visible = xlSheetVeryHidden Worksheets(2).Visible = 2 隐藏工作表，但在excel表无法显示，只能在VBA属性窗口修改或者代码修改为显示。
	- For Each sh In Worksheets sh.Visible = True Next 显示所有工作表
- 工作表数目 Worksheets.Count  返回一个long值
- sheets ,worksheets 区别 worksheets是sheets的一个子集 excel有4种不同的工作表，sheets表示所有工作簿所有类型的工作表集合，worksheets只是表示普通工作表集合
- worksheet或range对象的range属性
	- 引用单元格
```
	Range("A1:A10") = 200
	Dim n As String
	n = "A1:A10"
 	Range(n) = 300
```
	- 引用不连续区域 Range("A1:A10,B3:E10").Select 1个引号里面2个区域，逗号隔开 (一个参数)
	- 引用2个区域围成的最小区域 Range("A1", "B3:E10").Select 2个引号各1个区域，逗号隔开 二个参数
	- 引用2个区域围成的重叠区域 Range("A1:B7 B3:E10").Select 1个引号里面2个区域，空格隔开 二个参数
- worksheet或range对象的cell属性
	- Worksheet.Cells返回一个 Range 对象，该对象代表工作表上的所有单元格（不仅仅是当前正在使用的单元格）。在不使用对象识别符的情况下，使用此属性将返回一个 Range 对象，它代表活动工作表中所有的单元格。
	- cells.clear 可以清除工作表所有单元格
	- ActiveSheet.Cells(3, 4).Value = 200 等价 ActiveSheet.Cells(3, "D").Value = 300 cell是（行数，列数）
	- Range("A1:D10").Cells(1, 4) = 400 range的cell属性
	- Range(Cells(1), Cells(7)) = 200 range里面用cell表示
	- ActiveSheet.Cells(3).Value = 700 cell可以只用1个参数，索引号，表示第几个，从开头从左往右算，
	- Range("A1:D10").Cells(100) = 400 range对象cell属性，当索引号超过范围，自动在行的方向进行扩展，列不变。
	- [A100] = 100 可直接用中括号引用，但里面不能使用变量
- 引用整行，列
	- ActiveSheet.Rows("3:5").Select
	- ActiveSheet.Columns("B:D").Select
	- ActiveSheet.Columns(3).Select 第三列
- union 返回两个或多个区域的合并区域。
	-  表达式. Union( Arg1, Arg2, [Arg3]，...., [Arg30] ) 两个必选，最多30个 表达式 一个代表 Application 对象的变量。
	- Application.Union(Range("A1:B2"), Range("C4:D5")).Select application可以省略
	- 小练习，选择与A1相同值的全部单元格
```
	Dim myrange, n As Range
	Set myrange = Range("A1")
	For Each n In Range("A1:D10")
		If n.Value = Range("A1").Value Then
			Set myrange = Union(myrange, n)
		End If
	Next
	myrange.Select
```
- range对象的offset属性 即偏移
	- Range("A1").Offset(3, 4).Value = 400 A1向下移动3行，再向右移动四列
	- range对象的resize属性 扩大缩小指定单元格区域
	- Range("A1").Resize(3, 4).Select 以range区域最左上角单元格为基准，设定3行4列区域 扩大
	- Range("A1：E10").Resize(2, 1).Select ，以range区域最左上角单元格为基准，设定2行1列区域 缩小
- worksheet对象的usedrange属性 返回一个 Range 对象，该对象表示指定工作表上所使用的区域。
	- ActiveSheet.UsedRange.Select 返回工作表已经使用的单元格所围成的区域
- range对象的CurrentRegion属性 返回一个 Range 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域
	- Range("D8").CurrentRegion.Select
	- range对象的end属性 返回一个 Range 对象，该对象代表包含源区域的区域尾端的单元格。
	- Range("D65536").End(xlUp).Offset(1, 0).Select
	- xlToLeft & xlDown & xlToRight & xlUp 四个方向
	- 小练习，取A列非空单元格并赋值
```
	Dim a As Range
	Set a = Range("A65535").End(xlUp)
	If a.Value <> "" Then
		Set a = a.Offset(1, 0)
	End If
	a.Value = 200
```
- range对象的value属性 默认属性，赋值时可忽略
	- Range("A1") = 250 等价 Range("A1").Value = 250
- range对象的count属性
	- MsgBox Range("A1,D10").Count 共40个单元格
	- ActiveSheet.UsedRange.Rows.Count 活动工作表已使用的行数
	- ActiveSheet.UsedRange.Column.Count 活动工作表已使用的列数
- 单元格地址。address属性
	- Range("a1") = Selection.Address
- 表达式.Selection 为 Application 对象返回在活动窗口中选定的对象。 表达式   一个代表 Application 对象的变量。在不使用对象识别符的情况下，使用此属性等效于使用 Application.Selection。
	- Selection.clear 等价 Application.Selection
- 选中单元格 select与activate 两者都可以选中，但区别在于activate 选中单元格区域后，再使用activate激活该区域里的单元格时，该区域仍呈现选中状态，只改变活动单元格为激活的单元格，而使用select则只选中最后一次使用时的单元格
	- Range("A1:B10").Select
	- Range("B5").Select
	- Range("A1:B10").Activate
	- Range("B5").Activate
- 清除单元格区域
	- Range.Clear 方法 清除整个对象。cells.clear
	- Range.ClearComments 方法 清除指定区域的所有单元格批注。
	- Range.ClearContents 方法 清除区域中的公式。 会清除内容，但不会清除格式
	- Range.ClearFormats 方法 清除对象的格式设置。
	- Range.ClearHyperlinks 方法 删除指定区域中的所有超链接。
	- Range.ClearNotes 方法 清除指定区域中所有单元格的批注和语音批注。
	- Range.ClearOutline 方法 清除指定区域的分级显示。
- 复制单元格区域 Range.Copy 方法 将单元格区域复制到指定的区域或剪贴板中。
	- 表达式.Copy(Destination) Destination 指定区域要复制到的新域。如果省略此参数，Microsoft Excel 会将区域复制到剪贴板。
	- Worksheets("Sheet1").Range("A1:A10").Copy Worksheets("Sheet2").Range("B1") 表1A1:A10f复制到表2【B1：B10】
	- Worksheets("Sheet1").Range("C4").CurrentRegion.Copy Worksheets("Sheet2").Range("B1")
- 剪切单元格区域，range.cut 与复制类似
- 删除单元格 Range.Delete 方法 表达式.删除(移位)
	- Shift 可选Variant只能与Range对象一起使用。指定如何移动单元格来替换删除的单元格。可以是下列的XlDeleteShiftDirection常量之一： xlShiftToLeft或xlShiftUp 。如果省略此参数，则 Microsoft Excel 将决定基于区域的形状上。
	- 不使用xlShiftToLeft或xlShiftUp ，则默认先上移，后左移
	- Range("B3").Delete xlShiftUp
	- Range("B3").Delete
	- Range("B3").EntireRow.Delete 删除整行
	- Range.EntireRow 返回range对象 Range.Row 返回行号，long类型
- 粘贴 Range.PasteSpecial 方法 将 Range 从剪贴板粘贴到指定的区域中。
	- 例如：
- Worksheets("Sheet1").Range("C4").CurrentRegion.Copy
	- Range("F1").PasteSpecial xlPasteValues 等价Range("F1").PasteSpecial Paste:=xlPasteValues等价Range("F1").PasteSpecial -4163
	- 表达式.PasteSpecial(Paste, Operation, SkipBlanks, Transpose) 表达式   一个代表 Range 对象的变量。
	- Paste 可选 XlPasteType 要粘贴的区域部分。
		- xlPasteAll -4104 粘贴全部内容。
		- xlPasteAllExceptBorders 7 粘贴除边框外的全部内容。
		- xlPasteAllMergingConditionalFormats 14 将粘贴所有内容，并且将合并条件格式。
		- xlPasteAllUsingSourceTheme 13 使用源主题粘贴全部内容。
		- xlPasteColumnWidths 8 粘贴复制的列宽。
		- xlPasteComments -4144 粘贴批注。
		- xlPasteFormats -4122 粘贴复制的源格式。
		- xlPasteFormulas -4123 粘贴公式。
		- xlPasteFormulasAndNumberFormats 11 粘贴公式和数字格式。
		- xlPasteValidation 6 粘贴有效性。
		- xlPasteValues -4163 粘贴值。
		- xlPasteValuesAndNumberFormats 12 粘贴值和数字格式。
	- Operation 可选 XlPasteSpecialOperation 粘贴操作。
	- SkipBlanks 可选 Variant 如果为 True，则不将剪贴板上区域中的空白单元格粘贴到目标区域中。默认值为 False。
	- Transpose 可选 Variant 如果为 True，则在粘贴区域时转置行和列。默认值为 False。
- 颜色 Interior.Color 属性 返回或设置对象的主要颜色，如注释部分中的表格所示。使用 RGB 函数可创建颜色值。 Variant 型，可读写。
	- 表达式 . Color 表达式 一个返回 Interior 对象的表达式。
	- Border 边框的颜色。
	- Borders 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 Color 返回的是 0（零）。
	- Font 字体的颜色。
	- Interior 单元格底纹的颜色或图形对象的填充颜色。
	- Tab 选项卡的颜色。
	- Range("B5").Borders.Color = RGB(255, 0, 0) 红框 
	- Range("B5").Interior.Color = RGB(255, 0, 0) 底纹红色
- 批注添加删除更改 Range.AddComment 方法 为区域添加批注。返回值：Comment
	- 表达式.AddComment（Text） Text可选Variant批注文字。
	- Range("B5").AddComment "我是大佬" 如果原本有批注，则会报错
	- Range("B5").Comment.Text "我不是大佬啊" 更改批注，原来需要有批注
	- Range("B5").Comment.Visible = True 显示批注 Flase隐藏
	- Range("B5").Comment.Delete 删除
- 字体 font 对象包含对象的字体属性（字体名称、字号、颜色等等）。
```
	With Range("B5").Font
		.Name = "宋体" 字体
		.Bold = True 加粗
		.Size = 12 大小
		.Color = RGB(255, 0, 0) 颜色
		.Italic = True 倾斜
		.Underline = True 下划线
	End With
```
- 边框 borders
- 内部 Interior 对象 代表一个对象的内部。
	- Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 3 将单元格 A1 的内部设置为红色。
- 创建，并保存到指定文件夹。
```
	 Dim Wb As Workbook, Ws As Worksheet
	 Set Wb = Workbooks.Add
	 Set Ws = Wb.Worksheets(1)
	 Ws.Name = "花名册"
	 With Ws.Range("A1:E1")
	 	.Value = Array("小区", "基站", "cellID", "频点", "地址")
	 	With .Font
	 		.Color = RGB(255, 0, 0)
	 		.Name = "宋体"
	 		.Size = 18
	 		.Bold = True
	 	End With
	 	.Interior.Color = RGB(0, 255, 0)
	 	.Borders.Color = RGB(0, 0, 0)
	 End With
	 Wb.SaveAs ThisWorkbook.Path & "\花名册.xlsx"
	 ActiveWorkbook.Close
```
- 判断文件是否打开

```
	 Dim Wb As Integer
	 For Wb = 1 To Workbooks.Count
	 	If Workbooks(Wb).Name = "花名册.xlsx" Then
	 		MsgBox "该文件已打开"
			Exit Sub ‘如果打开，则使用Exit sub 退出执行程序
	 	End If
	 Next
	 MsgBox "该文件未打开"
```
- 判断工作簿里面的工作表是否存在，并放置最前面
```
	 On Error Resume Next
	 If Worksheets("花名册") Is Nothing Then
	 	Worksheets.Add(before:=Worksheets(1)).Name = "花名册"
	 Else
	 Worksheets.Move before:=Worksheets(1)
	 End If
```
- 判断工作簿在某文件夹是否存在
```
	 Dim h As String
	 h = ThisWorkbook.Path & "\花名册.xlsx"
	 If Len(Dir(h)) <> 0 Then        'dir函数，如果存在，返回文件名，反之返回空字符串。
	 	MsgBox "存在"
	 Else
	 	MsgBox "不存在"
	 End If
```
- 1 把A列的每个不重复的单元格内容，建立一张表
```
	 Dim i As Integer, Wb As Worksheet
	 Set Wb = Worksheets("Sheet1")
	 i = 2
	 Do While Wb.Cells(i, "A") <> ""
	 On Error Resume Next
	 If Worksheets(Wb.Cells(i, "A").Value) Is Nothing Then
	 	Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = Wb.Cells(i, "A").Value
	 End If
	 i = i + 1
	 Loop
```
- 2 批量对数据分类到各个工作表
```
	 Dim i, j As Integer
	 Dim value As String
	 Dim ws As Worksheet, rng As Range
	 Set ws = Worksheets(1)
	 Worksheets(1).Range("A1:E1").Copy
	 For i = 2 To Worksheets.Count
	 	Worksheets(i).Range("A1:E1").PasteSpecial
	 Next
	 j = 2
	 value = Cells(j, "A").value
	 Do While value <> ""
	 	Set rng = Worksheets(value).Range("A65535").End(xlUp).Offset(1, 0)
	 	Cells(j, 1).Resize(1, 5).Copy rng
	 	j = j + 1
	 	value = Cells(j, "A").value
	 Loop
```
- 多个工作表保存到新的工作簿
```
	Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Copy
	With ActiveWorkbook
		.SaveAs Filename:=Environ("TEMP") &; "\New3.xlsx", FileFormat:=xlOpenXMLWorkbook
		.Close SaveChanges:=False
	End With
```
- Application.ScreenUpdating 属性 屏幕更新
	- 表达式 . ScreenUpdating 表达式 一个代表 Application 对象的变量。
	- 关闭屏幕更新可加快宏的执行速度。这样将看不到宏的执行过程，但宏的执行速度加快了。当宏结束运行后，请记住将 ScreenUpdating 属性设置回 True 。
	- Application.ScreenUpdating = False
- dir & Mkdir 函数
	- dir 返回一个字符串，表示文件、 目录或与指定的模式或文件属性匹配的文件夹名或驱动器卷标。
		- 在第一次调用 Dir 函数时，必须指定 pathname，否则会产生错误。如果也指定了文件属性，那么就必须包括 pathname。Dir 会返回匹配 pathname 的第一个文件名。若想得到其它匹配 pathname 的文件名，再一次调用 Dir，且不要使用参数。如果已没有合乎条件的文件，则 Dir 会返回一个零长度字符串 ("")。一旦返回值为零长度字符串，并要再次调用 Dir 时，就必须指定 pathname，否则会产生错误。不必访问到所有匹配当前 pathname 的文件名，就可以改变到一个新的 pathname 上。但是，不能以递归方式来调用 Dir 函数。以 vbDirectory 属性来调用 Dir 不能连续地返回子目录。
		- 在 Microsoft Windows 中， Dir 支持多字符 (*) 和单字符 (?) 的通配符来指定多重文件。
		- 语法目录[ （路径名[ ，属性] ) ]Dir函数的语法包含以下成分：
		- **vbNormal 0 （默认）指定没有属性的文件。**
		- vbReadOnly 1 指定只读文件以及不带属性的文件。
		- vbHidden 2 指定隐藏文件以及不带属性的文件。
		- VbSystem 4 指定系统文件以及不带属性的文件。在 Macintosh 上不可用。
		- vbVolume 8 指定卷标;如果指定了任何其他特性化，则vbVolume将被忽略。在 Macintosh 上不可用。
		- **vbDirectory 16 指定目录或文件夹以及不带属性的文件。 (文件夹，目录) 参考下面例子**
		- vbAlias 64 指定文件名为别名。仅在 Macintosh 上可用。
	- MkDir 新建目录或文件夹。
		- MkDir路径 所需的_路径_参数是字符串表达式，用于标识的目录或文件夹创建。路径_可以包含驱动器。如果未指定驱动器， MkDir在当前驱动器上创建新的目录或文件夹。
- 当前工作簿中的非活跃工作表，每个表保存为一个工作簿
```
	Dim Ws As Worksheet, n_path As String
	n_path = "C:\Users\zhang\Desktop\excel\默认excel存放位置\班级表"
	If Len(Dir(n_path, vbDirectory)) = 0 Then   '判断文件夹是否存在
		MkDir n_path                        '不存在则创建文件夹
	End If
	Application.ScreenUpdating = False      '禁止屏幕更新
	For Each Ws In Worksheets
		If Ws.Name <> ActiveSheet.Name Then
			Ws.Copy
			With ActiveWorkbook
				.Worksheets(1).Name = Ws.Name
				.SaveAs n_path & "\" & Ws.Name & ".xlsx"
				.Close
			End With
		End If
	Next
	Application.ScreenUpdating = True       '重新启动屏幕更新
```
- 删除非活动工作表
```
	Dim ws As Worksheet
	Application.DisplayAlerts = False  '不提示删除框，默认删除
	For Each ws In Worksheets
		If ws.Name <> ActiveSheet.Name Then
			ws.Delete
		End If
	Next
	Application.DisplayAlerts = True    '恢复默认属性
```
- 清除非激活窗体内容
```
	Dim ws As Worksheet
	For Each ws In Worksheets
		If ws.Name <> ActiveSheet.Name Then
			ws.Cells.clear
		End If
	Next
```
- 合并多个表数据到活动工作表
```
	Rows("2:65565").clear
	Dim temp As Worksheet, rng As Range
	Dim i As Integer
	For Each temp In Worksheets
		If temp.Name <> ActiveSheet.Name Then
			i = temp.Range("A2").CurrentRegion.Rows.Count 1
			Set rng = temp.Range("A2").Resize(i, 5)
			rng.Copy
			Range("A65535").End(xlUp).Offset(1, 0).PasteSpecial
		End If
	Next
```
- GetObject 函数 open方法原理相同，不过open方法打开工作簿是显示的，getobject打开是隐藏的，再次打开不会显示出来，可以在Excel表中 视图>"取消隐藏"  也可以使用以下代码解决
	- 解决用GetObject打开的工作表修改后保存，再次打开工作表不显示
	- 通过getobject打开的Excel文件只要被修改（写）并保存后，就只能在VBE中看到，但用户界面却看不到。就算你重启Excel，再去手动打开此文件，也是什么都看不到。不保存就没有这个问题！如果要解决这个问题，必须在wb.close 前加一句Application.Windows( wb.name).Visible = True。
```
	Private Sub CommandButton1_Click()
	On Error Resume Next
	文件目录 = ThisWorkbook.Path & "\Excel\"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fldr = fso.GetFolder(文件目录)
	For Each s In fldr.Files
		With GetObject(文件目录 & s.Name)
			.Sheets(1).Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False    '随便做一点改动
			.SaveAs ThisWorkbook.Path & "\Excel_修改后\" & s.Name    '保存
			.Windows(1).Visible = True    '工作表可见
			.Close (True)    '保存改动
		End With
	Next
	End Sub
```
- 合并多个工作簿成一张表
```
	Cells.clear
	Dim wb As Workbook, s_path As String
	Dim arr As Variant, i As String
	Dim dr, jr As Integer
	Dim ws As Worksheet
	Dim num As Integer
	num = 1
	i = "C:\Users\Administrator\Desktop\excel\default_path\班级表"
	s_path = Dir(i & "\*.xlsx")
	Application.ScreenUpdating = False
	Do While s_path <> ""
		Set wb = GetObject(i & "\" & s_path)
		Set ws = wb.Worksheets(1)
		dr = ws.Range("A1").CurrentRegion.Rows.Count 1
		jr = ws.Range("A1").CurrentRegion.Columns.Count
		arr = ws.Range(ws.Cells(num, "A"), ws.Cells(dr + 1, jr))
		If num = 1 Then       '先取包括工作簿第一行在内的数值，后面不重复取第一行
			Range("A65535").End(xlUp).Resize(dr, jr) = arr
		Else
			Range("A65535").End(xlUp).Offset(1, 0).Resize(dr, jr) = arr
		End If
		num = 2
		wb.Close False
		s_path = Dir
	Loop
	Application.ScreenUpdating = True
```
- Application.WindowState  返回或设置窗口的状态。读/写XlWindowState 。
	- application.WindowState = xlMaximized '窗口最大化
- Hyperlinks.Add 方法 向指定的区域或形状添加超链接。
	- Hyperlinks.Add 方法
	- 表达式 . Add( Anchor, Address, SubAddress, ScreenTip, TextToDisplay ) 表达式 一个代表 Hyperlinks 对象的变量。
	- Anchor 必需 Object 超链接的位置。可为 Range 或 Shape 对象。
	- Address 必需 String 超链接的地址。
	- SubAddress 可选 Variant 超链接的子地址。
	- ScreenTip 可选 Variant 当鼠标指针停留在超链接上时所显示的屏幕提示。
	- TextToDisplay 可选 Variant要显示的超链接的文本。
- 为工作簿的工作表创建目录
	- hyperlinks 的  地址参数和子地址参数 可以共存吗？ 
	- 可以共存的：address：指定链接的地址。此地址可以是电子邮件地址、Internet 地址或文件名。但Excel不检查该地址的正确性。subaddress：目标文件内的位置名，如已命名的区域或单元格。
	- SubAddress:=Ws.Name & "!A1" 
	- SubAddress:="'" & ws.Name & "'!A1"
	- **加一对撇号是 Hyperlink 函数对工作表名中含有特殊字符时的特定要求，在没有特殊字符的情况下可以不要，也可以要，当含有特殊字符的时候必须要有一对撇号。**
```
	Cells.clear
	Range("A1:B1") = Array("序号", "班级名称")
	Dim ws As Worksheet, i As Integer
	i = 1
	For Each ws In Worksheets
		If ws.Name <> ActiveSheet.Name Then
			i = i + 1
			Cells(i, 1) = i 1
			ActiveSheet.Hyperlinks.Add Cells(i, 2), Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
		End If
	Next
```
- 事件-能被对象识别的操作
- Workbook_Open事件
```
	Sub Workbook_Open() ‘Workbook是对象 Open是事件名称，两者用下划线连接，事件控制程序的规则，当打开对象时自动运行
	msgbox "excel 欢迎你~~"
	End Sub
```
- Worksheet_Change事件 手动或者使用VBA代码修改单元格，都会触发
```
	Private Sub Worksheet_Change(ByVal Target As Range)
		Application.EnableEvents = False '禁止事件，让下面发生的不会再触发事件
			MsgBox "值被改啦"
			Target.Value = "新" & Target.Value '在目标单元格前面插入“新”字，如果没有启用禁止事件，将无限重复，陷入死循环
		Application.EnableEvents = True '重新启动
	End Sub
```
- Worksheet_SelectionChange 选中工作表中的单元格改变为其他单元格时则触发
```
	Private Sub Worksheet_SelectionChange(ByVal Target As Range)
		If Target.Column <> 1 Then '目标单元格不在A列时
			Cells(Target.Row, 1).Select ‘自动选中目标单元格所在同行A列单元格
			Selection.Clear
		End If
	End Sub
```
- Worksheet_Activate 激活工作表时触发（从其他工作表到当前事件工作表）
```
	Private Sub Worksheet_Activate()
		MsgBox "当前工作表为：" &ActiveSheet.Name
	End Sub
```
- Worksheet_Deactivate 但当前事件工作表为活动工作表，并选中其他工作表时触发
```
	Private Sub Worksheet_Deactivate()
		MsgBox "不要乱跑哟"
		Worksheets("Sheet8").Select '重新选中工作表，实现禁止选中其他工作表
	End Sub
```
- Workbook_BeforeClose 关闭工作簿触发
```
	Private Sub Workbook_BeforeClose(Cancel As Boolean)
	If MsgBox("确定要关闭", vbYesNo) = vbNo Then
	Cancel = True '选择否，则不关闭
	End If
	End Sub
```
- Workbook_SheetChange 改变工作簿任意单元格触发，Sh为工作表，Target 为单元格
```
	Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
		MsgBox "当前改变的单元格所在的工作表：" & Sh.Name & _
		"当前改变的单元格地址：" & Target.Address
	End Sub
```
- CommandButton1_MouseMove
	- Int(number) Fix(number) 必要的 number 参数是 Double 或任何有效的数值表达式。如果 number 包含 Null，则返回 Null。
	- Int 和 Fix 都会删除 number 的小数部份而返回剩下的整数。Int 和 Fix 的不同之处在于，如果 number 为负数，则 Int 返回小于或等于 number 的第一个负整数，而 Fix 则会返回大于或等于 number 的第一个负整数。例如，Int 将 -8.4 转换成 -9，而 Fix 将 -8.4 转换成 -8。
```
	Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
		Dim i, j As Integer
		i = Int(Rnd() * 10 + 125) * (Int(Rnd() * 3 + 1) - 1)
		j = Int(Rnd() * 10 + 30) * (Int(Rnd() * 3 + 1) - 2)
		CommandButton1.Top = CommandButton1.Top + i
		CommandButton1.Left = CommandButton1.Left + j
	End Sub
```
- CommandBarButton.Left 属性
	- 设置或获取指定的 CommandBarButton 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。
- CommandBarButton.Top 属性
	- 获取指定的 CommandBarButton 控件顶边到屏幕顶边的距离（以像素为单位）。只读。
- 移动鼠标到控件，按钮会跟着随机跳动
```
	Private Sub Cmd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
		Dim i, j As Integer
		i = Int(Rnd() * 10 + 125) * (Int(Rnd() * 3 + 1) - 2)
		j = Int(Rnd() * 10 + 30) * (Int(Rnd() * 3 + 1) - 2)
		Cmd.Top = Cmd.Top + i
		Cmd.Left = Cmd.Left + j
	End Sub
```
- Application.OnKey 方法
	- 当按特定键或特定的组合键时运行指定的过程。表达式.OnKey(Key, Procedure)
	- Key 必选 String 表示要按的键的字符串。
	- Procedure 可选 Variant 表示要运行的过程名称的字符串。如果 Procedure 为空文本 ("")，则按 Key 时不发生任何操作。该格式的 OnKey 将更改键击在 Microsoft Excel 中产生的正常结果。如果省略 Procedure 参数，则 Key 恢复为 Microsoft Excel 中的正常结果，同时清除先前使用 OnKey 方法所做的特殊键击设置。
	- Key 参数可指定任何与 Alt、Ctrl 或 Shift 组合使用的键，还可以指定这些键的任何组合。每一个键可由一个或多个字符表示，比如 "a" 表示字符 a，"{ENTER}" 表示 Enter。若要指定按对应的键（例如 Enter 或 Tab）时的非显示字符，请使用下表所列出的代码。表中的每个代码表示键盘上的一个对应键。
	- 按键 					代码
	- Backspace {BACKSPACE} 或 {BS}
	- Break {BREAK}
	- Caps Lock {CAPSLOCK}
	- Clear {CLEAR}
	- Delete 或 Del {DELETE} 或 {DEL}
	- 向下键 {DOWN}
	- End {END}
	- Enter（数字小键盘） {ENTER}
	- Enter ~（波形符）
	- Esc {ESCAPE} 或 {ESC}
	- Help {HELP}
	- Home {HOME}
	- Ins {INSERT}
	- 向左键 {LEFT}
	- Num Lock {NUMLOCK}
	- PageDown {PGDN}
	- PageUp {PGUP}
	- Return {RETURN}
	- 向右键 {RIGHT}
	- Scroll Lock {SCROLLLOCK}
	- Tab {TAB}
	- 向上键 {UP}
	- F1 到 F15 {F1} 到 {F15}
	- 要组合的键在键代码之前添加
	- Shift +（加号）
	- Ctrl ^（插入符号）
	- Alt %（百分号）
	- 如下例子：使用说shift+e组合键，调用test过程

```	
	Public Sub OK()
		Application.OnKey "+e", "test"
	End Sub
	Public Sub test()
		MsgBox "VBA最牛逼~"
	End Sub
```
- Application.OnTime 方法
	- 安排一个过程在将来的特定时间运行（既可以是具体指定的某个时间，也可以是指定的一段时间之后）。
	- EarliestTime 必选 Variant 希望此过程运行的时间。
	- Procedure 必选 String 要运行的过程名。
	- LatestTime 可选 Variant 过程开始运行的最晚时间。例如，如果 LatestTime 参数设置为 EarliestTime + 30，且当到达 EarliestTime 时间时，由于其他过程处于运行状态而导致 Microsoft Excel 不能处于“就绪”、“复制”、“剪切”或“查找”模式，则 Microsoft Excel 将等待 30 秒让第一个过程先完成。如果 Microsoft Excel 不能在 30 秒内回到“就绪”模式，则不运行此过程。如果省略该参数，Microsoft Excel 将一直等待到可以运行该过程为止。
	- Schedule 可选 Variant 如果为 True，则预定一个新的 OnTime 过程。如果为 False，则清除先前设置的过程。默认值为 True。
```
	Public Sub test()
		MsgBox "注意休息"
	End Sub
	Public Sub time_out()
		Application.OnTime Now() + TimeValue("00:00:10"), "test" '设置从现在开始 10 秒后运行
	End Sub
```
- 实现每10秒一次
```
	Public Sub test()
		MsgBox "VBA最牛逼~"
		Call time_out
	End Sub
	Public Sub time_out()
		Application.OnTime Now() + TimeValue("00:00:10"), "test"
	End Sub
```
- 以上两个过程，如果不运行一遍，对应指定的过程也不会触发，可以使用open事件，使用call函数，实现打开时自动运行。
```
	Private Sub Workbook_Open()
		Call time_out
		Call OK
	End Sub
```
- UCase 函数
	- 返回 Variant (String)，其中包含转成大写的字符串。只有小写的字母会转成大写；原本大写或非字母之字符保持不变。
- Application.Intersect 方法
	- 返回一个 Range 对象，该对象表示两个或多个区域重叠的矩形区域。表达式.Intersect(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
- 利用change事件快速录入数据
![](https://i.imgur.com/3EJ8mqM.png)
```
	Private Sub Worksheet_Change(ByVal Target As Range)
	If Application.Intersect(Target, Range("B3:B65535")) Is Nothing Or 	Target.Count > 1 Then '目标区域必须为1个且在B列第三个以下
		Exit Sub
	End If
	Dim i As Integer
	i = 3
	Do While Cells(i, "l").value <> ""
		If UCase(Target.value)= Cells(i, "l").value Then '转换成大写
	 
			Application.EnableEvents = False
			Target.value = Cells(i, "l").Offset(0, 1).value
			Target.Offset(0, -1).value = Date
			Target.Offset(0, 1).value = Cells(i, "l").Offset(0, 2).value
			Target.Offset(0, 2).value = Cells(i, "l").Offset(0, 3).value
			Target.Offset(0, 3).Select
			Application.EnableEvents = True
			Exit Sub       '退出过程，减少循环
		End If
		i = i + 1
	Loop
End Sub
```
- 利用SelectionChange事件快速显示选中单元格所有相同单元格内容
![](https://i.imgur.com/trOyTgP.png)
```
	Private Sub Worksheet_SelectionChange(ByVal Target As Range)
	If Application.Intersect(Target, Range("A14").CurrentRegion) Is Nothing Then
		Exit Sub
	End If
	Range("A14").CurrentRegion.Interior.ColorIndex = xlColorIndexNone '清空表格颜色
	If Target.Count > 1 Then
		Set Target = Target.Resize(1, 1)
	End If
	Dim i As Range
	For Each i In Range("A14").CurrentRegion
		If i.value = Target.value Then
			i.Interior.Color = RGB(255, 0, 0)
		End If
	Next
	End Sub
	'高亮行列
	Cells.Interior.ColorIndex = xlColorIndexNone
	Rows(Target.Row).Interior.Color = RGB(255, 0, 0)
	Columns(Target.Column).Interior.Color = RGB(255, 0, 0)
	Target.Interior.ColorIndex = xlColorIndexNone
```
- 移动控件时，按住alt,调整可随控制准确的高度和宽度
	- 如果要将控件与其基础单元格一起排序和筛选，那么为了获得最佳效果，请使用 ActiveX 控件并将该控件调整为基础单元格的准确高度和宽度（在移动单元格和调整其大小时按住 Alt。）
	- 选项按钮控件，选择性别~
```
	Private Sub xb1_Click()
	If xb1.value = True Then
		Range("S3") = "男"
		MsgBox xb2.value
	End If
	End Sub
	Private Sub xb2_Click()
		If xb2.value = True Then
		Range("S3") = "女"
		End If
	End Sub
```
- InputBox 函数
	- 在一对话框来中显示提示，等待用户输入正文或按下按钮，并返回包含文本框内容的 String。
	- InputBox(prompt[, title] [, default] [, xpos] [, ypos] [, helpfile, context])
	- Prompt 必需的。作为对话框消息出现的字符串表达式。prompt 的最大长度大约是 1024 个字符，由所用字符的宽度决定。如果 prompt 包含多个行，则可在各行之间用回车符 (Chr(13))、换行符 (Chr(10)) 或回车换行符的组合 (Chr(13) & Chr(10)) 来分隔。
	- Title 可选的。显示对话框标题栏中的字符串表达式。如果省略 title，则把应用程序名放入标题栏中。
	- Default 可选的。显示文本框中的字符串表达式，在没有其它输入时作为缺省值。如果省略 default，则文本框为空。
	- Xpos 可选的。数值表达式，成对出现，指定对话框的左边与屏幕左边的水平距离。如果省略 xpos，则对话框会在水平方向居中。
	- Ypos 可选的。数值表达式，成对出现，指定对话框的上边与屏幕上边的距离。如果省略 ypos，则对话框被放置在屏幕垂直方向距下边大约三分之一的位置。
	- Helpfile 可选的。字符串表达式，识别帮助文件，用该文件为对话框提供上下文相关的帮助。如果已提供 helpfile，则也必须提供 context。
	- Context 可选的。数值表达式，由帮助文件的作者指定给某个帮助主题的帮助上下文编号。如果已提供 context，则也必须要提供 helpfile。
- Application.InputBox 方法 **比函数多了一个type**
	- 显示一个接收用户输入的对话框。返回此对话框中输入的信息。
	- 表达式.InputBox(Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type)
	- Prompt 必选 String 要在对话框中显示的消息。可为字符串、数字、日期、或布尔值（在显示之前，Microsoft Excel 自动将其值强制转换为 String）。
	- Title 可选 Variant 输入框的标题。如果省略该参数，默认标题将为“Input”。
	- Default 可选 Variant 指定一个初始值，该值在对话框最初显示时出现在文本框中。如果省略该参数，文本框将为空。该值可以是 Range 对象。
	- Left 可选 Variant 指定对话框相对于屏幕左上角的 X 坐标（以磅 （磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位）。
	- Top 可选 Variant 指定对话框相对于屏幕左上角的 Y 坐标（以磅为单位）。
	- HelpFile 可选 Variant 此输入框使用的帮助文件名。如果存在 HelpFile 和 HelpContextID 参数，对话框中将出现一个帮助按钮。
	- HelpContextID 可选 Variant HelpFile 中帮助主题的上下文 ID 号。
	- Type 可选 Variant 指定返回的数据类型。如果省略该参数，对话框将返回文本。
	- 表列出了可以在 Type 参数中传递的值。可以为下列值之一或其中几个值的和。例如，对于一个可接受文本和数字的输入框，将 Type 设置为 1 + 2。
		- 0 公式
		- 1 数字
		- 2 文本（字符串）
		- 4 逻辑值（True 或 False）
		- 8 单元格引用，作为一个 Range 对象
		- 16 错误值，如 #N/A
		- 64 数值数组
- Dim str As Range
- Set str = Application.InputBox("请输入姓名", "我是中华小当家", "姓名", 500, 500,,, 8)
- str.value = 100 '在区域内输入100
-  ![](https://i.imgur.com/r8Y7bm3.png)
- MsgBox 函数
	- 在对话框中显示消息，等待用户单击按钮，并返回一个 Integer 告诉用户单击哪一个按钮。
	- MsgBox(prompt[, buttons] [, title] [, helpfile, context])
	- Prompt 必需的。字符串表达式，作为显示在对话框中的消息。prompt 的最大长度大约为 1024 个字符，由所用字符的宽度决定。如果 prompt 的内容超过一行，则可以在每一行之间用回车符 (Chr(13))、换行符 (Chr(10)) 或是回车与换行符的组合 (Chr(13) & Chr(10)) 将各行分隔开来。
	- Buttons 可选的。数值表达式是值的总和，指定显示按钮的数目及形式，使用的图标样式，缺省按钮是什么以及消息框的强制回应等。如果省略，则 buttons 的缺省值为 0。
	- Title 可选的。在对话框标题栏中显示的字符串表达式。如果省略 title，则将应用程序名放在标题栏中。
	- Helpfile 可选的。字符串表达式，识别用来向对话框提供上下文相关帮助的帮助文件。如果提供了 helpfile，则也必须提供 context。
	- Context 可选的。数值表达式，由帮助文件的作者指定给适当的帮助主题的帮助上下文编号。如果提供了 context，则也必须提供 helpfile。
	- 参数
	- ![](https://i.imgur.com/yracimx.png)
	- 返回值
	- ![](https://i.imgur.com/fGUqDoN.png)
```
	Public Sub box()
	Dim num As Integer
	num = MsgBox("你是谁", 36)
	If num = vbNo Then
	    Range("G11") = num   '结果返回7
	End If
	End Sub
```

