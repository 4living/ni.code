**声明：**这个文档只是用来记一些自己在日常办公中常用到的一些**命令行**和**代码**，泥萌不准笑话我(/▽╲)

#[目录]

*	[**BAT**](#BAT)
	*	[合并文件](#Combine)
	*	[命名文件](#Delete)
		*	[删除关键字](#Delete)
		*	[替换关键字](#Replace)
	*	[新建文件夹](#Md)
	*	[新建连续文件夹并复制文件](#Md&Copy)
*	[**CMD**](#CMD)
	*	[Copy](#Copy)
	*	[Xcopy](#Xcopy)
	*	[Hash](#Hash)
*	[**IPV6**](#IPV6)
	*	[关闭IPV6接口](#Clsipv6)
	*	[还原IPV6隧道](#Rstipv6)
*	[**VB**](#VB)
	*	[CommonDialog Save](#cmdSave)
*	[**VBA**](#VBA)
	*	[Excel](#CBAS)
		*	[合并所有工作表](#CBAS)
		*	[合并所有工作薄](#CBAF)
	*	[Word](#Hiline)
		*	[高亮含关键字行](#Hiline)
		*	[删除含关键字行](#Deline)
*	[**VBS**](#VBS)
*	[**Word**](#Word)
	*	[查找/替换](#Find&Replace)
*	[**Excel**](#Excel)
	*	[单元格格式](#CellFormat)


---
<h4 id="BAT">BAT</h4>

*	<h5 id="Combine">合并文件:</h5>
```Bash
@echo off
	copy /b 000.ts+001.ts+002.ts+003.ts+004.ts 005.ts
pause
```  

```Bash
@echo off
	cd.>文件名.txt
	for /f "tokens=*" %%i in ('dir/on/b *.txt') do type "%%i">>文件名.txt
pause
```  
*	<h5 id="Delete">删除关键字:</h5>
```Bash
@echo off& setlocal enabledelayedexpansion
for /f "delims=" %%1 in ('dir /a /b') do (
	set wind=%%1
	ren "%%~1" "!wind:关键字=!")
pause
```
*	<h5 id="Replace">替换关键字:</h5>
```Bash
@echo off
for /f "delims=" %%i in ('dir/b/a-d *目标*')do (
	set f=%%i
    echo.%%i
    call set f=%%f:目标=替换%%
    call ren "%%i" "%%f%%")
pause
```

*	<h5 id="Md">新建文件夹:</h5>
```Bash		
@echo off
	cd\ 
	md A:\...\[file name]
	md B:\...\[file name]
	···
	md X:\...\[file name]
pause
```

*	<h5 id="Md&Copy">新建连续文件夹并复制文件:</h5>
```Bash		
@echo off
for /l %%k in (start_number,step,end_number) do (
	md %%k
	copy /y combine.bat %%k
	copy /y result.xlsx %%k
)
pause
```

---
<h4 id="CMD">CMD</h4>

*	<h5 id="Copy">Copy:</h5>
```CMD
copy *.txt all.txt
```
*	<h5 id="Hash">Hash:</h5>
```CMD
FCIV -md5 -sha1 path\filename.ext
```
*	<h5 id="Mklink">Mklink:</h5>
```CMD	
mklink /j "X:\...\..." "Y:\...\..."
```
*	<h5 id="Xcopy">Xcopy:</h5>
```CMD
xcopy A\*.* B: /s /h /d /y
```

---
<h4 id="IPV6">IPV6</h4>

*	<h5 id="Clsipv6">关闭IPV6接口：</h5>
```Bash
netsh interface teredo set state disable 
netsh interface 6to4 set state disabled 
netsh interface isatap set state disabled 
```
*	<h5 id="Rstipv6">还原IPV6隧道：</h5>
```Bash
netsh interface teredo set state disable 
netsh interface 6to4 set state disabled 
netsh interface isatap set state disabled 
```
---
<h4 id="VB">VB</h4>
*	<h5 id="cmdSave">CommonDialog Save：</h5>
```vb
Private Sub cmdSave_Click()
On Error GoTo userCanceled
    With dlg
        .FileName = "result"
        .InitDir = App.Path
        .CancelError = True
        .Filter = "文本文件(*.txt)|*.txt"
        .ShowSave
    End With
    'Text1.Text = CommonDialog1.FileName
userCanceled:
End Sub
```

---
<h4 id="VBA">VBA</h4>

*	<h5 id="CBAS">合并当前工作簿的全部工作表：</h5>
```vb
Sub 合并当前工作簿的全部工作表()
Dim FilesToOpen, ft
Dim x As Integer

Application.ScreenUpdating = False
On Error GoTo errhandler

FilesToOpen = Application.GetOpenFilename _
(FileFilter:="Micrsofe Excel文件(*.xls), *.xls", _
MultiSelect:=True, Title:="要合并的文件")
 
If TypeName(FilesToOpen) = "boolean" Then
MsgBox "没有选定文件"
'GoTo errhandler
End If
x = 1
While x <= UBound(FilesToOpen)
Set wk = Workbooks.Open(Filename:=FilesToOpen(x))
wk.Sheets().Move after:=ThisWorkbook.Sheets _
(ThisWorkbook.Sheets.Count)
x = x + 1
Wend

MsgBox "合并成功完成！"

errhandler:
'MsgBox Err.Description
'Resume errhandler
End Sub
```
*	<h5 id="CBAF">合并当前目录下所有工作簿的全部工作表：</h5>
```VB		
Sub 合并当前目录下所有工作簿的全部工作表()
Dim MyPath, MyName, AWbName
Dim Wb As Workbook, WbN As String
Dim G As Long
Dim Num As Long
Dim BOX As String
Application.ScreenUpdating = False
MyPath = ActiveWorkbook.Path
MyName = Dir(MyPath & "\" & "*.xls")
AWbName = ActiveWorkbook.Name
Num = 0
Do While MyName <> ""
If MyName <> AWbName Then
Set Wb = Workbooks.Open(MyPath & "\" & MyName)
Num = Num + 1
With Workbooks(1).ActiveSheet
.Cells(.Range("B65536").End(xlUp).Row + 2, 1) = Left(MyName, Len(MyName) - 4)
For G = 1 To Sheets.Count
Wb.Sheets(G).UsedRange.Copy .Cells(.Range("B65536").End(xlUp).Row + 1, 1)
Next
WbN = WbN & Chr(13) & Wb.Name
Wb.Close False
End With
End If
MyName = Dir
Loop
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "共合并了" & Num & "个工作薄下的全部工作表。如下：" & Chr(13) & WbN, vbInformation, "提示"
End Sub
```
*	<h5 id="Hiline">删除含关键字行：</h5>
```VB
Sub 高亮含关键字行()
Dim i As Paragraph
n = InputBox("请输入关键字")
Application.ScreenUpdating = False
For Each i In ActiveDocument.Paragraphs
Selection.Find.ClearFormatting
Selection.Find.text = n
Selection.Find.Execute
If Selection.text = n Then
Selection.HomeKey Unit:=wdLine
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
'添加字符底纹
Selection.Shading.BackgroundPatternColor = wdColorLightYellow
End If
Selection.Find.Execute
Next i
End Sub
```
*	<h5 id="Deline">删除含关键字行：</h5>
```VB	
Sub 删除含关键字行()
Dim i As Paragraph
n = InputBox("请输入关键字")
Application.ScreenUpdating = False
For Each i In ActiveDocument.Paragraphs
Selection.Find.ClearFormatting
Selection.Find.text = n
Selection.Find.Execute
If Selection.text = n Then
Selection.HomeKey Unit:=wdLine
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
'删除含关键字行
Selection.Delete Unit:=wdCharacter, Count:=1
End If
Selection.Find.Execute
Next i
End Sub
```
---
<h4 id="VBS">VBS</h4>

```VBS
CreateObject("SAPI.SpVoice").Speak "德玛西亚"
```
---
<h4 id="Word">Word</h4>

*	<h5 id="Find&Replace">查找/替换：</h5>
	- [x]	使用通配符
	*	**中文**：`[!^1-^127]` 或 `[一-龥]`
	*	**英文**：`[a-z]` 或 `[A-Z]`
	*	**数字**：`[数字-数字]`
	*	**组合**：`[a-z/A-Z/1-9]`
	*	**本体**：`^&`

---
<h4 id="Excel">Excel</h4>

*	<h5 id="CellFormat">批量修改A列单元格部分内容的格式：</h5>
把公式下拉至B5，复制B1：B5，调出剪贴版，单击剪贴板的内容，这时单击右键—选择性粘贴—Unicode 文本—确定
```html
="<table><tr><td><font size=16 face=隶书 color=red>"&SUBSTITUTE(A1,RIGHT(A1,3),"")&"</font>"&RIGHT(A1,3)
```
