**声明：**这个文档只是用来记一些自己在日常办公中常用到的一些**命令行**和**代码**，泥萌不准笑话我(/▽╲)

#[目录]

*	[**BAT**](#BAT)
	*	[合并文件](#Combine)
	*	[命名文件](#Delete)
		*	[删除关键字](#Delete)
		*	[替换关键字](#Replace)
*	[**CMD**](#CMD)
	*	[Copy](#Copy)
	*	[Xcopy](#Xcopy)
	*	[Hash](#Hash)
*	[**VBA**](#VBA)
	*	[Excel](#CBAS)
		*	[合并所有工作表](#CBAS)
		*	[合并所有工作薄](#CBAF)
	*	[Word](#Hiline)
		*	[高亮含关键字行](#Hiline)
		*	[删除含关键字行](#Deline)
*	[**VBS**](#VBS)


---
<h4 id="BAT">BAT</h4>

*	<h5 id="Combine">合并文件:</h5>
```Bash
@echo off
	copy /b 000.ts+001.ts+002.ts+003.ts+004.ts 005.ts
pause
```  
*	<h5 id="Delete">删除关键字:</h5>
```Bash
@echo off& setlocal enabledelayedexpansion
for /f "delims=" %%1 in ('dir /a /b') do (set wind=%%1
	ren "%%~1" "!wind:关键字=!")
pause
```
*	<h5 id="Replace">替换关键字:</h5>
```Bash
@echo off
for /f "delims=" %%i in ('dir/b/a-d *目标*')do (set f=%%i
    echo.%%i
    call set f=%%f:目标=替换%%
    call ren "%%i" "%%f%%")
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
