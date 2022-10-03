Option Explicit
Dim objXL
Dim GetPathArray
Dim cnt
Dim FileName()
Dim FilePath()
Dim objFSO
Dim scriptPath
Dim ddPath
Set GetPathArray = WScript.Arguments   
Set objFSO = CreateObject("Scripting.FileSystemObject") 
scriptPath = left(WScript.ScriptFullName,len(WScript.ScriptFullName)-len(WScript.ScriptName))

'配列宣言を動的にする
For Each ddPath in GetPathArray
	cnt = cnt + 1
Next
    reDim FileName(cnt)
    reDim FilePath(cnt)
cnt = 0

'D&Dしたファイルパスや名前を配列に落とし込む
For Each ddPath in GetPathArray
    FileName(cnt) = objFSO.GetFileName(ddPath)
    FilePath(cnt) = ddPath
    cnt = cnt + 1
Next

'FilePath(0)～入れた分だけの引数を指定したエクセルを起動して、指定したマクロにぶち込んでを起動する
Set objXL=WScript.CreateObject("Excel.Application")
objXL.Visible=True
objXL.Workbooks.Open scriptPath + "エクセルファイル名.xlsm"
objXL.run "起動したいプロシージャ名",cstr(FilePath(0)),cstr(FilePath(1))'必要になる引数の数に応じて。今回は可変を想定していない

Set objFSO = Nothing