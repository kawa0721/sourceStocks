Option Explicit
 
Dim objFileSys
Dim objFolder
Dim objFile
Const conStrPath = "C:\Users\川喜田将之\Desktop\test"
Dim hoge

'ファイルシステムを扱うオブジェクトを作成
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'c:\temp フォルダのオブジェクトを取得
Set objFolder = objFileSys.GetFolder(conStrPath)
 

'管理者権限に設定
Dim WMI, OS, Value, Shell
do while WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set OS = WMI.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value in OS
     if left(Value.Version, 3) < 6.0 then exit do
    Next
    Set Shell = CreateObject("Shell.Application")
    Shell.ShellExecute "cscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"
    WScript.Quit
loop

Dim wsh
Dim strPath, strName, strFile, strCmd
Set wsh = WScript.CreateObject("WScript.Shell")

'ExecutionPolicyの変更
strCmd = "PowerShell -Command ""Set-ExecutionPolicy Bypass"""
wsh.Run strCmd,0,True

Wscript.Echo "処理開始!"

Dim progNum
progNum = 1
Dim fileCount
fileCount = objFolder.Files.Count
Dim progRate

Wscript.Echo "進捗率:0%"

For Each objFile In objFolder.Files
hoge = " " & objFile

'PowerShell スクリプトの実行
strCmd = "cmd /c powershell -file " + "C:\Users\川喜田将之\Desktop\抽出シェルvbs用.ps1" & hoge
wsh.Run strCmd,8,True

progRate = Round(progNum / fileCount * 100, 1)
Wscript.Echo "進捗率:" & progRate & "%"
WScript.Echo objFile.Name & "の処理が完了しました!!"
progNum = progNum + 1

Next

Set objFolder = Nothing
Set objFileSys = Nothing 
Wscript.Echo "処理完了!!!"
