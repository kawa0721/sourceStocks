Option Explicit
 
Dim objFileSys
Dim objFolder
Dim objFile
Const conStrPath = "C:\Users\���c���V\Desktop\test"
Dim hoge

'�t�@�C���V�X�e���������I�u�W�F�N�g���쐬
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'c:\temp �t�H���_�̃I�u�W�F�N�g���擾
Set objFolder = objFileSys.GetFolder(conStrPath)
 

'�Ǘ��Ҍ����ɐݒ�
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

'ExecutionPolicy�̕ύX
strCmd = "PowerShell -Command ""Set-ExecutionPolicy Bypass"""
wsh.Run strCmd,0,True

Wscript.Echo "�����J�n!"

Dim progNum
progNum = 1
Dim fileCount
fileCount = objFolder.Files.Count
Dim progRate

Wscript.Echo "�i����:0%"

For Each objFile In objFolder.Files
hoge = " " & objFile

'PowerShell �X�N���v�g�̎��s
strCmd = "cmd /c powershell -file " + "C:\Users\���c���V\Desktop\���o�V�F��vbs�p.ps1" & hoge
wsh.Run strCmd,8,True

progRate = Round(progNum / fileCount * 100, 1)
Wscript.Echo "�i����:" & progRate & "%"
WScript.Echo objFile.Name & "�̏������������܂���!!"
progNum = progNum + 1

Next

Set objFolder = Nothing
Set objFileSys = Nothing 
Wscript.Echo "��������!!!"
