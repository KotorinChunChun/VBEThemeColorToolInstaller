Rem �{�X�N���v�g��vbetctool.exe���^�X�N�X�P�W���[�����玩���I�Ɏ��s�ł���悤��
Rem �g�p���̃p�\�R���ɃC���X�g�[��/�A���C���X�g�[�����܂��B
Rem
Rem �t�H���_�\��
Rem - VBEThemeColorTool.vbs     ����񎁂��J�������C���X�g�[�����^�X�N�X�P�W���[���Ăяo���p�̖{vbs�t�@�C���ł��B
Rem - vbetctool.exe             ���������J������VBE7.dll�̏��������v���O�����{�̂ł��B
Rem - VBEThemeColorEditor.exe   ���xml�쐬�p�c�[���ł��BVBE7.dll�̏��������ɂ͎g���܂���B
Rem - Themes\*.xml              VBEThemeColorEditor.exe�ō쐬�����ύX�e�[�}��xml��ۑ�����Ƃ���ł��B
Rem
Rem ��L�̃t�@�C����Excel�̃C���X�g�[������Ă���r�b�g����ProgramFiles�ɃC���X�g�[������܂��B
Rem �Ȃ��AVBEThemeColorTool.vbs��VBEThemeColorEditor.exe�̓X�^�[�g�ɃV���[�g�J�b�g��o�^���܂��B
Rem
Rem
Rem �����ŕ��򂷂����
Rem - �Ȃ�     : VBS���Ǘ��҂Ƃ���uac�Ƒ��̈������܂߂ă��_�C���N�g���s���܂��B
Rem - uac      : �Ǘ��Ҏ��s���ł��邱�Ƃ��Ӗ����邽�߁A�ȉ��֏����𑱍s���܂��B
Rem - change   : �����XML�𔽉f����v���V�[�W�����Ăт܂��B
Rem - default  : �W����XML�𔽉f����v���V�[�W�����Ăт܂��B
Rem - ��L���� : GUI�ŃC���X�g�[��/�A���C���X�g�[��/�L�����Z���I���ł��܂��B
Rem   - �C���X�g�[�����͈ȉ���2���s���܂��B
Rem     1. �K�v�ȃt�@�C����Program Files�փR�s�[
Rem     2. �^�X�N�X�P�W���[���ɖ������s����悤�ɓo�^�i������Change�j
Rem   - �A���C���X�g�[�����͏�L2���������܂��B
Rem

Option Explicit

Const APP_NAME = "VBEThemeColorTool"

Rem ------VBA�Ńf�o�b�O����Ƃ��̓R�����g�A�E�g���镔��-----
Call CheckAdmin
Call VbsMain
Rem --------------------------------------------------------

Rem �J����VBE�f�o�b�O�p
Sub VbeDebugPathOverride()
    DebugSetScriptFullName "D:\OneDrive - �������邿��񂿂��\ExcelVBA\���J���x��\VBA�J�X�^�}�C�Y\VBEThemeColorTool\vbetctool_install.vbs"
End Sub

Sub VbetctoolSetMyColor()
    wsShell.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
    
    Rem --------------------------------------------------------
    Rem �{�X�N���v�g�ŃJ�X�^�}�C�Y���K�v�ȉӏ�
    Rem --------------------------------------------------------
    Dim ThemeXml, ForeColors, BackColors
    ThemeXml = """.\Themes\kc.xml"""
    ForeColors = """1 0 5 0 1 2 13 15 4 7"""
    BackColors = """4 0 4 7 6 4 4 4 13 4"""
    Rem --------------------------------------------------------
    
    Rem �R�}���h�ł̎��s��
    Rem cd���Ȃ��ƃo�b�`���s����.\Temes���s����
    Rem cd C:\Program Files\VBEThemeColorTool
    Rem vbetctool -l "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" -t ".\Themes\kc.xml" -f "1 0 5 0 1 2 13 15 4 7" -b "4 0 4 7 6 4 4 4 13 4" -V
    wsShell.Run "vbetctool -l " & GetPath_VBE7DllFileWQ & " -t " & ThemeXml & " -f " & ForeColors & " -b " & BackColors & " -V", 0, True
    Rem cmd /c ���āA������1�ɂ���ƁA�R�}���h�v�����v�g���\������Ē�~���邽�߃f�o�b�O�ŕ֗��i�������B���Ӂj
    'wsShell.Run "cmd /c vbetctool -l " & VBE7DLL & " -t " & ThemeXml & " -f " & ForeColors & " -b " & BackColors & " -V & pause", 1, true
End Sub

Sub VbetctoolSetDefaultColor()
    wsShell.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
    
    Dim ThemeXml, ForeColors, BackColors
    ThemeXml = """.\Themes\Default VBE.xml"""
    ForeColors = """1 0 5 0 1 2 13 15 4 7"""
    BackColors = """4 0 4 7 6 4 4 4 13 4"""
    
    wsShell.Run "vbetctool -l " & GetPath_VBE7DllFileWQ & " -t " & ThemeXml & " -f " & ForeColors & " -b " & BackColors & " -V", 0, True
End Sub

Rem �C���X�g�[���Ɋ܂߂�t�@�C���̑��΃p�X���X�g
Rem �c�[���{�́AVBS�AThemes�t�H���_��XML�S��
Function GetInstallFiles()
    Dim dic: Set dic = NewDic()
    dic.Add "VBEThemeColorEditor.exe", ""
    dic.Add GetPath_VbsFileName, ""
    dic.Add "vbetctool.exe", ""
    Dim fl
    For Each fl In fso.GetFolder(GetPath_VbsFolderPath() & "Themes").Files
        Rem ���΃p�X�̏�������
        dic.Add Replace(fl.Path, GetPath_VbsFolderPath, ""), ""
    Next
    GetInstallFiles = dic.Keys()
End Function

Function SW_Change(): SW_Change = "change": End Function
Function SW_Default(): SW_Default = "default": End Function

Sub VbsMain()
    Rem �p�����[�^ change �� default ������Ƃ��͐F�ݒ���s���I������i�X�P�W���[������̌Ăяo���p�j
    Dim arg
    For Each arg In WScript.Arguments
        If arg = SW_Change Then: Call VbetctoolSetMyColor: Exit Sub
        If arg = SW_Default Then: Call VbetctoolSetDefaultColor: Exit Sub
    Next
    
    Rem �p�����[�^���������̓C���X�g�[��or�A���C���X�g�[�����s��
    Select Case MsgBox(" �͂� �F�C���X�g�[��" & vbLf & "�������F�A���C���X�g�[��", vbYesNoCancel, APP_NAME)
        Case vbYes
            Call Install
            MsgBox "�C���X�g�[�� ����", vbOKOnly, APP_NAME
        Case vbNo
            Call Uninstall
            MsgBox "�A���C���X�g�[�� ����", vbOKOnly, APP_NAME
        Case Else
            MsgBox "�L�����Z������܂����B", vbOKOnly, APP_NAME
    End Select
End Sub

Sub Install()
    
    Dim VbsInstalledFolder: VbsInstalledFolder = GetPath_AppInstallFolder()
    Dim VbsInstalledFullName: VbsInstalledFullName = VbsInstalledFolder & GetPath_VbsFileName
    
    Rem Excel�̃I�v�V�����̕ύX
    'wsShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "�l�r �o�S�V�b�N,11", "REG_SZ"

    Rem �t�@�C�����O����
    Dim fl
    For Each fl In GetInstallFiles
        CopyFileEx GetPath_VbsFolderPath & "\" & fl, VbsInstalledFolder & "\" & fl
    Next
    
    Dim SW
    Select Case MsgBox("����VBE7�̐F��ύX���܂��B�\�񃂁[�h�@�͂��F�F��ύX����@�������F�F��߂�", vbYesNo, APP_NAME)
        Case vbYes: SW = SW_Change
        Case vbNo: SW = SW_Default
    End Select
    
    Rem 1�������Ɏ��s������@
    Rem schtasks /create /tn VBEThemeColorTool /tr "\"wscript\" \"C:\Program Files\\VBEThemeColorTool\VBEThemeColorTool.vbs\" \"change\"" /sc minute /mo 1 /rl highest /F
    Rem schtasks /create /tn VBEThemeColorTool /tr "\"wscript\" \"C:\Program Files\\VBEThemeColorTool\VBEThemeColorTool.vbs\" \"default\"" /sc minute /mo 1 /rl highest /F
    Const CmdTemplate = "schtasks /create /tn [TASK_NAME] /tr ""[EXE_PATH] [PARAM_TEXT]"" /sc minute /mo 1 /rl highest /F"
    Dim ss: ss = CmdTemplate
    ss = Replace(ss, "[TASK_NAME]", APP_NAME)
    ss = Replace(ss, "[EXE_PATH]", "\""wscript\""")
    ss = Replace(ss, "[PARAM_TEXT]", "\""" & VbsInstalledFullName & "\"" \""" & SW & "\""")
    wsShell.Run ss, 0, True
    
    MsgBox ss
    
    Rem schtasks /create /tn AUTO_BUILD /tr c:\test.vbs /sc minute /mo 1
'    Const CmdTemplate = "schtasks /create /tn [TASK_NAME] /tr \""""[EXE_PATH][PARAM_TEXT]\"""" /sc minute /mo 1 /rl highest /F"
'    Dim ss: ss = CmdTemplate
'    ss = Replace(ss, "[TASK_NAME]", APP_NAME)
'    ss = Replace(ss, "[APP_NAME]", "wscript")
'    ss = Replace(ss, "[EXE_PATH]", fd & FnFile1)
'    ss = Replace(ss, "[PARAM_TEXT]", " \""" & VbsInstalledFullName & "\"" \""/hide\""")
'    ss = Replace(ss, "[PARAM_TEXT]", "")
'    wsShell.exec ss
    'msgbox ss
    
    Rem �Ƃ肠����������1����s���Ă���
    Rem �쐬����͎��s����Ȃ��炵���̂Œx��������B
    'WScript.Sleep 1000
    'wsShell.exec "schtasks /run /tn " & APP_NAME & ""
    

    'Rem �V�K�쐬�̃��W�X�g���ǉ�
    'wsShell.RegWrite "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName", PATH_NEW & "EXCEL12.XLSM", "REG_SZ"
    'Rem ���L�͌��ʂȂ��B Excel.Sheet.8��xls_auto_file�̒�`���Ȃ����炾�Ǝv����
    'Rem REG ADD "HKEY_CLASSES_ROOT\.xls\Excel.Sheet.8\ShellNew" /v "FileName" /t REG_SZ /d "C:\Program Files (x86)\Microsoft Office\Root\VFS\Windows\ShellNew\EXCEL8.XLS" /f
    'wsShell.RegWrite "HKCR\.xls\ShellNew\FileName", PATH_NEW & "EXCEL8.XLS", "REG_SZ"
    'Rem [�t�@�C���̎��]�̒�`�B�R�R���󗓂���ShellNew�ɓo�^���Ă����j���[�ɑ����Ȃ��B
    'wsShell.RegWrite "HKCU\Software\Classes\xls_auto_file", "Microsoft Excel 97-2003 �݊��u�b�N", "REG_SZ"
    'Rem ����l��\�ŏI��
    'wsShell.RegWrite "HKCR\xls_auto_file\", "Microsoft Excel 97-2003 �݊��u�b�N", "REG_SZ"
    
    Rem 1�ڂ�2�ڂɃX�^�[�g�ɓo�^�������v���O����������O��
    Call CreateShortcutInStartMenu(VbsInstalledFolder & GetInstallFiles()(0))
    Call CreateShortcutInStartMenu(VbsInstalledFolder & GetInstallFiles()(1))
End Sub

Sub Uninstall()

    Rem �F��߂�
    Call VbetctoolSetDefaultColor
    
    Rem �X�^�[�g�̃V���[�g�J�b�g���폜����
    On Error Resume Next
    fso.DeleteFile GetPath_StartPrograms() & "\" & fso.GetBaseName(GetInstallFiles()(0)) & ".lnk"
    fso.DeleteFile GetPath_StartPrograms() & "\" & fso.GetBaseName(GetInstallFiles()(1)) & ".lnk"
    On Error GoTo 0

    Dim fd: fd = GetPath_AppInstallFolder()
    
    Call DeleteFolderAndFiles(fd)

    Const CmdTemplate = "schtasks /delete /tn [TASK_NAME] /F"
    Dim ss: ss = CmdTemplate
    ss = Replace(ss, "[TASK_NAME]", APP_NAME)
    wsShell.exec ss

    Rem �e���v���[�g�uBook.xltx�v�̍폜
    Rem If fso.FileExists( fd & "\" & FnFile1) Then
    Rem     fso.DeleteFile  fd & "\" & FnFile1
    Rem End If

End Sub



Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------

Rem ----------------------------------------------------------------------------------------------------
Rem Excel�C���X�g�[�����擾�֐�

Function GetPath_ExcelInstalledBit()
    Rem Excel.EXE�̑��݂���ProgramFiles�̃p�X�����
    If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Then
        GetPath_ExcelInstalledBit = 32
    ElseIf fso.FileExists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then
        GetPath_ExcelInstalledBit = 64
    Else
        GetPath_ExcelInstalledBit = 0
    End If
End Function

Function GetPath_ExcelInstalledProgramFiles()
    Select Case GetPath_ExcelInstalledBit
        Case 32: GetPath_ExcelInstalledProgramFiles = "C:\Program Files (x86)\"
        Case 64: GetPath_ExcelInstalledProgramFiles = "C:\Program Files\"
        Case Else:
            MsgBox "�Ή����Ă���o�[�W������Excel���C���X�g�[������Ă��܂���B"
            GetPath_ExcelInstalledProgramFiles = ""
    End Select
End Function

Function GetPath_ExcelTemplateFolder()
    GetPath_ExcelTemplateFolder = GetPath_ExcelInstalledProgramFiles() & "\Microsoft Office\root\Office16\XLSTART\"
End Function

Function GetPath_ExcelShellNewFolder()
    GetPath_ExcelShellNewFolder = GetPath_ExcelInstalledProgramFiles() & "\Microsoft Office\root\VFS\Windows\SHELLNEW\"
End Function

Function GetPath_VBE7DllFileWQ()
    Select Case GetPath_ExcelInstalledBit
        Case 32: GetPath_VBE7DllFileWQ = """" & GetPath_ExcelInstalledProgramFiles() & "\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" & """"
        Case 64: GetPath_VBE7DllFileWQ = """" & GetPath_ExcelInstalledProgramFiles() & "\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" & """"
        Case Else: GetPath_VBE7DllFileWQ = ""
    End Select
End Function

Function GetPath_AppInstallFolder()
    GetPath_AppInstallFolder = GetPath_ExcelInstalledProgramFiles & "\" & APP_NAME & "\"
End Function

Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------
Rem ---------------------------------------------------------------------------------------------------

Rem �Ǘ��Ҍ����m�F
Sub CheckAdmin()

    Dim Args
    Dim UacFlag '�Ǘ��Ҍ����t���O�i�p�����[�^��uac���܂ގ�True�j
    Set Args = WScript.Arguments

    Dim i
    For i = 0 To Args.Count - 1
      If Args(i) = "uac" Then UacFlag = True
    Next

    ' �Ǘ��Ҍ����ɏ��i
    'Dim WScript    'VBE�ł̃R�[�h�`�F�b�N�p
    Dim Param
    Do While UacFlag = False And WScript.Version >= 5.7

      '���݂̃p�����[�^���X�y�[�X��؂�ɕϊ�
      Param = ""
      For i = 0 To Args.Count - 1
        Param = Param & " " & Args(i)
      Next

      ' Check WScript5.7~ and Vista~
      Dim os, wmi, Value
      Set wmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
      Set os = wmi.ExecQuery("SELECT *FROM Win32_OperatingSystem")
      For Each Value In os
        If Left(Value.Version, 3) < 6 Then Exit Do    'Exit if not vista
      Next

      ' Run this script as admin.
      Dim sha
      Set sha = CreateObject("Shell.Application")
      sha.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac" & Param, "", "runas"

      WScript.Quit
    Loop

End Sub

Rem ------------------------�֐�-------------------------------
Rem �悭�g���I�u�W�F�N�g���܂Ƃ߂Ē�`
Rem Dim WScript Rem ��VBS�ł͒�`�s�v
Public Function wsShell(): Set wsShell = WScript.CreateObject("WScript.Shell"): End Function
Public Function wsNet(): Set wsNet = WScript.CreateObject("WScript.Network"): End Function
Public Function shellAp(): Set shellAp = WScript.CreateObject("Shell.Application"): End Function
Public Function fso(): Set fso = WScript.CreateObject("Scripting.FileSystemObject"): End Function
Public Function NewDic(): Set NewDic = WScript.CreateObject("Scripting.Dictionary"): End Function

Rem VBS�̃t���p�X�AVBS�̃t�H���_�p�X�A�t�@�C�����A�g���q�������t�@�C����
Public Function GetPath_VbsFullName(): GetPath_VbsFullName = WScript.ScriptFullName: End Function
Public Function GetPath_VbsFolderPath(): GetPath_VbsFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, ""): End Function
Public Function GetPath_VbsFileName(): GetPath_VbsFileName = WScript.ScriptName: End Function
Public Function GetPath_VbsBaseName(): GetPath_VbsBaseName = Mid(WScript.ScriptName, 1, InStrRev(WScript.ScriptName, ".") - 1): End Function
Rem C:\Users\XXXXX\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
Public Function GetPath_StartPrograms(): GetPath_StartPrograms = wsShell.SpecialFolders("Programs"): End Function
Public Function GetPath_Desktop(): GetPath_Desktop = wsShell.SpecialFolders("Desktop"): End Function

Rem �X�^�[�g���j���[�ɃV���[�g�J�b�g���쐬����
Sub CreateShortcutInStartMenu(ssTargetFullName)
    Dim shortcut
    Set shortcut = wsShell.CreateShortcut(GetPath_StartPrograms() & "\" & fso.GetBaseName(ssTargetFullName) & ".lnk")
    With shortcut
        .targetPath = ssTargetFullName
        .WorkingDirectory = GetPath_StartPrograms()
        .Save
    End With
End Sub

Rem �t�H���_�̈ꊇ�쐬
Private Sub CreateDirectoryExFso(ByVal strPath)
    '�p�X���t�@�C���̎��Ƀt�H���_�쐬��h������
    Dim without_LastFileName: without_LastFileName = False
    Dim s, v, f
    Dim i
    v = Split(strPath, "\")
    'On Error Resume Next
    For i = LBound(v) To UBound(v)
        If without_LastFileName And i = UBound(v) Then Exit For
    
        If f = "" Then f = v(i) Else f = f & "\" & v(i)
        
        If Not fso.FolderExists(f) Then
'            MsgBox f
            fso.CreateFolder f & "\"
        End If
    Next

End Sub

Function CopyFileEx(srcFp, destFp)
    Call CreateDirectoryExFso(fso.GetParentFolderName(destFp))
    If Not fso.FileExists(srcFp) Then
        MsgBox "�R�s�[���F" & srcFp & "���݂���܂���"
        Exit Function
    End If
    On Error Resume Next
    fso.CopyFile srcFp, destFp
    Dim errMsg: errMsg = Err.Description
    On Error GoTo 0
    If errMsg <> "" Then
        Const MsgTemplate = "�R�s�[���F[src]\n�R�s�[��F[dest]\n�G���[���b�Z�[�W�F[errMsg]"
        Dim ss: ss = MsgTemplate
        ss = Replace(ss, "\n", vbLf)
        ss = Replace(ss, "[src]", srcFp)
        ss = Replace(ss, "[dest]", destFp)
        ss = Replace(ss, "[errMsg]", errMsg)
        MsgBox ss
    End If
End Function

Rem �w�肵���t�H���_�̒��g�����ׂč폜����B
Sub DeleteFolderAndFiles(DirectoryPath)
    Dim objFolder: Set objFolder = fso.GetFolder(DirectoryPath)
    
    On Error Resume Next
    ' �T�u�t�H���_���擾���čċN�Ăяo���B
    Dim objSubFolder
    For Each objSubFolder In objFolder.SubFolders
        DeleteFolderAndFiles (DirectoryPath & "\" & objSubFolder.Name)
        fso.DeleteFolder DirectoryPath & "\" & objSubFolder.Name, True
    Next

    ' �t�H���_���̃t�@�C�����擾���A�폜����B
    Dim fileName
    For Each fileName In objFolder.Files
        fso.DeleteFile DirectoryPath & "\" & fileName.Name, True
    Next
End Sub

Rem----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------

