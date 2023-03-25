Rem 本スクリプトはvbetctool.exeをタスクスケジューラから自動的に実行できるように
Rem 使用中のパソコンにインストール/アンインストールします。
Rem
Rem フォルダ構成
Rem - VBEThemeColorTool.vbs     ちゅん氏が開発したインストーラ兼タスクスケジューラ呼び出し用の本vbsファイルです。
Rem - vbetctool.exe             風柳氏が開発したVBE7.dllの書き換えプログラム本体です。
Rem - VBEThemeColorEditor.exe   例のxml作成用ツールです。VBE7.dllの書き換えには使いません。
Rem - Themes\*.xml              VBEThemeColorEditor.exeで作成した変更テーマのxmlを保存するところです。
Rem
Rem 上記のファイルがExcelのインストールされているビット数のProgramFilesにインストールされます。
Rem なお、VBEThemeColorTool.vbsとVBEThemeColorEditor.exeはスタートにショートカットを登録します。
Rem
Rem
Rem 引数で分岐する役割
Rem - なし     : VBSを管理者としてuacと他の引数を含めてリダイレクト実行します。
Rem - uac      : 管理者実行中であることを意味するため、以下へ処理を続行します。
Rem - change   : 自作のXMLを反映するプロシージャを呼びます。
Rem - default  : 標準のXMLを反映するプロシージャを呼びます。
Rem - 上記無し : GUIでインストール/アンインストール/キャンセル選択できます。
Rem   - インストール時は以下の2つを行います。
Rem     1. 必要なファイルをProgram Filesへコピー
Rem     2. タスクスケジューラに毎分実行するように登録（引数はChange）
Rem   - アンインストール時は上記2つを取り消します。
Rem

Option Explicit

Const APP_NAME = "VBEThemeColorTool"

Rem ------VBAでデバッグするときはコメントアウトする部分-----
Call CheckAdmin
Call VbsMain
Rem --------------------------------------------------------

Rem 開発者VBEデバッグ用
Sub VbeDebugPathOverride()
    DebugSetScriptFullName "D:\OneDrive - えくせるちゅんちゅん\ExcelVBA\●開発支援\VBAカスタマイズ\VBEThemeColorTool\vbetctool_install.vbs"
End Sub

Sub VbetctoolSetMyColor()
    wsShell.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
    
    Rem --------------------------------------------------------
    Rem 本スクリプトでカスタマイズが必要な箇所
    Rem --------------------------------------------------------
    Dim ThemeXml, ForeColors, BackColors
    ThemeXml = """.\Themes\kc.xml"""
    ForeColors = """1 0 5 0 1 2 13 15 4 7"""
    BackColors = """4 0 4 7 6 4 4 4 13 4"""
    Rem --------------------------------------------------------
    
    Rem コマンドでの実行例
    Rem cdしないとバッチ実行時に.\Temes失敗する
    Rem cd C:\Program Files\VBEThemeColorTool
    Rem vbetctool -l "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL" -t ".\Themes\kc.xml" -f "1 0 5 0 1 2 13 15 4 7" -b "4 0 4 7 6 4 4 4 13 4" -V
    wsShell.Run "vbetctool -l " & GetPath_VBE7DllFileWQ & " -t " & ThemeXml & " -f " & ForeColors & " -b " & BackColors & " -V", 0, True
    Rem cmd /c して、引数を1にすると、コマンドプロンプトが表示されて停止するためデバッグで便利（毎分増殖注意）
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

Rem インストールに含めるファイルの相対パスリスト
Rem ツール本体、VBS、ThemesフォルダのXML全部
Function GetInstallFiles()
    Dim dic: Set dic = NewDic()
    dic.Add "VBEThemeColorEditor.exe", ""
    dic.Add GetPath_VbsFileName, ""
    dic.Add "vbetctool.exe", ""
    Dim fl
    For Each fl In fso.GetFolder(GetPath_VbsFolderPath() & "Themes").Files
        Rem 相対パスの書き込み
        dic.Add Replace(fl.Path, GetPath_VbsFolderPath, ""), ""
    Next
    GetInstallFiles = dic.Keys()
End Function

Function SW_Change(): SW_Change = "change": End Function
Function SW_Default(): SW_Default = "default": End Function

Sub VbsMain()
    Rem パラメータ change や default があるときは色設定を行い終了する（スケジュールからの呼び出し用）
    Dim arg
    For Each arg In WScript.Arguments
        If arg = SW_Change Then: Call VbetctoolSetMyColor: Exit Sub
        If arg = SW_Default Then: Call VbetctoolSetDefaultColor: Exit Sub
    Next
    
    Rem パラメータが無い時はインストールorアンインストールを行う
    Select Case MsgBox(" はい ：インストール" & vbLf & "いいえ：アンインストール", vbYesNoCancel, APP_NAME)
        Case vbYes
            Call Install
            MsgBox "インストール 完了", vbOKOnly, APP_NAME
        Case vbNo
            Call Uninstall
            MsgBox "アンインストール 完了", vbOKOnly, APP_NAME
        Case Else
            MsgBox "キャンセルされました。", vbOKOnly, APP_NAME
    End Select
End Sub

Sub Install()
    
    Dim VbsInstalledFolder: VbsInstalledFolder = GetPath_AppInstallFolder()
    Dim VbsInstalledFullName: VbsInstalledFullName = VbsInstalledFolder & GetPath_VbsFileName
    
    Rem Excelのオプションの変更
    'wsShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "ＭＳ Ｐゴシック,11", "REG_SZ"

    Rem ファイル事前準備
    Dim fl
    For Each fl In GetInstallFiles
        CopyFileEx GetPath_VbsFolderPath & "\" & fl, VbsInstalledFolder & "\" & fl
    Next
    
    Dim SW
    Select Case MsgBox("毎分VBE7の色を変更します。予約モード　はい：色を変更する　いいえ：色を戻す", vbYesNo, APP_NAME)
        Case vbYes: SW = SW_Change
        Case vbNo: SW = SW_Default
    End Select
    
    Rem 1分おきに実行する方法
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
    
    Rem とりあえず今すぐ1回実行しておく
    Rem 作成直後は実行されないらしいので遅延させる。
    'WScript.Sleep 1000
    'wsShell.exec "schtasks /run /tn " & APP_NAME & ""
    

    'Rem 新規作成のレジストリ追加
    'wsShell.RegWrite "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName", PATH_NEW & "EXCEL12.XLSM", "REG_SZ"
    'Rem 下記は効果なし。 Excel.Sheet.8はxls_auto_fileの定義がないからだと思われる
    'Rem REG ADD "HKEY_CLASSES_ROOT\.xls\Excel.Sheet.8\ShellNew" /v "FileName" /t REG_SZ /d "C:\Program Files (x86)\Microsoft Office\Root\VFS\Windows\ShellNew\EXCEL8.XLS" /f
    'wsShell.RegWrite "HKCR\.xls\ShellNew\FileName", PATH_NEW & "EXCEL8.XLS", "REG_SZ"
    'Rem [ファイルの種類]の定義。ココが空欄だとShellNewに登録してもメニューに増えない。
    'wsShell.RegWrite "HKCU\Software\Classes\xls_auto_file", "Microsoft Excel 97-2003 互換ブック", "REG_SZ"
    'Rem 既定値は\で終了
    'wsShell.RegWrite "HKCR\xls_auto_file\", "Microsoft Excel 97-2003 互換ブック", "REG_SZ"
    
    Rem 1つ目と2つ目にスタートに登録したいプログラムがある前提
    Call CreateShortcutInStartMenu(VbsInstalledFolder & GetInstallFiles()(0))
    Call CreateShortcutInStartMenu(VbsInstalledFolder & GetInstallFiles()(1))
End Sub

Sub Uninstall()

    Rem 色を戻す
    Call VbetctoolSetDefaultColor
    
    Rem スタートのショートカットを削除する
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

    Rem テンプレート「Book.xltx」の削除
    Rem If fso.FileExists( fd & "\" & FnFile1) Then
    Rem     fso.DeleteFile  fd & "\" & FnFile1
    Rem End If

End Sub



Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------

Rem ----------------------------------------------------------------------------------------------------
Rem Excelインストール情報取得関数

Function GetPath_ExcelInstalledBit()
    Rem Excel.EXEの存在からProgramFilesのパスを特定
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
            MsgBox "対応しているバージョンのExcelがインストールされていません。"
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

Rem 管理者権限確認
Sub CheckAdmin()

    Dim Args
    Dim UacFlag '管理者権限フラグ（パラメータにuacを含む時True）
    Set Args = WScript.Arguments

    Dim i
    For i = 0 To Args.Count - 1
      If Args(i) = "uac" Then UacFlag = True
    Next

    ' 管理者権限に昇格
    'Dim WScript    'VBEでのコードチェック用
    Dim Param
    Do While UacFlag = False And WScript.Version >= 5.7

      '現在のパラメータをスペース区切りに変換
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

Rem ------------------------関数-------------------------------
Rem よく使うオブジェクトをまとめて定義
Rem Dim WScript Rem ※VBSでは定義不要
Public Function wsShell(): Set wsShell = WScript.CreateObject("WScript.Shell"): End Function
Public Function wsNet(): Set wsNet = WScript.CreateObject("WScript.Network"): End Function
Public Function shellAp(): Set shellAp = WScript.CreateObject("Shell.Application"): End Function
Public Function fso(): Set fso = WScript.CreateObject("Scripting.FileSystemObject"): End Function
Public Function NewDic(): Set NewDic = WScript.CreateObject("Scripting.Dictionary"): End Function

Rem VBSのフルパス、VBSのフォルダパス、ファイル名、拡張子を除くファイル名
Public Function GetPath_VbsFullName(): GetPath_VbsFullName = WScript.ScriptFullName: End Function
Public Function GetPath_VbsFolderPath(): GetPath_VbsFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, ""): End Function
Public Function GetPath_VbsFileName(): GetPath_VbsFileName = WScript.ScriptName: End Function
Public Function GetPath_VbsBaseName(): GetPath_VbsBaseName = Mid(WScript.ScriptName, 1, InStrRev(WScript.ScriptName, ".") - 1): End Function
Rem C:\Users\XXXXX\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
Public Function GetPath_StartPrograms(): GetPath_StartPrograms = wsShell.SpecialFolders("Programs"): End Function
Public Function GetPath_Desktop(): GetPath_Desktop = wsShell.SpecialFolders("Desktop"): End Function

Rem スタートメニューにショートカットを作成する
Sub CreateShortcutInStartMenu(ssTargetFullName)
    Dim shortcut
    Set shortcut = wsShell.CreateShortcut(GetPath_StartPrograms() & "\" & fso.GetBaseName(ssTargetFullName) & ".lnk")
    With shortcut
        .targetPath = ssTargetFullName
        .WorkingDirectory = GetPath_StartPrograms()
        .Save
    End With
End Sub

Rem フォルダの一括作成
Private Sub CreateDirectoryExFso(ByVal strPath)
    'パスがファイルの時にフォルダ作成を防ぐため
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
        MsgBox "コピー元：" & srcFp & "がみつかりません"
        Exit Function
    End If
    On Error Resume Next
    fso.CopyFile srcFp, destFp
    Dim errMsg: errMsg = Err.Description
    On Error GoTo 0
    If errMsg <> "" Then
        Const MsgTemplate = "コピー元：[src]\nコピー先：[dest]\nエラーメッセージ：[errMsg]"
        Dim ss: ss = MsgTemplate
        ss = Replace(ss, "\n", vbLf)
        ss = Replace(ss, "[src]", srcFp)
        ss = Replace(ss, "[dest]", destFp)
        ss = Replace(ss, "[errMsg]", errMsg)
        MsgBox ss
    End If
End Function

Rem 指定したフォルダの中身をすべて削除する。
Sub DeleteFolderAndFiles(DirectoryPath)
    Dim objFolder: Set objFolder = fso.GetFolder(DirectoryPath)
    
    On Error Resume Next
    ' サブフォルダを取得して再起呼び出し。
    Dim objSubFolder
    For Each objSubFolder In objFolder.SubFolders
        DeleteFolderAndFiles (DirectoryPath & "\" & objSubFolder.Name)
        fso.DeleteFolder DirectoryPath & "\" & objSubFolder.Name, True
    Next

    ' フォルダ内のファイルを取得し、削除する。
    Dim fileName
    For Each fileName In objFolder.Files
        fso.DeleteFile DirectoryPath & "\" & fileName.Name, True
    Next
End Sub

Rem----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------
Rem ----------------------------------------------------------------------------------------------------

