Option Explicit

' -------------------------------------------------------------- '
'初期設定

'絶対パスで指定する
Dim VBAFilePath : VBAFilePath = "D:\趣味\AutomationCopyToCD\VBA\CopyXlsx.xlsm"

'コピー先ドライブを指定する
Dim checkDrive : checkDrive = "D"

'コピー元パス格納ファイルのパスを指定する
Dim dataPathPC : dataPathPC = "D:\趣味\AutomationCopyToCD\data\PC_tg_path.txt"

'コピー先パス格納ファイルのパスを指定する
Dim dataPathCD : dataPathCD = "D:\趣味\AutomationCopyToCD\data\CD_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'動作選択
Dim flag
flag = MsgBox("CDはしっかりセットしましたか？" &vbCrLf& "プリンターは電源つけましたか？", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    'ターゲットpathを読み込む
    Dim path : path = ReadPath(dataPathPC, dataPathCD)
    ' ------------------------ '

    ' ------------------------ '
    'コピーを実行
    flag = CopyFnc(path(0), path(1), checkDrive)
    If flag = False Then
        MsgBox "終了します"
        WScript.Quit
    End If
    ' ------------------------ '

    flag = MsgBox("印刷しますか？", vbYesNo+vbQuestion, "info")
    IF flag = vbYes Then

        ' ------------------------ '
        'VBAを実行して画像を印刷する
        VBARunFn(VBAFilePath)
        ' ------------------------ '

    End If

    MsgBox "終了します"

Else

    MsgBox "終了します"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_from, str_to, checkDrive)
    'コピーを実行

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")

    Dim flag : flag = CheckCapacity(checkDrive, str_from)
    If flag  = True Then

        Dim infoPathFlag : infoPathFlag = MsgBox("コピー元フォルダ: " + str_from &vbCrLf& "コピー先フォルダ: " + str_to &vbCrLf& "コピーを開始してよろしいですか？", vbOKCancel+vbQuestion, "info")

        If infoPathFlag = vbOK Then

            Call objFS.CopyFolder(str_from, str_to)

            MsgBox "コピーが完了しました"

            CopyFnc = True

        Else

            MsgBox "コピー先フォルダやコピー元フォルダの設定は [setUpScript] を実行して設定してください", vbOKOnly, "info"

            CopyFnc = False

        End If

    Else

        MsgBox "コピーするには容量が足りません"

        CopyFnc = False

    End If
End Function

Function ReadPath(dataPathPC, dataPathCD)
    'CDへコピーする

    ' -------------------------------------------------------------- '
    'コピー先, コピー元 パスを取得

    'コピー元パスデータ格納ファイル
    Dim strFromFile : strFromFile = dataPathPC
    Dim strToFile : strToFile = dataPathCD

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")
    Dim objTextFrom : Set objTextFrom = objFS.OpenTextFile(strFromFile)
    Dim objTextTo : Set objTextTo = objFS.OpenTextFile(strToFile)

    Dim str_from_data : str_from_data = ""
    Dim str_to_data : str_to_data = ""
    Do While objTextFrom.AtEndOfLine <> True
        str_from_data = objTextFrom.ReadLine
    Loop
    Do While objTextTo.AtEndOfLine <> True
        str_to_data = objTextTo.ReadLine
    Loop

    objTextFrom.Close
    objTextTo.Close
    ' -------------------------------------------------------------- '

    ReadPath = Array(str_from_data, str_to_data)

End Function

Function CheckCapacity(tgDrive, tgFolder)
    '容量が足りるかの判定

    Dim objFSO : Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    Dim objDrive : Set objDrive = objFSO.GetDrive(tgDrive)
    Dim objFolder : Set objFolder = objFSO.GetFolder(tgFolder)

    If objDrive.FreeSpace > objFolder.Size Then

        'MsgBox FormatNumber(objDrive.FreeSpace/1048576, 0) + " MB > " + FormatNumber(objFolder.Size/1048576, 0) + " MB"
        CheckCapacity = True

    Else

        'MsgBox FormatNumber(objDrive.FreeSpace/1048576, 0) + " MB < " + FormatNumber(objFolder.Size/1048576, 0) + " MB"
        CheckCapacity = False

    End If

    Set objDrive = Nothing
    Set objFSO = Nothing

End Function

Function VBARunFn(VBAFilePath)
    'VBAを実行して画像を印刷する

    MsgBox "印刷を開始します"
    
    Dim objExcel : Set objExcel = CreateObject("Excel.Application")

    'VBAファイルの場所を記述する'
    Dim ExcelBook : Set ExcelBook = objExcel.Workbooks.Open(VBAFilePath)

    'Excelファイルを非表示
    objExcel.Visible = False

    objExcel.Run "Module1.PrintOutImg"

    ExcelBook.Close True

    'Excelを終了
    objExcel.Quit

    Set objExcel = Nothing

    MsgBox "印刷が完了しました"

End Function