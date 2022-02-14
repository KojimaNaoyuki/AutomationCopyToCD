Dim str_from
Dim str_to
Dim Input

Dim objFS
Dim objTextFrom
Dim objTextTo
Dim str_from_data
Dim str_to_data
Dim strFromFile
Dim strToFile

' -------------------------------------------------------------- '
'初期設定

Dim VBAFilePath : VBAFilePath = "E:\趣味\AutomationCopyToCD\VBA\CopyXlsx.xlsm" '絶対パスで指定する

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'動作選択
Dim flag
flag = MsgBox("CDはしっかりセットしましたか？", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    'ターゲットpathを読み込む
    path = ReadPath()
    ' ------------------------ '

    ' ------------------------ '
    'コピーを実行
    CopyFnc path(0), path(1)
    ' ------------------------ '

    ' ------------------------ '
    'VBAを実行して画像を印刷する
    VBARunFn(VBAFilePath)
    ' ------------------------ '

    MsgBox "終了します"

Else

    MsgBox "終了します"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_from, str_to)
    'コピーを実行

    Dim flag : flag = CheckCapacity("C", str_from)
    If flag  = True Then

        infoPathFlag = MsgBox("コピー元フォルダ: " + str_from &vbCrLf& "コピー先フォルダ: " + str_to &vbCrLf& "コピーを開始してよろしいですか？", vbOKCancel+vbQuestion, "info")

        If infoPathFlag = vbOK Then

            Call objFS.CopyFolder(str_from, str_to)

            MsgBox "コピーが完了しました"

        Else

            MsgBox "コピー先フォルダやコピー元フォルダの設定は [setUpScript] を実行して設定してください", vbOKOnly, "info"

        End If

    Else

        MsgBox "コピーするには容量が足りません"

    End If
End Function

Function ReadPath()
    'CDへコピーする

    ' -------------------------------------------------------------- '
    'コピー先, コピー元 パスを取得

    'コピー元パスデータ格納ファイル
    strFromFile = "./data/PC_tg_path.txt"
    strToFile = "./data/CD_tg_path.txt"

    Set objFS = CreateObject("Scripting.FilesystemObject")
    Set objTextFrom = objFS.OpenTextFile(strFromFile)
    Set objTextTo = objFS.OpenTextFile(strToFile)

    str_from_data = ""
    str_to_data = ""
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
    
    Dim objFSO
    Dim objFolder
    Dim objDrive

    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    Set objDrive = objFSO.GetDrive(tgDrive)
    Set objFolder = objFSO.GetFolder(tgFolder)

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