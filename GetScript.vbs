Option Explicit

' -------------------------------------------------------------- '
'初期設定

'コピー元パス格納ファイルのパスを指定する
Dim dataPathPC : dataPathPC = "E:\趣味\AutomationCopyToCD\data\PC_tg_path.txt"

'コピー先パス格納ファイルのパスを指定する
Dim dataPathCD : dataPathCD = "E:\趣味\AutomationCopyToCD\data\CD_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'動作選択
Dim flag
flag = MsgBox("CDからパソコンに写真を取り込んでよろしいですか？", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    'ターゲットpathを読み込む
    Dim path : path = ReadPath(dataPathPC, dataPathCD)
    ' ------------------------ '

    ' ------------------------ '
    'コピーを実行
    flag = CopyFnc(path(0), path(1))
    If flag = False Then
        MsgBox "終了します"
        WScript.Quit
    End If
    ' ------------------------ '

    MsgBox "終了します"

Else

    MsgBox "終了します"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_to, str_from)
    'コピーを実行

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")

    Dim infoPathFlag : infoPathFlag = MsgBox("コピー元フォルダ: " + str_from &vbCrLf& "コピー先フォルダ: " + str_to &vbCrLf& "取り込みを開始してよろしいですか？", vbOKCancel+vbQuestion, "info")

    If infoPathFlag = vbOK Then

        Call objFS.CopyFolder(str_from, str_to)

        MsgBox "写真の取り込みが完了しました"

        CopyFnc = True

    Else

        MsgBox "コピー先フォルダやコピー元フォルダの設定は [setUpScript] を実行して設定してください", vbOKOnly, "info"

        CopyFnc = False

    End If
End Function

Function ReadPath(dataPathPC, dataPathCD)
    'パスを読み込む

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