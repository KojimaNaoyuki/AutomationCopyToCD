Option Explicit

' -------------------------------------------------------------- '
'初期設定

'コピー元パス格納ファイルのパスを指定する
Dim dataPathPC : dataPathPC = "E:\趣味\AutomationCopyToCD\data\PC_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '

Dim flag
flag = MsgBox("パソコンの中にある写真を全て削除してよろしいですか？", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    flag = MsgBox("本当に削除してよろしいですか？" &vbCrLf& "削除した写真は復元できません", vbYesNo+vbQuestion, "info")
    If flag = vbYes Then

        Dim DelFilePath : DelFilePath = ReadPath(dataPathPC)

        DelFiles(DelFilePath)

        MsgBox "削除が完了しました"

    Else

        MsgBox "終了します"

    End If

Else

    MsgBox "終了します"

End IF

' -------------------------------------------------------------- '


Function ReadPath(dataPathPC)
    'コピーパスを読み込む

    ' -------------------------------------------------------------- '
    'コピー先, コピー元 パスを取得

    'コピー元パスデータ格納ファイル
    Dim strFromFile : strFromFile = dataPathPC

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")
    Dim objTextFrom : Set objTextFrom = objFS.OpenTextFile(strFromFile)

    Dim str_from_data : str_from_data = ""
    Do While objTextFrom.AtEndOfLine <> True
        str_from_data = objTextFrom.ReadLine
    Loop

    objTextFrom.Close
    ' -------------------------------------------------------------- '

    ReadPath = str_from_data

End Function

Function DelFiles(DelFilePath)
    'ファイル削除を実行

    Dim objFS : Set objFS = CreateObject("Scripting.FileSystemObject")

    objFS.DeleteFolder DelFilePath & "/*"
End Function