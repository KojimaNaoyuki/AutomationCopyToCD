Option Explicit

' -------------------------------------------------------------- '
'�����ݒ�

'�R�s�[���p�X�i�[�t�@�C���̃p�X���w�肷��
Dim dataPathPC : dataPathPC = "E:\�\AutomationCopyToCD\data\PC_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '

Dim flag
flag = MsgBox("�p�\�R���̒��ɂ���ʐ^��S�č폜���Ă�낵���ł����H", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    flag = MsgBox("�{���ɍ폜���Ă�낵���ł����H" &vbCrLf& "�폜�����ʐ^�͕����ł��܂���", vbYesNo+vbQuestion, "info")
    If flag = vbYes Then

        Dim DelFilePath : DelFilePath = ReadPath(dataPathPC)

        DelFiles(DelFilePath)

        MsgBox "�폜���������܂���"

    Else

        MsgBox "�I�����܂�"

    End If

Else

    MsgBox "�I�����܂�"

End IF

' -------------------------------------------------------------- '


Function ReadPath(dataPathPC)
    '�R�s�[�p�X��ǂݍ���

    ' -------------------------------------------------------------- '
    '�R�s�[��, �R�s�[�� �p�X���擾

    '�R�s�[���p�X�f�[�^�i�[�t�@�C��
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
    '�t�@�C���폜�����s

    Dim objFS : Set objFS = CreateObject("Scripting.FileSystemObject")

    objFS.DeleteFolder DelFilePath & "/*"
End Function