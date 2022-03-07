Option Explicit

' -------------------------------------------------------------- '
'�����ݒ�

'�R�s�[���p�X�i�[�t�@�C���̃p�X���w�肷��
Dim dataPathPC : dataPathPC = "E:\�\AutomationCopyToCD\data\PC_tg_path.txt"

'�R�s�[��p�X�i�[�t�@�C���̃p�X���w�肷��
Dim dataPathCD : dataPathCD = "E:\�\AutomationCopyToCD\data\CD_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'����I��
Dim flag
flag = MsgBox("CD����p�\�R���Ɏʐ^����荞��ł�낵���ł����H", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    '�^�[�Q�b�gpath��ǂݍ���
    Dim path : path = ReadPath(dataPathPC, dataPathCD)
    ' ------------------------ '

    ' ------------------------ '
    '�R�s�[�����s
    flag = CopyFnc(path(0), path(1))
    If flag = False Then
        MsgBox "�I�����܂�"
        WScript.Quit
    End If
    ' ------------------------ '

    MsgBox "�I�����܂�"

Else

    MsgBox "�I�����܂�"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_to, str_from)
    '�R�s�[�����s

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")

    Dim infoPathFlag : infoPathFlag = MsgBox("�R�s�[���t�H���_: " + str_from &vbCrLf& "�R�s�[��t�H���_: " + str_to &vbCrLf& "��荞�݂��J�n���Ă�낵���ł����H", vbOKCancel+vbQuestion, "info")

    If infoPathFlag = vbOK Then

        Call objFS.CopyFolder(str_from, str_to)

        MsgBox "�ʐ^�̎�荞�݂��������܂���"

        CopyFnc = True

    Else

        MsgBox "�R�s�[��t�H���_��R�s�[���t�H���_�̐ݒ�� [setUpScript] �����s���Đݒ肵�Ă�������", vbOKOnly, "info"

        CopyFnc = False

    End If
End Function

Function ReadPath(dataPathPC, dataPathCD)
    '�p�X��ǂݍ���

    ' -------------------------------------------------------------- '
    '�R�s�[��, �R�s�[�� �p�X���擾

    '�R�s�[���p�X�f�[�^�i�[�t�@�C��
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