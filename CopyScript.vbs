Option Explicit

' -------------------------------------------------------------- '
'�����ݒ�

'��΃p�X�Ŏw�肷��
Dim VBAFilePath : VBAFilePath = "D:\�\AutomationCopyToCD\VBA\CopyXlsx.xlsm"

'�R�s�[��h���C�u���w�肷��
Dim checkDrive : checkDrive = "D"

'�R�s�[���p�X�i�[�t�@�C���̃p�X���w�肷��
Dim dataPathPC : dataPathPC = "D:\�\AutomationCopyToCD\data\PC_tg_path.txt"

'�R�s�[��p�X�i�[�t�@�C���̃p�X���w�肷��
Dim dataPathCD : dataPathCD = "D:\�\AutomationCopyToCD\data\CD_tg_path.txt"

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'����I��
Dim flag
flag = MsgBox("CD�͂�������Z�b�g���܂������H" &vbCrLf& "�v�����^�[�͓d�����܂������H", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    '�^�[�Q�b�gpath��ǂݍ���
    Dim path : path = ReadPath(dataPathPC, dataPathCD)
    ' ------------------------ '

    ' ------------------------ '
    '�R�s�[�����s
    flag = CopyFnc(path(0), path(1), checkDrive)
    If flag = False Then
        MsgBox "�I�����܂�"
        WScript.Quit
    End If
    ' ------------------------ '

    flag = MsgBox("������܂����H", vbYesNo+vbQuestion, "info")
    IF flag = vbYes Then

        ' ------------------------ '
        'VBA�����s���ĉ摜���������
        VBARunFn(VBAFilePath)
        ' ------------------------ '

    End If

    MsgBox "�I�����܂�"

Else

    MsgBox "�I�����܂�"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_from, str_to, checkDrive)
    '�R�s�[�����s

    Dim objFS : Set objFS = CreateObject("Scripting.FilesystemObject")

    Dim flag : flag = CheckCapacity(checkDrive, str_from)
    If flag  = True Then

        Dim infoPathFlag : infoPathFlag = MsgBox("�R�s�[���t�H���_: " + str_from &vbCrLf& "�R�s�[��t�H���_: " + str_to &vbCrLf& "�R�s�[���J�n���Ă�낵���ł����H", vbOKCancel+vbQuestion, "info")

        If infoPathFlag = vbOK Then

            Call objFS.CopyFolder(str_from, str_to)

            MsgBox "�R�s�[���������܂���"

            CopyFnc = True

        Else

            MsgBox "�R�s�[��t�H���_��R�s�[���t�H���_�̐ݒ�� [setUpScript] �����s���Đݒ肵�Ă�������", vbOKOnly, "info"

            CopyFnc = False

        End If

    Else

        MsgBox "�R�s�[����ɂ͗e�ʂ�����܂���"

        CopyFnc = False

    End If
End Function

Function ReadPath(dataPathPC, dataPathCD)
    'CD�փR�s�[����

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

Function CheckCapacity(tgDrive, tgFolder)
    '�e�ʂ�����邩�̔���

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
    'VBA�����s���ĉ摜���������

    MsgBox "������J�n���܂�"
    
    Dim objExcel : Set objExcel = CreateObject("Excel.Application")

    'VBA�t�@�C���̏ꏊ���L�q����'
    Dim ExcelBook : Set ExcelBook = objExcel.Workbooks.Open(VBAFilePath)

    'Excel�t�@�C�����\��
    objExcel.Visible = False

    objExcel.Run "Module1.PrintOutImg"

    ExcelBook.Close True

    'Excel���I��
    objExcel.Quit

    Set objExcel = Nothing

    MsgBox "������������܂���"

End Function