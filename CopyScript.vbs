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
'�����ݒ�

Dim VBAFilePath : VBAFilePath = "E:\�\AutomationCopyToCD\VBA\CopyXlsx.xlsm" '��΃p�X�Ŏw�肷��

' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
'����I��
Dim flag
flag = MsgBox("CD�͂�������Z�b�g���܂������H", vbYesNo+vbQuestion, "info")
If flag = vbYes Then

    ' ------------------------ '
    '�^�[�Q�b�gpath��ǂݍ���
    path = ReadPath()
    ' ------------------------ '

    ' ------------------------ '
    '�R�s�[�����s
    CopyFnc path(0), path(1)
    ' ------------------------ '

    ' ------------------------ '
    'VBA�����s���ĉ摜���������
    VBARunFn(VBAFilePath)
    ' ------------------------ '

    MsgBox "�I�����܂�"

Else

    MsgBox "�I�����܂�"

End If
' -------------------------------------------------------------- '

Function CopyFnc(str_from, str_to)
    '�R�s�[�����s

    Dim flag : flag = CheckCapacity("C", str_from)
    If flag  = True Then

        infoPathFlag = MsgBox("�R�s�[���t�H���_: " + str_from &vbCrLf& "�R�s�[��t�H���_: " + str_to &vbCrLf& "�R�s�[���J�n���Ă�낵���ł����H", vbOKCancel+vbQuestion, "info")

        If infoPathFlag = vbOK Then

            Call objFS.CopyFolder(str_from, str_to)

            MsgBox "�R�s�[���������܂���"

        Else

            MsgBox "�R�s�[��t�H���_��R�s�[���t�H���_�̐ݒ�� [setUpScript] �����s���Đݒ肵�Ă�������", vbOKOnly, "info"

        End If

    Else

        MsgBox "�R�s�[����ɂ͗e�ʂ�����܂���"

    End If
End Function

Function ReadPath()
    'CD�փR�s�[����

    ' -------------------------------------------------------------- '
    '�R�s�[��, �R�s�[�� �p�X���擾

    '�R�s�[���p�X�f�[�^�i�[�t�@�C��
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
    '�e�ʂ�����邩�̔���
    
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