Dim objFSO
Dim objFileFrom
Dim objFileTo

Dim InputFrom
Dim InputTo

InputFrom = InputBox("PC���t�@���_�̃p�X���w�肵�Ă�������(��΃p�X)")
InputTo = InputBox("CD���t�@���_�̃p�X���w�肵�Ă�������(��΃p�X)")

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

Set objFileFrom = objFSO.OpenTextFile("./data/PC_tg_path.txt", 8, True)
Set objFileTo = objFSO.OpenTextFile("./data/CD_tg_path.txt", 8, True)

objFileFrom.WriteLine InputFrom
objFileTo.WriteLine InputTo

objFileFrom.Close
objFileTo.Close
Set objFileFrom = Nothing
Set objFileTo = Nothing

Set objFSO = Nothing

MsgBox "�ݒ芮��"