Attribute VB_Name = "Module1"
Option Explicit
Sub selsctDemo()
Attribute selsctDemo.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Integer
i = CInt(InputBox("�п�J�W��"))
Select Case i
Case 1
    MsgBox ("�a�x")
Case 2
    MsgBox ("�ȭx")
End Select
End Sub
Sub selectDemo2()
Dim i As Integer
Dim tName As Variant
'tName = TypeName(InputBox("�п�J�W��"))
'MsgBox (tName)
i = CInt(InputBox("�п�J�W��"))
Select Case i
Case 1
    MsgBox ("�a�x")
Case 2
    MsgBox ("�ȭx")
Case 3
    MsgBox ("�u�x")
Case 4
    MsgBox ("���x")
End Select
End Sub
