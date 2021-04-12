Attribute VB_Name = "Module1"
Option Explicit
Sub selsctDemo()
Attribute selsctDemo.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Integer
i = CInt(InputBox("請輸入名次"))
Select Case i
Case 1
    MsgBox ("冠軍")
Case 2
    MsgBox ("亞軍")
End Select
End Sub
Sub selectDemo2()
Dim i As Integer
Dim tName As Variant
'tName = TypeName(InputBox("請輸入名次"))
'MsgBox (tName)
i = CInt(InputBox("請輸入名次"))
Select Case i
Case 1
    MsgBox ("冠軍")
Case 2
    MsgBox ("亞軍")
Case 3
    MsgBox ("季軍")
Case 4
    MsgBox ("殿軍")
End Select
End Sub
