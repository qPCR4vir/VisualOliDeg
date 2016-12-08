VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Align"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "toList, 5, 0, MSForms, CommandButton"
Attribute VB_Control = "Worst, 456, 1, MSForms, CommandButton"
Attribute VB_Control = "Clear, 458, 2, MSForms, CommandButton"
Private Sub Clear_Click()
    Sheets("Align").Columns("J:J").ClearContents
    Range("OnlyWorst") = False
End Sub

Private Sub toList_Click()
    
    If (Range("ActivPrim").Value = "") Or (Range("PrimerLen").Value < 3) Then
    
        MsgBox ("First select (type) a valid primer label in cell ActivPrim")
        
    Else
        Sheets("Oligos").Rows("4:4").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets("Oligos").Rows("1:1").Copy
        Sheets("Oligos").Rows("4:4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Sheets("Align").Select
    End If
    
End Sub

Private Sub Worst_Click()
    Sheets("Align").Columns("F:F").Copy
    Sheets("Align").Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("OnlyWorst") = True
End Sub
