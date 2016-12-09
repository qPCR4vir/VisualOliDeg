VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "cbt_Set_FASTAfile, 3, 0, MSForms, CommandButton"
Attribute VB_Control = "cbt_AddSeq, 4, 1, MSForms, CommandButton"
Attribute VB_Control = "Reset, 5, 2, MSForms, CommandButton"
Attribute VB_Control = "cbt_ExportSeq, 6, 3, MSForms, CommandButton"
Private Sub cbt_ExportSeq_Click()
    Call VisualOliDeg.ExportFASTAfile
End Sub


Private Sub cbt_Set_FASTAfile_Click()
    Dim file
    file = Excel.Application.GetOpenFilename("FASTA files (*.fasta;*.fas;*seq;*.txt),*.fasta;*.fas;*seq;*.txt, All Files (*.*),*.* ", , "Select yours aligned sequences in FASTA format")
    If file <> False Then Range("FastaFileNAme") = file
End Sub


Private Sub cbt_AddSeq_Click()
    Application.ScreenUpdating = False
    Excel.Application.Calculation = xlCalculationManual
    
    Call AddSeqFromFASTAfile
    
    'Excel.Calculate                 ' comment??
    Excel.Application.Calculation = xlCalculationAutomatic

    Application.ScreenUpdating = True
    
    Range("ClassHeaders").Offset(1, 0).Select
    Range("ClassHeaders").Offset(1, 0).Show
    'Range("Align.Data").Offset(1, 0).Activate
    'Range("Align.Data").Offset(1, 0).Show
    'Range("Gr").Select
End Sub


Private Sub Reset_Click()
    Call VisualOliDeg.ResetAll_Click
End Sub
