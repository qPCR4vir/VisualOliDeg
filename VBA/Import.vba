VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
