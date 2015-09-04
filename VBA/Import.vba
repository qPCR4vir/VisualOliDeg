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
    Call VisualOliDeg.SetFASTAfile
End Sub


Private Sub cbt_AddSeq_Click()
    Call VisualOliDeg.AddSeq_Click
End Sub


Private Sub Reset_Click()
    Call VisualOliDeg.ResetAll_Click
End Sub
