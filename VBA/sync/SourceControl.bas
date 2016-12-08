Attribute VB_Name = "SourceControl"
'   http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
'   http://www.pretentiousname.com/excel_extractvba/

' https://github.com/brucemcpherson/VbaGit/blob/master/scripts/VbaGit.vba#L299

' https://support.office.com/en-us/article/Overview-of-Spreadsheet-Compare-13fafa61-62aa-451b-8674-242ce5f2c986

' https://wiki.ucl.ac.uk/display/~ucftpw2/2013/10/18/Using+git+for+version+control+of+spreadsheet+models+-+part+1+of+3


' Function VersionNum() As String
'     VersionNum = "v1.06"
     ' Introducing version number in Code to facilitate commits comments in git/GitHub. 2016-12-06.
' End Function


Sub CommitToLaptoop()
    CommitVBA ("C:\Prog\VisualOliDeg\VBA\")
End Sub

Sub CommitToC()
    CommitVBA ("C:\Prog\VisualOliDeg\VBA\")
End Sub

Sub CommitToDrive() 'dont works
    CommitVBA ("C:\Users\Ariel\OneDrive\Documents\Tesis\Flavivirus\VBA\")
End Sub



Sub CommitVBA(dir As String)

    ' Exports all VBA modules
    Dim i%, sName$

    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.name
                sName$ = dir & sName$ & ".vba"
                .VBComponents(i%).Export (sName$)
            End If
        Next i
    End With


End Sub


Sub RevertVBA()
  
  Dim Line As String

  For Each VBComp In ThisWorkbook.VBProject.VBComponents
     With VBComp.CodeModule
        On Error GoTo Skip
           Open "C:\Prog\VisualOliDeg\VBA\" & .name & ".vba" For Input As #VBA
           .DeleteLines 1, .CountOfLines
           Do While Not EOF(VBA)
              Line Input #VBA, Line
              .InsertLines Line
           Loop
           Close #VBA
Skip:

     End With
  Next VBComp


End Sub
