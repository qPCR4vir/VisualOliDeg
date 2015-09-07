Attribute VB_Name = "VisualOliDeg"
' see: http://msdn.microsoft.com/en-us/library/office/ff196650%28v=office.15%29.aspx
' and: http://blogs.office.com/2005/10/14/conditional-formatting-using-vba-some-examples/
'    : http://www.mrexcel.com/forum/excel-questions/522196-conditional-formatting-excel-visual-basic-applications.html

Option Explicit

Private Sub Tm_CBox_Change()
    Call ThDy.Tm_Set_Change
End Sub

Sub AddSeqFromFASTAfile()
'Load a region of each seq from the FASTA file to separated rows in sheet Import
    
    Dim LoadSeqFrom   As Long
        LoadSeqFrom = Range("LoadSeqFrom")
        If LoadSeqFrom < 1 Then
           LoadSeqFrom = 1
           Range("LoadSeqFrom") = LoadSeqFrom
        End If
    Dim LoadSeqTo    As Long
        LoadSeqTo = Range("LoadSeqTo")
    Dim MaxSeqLength As Long
        MaxSeqLength = Range("MaxSeqLength")
        
    If LoadSeqTo <= 0 Or LoadSeqTo > LoadSeqFrom + MaxSeqLength - 1 Then LoadSeqTo = LoadSeqFrom + MaxSeqLength - 1
    
    Dim Descript   As Range
    Set Descript = Range("Description") '        Table HEAD for seq: Descriptions
    
    Dim nrSeqDesc As String
        nrSeqDesc = "SeqDescriptions"
    Dim rSeqDesc   As Range
    Set rSeqDesc = Range(nrSeqDesc)        '    Todos los datos de la tabla
    Dim ClassHeaders As Range
    Set ClassHeaders = Range("ClassHeaders")
    Dim NumClass    As Long
        NumClass = ClassHeaders.Columns.Count      ' 4  !
    
    Dim NameCol As Long
        NameCol = 1
    Dim FastaNames As Range
    Set FastaNames = rSeqDesc.Columns(NameCol) ' Columna con los nombres de sec.
    
    Dim SeqCol As Long
        SeqCol = 2
    Dim Seq As Range
    Set Seq = rSeqDesc.Columns(SeqCol)  '        Columna de sec.
    
    Dim NtImportedCol As Long
        NtImportedCol = 8
    Dim DescCol As Long   ' Class1Col As Long,
        DescCol = 7
    
                                                 'Set VisODegWB = Excel.Application.ActiveWorkbook '     ----- ??
    Dim NofSeq As Long
        NofSeq = 0
    Dim Nt  As Long             ' The running nt in the current seq -pos del nt ya leido
        Nt = 0
    Dim maxNt As Long           ' the largest seq read until now
        maxNt = 0
    
    Dim L    As Long
    Dim ci   As Long, cf As Long
    Dim p1   As Long, p2 As Long
    Dim clas As Long
    
    Dim eqCell As Range
    
    
    Dim Line   As String
    Dim sequence As String
    Dim Class  As String
    Dim ClassH As String
    Dim SeqCell As Range
    
    Dim FastaFile As Integer
    FastaFile = FreeFile '                    genera el siguiente # disponible para ID de file
    Open Range("FastaFileNAme") For Input As #FastaFile
    
    Do While Not EOF(FastaFile)
      Line Input #FastaFile, Line
      Line = Trim(Line) '                 ---> Hace falta?
      If (Line <> "") Then   ' 11                  Ignora lineas en blanco
      
        If (Line Like ">*") Then        '  10
        
        '   New seq. Example: >EU303182.Apoi Grupo:TB Clas:RioBrVG Especie:Apoi Lineage:
        '            parse name, description and possible classifications
        
            
            If NofSeq < 1 Then
                rSeqDesc.ClearContents:
                ClassHeaders.ClearContents
            Else
                SeqCell.Value = sequence
                If Nt < LoadSeqFrom Then
                    rSeqDesc(NofSeq, NtImportedCol) = 0
                Else
                    rSeqDesc(NofSeq, NtImportedCol) = Nt - LoadSeqFrom + 1 ' write the actual number of nt readed (well, counting all sort of gaps)
                    If Nt > maxNt Then maxNt = Nt
                End If
            End If
            
            NofSeq = NofSeq + 1
            Set SeqCell = Seq.Rows(NofSeq)
            sequence = ""
            Line = LTrim(Mid$(Line, 2))       '  Erase >  Mid(string, start[, length])   start - index 1-based
            Nt = 0
            
            Dim name As String
            p2 = InStr(Line, " ")         'p2- position of the first space, 0 if no space, tentatively after name
            If p2 = 0 Then '   == 0
              name = Line
              Line = ""
              p1 = 0
            Else: '                           Contiene algo mas que el nombre
              name = Left$(Line, p2 - 1)
              Line = LTrim(Mid$(Line, p2))
              p1 = InStr(Line, ":")         'p1- position of the first :, tentatively after first class
            End If
              rSeqDesc(NofSeq, NameCol) = name
            
            Dim curClass As Long
            curClass = 0
            Do While p1 > 0 And curClass < NumClass '      Contiene al menos un level: class1   ----> restaurar en Header o agragar ERROR
                curClass = curClass + 1             '      like Grupo:TB
                ClassH = RTrim(Left$(Line, p1 - 1)) '   class name, to find or set the class header
                Line = Mid$(Line, p1 + 1)
                
                p2 = InStr(Line, " ")  '          Contain more classes ?
                If p2 = 0 Then '7
                  Class = Line
                  Line = ""
                  p1 = 0
                Else: ' 7                         Contain more classes
                  Class = Left$(Line, p2 - 1)
                  Line = LTrim(Mid$(Line, p2))
                  p1 = InStr(Line, ":")
                End If '7
                
                If NofSeq = 1 Then             '  8   La primera seq es la que pone los Class headers.
                  ClassHeaders(, curClass) = ClassH
                  ClassHeaders(NofSeq + 1, curClass) = Class
                Else             '   8
                  Dim i     As Long
                  Dim found As Boolean
                  found = False
                  For i = 1 To NumClass    '         busco classH en los headesrs.
                    If ClassH = ClassHeaders(, i) Then
                      If ClassH <> "" Or i >= curClass Then  ' si lo encontro o si es un Header en blanco y no hemos usado esa pos
                        found = True
                        curClass = i
                        ClassHeaders(NofSeq + 1, curClass) = Class
                        Exit For
                      End If
                    End If
                  Next i
                  If Not found Then
                    If ClassH = "" Then
                        ClassHeaders(NofSeq + 1, curClass) = Class
                    Else
                        rSeqDesc(NofSeq, DescCol) = rSeqDesc(NofSeq, DescCol) & ClassH & ":" & Class & " "
                    End If
                  End If
                End If   '8
            Loop
            rSeqDesc(NofSeq, DescCol) = rSeqDesc(NofSeq, DescCol) & Line
          
          
          '   Parse the nt seq.
          
          
        ElseIf LoadSeqTo > Nt Then  '10
            L = Len(Line)
            
            ci = LoadSeqFrom - Nt
            If ci < 1 Then ' b
                ci = 1
            ElseIf ci > L Then 'b
                ci = L + 1
            End If 'b
            
            cf = LoadSeqTo - Nt
            If cf < 1 Then 'a
                cf = 0
            ElseIf cf > L Then 'a
                cf = L
            End If 'a
            
            L = cf - ci + 1
            If L > 0 Then sequence = sequence + Mid(Line, ci, L)
            
            Nt = Nt + cf
            
        End If '10
        
        
     End If '11  -->  If (Line <> "") Then   ' 11                  Ignora lineas en blanco
    Loop
    Close #FastaFile
    
    If NofSeq < 1 Then
        MsgBox "No sequences were found and this file is ignored."
        Return
    End If
                
                SeqCell.Value = sequence
                If Nt < LoadSeqFrom Then
                    rSeqDesc(NofSeq, NtImportedCol) = 0
                Else
                    rSeqDesc(NofSeq, NtImportedCol) = Nt - LoadSeqFrom + 1 ' write the actual number of nt readed (well, counting all sort of gaps)
                    If Nt > maxNt Then maxNt = Nt
                End If
    
    If maxNt < 2 Then
        MsgBox "No sequences which at least 2 nt were found. Revise the sequences and the From/To range."
        maxNt = 2
    Else
        maxNt = maxNt - LoadSeqFrom + 1
    End If
                
    
    Range("NoSeq") = NofSeq
    Range("NoNt") = maxNt
   
    
    Call AdjustColHrow("Align.primer_mark", maxNt, Clear:=True)
    Excel.Range("ActivPrim").ClearContents
    Call AdjustRowHcol("Match.CountErr", NofSeq)
    
    Call AdjustRange("Align.Data", maxNt, NofSeq)
    Call AdjustRange("Match.Data", maxNt, NofSeq)

   
 If rSeqDesc.Rows.Count > NofSeq Then
   Range(rSeqDesc.Rows(NofSeq + 1), rSeqDesc.Rows(rSeqDesc.Rows.Count)).EntireRow.Delete
 End If
 
 Set rSeqDesc = rSeqDesc.Resize(NofSeq)
    
 rSeqDesc.Rows(1).Copy
 rSeqDesc.PasteSpecial (xlPasteFormats)
        
 Excel.Names.Add nrSeqDesc, rSeqDesc
 'Range("SeqName").Offset(1, 0).sc
 

End Sub


Sub AdjustRowHcol(ColName As String, ByVal numR As Long, Optional formula, Optional Clear = False)
    
 If numR <= 0 Then numR = 1    'Para que siempre quede al menos la primera Row para poderla copiar

 With Excel.Range(ColName)
   Excel.ThisWorkbook.Names.Add ColName, .Resize(numR)
   If Clear <> False Then
        .Rows(1).ClearContents
   Else
   If Not IsMissing(formula) Then 'If formula <> "none" Then '
        .Rows(1).formula = formula
     End If
   End If
      
 End With
End Sub

Sub AdjustColHrow(RowName As String, ByVal numC As Long, Optional formula = "none", Optional Clear = False)

    
 If numC <= 0 Then numC = 1    'Para que siempre quede al menos la primera col para poderla copiar

 With Excel.Range(RowName)
   Excel.ThisWorkbook.Names.Add RowName, .Resize(, numC)          'Excel.Range(.Columns(1), .Columns(numC))
   If Clear <> False Then
        .Columns(1).ClearContents
   Else
     If formula <> "none" Then
        .Columns(1).formula = formula
     End If
   End If
      
 End With
End Sub



Sub AdjustRange(RangeName As String, ByVal numCols As Long, ByVal numRows As Long)    ' AdjustRange("Align.ColsAll", maxNt, NofSeq)

    If numRows <= 0 Then numRows = 1  'Para que siempre quede al menos la primera Row para poderla copiar
    If numCols <= 0 Then numCols = 1  'Para que siempre quede al menos la primera Row para poderla copiar
 
 Dim Data As Range
 Set Data = Excel.Range(RangeName)
 
 If Data.Rows.Count > numRows Then
   Excel.Range(Data.Rows(numRows + 1), Data.Rows(Data.Rows.Count)).EntireRow.Delete
 End If
 If Data.Columns.Count > numCols Then
   Excel.Range(Data.Columns(numCols + 1), Data.Columns(Data.Columns.Count)).EntireColumn.Delete
 End If
 
 Set Data = Data.Resize(numRows, numCols)
 Data(1, 1).Copy
 Data.PasteSpecial (xlPasteAll)
 Data.ColumnWidth = Data(1, 1).ColumnWidth
 Data.RowHeight = Data(1, 1).RowHeight
 
 Dim ColH As Range
 Dim RowH As Range
 With Data.Worksheet
   Set ColH = .Range(.Cells(1, Data.Column), .Cells(Data.Row - 1, Data.Columns.Count + Data.Column - 1))
    ColH.Columns(1).Copy
    ColH.PasteSpecial (xlPasteAll)
    
   Set RowH = .Range(.Cells(Data.Row, 1), .Cells(Data.Row + numRows - 1, Data.Column - 1))
    RowH.Rows(1).Copy
    RowH.PasteSpecial (xlPasteAll)
 End With
 
 Excel.ThisWorkbook.Names.Add RangeName, Data
    
End Sub

Sub ExportFASTAfile()
    ExportFileNAme = Excel.Application.GetSaveAsFilename(Range("FastaFileNAme"), "FASTA files (*.fasta;*.fas;*seq;*.txt),*.fasta;*.fas;*seq;*.txt, All Files (*.*),*.* ", , "Export current sequences in FASTA format to a file:")
    Dim Ident As String, Class As String, Clas1 As String, Clas2 As String, Clas3 As String, Clas4 As String
    Dim HeadName As Range, rSeqDesc As Range
    Set HeadName = Range("SeqName")
    Clas1 = HeadName(1, 3)
    Clas2 = HeadName(1, 4)
    Clas3 = HeadName(1, 5)
    Clas4 = HeadName(1, 6)

    FastaFile = FreeFile
    Open ExportFileNAme For Output As #FastaFile
    Set rSeqDesc = Range("SeqDescriptions")
    For Each Row In rSeqDesc.Rows
       Ident = ">" & Row.Cells(1, 1) & " "
       clas = Clas1 & ":" & CStr(Row.Cells(1, 3)) & " "
       clas = clas & Clas2 & ":" & CStr(Row.Cells(1, 4)) & " "
       clas = clas & Clas3 & ":" & CStr(Row.Cells(1, 5)) & " "
       clas = clas & Clas4 & ":" & CStr(Row.Cells(1, 6)) & " "
       Print #FastaFile, Ident; clas;
       Print #FastaFile, Row.Cells(1, 7)
       Print #FastaFile, Row.Cells(1, 2)
    Next
    Close #FastaFile
End Sub





Sub ResetAll_Click()

End Sub




Function BinAnd(a, b)
Attribute BinAnd.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BinAnd Makro
'
BinAnd = a And b
'
End Function

