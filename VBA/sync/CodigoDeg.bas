Attribute VB_Name = "CodigoDeg"
' Match.DataBeg =WENN(ODER(Align!M$10=FALSCH;Align!M22=Align!M$9;Align!M22="");"";WENN(Align!M22="-";3;WENN(BinAnd(FINDEN(Align!M22;CodDegElegir)-1          ;M$9);1;2)))
' Match.DataBeg =WENN(ODER(Align!M$10=FALSCH;Align!M22=Align!M$9;Align!M22="");"";WENN(Align!M22="-";3;WENN(BinAnd(INDEX(CodDeg!C;CODE(Align!M22));M$9);1;2)))

Dim deg_code(1 To 255) As Byte
 
'Call InitDegCod

'Sub Auto_Open()
'   Call InitDegCod
'End Sub


Function opt_A() As Boolean

    Debug.Print "A"
    opt_A = True
    

End Function

Function opt_B() As Boolean

    Debug.Print "B"
    opt_B = True

End Function


Public Function cod2txt(cod As Byte) As String
    
End Function

Sub InitDegCod()
' initialize all variable and array for convertions code / txt
    Dim i As Byte
    
    For i = 1 To 254
        deg_code(i) = Range("txt2code").Rows(i)
    Next

End Sub
