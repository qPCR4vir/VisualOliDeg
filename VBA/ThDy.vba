Attribute VB_Name = "ThDy"

Function Tm_min_basic_formula() As String
    Tm_min_basic_formula = "=64.9+41*(PrimerLen-Match!R12C-16.4)/PrimerLen"
End Function


Function Tm_max_basic_formula() As String
    Tm_max_basic_formula = "=64.9+41*(Match!R13C-16.4)/PrimerLen"
End Function




Sub Tm_Set_Change()

Dim Tmmin As Range
Set Tmmin = Range("Tm_min").Resize(, Range("NoNt")) '[Align.ColsAll].Rows(6)
Dim Tmmax As Range
Set Tmmax = Range("Tm_max").Resize(, Range("NoNt")) ' Range("Align.ColsAll").Rows(7)

   Select Case Range("Tm_choice")
   
   Case 1 ' Tm basic
   
      Tmmin.FormulaR1C1 = "=64.9+41*(PrimerLen-Match!R12C-16.4)/PrimerLen" ' Match!R12C=max count of AT elimina formula de Tm bas min en Match
      Tmmax.FormulaR1C1 = "=64.9+41*(Match!R13C-16.4)/PrimerLen" ' Match!R13C=max count of CG elimina formula de Tm bas max en Match
      '=64.9+41*(PrimerLen-Z(-2)S-16.4)/PrimerLen     ' Formula ori in Match for Tm bas min
      'Range("Tm_min").FormulaR1C1 = "=Match!R14C"    ' Elimina seleccion de Tm_choice: direct Match!R14C=Tm bas min
      'Range("Tm_max").FormulaR1C1 = "=Match!R15C"    ' Elimina seleccion de Tm_choice
      
   Case 2 ' G calculated by NN
   
      Tmmin.FormulaR1C1 = "=(Match!R17C-TaK*Match!R18C)/1000" ' R17C=dH,max  ;  R18C=dS,min
      Tmmax.FormulaR1C1 = "=(Match!R16C-TaK*Match!R19C)/1000" ' R16C=dH,min  ;  R19C=dS,max
      'Match!R4C ---  =(Z17S-TaK*Z18S)/1000
      'Range("Tm_min").FormulaR1C1 = "=Match!R4C"   '  R4C=Suma dGmax
      'Range("Tm_max").FormulaR1C1 = "=Match!R5C"   '  R5C=Suma dGmin
      
   Case 3 ' Tm calculated by NN
   
      Tmmin.FormulaR1C1 = "=1000*(Match!R20C-3.4)/(Match!R22C+RlnPC)+Kelv_Salt" 'R20C=Suma dHmin ; R22C=Suma dSmin
      Tmmax.FormulaR1C1 = "=1000*(Match!R21C-3.4)/(Match!R23C+RlnPC)+Kelv_Salt" 'R21C=Suma dHmax ; R23C=Suma dSmax
      'R24C==1000*(Z(-4)S-3.4)/(Z(-2)S+RlnPC)+Kelv_Salt  'R20C=Suma dHmin ; R22C=Suma dSmin
      'R25C==1000*(Z(-4)S-3.4)/(Z(-2)S+RlnPC)+Kelv_Salt 'R20C=Suma dHmin ; R22C=Suma dSmin
      'Range("Tm_min").FormulaR1C1 = "=Match!R24C" 'R24C =Tm,min NearNeig
      'Range("Tm_max").FormulaR1C1 = "=Match!R25C" 'R25C =Tm,max NearNeig
      
   Case 4 ' I calculated by NN??
   
      Tmmin.FormulaR1C1 = "= IF((Match!R17C-TaK*Match!R18C)/1000<G_sat,1,te*EXP((Match!R17C-TaK*Match!R18C)/1000*ro))"   'R4C=
      Tmmax.FormulaR1C1 = "= IF((Match!R16C-TaK*Match!R19C)/1000<G_sat,1,te*EXP((Match!R16C-TaK*Match!R19C)/1000*ro))"   'R4C=
      '= WENN(Z4S<G_sat;1;te*EXP(Z4S*ro))
      '= WENN(Z5S<G_sat;1;te*EXP(Z5S*ro))
      'Range("Tm_min").FormulaR1C1 = "=Match!R6C"  'R6C=Suma I min
      'Range("Tm_max").FormulaR1C1 = "=Match!R7C"  'R7C=Suma I max
      
   End Select
End Sub




Function PrimerSeq(ByVal PrimerPosition As Long, ByVal PrimerLen As Long) As String

    Dim primer As Range
    Set primer = Range("UsedCons").Columns(PrimerPosition).Resize(, PrimerLen)
    
    PrimerPosition = PrimerPosition - Range("SeqStart") + 1
    
    
    PrimerSeq = ""
    Dim base As Range
    For Each base In primer.Cells
        PrimerSeq = PrimerSeq + base
    Next
    
End Function


Function Tm_min_basic(ByVal PrimerPosition As Long, ByVal PrimerLen As Long) As Double
    
    Dim GC As Long
    Dim AT As Long
    
    GC = Range("Match.ColH_first").Cells(12, PrimerPosition - Range("SeqStart") + 1) 'Match.sumATmax
    AT = PrimerLen - GC
       
    Tm_min_basic = 64.9 + 41 * (AT - 16.4) / PrimerLen
End Function


Function Tm_max_basic() As Double
    Dim cMatchAll As Range
    Set cMatchAll = Range("Match.ColH_first")
    
    Dim PrimerLen As Range
    Set PrimerLen = Range("PrimerLen")
    
    Dim currPrimerPosition As Range
    Set currPrimerPosition = Range("currPrimerPosition")
    
    Dim SeqStart As Range
    Set SeqStart = Range("SeqStart")

    Tm_max_basic = 64.9 + 41 * (cMatchAll(13, currPrimerPosition - SeqStart + 1) - 16.4) / PrimerLen
End Function

Function Tm_min_NN() As Double
    Dim cMatchAll As Range
    Set cMatchAll = Range("Match.ColH_first")
    
    Dim RlnPC As Range
    Set RlnPC = Range("RlnPC")
    
    Dim Kelv_Salt As Range
    Set Kelv_Salt = Range("Kelv_Salt")
    
    Dim currPrimerPosition As Range
    Set currPrimerPosition = Range("currPrimerPosition")
    
    Dim SeqStart As Long
    SeqStart = Range("SeqStart")
    
    Dim Suma_dHmin As Double, Suma_dSmin As Double
    Suma_dHmin = cMatchAll(20, currPrimerPosition - SeqStart + 1)
    Suma_dSmin = cMatchAll(22, currPrimerPosition - SeqStart + 1)

    Tm_min_NN = 1000 * (Suma_dHmin - 3.4) / (Suma_dSmin + RlnPC) + Kelv_Salt
    
End Function


Function Tm_max_NN() As Double
    Dim cMatchAll As Range
    Set cMatchAll = Range("Match.ColH_first")
    
    Dim RlnPC As Range
    Set RlnPC = Range("RlnPC")
    
    Dim Kelv_Salt As Range
    Set Kelv_Salt = Range("Kelv_Salt")
    
    Dim currPrimerPosition As Range
    Set currPrimerPosition = Range("currPrimerPosition")
    
    Dim SeqStart As Long
    SeqStart = Range("SeqStart")
    
    Dim Suma_dHmax As Double, Suma_dSmax As Double
    'R21C=Suma dHmax ; R23C=Suma dSmax
    Suma_dHmax = cMatchAll(21, currPrimerPosition - SeqStart + 1)
    Suma_dSmax = cMatchAll(23, currPrimerPosition - SeqStart + 1)

    Tm_max_NN = 1000 * (Suma_dHmax - 3.4) / (Suma_dSmax + RlnPC) + Kelv_Salt
End Function


