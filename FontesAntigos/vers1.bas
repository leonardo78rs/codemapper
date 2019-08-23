Sub Principal()
Dim cFile As String
Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double

cPath = "C:\cvs\ccrp\05-construcao\V649\libsiret\"
cFile = "ccrpcb.prg"
cBusca = "CCRPPORT"

na = ImportTxtFile(cPath, cFile, cBusca, 1, 1)
nb = ImportTxtFile(cPath, cFile, "function", 2, 1)
nc = ExtraiNomesFunc()

' nc = MergeFunc(cFile)

cFile = "ccrpca.prg"

na = ImportTxtFile(cPath, cFile, cBusca, 1, na)
nb = ImportTxtFile(cPath, cFile, "function", 2, nb)
nc = ExtraiNomesFunc()

nc = MergeFunc(cFile)

nc = GroupFunc()

End Sub


Function ImportTxtFile(cPath As String, cFile As String, cBusca As String, nPlanilha As Integer, a As Double) As Double

Dim strTextLine
Dim strTextFile
Dim intFileNumber
Dim wPlan1 As Worksheet
Dim wPlan2 As Worksheet
Dim lin, col As Integer
Dim LinFonte As Double
Dim ncoment As Integer


If nPlanilha = 1 Then
   Set wPlan1 = Sheets("Sheet1")
Else
   Set wPlan1 = Sheets("Sheet2")
End If
 
lin = a
LinFonte = 0
intFileNumber = 1  'Criar numeração
strTextFile = cPath + cFile


Open strTextFile For Input As #intFileNumber 'Criar conexão com o arquivo txt

'Loop para percorrer as linhas do arquivo até o seu final
Do While Not EOF(intFileNumber)
   Line Input #intFileNumber, strTextLine
   LinFonte = LinFonte + 1
         
   If Mid(Trim(strTextLine), 1, 1) <> "/" Then
   
      ncoment = 0
      ncoment = InStr(1, strTextLine, "//", vbTextCompare)
      If ncoment > 0 Then
         strTextLine = Mid(strTextLine, 1, ncoment - 1)
      End If
    
      col = 0
      col = InStr(1, strTextLine, cBusca, vbTextCompare)
      
      If col <> 0 Then
         lin = lin + 1
         wPlan1.Cells(lin, 1) = cFile
         wPlan1.Cells(lin, 2) = LinFonte
         wPlan1.Cells(lin, 3) = col
         wPlan1.Cells(lin, 5) = strTextLine
      End If
   
   End If
     
Loop
ImportTxtFile = lin

'Fechar a conexão com o arquivo
Close #intFileNumber

End Function

Function ExtraiNomesFunc()

Dim wPlan1 As Worksheet
Dim wPlan2 As Worksheet
Dim lin1, lin2, col As Integer
Dim nIni, nFim As Double

Set wPlan1 = Sheets("Sheet1")
Set wPlan2 = Sheets("Sheet2")

For lin1 = 1 To 600
    nIni = 0
    nFim = 0
    nIni = InStr(1, wPlan2.Cells(lin1, 5), "function", vbTextCompare) + 8
    nFim = InStr(nIni, wPlan2.Cells(lin1, 5), "(", vbTextCompare)
   
    If (nFim - nIni) >= 0 Then
       wPlan2.Cells(lin1, 4) = Mid(wPlan2.Cells(lin1, 5), nIni, (nFim - nIni) + 1) + ")"
    End If
         
Next
ExtraiNomesFunc = True

End Function

Function MergeFunc(cFile As String)

Dim wPlan1 As Worksheet
Dim wPlan2 As Worksheet
Dim lin1, lin2, lin3, col As Integer
Dim nIni, nFim As Double

Set wPlan1 = Sheets("Sheet1")
Set wPlan2 = Sheets("Sheet2")
lin2 = 1
lin3 = 1 

For lin1 = 1 To 900
    lin2 = lin3 
    
    Do While TRUE 
        
        lin2 = lin2 + 1
        
        If wPlan2.Cells(lin2, 1) = "" OR (wPlan1.Cells(lin1, 1) <> wPlan2.Cells(lin2, 1)) Then
           Exit Do
        End If
    
        If wPlan1.Cells(lin1, 2) > wPlan2.Cells(lin2, 2) And _
           wPlan1.Cells(lin1, 2) < wPlan2.Cells(lin2 + 1, 2) Then           
           
           wPlan1.Cells(lin1, 4) = wPlan2.Cells(lin2, 4)           
        End If
        
        lin3 = lin2 
    Loop 'Next
    
    
Next

MergeFunc = True

End Function


Function GroupFunc()
Dim wPlan1 As Worksheet
Dim wPlan4 As Worksheet
Dim lin1, lin2, col As Integer
Dim nIni, nFim, nCountGroup As Double
Dim cFonte, cFunct As String

Set wPlan1 = Sheets("Sheet1")
Set wPlan4 = Sheets("Sheet4")
nCountGroup = 0

' Primeira
cFonte = wPlan1.Cells(2, 1)
cFunct = wPlan1.Cells(2, 4)
nCountGroup = 1
lin2 = 3

For lin1 = 3 To 900
    If wPlan1.Cells(lin1, 1) = cFonte And _
       wPlan1.Cells(lin1, 4) = cFunct Then
           
       cFonte = wPlan1.Cells(lin1, 1)
       cFunct = wPlan1.Cells(lin1, 4)
       nCountGroup = nCountGroup + 1
    Else
       wPlan4.Cells(lin2, 1) = cFonte
       wPlan4.Cells(lin2, 2) = cFunct
       wPlan4.Cells(lin2, 3) = nCountGroup
       lin2 = lin2 + 1
       
       
       cFonte = wPlan1.Cells(lin1, 1)
       cFunct = wPlan1.Cells(lin1, 4)
       nCountGroup = 1
    End If

Next

End Function
