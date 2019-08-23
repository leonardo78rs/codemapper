Sub Principal()
Dim cFile As String
Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double
Dim nfontes As Double
Dim wFontes As Worksheet
Dim wOcorr As Worksheet
Dim numfontes  As Integer
Dim wResumo As Worksheet
Dim lAnimado As Boolean
Dim cTimeIni As String
Dim cTimeFim As String
Dim nTime    As Double

Set wFontes = Sheets("Fontes")
Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")

cTimeIni = Time()
wOcorr.Cells(8, 9) = 0

cPath = wOcorr.Cells(3, 2)     'Ocorrencias!B3   -- Pasta que cont‚m os arquivos
cBusca = wOcorr.Cells(2, 2)    'Ocorrencias!B2   -- Palavra a buscar
lAnimado = wOcorr.Cells(6, 4)  'Indica se a tela ficar  mostrando o que est  acontecendo (demora mais)

numfontes = 0                  'quantidade de fontes pesquisados

If Not lAnimado Then
   wOcorr.Cells(6, 2) = ""
   wOcorr.Cells(7, 2) = numfontes
End If
   
na = 12   'retorno da outra funcao e linha inicial da planilha ocorrencias

LimpaOcorrencias

For nfontes = 1 To 800
    'wFontes.Cells(nfontes, 3) = ""
    'If wFontes.Cells(nfontes, 2) = "S" Then
       cFile = wFontes.Cells(nfontes, 1)
       numfontes = numfontes + 1
       
       If Len(Trim(cFile)) > 0 Then
       
          If lAnimado Then
             wOcorr.Cells(6, 2) = cFile ' Colocar em tela o fonte que est  buscando...
             wOcorr.Cells(7, 2) = numfontes
          End If
       
          na = ImportTxtFile(cPath, cFile, cBusca, na)
          'wFontes.Cells(nfontes, 3) = "Verif"
       End If
    'End If
Next

If Not lAnimado Then
  wOcorr.Cells(6, 2) = cFile ' Colocar em tela o fonte que est  buscando...
  wOcorr.Cells(7, 2) = numfontes
End If

LimpaResumo ("ABC")

' Faz o resumo das fun‡äes
GroupFunc (lAnimado)

LimpaResumo ("KLM")

cTimeFim = Time()

nTime = Round((Mid(cTimeFim, 1, 2) * 3600 + Mid(cTimeFim, 4, 2) * 60 + Mid(cTimeFim, 7, 2) - _
               Mid(cTimeIni, 1, 2) * 3600 - Mid(cTimeIni, 4, 2) * 60 - Mid(cTimeIni, 7, 2)), 2)
        
        
wOcorr.Cells(8, 9) = nTime



End Sub

'--------------------------------------------------------------------------
Sub Secundario()

Dim cFile As String
Dim cBusca As String
Dim cPath As String
Dim na As Double
Dim nb As Double
Dim nc As Double
Dim wFontes As Worksheet
Dim nfontes As Integer
Dim nLimpa  As Integer
Dim wResumo As Worksheet
Dim nBuscas As Integer
Dim nMaxBuscas As Integer
Dim wOcorr As Worksheet
Dim cTimeIni As String
Dim cTimeFim As String
Dim nTime    As Double

Set wFontes = Sheets("Fontes")
Set wResumo = Sheets("Resumo")
Set wOcorr = Sheets("Ocorrencias")



cTimeIni = Time()
wOcorr.Cells(9, 18) = cTimeIni
wOcorr.Cells(10, 15) = 0

cPath = wOcorr.Cells(3, 2)     'Ocorrencias!B3   -- Pasta que cont‚m os arquivos

'Copia de A:C para  K:M
For nLimpa = 3 To 200
    wResumo.Cells(nLimpa + 3, 11) = wResumo.Cells(nLimpa, 1)
    wResumo.Cells(nLimpa + 3, 12) = wResumo.Cells(nLimpa, 2)
    wResumo.Cells(nLimpa + 3, 13) = wResumo.Cells(nLimpa, 3)
Next


' cBusca = "F_CalcCorrMonetaria"
For nBuscas = 6 To 46

    LimpaOcorrencias
    
    wResumo.Activate
    LimpaResumo ("ABC")

    cBusca = wResumo.Cells(nBuscas, 12)  'Coluna L
    cBusca = Trim(cBusca)
    
    If cBusca <> Empty Then
       If Mid(cBusca, 1, LenB(cBusca) - 2) <> wOcorr.Cells(2, 2) Then 'Ocorrencias!B2
       
            cBusca = Mid(cBusca, 1, Len(cBusca) - 1)
            
            na = 12
            
            For nfontes = 1 To 800
                If Len(Trim(cFile)) > 0 Then
                    'wFontes.Cells(nfontes, 3) = ""
                    'If wFontes.Cells(nfontes, 2) = "S" Then
                       cFile = wFontes.Cells(nfontes, 1)
                       na = ImportTxtFile(cPath, cFile, cBusca, na)
                       'wFontes.Cells(nfontes, 3) = "Verif"
                    'End If
                End If
            Next
            
            ' Agrupa fun‡äes  (False - nÆo exibe anima‡Æo no secund rio)
            GroupFunc (False)
           
            nColIni = 15 'coluna à
            For nAchados = 3 To 45   'ver ultima
            
                If Trim(wResumo.Cells(nAchados, 1)) = Empty Then
                   Exit For
                End If
            
                If wResumo.Cells(nAchados, 2) <> wResumo.Cells(nBuscas, 12) And _
                   wResumo.Cells(nAchados, 2) <> Empty Then        'nao colocar a propria funcao
        
                   wResumo.Cells(nBuscas, nColIni) = wResumo.Cells(nAchados, 2)
                   nColIni = nColIni + 1
                
                End If
            
            
            Next
            
        End If
     End If

Next


cTimeFim = Time()
wOcorr.Cells(10, 18) = cTimeFim

nTime = Round((Mid(cTimeFim, 1, 2) * 3600 + Mid(cTimeFim, 4, 2) * 60 + Mid(cTimeFim, 7, 2) - _
               Mid(cTimeIni, 1, 2) * 3600 - Mid(cTimeIni, 4, 2) * 60 - Mid(cTimeIni, 7, 2)), 2)
        
 
wOcorr.Cells(10, 15) = nTime

End Sub

'--------------------------------------------------------------------------
Function ImportTxtFile(cPath As String, cFile As String, cBusca As String, LinOcorrencias As Double) As Double

Dim strTextLine
Dim strTextFile
Dim intFileNumber
Dim wOcorr As Worksheet
Dim wPlan2 As Worksheet
Dim lin, col, colFunc As Integer
Dim LinFonte As Double
Dim ncoment As Integer
Dim ncomentSql As Integer

Dim cFuncAtual As String
Dim nini, nfim As Integer
Dim colComent As Integer

Set wOcorr = Sheets("Ocorrencias")

 
lin = LinOcorrencias
LinFonte = 0
intFileNumber = 1  'Criar numera‡Æo
strTextFile = cPath + cFile
cFuncAtual = ""

Open strTextFile For Input As #intFileNumber 'Criar conexÆo com o arquivo txt

'Loop para percorrer as linhas do arquivo at‚ o seu final
Do While Not EOF(intFileNumber)
   Line Input #intFileNumber, strTextLine
   LinFonte = LinFonte + 1
   
   colComent = 0
   colComent = InStr(1, strTextLine, "/*", vbTextCompare)
   colComent2 = InStr(1, strTextLine, "*/", vbTextCompare)
   
   If colComent = 0 Or (colComent > 0 And colComent2 > 0) Then
   
        If Mid(Trim(strTextLine), 1, 1) <> "/" Then
        
           ncoment = 0
           ncoment = InStr(1, strTextLine, "//", vbTextCompare)
           If ncoment > 0 Then
              strTextLine = Mid(strTextLine, 1, ncoment - 1)         ' se encontrou //, corta tudo que vem depois
           Else
              ncomentSql = 0
              ncomentSql = InStr(1, strTextLine, "--", vbTextCompare)
              If ncoment > 0 Then
                 strTextLine = Mid(strTextLine, 1, ncomentSql - 1)   ' se encontrou --, corta tudo que vem depois
              End If
           End If
           
                   
           
           colFunc = 0
           colProc = 0
           
           colFunc = InStr(1, strTextLine, "function", vbTextCompare)
           colProc = InStr(1, strTextLine, "Procedure", vbTextCompare)
           If colFunc = 0 And colProc <> 0 Then
              colFunc = colProc
           End If
           
           If colFunc <> 0 Then
              nini = 0
              nfim = 0
              nini = InStr(1, strTextLine, "function", vbTextCompare) + 8
              If colProc <> 0 Then
                 nini = InStr(1, strTextLine, "Procedure", vbTextCompare) + 9
              End If
              
              nfim = InStr(nini, strTextLine, "(", vbTextCompare)
             
              If nini > 0 And (nfim - nini) >= 0 Then
                 cFuncAtual = Mid(strTextLine, nini, (nfim - nini) + 1) + ")"
              End If
           End If
               
           col = 0
           col = InStr(1, strTextLine, cBusca, vbTextCompare)
           
           If col <> 0 Then
              lin = lin + 1
              wOcorr.Cells(lin, 1) = cFile
              wOcorr.Cells(lin, 2) = LinFonte
              wOcorr.Cells(lin, 3) = col
              wOcorr.Cells(lin, 4) = cFuncAtual
              wOcorr.Cells(lin, 5) = strTextLine
           End If
        
        End If
     
     End If
     
Loop
ImportTxtFile = lin

'Fechar a conexÆo com o arquivo
Close #intFileNumber

End Function

'--------------------------------------------------------------------------
Function ExtraiNomesFunc()

'Dim wOcorr As Worksheet
Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim As Double

'Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")

For lin1 = 1 To 600    'criar variavel ou celula para total
    nini = 0
    nfim = 0
    nini = InStr(1, wResumo.Cells(lin1, 5), "function", vbTextCompare) + 8
    If nini = 0 Then
       nini = InStr(1, wResumo.Cells(lin1, 5), "procedure", vbTextCompare) + 9
    End If
    
    
    nfim = InStr(nini, wResumo.Cells(lin1, 5), "(", vbTextCompare)
   
    If (nfim - nini) >= 0 Then
       wResumo.Cells(lin1, 4) = Mid(wResumo.Cells(lin1, 5), nini, (nfim - nini) + 1) + ")"
    End If
         
Next
ExtraiNomesFunc = True

End Function

'--------------------------------------------------------------------------
'Function MergeFunc(cFile As String)
'
'Dim wOcorr As Worksheet
'Dim wResumo As Worksheet
'Dim lin1, lin2, lin3, col As Integer
'Dim nini, nfim As Double
'
'Set wOcorr = Sheets("Ocorrencias")
'Set wResumo = Sheets("Resumo")
'lin2 = 2
'lin3 = 3
'
'For lin1 = 13 To 900      'criar variavel ou celula para total
'    lin3 = lin2
'    Do While True
'
'        If wResumo.Cells(lin2, 1) = "" Then
'           lin2 = lin2 + 1
'           Exit Do
'        End If
'
'        If wOcorr.Cells(lin1, 1) <> wResumo.Cells(lin2, 1) Then
'           lin2 = lin2 + 1
'        End If
'
'        If wOcorr.Cells(lin1, 2) > wResumo.Cells(lin2, 2) And _
'           wOcorr.Cells(lin1, 2) < wResumo.Cells(lin2 + 1, 2) Then
'
'           wOcorr.Cells(lin1, 4) = wResumo.Cells(lin2, 4)
'           lin2 = lin3
'           Exit Do
'        End If
'
'
'    Loop 'Next
'
'
'Next
'
'MergeFunc = True
'
'End Function

'--------------------------------------------------------------------------
' GroupFunc ‚ uma fun‡Æo para agrupar e contar quantas vezes a fun‡Æo cont‚m a expressÆo procurada
' O resultado ‚ colocado na planilha Resumo

Sub GroupFunc(lAnimado As Boolean)

Dim wOcorr As Worksheet
Dim wResumo As Worksheet
Dim lin1, lin2, col As Integer
Dim nini, nfim, nCountGroup As Double
Dim cFonte, cFunct As String

Set wOcorr = Sheets("Ocorrencias")
Set wResumo = Sheets("Resumo")
nCountGroup = 0

' Primeira linha tem que pegar separado para fazer a compara‡Æo no FOR abaixo
cFonte = wOcorr.Cells(13, 1)
cFunct = wOcorr.Cells(13, 4)
nCountGroup = 1
lin2 = 3                            ' Linha do resumo que come‡a a gravar

If lAnimado Then
   wOcorr.Activate
End If

For lin1 = 14 To 500

    If wOcorr.Cells(lin1, 1) = cFonte And _
       wOcorr.Cells(lin1, 4) = cFunct Then
           
       cFonte = wOcorr.Cells(lin1, 1)
       cFunct = wOcorr.Cells(lin1, 4)
       nCountGroup = nCountGroup + 1
    Else
    
       wResumo.Cells(lin2, 1) = cFonte
       wResumo.Cells(lin2, 2) = cFunct
       wResumo.Cells(lin2, 3) = nCountGroup
       lin2 = lin2 + 1
              
       cFonte = wOcorr.Cells(lin1, 1)
       cFunct = wOcorr.Cells(lin1, 4)
       nCountGroup = 1
    End If

Next

If lAnimado Then
   wResumo.Activate
End If

End Sub

'--------------------------------------------------------------------------
Sub LimpaOcorrencias()

Dim nLimpa As Integer
Dim wOcorr As Worksheet

Set wOcorr = Sheets("Ocorrencias")

For nLimpa = 13 To 500
    wOcorr.Cells(nLimpa, 1) = ""
    wOcorr.Cells(nLimpa, 2) = ""
    wOcorr.Cells(nLimpa, 3) = ""
    wOcorr.Cells(nLimpa, 4) = ""
    wOcorr.Cells(nLimpa, 5) = ""
Next
    

End Sub


'--------------------------------------------------------------------------
' Fun‡Æo para limpar resumo
' Colunas ABC e KLM precisam ser limpas na funcao principal e na secund ria
' Mas em momentos diferentes.
' Na secund ria, vai limpar uma vez a KLM e vai limpar v rias vezes (loop) as colunas ABC.

Sub LimpaResumo(cColunasLimpar As String)

Dim nBuscas As Integer
Dim wResumo As Worksheet

Set wResumo = Sheets("Resumo")

increm = 0 ' Limpar colunas A,B,C ou colunas K,L,M


If cColunasLimpar = "KLM" Then
   increm = 10
   
   ' J  aproveita e limpar tudo da celula N para a direita
   wResumo.Activate
   Columns("N:N").Select
   Range(Selection, Selection.End(xlToRight)).Select
   Selection.ClearContents
   Columns("A:C").Select
   Sheets("Ocorrencias").Select
   
End If

'Colunas ABC, increm = 0
'Colunas KLM, increm = 10
For nBuscas = 3 To 200
    wResumo.Cells(nBuscas, increm + 1) = ""
    wResumo.Cells(nBuscas, increm + 2) = ""
    wResumo.Cells(nBuscas, increm + 3) = ""
Next
    

End Sub


Sub Instrucoes()

MsgBox ("- Tem que copiar todos arquivos para uma pasta" + Chr(10) + _
        "- Tem que colocar a lista de arquivos a serem analisador na planilha 'Fontes' " + Chr(10) + Chr(10) + _
        "- NÆo procura em linhas comentadas com // , e */ /* quando estiver na mesma linha" + Chr(10) + _
        "- Por‚m ainda nÆo desconsidera bloco de c¢digo comentado */ e /* em linhas diferentes " + Chr(10) + Chr(10) + _
        "- NÆo testei ainda PL/SQL, mas deve funcionar " + Chr(10) + Chr(10) + _
        "- O sistema nÆo entende est tico (STATIC FUNCTION), ou seja, acaba procurando em todos os arquivos fonte")
        
End Sub

