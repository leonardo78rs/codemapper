----- isto aqui é so exemplo
Sub lista_arquivos()
    Dim pasta As String
    Dim linha As Integer
    Dim arquivo As String
    Dim x As Integer
    linha = 7
        
    'Pega o caminho completo da pasta
    pasta = InputBox("Digite o caminho da pasta" + Chr(13) + Chr(13) + "É preciso terminar com a barra \ final", "LISTA ARQUIVOS", Cells(5, 6))
        
    Cells(5, 6) = pasta
    
    'Cabeçalho
    Cells(linha, 2) = "Nome do Arquivo"
    Range("A1:C1").Font.Bold = True
    linha = linha + 1

    'Lista o primeiro arquivo da pasta
    arquivo = Dir(pasta, 7)
    Cells(linha, 2) = arquivo
         
    'Lista os arquivos restantes
    Do While arquivo <> "" ' And linha <= Cells(10, 7)
        arquivo = Dir
        If arquivo <> "" Then ' Or linha > Cells(10, 7) Then
            linha = linha + 1
            Cells(linha, 2) = arquivo
        End If
    Loop
    Cells(7, 8) = linha
    
    Do While Cells(linha + 1, 2) <> ""
            linha = linha + 1
            Cells(linha, 2) = ""
    Loop
    
   
End Sub



Sub renomeia()

Dim pasta As String
Dim linha As Integer
pasta = InputBox("Confirme a pasta" + Chr(13) + Chr(13) + "É preciso terminar com a barra \ final", "RENOMEIA CONFORME JÁ TÁ APARECENDO ABAIXO", Cells(5, 6))
        
Cells(5, 6) = pasta
linha = Cells(7, 8)

    
For x = 8 To linha
    Name pasta + Cells(x, 2) As pasta + Cells(x, 3)
Next x
    

End Sub



