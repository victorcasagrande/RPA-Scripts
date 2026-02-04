'==========================================================================
' FUNÇÃO: ObterRelatorioDiscos
' DESCRIÇÃO: Coleta informações de discos locais (HD/SSD) e retorna uma
'            string formatada com o relatório. A função não exibe
'            nenhuma informação na tela, apenas retorna o valor.
' RETORNO: Uma string contendo o relatório completo dos discos.
'==========================================================================
Function ObterRelatorioDiscos()
    ' --- Constantes e Variáveis ---
    Const ToGB = 1073741824 ' 1024 * 1024 * 1024 (Bytes para Gigabytes)
    Const ToMB = 1048576    ' 1024 * 1024 (Bytes para Megabytes) - ADICIONADO

    Dim strComputer, objWMIService, colDisks, objDisk
    Dim nomeLimpo
    Dim capacidadeTotalGB, espacoLivreGB, espacoUtilizadoGB
    Dim capacidadeTotalMB, espacoLivreMB, espacoUtilizadoMB ' Variáveis para MB - ADICIONADO
    Dim porcentagemLivre
    Dim reportString ' String para construir o relatório
    reportString = ""

    ' --- Lógica WMI ---
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = 3")

    ' --- Processamento e Construção da String de Retorno ---
    If colDisks.Count > 0 Then
        For Each objDisk in colDisks
            nomeLimpo = Replace(objDisk.DeviceID, ":", "")
            
            ' --- Cálculos em GB (Gigabytes) ---
            capacidadeTotalGB = Round(objDisk.Size / ToGB, 1)
            espacoLivreGB = Round(objDisk.FreeSpace / ToGB, 1)
            espacoUtilizadoGB = Round((objDisk.Size - objDisk.FreeSpace) / ToGB, 1)

            ' --- Cálculos em MB (Megabytes) - ADICIONADO ---
            ' Usamos FormatNumber para formatar o número com separador de milhar e sem casas decimais
            capacidadeTotalMB = FormatNumber(objDisk.Size / ToMB, 0)
            espacoLivreMB = FormatNumber(objDisk.FreeSpace / ToMB, 0)
            espacoUtilizadoMB = FormatNumber((objDisk.Size - objDisk.FreeSpace) / ToMB, 0)
            
            If objDisk.Size > 0 Then
                porcentagemLivre = Round((objDisk.FreeSpace / objDisk.Size) * 100, 2)
            Else
                porcentagemLivre = 0
            End If

            ' --- Montagem da String de Relatório - ALTERADO ---
            reportString = reportString & "Disco: " & nomeLimpo & " - " & capacidadeTotalGB & " GB (" & capacidadeTotalMB & " MB)" & vbCrLf
            reportString = reportString & "Total utilizado: " & espacoUtilizadoGB & " GB (" & espacoUtilizadoMB & " MB)" & vbCrLf
            reportString = reportString & "Espaço livre: " & espacoLivreGB & " GB (" & espacoLivreMB & " MB)" & vbCrLf
            reportString = reportString & "Porcentagem livre: " & porcentagemLivre & " %" & vbCrLf
            reportString = reportString & vbCrLf ' Adiciona uma linha em branco para separar os discos
         
        Next
    Else
        reportString = "Nenhum disco local encontrado."
    End If

    ' --- Limpeza de Objetos ---
    Set objDisk = Nothing
    Set colDisks = Nothing
    Set objWMIService = Nothing

    ' --- RETORNO DA FUNÇÃO ---
    ' Atribui a string final ao nome da função para que ela seja retornada.
    ObterRelatorioDiscos = reportString
End Function


' --- Exemplo de como usar a função ---
' Dim relatorio
' relatorio = ObterRelatorioDiscos()
' WScript.Echo relatorio ' Exibe o relatório na tela
' ' Ou MsgBox relatorio