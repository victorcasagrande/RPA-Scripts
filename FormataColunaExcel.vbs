' =====================================================================
' SCRIPT: FormatarColunaE_Data.vbs
' DESCRIÇÃO: Formata a coluna E da planilha ativa no Excel para o 
'            formato de data "dd/mm/yyyy".
' AUTOR: Seu Nome/Empresa
' DATA: 23/05/2024
' =====================================================================

Option Explicit ' Força a declaração de todas as variáveis

Dim objExcel, objSheet, strSheetName

' Ativa o tratamento de erros. Se o Excel não estiver aberto, o script não irá travar.
On Error Resume Next

' Tenta se conectar a uma instância do Excel que já está em execução
Set objExcel = GetObject(, "Excel.Application")

' Verifica se a conexão com o Excel falhou (Err.Number diferente de 0)
If Err.Number <> 0 Then
    On Error GoTo 0 ' Desativa o tratamento de erros
    WScript.Echo "ERRO: Não foi possível encontrar uma instância do Excel em execução." & vbCrLf & "Por favor, abra a planilha desejada e execute o script novamente."
    WScript.Quit ' Encerra o script
End If

On Error GoTo 0 ' Desativa o tratamento de erros para que outros erros possam ser vistos

' Verifica se há uma planilha ativa
If objExcel.ActiveSheet Is Nothing Then
    WScript.Echo "ERRO: O Excel está aberto, mas não há uma planilha ativa." & vbCrLf & "Por favor, selecione uma planilha e tente novamente."
    Set objExcel = Nothing ' Libera o objeto
    WScript.Quit
End If

' Pega a referência para a planilha ativa
Set objSheet = objExcel.ActiveSheet
strSheetName = objSheet.Name

' === AÇÃO PRINCIPAL: Formata a coluna "E" ===
objSheet.Columns("E").NumberFormat = "dd/mm/yyyy"

' Exibe uma mensagem de sucesso para o usuário
WScript.Echo "A formatação da coluna E foi aplicada com sucesso na planilha '" & strSheetName & "'."

' Libera os objetos da memória para evitar que o Excel fique "preso"
Set objSheet = Nothing
Set objExcel = Nothing

WScript.Quit