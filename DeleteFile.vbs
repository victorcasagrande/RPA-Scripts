Function DeleteFile(strFilePath)
    Dim objFSO, objWMIService, colProcesses, objProcess
    
    ' Tenta encerrar processos do Excel se estiverem em execução
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    For Each objProcess In colProcesses
        objProcess.Terminate
    Next
    On Error GoTo 0
    
    ' Aguarda um pequeno tempo para garantir que o processo foi encerrado
    WScript.Sleep 2000
    
    ' Cria o objeto FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verifica se o arquivo existe antes de tentar deletá-lo
    If objFSO.FileExists(strFilePath) Then
        objFSO.DeleteFile strFilePath, True
        DeleteFile = "Arquivo deletado com sucesso: " & strFilePath
    Else
        DeleteFile = "Arquivo não encontrado: " & strFilePath
    End If
    
    Set objFSO = Nothing
End Function

' Exemplo de uso
Dim strFilePath
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Uso: cscript delete_file.vbs <caminho_do_arquivo>"
    WScript.Quit
End If

strFilePath = WScript.Arguments(0)
WScript.Echo DeleteFile(strFilePath)
