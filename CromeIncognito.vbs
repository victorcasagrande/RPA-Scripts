' Declaração de variáveis
Dim objShell
Dim chromePath

' Define o caminho completo para o executável do Google Chrome.
' As aspas duplas em volta do caminho são importantes caso ele contenha espaços.
' Este é o caminho mais comum em sistemas de 64 bits.
chromePath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""

' Cria o objeto WScript.Shell
Set objShell = CreateObject("WScript.Shell")

' Executa o comando usando o caminho completo + o parâmetro -incognito
objShell.Run chromePath & " -incognito"

' Libera o objeto da memória
Set objShell = Nothing