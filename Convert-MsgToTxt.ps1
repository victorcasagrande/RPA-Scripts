# SCRIPT: Convert-MsgToTxt.ps1
# DESCRIÇÃO: Busca '1.msg' ou '1.eml' e extrai seus dados para '1.txt'.
#            - Para .msg: Usa Outlook para extrair data e corpo do e-mail.
#            - Para .eml: Usa um parser nativo para extrair data, corpo e o primeiro link.

param(
    [Parameter(Mandatory=$false, HelpMessage="Forneça o caminho do diretório onde o arquivo '1.msg' ou '1.eml' está localizado.")]
    [string]$DirectoryPath
)

# --- FUNÇÕES AUXILIARES ---

function Show-PopupMessage {
    param([string]$Message, [string]$Title, [int]$Timeout = 3)
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.Popup($Message, $Timeout, $Title, 64)
    } catch { Write-Warning "Não foi possível exibir o popup: $Message" }
}

$script:logFilePath = $null

function Write-Log {
    param([string]$Message)
    if (-not $script:logFilePath) { return }
    $logEntry = "[{0}] - {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $script:logFilePath -Value $logEntry
}

function ConvertFrom-QuotedPrintableToUtf8 {
    param([string]$InputString)
    $joinedLines = $InputString -replace '=\r?\n', ''
    $memStream = New-Object System.IO.MemoryStream
    $matches = [regex]::Matches($joinedLines, '(=[A-Fa-f0-9]{2})|(.+?)')
    foreach ($match in $matches) {
        if ($match.Groups[1].Success) {
            $hex = $match.Value.Substring(1)
            $byte = [System.Convert]::ToByte($hex, 16)
            $memStream.WriteByte($byte)
        } else {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($match.Value)
            $memStream.Write($bytes, 0, $bytes.Length)
        }
    }
    return [System.Text.Encoding]::UTF8.GetString($memStream.ToArray())
}

function Parse-EmlFile {
    param([string]$FilePath)
    Write-Log "Iniciando parser nativo para arquivo .eml."
    $rawContent = Get-Content -Path $FilePath -Raw
    $dateMatch = [regex]::Match($rawContent, "(?im)^Date:\s*(.*)")
    $sentDateFormatted = "Data não disponível"
    if ($dateMatch.Success) {
        try {
            $dateString = $dateMatch.Groups[1].Value.Trim() -replace '\s\(.*\)'
            $sentDate = [datetime]::Parse($dateString)
            $sentDateFormatted = $sentDate.ToString("dd/MM/yyyy HH:mm:ss")
            Write-Log "Data extraída com sucesso: $sentDateFormatted"
        } catch { Write-Log "AVISO: Falha ao converter a string de data do EML." }
    } else { Write-Log "AVISO: Cabeçalho 'Date:' não foi encontrado." }

    $bodyText = ""
    $firstLink = ""
    $htmlMatch = [regex]::Match($rawContent, 'Content-Type: text/html;.*?\r?\n\r?\n(.*?)(?=\r?\n--|\Z)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($htmlMatch.Success) {
        $encodedHtml = $htmlMatch.Groups[1].Value
        $decodedHtml = ""
        if ($rawContent -match 'Content-Transfer-Encoding: quoted-printable') {
            $decodedHtml = ConvertFrom-QuotedPrintableToUtf8 -InputString $encodedHtml
        } else { $decodedHtml = $encodedHtml }
        try {
            $html = New-Object -ComObject "HTMLFile"
            $html.IHTMLDocument2_write($decodedHtml)
            $bodyText = $html.body.innerText
            $links = $html.getElementsByTagName('a')
            if ($links.length -gt 0) { $firstLink = $links[0].href }
        } catch { Write-Log "ERRO ao usar o parser HTML." }
    } else { Write-Log "AVISO: Não foi encontrada uma seção de corpo HTML no .eml." }
    return @{ Date = $sentDateFormatted; Body = $bodyText.Trim(); FirstLink = $firstLink }
}

# --- LÓGICA PRINCIPAL ---
$emailObject = $null
$outlook = $null
try {
    # 1. Setup inicial
    if ([string]::IsNullOrWhiteSpace($DirectoryPath)) { $targetDirectory = $PSScriptRoot } else { $targetDirectory = $DirectoryPath }
    $script:logFilePath = Join-Path -Path $targetDirectory -ChildPath "email_export_log.txt"
    Write-Log "--- Iniciando o script de exportação ---"; Write-Log "Usando diretório: '$targetDirectory'"
    if (-not (Test-Path -Path $targetDirectory -PathType Container)) { throw "O diretório '$targetDirectory' não existe." }
    
    $outputFilePath = Join-Path -Path $targetDirectory -ChildPath "1.txt"
    $msgPath = Join-Path -Path $targetDirectory -ChildPath "1.msg"
    $emlPath = Join-Path -Path $targetDirectory -ChildPath "1.eml"
    $emailData = $null
    
    # 2. LÓGICA HÍBRIDA
    if (Test-Path -Path $msgPath -PathType Leaf) {
        Write-Log "Arquivo '1.msg' encontrado. Usando automação do Outlook."
        $outlook = New-Object -ComObject Outlook.Application
        $emailObject = $outlook.CreateItemFromTemplate($msgPath)
        
        # Extrair a Data (lógica original para .msg)
        $sentDateFormatted = "Data não disponível"
        $sentDate = $emailObject.SentOn
        if ($sentDate.Year -lt 4000) { 
            $sentDateFormatted = $sentDate.ToString("dd/MM/yyyy HH:mm:ss")
        } else { 
            $sentDateFormatted = "[Data de Criação] " + $emailObject.CreationTime.ToString("dd/MM/yyyy HH:mm:ss")
        }
        
        # Extrair o Corpo (lógica original para .msg)
        $bodyText = $emailObject.Body
        if ([string]::IsNullOrWhiteSpace($bodyText)) {
            $htmlBody = $emailObject.HTMLBody
            if ($htmlBody) {
                $html = New-Object -ComObject "HTMLFile"
                $html.IHTMLDocument2_write($htmlBody)
                $bodyText = $html.body.innerText
            }
        }
        
        # Cria a estrutura de dados unificada (sem link para .msg)
        $emailData = @{ Date = $sentDateFormatted; Body = $bodyText; FirstLink = "" }

    } elseif (Test-Path -Path $emlPath -PathType Leaf) {
        Write-Log "Arquivo '1.eml' encontrado. Usando parser nativo."
        $emailData = Parse-EmlFile -FilePath $emlPath
        
    } else {
        throw "Nenhum arquivo '1.msg' ou '1.eml' foi encontrado no diretório."
    }

    # 3. Construir e salvar o conteúdo final (funciona para ambos os casos)
    $finalContent = "DATA DO E-MAIL: $($emailData.Date)`r`n"
    $finalContent += "----------------------------------------`r`n"
    $finalContent += "CONTEÚDO DO E-MAIL:`r`n`r`n"
    $finalContent += $emailData.Body

    # Esta seção só será adicionada se um link for encontrado (apenas para .eml)
    if (-not [string]::IsNullOrWhiteSpace($emailData.FirstLink)) {
        $finalContent += "`r`n`r`n----------------------------------------`r`n"
        $finalContent += "PRIMEIRO LINK ENCONTRADO:`r`n"
        $finalContent += $emailData.FirstLink
    }
    
    Set-Content -Path $outputFilePath -Value $finalContent -Encoding UTF8
    Write-Log "Dados extraídos e salvos com sucesso em '$outputFilePath'."
    Show-PopupMessage "Os dados do e-mail foram exportados com sucesso para `n$outputFilePath" "Operação Concluída"

} catch {
    $errorMessage = $_.Exception.Message
    if ($script:logFilePath) { Write-Log "ERRO: $errorMessage" } else { Write-Error "ERRO: $errorMessage" }
    Show-PopupMessage "Ocorreu um erro:`n$errorMessage" "Erro Crítico"
} finally {
    # Libera os objetos COM se eles foram criados
    if ($emailObject) {
        $emailObject.Close(1)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($emailObject) | Out-Null
    }
    if ($outlook) {
        Write-Log "Liberando objeto Outlook da memória."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    Write-Log "--- Script finalizado ---`n"
}