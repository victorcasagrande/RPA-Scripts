<#
.SYNOPSIS
    Coleta informaÃ§Ãµes de diagnÃ³stico de sistema, tela e sessÃ£o para anÃ¡lise de problemas
    de automaÃ§Ã£o de GUI (OCR, Captura de Imagem) em ambientes RDP.

.DESCRIPTION
    Este script foi projetado para investigar falhas em actions de OCR no Automation Anywhere
    executadas via RDP. Ele coleta dados crÃ­ticos sobre:
    
    1. ResoluÃ§Ã£o e Escala: A base para qualquer automaÃ§Ã£o visual.
    2. Drivers de VÃ­deo: Problemas de driver podem afetar a renderizaÃ§Ã£o.
    3. SessÃ£o RDP: Verifica se a sessÃ£o do bot estÃ¡ ativa e visÃ­vel.
    4. ConfiguraÃ§Ãµes de RenderizaÃ§Ã£o: Profundidade de cor e aceleraÃ§Ã£o de hardware.
    5. Contexto do Sistema: SO, arquitetura e permissÃµes do usuÃ¡rio.
    6. Desempenho: Identifica se a mÃ¡quina estÃ¡ sobrecarregada.
    
    O script gera um arquivo chamado 'ResolucaoTela.txt' no mesmo diretÃ³rio
    de sua execuÃ§Ã£o, substituindo qualquer versÃ£o anterior.
#>

# --- [ ETAPA 1: DEFINIÃ‡Ã•ES INICIAIS ] ---

# Obter o caminho do diretÃ³rio onde o script estÃ¡ sendo executado.
$caminhoScript = $PSScriptRoot

# Definir o nome do arquivo de log e construir o caminho completo.
$nomeArquivo = "ResolucaoTela.txt"
$caminhoCompletoArquivo = Join-Path -Path $caminhoScript -ChildPath $nomeArquivo

# Obter a data e a hora atuais para o registro no relatÃ³rio.
$timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

# Iniciar um array que armazenarÃ¡ todas as linhas do nosso relatÃ³rio.
$conteudoRelatorio = @(
    "=============================================================",
    "  RelatÃ³rio de DiagnÃ³stico para AutomaÃ§Ã£o de GUI (OCR/RDP)",
    "=============================================================",
    "Gerado em: $timestamp",
    ""
)


# --- [ ETAPA 2: RESOLUÃ‡ÃƒO DA TELA ] ---
# RelevÃ¢ncia: O OCR depende de uma resoluÃ§Ã£o estÃ¡vel e esperada.

$conteudoRelatorio += "--- [ 1. ResoluÃ§Ã£o da Tela Principal ] ---"
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $telaPrincipal = [System.Windows.Forms.Screen]::PrimaryScreen
    $resolucao = $telaPrincipal.Bounds
    
    $conteudoRelatorio += "Largura: $($resolucao.Width) pixels"
    $conteudoRelatorio += "Altura:  $($resolucao.Height) pixels"
}
catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel obter as informaÃ§Ãµes de resoluÃ§Ã£o."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 3: ESCALA (DPI) ] ---
# RelevÃ¢ncia: Uma escala diferente daquela em que o bot foi desenvolvido quebra o reconhecimento de imagem e OCR.

$conteudoRelatorio += "--- [ 2. Escala de Tela (DPI) ] ---"
try {
    $valorDPI = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI -ErrorAction Stop).AppliedDPI
    $escalaPercentual = [math]::Round(($valorDPI / 96) * 100)
    
    $conteudoRelatorio += "Valor DPI Aplicado: $valorDPI"
    $conteudoRelatorio += "Escala em Porcentagem: $($escalaPercentual)%"
}
catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel obter o valor de DPI do registro."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 4: DRIVER(S) DE VÃDEO ] ---
# RelevÃ¢ncia: Drivers desatualizados ou genÃ©ricos (comuns em VMs) podem causar problemas de renderizaÃ§Ã£o.

$conteudoRelatorio += "--- [ 3. Driver(s) de VÃ­deo ] ---"
try {
    $controladoresVideo = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
    if ($null -eq $controladoresVideo -or $controladoresVideo.Count -eq 0) {
        $conteudoRelatorio += "Nenhum dispositivo de vÃ­deo foi encontrado no sistema."
    } else {
        foreach ($controlador in $controladoresVideo) {
            $dataDriverFormatada = if ($controlador.DriverDate) { Get-Date -Date $controlador.DriverDate -Format "dd/MM/yyyy" } else { "NÃ£o disponÃ­vel" }
            $conteudoRelatorio += "Dispositivo: $($controlador.Name)"
            $conteudoRelatorio += "  - Fabricante: $($controlador.AdapterCompatibility)"
            $conteudoRelatorio += "  - VersÃ£o do Driver: $($controlador.DriverVersion)"
            $conteudoRelatorio += "  - Data do Driver: $dataDriverFormatada"
            $conteudoRelatorio += ""
        }
    }
}
catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel obter as informaÃ§Ãµes do driver de vÃ­deo."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}


# --- [ ETAPA 5: DETALHES DA SESSÃƒO DO USUÃRIO (RDP) ] --- # 
# RelevÃ¢ncia: Causa nÂº 1 de falhas de bot em RDP. O bot precisa de uma sessÃ£o grÃ¡fica ATIVA para "ver" a tela.
# Uma sessÃ£o 'Disc' (Desconectada) ou um computador bloqueado apresentarÃ£o uma tela preta/vazia para o bot.

$conteudoRelatorio += "--- [ 4. Status da SessÃ£o do UsuÃ¡rio (RDP) ] ---"
try {
    # MÃ‰TODO CORRIGIDO: Usa 'query session' para identificar a sessÃ£o atual (marcada com '>').
    # Este mÃ©todo Ã© mais confiÃ¡vel do que a variÃ¡vel de ambiente $env:SESSIONNAME, que pode ser herdada do processo pai.
    $queryOutput = query session
    $activeSessionLine = $queryOutput | Where-Object { $_ -like '>*' }

    # Verifica se o nome da sessÃ£o ativa contÃ©m "rdp"
    if ($activeSessionLine -match 'rdp') {
        $isRdp = "Sim"
    } else {
        $isRdp = "NÃ£o (SessÃ£o de Console Local)"
    }
    
    $conteudoRelatorio += "Executando em sessÃ£o RDP: $isRdp"

    # A variÃ¡vel $env:CLIENTNAME ainda Ã© Ãºtil para obter o nome do computador de origem em uma sessÃ£o RDP.
    if ($isRdp -eq "Sim" -and $env:CLIENTNAME) {
        $conteudoRelatorio += "Nome do Cliente (de onde conecta): $($env:CLIENTNAME)"
    }

    # Usa o comando 'quser' para verificar o status da sessÃ£o do usuÃ¡rio atual.
    $quserOutput = quser $env:USERNAME 2>$null
    if ($quserOutput) {
        $sessionInfo = $quserOutput | Select-Object -Skip 1
        # A coluna de estado pode variar de posiÃ§Ã£o, mas geralmente Ã© a 3Âª ou 4Âª.
        $sessionState = ($sessionInfo -split '\s+', 5)[3].Trim()
        $conteudoRelatorio += "Status da SessÃ£o Atual: $sessionState"
        if ($sessionState -ne "Active" -and $sessionState -ne "Ativo") {
             $conteudoRelatorio += "  -> ALERTA: A sessÃ£o nÃ£o estÃ¡ ativa! O OCR provavelmente falharÃ¡."
        }
    } else {
        $conteudoRelatorio += "Status da SessÃ£o Atual: NÃ£o foi possÃ­vel determinar via 'quser'."
    }
} catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel verificar os detalhes da sessÃ£o."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 6: CONFIGURAÃ‡Ã•ES DE VÃDEO ADICIONAIS ] ---
# RelevÃ¢ncia: ConfiguraÃ§Ãµes de baixa qualidade podem degradar a imagem a ponto de o OCR nÃ£o conseguir ler o texto.

$conteudoRelatorio += "--- [ 5. ConfiguraÃ§Ãµes de RenderizaÃ§Ã£o GrÃ¡fica ] ---"
try {
    $displayConfig = (Get-CimInstance -ClassName Win32_DisplayConfiguration | Select-Object -First 1)
    $conteudoRelatorio += "Profundidade de Cor: $($displayConfig.BitsPerPel) bits por pixel"
    if ($displayConfig.BitsPerPel -lt 24) {
        $conteudoRelatorio += "  -> ATENÃ‡ÃƒO: Profundidade de cor baixa. Pode impactar a qualidade da imagem para o OCR."
    }

    $dwmComposition = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows" -Name "DWMComposition" -ErrorAction SilentlyContinue
    if ($dwmComposition) {
        $statusAceleracao = if ($dwmComposition.DWMComposition -eq 1) { "Habilitada (DWM Ativo)" } else { "Desabilitada (DWM Inativo)" }
        $conteudoRelatorio += "AceleraÃ§Ã£o de Hardware (Desktop Window Manager): $statusAceleracao"
    } else {
        $conteudoRelatorio += "AceleraÃ§Ã£o de Hardware (Desktop Window Manager): NÃ£o foi possÃ­vel determinar (Comum no Windows Server)"
    }
} catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel obter configuraÃ§Ãµes de renderizaÃ§Ã£o."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 7: INFORMAÃ‡Ã•ES GERAIS DO SISTEMA ] ---
# RelevÃ¢ncia: Fornece contexto sobre o ambiente de execuÃ§Ã£o do bot.

$conteudoRelatorio += "--- [ 6. InformaÃ§Ãµes do Sistema ] ---"
try {
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $conteudoRelatorio += "Sistema Operacional: $($osInfo.Caption) ($($osInfo.OSArchitecture))"
    $conteudoRelatorio += "VersÃ£o do SO: $($osInfo.Version)"
    $conteudoRelatorio += "UsuÃ¡rio Executando o Script: $env:USERDOMAIN\$env:USERNAME"
    $conteudoRelatorio += "VersÃ£o do PowerShell: $($PSVersionTable.PSVersion.ToString())"
} catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel obter as informaÃ§Ãµes do sistema operacional."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 8: DESEMPENHO ATUAL DO SISTEMA ] ---
# RelevÃ¢ncia: Sobrecarga de CPU ou RAM causa lentidÃ£o e inconsistÃªncia na renderizaÃ§Ã£o dos elementos da tela.

$conteudoRelatorio += "--- [ 7. Desempenho Atual do Sistema (Snapshot) ] ---"
try {
    # Uso de CPU
    $cpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 1).CounterSamples.CookedValue
    $conteudoRelatorio += "Uso de CPU (Total): $([math]::Round($cpuUsage, 2))%"

    # Uso de MemÃ³ria RAM
    $memInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $totalRamGB = [math]::Round($memInfo.TotalVisibleMemorySize / 1MB / 1024, 2)
    $freeRamGB = [math]::Round($memInfo.FreePhysicalMemory / 1MB / 1024, 2)
    $usedRamGB = $totalRamGB - $freeRamGB
    $usedRamPercent = if ($totalRamGB -gt 0) { [math]::Round(($usedRamGB / $totalRamGB) * 100, 2) } else { 0 }
    
    $conteudoRelatorio += "MemÃ³ria RAM Total: $totalRamGB GB"
    $conteudoRelatorio += "MemÃ³ria RAM em Uso: $usedRamGB GB ($($usedRamPercent)%)"

    if ($usedRamPercent -gt 90 -or $cpuUsage -gt 90) {
        $conteudoRelatorio += "  -> ALERTA: Alto uso de recursos! Pode causar lentidÃ£o e falhas na automaÃ§Ã£o."
    }
} catch {
    $conteudoRelatorio += "ERRO: NÃ£o foi possÃ­vel coletar dados de desempenho."
    $conteudoRelatorio += "      Detalhes: $($_.Exception.Message)"
}
$conteudoRelatorio += ""


# --- [ ETAPA 9: SALVAR O RELATÃ“RIO NO ARQUIVO ] ---

try {
    $conteudoRelatorio | Out-File -FilePath $caminhoCompletoArquivo -Encoding UTF8 -Force
    Write-Host "RelatÃ³rio de diagnÃ³stico 'ResolucaoTela.txt' gerado com sucesso em '$caminhoScript'" -ForegroundColor Green
}
catch {
    Write-Host "ERRO CRÃTICO: NÃ£o foi possÃ­vel escrever no arquivo '$caminhoCompletoArquivo'." -ForegroundColor Red
    Write-Host "Detalhes do erro: $($_.Exception.Message)" -ForegroundColor Red
}

# --- Fim do Script ---