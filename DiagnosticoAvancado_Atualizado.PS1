#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnóstico de ambiente para automações RPA baseadas em interface gráfica (OCR, reconhecimento
    de imagem, sessões RDP).

.DESCRIPTION
    Coleta, em modo somente leitura, as informações que mais costumam explicar falhas de OCR e de
    captura de imagem em bots (ex.: Automation Anywhere) executados em sessões RDP ou console:

      - Contexto de execução: sessão 0, processo 32/64 bits, modo de linguagem, assinatura do script
      - Estado real da sessão (ativa/desconectada), desktop de entrada (tela bloqueada) e cliente RDP
      - Monitores, resolução física, DPI efetivo da sessão e profundidade de cor
      - Suavização de fontes (ClearType), composição DWM e políticas gráficas do RDP
      - Proteção de tela, bloqueio por inatividade e timeouts de sessão RDP
      - Drivers de vídeo, reinicialização pendente, idioma e layout de teclado
      - Desempenho (CPU, RAM, commit, disco), processos mais pesados e objetos GDI/USER
      - Serviços/processos do agente RPA e instâncias possivelmente órfãs (Excel, Word etc.)
      - Eventos recentes: TDR de vídeo, travamentos de aplicações, desconexões RDP, desligamentos
      - Teste real de captura de tela (detecta tela preta) sem gravar a imagem, salvo se solicitado

    Gera um relatório .txt (leitura humana) e um .json (consumo pelo bot ou monitoramento) com nome
    único por execução, em uma pasta com permissões restritas, e devolve um código de saída que o
    bot pode usar para decidir se continua ou aborta.

.PARAMETER PastaSaida
    Pasta dos relatórios. Padrão: %ProgramData%\DiagnosticoRPA (criada com acesso apenas para
    SYSTEM, Administradores e o usuário que executa).

.PARAMETER HorasEventos
    Janela, em horas, para a leitura dos logs de eventos. Padrão: 24.

.PARAMETER DiasRetencao
    Relatórios mais antigos que isso são apagados da pasta de saída. Padrão: 30.

.PARAMETER PularTesteCaptura
    Não executa o teste de captura de tela.

.PARAMETER SalvarCaptura
    Salva a captura de tela em PNG na pasta de saída. ATENÇÃO: a imagem pode conter dados
    sensíveis exibidos na tela. Por padrão, só estatísticas da imagem são registradas.

.PARAMETER Anonimizar
    Substitui nomes de usuário, computador e cliente RDP por um hash curto no relatório.

.PARAMETER PadroesServicoRpa
    Padrões (curingas) de nome de exibição dos serviços do agente RPA a verificar.

.PARAMETER PadroesProcessoRpa
    Padrões (curingas) de nome dos processos do agente RPA a verificar.

.EXAMPLE
    powershell.exe -NoProfile -NonInteractive -File "C:\RPA\DiagnosticoAvancado.ps1"

.EXAMPLE
    .\DiagnosticoAvancado.ps1 -HorasEventos 48 -Anonimizar -Verbose

.NOTES
    Versão 2.0.0
    Códigos de saída: 0 = OK | 1 = alertas de atenção | 2 = alertas críticos | 3 = falha do script.
    A última linha da saída padrão tem o formato:
        STATUS=<OK|ATENCAO|CRITICO>;CRITICOS=n;ATENCAO=n;RELATORIO=<caminho>;JSON=<caminho>
    Salve este arquivo como UTF-8 COM BOM para o Windows PowerShell 5.1 ler os acentos corretamente.
    Em produção, assine o script (Set-AuthenticodeSignature) e use a política AllSigned em vez de
    -ExecutionPolicy Bypass.
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$PastaSaida = (Join-Path -Path $env:ProgramData -ChildPath 'DiagnosticoRPA'),

    [ValidateRange(1, 720)]
    [int]$HorasEventos = 24,

    [ValidateRange(1, 365)]
    [int]$DiasRetencao = 30,

    [switch]$PularTesteCaptura,
    [switch]$SalvarCaptura,
    [switch]$Anonimizar,

    [string[]]$PadroesServicoRpa = @('*Automation Anywhere*'),
    [string[]]$PadroesProcessoRpa = @('*AutomationAnywhere*', '*BotAgent*')
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:VersaoScript     = '2.0.0'
$script:Achados          = New-Object 'System.Collections.Generic.List[object]'
$script:Dados            = [ordered]@{}
$script:NativoDisponivel = $false
$script:CaminhoScript    = $PSCommandPath
$script:IdSessao         = [System.Diagnostics.Process]::GetCurrentProcess().SessionId

#region ---------------------------- APIs nativas ----------------------------

$codigoNativo = @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class RpaDiagNativo
{
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);
    [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr hProcess, uint uiFlags);
    [DllImport("dwmapi.dll")] public static extern int DwmIsCompositionEnabled(out bool enabled);

    [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr hWnd);
    [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
    [DllImport("gdi32.dll")]  static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr OpenInputDesktop(uint dwFlags, bool fInherit, uint dwDesiredAccess);
    [DllImport("user32.dll", SetLastError = true)]
    static extern bool CloseDesktop(IntPtr hDesktop);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool GetUserObjectInformation(IntPtr hObj, int nIndex, StringBuilder pvInfo, int nLength, out int lpnLengthNeeded);

    [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "WTSQuerySessionInformationW")]
    static extern bool WTSQuerySessionInformation(IntPtr hServer, int sessionId, int infoClass, out IntPtr ppBuffer, out int pBytesReturned);
    [DllImport("wtsapi32.dll")]
    static extern void WTSFreeMemory(IntPtr pMemory);

    // GetDeviceCaps da tela inteira (LOGPIXELSX = 88, BITSPIXEL = 12)
    public static int ObterDeviceCap(int indice)
    {
        IntPtr hdc = GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero) return -1;
        try { return GetDeviceCaps(hdc, indice); }
        finally { ReleaseDC(IntPtr.Zero, hdc); }
    }

    // Nome do desktop que recebe a entrada do usuário. "Default" = desktop normal.
    // Retorna null quando não é acessível (tela bloqueada, UAC seguro, sem desktop interativo).
    public static string ObterDesktopDeEntrada()
    {
        IntPtr h = OpenInputDesktop(0, false, 0x0001); // DESKTOP_READOBJECTS
        if (h == IntPtr.Zero) return null;
        try
        {
            StringBuilder sb = new StringBuilder(256);
            int necessario;
            return GetUserObjectInformation(h, 2, sb, sb.Capacity * 2, out necessario) ? sb.ToString() : null; // UOI_NAME
        }
        finally { CloseDesktop(h); }
    }

    // WTSConnectState (classe 8)
    public static int ObterEstadoSessao(int sessionId)
    {
        IntPtr buffer; int bytes;
        if (!WTSQuerySessionInformation(IntPtr.Zero, sessionId, 8, out buffer, out bytes)) return -1;
        try { return Marshal.ReadInt32(buffer); }
        finally { WTSFreeMemory(buffer); }
    }

    // Informações textuais da sessão (6 = WinStationName, 10 = ClientName)
    public static string ObterTextoSessao(int sessionId, int infoClass)
    {
        IntPtr buffer; int bytes;
        if (!WTSQuerySessionInformation(IntPtr.Zero, sessionId, infoClass, out buffer, out bytes)) return null;
        try { return Marshal.PtrToStringUni(buffer); }
        finally { WTSFreeMemory(buffer); }
    }
}
'@

#endregion

#region ---------------------------- Funções auxiliares ----------------------------

function Add-Achado {
    param(
        [Parameter(Mandatory, Position = 0)][ValidateSet('INFO', 'ATENCAO', 'CRITICO')][string]$Severidade,
        [Parameter(Mandatory, Position = 1)][string]$Secao,
        [Parameter(Mandatory, Position = 2)][string]$Mensagem,
        [Parameter(Position = 3)][string]$Recomendacao = ''
    )
    $script:Achados.Add([pscustomobject]@{
        Severidade   = $Severidade
        Secao        = $Secao
        Mensagem     = $Mensagem
        Recomendacao = $Recomendacao
    })
}

# Executa uma seção isolada: uma falha não interrompe as demais e vira um achado.
function Invoke-Secao {
    param(
        [Parameter(Mandatory, Position = 0)][string]$Nome,
        [Parameter(Mandatory, Position = 1)][scriptblock]$Bloco
    )
    Write-Verbose "Coletando: $Nome"
    $cronometro = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $script:Dados[$Nome] = & $Bloco
    }
    catch {
        $script:Dados[$Nome] = [ordered]@{ Erro = $_.Exception.Message }
        Add-Achado ATENCAO $Nome "Falha na coleta desta seção: $($_.Exception.Message)"
    }
    Write-Verbose ('{0} concluída em {1} ms' -f $Nome, $cronometro.ElapsedMilliseconds)
}

function Get-ValorRegistro {
    param([string]$Caminho, [string]$Nome)
    try { (Get-ItemProperty -LiteralPath $Caminho -Name $Nome -ErrorAction Stop).$Nome }
    catch { $null }
}

function Protect-Valor {
    param([string]$Valor)
    if (-not $Anonimizar -or [string]::IsNullOrEmpty($Valor)) { return $Valor }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Valor.ToUpperInvariant())) }
    finally { $sha.Dispose() }
    'anon-' + (($hash[0..3] | ForEach-Object { $_.ToString('x2') }) -join '')
}

# Lê eventos sem confundir "nenhum evento" com "erro de acesso".
function Get-EventosSeguros {
    param([hashtable]$Filtro, [int]$Maximo = 300)
    try {
        $eventos = @(Get-WinEvent -FilterHashtable $Filtro -MaxEvents $Maximo -ErrorAction Stop)
        [ordered]@{ Erro = $null; Eventos = $eventos }
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') {
            [ordered]@{ Erro = $null; Eventos = @() }
        }
        else {
            [ordered]@{ Erro = $_.Exception.Message; Eventos = @() }
        }
    }
}

# Cria a pasta de saída e restringe o acesso (SIDs, para funcionar em Windows em qualquer idioma).
function Initialize-PastaSaida {
    param([string]$Caminho)
    if (-not (Test-Path -LiteralPath $Caminho)) {
        $null = New-Item -ItemType Directory -Path $Caminho -Force
    }
    try {
        # Lê e grava SOMENTE a seção de acesso (DACL). Um DirectorySecurity novo passado ao Set-Acl
        # tenta gravar também a SACL (auditoria), o que exige SeSecurityPrivilege.
        $pasta  = Get-Item -LiteralPath $Caminho
        $secoes = [System.Security.AccessControl.AccessControlSections]::Access
        if ($PSVersionTable.PSEdition -eq 'Core') {
            $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($pasta, $secoes)
        }
        else {
            $acl = $pasta.GetAccessControl($secoes)
        }

        # Remove a herança (descartando regras herdadas) e qualquer regra explícita existente.
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($regraExistente in @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
            [void]$acl.RemoveAccessRuleSpecific($regraExistente)
        }

        $heranca = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        $propagacao = [System.Security.AccessControl.PropagationFlags]::None
        $sids = @(
            (New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-18'),      # SYSTEM
            (New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-544'),  # Administradores
            [System.Security.Principal.WindowsIdentity]::GetCurrent().User             # usuário atual
        )
        foreach ($sid in $sids) {
            $regra = New-Object System.Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', $heranca, $propagacao, 'Allow')
            $acl.AddAccessRule($regra)
        }
        if ($PSVersionTable.PSEdition -eq 'Core') {
            [System.IO.FileSystemAclExtensions]::SetAccessControl($pasta, $acl)
        }
        else {
            $pasta.SetAccessControl($acl)
        }
    }
    catch {
        Write-Warning "Não foi possível restringir as permissões da pasta '$Caminho': $($_.Exception.Message)"
        Add-Achado ATENCAO 'Saida' "Permissões da pasta de relatórios não foram restringidas: $($_.Exception.Message)" 'Verifique quem tem acesso à pasta; os relatórios contêm dados do ambiente.'
    }
}

# Converte a estrutura coletada em linhas de texto indentadas.
function ConvertTo-LinhasTexto {
    param($Valor, [int]$Nivel = 0)
    $recuo = '  ' * $Nivel
    $linhas = New-Object 'System.Collections.Generic.List[string]'
    if ($Valor -is [System.Collections.IDictionary]) {
        foreach ($chave in $Valor.Keys) {
            $item = $Valor[$chave]
            if ($item -is [System.Collections.IDictionary] -or $item -is [array]) {
                $linhas.Add("$recuo${chave}:")
                foreach ($l in (ConvertTo-LinhasTexto -Valor $item -Nivel ($Nivel + 1))) { $linhas.Add($l) }
            }
            else {
                $linhas.Add("$recuo${chave}: $item")
            }
        }
    }
    elseif ($Valor -is [array]) {
        if ($Valor.Count -eq 0) { $linhas.Add("$recuo(nenhum)") }
        foreach ($elemento in $Valor) {
            if ($elemento -is [System.Collections.IDictionary]) {
                foreach ($l in (ConvertTo-LinhasTexto -Valor $elemento -Nivel $Nivel)) { $linhas.Add($l) }
                $linhas.Add('')
            }
            else {
                $linhas.Add("$recuo- $elemento")
            }
        }
    }
    else {
        $linhas.Add("$recuo$Valor")
    }
    $linhas
}

#endregion

try {
    #region ---------------------------- Preparação ----------------------------

    # A consciência de DPI precisa ser ativada ANTES de consultar telas; sem isso o Windows
    # "virtualiza" as coordenadas e a resolução é reportada já dividida pela escala.
    try {
        if (-not ('RpaDiagNativo' -as [type])) {
            Add-Type -TypeDefinition $codigoNativo -ErrorAction Stop
        }
        [void][RpaDiagNativo]::SetProcessDPIAware()
        $script:NativoDisponivel = $true
    }
    catch {
        Write-Verbose "APIs nativas indisponíveis: $($_.Exception.Message)"
    }

    Add-Type -AssemblyName System.Windows.Forms, System.Drawing

    Initialize-PastaSaida -Caminho $PastaSaida
    $carimbo  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $nomeBase = 'DiagRPA_{0}_{1}' -f (Protect-Valor $env:COMPUTERNAME), $carimbo

    #endregion

    #region ---------------------------- 1. Contexto de execução ----------------------------

    Invoke-Secao 'Contexto de Execucao' {
        $secao      = 'Contexto de Execucao'
        $identidade = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal  = New-Object System.Security.Principal.WindowsPrincipal($identidade)
        $modo       = $ExecutionContext.SessionState.LanguageMode.ToString()

        $assinatura = 'N/D (não executado a partir de um arquivo)'
        if ($script:CaminhoScript) {
            $assinatura = (Get-AuthenticodeSignature -LiteralPath $script:CaminhoScript).Status.ToString()
        }

        if ($script:IdSessao -eq 0 -or -not [Environment]::UserInteractive) {
            Add-Achado CRITICO $secao 'Execução na sessão 0 ou em contexto não interativo (serviço ou tarefa agendada sem logon).' 'OCR e captura de imagem exigem um desktop interativo. Execute o bot em uma sessão de usuário logada (console ou RDP).'
        }
        if (-not $script:NativoDisponivel) {
            Add-Achado ATENCAO $secao "APIs nativas indisponíveis (modo de linguagem: $modo). Estado da sessão, DPI efetivo e GDI não foram verificados." 'Se houver AppLocker/WDAC em modo restrito, assine o script e libere-o na política.'
        }
        if ($assinatura -ne 'Valid') {
            Add-Achado INFO $secao "Assinatura do script: $assinatura." 'Em produção, assine o script com certificado de code signing e use a política AllSigned em vez de Bypass.'
        }

        [ordered]@{
            Computador          = Protect-Valor $env:COMPUTERNAME
            Usuario             = Protect-Valor $identidade.Name
            Administrador       = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
            IdSessao            = $script:IdSessao
            ProcessoDe64Bits    = [Environment]::Is64BitProcess
            SistemaDe64Bits     = [Environment]::Is64BitOperatingSystem
            VersaoPowerShell    = $PSVersionTable.PSVersion.ToString()
            ModoLinguagem       = $modo
            PoliticaExecucao    = (Get-ExecutionPolicy).ToString()
            AssinaturaScript    = $assinatura
            ContextoInterativo  = [Environment]::UserInteractive
        }
    }

    #endregion

    #region ---------------------------- 2. Sessão e desktop ----------------------------

    Invoke-Secao 'Sessao e Desktop' {
        $secao   = 'Sessao e Desktop'
        $remota  = [System.Windows.Forms.SystemInformation]::TerminalServerSession
        $retorno = [ordered]@{ SessaoRemota = $remota }

        if ($script:NativoDisponivel) {
            $estados = @{ 0 = 'Active'; 1 = 'Connected'; 2 = 'ConnectQuery'; 3 = 'Shadow'; 4 = 'Disconnected';
                          5 = 'Idle'; 6 = 'Listen'; 7 = 'Reset'; 8 = 'Down'; 9 = 'Init' }
            $codigo  = [RpaDiagNativo]::ObterEstadoSessao($script:IdSessao)
            $estado  = if ($estados.ContainsKey($codigo)) { $estados[$codigo] } else { "Desconhecido ($codigo)" }
            $desktop = [RpaDiagNativo]::ObterDesktopDeEntrada()

            $retorno.EstadoSessao     = $estado
            $retorno.Estacao          = [RpaDiagNativo]::ObterTextoSessao($script:IdSessao, 6)
            $retorno.ClienteRdp       = Protect-Valor ([RpaDiagNativo]::ObterTextoSessao($script:IdSessao, 10))
            $retorno.DesktopDeEntrada = if ($desktop) { $desktop } else { 'Inacessível' }

            if ($estado -ne 'Active') {
                Add-Achado CRITICO $secao "A sessão do bot não está ativa (estado: $estado)." 'Sessões desconectadas não renderizam a tela. Use o mecanismo de auto-login do agente RPA ou mantenha a sessão conectada.'
            }
            if ($desktop -ne 'Default') {
                Add-Achado CRITICO $secao "O desktop de entrada não é o desktop normal ('$($retorno.DesktopDeEntrada)'): tela bloqueada, prompt seguro do UAC ou proteção de tela." 'Desbloqueie a sessão e revise as políticas de bloqueio (ver seção Bloqueio e Timeouts).'
            }
        }
        else {
            $logonUi = @(Get-Process -Name LogonUI -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $script:IdSessao })
            $retorno.LogonUIAtivo = ($logonUi.Count -gt 0)
            if ($logonUi.Count -gt 0) {
                Add-Achado CRITICO $secao 'LogonUI.exe em execução na sessão do bot: a estação está bloqueada.' 'Desbloqueie a sessão e revise as políticas de bloqueio.'
            }
        }

        if ($remota) {
            Add-Achado INFO $secao 'Sessão RDP: se a janela do cliente RDP for minimizada, o Windows pode parar de renderizar a tela remota.' 'No computador CLIENTE, crie o DWORD HKCU\Software\Microsoft\Terminal Server Client\RemoteDesktop_SuppressWhenMinimized = 2.'
        }
        $retorno
    }

    #endregion

    #region ---------------------------- 3. Telas e escala ----------------------------

    Invoke-Secao 'Telas e Escala' {
        $secao = 'Telas e Escala'
        $telas = @([System.Windows.Forms.Screen]::AllScreens)
        $lista = @(foreach ($t in $telas) {
            [ordered]@{
                Dispositivo  = $t.DeviceName
                Principal    = $t.Primary
                Resolucao    = '{0}x{1}' -f $t.Bounds.Width, $t.Bounds.Height
                Posicao      = '{0},{1}' -f $t.Bounds.X, $t.Bounds.Y
                AreaUtil     = '{0}x{1}' -f $t.WorkingArea.Width, $t.WorkingArea.Height
                BitsPorPixel = $t.BitsPerPixel
            }
        })

        $dpiRegistro = Get-ValorRegistro 'HKCU:\Control Panel\Desktop\WindowMetrics' 'AppliedDPI'
        $dpiEfetivo  = $null
        $bpp         = $null
        if ($script:NativoDisponivel) {
            $dpiEfetivo = [RpaDiagNativo]::ObterDeviceCap(88)
            $bpp        = [RpaDiagNativo]::ObterDeviceCap(12)
        }

        $principal = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        if ($principal.Width -lt 1280 -or $principal.Height -lt 720) {
            Add-Achado ATENCAO $secao ('Resolução principal baixa ({0}x{1}). Comum quando a sessão RDP é criada sem resolução fixa (ex.: reconexão pelo console).' -f $principal.Width, $principal.Height) 'Fixe a resolução da sessão do bot igual à usada no desenvolvimento (configuração do agente RPA ou do arquivo .rdp).'
        }
        if ($dpiEfetivo -and $dpiEfetivo -gt 0 -and $dpiEfetivo -ne 96) {
            Add-Achado ATENCAO $secao ('Escala da sessão em {0}% (DPI {1}).' -f [math]::Round($dpiEfetivo / 96 * 100), $dpiEfetivo) 'Confirme que é a mesma escala da máquina onde o bot foi desenvolvido; escala diferente quebra OCR e reconhecimento de imagem.'
        }
        if ($dpiRegistro -and $dpiEfetivo -and $dpiEfetivo -gt 0 -and [int]$dpiRegistro -ne $dpiEfetivo) {
            Add-Achado ATENCAO $secao "DPI do registro ($dpiRegistro) difere do DPI efetivo da sessão ($dpiEfetivo)." 'Comum em RDP com escala herdada do cliente. Um logoff/logon aplica a configuração de forma consistente.'
        }
        if ($bpp -and $bpp -gt 0 -and $bpp -lt 24) {
            Add-Achado ATENCAO $secao "Profundidade de cor baixa ($bpp bits)." 'Use 24 ou 32 bits na conexão RDP; cores reduzidas degradam OCR e comparação de imagens.'
        }
        if ($telas.Count -gt 1) {
            Add-Achado INFO $secao "$($telas.Count) monitores detectados." 'Coordenadas e capturas podem referir-se a outro monitor; prefira um único monitor na máquina do bot.'
        }

        [ordered]@{
            QuantidadeTelas = $telas.Count
            Telas           = $lista
            DpiEfetivo      = if ($dpiEfetivo -and $dpiEfetivo -gt 0) { $dpiEfetivo } else { 'N/D' }
            EscalaEfetiva   = if ($dpiEfetivo -and $dpiEfetivo -gt 0) { '{0}%' -f [math]::Round($dpiEfetivo / 96 * 100) } else { 'N/D' }
            DpiRegistro     = if ($dpiRegistro) { $dpiRegistro } else { 'N/D' }
            ProfundidadeCor = if ($bpp -and $bpp -gt 0) { "$bpp bits" } else { 'N/D' }
        }
    }

    #endregion

    #region ---------------------------- 4. Renderização ----------------------------

    Invoke-Secao 'Renderizacao' {
        $secao   = 'Renderizacao'
        $suave   = [System.Windows.Forms.SystemInformation]::IsFontSmoothingEnabled
        $tipoCod = [System.Windows.Forms.SystemInformation]::FontSmoothingType
        $tipo    = switch ($tipoCod) { 1 { 'Padrão (escala de cinza)' } 2 { 'ClearType' } default { "Desconhecido ($tipoCod)" } }

        $dwm = 'N/D'
        if ($script:NativoDisponivel) {
            $habilitado = $false
            if ([RpaDiagNativo]::DwmIsCompositionEnabled([ref]$habilitado) -eq 0) { $dwm = $habilitado }
        }

        $politicaTs = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
        $gpuRdp     = Get-ValorRegistro $politicaTs 'bEnumerateHWBeforeSW'
        $avc444     = Get-ValorRegistro $politicaTs 'AVC444ModePreferred'

        if (-not $suave) {
            Add-Achado ATENCAO $secao 'Suavização de fontes desativada nesta sessão.' 'O texto é desenhado de forma diferente de uma sessão com ClearType. Padronize entre desenvolvimento e produção (no cliente RDP: opção "Suavização de fonte").'
        }

        [ordered]@{
            SuavizacaoFontes    = $suave
            TipoSuavizacao      = $tipo
            ComposicaoDwm       = $dwm
            GpuNoRdp_Politica   = if ($null -ne $gpuRdp) { $gpuRdp } else { 'Não configurada' }
            Avc444_Politica     = if ($null -ne $avc444) { $avc444 } else { 'Não configurada' }
        }
    }

    #endregion

    #region ---------------------------- 5. Bloqueio e timeouts ----------------------------

    Invoke-Secao 'Bloqueio e Timeouts' {
        $secao      = 'Bloqueio e Timeouts'
        $politicaTs = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
        $rdpTcp     = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'

        $timeouts = [ordered]@{}
        foreach ($nome in 'MaxIdleTime', 'MaxDisconnectionTime', 'MaxConnectionTime') {
            $valor  = Get-ValorRegistro $politicaTs $nome
            $origem = 'GPO'
            if ($null -eq $valor) { $valor = Get-ValorRegistro $rdpTcp $nome; $origem = 'RDP-Tcp' }
            if ($null -eq $valor -or [int64]$valor -eq 0) {
                $timeouts[$nome] = 'Sem limite'
            }
            else {
                $minutos = [math]::Round([int64]$valor / 60000)
                $timeouts[$nome] = '{0} min ({1})' -f $minutos, $origem
                if ($nome -eq 'MaxIdleTime') {
                    Add-Achado ATENCAO $secao "Sessões RDP ociosas são desconectadas após $minutos min." 'Bots que só leem a tela não geram entrada e podem perder a sessão. Avalie remover o limite para a conta do bot.'
                }
            }
        }

        $chaveUsuario  = 'HKCU:\Control Panel\Desktop'
        $chavePolitica = 'HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop'
        $lerProtecao = {
            param($nomeValor)
            $v = Get-ValorRegistro $chavePolitica $nomeValor
            if ($null -eq $v) { $v = Get-ValorRegistro $chaveUsuario $nomeValor }
            $v
        }
        $protecaoAtiva  = (& $lerProtecao 'ScreenSaveActive') -eq '1'
        $protecaoSenha  = (& $lerProtecao 'ScreenSaverIsSecure') -eq '1'
        $protecaoTempo  = & $lerProtecao 'ScreenSaveTimeOut'
        $inatividadeMaq = Get-ValorRegistro 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'InactivityTimeoutSecs'

        if ($protecaoAtiva) {
            $sev = if ($protecaoSenha) { 'CRITICO' } else { 'ATENCAO' }
            Add-Achado $sev $secao ("Proteção de tela ativa após {0} s{1}." -f $protecaoTempo, $(if ($protecaoSenha) { ', com bloqueio por senha' } else { '' })) 'A proteção de tela cobre a área que o OCR lê. Desative-a para a conta do bot (via GPO, se for imposta).'
        }
        if ($inatividadeMaq -and [int]$inatividadeMaq -gt 0) {
            Add-Achado CRITICO $secao "Política de bloqueio por inatividade da máquina: $inatividadeMaq s." 'A estação será bloqueada durante execuções longas. Solicite exceção da política para as máquinas de bot.'
        }

        [ordered]@{
            TimeoutsRdp             = $timeouts
            ProtecaoTelaAtiva       = $protecaoAtiva
            ProtecaoTelaComSenha    = $protecaoSenha
            ProtecaoTelaTempoSeg    = if ($protecaoTempo) { $protecaoTempo } else { 'N/D' }
            BloqueioInatividadeSeg  = if ($inatividadeMaq) { $inatividadeMaq } else { 'Não configurado' }
        }
    }

    #endregion

    #region ---------------------------- 6. Drivers de vídeo ----------------------------

    Invoke-Secao 'Drivers de Video' {
        $secao        = 'Drivers de Video'
        $controlador  = @(Get-CimInstance -ClassName Win32_VideoController)
        if ($controlador.Count -eq 0) {
            Add-Achado ATENCAO $secao 'Nenhum adaptador de vídeo encontrado.'
        }
        $lista = @(foreach ($c in $controlador) {
            if ($c.ConfigManagerErrorCode -and $c.ConfigManagerErrorCode -ne 0) {
                Add-Achado ATENCAO $secao "Adaptador '$($c.Name)' com erro no Gerenciador de Dispositivos (código $($c.ConfigManagerErrorCode))." 'Reinstale ou atualize o driver; erros de driver causam falhas intermitentes de renderização.'
            }
            [ordered]@{
                Nome         = $c.Name
                Fabricante   = $c.AdapterCompatibility
                VersaoDriver = $c.DriverVersion
                DataDriver   = if ($c.DriverDate) { $c.DriverDate.ToString('dd/MM/yyyy') } else { 'N/D' }
                Status       = $c.Status
            }
        })
        [ordered]@{ Adaptadores = $lista }
    }

    #endregion

    #region ---------------------------- 7. Sistema ----------------------------

    Invoke-Secao 'Sistema' {
        $secao  = 'Sistema'
        $so     = Get-CimInstance -ClassName Win32_OperatingSystem
        $ubr    = Get-ValorRegistro 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'UBR'
        $ligado = (Get-Date) - $so.LastBootUpTime

        $reinicio = New-Object 'System.Collections.Generic.List[string]'
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reinicio.Add('Component Based Servicing') }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reinicio.Add('Windows Update') }
        if (Get-ValorRegistro 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations') { $reinicio.Add('Renomeação de arquivos pendente') }

        if ($reinicio.Count -gt 0) {
            Add-Achado ATENCAO $secao "Reinicialização pendente: $($reinicio -join ', ')." 'Atualizações pendentes podem gerar janelas e prompts inesperados. Reinicie fora da janela de execução dos bots.'
        }
        if ($ligado.TotalDays -gt 30) {
            Add-Achado INFO $secao ('Máquina ligada há {0} dias.' -f [math]::Floor($ligado.TotalDays)) 'Reinicializações periódicas reduzem vazamentos de memória e de objetos GDI.'
        }

        $teclado = try { [System.Windows.Forms.InputLanguage]::CurrentInputLanguage.Culture.Name } catch { 'N/D' }

        [ordered]@{
            SistemaOperacional      = '{0} ({1})' -f $so.Caption, $so.OSArchitecture
            Versao                  = if ($ubr) { '{0}.{1}' -f $so.Version, $ubr } else { $so.Version }
            UltimaInicializacao     = $so.LastBootUpTime.ToString('dd/MM/yyyy HH:mm:ss')
            TempoLigado             = '{0}d {1}h {2}min' -f $ligado.Days, $ligado.Hours, $ligado.Minutes
            CulturaRegional         = (Get-Culture).Name
            IdiomaInterface         = (Get-UICulture).Name
            LayoutTeclado           = $teclado
            FusoHorario             = [System.TimeZoneInfo]::Local.Id
            ReinicializacaoPendente = if ($reinicio.Count -gt 0) { $reinicio -join ', ' } else { 'Não' }
        }
    }

    #endregion

    #region ---------------------------- 8. Desempenho ----------------------------

    Invoke-Secao 'Desempenho' {
        $secao = 'Desempenho'

        # Classes WMI de desempenho têm nome fixo; contadores do Get-Counter são traduzidos
        # (ex.: "\Processador(_Total)\% Tempo de Processador" em pt-BR) e falham fora do inglês.
        $amostras = New-Object 'System.Collections.Generic.List[double]'
        for ($i = 0; $i -lt 4; $i++) {
            $valor = (Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
            if ($i -gt 0) { $amostras.Add([double]$valor) }   # a 1ª leitura costuma vir zerada
            Start-Sleep -Milliseconds 750
        }
        $cpu = [math]::Round(($amostras | Measure-Object -Average).Average, 1)

        $so       = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalGB  = [math]::Round($so.TotalVisibleMemorySize / 1MB, 2)   # valor em KB -> GB
        $livreGB  = [math]::Round($so.FreePhysicalMemory / 1MB, 2)
        $usoRam   = [math]::Round((1 - $so.FreePhysicalMemory / $so.TotalVisibleMemorySize) * 100, 1)
        $usoCommit = [math]::Round((1 - $so.FreeVirtualMemory / $so.TotalVirtualMemorySize) * 100, 1)

        if ($cpu -ge 90)       { Add-Achado ATENCAO $secao "CPU em $cpu% (média de 3 amostras)." 'Sobrecarga atrasa a renderização; aumente os tempos de espera do bot ou investigue os processos listados.' }
        if ($usoRam -ge 90)    { Add-Achado ATENCAO $secao "Memória física em $usoRam% de uso." 'Verifique instâncias órfãs (Excel, navegadores) deixadas por execuções anteriores.' }
        if ($usoCommit -ge 90) { Add-Achado ATENCAO $secao "Memória comprometida (commit) em $usoCommit%." 'Próximo do limite, aplicações falham ao abrir janelas. Revise o arquivo de paginação e os processos.' }

        $discos = @(foreach ($d in @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3')) {
            if (-not $d.Size) { continue }
            $livrePct = [math]::Round($d.FreeSpace / $d.Size * 100, 1)
            $livreDisco = [math]::Round($d.FreeSpace / 1GB, 1)
            if ($livrePct -lt 10 -or $livreDisco -lt 5) {
                Add-Achado ATENCAO $secao "Pouco espaço livre em $($d.DeviceID): $livreDisco GB ($livrePct%)." 'Logs, capturas e arquivos temporários do bot podem falhar ao gravar.'
            }
            '{0} {1} GB livres de {2} GB ({3}%)' -f $d.DeviceID, $livreDisco, [math]::Round($d.Size / 1GB, 1), $livrePct
        })

        $nucleos = [Environment]::ProcessorCount
        $null = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process   # aquecimento
        Start-Sleep -Seconds 1
        $topCpu = @(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process |
            Where-Object { $_.Name -notin '_Total', 'Idle' } |
            Sort-Object -Property PercentProcessorTime -Descending |
            Select-Object -First 5 |
            ForEach-Object { '{0} (PID {1}): {2}%' -f $_.Name, $_.IDProcess, [math]::Round($_.PercentProcessorTime / $nucleos, 1) })

        $topMemoria = @(Get-Process |
            Sort-Object -Property WorkingSet64 -Descending |
            Select-Object -First 5 |
            ForEach-Object { '{0} (PID {1}): {2} MB' -f $_.ProcessName, $_.Id, [math]::Round($_.WorkingSet64 / 1MB) })

        [ordered]@{
            CpuMedia        = "$cpu%"
            RamTotal        = "$totalGB GB"
            RamLivre        = "$livreGB GB"
            RamEmUso        = "$usoRam%"
            CommitEmUso     = "$usoCommit%"
            Discos          = $discos
            TopCpu          = $topCpu
            TopMemoria      = $topMemoria
        }
    }

    #endregion

    #region ---------------------------- 9. Recursos de GUI (GDI/USER) ----------------------------

    if ($script:NativoDisponivel) {
        Invoke-Secao 'Recursos GUI' {
            $secao = 'Recursos GUI'
            $quota = Get-ValorRegistro 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' 'GDIProcessHandleQuota'
            $limite = if ($quota) { [int]$quota } else { 10000 }

            $medidas = @(foreach ($p in @(Get-Process | Where-Object { $_.SessionId -eq $script:IdSessao })) {
                try {
                    [pscustomobject]@{
                        Nome = $p.ProcessName
                        PID  = $p.Id
                        GDI  = [int][RpaDiagNativo]::GetGuiResources($p.Handle, 0)
                        USER = [int][RpaDiagNativo]::GetGuiResources($p.Handle, 1)
                    }
                }
                catch { }   # processos sem permissão de consulta são ignorados
            })
            $top = @($medidas | Sort-Object -Property GDI -Descending | Select-Object -First 5)

            foreach ($m in $top) {
                if ($m.GDI -ge $limite * 0.8) {
                    Add-Achado CRITICO $secao "$($m.Nome) (PID $($m.PID)) usa $($m.GDI) objetos GDI (limite $limite)." 'Perto do limite o processo deixa de desenhar janelas e textos. Reinicie o processo e verifique vazamento.'
                }
                elseif ($m.GDI -ge $limite * 0.5) {
                    Add-Achado ATENCAO $secao "$($m.Nome) (PID $($m.PID)) usa $($m.GDI) objetos GDI." 'Acompanhe a evolução; crescimento contínuo indica vazamento.'
                }
            }

            [ordered]@{
                LimiteGdiPorProcesso = $limite
                TopGdi = @($top | ForEach-Object { '{0} (PID {1}): GDI={2} USER={3}' -f $_.Nome, $_.PID, $_.GDI, $_.USER })
            }
        }
    }

    #endregion

    #region ---------------------------- 10. Agente RPA ----------------------------

    Invoke-Secao 'Agente RPA' {
        $secao = 'Agente RPA'

        $servicos = @(Get-Service -ErrorAction SilentlyContinue | Where-Object {
            $exibicao = $_.DisplayName
            @($PadroesServicoRpa | Where-Object { $exibicao -like $_ }).Count -gt 0
        })
        $listaServicos = @(foreach ($s in $servicos) {
            if ($s.Status -ne 'Running') {
                Add-Achado CRITICO $secao "Serviço '$($s.DisplayName)' não está em execução (estado: $($s.Status))." 'Inicie o serviço e verifique o log do agente antes de reexecutar o bot.'
            }
            [ordered]@{
                Nome          = $s.Name
                Exibicao      = $s.DisplayName
                Estado        = $s.Status.ToString()
                Inicializacao = $s.StartType.ToString()
            }
        })
        if ($servicos.Count -eq 0) {
            Add-Achado INFO $secao "Nenhum serviço corresponde a: $($PadroesServicoRpa -join ', ')." 'Ajuste o parâmetro -PadroesServicoRpa com o nome do serviço do seu agente.'
        }

        $processos = @(Get-Process | Where-Object {
            $nomeProc = $_.ProcessName
            @($PadroesProcessoRpa | Where-Object { $nomeProc -like $_ }).Count -gt 0
        })
        $listaProcessos = @(foreach ($p in $processos) {
            $inicio = try { $p.StartTime.ToString('dd/MM/yyyy HH:mm:ss') } catch { 'N/D' }
            [ordered]@{
                Nome      = $p.ProcessName
                PID       = $p.Id
                Sessao    = $p.SessionId
                MemoriaMB = [math]::Round($p.WorkingSet64 / 1MB, 1)
                Inicio    = $inicio
            }
        })

        # Instâncias de Office/IE sem janela visível na sessão do bot costumam ser restos de
        # execuções que terminaram com erro sem fechar a aplicação.
        $orfaos = @(foreach ($nome in 'EXCEL', 'WINWORD', 'POWERPNT', 'iexplore') {
            $instancias = @(Get-Process -Name $nome -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $script:IdSessao })
            if ($instancias.Count -eq 0) { continue }
            $semJanela = @($instancias | Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
            if ($semJanela -gt 0) {
                Add-Achado ATENCAO $secao "$semJanela instância(s) de $nome sem janela visível na sessão do bot." 'Instâncias órfãs travam arquivos e consomem memória/GDI. Feche-as no bloco de tratamento de erro (Finally) do bot.'
            }
            '{0}: {1} instância(s), {2} sem janela' -f $nome, $instancias.Count, $semJanela
        })

        [ordered]@{
            Servicos              = $listaServicos
            Processos             = $listaProcessos
            AplicacoesNaSessao    = $orfaos
        }
    }

    #endregion

    #region ---------------------------- 11. Eventos recentes ----------------------------

    Invoke-Secao 'Eventos Recentes' {
        $secao  = 'Eventos Recentes'
        $inicio = (Get-Date).AddHours(-$HorasEventos)

        $tdr     = Get-EventosSeguros @{ LogName = 'System'; ProviderName = 'Display'; Id = 4101; StartTime = $inicio }
        $quedas  = Get-EventosSeguros @{ LogName = 'System'; Id = 41, 6008; StartTime = $inicio }
        $falhas  = Get-EventosSeguros @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; Id = 1000, 1002; StartTime = $inicio } -Maximo 500
        $rdp     = Get-EventosSeguros @{ LogName = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'; Id = 24, 25; StartTime = $inicio } -Maximo 500

        $formatarContagem = {
            param($resultado)
            if ($resultado.Erro) { "Não lido: $($resultado.Erro)" } else { @($resultado.Eventos).Count }
        }

        if (@($tdr.Eventos).Count -gt 0) {
            Add-Achado ATENCAO $secao "$(@($tdr.Eventos).Count) reinício(s) do driver de vídeo (TDR) nas últimas $HorasEventos h." 'Telas pretas ou congeladas momentâneas coincidem com esses eventos. Atualize o driver de vídeo.'
        }
        if (@($quedas.Eventos).Count -gt 0) {
            Add-Achado ATENCAO $secao "$(@($quedas.Eventos).Count) desligamento(s) inesperado(s) nas últimas $HorasEventos h." 'Investigue energia, hipervisor ou travamentos do sistema.'
        }

        $desconexoes = @($rdp.Eventos | Where-Object { $_.Id -eq 24 })
        if ($desconexoes.Count -ge 3) {
            Add-Achado ATENCAO $secao "$($desconexoes.Count) desconexões de sessão nas últimas $HorasEventos h." 'Desconexões frequentes indicam rede instável, timeouts ou a sessão sendo "roubada" por outro logon.'
        }

        $aplicacoesComFalha = @($falhas.Eventos |
            Group-Object -Property { $_.Properties[0].Value } |
            Sort-Object -Property Count -Descending |
            Select-Object -First 5 |
            ForEach-Object { '{0}: {1} ocorrência(s)' -f $_.Name, $_.Count })

        $ultimosRdp = @($rdp.Eventos |
            Select-Object -First 5 |
            ForEach-Object { '{0} - {1}' -f $_.TimeCreated.ToString('dd/MM HH:mm:ss'), $(if ($_.Id -eq 24) { 'Desconectada' } else { 'Reconectada' }) })

        [ordered]@{
            JanelaHoras            = $HorasEventos
            ResetsDriverVideo      = & $formatarContagem $tdr
            DesligamentosInesperados = & $formatarContagem $quedas
            FalhasAplicacoes       = & $formatarContagem $falhas
            AplicacoesMaisAfetadas = $aplicacoesComFalha
            EventosSessaoRdp       = & $formatarContagem $rdp
            UltimosEventosRdp      = $ultimosRdp
        }
    }

    #endregion

    #region ---------------------------- 12. Teste de captura de tela ----------------------------

    if (-not $PularTesteCaptura) {
        Invoke-Secao 'Teste de Captura' {
            $secao   = 'Teste de Captura'
            $limites = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            $imagem  = New-Object System.Drawing.Bitmap($limites.Width, $limites.Height)
            $grafico = [System.Drawing.Graphics]::FromImage($imagem)
            try {
                try {
                    $grafico.CopyFromScreen($limites.Location, [System.Drawing.Point]::Empty, $limites.Size)
                }
                catch {
                    Add-Achado CRITICO $secao "Não foi possível capturar a tela: $($_.Exception.Message)" 'Esse é o mesmo mecanismo usado por OCR e reconhecimento de imagem: a sessão provavelmente está bloqueada, desconectada ou sem desktop.'
                    return [ordered]@{ Sucesso = $false; Erro = $_.Exception.Message }
                }

                # Amostra uma grade de ~60x60 pontos para medir se a imagem é preta ou uniforme.
                $passoX = [math]::Max(1, [int]($limites.Width / 60))
                $passoY = [math]::Max(1, [int]($limites.Height / 60))
                $total = 0; $pretos = 0
                $cores = New-Object 'System.Collections.Generic.HashSet[int]'
                for ($x = 0; $x -lt $limites.Width; $x += $passoX) {
                    for ($y = 0; $y -lt $limites.Height; $y += $passoY) {
                        $pixel = $imagem.GetPixel($x, $y)
                        $total++
                        if ($pixel.R -lt 16 -and $pixel.G -lt 16 -and $pixel.B -lt 16) { $pretos++ }
                        [void]$cores.Add($pixel.ToArgb())
                    }
                }
                $percentualPreto = [math]::Round($pretos / $total * 100, 1)

                if ($percentualPreto -ge 98) {
                    Add-Achado CRITICO $secao "A captura de tela saiu praticamente preta ($percentualPreto% de pixels pretos)." 'Causa clássica de OCR vazio: sessão minimizada, desconectada ou bloqueada.'
                }
                elseif ($cores.Count -le 3) {
                    Add-Achado ATENCAO $secao "A captura de tela é uniforme ($($cores.Count) cor(es) distintas)." 'A tela pode estar coberta por proteção de tela, papel de parede sem janelas ou sessão não renderizada.'
                }

                $arquivoCaptura = 'Não salva (use -SalvarCaptura)'
                if ($SalvarCaptura) {
                    $arquivoCaptura = Join-Path -Path $PastaSaida -ChildPath "$nomeBase`_captura.png"
                    $imagem.Save($arquivoCaptura, [System.Drawing.Imaging.ImageFormat]::Png)
                }

                [ordered]@{
                    Sucesso          = $true
                    AreaCapturada    = '{0}x{1}' -f $limites.Width, $limites.Height
                    PontosAmostrados = $total
                    PercentualPreto  = "$percentualPreto%"
                    CoresDistintas   = $cores.Count
                    Arquivo          = $arquivoCaptura
                }
            }
            finally {
                $grafico.Dispose()
                $imagem.Dispose()
            }
        }
    }

    #endregion

    #region ---------------------------- Relatório ----------------------------

    $ordem = @{ CRITICO = 0; ATENCAO = 1; INFO = 2 }
    $achadosOrdenados = @($script:Achados | Sort-Object -Property { $ordem[$_.Severidade] })
    $qtdCritico = @($script:Achados | Where-Object { $_.Severidade -eq 'CRITICO' }).Count
    $qtdAtencao = @($script:Achados | Where-Object { $_.Severidade -eq 'ATENCAO' }).Count
    $qtdInfo    = @($script:Achados | Where-Object { $_.Severidade -eq 'INFO' }).Count

    $status = 'OK'; $codigoSaida = 0
    if ($qtdAtencao -gt 0) { $status = 'ATENCAO'; $codigoSaida = 1 }
    if ($qtdCritico -gt 0) { $status = 'CRITICO'; $codigoSaida = 2 }

    $linhas = New-Object 'System.Collections.Generic.List[string]'
    $separador = '=' * 78
    $linhas.Add($separador)
    $linhas.Add('  RELATÓRIO DE DIAGNÓSTICO PARA AUTOMAÇÃO RPA (GUI / OCR / RDP)')
    $linhas.Add($separador)
    $linhas.Add("Gerado em : $((Get-Date).ToString('dd/MM/yyyy HH:mm:ss'))")
    $linhas.Add("Computador: $(Protect-Valor $env:COMPUTERNAME)")
    $linhas.Add("Script    : v$($script:VersaoScript)")
    $linhas.Add("STATUS    : $status  (críticos: $qtdCritico | atenção: $qtdAtencao | informativos: $qtdInfo)")
    $linhas.Add('')
    $linhas.Add('--- [ ACHADOS ] ---')
    if ($achadosOrdenados.Count -eq 0) { $linhas.Add('Nenhum problema identificado.') }
    foreach ($achado in $achadosOrdenados) {
        $linhas.Add(('[{0}] {1}: {2}' -f $achado.Severidade, $achado.Secao, $achado.Mensagem))
        if ($achado.Recomendacao) { $linhas.Add("    -> $($achado.Recomendacao)") }
    }
    foreach ($nomeSecao in $script:Dados.Keys) {
        $linhas.Add('')
        $linhas.Add("--- [ $nomeSecao ] ---")
        foreach ($l in (ConvertTo-LinhasTexto -Valor $script:Dados[$nomeSecao])) { $linhas.Add($l) }
    }

    $arquivoTxt  = Join-Path -Path $PastaSaida -ChildPath "$nomeBase.txt"
    $arquivoJson = Join-Path -Path $PastaSaida -ChildPath "$nomeBase.json"

    # TXT com BOM (abre corretamente no Bloco de Notas e no PowerShell 5.1); JSON sem BOM (parsers).
    [System.IO.File]::WriteAllText($arquivoTxt, ($linhas -join [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($true)))

    $json = [ordered]@{
        VersaoScript = $script:VersaoScript
        GeradoEm     = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        Status       = $status
        Achados      = $achadosOrdenados
        Dados        = $script:Dados
    } | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($arquivoJson, $json, (New-Object System.Text.UTF8Encoding($false)))

    # Retenção: remove relatórios antigos deste script.
    $limiteRetencao = (Get-Date).AddDays(-$DiasRetencao)
    Get-ChildItem -LiteralPath $PastaSaida -File -Filter 'DiagRPA_*' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limiteRetencao } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    Write-Output ('STATUS={0};CRITICOS={1};ATENCAO={2};RELATORIO={3};JSON={4}' -f $status, $qtdCritico, $qtdAtencao, $arquivoTxt, $arquivoJson)
    exit $codigoSaida

    #endregion
}
catch {
    Write-Warning "Falha inesperada no diagnóstico: $($_.Exception.Message) (linha $($_.InvocationInfo.ScriptLineNumber))"
    Write-Output "STATUS=ERRO_SCRIPT;MENSAGEM=$($_.Exception.Message)"
    exit 3
}