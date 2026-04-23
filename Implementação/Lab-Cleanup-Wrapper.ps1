# ==============================================================================
# LAB-CLEANUP-WRAPPER.PS1
# Executor oculto do protocolo de higienização de laboratório
# Executado como: SYSTEM (via Task Scheduler)
# Gatilho:        Logon de qualquer usuário
# ==============================================================================

# --- [0] LOG DE EXECUÇÃO -------------------------------------------------------
$LogDir  = "C:\Windows\Logs\LabCleanup"
$LogFile = "$LogDir\cleanup_$(Get-Date -Format 'yyyyMMdd').log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

Write-Log "============================================================"
Write-Log "Wrapper iniciado. Identidade: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# --- [1] AGUARDAR O DESKTOP DO USUÁRIO ESTAR PRONTO ----------------------------
# Pausa para garantir que o perfil do usuário foi completamente carregado
Start-Sleep -Seconds 45

# --- [2] DETECTAR O USUÁRIO INTERATIVO LOGADO ----------------------------------
try {
    # Consulta o processo Explorer.exe para identificar o usuário com sessão ativa
    $explorerProc = Get-WmiObject -Query "SELECT * FROM Win32_Process WHERE Name='explorer.exe'" |
        Select-Object -First 1

    if (-not $explorerProc) {
        Write-Log "Nenhuma sessão interativa (Explorer) detectada. Encerrando." "WARN"
        exit 0
    }

    # Obtém o proprietário do processo Explorer
    $ownerResult  = $explorerProc.GetOwner()
    $UserDomain   = $ownerResult.Domain
    $UserName     = $ownerResult.User

    if (-not $UserName) {
        Write-Log "Não foi possível determinar o usuário logado. Encerrando." "WARN"
        exit 0
    }

    Write-Log "Usuário detectado: $UserDomain\$UserName"

} catch {
    Write-Log "Erro ao detectar usuário: $_" "ERROR"
    exit 1
}

# --- [3] CONSTRUIR OS CAMINHOS DE PERFIL DO USUÁRIO ALVO ----------------------
# Necessário porque o Task roda como SYSTEM: $env:APPDATA apontaria para o perfil
# do SYSTEM e não do usuário logado. Resolvemos via registro de perfis do Windows.

try {
    $ProfileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $ProfileKey = Get-ChildItem $ProfileListPath |
        Where-Object {
            $sid = $_.PSChildName
            try {
                $ntAccount = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate(
                    [System.Security.Principal.NTAccount]
                ).Value
                $ntAccount -like "*\$UserName"
            } catch { $false }
        } | Select-Object -First 1

    if (-not $ProfileKey) {
        Write-Log "Perfil de registro não localizado para '$UserName'. Tentando via C:\Users..." "WARN"
        $UserProfile = "C:\Users\$UserName"
    } else {
        $UserProfile = (Get-ItemProperty -Path $ProfileKey.PSPath).ProfileImagePath
    }

    if (-not (Test-Path $UserProfile)) {
        Write-Log "Caminho de perfil inválido: '$UserProfile'. Encerrando." "ERROR"
        exit 1
    }

    Write-Log "Perfil do usuário resolvido: $UserProfile"

    # Define variáveis de ambiente apontando para o perfil do usuário real
    $env:USERPROFILE  = $UserProfile
    $env:APPDATA      = "$UserProfile\AppData\Roaming"
    $env:LOCALAPPDATA = "$UserProfile\AppData\Local"
    $env:TEMP         = "$UserProfile\AppData\Local\Temp"
    $env:TMP          = "$UserProfile\AppData\Local\Temp"
    $env:HOMEPATH     = $UserProfile -replace "^[A-Za-z]:", ""
    $env:USERNAME     = $UserName

} catch {
    Write-Log "Erro ao resolver perfil: $_" "ERROR"
    exit 1
}

# --- [4] BAIXAR O SCRIPT DO GITHUB ---------------------------------------------
$GitHubRawUrl = "https://raw.githubusercontent.com/abneraprigio/AutomacaoPowerShell/master/Apaga%20Registros/Script%20Desloga%20Usuarios.txt"

try {
    Write-Log "Baixando script de higienização: $GitHubRawUrl"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $webClient    = New-Object System.Net.WebClient
    $ScriptContent = $webClient.DownloadString($GitHubRawUrl)

    if ([string]::IsNullOrWhiteSpace($ScriptContent)) {
        Write-Log "Script baixado está vazio. Encerrando." "ERROR"
        exit 1
    }

    Write-Log "Script baixado com sucesso ($($ScriptContent.Length) bytes)."

} catch {
    Write-Log "Falha ao baixar script do GitHub: $_" "ERROR"
    exit 1
}

# --- [5] SANITIZAR O SCRIPT PARA EXECUÇÃO NÃO-INTERATIVA ---------------------
# Remove blocos que causariam travamento ou redundância ao rodar como SYSTEM

# 5a. Remove o bloco de auto-elevação (já somos SYSTEM, sem necessidade de UAC)
$ScriptContent = $ScriptContent -replace (
    '(?s)if\s*\(\s*!\s*\(\[Security\.Principal\..*?\)\s*\{.*?exit\s*\}',
    '# [WRAPPER] Bloco de elevação removido - executando como SYSTEM'
)

# 5b. Remove o ReadKey interativo no final (travaria em modo não-interativo)
$ScriptContent = $ScriptContent -replace 'Write-Host\s+"Pressione qualquer tecla.*"[^\r\n]*', ''
$ScriptContent = $ScriptContent -replace '\$null\s*=\s*\$Host\.UI\.RawUI\.ReadKey\s*\(.*?\)', ''

# 5c. Remove Write-Host que requerem console visual (opcional: mantém log limpo)
# Mantemos os Write-Host pois eles simplesmente não produzem saída em modo oculto

Write-Log "Script sanitizado para execução não-interativa."

# --- [6] EXECUTAR O SCRIPT SANITIZADO ------------------------------------------
try {
    Write-Log "Iniciando execução do protocolo de higienização para '$UserName'..."

    # Executa no escopo atual (SYSTEM com env vars do usuário alvo)
    Invoke-Expression $ScriptContent

    Write-Log "Protocolo de higienização concluído com sucesso para '$UserName'."

} catch {
    Write-Log "Erro durante execução do script principal: $_" "ERROR"
    exit 1
}

Write-Log "Wrapper finalizado."
Write-Log "============================================================"
exit 0
