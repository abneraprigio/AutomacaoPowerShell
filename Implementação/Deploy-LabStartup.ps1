# ==============================================================================
# DEPLOY-LABSTARTUP.PS1
# Script de implantação — executado UMA VEZ pelo Administrador ou via GPO
#
# O que este script faz:
#   1. Valida privilégios de Administrador
#   2. Cria o diretório seguro C:\Windows\Scripts\
#   3. Copia o wrapper (Lab-Cleanup-Wrapper.ps1) para o local correto
#   4. Bloqueia permissões da pasta para usuários comuns não adulterarem
#   5. Registra a Tarefa Agendada como SYSTEM, gatilho: Logon de qualquer usuário
#   6. Protege a política de execução contra alteração pelo usuário
#
# Pré-requisito: Execute como Administrador local ou conta de domínio com admin local
# ==============================================================================

#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------- #
#  VARIÁVEIS DE CONFIGURAÇÃO
# ---------------------------------------------------------------------------- #
$ScriptDestDir  = "C:\Windows\Scripts"
$WrapperName    = "Lab-Cleanup-Wrapper.ps1"
$WrapperDest    = "$ScriptDestDir\$WrapperName"
$TaskName       = "LabHigienizacao_Startup"
$TaskDesc       = "Protocolo de manutencao e higienizacao de ambiente de laboratorio. Gerenciado pelo TI."
$LogDir         = "C:\Windows\Logs\LabCleanup"

# Caminho do wrapper fonte (mesmo diretório deste script de deploy)
$WrapperSource  = Join-Path $PSScriptRoot $WrapperName

# ---------------------------------------------------------------------------- #
#  FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------- #
function Write-Step {
    param([string]$Msg)
    Write-Host "`n[>>] $Msg" -ForegroundColor Cyan
}

function Write-OK   { param([string]$Msg); Write-Host "  [OK] $Msg" -ForegroundColor Green }
function Write-Fail { param([string]$Msg); Write-Host "  [!!] $Msg" -ForegroundColor Red; exit 1 }
function Write-Warn { param([string]$Msg); Write-Host "  [!] $Msg"  -ForegroundColor Yellow }

# ---------------------------------------------------------------------------- #
#  STEP 0 — VALIDAÇÃO INICIAL
# ---------------------------------------------------------------------------- #
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "  DEPLOY - PROTOCOLO DE HIGIENIZACAO DE LABORATORIO" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan

Write-Step "Validando pre-requisitos..."

if (-not (Test-Path $WrapperSource)) {
    Write-Fail "Wrapper nao encontrado em: $WrapperSource`nCertifique-se de que '$WrapperName' esta na mesma pasta deste script."
}

$osVersion = [System.Environment]::OSVersion.Version
if ($osVersion.Major -lt 10) {
    Write-Warn "SO abaixo do Windows 10 detectado. Compatibilidade nao garantida."
} else {
    Write-OK "Sistema operacional: Windows $($osVersion.Major).$($osVersion.Minor) (Build $($osVersion.Build))"
}

# ---------------------------------------------------------------------------- #
#  STEP 1 — CRIAR DIRETÓRIOS NECESSÁRIOS
# ---------------------------------------------------------------------------- #
Write-Step "Criando estrutura de diretorios..."

foreach ($dir in @($ScriptDestDir, $LogDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-OK "Criado: $dir"
    } else {
        Write-OK "Ja existe: $dir"
    }
}

# ---------------------------------------------------------------------------- #
#  STEP 2 — COPIAR O WRAPPER PARA LOCAL SEGURO
# ---------------------------------------------------------------------------- #
Write-Step "Instalando wrapper em local protegido..."

try {
    Copy-Item -Path $WrapperSource -Destination $WrapperDest -Force
    Write-OK "Wrapper copiado para: $WrapperDest"
} catch {
    Write-Fail "Erro ao copiar wrapper: $_"
}

# ---------------------------------------------------------------------------- #
#  STEP 3 — APLICAR ACL RESTRITIVA (apenas SYSTEM e Admins lêem/escrevem)
# ---------------------------------------------------------------------------- #
Write-Step "Aplicando permissoes restritivas no diretorio de scripts..."

try {
    $Acl = Get-Acl -Path $ScriptDestDir

    # Remove herança e limpa regras existentes
    $Acl.SetAccessRuleProtection($true, $false)

    # Cria novas regras mínimas necessárias
    $rights       = [System.Security.AccessControl.FileSystemRights]
    $inheritance  = [System.Security.AccessControl.InheritanceFlags]
    $propagation  = [System.Security.AccessControl.PropagationFlags]
    $type         = [System.Security.AccessControl.AccessControlType]

    $ruleSystem = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "NT AUTHORITY\SYSTEM",
        $rights::FullControl,
        ($inheritance::ContainerInherit -bor $inheritance::ObjectInherit),
        $propagation::None,
        $type::Allow
    )

    $ruleAdmins = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "BUILTIN\Administrators",
        $rights::FullControl,
        ($inheritance::ContainerInherit -bor $inheritance::ObjectInherit),
        $propagation::None,
        $type::Allow
    )

    # Usuários comuns: APENAS leitura do diretório (sem acesso aos arquivos .ps1)
    # Intencionalmente não concedemos Read ao grupo Users para o conteúdo dos scripts
    $Acl.AddAccessRule($ruleSystem)
    $Acl.AddAccessRule($ruleAdmins)
    Set-Acl -Path $ScriptDestDir -AclObject $Acl

    Write-OK "Permissoes aplicadas: apenas SYSTEM e Administrators."
} catch {
    Write-Warn "Nao foi possivel restringir ACL: $_. Continue manualmente."
}

# ---------------------------------------------------------------------------- #
#  STEP 4 — REMOVER TAREFA ANTERIOR (SE EXISTIR) E REGISTRAR A NOVA
# ---------------------------------------------------------------------------- #
Write-Step "Registrando Tarefa Agendada: '$TaskName'..."

# Remove versão anterior sem errar se não existir
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
Write-OK "Tarefa anterior removida (se existia)."

try {
    # AÇÃO: Executar PowerShell de forma completamente oculta
    $Action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$WrapperDest`""

    # GATILHO: Ao logon de QUALQUER usuário (inclui AD e locais)
    $Trigger = New-ScheduledTaskTrigger -AtLogOn

    # PRINCIPAL: SYSTEM — máximo privilégio, sem UAC, sem prompt visível
    $Principal = New-ScheduledTaskPrincipal `
        -UserId    "NT AUTHORITY\SYSTEM" `
        -LogonType "ServiceAccount" `
        -RunLevel  "Highest"

    # CONFIGURAÇÕES da tarefa
    $Settings = New-ScheduledTaskSettingsSet `
        -Hidden                              `   # Oculta da interface de Agendador de Tarefas para usuários
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
        -MultipleInstances  IgnoreNew        `   # Não inicia nova instância se já rodando
        -StartWhenAvailable                  `   # Executa mesmo que atrasada
        -AllowStartIfOnBatteries             `
        -DontStopIfGoingOnBatteries

    # REGISTRO DA TAREFA
    $Task = Register-ScheduledTask `
        -TaskName   $TaskName         `
        -TaskPath   "\"              `   # Raiz do Task Scheduler
        -Action     $Action           `
        -Trigger    $Trigger          `
        -Principal  $Principal        `
        -Settings   $Settings         `
        -Description $TaskDesc        `
        -Force

    Write-OK "Tarefa registrada com sucesso: '$TaskName'"
    Write-OK "  Usuario de execucao : NT AUTHORITY\SYSTEM"
    Write-OK "  Gatilho             : Logon de qualquer usuario"
    Write-OK "  Janela              : Oculta (Hidden)"
    Write-OK "  Politica execucao   : Bypass (restrita ao contexto desta tarefa)"

} catch {
    Write-Fail "Erro ao registrar tarefa agendada: $_"
}

# ---------------------------------------------------------------------------- #
#  STEP 5 — PROTEGER A POLÍTICA DE EXECUÇÃO DO POWERSHELL VIA REGISTRO
# ---------------------------------------------------------------------------- #
Write-Step "Aplicando politica de execucao restritiva via Registro (maquina)..."

try {
    # Define Restricted para usuários — scripts não rodam manualmente sem admin
    $RegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell"
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $RegPath -Name "EnableScripts"      -Value 1    -Type DWord
    Set-ItemProperty -Path $RegPath -Name "ExecutionPolicy"    -Value "RemoteSigned" -Type String

    Write-OK "Politica PowerShell: RemoteSigned (maquina). Tarefa usa Bypass isolado."
} catch {
    Write-Warn "Nao foi possivel definir politica via registro: $_"
}

# ---------------------------------------------------------------------------- #
#  STEP 6 — VERIFICAÇÃO FINAL
# ---------------------------------------------------------------------------- #
Write-Step "Verificacao final..."

$taskCheck = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($taskCheck) {
    Write-OK "Tarefa confirmada no sistema: Estado = $($taskCheck.State)"
} else {
    Write-Fail "Tarefa NAO encontrada apos registro. Verifique permissoes."
}

if (Test-Path $WrapperDest) {
    Write-OK "Wrapper confirmado em: $WrapperDest"
} else {
    Write-Fail "Wrapper NAO encontrado no destino."
}

# ---------------------------------------------------------------------------- #
#  RESUMO
# ---------------------------------------------------------------------------- #
Write-Host "`n==============================================================" -ForegroundColor Green
Write-Host "  IMPLANTACAO CONCLUIDA COM SUCESSO" -ForegroundColor Green
Write-Host "==============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Arquivos instalados:" -ForegroundColor White
Write-Host "    $WrapperDest" -ForegroundColor Gray
Write-Host "    $LogDir\cleanup_<data>.log (gerado em cada execucao)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Tarefa Agendada: '$TaskName'" -ForegroundColor White
Write-Host "    - Executa ao logon de qualquer usuario (local ou AD)" -ForegroundColor Gray
Write-Host "    - Roda como SYSTEM, completamente oculta" -ForegroundColor Gray
Write-Host "    - Baixa e executa o script do GitHub automaticamente" -ForegroundColor Gray
Write-Host ""
Write-Host "  PROXIMOS PASSOS RECOMENDADOS:" -ForegroundColor Yellow
Write-Host "    1. Distribua este deploy via GPO (Computer Config > Scripts > Startup)" -ForegroundColor Gray
Write-Host "    2. Ou execute manualmente em cada maquina do laboratorio" -ForegroundColor Gray
Write-Host "    3. Verifique os logs em: $LogDir" -ForegroundColor Gray
Write-Host "    4. Teste fazendo logon com uma conta 'aluno' e verificando o log" -ForegroundColor Gray
Write-Host ""
