#Requires -Version 5.1

<#
.SYNOPSIS
    Configura o ambiente de laboratório: baixa o script de reset e registra tarefa agendada.
.DESCRIPTION
    Deve ser executado como Administrador.
    1. Faz download do script Reset-LabSession do GitHub.
    2. Registra tarefa agendada que dispara no logon do usuário "aluno",
       executando com privilégios SYSTEM.
#>

# ─────────────────────────────────────────────────────────────────────────────
# VERIFICAÇÃO DE PRIVILÉGIOS DE ADMINISTRADOR
# ─────────────────────────────────────────────────────────────────────────────
$identidade  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal   = [Security.Principal.WindowsPrincipal] $identidade
$ehAdmin     = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $ehAdmin) {
    Write-Warning "Este script precisa ser executado como Administrador."
    Write-Warning "Clique com o botão direito no PowerShell e escolha 'Executar como administrador'."
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# VARIÁVEIS DE CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────
$urlDownload    = "https://raw.githubusercontent.com/abneraprigio/AutomacaoPowerShell/main/ResetLab/Reset-LabSession%20-%20Adm.ps1"
$pastaDestino   = "C:\Arquivo de Programas"
$arquivoDestino = Join-Path $pastaDestino "Reset-LabSession - Adm.ps1"
$nomeTarefa     = "ResetLabSession"
$usuarioLogon   = "aluno"

# ─────────────────────────────────────────────────────────────────────────────
# PASSO 1 — DOWNLOAD DO SCRIPT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n[1/3] Iniciando download do script..." -ForegroundColor Cyan

try {
    if (-not (Test-Path -Path $pastaDestino)) {
        New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
        Write-Host "      Pasta criada: $pastaDestino"
    } else {
        Write-Host "      Pasta já existe: $pastaDestino"
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Invoke-WebRequest -Uri $urlDownload -OutFile $arquivoDestino -UseBasicParsing -ErrorAction Stop

    if (-not (Test-Path -Path $arquivoDestino)) {
        throw "O arquivo não foi encontrado após o download: $arquivoDestino"
    }
    $tamanho = (Get-Item $arquivoDestino).Length
    if ($tamanho -eq 0) {
        throw "O arquivo baixado está vazio. Verifique a URL e a conectividade."
    }

    Write-Host "      Download concluido com sucesso." -ForegroundColor Green
    Write-Host "      Destino : $arquivoDestino"
    Write-Host "      Tamanho : $tamanho bytes"

} catch {
    Write-Error "Falha no download do script.`nDetalhes: $_"
    exit 2
}

# ─────────────────────────────────────────────────────────────────────────────
# PASSO 2 — CRIAÇÃO DA TAREFA AGENDADA
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n[2/3] Registrando tarefa agendada '$nomeTarefa'..." -ForegroundColor Cyan

try {
    $tarefaExistente = Get-ScheduledTask -TaskName $nomeTarefa -ErrorAction SilentlyContinue
    if ($tarefaExistente) {
        Unregister-ScheduledTask -TaskName $nomeTarefa -Confirm:$false
        Write-Host "      Tarefa anterior removida para recriação."
    }

    $acao = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File `"$arquivoDestino`""

    $gatilho = New-ScheduledTaskTrigger -AtLogOn -User $usuarioLogon

    $configuracoes = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
        -MultipleInstances IgnoreNew `
        -StartWhenAvailable

    $principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    $tarefaDefinicao = New-ScheduledTask `
        -Action      $acao `
        -Trigger     $gatilho `
        -Principal   $principal `
        -Settings    $configuracoes `
        -Description "Reseta o ambiente de laboratorio ao logon do usuario aluno. Gerenciado por Setup-LabTask.ps1."

    Register-ScheduledTask `
        -TaskName    $nomeTarefa `
        -InputObject $tarefaDefinicao `
        -Force `
        -ErrorAction Stop | Out-Null

    Write-Host "      Tarefa registrada com sucesso." -ForegroundColor Green

} catch {
    Write-Error "Falha ao registrar a tarefa agendada.`nDetalhes: $_"
    exit 3
}

# ─────────────────────────────────────────────────────────────────────────────
# PASSO 3 — VALIDAÇÃO FINAL
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n[3/3] Validando configuracao final..." -ForegroundColor Cyan

$tarefaRegistrada = Get-ScheduledTask -TaskName $nomeTarefa -ErrorAction SilentlyContinue
$arquivoPresente  = Test-Path -Path $arquivoDestino

if ($tarefaRegistrada -and $arquivoPresente) {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║           CONFIGURACAO CONCLUIDA COM SUCESSO                ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Script baixado : $arquivoDestino"
    Write-Host "  Tarefa criada  : $nomeTarefa"
    Write-Host "  Executa como   : SYSTEM (privilegios elevados)"
    Write-Host "  Gatilho        : Logon do usuario '$usuarioLogon'"
    Write-Host ""
    Write-Host "  A tarefa sera executada automaticamente na proxima vez que"
    Write-Host "  o usuario '$usuarioLogon' fizer logon neste computador."
    Write-Host ""
} else {
    Write-Host ""
    if (-not $arquivoPresente) {
        Write-Warning "AVISO: O arquivo de script NAO foi encontrado em '$arquivoDestino'."
    }
    if (-not $tarefaRegistrada) {
        Write-Warning "AVISO: A tarefa agendada '$nomeTarefa' NAO foi encontrada no sistema."
    }
    Write-Warning "Revise os erros acima e execute o script novamente."
    exit 4
}
