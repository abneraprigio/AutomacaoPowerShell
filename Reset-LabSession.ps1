#Requires -Version 5.1
<#
.SYNOPSIS
    Reset-LabSession.ps1 — Script de Wipe/Reset de Sessão para Laboratório de Informática
.DESCRIPTION
    Restaura o computador ao estado "limpo e original" ao final de cada sessão de aluno.
    Encerra processos, apaga perfis de navegadores, tokens Microsoft, configurações de
    ferramentas de desenvolvimento, limpa a Área de Trabalho e realiza RESTAURAÇÃO COMPLETA
    das pastas pessoais do usuário (Downloads, Documentos, Fotos e Vídeos),
    sem alterar permanentemente o SO.
.NOTES
    Autor      : Engenharia de Infraestrutura
    Versão     : 3.0
    Requer     : PowerShell 5.1+, execução como Administrador
    Testado em : Windows 10 21H2 / Windows 11 22H2+
    Changelog  : v3.0 — Adicionado wipe completo de Downloads, Documentos, Fotos e Vídeos
#>

# ==============================================================================
# BLOCO 0 — AUTO-ELEVAÇÃO PARA ADMINISTRADOR
# Se o script não estiver rodando como Admin, ele se relança com privilégios elevados.
# ==============================================================================
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "[INFO] Privilégios insuficientes. Relançando como Administrador..." -ForegroundColor Yellow

    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs
    exit
}

# ==============================================================================
# BLOCO 1 — CONFIGURAÇÃO GLOBAL E FUNÇÕES AUXILIARES
# ==============================================================================

# Impede que erros não críticos interrompam o script globalmente.
$ErrorActionPreference = "SilentlyContinue"

# Captura o perfil do usuário atualmente logado na sessão interativa (não o SYSTEM).
# Isso é necessário porque o script roda como Admin/SYSTEM mas precisa limpar o perfil
# do usuário comum que estava logado.
$TargetUser    = (Get-WmiObject -Class Win32_ComputerSystem).UserName -replace '.*\\'
$TargetProfile = "C:\Users\$TargetUser"

# Fallback: se não conseguir determinar o usuário, usa variáveis de ambiente padrão.
if ([string]::IsNullOrEmpty($TargetUser) -or -not (Test-Path $TargetProfile)) {
    $TargetProfile = $env:USERPROFILE
    $TargetUser    = $env:USERNAME
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   RESET DE SESSÃO DE LABORATÓRIO — Iniciando..." -ForegroundColor Cyan
Write-Host "   Perfil alvo : $TargetProfile" -ForegroundColor Cyan
Write-Host "   Usuário alvo: $TargetUser" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Função auxiliar: remove caminhos com segurança (arquivo ou pasta) ---
function Remove-ItemSafe {
    param([string[]]$Paths)
    foreach ($p in $Paths) {
        if (Test-Path $p) {
            Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  [DEL] $p" -ForegroundColor DarkGray
        }
    }
}

# --- Função auxiliar: mata um processo pelo nome, ignorando se não existir ---
function Stop-ProcessSafe {
    param([string]$ProcessName)
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($procs) {
        # /F = forçar | /T = encerrar árvore de processos filhos também
        taskkill /F /T /IM "$ProcessName.exe" 2>$null | Out-Null
        Write-Host "  [KILL] $ProcessName" -ForegroundColor DarkGray
    }
}

# ==============================================================================
# BLOCO 2 — ENCERRAMENTO FORÇADO DE PROCESSOS (FILE LOCKS)
# Deve ser o PRIMEIRO passo. Arquivos abertos não podem ser deletados.
# Usamos taskkill /F /T para matar a árvore completa de processos.
# ==============================================================================
Write-Host "[ETAPA 1/7] Encerrando processos ativos..." -ForegroundColor Yellow

$ProcessList = @(
    # Navegadores
    "chrome", "msedge", "firefox", "opera", "opera_autoupdate",
    # Microsoft Office
    "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSPUB", "MSACCESS",
    "lync", "Teams", "Teams2",
    # Power BI
    "PBIDesktop",
    # IDEs e Dev Tools
    "Code",           # VS Code
    "GithubDesktop",  # GitHub Desktop
    "git",
    "node", "npm",
    "python", "pythonw",
    "java", "javaw",
    # Cisco / Educação
    "PacketTracer",
    "PacketTracer8",   # versões futuras
    "Antigravity",
    # Bloco de Notas UWP (pode ter processo Win32 auxiliar)
    "notepad",
    "WindowsNotepad"
)

foreach ($proc in $ProcessList) {
    Stop-ProcessSafe -ProcessName $proc
}

# Aguarda 2 segundos para garantir que os processos foram finalizados
# e os file locks foram liberados antes de prosseguir com as deleções.
Start-Sleep -Seconds 2
Write-Host "  [OK] Processos encerrados." -ForegroundColor Green

# ==============================================================================
# BLOCO 3 — LIMPEZA DE NAVEGADORES (PERFIS, CACHE, HISTÓRICO, SESSÕES)
# Estratégia: apagar os diretórios "User Data" e "Profiles" do LocalAppData/Roaming.
# Isso destrói: cookies, histórico, senhas salvas, sessões ativas, extensões de sessão.
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 2/7] Limpando perfis e dados de navegadores..." -ForegroundColor Yellow

# --- 3.1 GOOGLE CHROME ---
# O Chrome armazena TUDO dentro de "User Data". Apagar esta pasta equivale
# a um reset completo. O Chrome recria a estrutura mínima na próxima abertura.
Write-Host "  -> Google Chrome" -ForegroundColor White
$ChromePaths = @(
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Cache",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Code Cache",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\History",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Cookies",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Login Data",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Web Data",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Sessions",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Session Storage",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Local Storage",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\IndexedDB",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Extension Cookies",
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Network",
    # Token de login da conta Google sincronizada
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Default\Google Profile.ico",
    # Limpa perfis secundários (Profile 1, Profile 2, ...)
    "$TargetProfile\AppData\Local\Google\Chrome\User Data\Profile*"
)
Remove-ItemSafe -Paths $ChromePaths

# Apaga o arquivo "Local State" que guarda qual conta Google está conectada
$ChromeLocalState = "$TargetProfile\AppData\Local\Google\Chrome\User Data\Local State"
if (Test-Path $ChromeLocalState) {
    # Em vez de apagar, sobrescreve com um JSON mínimo para não quebrar o Chrome
    '{"browser":{"enabled_labs_experiments":[]}}' | Set-Content -Path $ChromeLocalState -Force
    Write-Host "  [RESET] Chrome Local State redefinido." -ForegroundColor DarkGray
}

# --- 3.2 MICROSOFT EDGE ---
# Edge Chromium usa estrutura idêntica ao Chrome, mas em "Microsoft\Edge"
Write-Host "  -> Microsoft Edge" -ForegroundColor White
$EdgePaths = @(
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Cache",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Code Cache",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\History",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Cookies",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Login Data",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Web Data",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Sessions",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Session Storage",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Local Storage",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\IndexedDB",
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Default\Network",
    # Perfis secundários do Edge
    "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Profile*",
    # Edge também salva dados no Roaming
    "$TargetProfile\AppData\Roaming\Microsoft\Edge"
)
Remove-ItemSafe -Paths $EdgePaths

$EdgeLocalState = "$TargetProfile\AppData\Local\Microsoft\Edge\User Data\Local State"
if (Test-Path $EdgeLocalState) {
    '{"browser":{"enabled_labs_experiments":[]}}' | Set-Content -Path $EdgeLocalState -Force
    Write-Host "  [RESET] Edge Local State redefinido." -ForegroundColor DarkGray
}

# --- 3.3 MOZILLA FIREFOX ---
# Firefox usa um sistema de "Profiles" diferente. O perfil fica em Roaming.
# Cada perfil é uma pasta com nome aleatório (ex: a1b2c3d4.default-release).
Write-Host "  -> Mozilla Firefox" -ForegroundColor White
$FirefoxProfilesRoot = "$TargetProfile\AppData\Roaming\Mozilla\Firefox\Profiles"
if (Test-Path $FirefoxProfilesRoot) {
    Get-ChildItem -Path $FirefoxProfilesRoot -Directory | ForEach-Object {
        $ffProfile = $_.FullName
        # Apaga os artefatos de sessão dentro de cada perfil individualmente
        $FirefoxItems = @(
            "$ffProfile\cache2",
            "$ffProfile\cookies.sqlite",
            "$ffProfile\places.sqlite",        # histórico e bookmarks
            "$ffProfile\formhistory.sqlite",
            "$ffProfile\logins.json",           # senhas salvas
            "$ffProfile\key4.db",               # chave de criptografia das senhas
            "$ffProfile\sessionstore.jsonlz4",  # abas da última sessão
            "$ffProfile\sessionstore-backups",
            "$ffProfile\storage",               # LocalStorage / IndexedDB
            "$ffProfile\thumbnails",
            "$ffProfile\crashes",
            "$ffProfile\datareporting"
        )
        Remove-ItemSafe -Paths $FirefoxItems
    }
}
# Limpa também o cache local do Firefox
Remove-ItemSafe -Paths @(
    "$TargetProfile\AppData\Local\Mozilla\Firefox\Profiles"
)

# --- 3.4 OPERA ---
# Opera Stable usa "Opera Software\Opera Stable" tanto no Local quanto no Roaming.
Write-Host "  -> Opera" -ForegroundColor White
$OperaPaths = @(
    "$TargetProfile\AppData\Roaming\Opera Software\Opera Stable",
    "$TargetProfile\AppData\Local\Opera Software\Opera Stable",
    # Opera GX (versão gamer)
    "$TargetProfile\AppData\Roaming\Opera Software\Opera GX Stable",
    "$TargetProfile\AppData\Local\Opera Software\Opera GX Stable"
)
Remove-ItemSafe -Paths $OperaPaths

Write-Host "  [OK] Navegadores limpos." -ForegroundColor Green

# ==============================================================================
# BLOCO 4 — DESLOGAMENTO DO ECOSSISTEMA MICROSOFT
# Remove tokens de autenticação do Office, Power BI e contas Windows.
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 3/7] Removendo contas e tokens Microsoft..." -ForegroundColor Yellow

# --- 4.1 MICROSOFT OFFICE (Word, Excel, PowerPoint, Outlook, etc.) ---
# O Office armazena tokens OAuth e cache de identidade em duas localizações no Registro.
# Apagando essas chaves, o Office solicitará login novamente.
Write-Host "  -> Microsoft Office (tokens de identidade)" -ForegroundColor White

$OfficeRegistryPaths = @(
    # Tokens de identidade do Office 365 / Microsoft 365 (chave principal)
    "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
    "HKCU:\Software\Microsoft\Office\15.0\Common\Identity",
    # Cache de usuário conectado no Office
    "HKCU:\Software\Microsoft\Office\16.0\Common\ServicesManagerCache",
    # Configurações de conta do Outlook
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles",
    # Dados do Teams clássico integrado ao Office
    "HKCU:\Software\Microsoft\Office\Teams"
)
foreach ($regPath in $OfficeRegistryPaths) {
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [DEL REG] $regPath" -ForegroundColor DarkGray
    }
}

# Limpa também os arquivos de cache de identidade do Office em disco
$OfficeCachePaths = @(
    # MSAL (Modern Authentication) token cache — usado pelo Office 365
    "$TargetProfile\AppData\Local\Microsoft\Office\16.0\Licensing",
    "$TargetProfile\AppData\Local\Microsoft\TokenBroker\Cache",  # (coberto abaixo também)
    # Arquivos temporários de documentos abertos
    "$TargetProfile\AppData\Local\Microsoft\Windows\INetCache\Content.MSO",
    "$TargetProfile\AppData\Roaming\Microsoft\Templates",
    # Histórico de documentos recentes do Office
    "$TargetProfile\AppData\Roaming\Microsoft\Office\Recent"
)
Remove-ItemSafe -Paths $OfficeCachePaths

# --- 4.2 POWER BI DESKTOP ---
# O Power BI salva tokens OAuth e credenciais de fontes de dados nesses caminhos.
Write-Host "  -> Power BI Desktop" -ForegroundColor White
$PowerBIPaths = @(
    "$TargetProfile\AppData\Local\Microsoft\Power BI Desktop",
    "$TargetProfile\AppData\Roaming\Microsoft\Power BI Desktop",
    # Cache de análise e modelos temporários
    "$TargetProfile\AppData\Local\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"
)
Remove-ItemSafe -Paths $PowerBIPaths

$PowerBIRegistryPaths = @(
    "HKCU:\Software\Microsoft\Microsoft Power BI Desktop"
)
foreach ($regPath in $PowerBIRegistryPaths) {
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [DEL REG] $regPath" -ForegroundColor DarkGray
    }
}

# --- 4.3 WINDOWS ACCOUNTS — TOKENBROKER ---
# O TokenBroker é o gerenciador de tokens SSO do Windows (WAM - Windows Account Manager).
# Ele guarda tokens de acesso para qualquer conta Microsoft adicionada ao Windows
# (pessoal, corporativa, escolar). Destruir esta pasta invalida todos os tokens ativos.
Write-Host "  -> Windows TokenBroker (contas Microsoft do sistema)" -ForegroundColor White
$TokenBrokerPaths = @(
    "$TargetProfile\AppData\Local\Microsoft\TokenBroker",
    "$TargetProfile\AppData\Local\Microsoft\Windows\CloudStore",  # Estado de sincronização da conta
    # Wallet / passkeys
    "$TargetProfile\AppData\Local\Microsoft\Windows\WebAccountManager"
)
Remove-ItemSafe -Paths $TokenBrokerPaths

# --- 4.4 GERENCIADOR DE CREDENCIAIS DO WINDOWS (WINDOWS VAULT) ---
# O cmdkey é a ferramenta nativa para listar e remover credenciais do Credential Manager.
# Isso remove senhas de rede, RDP, aplicativos que usam DPAPI, etc.
Write-Host "  -> Windows Credential Manager (Vault)" -ForegroundColor White

# Lista todas as credenciais armazenadas e as remove uma a uma
$cmdkeyList = cmdkey /list 2>$null
if ($cmdkeyList) {
    $cmdkeyList | Where-Object { $_ -match "^\s+Destino:\s*(.+)$" -or $_ -match "^\s+Target:\s*(.+)$" } | ForEach-Object {
        # Extrai o nome do destino da credencial (compatível com PT-BR e EN)
        if ($_ -match "Destino:\s*(.+)$" -or $_ -match "Target:\s*(.+)$") {
            $credTarget = $Matches[1].Trim()
            cmdkey /delete:"$credTarget" 2>$null | Out-Null
            Write-Host "  [DEL CRED] $credTarget" -ForegroundColor DarkGray
        }
    }
}

Write-Host "  [OK] Ecossistema Microsoft limpo." -ForegroundColor Green

# ==============================================================================
# BLOCO 5 — LIMPEZA DE FERRAMENTAS DE DESENVOLVIMENTO E EDUCAÇÃO
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 4/7] Limpando ferramentas de desenvolvimento e educação..." -ForegroundColor Yellow

# --- 5.1 VISUAL STUDIO CODE ---
# O VS Code guarda: extensões instaladas pelo usuário, configurações (settings.json),
# snippets personalizados, estado da janela e histórico de arquivos recentes.
# A pasta "User" dentro de "Code" contém TUDO relacionado ao perfil do usuário.
Write-Host "  -> Visual Studio Code" -ForegroundColor White
$VSCodePaths = @(
    # Configurações, snippets e keybindings do usuário
    "$TargetProfile\AppData\Roaming\Code\User\settings.json",
    "$TargetProfile\AppData\Roaming\Code\User\keybindings.json",
    "$TargetProfile\AppData\Roaming\Code\User\snippets",
    # Histórico de arquivos abertos recentemente (workspaceStorage)
    "$TargetProfile\AppData\Roaming\Code\User\workspaceStorage",
    "$TargetProfile\AppData\Roaming\Code\User\History",          # Edições não salvas (Ctrl+Z history)
    "$TargetProfile\AppData\Roaming\Code\User\globalStorage",    # Estado de extensões
    # Cache local do VS Code (thumbnails, extensões baixadas pelo usuário)
    "$TargetProfile\AppData\Local\Programs\Microsoft VS Code\resources\app\user-data",
    # Pasta de logs
    "$TargetProfile\AppData\Roaming\Code\logs"
)
Remove-ItemSafe -Paths $VSCodePaths

# Extensões instaladas pelo usuário ficam em ~/.vscode/extensions
# Removemos para não acumular extensões entre alunos
$VSCodeExtensions = "$TargetProfile\.vscode\extensions"
if (Test-Path $VSCodeExtensions) {
    Remove-Item -Path $VSCodeExtensions -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  [DEL] Extensões do VS Code: $VSCodeExtensions" -ForegroundColor DarkGray
}

# --- 5.2 GIT E GITHUB DESKTOP ---
# .gitconfig guarda: nome, e-mail, chaves de assinatura, configurações globais.
# .ssh guarda: chaves privadas RSA/ED25519 do aluno (CRÍTICO para segurança!).
Write-Host "  -> Git (configuração global e chaves SSH)" -ForegroundColor White
$GitPaths = @(
    # Identidade global do Git — nome e e-mail do committer
    "$TargetProfile\.gitconfig",
    "$TargetProfile\.gitconfig_local",
    # Chaves SSH (incluindo known_hosts e authorized_keys)
    "$TargetProfile\.ssh",
    # Cache de credenciais do Git (git-credential-store)
    "$TargetProfile\.git-credentials",
    # Configuração do Git Credential Manager (GCM)
    "$TargetProfile\AppData\Local\GitCredentialManager",
    "$TargetProfile\AppData\Roaming\GitHub Desktop",  # Sessão do GitHub Desktop
    "$TargetProfile\AppData\Local\GitHub Desktop"     # Cache local do GitHub Desktop
)
Remove-ItemSafe -Paths $GitPaths

# Remove tokens do GitHub armazenados no Registro pelo GCM
$GCMRegistryPaths = @(
    "HKCU:\Software\GitCredentialManager"
)
foreach ($regPath in $GCMRegistryPaths) {
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [DEL REG] $regPath" -ForegroundColor DarkGray
    }
}

# --- 5.3 CISCO PACKET TRACER ---
# O Packet Tracer salva: arquivos de atividades recentes, configurações de usuário
# e (em versões recentes) credenciais de conta Cisco Networking Academy (NetAcad).
Write-Host "  -> Cisco Packet Tracer" -ForegroundColor White
$PacketTracerPaths = @(
    # Atividades e configurações do usuário (versões 7.x e 8.x)
    "$TargetProfile\Cisco Packet Tracer $([string]([System.Version]'8.0'))*",
    "$TargetProfile\AppData\Roaming\Cisco\Packet Tracer",
    "$TargetProfile\AppData\Local\Cisco\Packet Tracer",
    # Diretório home padrão criado pelo PT
    "$TargetProfile\PT Saves",
    "$TargetProfile\Documents\Packet Tracer"
)
Remove-ItemSafe -Paths $PacketTracerPaths

# Limpa o diretório específico por versão do Packet Tracer
Get-ChildItem -Path $TargetProfile -Directory -Filter "Cisco Packet Tracer *" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-ItemSafe -Paths @($_.FullName) }

# --- 5.4 ANTIGRAVITY ---
# Limpa configurações e dados de sessão do Antigravity.
Write-Host "  -> Antigravity" -ForegroundColor White
$AntigravityPaths = @(
    "$TargetProfile\AppData\Roaming\Antigravity",
    "$TargetProfile\AppData\Local\Antigravity"
)
Remove-ItemSafe -Paths $AntigravityPaths

Write-Host "  [OK] Ferramentas de desenvolvimento limpas." -ForegroundColor Green

# ==============================================================================
# BLOCO 6 — LIMPEZA DA ÁREA DE TRABALHO
# Remove arquivos e pastas criados pelo aluno, PRESERVANDO atalhos (.lnk)
# e arquivos de configuração (.ini), que pertencem à imagem do laboratório.
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 5/7] Limpando Área de Trabalho do usuário..." -ForegroundColor Yellow

$DesktopPath = "$TargetProfile\Desktop"

# Extensões protegidas — pertencem à configuração do laboratório, NÃO deletar.
$ProtectedExtensions = @(".lnk", ".ini", ".url")

if (Test-Path $DesktopPath) {
    # Itera sobre todos os itens da Área de Trabalho
    Get-ChildItem -Path $DesktopPath -Force | ForEach-Object {
        $item = $_
        $shouldProtect = $false

        # Verifica se o item é um arquivo com extensão protegida
        if (-not $item.PSIsContainer) {
            foreach ($ext in $ProtectedExtensions) {
                if ($item.Extension -ieq $ext) {
                    $shouldProtect = $true
                    break
                }
            }
        }

        if (-not $shouldProtect) {
            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  [DEL] Desktop\$($item.Name)" -ForegroundColor DarkGray
        } else {
            Write-Host "  [KEEP] Desktop\$($item.Name)" -ForegroundColor DarkCyan
        }
    }
}

Write-Host "  [OK] Área de Trabalho limpa (atalhos e .ini preservados)." -ForegroundColor Green

# ==============================================================================
# BLOCO 6.5 — RESTAURAÇÃO COMPLETA DAS PASTAS PESSOAIS DO USUÁRIO
# Apaga TODO o conteúdo de Downloads, Documentos, Fotos e Vídeos.
# Esta etapa transforma o reset parcial numa RESTAURAÇÃO COMPLETA do perfil,
# garantindo que nenhum arquivo produzido ou baixado pelo aluno persista.
#
# ARQUITETURA DE SEGURANÇA:
#   - Usamos Shell.Application (COM) para resolver os caminhos reais das pastas
#     conhecidas (KnownFolders), pois no Windows 11 o usuário pode ter redirecionado
#     Downloads ou Documentos para outro drive (ex.: D:\Downloads).
#   - Como fallback, usamos os caminhos padrão relativos ao perfil.
#   - A pasta em si é PRESERVADA — apenas o conteúdo é removido.
#     Apagar a pasta raiz quebraria os ponteiros do Shell e atalhos de terceiros.
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 6/7] Restauração completa — Pastas pessoais do usuário..." -ForegroundColor Yellow

# --- Função auxiliar especializada: esvazia uma pasta preservando sua raiz ---
# Parâmetros:
#   $FolderPath  : caminho completo da pasta a esvaziar
#   $FolderLabel : nome amigável exibido no log
function Clear-UserFolder {
    param(
        [string]$FolderPath,
        [string]$FolderLabel
    )

    Write-Host "  -> $FolderLabel" -ForegroundColor White

    if (-not (Test-Path $FolderPath)) {
        Write-Host "     [SKIP] Pasta não encontrada: $FolderPath" -ForegroundColor DarkGray
        return
    }

    # Conta itens antes para o relatório
    $itemsBefore = (Get-ChildItem -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue).Count

    # Remove TUDO dentro da pasta (arquivos + subpastas), mas não a pasta raiz
    Get-ChildItem -Path $FolderPath -Force -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Verifica se ainda restou algo (arquivos bloqueados por outro processo)
    $itemsAfter = (Get-ChildItem -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue).Count

    if ($itemsAfter -eq 0) {
        Write-Host "     [OK] $itemsBefore item(s) removido(s). Pasta vazia." -ForegroundColor Green
    } else {
        Write-Host "     [WARN] $itemsAfter item(s) não puderam ser removidos (file lock?). Verifique." -ForegroundColor DarkYellow
    }
}

# --- Resolve os caminhos reais das pastas conhecidas via Shell COM ---
# Esta abordagem respeita redirecionamentos do usuário (OneDrive, outro drive, etc.)
$ShellApp = $null
try {
    $ShellApp = New-Object -ComObject Shell.Application

    # GUIDs / índices das pastas conhecidas no Shell.NameSpace:
    #   0x10 (16) = Downloads   |  0x05 (5)  = Documentos
    #   0x27 (39) = Imagens     |  0x0E (14) = Vídeos
    $DownloadsReal  = $ShellApp.NameSpace(0x10).Self.Path
    $DocumentsReal  = $ShellApp.NameSpace(0x05).Self.Path
    $PicturesReal   = $ShellApp.NameSpace(0x27).Self.Path
    $VideosReal     = $ShellApp.NameSpace(0x0E).Self.Path
} catch {
    # Fallback: caminhos padrão caso o COM falhe (ex.: contexto SYSTEM puro)
    $DownloadsReal = $null
    $DocumentsReal = $null
    $PicturesReal  = $null
    $VideosReal    = $null
}

# Fallback manual usando o perfil do usuário alvo
# (necessário quando o script roda como SYSTEM e o Shell.App retorna caminhos do SYSTEM)
function Resolve-FolderPath {
    param([string]$ComPath, [string]$DefaultRelative)
    # Se o COM retornou um caminho válido E ele está dentro do perfil alvo, usa-o.
    # Caso contrário, usa o caminho padrão.
    $defaultFull = Join-Path $TargetProfile $DefaultRelative
    if ($ComPath -and (Test-Path $ComPath) -and $ComPath -like "*$TargetUser*") {
        return $ComPath
    }
    return $defaultFull
}

$PathDownloads = Resolve-FolderPath -ComPath $DownloadsReal  -DefaultRelative "Downloads"
$PathDocuments = Resolve-FolderPath -ComPath $DocumentsReal  -DefaultRelative "Documents"
$PathPictures  = Resolve-FolderPath -ComPath $PicturesReal   -DefaultRelative "Pictures"
$PathVideos    = Resolve-FolderPath -ComPath $VideosReal     -DefaultRelative "Videos"

Write-Host "  Caminhos resolvidos:" -ForegroundColor DarkCyan
Write-Host "    Downloads  : $PathDownloads"  -ForegroundColor DarkCyan
Write-Host "    Documentos : $PathDocuments"  -ForegroundColor DarkCyan
Write-Host "    Fotos      : $PathPictures"   -ForegroundColor DarkCyan
Write-Host "    Vídeos     : $PathVideos"     -ForegroundColor DarkCyan
Write-Host ""

# --- Executa a limpeza de cada pasta ---
Clear-UserFolder -FolderPath $PathDownloads -FolderLabel "Downloads"
Clear-UserFolder -FolderPath $PathDocuments -FolderLabel "Documentos"
Clear-UserFolder -FolderPath $PathPictures  -FolderLabel "Fotos (Pictures)"
Clear-UserFolder -FolderPath $PathVideos    -FolderLabel "Vídeos"

# --- Pastas adicionais relacionadas que também acumulam conteúdo do usuário ---

# Música: raramente usada em laboratório, mas incluída para restauração total
$PathMusic = Resolve-FolderPath -ComPath $null -DefaultRelative "Music"
Clear-UserFolder -FolderPath $PathMusic -FolderLabel "Música (Music)"

# OneDrive local: se o aluno conectou o OneDrive, apaga os arquivos sincronizados localmente.
# ATENÇÃO: isso apaga apenas a cópia LOCAL. A nuvem não é afetada.
$OneDrivePaths = @(
    "$TargetProfile\OneDrive",
    "$TargetProfile\OneDrive - *"   # Cobre variantes corporativas (ex.: OneDrive - Empresa)
)
Write-Host "  -> OneDrive (cache local)" -ForegroundColor White
foreach ($odPath in $OneDrivePaths) {
    $resolvedPaths = Resolve-Path -Path $odPath -ErrorAction SilentlyContinue
    foreach ($rp in $resolvedPaths) {
        if (Test-Path $rp.Path) {
            Get-ChildItem -Path $rp.Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
                # Preserva as pastas de sistema do OneDrive (começam com ponto ou são ocultas do sistema)
                if (-not ($_.Attributes -band [System.IO.FileAttributes]::System)) {
                    Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "     [DEL] $($_.Name)" -ForegroundColor DarkGray
                }
            }
            Write-Host "     [OK] Cache local do OneDrive limpo: $($rp.Path)" -ForegroundColor Green
        }
    }
}

# Desconecta o OneDrive (remove a conta vinculada) via Registro
# Isso impede que o próximo aluno veja os arquivos em nuvem do anterior
$OneDriveRegPaths = @(
    "HKCU:\Software\Microsoft\OneDrive\Accounts",
    "HKCU:\Software\Microsoft\SkyDrive"   # chave legada
)
foreach ($regPath in $OneDriveRegPaths) {
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [DEL REG] OneDrive account key: $regPath" -ForegroundColor DarkGray
    }
}

Write-Host "  [OK] Pastas pessoais completamente restauradas." -ForegroundColor Green

# ==============================================================================
# BLOCO 7 — BLOCO DE NOTAS DO WINDOWS 11 (PACOTE UWP)
# O novo Bloco de Notas (Windows 11 22H2+) é um pacote UWP publicado na Store.
# Ele salva rascunhos e o estado das abas abertas em LocalAppData\Packages\*.
# Destruindo essa pasta, o app sempre abre em branco na próxima sessão.
# ==============================================================================
Write-Host ""
Write-Host "[ETAPA 7/7] Limpando cache do Bloco de Notas (Windows 11)..." -ForegroundColor Yellow

# O nome do pacote UWP do Notepad contém "Microsoft.WindowsNotepad"
$UWPPackagesRoot = "$TargetProfile\AppData\Local\Packages"

if (Test-Path $UWPPackagesRoot) {
    Get-ChildItem -Path $UWPPackagesRoot -Directory -Filter "*WindowsNotepad*" -ErrorAction SilentlyContinue |
        ForEach-Object {
            $notepadPkg = $_.FullName
            # Apaga apenas os dados do usuário (LocalState, Settings, RoamingState)
            # NÃO apaga o pacote inteiro (isso corromperia a instalação do app)
            $NotepadDataPaths = @(
                "$notepadPkg\LocalState",     # Rascunhos e abas salvas automaticamente
                "$notepadPkg\Settings",       # Configurações do usuário (tema, fonte)
                "$notepadPkg\RoamingState",   # Estado sincronizado
                "$notepadPkg\TempState"       # Arquivos temporários
            )
            Remove-ItemSafe -Paths $NotepadDataPaths
            Write-Host "  [RESET] $($_.Name)" -ForegroundColor DarkGray
        }
}

# Cobre também o Bloco de Notas clássico (Win32) que ainda existe no Win11
# Ele não tem estado persistente relevante, mas limpa arquivos temporários
Remove-ItemSafe -Paths @(
    "$TargetProfile\AppData\Roaming\Microsoft\Windows\Recent\*.txt"
)

Write-Host "  [OK] Bloco de Notas resetado." -ForegroundColor Green

# ==============================================================================
# BLOCO 8 — LIMPEZA COMPLEMENTAR (BOA PRÁTICA DE HIGIENE)
# Itens adicionais que garantem uma sessão ainda mais limpa.
# ==============================================================================
Write-Host ""
Write-Host "[COMPLEMENTAR] Higiene adicional do perfil..." -ForegroundColor Yellow

# Limpa arquivos temporários do usuário (%TEMP%)
$TempPaths = @(
    "$TargetProfile\AppData\Local\Temp",
    # Histórico de arquivos recentes do Windows Explorer
    "$TargetProfile\AppData\Roaming\Microsoft\Windows\Recent",
    # JumpLists (arquivos recentes ancorados na Barra de Tarefas)
    "$TargetProfile\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations",
    "$TargetProfile\AppData\Roaming\Microsoft\Windows\Recent\CustomDestinations",
    # Cache de thumbnails do Explorer
    "$TargetProfile\AppData\Local\Microsoft\Windows\Explorer",
    # Cache de aplicativos UWP em geral
    "$TargetProfile\AppData\Local\Microsoft\Windows\INetCache",
    # Histórico da Área de Transferência (Clipboard)
    "$TargetProfile\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\*.automaticDestinations-ms"
)
Remove-ItemSafe -Paths $TempPaths

# Limpa o histórico da Área de Transferência via Registro
try {
    # Desativa e reativa o histórico do clipboard, efetivamente limpando-o
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -Value 0 -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -Value 1 -ErrorAction SilentlyContinue
} catch {}

# Apaga chaves de registro de arquivos recentes do Office e Explorer
$RecentRegistryKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\File MRU",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\File MRU",
    "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\File MRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU",     # histórico do Win+R
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"  # diálogos Abrir/Salvar
)
foreach ($regPath in $RecentRegistryKeys) {
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [DEL REG] $regPath" -ForegroundColor DarkGray
    }
}

Write-Host "  [OK] Higiene complementar concluída." -ForegroundColor Green

# ==============================================================================
# FINALIZAÇÃO — RELATÓRIO E ENCERRAMENTO
# ==============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   RESET DE SESSÃO CONCLUÍDO COM SUCESSO!" -ForegroundColor Green
Write-Host "   Usuário resetado : $TargetUser" -ForegroundColor Cyan
Write-Host "   Data/Hora        : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "A máquina está pronta para o próximo aluno." -ForegroundColor White
Write-Host ""

# Opcional: Gravar um log de auditoria em pasta central de rede (ajuste o caminho)
# $LogPath = "\\SERVIDOR\Logs\LabReset"
# $LogFile = "$LogPath\reset_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
# if (Test-Path $LogPath) {
#     "Reset realizado em $(Get-Date) | Computador: $(hostname) | Usuário: $TargetUser" |
#         Out-File -FilePath $LogFile -Encoding UTF8
# }

# Aguarda 5 segundos antes de fechar a janela, para que o técnico possa ler o resultado.
Write-Host "Esta janela fechará automaticamente em 5 segundos..." -ForegroundColor DarkGray
Start-Sleep -Seconds 5
