# Automação PowerShell — Laboratório de Informática

Coleção de scripts PowerShell para gerenciamento de máquinas em laboratório.

---

## Índice

- [Script Desloga Usuarios](#apaga-registros--script-desloga-usuariostxt)
- [Script Atualiza Imagem](#atualizar-imagem--script-atualiza-imagemtxt)
- [Deploy-LabStartup](#implementação--deploy-labstartupps1)
- [Lab-Cleanup-Wrapper](#implementação--lab-cleanup-wrapperps1)
- [Reset-LabSession (Admin)](#resetlab--reset-labsession---admps1)
- [Reset de Sessão (Usuário)](#resetlab--script-de-reset-de-sessão-para-usuáriotxt)
- [Setup-LabTask](#implementação--setup-labtaskps1)

---

## Scripts

### `Apaga Registros / Script Desloga Usuarios.txt`
Higienização expressa da sessão: encerra processos (navegadores, Office, IDEs, etc.), limpa credenciais do Windows Vault, apaga caches de navegadores e apps, esvazia a Área de Trabalho, limpa `%TEMP%` e recarrega o DNS.

---

### `Atualizar Imagem / Script Atualiza Imagem.txt`
Aplica wallpaper e tela de bloqueio corporativos em toda a máquina. Baixa as imagens do GitHub (ou usa cópia local), injeta as políticas no registro de todos os perfis de usuário via `NTUSER.DAT`, bloqueia Microsoft Store e Widgets, remove jogos/bloatware e configura senhas para nunca expirarem.

Executa Script altera Imagem.
 - `iex(irm is.gd/VWgMgD)`

---

### `Implementação / Deploy-LabStartup.ps1`
Script de implantação executado uma única vez pelo Administrador. Cria a pasta segura `C:\Windows\Scripts\`, copia o wrapper, aplica permissões restritivas e registra a Tarefa Agendada `LabHigienizacao_Startup` (executa como SYSTEM a cada logon de qualquer usuário).

---

### `Implementação / Lab-Cleanup-Wrapper.ps1`
Executor oculto acionado pela Tarefa Agendada a cada logon. Detecta o usuário interativo via WMI, resolve o perfil correto, baixa o script de higienização do GitHub e o executa no contexto do usuário logado. Gera log em `C:\Windows\Logs\LabCleanup\`.

---

### `ResetLab / Reset-LabSession - Adm.ps1`
Reset completo de sessão (modo Administrador). Encerra processos, limpa navegadores (Chrome, Edge, Firefox, Opera), remove tokens Microsoft (Office, Power BI, TokenBroker, Credential Manager), apaga dados de VS Code, Git, Packet Tracer, esvazia Downloads, Documentos, Fotos, Vídeos, Música, OneDrive local e reseta o Bloco de Notas do Windows 11.

---

### `ResetLab / Script de Reset de Sessão para Usuário.txt`
Mesma cobertura do script Admin, porém executável pelo próprio aluno sem privilégios elevados. Limpa o perfil do usuário atual usando variáveis de ambiente do próprio contexto.

---

### `Implementação / Setup-LabTask.ps1`
Script de configuração inicial executado pelo Administrador. Realiza o download automático do script `Reset-LabSession - Adm.ps1` do GitHub para `C:\Arquivo de Programas\` e registra a Tarefa Agendada `ResetLabSession` (executa como SYSTEM no logon do usuário **aluno**). Inclui validação de privilégios, verificação de integridade do arquivo baixado e confirmação final do ambiente configurado.
