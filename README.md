# M365 Users and Groups Exporter

Script em PowerShell projetado para extrair a relação de usuários do Microsoft 365 e os grupos aos quais pertencem, utilizando o Microsoft Graph API. O resultado é exportado automaticamente para um arquivo Excel (.xlsx) estruturado.

## Pré-requisitos

* PowerShell 5.1 ou superior (PowerShell 7 recomendado).
* Módulo `Microsoft.Graph` (instalação manual exigida caso não esteja presente).
* Módulo `ImportExcel` (o script tentará instalar automaticamente no escopo do usuário corrente, caso ausente).
* Conta com privilégios administrativos no Microsoft 365 para consentimento dos escopos exigidos pelo Graph API (`User.Read.All`, `Group.Read.All`, `Directory.Read.All`).

## Como Executar

1. Abra o terminal do PowerShell. (Privilégios de Administrador podem ser necessários na primeira execução para a instalação dos módulos).
2. Navegue até o diretório onde o arquivo se encontra.
3. Execute o script:
   ```powershell
   .\Get-M365UsersAndGroups.ps1
