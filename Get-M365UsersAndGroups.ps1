# Script para extrair usuários e grupos do Microsoft 365
# Requer o módulo Microsoft.Graph.Users e Microsoft.Graph.Groups

# Switch interno: usado quando o script se reinicia numa sessão limpa após instalar módulos.
# NÃO usar manualmente na primeira execução.
param(
    [switch]$SkipModuleCheck
)

# Módulos necessários
$requiredModules = @(
    "Microsoft.Graph.Authentication",  # Connect-MgGraph
    "Microsoft.Graph.Users",           # Get-MgUser / Get-MgUserMemberOf
    "Microsoft.Graph.Groups",          # Get-MgGroup
    "ImportExcel"                      # Export-Excel (.xlsx)
)

# Validação e instalação automática dos módulos faltantes.
# IMPORTANTE: no Windows PowerShell 5.1, instalar e importar os módulos do Microsoft.Graph
# na MESMA sessão causa "TypeLoadException: GetTokenAsync ... não tem uma implementação"
# (conflito de assembly em uso). Por isso, se algo for instalado, reiniciamos o script
# numa sessão nova e limpa antes de importar/usar os módulos.
if (-not $SkipModuleCheck) {
    $installedAny = $false
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Módulo '$module' não encontrado. Instalando..." -ForegroundColor Cyan
            try {
                Install-Module $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $installedAny = $true
                Write-Host "Módulo '$module' instalado com sucesso." -ForegroundColor Green
            }
            catch {
                Write-Host "Falha ao instalar o módulo '$module': $($_.Exception.Message)" -ForegroundColor Red
                exit 1
            }
        }
    }

    if ($installedAny) {
        Write-Host "Módulos recém-instalados. Reiniciando em uma sessão limpa para evitar conflitos de assembly..." -ForegroundColor Yellow
        $psExe = (Get-Process -Id $PID).Path   # mesmo host (powershell.exe)
        & $psExe -NoProfile -ExecutionPolicy Bypass -File "$PSCommandPath" -SkipModuleCheck
        exit $LASTEXITCODE
    }
}

# Importa os módulos na sessão atual (já livre do churn de instalação)
foreach ($module in $requiredModules) {
    try {
        Import-Module $module -ErrorAction Stop
    }
    catch {
        Write-Host "Falha ao importar o módulo '$module': $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Feche esta janela do PowerShell, abra uma nova e execute o script novamente." -ForegroundColor Yellow
        exit 1
    }
}

# Conectar ao Microsoft Graph
# Scopes necessários: User.Read.All, Group.Read.All, Directory.Read.All
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All"

# Obter todos os usuários
Write-Host "Obtendo usuários..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName

$results = @()

foreach ($user in $users) {
    # Obter grupos do usuário
    # Get-MgUserMemberOf retorna objetos de diretório, filtramos por grupos
    $groups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" }
    
    $groupNames = @()
    foreach ($group in $groups) {
        # Em algumas versões do SDK, o objeto retornado pode não ter o DisplayName diretamente acessível se for um DirectoryObject genérico
        # Mas geralmente para grupos ele vem. Se não vier, teríamos que fazer Get-MgGroup -GroupId $group.Id
        
        # Tentativa de pegar o DisplayName direto (funciona na maioria das versões recentes do SDK para MemberOf)
        if ($group.AdditionalProperties.ContainsKey("displayName")) {
            $groupNames += $group.AdditionalProperties["displayName"]
        }
        else {
            # Fallback: buscar o grupo individualmente (mais lento, mas garantido)
            $fullGroup = Get-MgGroup -GroupId $group.Id
            $groupNames += $fullGroup.DisplayName
        }
    }
    
    # Formatar a saída
    $groupsString = $groupNames -join ", "
    
    # Exibir no console no formato solicitado
    if ($groupNames.Count -gt 0) {
        Write-Host "$($user.UserPrincipalName) -> $groupsString"
        
        # Adicionar ao objeto para exportação APENAS se tiver grupos
        $results += [PSCustomObject]@{
            User        = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            Groups      = $groupsString
        }
    }
    else {
        # Opcional: Comentar se quiser ver quem não tem grupos no console, mas não exportar
        # Write-Host "$($user.UserPrincipalName) -> (sem grupos)"
    }
}

# Exportar para Excel (.xlsx)
$excelPath = ".\M365UsersAndGroups.xlsx"
# Remove arquivo existente para evitar erros de append indesejados ou bloqueios
if (Test-Path $excelPath) { Remove-Item $excelPath -Force }

$results | Export-Excel -Path $excelPath -WorksheetName "UsuariosGrupos" -AutoSize -TableStyle Medium2 -FreezeTopRow
Write-Host "Exportação concluída: $excelPath" -ForegroundColor Green
