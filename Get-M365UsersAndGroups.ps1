# Script para extrair usuários e grupos do Microsoft 365
# Requer o módulo Microsoft.Graph.Users e Microsoft.Graph.Groups

# Instalar o módulo se necessário (descomente a linha abaixo se não tiver o módulo instalado)
# Install-Module Microsoft.Graph -Scope CurrentUser

# Instalar o módulo ImportExcel se necessário (para gerar .xlsx)
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Instalando módulo ImportExcel..." -ForegroundColor Cyan
    Install-Module ImportExcel -Scope CurrentUser -Force
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
