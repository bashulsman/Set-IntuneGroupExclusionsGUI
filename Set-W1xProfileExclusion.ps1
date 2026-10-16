<#
.SYNOPSIS
    Voegt een exclude-groep toe aan alle Intune Configuration Profiles waarvan de naam begint met W1x.

.DESCRIPTION
    Doorloopt de drie soorten Configuration Profiles in Intune:
      - Templates            (deviceManagement/deviceConfigurations)
      - Settings catalog     (deviceManagement/configurationPolicies)
      - Administrative templ.(deviceManagement/groupPolicyConfigurations)

    Voor elk profiel waarvan de naam matcht met -NamePattern wordt de groep als
    exclusion toegevoegd. Bestaande assignments (incl. filters) blijven behouden,
    want de Graph /assign-actie vervangt altijd de volledige set assignments.

    Profielen waar de groep al aan gekoppeld is (include of exclude) worden overgeslagen.

.PARAMETER GroupName
    Weergavenaam van de Entra ID-groep die ge-exclude moet worden.

.PARAMETER NamePattern
    Regex waarop de profielnaam moet matchen (case-insensitive). Standaard '^W1x\s*-'
    (dus "W1x - ..."). Alleen CUS-profielen: '^W1x\s*-.*- CUS -'

.EXAMPLE
    .\Set-W1xProfileExclusion.ps1 -WhatIf
    Toont welke profielen aangepast zouden worden, zonder iets te wijzigen.

.EXAMPLE
    .\Set-W1xProfileExclusion.ps1
    Voert de wijziging daadwerkelijk uit.

.NOTES
    Vereist module Microsoft.Graph.Authentication:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$GroupName   = 'AAD_UG_Update-Ring-00',
    [string]$NamePattern = '^W1x\s*-'
)

$ErrorActionPreference = 'Stop'

#region Verbinden
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'DeviceManagementConfiguration.ReadWrite.All', 'Group.Read.All' -NoWelcome
}
#endregion

#region Helpers
function Get-GraphAll {
    param([string]$Uri)
    $items = @()
    do {
        $resp   = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $items += $resp.value
        $Uri    = $resp.'@odata.nextLink'
    } while ($Uri)
    return $items
}
#endregion

#region Groep opzoeken
$escapedName = $GroupName.Replace("'", "''")
$groups = Get-GraphAll -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escapedName'&`$select=id,displayName"

if ($groups.Count -ne 1) {
    throw "Verwacht precies 1 groep met naam '$GroupName', gevonden: $($groups.Count)."
}
$groupId = $groups[0].id
Write-Host "Groep gevonden: $GroupName ($groupId)" -ForegroundColor Cyan
#endregion

#region Profieltypes
$base = 'https://graph.microsoft.com/beta/deviceManagement'
$profileTypes = @(
    @{ Type = 'Template';          Endpoint = 'deviceConfigurations';      NameProp = 'displayName' }
    @{ Type = 'Settings catalog';  Endpoint = 'configurationPolicies';     NameProp = 'name' }
    @{ Type = 'Admin templates';   Endpoint = 'groupPolicyConfigurations'; NameProp = 'displayName' }
)
#endregion

$results = [System.Collections.Generic.List[object]]::new()

foreach ($pt in $profileTypes) {
    Write-Host "`n=== $($pt.Type) ===" -ForegroundColor Cyan

    $profiles = Get-GraphAll -Uri "$base/$($pt.Endpoint)?`$select=id,$($pt.NameProp)" |
        Where-Object { $_[$pt.NameProp] -match $NamePattern }

    if (-not $profiles) {
        Write-Host "Geen profielen gevonden die matchen met '$NamePattern'."
        continue
    }

    foreach ($p in $profiles) {
        $name = $p[$pt.NameProp]
        $id   = $p.id

        $assignments = Get-GraphAll -Uri "$base/$($pt.Endpoint)/$id/assignments"

        # Al gekoppeld? (include of exclude)
        $existing = $assignments | Where-Object { $_.target.groupId -eq $groupId }
        if ($existing) {
            $kind = if ($existing.target.'@odata.type' -like '*exclusion*') { 'al ge-exclude' } else { 'INCLUDED (niet aangepast!)' }
            $color = if ($kind -like 'INCLUDED*') { 'Yellow' } else { 'DarkGray' }
            Write-Host "  [skip] $name - groep $kind" -ForegroundColor $color
            $results.Add([pscustomobject]@{ Type = $pt.Type; Profiel = $name; Status = "Overgeslagen: $kind" })
            continue
        }

        # Bestaande assignments overnemen (alleen target, read-only velden weglaten)
        $newAssignments = @(
            foreach ($a in $assignments) { @{ target = $a.target } }
        )
        $newAssignments += @{
            target = @{
                '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'
                groupId       = $groupId
            }
        }

        $body = @{ assignments = $newAssignments } | ConvertTo-Json -Depth 10

        if ($PSCmdlet.ShouldProcess("$($pt.Type): $name", "Exclude groep '$GroupName'")) {
            try {
                Invoke-MgGraphRequest -Method POST -Uri "$base/$($pt.Endpoint)/$id/assign" -Body $body -ContentType 'application/json' | Out-Null
                Write-Host "  [ok]   $name" -ForegroundColor Green
                $results.Add([pscustomobject]@{ Type = $pt.Type; Profiel = $name; Status = 'Exclusion toegevoegd' })
            }
            catch {
                Write-Host "  [fout] $name - $($_.Exception.Message)" -ForegroundColor Red
                $results.Add([pscustomobject]@{ Type = $pt.Type; Profiel = $name; Status = "Fout: $($_.Exception.Message)" })
            }
        }
        else {
            $results.Add([pscustomobject]@{ Type = $pt.Type; Profiel = $name; Status = 'WhatIf: zou exclusion toevoegen' })
        }
    }
}

Write-Host "`n=== Samenvatting ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize
