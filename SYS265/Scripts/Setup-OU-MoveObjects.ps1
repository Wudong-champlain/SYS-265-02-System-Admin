<# 
.SYNOPSIS
  Interactive OU setup + move users/computers + optional delete old OU.

.NOTES
  Run as Domain Admin on a Domain Controller (or in Enter-PSSession to AD01).
#>

# --- Safety / prerequisites ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "ActiveDirectory module not available. Run this on AD01 (DC) or install RSAT AD tools."
    exit 1
}

# --- Helpers ---
function Get-DomainDN {
    (Get-ADDomain).DistinguishedName
}

function Ensure-OU {
    param(
        [Parameter(Mandatory=$true)][string]$OUName,
        [Parameter(Mandatory=$true)][string]$DomainDN
    )
    $ouDN = "OU=$OUName,$DomainDN"
    $existing = Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue

    if ($null -eq $existing) {
        Write-Host "[+] Creating OU: $OUName"
        New-ADOrganizationalUnit -Name $OUName -Path $DomainDN | Out-Null
    } else {
        Write-Host "[i] OU already exists: $OUName"
    }
    return $ouDN
}

function Show-Computers {
    param([string]$FilterText = "*")
    Get-ADComputer -Filter "Name -like '$FilterText'" |
        Select-Object Name, DistinguishedName |
        Sort-Object Name
}

function Show-Users {
    param([string]$FilterText = "*")
    Get-ADUser -Filter "Name -like '$FilterText'" |
        Select-Object Name, SamAccountName, DistinguishedName |
        Sort-Object Name
}

function Move-ComputersToOU {
    param(
        [Parameter(Mandatory=$true)][string[]]$ComputerNames,
        [Parameter(Mandatory=$true)][string]$TargetOUDN
    )
    foreach ($c in $ComputerNames) {
        $cTrim = $c.Trim()
        if ([string]::IsNullOrWhiteSpace($cTrim)) { continue }

        $obj = Get-ADComputer -Identity $cTrim -ErrorAction SilentlyContinue
        if ($null -eq $obj) {
            Write-Warning "[-] Computer not found: $cTrim"
            continue
        }

        Move-ADObject -Identity $obj.DistinguishedName -TargetPath $TargetOUDN
        Write-Host "[+] Moved computer '$($obj.Name)' to $TargetOUDN"
    }
}

function Move-UsersToOU {
    param(
        [Parameter(Mandatory=$true)][string[]]$UserSamAccountNames,
        [Parameter(Mandatory=$true)][string]$TargetOUDN
    )
    foreach ($u in $UserSamAccountNames) {
        $uTrim = $u.Trim()
        if ([string]::IsNullOrWhiteSpace($uTrim)) { continue }

        $obj = Get-ADUser -Identity $uTrim -ErrorAction SilentlyContinue
        if ($null -eq $obj) {
            Write-Warning "[-] User not found (SamAccountName): $uTrim"
            continue
        }

        Move-ADObject -Identity $obj.DistinguishedName -TargetPath $TargetOUDN
        Write-Host "[+] Moved user '$($obj.SamAccountName)' to $TargetOUDN"
    }
}

function Delete-OUIfRequested {
    param(
        [Parameter(Mandatory=$true)][string]$OUName,
        [Parameter(Mandatory=$true)][string]$DomainDN
    )
    $ouDN = "OU=$OUName,$DomainDN"
    $existing = Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Write-Host "[i] OU not found, nothing to delete: $OUName"
        return
    }

    Write-Host "[!] Deleting OU: $OUName"
    # remove accidental deletion protection if set
    try {
        Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
        Write-Host "[+] Removed accidental deletion protection."
    } catch {
        Write-Warning "[!] Could not change ProtectedFromAccidentalDeletion (may already be false)."
    }

    try {
        Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        Write-Host "[+] Deleted OU '$OUName'."
    } catch {
        Write-Error "Failed to delete OU '$OUName': $($_.Exception.Message)"
    }
}

# --- Main ---
$domainDN = Get-DomainDN
Write-Host "Domain DN: $domainDN"
Write-Host ""

# Ask for OU name to create/use
$targetOUName = Read-Host "Enter the OU name to create/use (example: Software Deploy)"
if ([string]::IsNullOrWhiteSpace($targetOUName)) {
    Write-Error "OU name cannot be blank."
    exit 1
}

$targetOUDN = Ensure-OU -OUName $targetOUName -DomainDN $domainDN
Write-Host ""

# Optionally show available computers
$showComps = Read-Host "Show available computers before choosing? (Y/N)"
if ($showComps.Trim().ToUpper() -eq "Y") {
    $compFilter = Read-Host "Computer name filter (example: WKS*). Leave blank for all"
    if ([string]::IsNullOrWhiteSpace($compFilter)) { $compFilter = "*" }
    Show-Computers -FilterText $compFilter | Format-Table -AutoSize
    Write-Host ""
}

# Choose computers to move
$compInput = Read-Host "Enter computer NAME(s) to move (comma-separated). Example: WKS01-WU"
$computerNames = @()
if (-not [string]::IsNullOrWhiteSpace($compInput)) {
    $computerNames = $compInput.Split(",")
    Move-ComputersToOU -ComputerNames $computerNames -TargetOUDN $targetOUDN
}
Write-Host ""

# Optionally show available users
$showUsers = Read-Host "Show available users before choosing? (Y/N)"
if ($showUsers.Trim().ToUpper() -eq "Y") {
    $userFilter = Read-Host "User display-name filter (example: wu*). Leave blank for all"
    if ([string]::IsNullOrWhiteSpace($userFilter)) { $userFilter = "*" }
    Show-Users -FilterText $userFilter | Format-Table -AutoSize
    Write-Host ""
}

# Choose users to move (SamAccountName)
Write-Host "NOTE: Enter user SamAccountName(s) (ex: wu, wu.dong) not the display name."
$userInput = Read-Host "Enter user SamAccountName(s) to move (comma-separated)"
$userSam = @()
if (-not [string]::IsNullOrWhiteSpace($userInput)) {
    $userSam = $userInput.Split(",")
    Move-UsersToOU -UserSamAccountNames $userSam -TargetOUDN $targetOUDN
}
Write-Host ""

# Optional delete old OU (like Test OU)
$delOld = Read-Host "Do you want to delete an old OU (example: Test OU)? (Y/N)"
if ($delOld.Trim().ToUpper() -eq "Y") {
    $oldOUName = Read-Host "Enter the OU name to delete (exact OU name)"
    if (-not [string]::IsNullOrWhiteSpace($oldOUName)) {
        Delete-OUIfRequested -OUName $oldOUName -DomainDN $domainDN
    }
}
Write-Host ""

# Final verification
Write-Host "=== Objects currently in OU '$targetOUName' ==="
Get-ADObject -Filter * -SearchBase $targetOUDN |
    Select-Object Name, ObjectClass |
    Sort-Object ObjectClass, Name |
    Format-Table -AutoSize

Write-Host ""
Write-Host "[Done]"