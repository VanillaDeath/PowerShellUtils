# Local Administrator Password Lister
# v1.0, by Wilson
# Use  Reset-AdmPwdPassword -ComputerName Computer-Name-Here  to reset a local admin password

param (
    [string]$OrgUnit,
    [string]$ComputerName,
    [string]$Domain = "contoso.com"
)
$DC = "DC=" + ("$Domain" -replace "\.",",DC=")

$OUs = @(
    "Servers",
    "Workstations"
)

$scriptname = Split-Path $PSCommandPath -Leaf
$laname = "Administrator"

Import-Module AdmPwd.PS

function DoQuit {
    Clear-Host
    exit
}

function Show-Menu ($options) {
    ""
    "Select OU to show:"
    ""
    $i = 1
    ForEach ($option in $options) {
        "[" + $i.ToString().PadLeft(2) + "] $option"
        $i++
    }
    "[ 0] ALL"
    ""
    "     Or just type a computer name"
    ""
    "[<┛] Quit & Clear Screen (Just hit Enter with no input)"
    ""
}

function Get-Pass ($CN) {
    "CN=$CN,$DC"
    ""
    Get-AdmPwdPassword -ComputerName "$CN"
    ""
}

function Get-Passes ($OU) {
    if ($OU -eq "" -or $OU -eq "*") {
        "Listing All preset Organizationl Units in $DC" + "..."
        ""
        ForEach ($O in $OUs) {
            Get-Passes ($O)
            ""
        }
    } else {
        $SearchBase = "OU=$OU,$DC"
        "$SearchBase"
        ""
        Get-ADComputer -Filter * -SearchBase "$SearchBase" | Get-AdmPwdPassword -ComputerName {$_.Name}
        ""
    }
}

""
"-- Local Administrator Password Lister -- v1.0 -- by Wilson --"
" - DOMAIN: $Domain ($DC) - LOCAL ADMIN: $laname -"
""

if ($OrgUnit) {
    Get-Passes ($OrgUnit)
    pause
} elseif ($ComputerName) {
    Get-Pass ($ComputerName)
    pause
} else {
    ""
    "Command-line usage:"
    ".\$scriptname [-OrgUnit OU-Name] [-ComputerName Computer-Name] [-Domain fqdn.tld]"
    "Use  -OrgUnit *  to show all."
    ""

    do {
        Show-Menu ($OUs)
        $in = Read-Host "Make Selection"
        $in = $in.Trim()
        ""

        switch ($in) {
            {$_ -match "^\d+$" -and [convert]::ToInt32($_, 10) -ge 1 -and [convert]::ToInt32($_, 10) -le $OUs.Count} {
                Get-Passes ($OUs[$_ - 1])
                pause
                break
            }
            "0" {
                Get-Passes ("*")
                pause
                break
            }
            "" {
                break
            }
            default {
                Get-Pass ($_)
            }
        }
    } until ($in -eq "")
}

DoQuit