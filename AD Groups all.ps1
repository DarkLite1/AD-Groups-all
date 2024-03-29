﻿#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.EventLog, Toolbox.HTML

<#
    .SYNOPSIS
        Create a list of AD groups within a specific OU.

    .DESCRIPTION
        Create an Excel sheet of all active directory groups found within a
        specific organizational unit. This list contains the group name,
        OU, member count, ... but not the members itself.

        The report is sent by mail for further review.

    .PARAMETER ImportFile
        A .json file containing the script arguments.

    .PARAMETER LogFolder
        Location for the log files.
#>

Param (
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Parameter(Mandatory)]
    [String]$ScriptName = 'AD Groups all (BNL)',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Groups all\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($OUs = $File.AD.OU)) {
            throw "Input file '$ImportFile': No 'AD.OU' found."
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $adGroups = foreach ($ou in $OUs) {
            Get-ADGroup -SearchBase $ou -Filter * -EA Stop -Properties Created,
            Modified, Members, Description, CanonicalName, SamAccountName,
            GroupCategory, GroupScope, info, ManagedBy |
            Select-Object -Property Created, Modified, Description,
            SamAccountName, DisplayName, Name, GroupCategory, GroupScope, info,
            @{
                Name       = 'OU'
                Expression = { ConvertTo-OuNameHC $_.CanonicalName }
            },
            @{
                Name       = 'ManagedBy'
                Expression = {
                    if ($_.ManagedBy) { Get-ADDisplayNameHC $_.ManagedBy }
                }
            },
            @{
                Name       = 'Members'
                Expression = { $_.Members.Count }
            }
        }

        $excelParams = @{
            Path               = $logFile + ' - Result.xlsx'
            AutoSize           = $true
            BoldTopRow         = $true
            FreezeTopRow       = $true
            WorkSheetName      = 'Users'
            TableName          = 'Users'
            NoNumberConversion = @(
                'Employee ID', 'OfficePhone', 'HomePhone', 'MobilePhone', 'ipPhone', 'Fax', 'Pager'
            )
            ErrorAction        = 'Stop'
        }
        Remove-Item $excelParams.Path -Force -EA Ignore
        $adGroups | Export-Excel @excelParams

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$($adGroups.Count) groups"
            LogFolder = $logParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        #region Format HTML
        $mailParams.Attachments = $excelParams.Path

        $mailParams.Message = "A total of <b>$(@($adGroups).count) groups</b> have been found. <p><i>* Check the attachment for details </i></p>
            $($OUs | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:')"
        #endregion

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    finally {
        Write-EventLog @EventEndParams
    }
}