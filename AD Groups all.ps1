#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.Remoting

<#
    .SYNOPSIS
        Retrieve all AD groups within a specific OU.

    .DESCRIPTION
        Report all the groups within a specific OU in AD.. The import file is
        read for getting the correct parameters. Then the groups are collected
        and a mail is send to the end user with an Excel sheet in attachment
        containing the groups.

    .PARAMETER ImportFile
        Contains all the OU's where we need to search

    .PARAMETER LogFolder
        Location for the log files
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
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        $File = Get-Content $ImportFile -EA Stop | Remove-CommentsHC

        if (-not (
                $MailTo = $File | Get-ValueFromArrayHC MailTo -Delimiter ',')
        ) {
            throw "No 'MailTo' found in the input file."
        }

        if (-not ($OUs = $File | Get-ValueFromArrayHC -Exclude MailTo)) {
            throw "No organizational units found in the input file."
        }

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
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority High -Message $_ -Header $ScriptName
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
                Expression = { @($_.Members).Count }
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
            Subject   = "$(@($adGroups).count) groups"
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
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}