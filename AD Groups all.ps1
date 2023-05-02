#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.Remoting

<# 
    .SYNOPSIS   
        Retrieve all AD groups within a specific OU.

    .DESCRIPTION
        Report all the groups within a specific OU in AD.. The import file is read 
        for getting the correct parameters. Then the groups are collected and a mail is send to 
        the end user with an Excel sheet in attachment containing the groups.

    .PARAMETER ImportFile
        Contains all the OU's where we need to search

    .PARAMETER LogFolder
        Location for the log files

    .NOTES
        CHANGELOG
        2018/06/14 Script born

        AUTHOR Brecht.Gijbels@heidelbergcement.com #>

Param (
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Parameter(Mandatory)]
    [String]$ScriptName = 'AD Groups all (BNL)',    
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Groups all\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        $File = Get-Content $ImportFile -EA Stop | Remove-CommentsHC

        if (-not ($MailTo = $File | Get-ValueFromArrayHC MailTo -Delimiter ',')) {
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
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $Groups = foreach ($O in $OUs) {
            Get-ADGroup -SearchBase $O -Filter * -EA Stop -Properties Created, Modified, Members,
            Description, CanonicalName, SamAccountName, GroupCategory, GroupScope, info, ManagedBy | 
            Select-Object -Property Created, Modified, Description, SamAccountName, 
            DisplayName, Name, GroupCategory, GroupScope, info,
            @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } },
            @{N = 'ManagedBy'; E = { if ($_.ManagedBy) { Get-ADDisplayNameHC $_.ManagedBy } } },
            @{N = 'Members'; E = { @($_.Members).Count } }
        }
        
        $ExcelParams = @{
            Path               = $LogFile + ' - Result.xlsx'
            AutoSize           = $true
            BoldTopRow         = $true
            FreezeTopRow       = $true
            WorkSheetName      = 'Users'
            TableName          = 'Users'
            NoNumberConversion = 'Employee ID', 'OfficePhone', 'HomePhone', 'MobilePhone', 'ipPhone', 'Fax', 'Pager'
            ErrorAction        = 'Stop'
        }
        Remove-Item $ExcelParams.Path -Force -EA Ignore
        $Groups | Export-Excel @ExcelParams

        $MailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$(@($Groups).count) groups"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        #region Format HTML
        $MailParams.Attachments = $ExcelParams.Path

        $MailParams.Message = "A total of <b>$(@($Groups).count) groups</b> have been found. <p><i>* Check the attachment for details </i></p>
            $($OUs | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:')"
        #endregion

        $null = Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message ($env:USERNAME + ' - ' + "FAILURE:`n`n- " + $_)
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}