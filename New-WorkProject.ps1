<#
.SYNOPSIS
Creates a new project directory with optional components, including a README file, worklog file, an Outlook/To Do task, and PowerShell script file.

.DESCRIPTION
The New-WorkProject function allows you to easily create a new project directory, README file, worklog file, Outlook/To Do task, and PowerShell script file. The function also includes the ability to specify a ticket number and task link for the project, which will be included in the README file and allows you to specify whether to create the README file, worklog file, and PowerShell script file. If none of these options are specified, all four components will be created by default.

.PARAMETER ProjectName
Specifies the name of the project. This parameter is required.

.PARAMETER ProjectDirectory
Specifies the directory where the project should be created. If this parameter is not specified, the project will be created in the user's Documents directory.

.PARAMETER PowerShell
Specifies whether to create a PowerShell script file for the project. If this parameter is not specified, a PowerShell script file will not be created.

.PARAMETER Readme
Specifies whether to create a README file for the project. If this parameter is not specified, a README file will not be created.

.PARAMETER WorkLog
Specifies whether to create a worklog file for the project. If this parameter is not specified, a worklog file will not be created.

.PARAMETER CreateTask
Specifies whether to create an Outlook/To Do item for the project. If this parameter is not specified, a task will not be created.

.PARAMETER TicketNumber
Specifies the ticket number for the project. If this parameter is not specified, the ticket number will not be included in the README file.

.PARAMETER TaskLink
Specifies the link to the task for the project. If this parameter is not specified, the task link will not be included in the README file.

.EXAMPLE
New-WorkProject -ProjectName "My Project" -ProjectDirectory "C:\Projects"  -TicketNumber "12345" -TaskLink "https://tasks.com/mytask"

Creates a new work project with a project directory, README file, worklog file, and PowerShell script file in the specified project directory, with the specified ticket number and task link included in the README file.

.EXAMPLE
New-WorkProject -ProjectName "My Project" -Readme -WorkLog

Creates a new work project with a project directory, README file, and worklog file in the user's Documents directory. No PowerShell script file or ticket number/task link will be included.

.NOTES
This function uses Visual Studio Code and Git, which must be installed on the system for the function to work correctly when PowerShell is specified.
#>

function New-WorkProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectName,
        
        [Parameter(Mandatory = $false)]
        [string]$ProjectDirectory,
        
        [Parameter(Mandatory = $false)]
        [switch]$PowerShell,
        
        [Parameter(Mandatory = $false)]
        [switch]$Readme,
        
        [Parameter(Mandatory = $false)]
        [switch]$WorkLog,

        [Parameter(Mandatory = $false)]
        [switch]$CreateTask,
        
        [Parameter(Mandatory = $false)]
        [string]$TicketNumber,
        
        [Parameter(Mandatory = $false)]
        [string]$TaskLink
    )

    if ((-not $PowerShell) -and (-not $Readme) -and (-not $WorkLog) -and (-not $CreateTask)) {
        $PowerShell = $TRUE
        $Readme = $TRUE
        $WorkLog = $TRUE
        $CreateTask = $TRUE
    }
    
    # Set the default project directory
    if (-not $ProjectDirectory) {
        $ProjectDirectory = "$env:USERPROFILE\Documents\$ProjectName"
    }
    
    if (Test-Path -Path $ProjectDirectory) {
        Write-Error "The Project Directory already exists. Please try again."
        exit
    }

    # Create the project directory
    
    New-Item -ItemType Directory -Path $ProjectDirectory | Out-Null

    # Create the worklog.md file if specified
    if ($WorkLog) {
        $worklogFile = "$projectDirectory\worklog.md"
        New-Item -ItemType File -Path $worklogFile | Out-Null

        # Add the main heading and date as the first subheading to the worklog.md file
        $currentDate = Get-Date -Format "yyyy-MM-dd"
        $worklogContent = @"
# $ProjectName

## $currentDate

"@
        Add-Content -Path $worklogFile -Value $worklogContent -NoNewline
    }

    # Create the Outlook/To Do task if specified
    if ($CreateTask) {

        $outlook = New-Object -ComObject Outlook.Application 
        $taskList = $outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderTasks) 
        $task = $taskList.Items.Add() 
        $task.Subject = "$ProjectName"
        $task.Save()
    }

    # Create the README.MD file if specified
    if ($Readme) {
        $readmeFile = "$ProjectDirectory\README.MD"
        New-Item -ItemType File -Path $readmeFile | Out-Null
        
        # Add the general information to the README.MD file
        $justification = ""
        if ($TicketNumber) {
            $justification += "Ticket Number: $TicketNumber"
        }
        if ($TaskLink) {
            if ($justification) {
                $justification += " - "
            }
            $justification += "Task Link: <$TaskLink>"
        }
        $generalInformation = @"
# $ProjectName

## Purpose

## Author

$env:USERNAME

## Published

## Updated

## Justification

$justification

## Description

## Dependencies

## Change Log

## Usage
"@
        Add-Content -Path $readmeFile -Value $generalInformation
    }
    
    # Create the Includes sub-directory if specified
    if ($PowerShell) {
        $projectFile = "$ProjectDirectory\$ProjectName.ps1"
        New-Item -ItemType File -Path $projectFile | Out-Null
        
        # Add the dynamic includes to the PowerShell file
        $includes = @'
# Dynamic Includes
(Get-ChildItem -Path $PSScriptRoot\Includes).Foreach({
        . $_.FullName
    })
'@
        Add-Content -Path $projectFile -Value $includes

        New-Item -ItemType Directory -Path "$ProjectDirectory\Includes" | Out-Null

        # Initialize the project directory as a git repository
        git -C $projectDirectory init
    }

    # Open the project directory in Visual Studio Code
    code $projectDirectory
}