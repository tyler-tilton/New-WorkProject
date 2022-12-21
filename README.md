# New-WorkProject

A PowerShell function for creating a new project directory and generating several files and subdirectories within it.

## Parameters

- `ProjectName`: a string value representing the name of the project (required)
- `ProjectDirectory`: a string value representing the directory path where the project will be created (optional, default is `$env:USERPROFILE\Documents\$ProjectName`)
- `PowerShell`, `Readme`, and `WorkLog`: switch values indicating whether to create a PowerShell file, README file, and worklog file, respectively (optional, default is to create all three if none are specified)
- `TicketNumber`: a string value representing a ticket number associated with the project (optional)
- `TaskLink`: a string value representing a link to a task associated with the project (optional)

## Examples

To create a new project named "MyProject" with all default options:

```POWERSHELL
New-WorkProject -ProjectName "MyProject"
```

To create a new project named "MyProject" with a custom project directory and only a PowerShell file:

 ```POWERSHELL
New-WorkProject -ProjectName "MyProject" -ProjectDirectory "C:\Projects\MyProject" -PowerShell -Readme:$false -WorkLog:$false
 ```

To create a new project named "MyProject" with a custom ticket number and task link:

```POWERSHELL
New-WorkProject -ProjectName "MyProject" -TicketNumber "ABC-123" -TaskLink "https://example.com/tasks/123"
 ```
