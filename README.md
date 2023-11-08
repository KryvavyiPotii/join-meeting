# Join-Meeting
Join-Meeting is a Powershell module for fast and comfortable connection to KPI classes.

## Installation

Put the module file into directory `C:\Users\<username>\Documents\WindowsPowerShell\Modules\Join-Meeting`;NBSP
If directory does not exist, create it manually.;NBSP
Put subjects.json wherever you want and specify it's path in variable `$path` in module file.;NBSP

```powershell
Import-Module Join-Meeting
```

## Update

In order to update module you should change the file and either remove it and import again or import it with `-Force` flag

```powershell
Remove-Module Join-Meeting
Import-Module Join-Meeting
# OR
Import-Module Join-Meeting -Force
```

Also you can manually add new subject full titles, subject short titles, links and extra info into `subjects.json`.

## Usage

```powershell
# Join RE Lecture.
Join-Meeting -Subject RE -Type Lecture

# Join English without specifying class type.
Join-Meeting -Subject English

# Join current/closest class (based on epi.kpi.ua).
Join-Meeting

# Join current/closest class without asking for permission.
Join-Meeting -Quiet
```

## Comments

* You may get extra info and examples of usage by [Get-Help](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/get-help?view=powershell-7.3)
* At the moment signing this module is not an option. So I would advise changing execution policy (see [Get-ExecutionPolicy](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-executionpolicy?view=powershell-7.3) and [Set-ExecutionPolicy](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.3)) or copying this module rather than downloading.

## Plans

* Add Linux and MacOS support.

## License

[MIT](https://choosealicense.com/licenses/mit/)
