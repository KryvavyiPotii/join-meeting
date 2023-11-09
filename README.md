# Join-Meeting

Join-Meeting is a Powershell module for fast and comfortable connection to KPI classes.

## Installation

Put the module file into directory `C:\Users\<username>\Documents\WindowsPowerShell\Modules\Join-Meeting`.
If directory does not exist, create it manually.
Put subjects.json wherever you want and specify it's path in variable `$path` in module file.

```powershell
Import-Module Join-Meeting
```

## Update

In order to update module you should change the module file. 
After that you need to remove it and import again.

```powershell
Remove-Module Join-Meeting
Import-Module Join-Meeting
```

A faster option is to use `Import-Module` command with `-Force` flag.

```powershell
Import-Module Join-Meeting -Force
```

In case you want to modify `subjects.json` you should keep in mind the format:
* Outer keys are short subject titles.
  * "Title" - full subject title that is present in epi.kpi.ua.
  * "Type" - hashtable of possible class types ("Lecture", "Lab" and "Practice").
    * "Link" - class link for a desired type.
    * "Info" - extra comments for user that aren't used in code.

It is not recommended to change any key names except for those of short and full subject titles.
After adding/modifying/removing short titles edit `ValidateSet` of `Subject` parameter accordingly.

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

* You may get extra info and examples of usage with [Get-Help](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/get-help?view=powershell-7.3)
* At the moment signing this module is not an option. So I would advise changing execution policy (see [Get-ExecutionPolicy](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-executionpolicy?view=powershell-7.3) and [Set-ExecutionPolicy](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.3)) or copying this module rather than downloading.

## Plans

* Add Linux and MacOS support.

## License

[MIT](https://choosealicense.com/licenses/mit/)
