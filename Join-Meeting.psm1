function Join-Meeting
{
    param
    (
        [ValidateSet(
            'English', 'RE', 'Crypto',
            'ProbTh', 'QAQC', 'PEA',
            'SysTech', 'MTAD', 'AlgoAn',
            'SpecRozd', 'Opers'
        )]
        $Subject,

        [ValidateSet('Lecture', 'Practice', 'Lab')]
        $Type,

        [switch] $Quiet
    )

    # Display current time.
    Write-Host "[OUT]`tTime:`t$(Get-Date -UFormat "%R")"

    # Display schedule data.
    $roz = 'http://epi.kpi.ua/Schedules/ViewSchedule.aspx?g=3a2a5666-0a50-4695-8602-8637ef7b6b62'
    $schedule = Get-ScheduleData -Link $roz
    if ($schedule -ne $null)
    {
        Write-Host "[OUT]`tGroup:`t`"$($schedule.Group)`""
        Write-Host "[OUT]`tCurrent/closest class: `"$($schedule.Title) ($($schedule.Type))`""
    }
    
    # Path to a .json file with subject short titles, links and extra info.
    $path = 'C:\Users\zh4rgone\Documents\WindowsPowerShell\Modules\Join-Meeting\subjects.json'
    $Subjects = ConvertTo-Hashtable -Path $path
    if ($Subjects -eq $null)
    {
        Write-Host "[OUT]`tFinishing work..."
        return
    }

    # Ukrainian strings for class types.
    $UATypes = @{
        'Лек on-line' = 'Lecture' 
        'Лаб on-line' = 'Lab'
        'Прак on-line' = 'Practice'
    }

    # Case join current/closest class.
    if (-not $Subject)
    {
        if ($schedule -ne $null)
        {
            # Find current/closest class in hashtable by title.
            foreach ($s in $Subjects.Keys)
            {
                if ($Subjects.$s.Title -eq $schedule.Title)
                {
                    $Subject = $s
                    $Type = $UATypes[$schedule.Type]
                    break
                }
            }
            $uatitle = $schedule.Title
            $uatype = $schedule.Type

            if (-not $Subject)
            {
                Write-Warning "Subject `"$($schedule.Title)`" is not supported"
            }
        }
        else
        {
            Write-Warning 'Not enough data for joining a meeting'
            Write-Host "[OUT]`tFinishing work..."
            return
        }
    }
    # Case join subject without specified type.
    elseif ($Subject -And -not $Type)
    {
        # Choose any type of a specified subject.
        if ($Subjects.$Subject.Lecture)
        {
            $uatype = 'Лек on-line'
        }
        elseif ($Subjects.$Subject.Lab)
        {
            $uatype = 'Лаб on-line'
        }
        elseif ($Subjects.$Subject.Practice)
        {
            $uatype = 'Прак on-line'
        }
        $Type = $UATypes.$uatype
        $uatitle = $Subjects.$Subject.Title

        # Check if Subject is not missing type in the hashtable.
        if ($Type)
        {
            Write-Host "[OUT]`tType was not specified: connecting to `"$($uatitle) ($($uatype))`""
        }
        else
        {
            Write-Warning "Subject `"$($uatitle)`" does not have a single type"
        }
    }
    # Case join subject with specified type.
    else
    {
        $uatitle = $Subjects.$Subject.Title
        
        # Set Ukrainian type.
        foreach ($t in $UATypes.Keys)
        {
            if ($UATypes.$t -eq $Type)
            {
                $uatype = $t
                break
            }
        }
    }

    # Connect to a meeting.
    $openLinkSplat = @{
        'Link' = $Subjects.$Subject.$Type.Link
        'Info' = $Subjects.$Subject.$Type.Info
    }
    if (-not $Quiet)
    {
        $ans = Read-Host "[IN]`tConnect to `"$($uatitle) ($($uatype))`"? [yn]"
        if ($ans -ne $null -And $ans -ne 'y')
        {
            Write-Host "[OUT]`tFinishing work..."
            return
        }
    }
    Open-Link @openLinkSplat

    Write-Host "[OUT]`tFinishing work..."

    <#
        .SYNOPSIS
        Connects to a meeting.

        .DESCRIPTION
        Connects to a meeting of current/next/specified subject.

        .PARAMETER Subject
        Specifies a subject title.

        .PARAMETER Type
        Specifies a class type of a class (lecture, practice or lab).

        .PARAMETER Quiet
        Do not ask user to accept or decline connecting.

        .INPUTS
        None. You can't pipe objects to Join-Meeting.

        .OUTPUTS
        System.String. Join-Meeting returns information strings about connection status.
  
        .EXAMPLE
        PS> Join-Meeting -Subject RE -Type Lecture
        [OUT]   Time:   22:11
        [OUT]   Group:  "ФБ-13"
        [OUT]   Current/closest class: "Технології забезпечення якості програмних засобів (Лаб on-line)"
        [IN]    Connect to "Зворотна розробка та аналіз шкідливого програмного забезпечення (Лек on-line)"? [y or n]: y
        [OUT]   Connection established.
        [OUT]   Access code: 079049
        [OUT]   Finishing work...

        .EXAMPLE
        PS> Join-Meeting -Subject English
        [OUT]   Time:   22:05
        [OUT]   Group:  "ФБ-13"
        [OUT]   Current/closest class: "Технології забезпечення якості програмних засобів (Лаб on-line)"
        [OUT]   Type was not specified: connecting to "Іноземна мова професійного спрямування. Частина 1 (Прак on-line)"
        [IN]    Connect to "Іноземна мова професійного спрямування. Частина 1 (Прак on-line)"? [y or n]: y
        [OUT]   Connection established.
        [OUT]   "I am present!" and "Bye bye"
        [OUT]   Finishing work...

        .EXAMPLE
        PS> Join-Meeting
        [OUT]   Time:   17:02
        [OUT]   Group:  "ФБ-13"
        [OUT]   Current/closest class: "Технології забезпечення якості програмних засобів (Лаб on-line)"
        [IN]    Connect to "Технології забезпечення якості програмних засобів (Лаб on-line)"? [y or n]: y
        [OUT]   Establishing connection to "Технології забезпечення якості програмних засобів (Лаб on-line)"...
        [OUT]   Connection established.
        [OUT]   Finishing work...

        .EXAMPLE
        PS> Join-Meeting -Quiet
        [OUT]   Time:   17:04
        [OUT]   Group:  "ФБ-13"
        [OUT]   Current/closest class: "Технології забезпечення якості програмних засобів (Лаб on-line)"
        [OUT]   Establishing connection to "Технології забезпечення якості програмних засобів (Лаб on-line)"...
        [OUT]   Connection established.
        [OUT]   Finishing work...
    #>
}

function ConvertTo-Hashtable
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $Path
    )

    if (Test-Path $Path)
    {
        try
        {
            $jsonObj = Get-Content -Path $Path -Raw | ConvertFrom-Json
            $hash = @{}
            foreach ($property in $jsonObj.PSObject.Properties)
            {
                $hash[$property.Name] = $property.Value
            }

            # Return created hashtable
            $hash
        }
        catch
        {
            Write-Warning 'Failed to create a hashtable'
            Write-Warning 'Check if path to a .json file is correct'

            $null
        }
    }
    else
    {
        Write-Warning "Subject file $Path does not exist"

        $null
    }

    <#
        .SYNOPSIS
        Converts .json file to a hashtable.

        .DESCRIPTION
        Converts .json file to a hashtable.

        .PARAMETER Path
        Specifies a path to .json file.

        .INPUTS
        None. You can't pipe objects to ConvertTo-Hashtable.

        .OUTPUTS
        System.Collections.Hashtable. On success Get-ScheduleData returns a hashtable that contains information about subjects.
  
        .EXAMPLE
        PS> ConvertTo-Hashtable -Path .\subjects.json

        Name                           Value                                                                                                            
        ----                           -----                                                                                                            
        SpecRozd                       @{Title=Спеціальні розділи обчислювальної математики; Lecture=; Practice=}                                       
        QAQC                           @{Title=Технології забезпечення якості програмних засобів; Lab=; Lecture=}                                       
        English                        @{Title=Іноземна мова професійного спрямування. Частина 1; Practice=}                                            
        Opers                          @{Title=Дослідження операцій; Lecture=; Practice=}                                                               
        SysTech                        @{Title=Системні технології для застосувань Windows; Lab=; Lecture=}                                             
        RE                             @{Title=Зворотна розробка та аналіз шкідливого програмного забезпечення; Lecture=; Practice=}                    
        ProbTh                         @{Title=Теорія ймовірностей та математична статистика; Lecture=; Practice=}                                      
        MTAD                           @{Title=Методи та технології аналітики даних; Lecture=; Practice=}                                               
        AlgoAn                         @{Title=Основи аналізу алгоритмів; Lab=}                                                                         
        Crypto                         @{Title=Криптографія; Lecture=; Lab=; Practice=}                                                                 
        PEA                            @{Title=Програмування ефективних алгоритмів; Lab=; Lecture=}
    #>
}

function Get-ScheduleData
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $Link
    )

    $timeout = 5

    # Show current/closest class.
    try
    {
        $html = Invoke-WebRequest $Link -UseBasicParsing -TimeoutSec $timeout
        $data = $html.Content

        # Define group ID.
        ($data -match "(?<group>[А-Я]{2}-\d{2})") | Out-Null
        $group = $Matches.group

        # Find current/closest class title.
        ($data -match "t_pair`">.{50,200}title=`"(?<title>.{1,150})`">") | Out-Null
        $title = $Matches.title

        # Find current/closest class type.
        ($data -match "t_pair`">.{200,500} (?<type>.{3,4} on-line)") | Out-Null
        $type = $Matches.type

        $ScheduleData = @{
            'Group' = $group
            'Title' = $title
            'Type' = $type
        }
        
        # Return info about group and current/closest class.
        $ScheduleData
    }
    catch
    {
        Write-Warning 'Failed to get info from $Link'
        Write-Warning 'Check if schedule link is up to date and servers work'

        $null
    }

    <#
        .SYNOPSIS
        Gets data from schedule link.

        .DESCRIPTION
        Parses HTML of a schedule and gets group ID and current/closest class.

        .PARAMETER Link
        Specifies a link to a schedule.

        .INPUTS
        None. You can't pipe objects to Get-ScheduleData.

        .OUTPUTS
        System.Collections.Hashtable. On success Get-ScheduleData returns a hashtable that contains group ID, class title and type.
  
        .EXAMPLE
        PS> Get-ScheduleData -Link http://epi.kpi.ua/Schedules/ViewSchedule.aspx?g=3a2a5666-0a50-4695-8602-8637ef7b6b62

        Name                           Value                                                                                                            
        ----                           -----                                                                                                            
        Group                          ФБ-13                                                                                                            
        Title                          Технології забезпечення якості програмних засобів                                                                
        Type                           Лаб on-line
    #>
}

function Open-Link
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $Link,

        [string] $Info
    )

    try 
    {
        Start-Process $Link
        Write-Host "[OUT]`tConnection established."

        # Show extra info if there is any.
        if ($Info)
        {
            Write-Host "[OUT]`t$Info"
        }
    }
    catch
    {
        Write-Warning 'Failed to connect'
    }

    <#
        .SYNOPSIS
        Opens a link.

        .DESCRIPTION
        Opens a link in browser (for example, a link to some Zoom meeting).

        .PARAMETER Link
        Specifies a link.

        .PARAMETER Info
        Specifies extra information about a link that should be written to output.

        .INPUTS
        None. You can't pipe objects to Open-Link.

        .OUTPUTS
        System.String. Open-Link returns information strings about connection status and extra info.
  
        .EXAMPLE
        PS> Open-Link -Link https://us04web.zoom.us/j/77493733844?pwd=VmtUWHNDSytHNnc1QkNXSGtseWJmUT09
        [OUT]   Connection established.

        .EXAMPLE
        PS> Open-Link -Link https://us05web.zoom.us/j/9442768212?pwd=nQWbqnK7bDPN0fATbZl53IBkZSIOcl.1 -Info "Pull request!"
        [OUT]   Connection established.
        [OUT]   Pull request!
    #>
}
