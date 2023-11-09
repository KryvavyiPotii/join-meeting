function Join-Meeting
{
    param
    (
        [ValidateSet(
            'English', 'RE', 'Crypto',
            'ProbTh', 'QAQC', 'PEA',
            'SysTech', 'MTAD', 'AlgoAn',
            'SpecRozd', 'Opers'
        )] $Subject,
        [ValidateSet('Lecture', 'Practice', 'Lab')] $Type,
        [switch] $Quiet
    )

    # Join meeting.
    function Establish-Connection
    {
        Write-Host "[INFO] Establishing connection to $($Subject) $($Type)..."

        try 
        {
            Start-Process "$($hash.$Subject.$Type.Link)"
            "[INFO] Connected to $($Subject) $($Type)"

            # Show extra info if there is.
            if ($hash.$Subject.$Type.Info)
            {
                Write-Host "[INFO] $($hash.$Subject.$Type.Info)"
            }
        }
        catch
        {
            Write-Warning 'Failed to open link'
        }
    }

    # Link to group's schedule.
    $roz = 'http://epi.kpi.ua/Schedules/ViewSchedule.aspx?g=3a2a5666-0a50-4695-8602-8637ef7b6b62'
    # Path to a .json file with subject short titles, links and extra info.
    $path = ''

    # Create hashtable from a file.
    try
    {
        $jsonObj = Get-Content -Path $path -Raw | ConvertFrom-Json
        $hash = @{}
        foreach ($property in $jsonObj.PSObject.Properties)
        {
            $hash[$property.Name] = $property.Value
        }
    }
    catch
    {
        Write-Warning 'Failed to create a hashtable'
        Write-Warning 'Check if path to a .json file is correct'
        return
    }

    # Display current time.
    Write-Host "[INFO] Time: $(Get-Date -UFormat "%R")"

    # Show current/closest class.
    try
    {
        $html = Invoke-WebRequest $roz -UseBasicParsing
        $data = $html.Content
        # Find area that contains class title.
        $ind = $data.IndexOf('<td class="closest_pair">')
        if ($ind -lt 0)
        {
            $ind = $data.IndexOf('<td class="current_pair">')
        }
        # Shorten found area.
        $inds = $data.IndexOf('<a', $ind)
        $inde = $data.IndexOf('</a', $inds)
        $tag = $data.Substring($inds, $inde-$inds)
        # Get class title.
        $class = $tag -replace '\<[^\>]*\>'

        Write-Host "[INFO] Current/closest class: $($class)"
    }
    catch
    {
        Write-Warning 'Failed to get info from epi.kpi.ua'
        Write-Warning 'Check if schedule link is up to date and servers work'
    }
        
    # Define current/closest class.
    if (-not $Subject -And -not ($data -eq $null))
    {
        $ClassTypes = @{
            'Лек' = 'Lecture' 
            'Лаб' = 'Lab'
            'Прак' = 'Practice'
        }

        # Find Subject with title Class.
        foreach ($key in $hash.Keys)
        {
            if ($hash.$key.Title -eq $class)
            {
                # Define class title.
                $Subject = "$($key)"

                # Define class type.
                $curr = $data.Substring($ind, 500) -replace '\<[^\>]*\>'
                foreach ($key in $ClassTypes.Keys)
                {
                    if ($curr.Contains($key))
                    {
                        $Type = $ClassTypes[$key]
                        break
                    }
                }  
            }
        }

        # Case hashtable doesn't have Subject.
        if (-not $Subject)
        {
            Write-Warning "Subject `"$($class)`" is not supported"
        }
    }

    # Case Type not specified.
    if (-not $Type)
    {
        # Choose any type of a specified subject.
        foreach ($key in $hash.$Subject.Keys)
        {
            if (-not ($key -eq 'Title'))
            {
                $Type = $hash.$Subject.$key
                break
            }
        }

        # Check if Subject is not missing type in the hashtable.
        if ($Type)
        {
            Write-Host "[INFO] Type was not specified: connecting to $($Subject) $($Type)"
        }
        else
        {
            Write-Warning "Subject $($Subject) does not have a single type"
        }
    }

    # Connect to a meeting.
    if (-not $Quiet)
    {
        $ans = Read-Host "[CHOICE] Do you want to connect to $($Subject) ($($Type))? [y or n]"
        if ($ans -eq 'y')
        {
            Establish-Connection
        }
    }
    else
    {
        Establish-Connection
    }

    Write-Host '[INFO] Finishing work...'

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
        System.String. Join-Meeting returns information strings about system time, connection status, extra subject info etc.
  
        .EXAMPLE
        PS> Join-Meeting -Subject RE -Type Lecture
        [INFO] Time: 19:19
        [INFO] Current/closest class: Теорія ймовірностей та математична статистика
        [CHOICE] Do you want to connect RE (Lecture)? [y or n]: y
        [INFO] Establishing connection to RE Lecture...
        [INFO] Connected to RE Lecture
        [INFO] Access code: 079049
        [INFO] Finishing work...

        .EXAMPLE
        PS> Join-Meeting -Subject English
        [INFO] Time: 19:11
        [INFO] Current/closest class: Теорія ймовірностей та математична статистика
        [INFO] Type was not specified: connecting to English Practice
        [CHOICE] Do you want to connect English (Practice)? [y or n]: y
        [INFO] Establishing connection to English Practice...
        [INFO] Connected to English Practice
        [INFO] "I am present!" and "Bye bye"
        [INFO] Finishing work...

        .EXAMPLE
        PS> Join-Meeting
        [INFO] Time: 19:07
        [INFO] Current/closest class: Теорія ймовірностей та математична статистика
        [CHOICE] Do you want to connect ProbTh Lecture? [y or n]: n
        [INFO] Finishing work...
    #>
}
