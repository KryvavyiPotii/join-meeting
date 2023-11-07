function Join-Meeting {
    param(
        [ValidateSet(
            'English', 'RE', 'Crypto',
            'ProbTh', 'QAQC', 'PEA',
            'SysTech', 'MTAD', 'AlgoAn',
            'SpecRozd', 'Opers'
        )] $Subject,
        [ValidateSet('Lecture', 'Practice', 'Lab')] $Type
    )

    $Semester5 = @{
        'English' = @{
            'Practice' = @{
                'Link' = 'https://us04web.zoom.us/j/2880207160?pwd=N3lsWFpURnY3TlFrZm9VWnFRaThHZz09'
                'Info' = '"I am present!" and "Bye bye"'
            }
        }

        'RE' = @{
            'Lecture' = @{
                'Link' = 'https://bbb.kpi.ua/b/myk-q4m-xn5-csj'
                'Info' = 'Access code: 079049'
            }
            'Practice' = @{
                'Link' = 'https://bbb.kpi.ua/b/uwq-ofo-gce-lhw'
                'Info' = 'Access code: 028474'
            }
        }

        'Crypto' = @{
            'Lecture' = @{
                'Link' = 'https://us04web.zoom.us/j/77493733844?pwd=VmtUWHNDSytHNnc1QkNXSGtseWJmUT09'
                'Info' = ''
            }
            'Practice' = @{
                'Link' = 'https://us05web.zoom.us/j/5757832035?pwd=aHlCdzVBcllqVzcwdnJTdG1BMTFNZz09'
                'Info' = 'Ask for link'
            }
            'Lab' = @{
                'Link' = 'https://us05web.zoom.us/j/9442768212?pwd=nQWbqnK7bDPN0fATbZl53IBkZSIOcl.1'
                'Info' = 'Pay attention to queue'
            }
        }

        'ProbTh' = @{
            'Lecture' = @{
                'Link' = 'https://us04web.zoom.us/j/7620255592?pwd=RFRveFIrbWR0TWIzRmRleHhOTjV5QT09'
                'Info' = 'Say "cheese" approx. at 9'
            }
            'Practice' = @{
                'Link' = 'https://us04web.zoom.us/j/7021457189?pwd=dytzTUFLSEluU1RIYkdtK3orSlkyUT09'
                'Info' = 'Prepare for shitshow :^)'
            }
        }

        'QAQC' = @{
            'Lecture' = @{
                'Link' = 'https://bth.zoom.us/j/62165940625'
                'Info' = 'Write down lecture'
            }
            'Lab' = @{
                'Link' = 'https://us04web.zoom.us/j/5095167397?pwd=bWc4QWdoNGM2bGxWRWp4bWs5eXFsdz09'
                'Info' = ''
            }
        }

        'PEA' = @{
            'Lecture' = @{
                'Link' = 'https://us06web.zoom.us/j/4292900213?pwd=UTJ3UndzT1g5cGtzM0sybTUwNU5aQT09'
                'Info' = 'Starts at 14:00 (without a break)'
            }
            'Lab' = @{
                'Link' = 'https://us06web.zoom.us/j/4292900213?pwd=UTJ3UndzT1g5cGtzM0sybTUwNU5aQT09'
                'Info' = 'Starts at 14:00 (without a break)'
            }
        }

        'SysTech' = @{
            'Lecture' = @{
                'Link' = 'https://web.telegram.org/a/#-1001700380029'
                'Info' = 'Check Telegram chat for link'
            }
            'Lab' = @{
                'Link' = 'https://us04web.zoom.us/j/5095167397?pwd=bWc4QWdoNGM2bGxWRWp4bWs5eXFsdz09'
                'Info' = ''
            }
        }

        'MTAD' = @{
            'Lecture' = @{
                'Link' = 'https://us02web.zoom.us/j/81033538779?pwd=Q3VqUHhuVDhCY0p1VnRiRk5SNEdRdz09'
                'Info' = ''
            }
            'Practice' = @{
                'Link' = 'https://meet.google.com/ffa-aetm-ogt'
                'Info' = ''
            }
        }

        'AlgoAn' = @{
            'Lab' = @{
                'Link' = 'https://us04web.zoom.us/j/9168981041?pwd=WmI4RS96KzNqM1p4MXZML0hEbk9vUT09'
                'Info' = ''
            }
        }

        'SpecRozd' = @{
            'Lecture' = @{
                'Link' = 'https://us04web.zoom.us/j/3908947683?pwd=UWhmb1B0NDZSU3dZc3ZVL0RaSUJqdz09'
                'Info' = 'Same link as for practice'
            }
            'Practice' = @{
                'Link' = 'https://us04web.zoom.us/j/3908947683?pwd=UWhmb1B0NDZSU3dZc3ZVL0RaSUJqdz09'
                'Info' = 'Same link as for lecture'
            }
        }
       
        'Opers' = @{
            'Lecture' = @{
                'Link' = 'https://t.me/+_15ip0MuBABkN2Uy'
                'Info' = 'Link in Telegram chat'
            }
            'Practice' = @{
                'Link' = 'https://t.me/+_15ip0MuBABkN2Uy'
                'Info' = 'Link in Telegram chat'
            }
        }
    }

    # Display current time.
    Write-Host "[INFO] Time: $(Get-Date -UFormat "%R")"

    # Show current/closest class.
    try {
        $roz = Invoke-WebRequest http://epi.kpi.ua/Schedules/ViewSchedule.aspx?g=3a2a5666-0a50-4695-8602-8637ef7b6b62
        # Find area that contains class title.
        $ind = $roz.Content.IndexOf('<td class="closest_pair">')
        if ($ind -lt 0) {
            $ind = $roz.Content.IndexOf('<td class="current_pair">')
        }
        # Shorten found area.
        $inds = $roz.Content.IndexOf('<a', $ind)
        $inde = $roz.Content.IndexOf('</a', $inds)
        $tag = $roz.Content.Substring($inds, $inde-$inds)
        # Get class title.
        $class = $tag -replace '\<[^\>]*\>'

        Write-Host "[INFO] Current/closest class: $($class)"
    }
    catch {
        Write-Warning 'Failed to get info from epi.kpi.ua'
    }

    if (-not $Subject -And -not ($roz -eq $null)) {
        # Define class title.
        $ClassTitles = @{
            'Іноземна мова професійного спрямування. Частина 1' = 'English'
            'Зворотна розробка та аналіз шкідливого програмного забезпечення' = 'RE'
            'Криптографія' = 'Crypto'
            'Теорія ймовірностей та математична статистика' = 'ProbTh'
            'Технології забезпечення якості програмних засобів' = 'QAQC'
            'Програмування ефективних алгоритмів' = 'PEA'
            'Системні технології для застосувань Windows' = 'SysTech'
            'Методи та технології аналітики даних' = 'MTAD'
            'Основи аналізу алгоритмів' = 'AlgoAn'
            'Спеціальні розділи обчислювальної математики' = 'SpecRozd'
            'Дослідження операцій' = 'Opers'
        }
        if ($ClassTitles.ContainsKey($class)) {
            $Subject = $ClassTitles[$class]    
        }
        else {
            Write-Warning "Subject `"$($class)`" is not supported"
            return
        }

        # Define class type.
        $ClassTypes = @{
            'Лек' = 'Lecture' 
            'Лаб' = 'Lab'
            'Прак' = 'Practice'
        }
        $data = $roz.Content.Substring($ind, 500) -replace '\<[^\>]*\>'
        foreach ($key in $ClassTypes.Keys) {
            if ($data.Contains($key)) {
                $Type = $ClassTypes[$key]
                break
            }
        }
    }

    # Join meeting.
    try {
        # Case Type not specified.
        if (-not $Type) {
            # Choose any type of a specified subject
            $Type = $Semester5.$Subject.Keys | Get-Random
            Write-Host "[INFO] Type was not specified: connecting to $($Subject) $($Type)"
        }

        Start-Process "$($Semester5.$Subject.$Type.Link)"
        "[INFO] Connected to $($Subject) $($Type)"

        # Show extra info if there is.
        if ($Semester5.$Subject.$Type.Info) {
            Write-Host "[INFO] $($Semester5.$Subject.$Type.Info)"
        }
    }
    catch {
        Write-Warning 'Failed to open link'
        return
    }

    <#
        .SYNOPSIS
        Connects to a meeting.

        .DESCRIPTION
        Connects to a meeting of current/next/specified subject.

        .PARAMETER Subject
        Specifies a subject title.

        .PARAMETER Type
        Specifies a class type of a class (lecture, practice or lab).

        .INPUTS
        None. You can't pipe objects to Join-Meeting.

        .OUTPUTS
        System.String. Join-Meeting returns information strings about system time, connection status, extra subject info etc.
  
        .EXAMPLE
        PS> Join-Meeting -Subject RE -Type Lecture
        [INFO] Time: 20:38
        [INFO] Current/closest class: Основи аналізу алгоритмів
        [INFO] Connected to RE Lecture
        [INFO] Access code: 079049

        .EXAMPLE
        PS> Join-Meeting -Subject English
        [INFO] Time: 21:27
        [INFO] Current/closest class: Основи аналізу алгоритмів
        [INFO] Type was not specified: connecting to English Practice
        [INFO] Connected to English Practice
        [INFO] "I am present!" and "Bye bye"

        .EXAMPLE
        PS> Join-Meeting
        [INFO] Time: 21:54
        [INFO] Current/closest class: Основи аналізу алгоритмів
        [INFO] Connected to AlgoAn Lab
    #>
}