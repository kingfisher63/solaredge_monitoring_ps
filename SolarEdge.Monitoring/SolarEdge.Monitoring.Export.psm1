# Copyright (C) 2025 Roger Hunen

Set-StrictMode -Version 3.0

# 0: date
$solarEdgeDateFormat = '{0:yyyy}-{0:MM}-{0:dd}'

#
# Support functions (not exported)
#

function PatternToFormat
{
    param(
        [string] $pattern,
        [object] $placeHolders
    )

    $percent = $false
    $format  = ''

    foreach ($c in $pattern.ToCharArray()) {
        $char = $c.ToString()

        if ($char -eq '%') {
            if ($percent) {
                $format += '%'
                $percent = $false
            } else {
                $percent = $true
            }
        } else {
            if ($percent) {
                if ($placeHolders.ContainsKey($char)) {
                    $format += "{$($placeHolders.$char[0])$($placeHolders.$char[1])}"
                } else {
                    throw "Invalid placeholder '$($char)'"
                }
                $percent = $false
            } else {
                $format += $char
            }
        }
    }

    if ($percent) {
        $format += '%'
    }

    return $format
}

#
# Exported functions
#

function Export-SolarEdgeInverterData
{
    <#
        .SYNOPSIS
        Gets inverter technical data from the SolarEdge monitoring platform and
        writes it to a CSV file.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER SerialNumber
        The inverter serial number.
        .PARAMETER StartDate
        The start date.
        .PARAMETER EndDate
        The end date.
        .PARAMETER OutFilePattern
        The output file name pattern. The pattern can contain the following
        placeholders:

          %I  The SolarEdge site ID
          %S  The inverter serial number
          %Y  The calendar year (4 digits)
          %M  The calendar month (01-12)
          %D  The calendar day (01-31)
          %J  The day of year (000-366)
          %%  Percent character

        The default pattern for month data is '%I %S %Y-%M-%D.csv'.
        .LINK
        Get-SolarEdgeInverterData
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [Parameter(Mandatory,Position=0)] [string]   $ApiKey,
        [Parameter(Mandatory,Position=1)] [string]   $Site,
        [Parameter(Mandatory,Position=2)] [string]   $SerialNumber,
        [parameter(Mandatory,Position=3)] [DateTime] $StartDate,
        [parameter(Mandatory,Position=4)] [DateTime] $EndDate,
                                          [string]   $OutFilePattern
    )

    begin {
        function FillDataRow
        {
            param (
                [System.Data.DataRow] $dataRow,
                [PSCustomObject]      $object,
                [int]                 $level,
                [string]              $prefix
            )

            process {
                foreach ($property in ($object | Get-Member -Type NoteProperty)) {
                    $columnName   = "${prefix}$($property.Name)"
                    $propertyType = ($property.Definition -split ' ')[0];

                    if ($propertyType -eq 'System.Management.Automation.PSCustomObject') {
                        FillDataRow $dataRow $object.($property.Name) ($level+1) "${columnName}_"
                    } else {
                        $dataRow[${columnName}] = $object.($property.Name)
                    }
                }
            }
        }

        function GetColumns
        {
            param (
                [hashtable]      $columns,
                [PSCustomObject] $object,
                [int]            $level,
                [string]         $prefix
            )

            process {
                foreach ($property in ($object | Get-Member -Type NoteProperty)) {
                    $columnName   = "${prefix}$($property.Name)"
                    $propertyType = ($property.Definition -split ' ')[0];

                    if ($propertyType -eq 'System.Management.Automation.PSCustomObject') {
                        GetColumns $columns $object.($property.Name) ($level+1) "${columnName}_"
                    } else {
                        $columns["${level}:${columnName}"] = $propertyType
                    }
                }
            }
        }
    }

    process {
        if ($StartDate.Hour -ne 0 -or $StartDate.Minute -ne 0 -or $StartDate.Second -ne 0 -or $StartDate.Millisecond -ne 0) {
            throw [ArgumentException]::New('Start date has a non-zero time component.', 'StartDate')
        }

        if ($EndDate.Hour -ne 0 -or $EndDate.Minute -ne 0 -or $EndDate.Second -ne 0 -or $EndDate.Millisecond -ne 0) {
            throw [ArgumentException]::New('End date has a non-zero time component.', 'EndDate')
        }

        if ($EndDate -lt $StartDate) {
            throw [ArgumentOutOfRangeException]::New("StartDate/EndDate", 'End date is before start date.')
        }

        $totalDays      = ($EndDate - $StartDate).TotalDays
        $queryStartDate = $StartDate

        if (-not $PSBoundParameters.ContainsKey('OutFilePattern')) {
            $OutFilePattern = '%I %S %Y-%M-%D.csv'
        }

        $placeHolders = @{
            'I' = (0, '')
            'S' = (1, '')
            'Y' = (2, ':yyyy')
            'M' = (2, ':MM')
            'D' = (2, ':dd')
            'J' = (3, ':D3')
        }

        $outFileFormat = PatternToFormat $OutFilePattern $placeHolders

        Write-Verbose "Site ID       : $Site"
        Write-Verbose "Serial number : $SerialNumber"
        Write-Verbose "Start date    : $($solarEdgeDateFormat -f $StartDate)"
        Write-Verbose "Days          : $totalDays"

        while ($totalDays -gt 0) {
            $queryDays    = if ($totalDays -lt 7 ) { $totalDays } else { 7 }
            $queryEndDate = $queryStartDate.AddDays($queryDays)

            $inverterData = (Get-SolarEdgeInverterData $apiKey $site $SerialNumber $queryStartDate $queryEndDate -ErrorAction Stop).inverterData

            Write-Verbose "Query start   : $($solarEdgeDateFormat -f $queryStartDate)"
            Write-Verbose "Query end     : $($solarEdgeDateFormat -f $queryEndDate)"
            Write-Verbose "Query days    : $queryDays"
            Write-Verbose "Data points   : $($inverterData.telemetries.Count)"

            if ($inverterData.telemetries.Count -eq 0) {
                break
            }

            for ($day=0; $day -lt $queryDays; $day++) {
                $date    = $queryStartDate.AddDays($day)
                $dateStr = $solarEdgeDateFormat -f $date

                $telemetries = $inverterData.telemetries | Where-Object { ([DateTime]$_.date).Date -eq $date }
                if ($null -eq $telemetries) {
                    continue
                }

                # Collect column names

                $columns = @{}
                foreach ($telemetry in $telemetries) {
                    GetColumns $columns $telemetry 0 ''
                }

                # Create data table

                $table = [System.Data.DataTable]::new()

                foreach ($key in ($columns.Keys | Sort-Object)) {
                    $column = [System.Data.DataColumn]::new(($key -split ':')[1])

                    switch ($columns[$key]) {
                        'decimal' { $column.DataType = [System.Decimal]; $column.DefaultValue = 0  }
                        'double'  { $column.DataType = [System.Double];  $column.DefaultValue = 0  }
                        'int'     { $column.DataType = [System.Int32];   $column.DefaultValue = 0  }
                        default   { $column.DataType = [System.String];  $column.DefaultValue = '' }
                    }

                    $table.Columns.Add($column)
                }

                # Add data rows

                foreach ($telemetry in $telemetries) {
                    $row = $table.NewRow()
                    FillDataRow $row $telemetry 0 ''
                    $table.Rows.Add($row)
                }

                # Export to CSV file

                $outFileName = $outFileFormat -f $Site, $SerialNumber, $date, $date.DayOfYear

                Write-Verbose ('{0}    : {1,-3} -> {2}' -f $dateStr, $telemetries.Count, $outFileName)
                $table | Export-Csv -Path $outFileName -NoTypeInformation
            }

            $queryStartDate = $queryEndDate
            $totalDays     -= $queryDays
        }
    }
}

function Export-SolarEdgeSiteEnergy
{
    <#
        .SYNOPSIS
        Gets site energy data from the SolarEdge monitoring platform and writes
        it to a CSV file.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER Year
        The year for which site energy data is exported. The default value is
        the current year.
        .PARAMETER Month
        The month for which site energy data is exported. If the month is not
        specified, site energy data is exported for the whole year.
        .PARAMETER TimeUnit
        The time granularity of the site energy data. DAY, WEEK, MONTH and YEAR
        are accepted when exporting year data. 15MIN, HOUR, DAY, WEEK and MONTH
        are accepted When exporting month data. The default value is DAY.
        .PARAMETER OutFilePattern
        The output file name pattern. The pattern can contain the following
        placeholders:

          %I  The SolarEdge site ID
          %N  The SolarEdge site name
          %Y  The calendar year (4 digits)
          %M  The calendar month (01-12)
          %U  The time unit
          %%  Percent character

        The default pattern for month data is '%N (%I) %Y-%M %U.csv'. The default
        pattern for year data is '%N (%I) %Y %U.csv'.
        .LINK
        Get-SolarEdgeSiteEnergy
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [Parameter(Mandatory,Position=0)] [string] $ApiKey,
        [Parameter(Mandatory,Position=1)] [string] $Site,
                                          [int]    $Year = [DateTime]::Now.Year,
                                          [int]    $Month,
                                          [string] $TimeUnit = 'DAY',
                                          [string] $OutFilePattern
    )
    
    process {
        if (($Year -lt 1980) -or ($Year -gt 9999)) {
            throw "Year '$Year' is out of range (1980-9999)"
        }

        # There is a flaw in the SolarEdge Monitoring API. In some cases data points
        # at the end date are included, in other cases (QUARTER_OF_AN_HOUR, HOUR) they
        # are not. We therefore set the end date at the 1st day of the next year/month
        # and drop data points at the end date later. (RH 2025/10/05)

        if ($PSBoundParameters.ContainsKey('Month')) {
            $timeUnits = '15MIN', 'QUARTER_OF_AN_HOUR', 'HOUR', 'DAY', 'MONTH'

            if (($Month -lt 1) -or ($Month -gt 12)) {
                throw "Month '$Month' is out of range (1-12)"
            }
            if ($timeUnits -notcontains $TimeUnit) {
                throw "Invalid time unit '$TimeUnit' for month data ($($timeUnits -join ', '))"
            }
            $TimeUnit = $TimeUnit.ToUpper()

            $startDate = [DateTime]::New($Year, $Month, 1)
            $endDate   = $StartDate.AddMonths(1)

            if (-not $PSBoundParameters.ContainsKey('OutFilePattern')) {
                $OutFilePattern = '%N (%I) %Y-%M %U.csv'
            }

            $placeHolders = @{
                'I' = (0, '')
                'N' = (1, '')
                'Y' = (2, ':yyyy')
                'M' = (2, ':MM')
                'U' = (3, '')
            }

            $outFileFormat = PatternToFormat $OutFilePattern $placeHolders
        } else {
            $Month     = 0
            $timeUnits = 'DAY', 'MONTH', 'YEAR'

            if ($timeUnits -notcontains $TimeUnit) {
                throw "Invalid time unit '$TimeUnit' for year data ($($timeUnits -join ', '))"
            }
            $TimeUnit = $TimeUnit.ToUpper()
    
            $startDate = [DateTime]::New($Year, 1, 1)
            $endDate   = $startDate.AddYears(1)

            if (-not $PSBoundParameters.ContainsKey('OutFilePattern')) {
                $OutFilePattern = '%N (%I) %Y %U.csv'
            }

            $placeHolders = @{
                'I' = (0, '')
                'N' = (1, '')
                'Y' = (2, ':yyyy')
                'U' = (3, '')
            }

            $outFileFormat = PatternToFormat $OutFilePattern $placeHolders
        }

        Write-Debug 'Download site details from monitoring platform'
        $details = (Get-SolarEdgeSiteDetails -ApiKey $ApiKey -Site $Site -ErrorAction Stop).siteDetails

        Write-Debug 'Download site energy from monitoring platform'
        $energy = (Get-SolarEdgeSiteEnergy -ApiKey $ApiKey -Site $Site -StartDate $startDate -EndDate $endDate -TimeUnit $TimeUnit -ErrorAction Stop).siteEnergy

        # Drop data points at or after the end date.

        $nvalues = 0
        foreach ($value in $energy.values) {
            if ([DateTime]$value.date -ge $endDate) {
                break;
            }
            $nvalues++
        }

        $values = $energy.values[0..$($nvalues-1)]

        foreach ($_value in $values) {
            if ($null -eq $_value.value) {
                $_value.value = [decimal]'0.0'
            }
        }

        $outFileName = $outFileFormat -f $Site, $details.name, $startDate, $TimeUnit

        Write-Verbose  "Site ID     : $Site"
        Write-Verbose  "Site name   : $($details.name)"
        Write-Verbose ("Start date  : $($solarEdgeDateFormat)" -f $startDate)
        Write-Verbose ("End date    : $($solarEdgeDateFormat)" -f $endDate)
        Write-Verbose  "Time unit   : $TimeUnit"
        Write-Verbose  "Output file : $outFileName"

        $csv = $values | ConvertTo-Csv -NoTypeInformation
        [System.IO.File]::WriteAllLines([System.IO.Path]::Combine((Get-Location), $outFileName), $csv)
    }
}

Export-ModuleMember -Function Export-SolarEdgeInverterData, Export-SolarEdgeSiteEnergy
