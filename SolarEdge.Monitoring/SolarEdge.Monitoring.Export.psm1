# Copyright (C) 2025 Roger Hunen

using namespace System.IO

Set-StrictMode -Version 3.0

$yearMin        = 2000
$yearMax        = 2099
$yearTimeUnits  = 'DAY', 'MONTH', 'YEAR'

$monthMin       = 1
$monthMax       = 12
$monthTimeUnits = '15MIN', 'QUARTER_OF_AN_HOUR', 'HOUR', 'DAY', 'MONTH'

# 0: date
$solarEdgeDateFormat = '{0:yyyy}-{0:MM}-{0:dd}'

# Export-SolarEdgeInverterData
#
# 0: site ID
# 1: inverter serial number
# 2: date
$inverterFileNameFormat = '{2:yyyy}{2:MM}{2:dd} {1} ({0}).csv'

# Export-SolarEdgeSiteEnergy
#
# 0: site ID
# 1: site name
# 2: year
# 3: month (not valid for year data)
# 4: time unit
$monthFileNameFormat = '{2:D4}-{3:D2} - {1} ({0}) - {4}.csv'                       
$yearFileNameFormat  = '{2:D4} - {1} ({0}) - {4}.csv'

#
# Support functions (not exported)
#

function AddProperty
{
    param (
        [PSCustomObject] $Object,
        [String]         $PropertyName,
        [String]         $PropertyValue
    )

    $Object | Add-Member NoteProperty -Name $PropertyName -Value $PropertyValue -Force
}

function ValueOrZero
{
    param(
        [Object] $Object,
        [String] $PropertyName
    )

    if ($Object | Get-Member -Name $PropertyName) {
        return $Object.$PropertyName
    } else {
        return [Decimal]'0.0'
    }
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
        .LINK
        Get-SolarEdgeInverterData
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [Parameter(Mandatory,Position=0)] [string]   $ApiKey,
        [Parameter(Mandatory,Position=1)] [string]   $Site,
        [Parameter(Mandatory,Position=2)] [string]   $SerialNumber,
        [parameter(Mandatory,Position=3)] [DateTime] $StartDate,
        [parameter(Mandatory,Position=4)] [DateTime] $EndDate
    )

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

        Write-Verbose "Site ID     : $Site"
        Write-Verbose "Start date  : $($solarEdgeDateFormat -f $StartDate)"
        Write-Verbose "Days        : $totalDays"

        while ($totalDays -gt 0) {
            $queryDays    = if ($totalDays -lt 7 ) { $totalDays } else { 7 }
        	$queryEndDate = $queryStartDate.AddDays($queryDays)

            $inverterData = (Get-SolarEdgeInverterData $apiKey $site $serialNumber $queryStartDate $queryEndDate -ErrorAction Stop).inverterData

	        Write-Verbose "Start date  : $($solarEdgeDateFormat -f $queryStartDate)"
	        Write-Verbose "End date    : $($solarEdgeDateFormat -f $queryEndDate)"
            Write-Verbose "Query days  : $queryDays"
            Write-Verbose "Data points : $($inverterData.telemetries.Count)"

            if ($inverterData.telemetries.Count -eq 0) {
                break
            }

            for ($day=0; $day -lt $queryDays; $day++) {
                $date    = $queryStartDate.AddDays($day)
                $dateStr = $solarEdgeDateFormat -f $date
                $outFile = $inverterFileNameFormat -f $Site, $SerialNumber, $date

                $telemetries = $inverterData.telemetries | Where-Object { ([DateTime]$_.date).Date -eq $date }
                if ($null -eq $telemetries) {
                    continue
                }

                $outItems = @()
                foreach ($telemetry in $telemetries) {
                    $outItem = [PSCustomObject]@{
                        date                  = $telemetry.date
                        totalActivePower      = $telemetry.totalActivePower
                        dcVoltage             = ValueOrZero $telemetry dcVoltage
                        groundFaultResistance = ValueOrZero $telemetry groundFaultResistance
                        powerLimit            = $telemetry.powerLimit
                        totalEnergy           = $telemetry.totalEnergy
                        temperature           = $telemetry.temperature
                        inverterMode          = $telemetry.inverterMode
                        operationMode         = $telemetry.operationMode
                    }

                    foreach ($phase in 'L1','L2','L3') {
                        $phaseDataProperty = "${phase}Data"

                        if ($telemetry | Get-Member -Name $phaseDataProperty) {
                            AddProperty $outItem "${phase}_acCurrent"     $telemetry.$phaseDataProperty.acCurrent
                            AddProperty $outItem "${phase}_acVoltage"     $telemetry.$phaseDataProperty.acVoltage
                            AddProperty $outItem "${phase}_acFrequency"   $telemetry.$phaseDataProperty.acFrequency
                            AddProperty $outItem "${phase}_apparentPower" $telemetry.$phaseDataProperty.apparentPower
                            AddProperty $outItem "${phase}_activePower"   $telemetry.$phaseDataProperty.activePower
                            AddProperty $outItem "${phase}_reactivePower" $telemetry.$phaseDataProperty.reactivePower
                            AddProperty $outItem "${phase}_cosPhi"        $telemetry.$phaseDataProperty.cosPhi
                        }
                    }

                    $outItems += $outItem
                }

                Write-Verbose ("{0}  : {1,3} -> {2}" -f $dateStr, $telemetries.Count, $outFile)

                $csv = $outItems | ConvertTo-Csv -NoTypeInformation
                [File]::WriteAllLines([Path]::Combine((Get-Location), $Outfile), $csv)
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
        The year for which site energy data is exported (2000-2099). The default
        value is the current year.
        .PARAMETER Month
        The month for which site energy data is exported (1-12). If the month
        is not specified, site energy data is exported for the whole year.
        .PARAMETER TimeUnit
        The time granularity of the site energy data. DAY, WEEK, MONTH and YEAR
        are  accepted when exporting data for a year. 15MIN, HOUR, DAY, WEEK and
        MONTH are accepted When exporting data for a month. The default value
        is DAY.
        .LINK
        Get-SolarEdgeSiteEnergy
    #>

    [CmdletBinding(DefaultParameterSetName='Year',PositionalBinding=$false)]

    param (
        [Parameter(Mandatory,Position=0)]     [string] $ApiKey,
        [Parameter(Mandatory,Position=1)]     [string] $Site,
                                              [int]    $Year = [DateTime]::Now.Year,
        [Parameter(ParameterSetName='Month')] [int]    $Month,
                                              [string] $TimeUnit = 'DAY'
    )
    
    process {
        if (($Year -lt $yearMin) -or ($Year -gt $yearMax)) {
            throw "Year '$Year' is out of range ($yearMin-$yearMax)"
        }

        # There is a flaw in the SolarEdge Monitoring API. In some cases data points
        # at the end date are included, in other cases (QUARTER_OF_AN_HOUR, HOUR) they
        # are not. We therefore set the end date at the 1st day of the next year/month
        # and drop data points at the end date later. (RH 2025/10/05)

        if ($PSCmdlet.ParameterSetName -eq 'Year') {
            if ($yearTimeUnits -notcontains $TimeUnit) {
                throw "Invalid time unit '$TimeUnit' for year data ($($yearTimeUnits -join ', '))"
            }
        
            $startDate = [DateTime]::New($Year, 1, 1)
            $endDate   = $startDate.AddYears(1)
        } else {
            if (($Month -lt $monthMin) -or ($Month -gt $monthMax)) {
                throw "Month '$Month' is out of range ($monthMin-$monthMax)"
            }
            if ($monthTimeUnits -notcontains $TimeUnit) {
                throw "Invalid time unit '$TimeUnit' for month data ($($monthTimeUnits -join ', '))"
            }

            $startDate = [DateTime]::New($Year, $Month, 1)
            $endDate   = $StartDate.AddMonths(1)
        }
		
        $TimeUnit = $TimeUnit.ToUpper()

        Write-Debug 'Download site details from monitoring platform'
        $details = (Get-SolarEdgeSiteDetails -ApiKey $ApiKey -Site $Site -ErrorAction Stop).siteDetails

        if ($PSCmdlet.ParameterSetName -eq 'Year') {
            $Outfile = $yearFileNameFormat  -f $Site, $details.name, $Year, 0,      $TimeUnit
        } else {
            $Outfile = $monthFileNameFormat -f $Site, $details.name, $Year, $Month, $TimeUnit
        }

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

        Write-Verbose  "Time span   : $($PSCmdlet.ParameterSetName.ToUpper())"
        Write-Verbose ("Start date  : {0:yyyy}-{0:MM}-{0:dd}" -f $startDate)
        Write-Verbose ("End date    : {0:yyyy}-{0:MM}-{0:dd}" -f $endDate)
        Write-Verbose  "Time unit   : $TimeUnit"
        Write-Verbose  "Site ID     : $Site"
        Write-Verbose  "Output file : $Outfile"

        $csv = $values | ConvertTo-Csv -NoTypeInformation
        [File]::WriteAllLines([Path]::Combine((Get-Location), $Outfile), $csv)
    }
}

Export-ModuleMember -Function Export-SolarEdgeInverterData, Export-SolarEdgeSiteEnergy
