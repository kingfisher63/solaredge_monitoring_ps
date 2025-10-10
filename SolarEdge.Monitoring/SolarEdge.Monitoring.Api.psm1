# Copyright (C) 2025 Roger Hunen
#
# Permission to use, copy, modify, and distribute this software and its
# documentation under the terms of the GNU General Public License is hereby 
# granted. No representations are made about the suitability of this software 
# for any purpose. It is provided "as is" without express or implied warranty.
# See the GNU General Public License for more details.

using namespace System.Management.Automation
using namespace System.Text

Set-StrictMode -Version 3.0

$allowedMeterValues        = 'Consumption', 'FeedIn', 'Production', 'Purchased', 'SelfConsumption'
$allowedPeriodLengths      = 'Week', 'Month', 'Year'
$allowedSystemUnitsValues  = 'Imperial', 'Metrics'
$allowedTimeUnitValues     = 'QUARTER_OF_AN_HOUR', 'HOUR', 'DAY', 'WEEK', 'MONTH', 'YEAR'

$errorFormat               = '{0}`n{1}`n    + CategoryInfo          : {2}`n    + FullyQualifiedErrorId : {3}`n'
$monitoringApiServer       = 'monitoringapi.solaredge.com'
$solarEdgeDateFormat       = '{0:yyyy}-{0:MM}-{0:dd}'
$solarEdgeDateTimeFormat   = '{0:yyyy}-{0:MM}-{0:dd} {0:HH}:{0:mm}:{0:ss}'

#
# Support functions (not exported)
#

function AddProperties
{
    param (
        [PSCustomObject] $Object,
        [hashtable]      $Properties
    )

    foreach ($_propertyName in $Properties.Keys) {
        $Object | Add-Member NoteProperty $_propertyName $Properties.$_propertyName -Force
    }
}

function GetApiData
{
    param (
        [String]    $RelativeUrl,
        [String]    $ApiKey,
        [Hashtable] $QueryParams
    )

    $url = "https://${monitoringApiServer}${RelativeUrl}?api_key=${Apikey}"

    if ($null -ne $QueryParams) {
        foreach ($_paramName in $QueryParams.Keys) {
            $url += "&${_paramName}=$([Uri]::EscapeDataString($QueryParams.$_paramName))"
        }
    }

    $json = (Invoke-WebRequest $url).RawContentStream.ToArray()
    $data = [Encoding]::UTF8.GetString($json) | ConvertFrom-Json

    return $data
}

function GetValidatedApiKey
{
    param (
        [String] $ApiKey
    )

    if ($ApiKey -notmatch '^[0-9A-Z]{32}$') {
        throw 'Invalid SolarEdge API key'
    }

    return $ApiKey.ToUpper()
}

function GetValidatedDateOnly
{
    param (
        [String]   $ParamName,
        [DateTime] $Date
    )

    if ($Date.Hour -ne 0 -or $Date.Minute -ne 0 -or  $Date.Second -ne 0 -or $Date.Millisecond -ne 0) {
        throw [ArgumentException]::New("Date-only argument has a non-zero time component.", $ParamName)
    }

    return $Date
}

function GetValidatedSerialNumber
{
    param (
        [String] $SerialNumber
    )

    if ($SerialNumber -notmatch '^[0-9A-F]{8}-[0-9A-F]{2}$') {
        throw "Invalid SolarEdge serial number '$SerialNumber' (invalid characters)."
    }

    $byte1 = [Convert]::ToInt32($SerialNumber.Substring(0,2), 16)
    $byte2 = [Convert]::ToInt32($SerialNumber.Substring(2,2), 16)
    $byte3 = [Convert]::ToInt32($SerialNumber.Substring(4,2), 16)
    $byte4 = [Convert]::ToInt32($SerialNumber.Substring(6,2), 16)
    $check = [Convert]::ToInt32($SerialNumber.Substring(9,2), 16)

    if ($check -ne ($byte1 + $byte2 + $byte3 + $byte4) % 256) {
        throw "Invalid SolarEdge serial number '$SerialNumber' (invalid checksum)."
    }

    return $SerialNumber.ToUpper()
}

function GetValidatedSetValue
{
    param (
        [String]   $ParamName,
        [String]   $Value,
        [String[]] $AllowedValues
    )

    foreach ($_allowedValue in $AllowedValues) {
        if ($_allowedValue -eq $Value) {
            return $_allowedValue
        }
    }

   throw [ArgumentOutOfRangeException]::New($ParamName, "Specified argument was out of the range of valid values ($($AllowedValues -join ', ')).")
}

function GetValidatedSetValueList
{
    param (
        [String]   $ParamName,
        [String[]] $Values,
        [String[]] $AllowedValues
    )

    $ret = @()
    foreach ($_value in $Values) {
        $ret += GetValidatedSetValue $ParamName $_value $AllowedValues
    }

    return $ret -join ','
}

function GetValidatedSiteID
{
    param (
        [String] $Site
    )

    if ($Site -notmatch '^[0-9]+$') {
        throw "Invalid SolarEdge Site ID '$Site'"
    }

    return $Site
}

function GetValidatedTimeUnit
{
    param (
        [String] $TimeUnit
    )

    if ($TimeUnit.ToUpper() -eq '15MIN') {
        return 'QUARTER_OF_AN_HOUR'
    } else {
        return GetValidatedSetValue TimeUnit $TimeUnit $allowedTimeUnitValues
    }
}

function GetValidatedUInt32
{
    param (
        [String] $ParamName,
        [UInt32] $Value,
        [UInt32] $Min,
        [UInt32] $Max
    )

    if (($Value -lt $Min) -or ($Value -gt $Max)) {
        throw [ArgumentOutOfRangeException]::New($ParamName, "Specified argument was out of the range of valid values ($Min-$Max).")
    }

    return $Value
}

function HandleError
{
    param (
        [ErrorRecord] $e
    )

    switch ($ErrorActionPreference) {
        'Continue' {
            $errorFields = $e.Exception.Message,
                           $e.InvocationInfo.PositionMessage,
                           $e.CategoryInfo.ToString(),
                           $e.FullyQualifiedErrorId

            Write-Host -ForegroundColor Red  ($errorFormat -f $errorFields)
        }
        'Ignore' {
        }
        'SilentlyContinue' {
        }
        Default {
            throw $e
        }
    }
}

function ValidateDateTimePeriod
{
    param (
        [String]   $StartDateTimeParam,
        [DateTime] $StartDateTime,
        [String]   $EndDateTimeParam,
        [DateTime] $EndDateTime,
        [String]   $PeriodLength
    )

    if ($EndDateTime -lt $StartDateTime) {
        throw [ArgumentOutOfRangeException]::New(
            "$StartDateTimeParam/$EndDateTimeParam",
            "$EndDateTimeParam is before $StartDateTimeParam"
        )
    }

    switch (GetValidatedSetValue PeriodLength $PeriodLength $allowedPeriodLengths) {
        'Week'  { $limitDateTime = $StartDateTime.AddDays(7)   }
        'Month' { $limitDateTime = $StartDateTime.AddMonths(1) }
        'Year'  { $limitDateTime = $StartDateTime.AddYears(1)  }
    }

    if ($EndDateTime -gt $limitDateTime) {
        throw [ArgumentOutOfRangeException]::New("$StartDateParam/$EndDateParam", "Too many days between $StartDateParam and $EndDateParam (max. $(($limitDateTime - $StartDateTime).TotalDays))")
    }
}

#
# Exported functions
#

function Get-SolarEdgeApiInfo
{
    <#
        .SYNOPSIS
        Gets the current and supported SolarEdge Monitoring API versions.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .OUTPUTS
        System.Management.Automation.PSCustomObject
        .LINK
        Write-SolarEdgeApiInfo
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)] [String] $ApiKey
    )

    process {
        try {
            $ApiKey = GetValidatedApiKey $ApiKey

            $version   = GetApiData "/version/current.json"   $ApiKey
            $supported = GetApidata "/version/supported.json" $ApiKey

            $ai = [PSCustomObject]@{
                apiInfo = [PSCustomObject]@{
                    version   = $version.version
                    supported = $supported.supported
                }
            }

            Write-Output $ai
        }
        catch {
            HandleError $_
        }
    }
}

function Get-SolarEdgeEquipmentChangeLog
{
    <#
        .SYNOPSIS
        Gets the equipment change log from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER SerialNumber
        The inverter serial number.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [String]   $SerialNumber
    )

    begin {
        $ApiKey       = GetValidatedApiKey       $ApiKey
        $SerialNumber = GetValidatedSerialNumber $SerialNumber
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/equipment/$(GetValidatedSiteID $_site)/$SerialNumber/changeLog.json" $ApiKey

                $ecl = [PSCustomObject]@{
                    siteId             = $_site
                    serialNumber       = $SerialNumber
                    equipmentChangeLog = $data.ChangeLog
                }

                Write-Output $ecl
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeInverterData
{
    <#
        .SYNOPSIS
        Gets inverter technical data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndTime
        The end date & time.
        .PARAMETER SerialNumber
        The inverter serial number.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [String]   $SerialNumber,
        [parameter(Mandatory,Position=3)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=4)]                   [DateTime] $EndTime
    )

    begin {
        $ApiKey       = GetValidatedApiKey       $ApiKey
        $SerialNumber = GetValidatedSerialNumber $SerialNumber

        ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Week

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/equipment/$(GetValidatedSiteID $_site)/$SerialNumber/data.json" $ApiKey $queryParams

                $id = [PSCustomObject]@{
                    siteId       = $_site
                    serialNumber = $SerialNumber
                    inverterData = $data.data
                }
                AddProperties $id $queryParams

                Write-Output $id
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeMeterData
{
    <#
        .SYNOPSIS
        Gets meter data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndDate
        The end date & time.
        .PARAMETER TimeUnit
        The granularity of the meter data. Valid values are QUARTER_OF_AN_HOUR, HOUR,
        DAY, WEEK, MONTH and YEAR as well as 15MIN (alias for QUARTER_OF_AN_HOUR).
        .PARAMETER Meters
        The list of meters to include in the data. Valid meter names are Consumption,
        FeedIn, Production, Purchased and SelfConsumption.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndTime,
                                                            [String]   $TimeUnit = 'DAY',
                                                            [String[]] $Meters
    )

    begin {
        $ApiKey   = GetValidatedApiKey   $ApiKey
        $TimeUnit = GetValidatedTimeUnit $TimeUnit

        switch ($TimeUnit) {
            'DAY'                { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Year  }
            'HOUR'               { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month }
            'QUARTER_OF_AN_HOUR' { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month }
        }

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
            timeUnit  = $TimeUnit
        }

        if ($PSBoundParameters.ContainsKey('Meters')) {
            $queryParams['meters'] = GetValidatedSetValueList Meters $Meters $allowedMeterValues
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/meters.json" $ApiKey $queryParams

                $md = [PSCustomObject]@{
                    siteId    = $_site
                    meterData = $data.meterEnergyDetails
                }
                AddProperties $md $queryParams

                Write-Output $md
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSensorData
{
    <#
        .SYNOPSIS
        Gets sensor data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartDate
        The start date.
        .PARAMETER EndDate
        The end date.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartDate,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndDate
    )

    begin {
        $ApiKey    = GetValidatedApiKey   $ApiKey
        $StartDate = GetValidatedDateOnly StartDate $StartDate
        $EndDate   = GetValidatedDateOnly EndDate   $EndDate

        ValidateDateTimePeriod StartDate $StartDate EndDate $EndDate Year

        $queryParams = @{
            startDate = $solarEdgeDateFormat -f $StartDate
            endDate   = $solarEdgeDateFormat -f $EndDate
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/equipment/$(GetValidatedSiteID $_site)/sensors.json" $ApiKey $queryParams | Write-Output

                $sd = [PSCustomObject]@{
                    siteId     = $_site
                    sensorData = $data.SiteSensors
                }
                AddProperties $sd $queryParams

                Write-Output $sd
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSensorList
{
    <#
        .SYNOPSIS
        Gets the list of sensors in a site from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/equipment/$(GetValidatedSiteID $_site)/sensors.json" $ApiKey

                $ssl = [PSCustomObject]@{
                    siteId     = $_site
                    sensorList = $data.SiteSensors
                }

                Write-Output $ssl
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSiteDataPeriod
{
    <#
        .SYNOPSIS
        Gets the site data period from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteDataPeriod
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
        $sites  = @()
    }

    process {
        foreach ($_site in $Site) {
            try {
                $sites += GetValidatedSiteID $_site
            }
            catch {
                HandleError $_
            }
        }
    }

    end {
        if ($sites.Length -ne 0) {
            $data = GetApiData "/sites/$($sites -join ',')/dataPeriod.json" $ApiKey

            foreach ($_siteDataPeriod in $data.datePeriodList.siteEnergyList) {
                $sdp = [PSCustomObject]@{
                    siteId         = $_siteDataPeriod.siteId
                    siteDataPeriod = $_siteDataPeriod.dataPeriod
                }

                Write-Output $sdp
            }
        }
    }
}

function Get-SolarEdgeSiteDetails
{
    <#
        .SYNOPSIS
        Gets the site details from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteDetails
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/details.json" $ApiKey

                $sd = [PSCustomObject]@{
                    siteId      = $_site
                    siteDetails = $data.details
                }

                Write-Output $sd
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSiteEnergy
{
    <#
        .SYNOPSIS
        Gets site energy data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartDate
        The start date.
        .PARAMETER EndDate
        The end date.
        .PARAMETER TimeUnit
        The time granularity of the site energy data. Valid values are QUARTER_OF_AN_HOUR,
        HOUR, DAY, WEEK, MONTH and YEAR as well as 15MIN (alias for QUARTER_OF_AN_HOUR).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteEnergy
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartDate,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndDate,
        [parameter()]                                       [String]   $TimeUnit = 'DAY'
    )

    begin {
        $ApiKey    = GetValidatedApiKey   $ApiKey
        $StartDate = GetValidatedDateOnly StartDate $StartDate
        $EndDate   = GetValidatedDateOnly EndDate   $EndDate
        $TimeUnit  = GetValidatedTimeUnit $TimeUnit

        switch ($TimeUnit) {
            'DAY'                { ValidateDateTimePeriod StartDate $StartDate EndDate $EndDate Year  }
            'HOUR'               { ValidateDateTimePeriod StartDate $StartDate EndDate $EndDate Month }
            'QUARTER_OF_AN_HOUR' { ValidateDateTimePeriod StartDate $StartDate EndDate $EndDate Month }
        }

        $queryParams = @{
            startDate = $solarEdgeDateFormat -f $StartDate
            endDate   = $solarEdgeDateFormat -f $EndDate
            timeUnit  = $TimeUnit
        }

        $sites = @()
    }

    process {
        foreach ($_site in $Site) {
            try {
                $sites += GetValidatedSiteID $_site
            }
            catch {
                HandleError $_
            }
        }
    }

    end {
        if ($sites.Length -ne 0) {
            $data = GetApiData "/sites/$($sites -join ',')/energy.json" $ApiKey $queryParams

            foreach ($_siteEnergy in $data.sitesEnergy.siteEnergyList) {
                $se = [PSCustomObject]@{
                    siteId     = $_siteEnergy.siteId
                    siteEnergy = [PSCustomObject]@{
                        unit   = $data.sitesEnergy.unit
                        values = $_siteEnergy.energyValues.values
                    }
                }
                AddProperties $se $queryParams

                Write-Output $se
            }
        }
    }
}

function Get-SolarEdgeSiteEnergyDetails
{
    <#
        .SYNOPSIS
        Gets detailed site energy data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndTime
        The end date & time.
        .PARAMETER TimeUnit
        The granularity of the site energy data. Valid values are QUARTER_OF_AN_HOUR,
        HOUR, DAY, WEEK, MONTH and YEAR as well as 15MIN (alias for QUARTER_OF_AN_HOUR).
        .PARAMETER Meters
        The list of meters to include in the data. Valid meter names are Consumption,
        FeedIn, Production, Purchased and SelfConsumption.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteEnergyDetails
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndTime,
                                                            [String]   $TimeUnit = 'DAY',
                                                            [String[]] $Meters
    )

    begin {
        $ApiKey   = GetValidatedApiKey   $ApiKey
        $TimeUnit = GetValidatedTimeUnit $TimeUnit

        switch ($TimeUnit) {
            'DAY'                { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Year  }
            'HOUR'               { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month }
            'QUARTER_OF_AN_HOUR' { ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month }
        }

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
            timeUnit  = $TimeUnit
        }

        if ($PSBoundParameters.ContainsKey('Meters')) {
            $queryParams['meters'] = GetValidatedSetValueList Meters $Meters $allowedMeterValues
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/energyDetails.json" $ApiKey $queryParams | Write-Output

                $sed = [PSCustomObject]@{
                    siteId            = $_site
                    siteEnergyDetails = $data.energyDetails
                }
                AddProperties $sed $queryParams

                Write-Output $sed
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSiteEnergySummary
{
    <#
        .SYNOPSIS
        Gets the total site energy from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartDate
        The start date.
        .PARAMETER EndDate
        The end date.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteEnergySummary
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartDate,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndDate
    )

    begin {
        $ApiKey    = GetValidatedApiKey   $ApiKey
        $StartDate = GetValidatedDateOnly StartDate $StartDate
        $EndDate   = GetValidatedDateOnly EndDate   $EndDate

        ValidateDateTimePeriod StartDate $StartDate EndDate $EndDate Year

        $queryParams = @{
            startDate = $solarEdgeDateFormat -f $StartDate
            endDate   = $solarEdgeDateFormat -f $EndDate
        }

        $sites = @()
    }

    process {
        foreach ($_site in $Site) {
            try {
                $sites += GetValidatedSiteID $_site
            }
            catch {
                HandleError $_
            }
        }
    }

    end {
        if ($sites.Length -ne 0) {
            $data = GetApiData "/sites/$($sites -join ',')/timeFrameEnergy.json" $ApiKey $queryParams
            
            foreach ($_timeFrameEnergy in $data.timeFrameEnergyList.timeFrameEnergyList) {
                $ses = [PSCustomObject]@{
                    siteId            = $_timeFrameEnergy.siteId
                    siteEnergySummary = $_timeFrameEnergy.timeFrameEnergy
                }
                AddProperties $ses $queryParams

                Write-Output $ses
            }
        }
    }
}

function Get-SolarEdgeSiteEnvBenefits
{
    <#
        .SYNOPSIS
        Gets the environmental benefits from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER SystemUnits
        The reporting data units. Valid values are Imperial and Metrics.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteEnvBenefits
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
                                                            [String]   $SystemUnits
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey

        $queryParams = @{}

        if ($PSBoundParameters.ContainsKey('SystemUnits')) {
            $queryParams['systemUnits'] = GetValidatedSetValue SystemUnits $SystemUnits $allowedSystemUnitsValues
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/envBenefits.json" $ApiKey $queryParams | Write-Output

                $seb = [PSCustomObject]@{
                    siteId          = $_site
                    siteEnvBenefits = $data.envBenefits
                }
                AddProperties $seb $queryParams

                Write-Output $seb
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSiteInventory
{
    <#
        .SYNOPSIS
        Gets the site inventory from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteInventory
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/inventory.json" $ApiKey

                $si = [PSCustomObject]@{
                    siteId        = $_site
                    siteInventory = $data.inventory
                }

                Write-Output $si
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSiteList
{
    <#
        .SYNOPSIS
        Gets the site list from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteDetails
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)] [String] $ApiKey
    )

    process {
        $ApiKey = GetValidatedApiKey $ApiKey
        $data   = GetApiData "/sites/list.json" $ApiKey

        foreach ($_site in $data.sites.site) {
            $sd = [PSCustomObject]@{
                siteId      = $_site.id
                siteDetails = $_site
            }

            Write-Output $sd
        }
    }
}

function Get-SolarEdgeSiteOverview
{
    <#
        .SYNOPSIS
        Gets the site overview from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSiteOverview
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
        $sites  = @()
    }

    process {
        foreach ($_site in $Site) {
            try {
                $sites += GetValidatedSiteID $_site
            }
            catch {
                HandleError $_
            }
        }
    }

    end {
        if ($sites.Length -ne 0) {
            $data = GetApiData "/sites/$($sites -join ',')/overview.json" $ApiKey

            foreach ($_siteOverview in $data.sitesOverviews.siteEnergyList) {
                $so = [PSCustomObject]@{
                    siteId       = $_siteOverview.siteId
                    siteOverview = $_siteOverview.siteOverview
                }

                Write-Output $so
            }
        }
    }
}

function Get-SolarEdgeSitePower
{
    <#
        .SYNOPSIS
        Gets site power data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndTime
        The end date & time.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSitePower
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndTime
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey

        ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
        }

        $sites = @()
    }

    process {
        foreach ($_site in $Site) {
            try {
                $sites += GetValidatedSiteID $_site
            }
            catch {
                HandleError $_
            }
        }
    }

    end {
        if ($sites.Length -ne 0) {
            $data = GetApiData "/sites/$($sites -join ',')/power.json" $ApiKey $queryParams

            foreach ($_sitePower in $data.powerDateValuesList.siteEnergyList) {
                $sp = [PSCustomObject]@{
                    siteId    = $_sitePower.siteId
                    sitePower = [PSCustomObject]@{
                        timeunit = $data.powerDateValuesList.timeunit
                        unit     = $data.powerDateValuesList.unit
                        values   = $_sitePower.powerDataValueSeries.values
                    }
                }
                AddProperties $sp $queryParams

                Write-Output $sp
            }
        }
    }
}

function Get-SolarEdgeSitePowerDetails
{
    <#
        .SYNOPSIS
        Gets detailed site power data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndTime
        The end date & time.
        .PARAMETER Meters
        The list of meters to include in the data. Valid meter names are Consumption,
        FeedIn, Production, Purchased and SelfConsumption.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSitePowerDetails
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndTime,
                                                            [String[]] $Meters
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey

        ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Month

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
        }

        if ($PSBoundParameters.ContainsKey('Meters')) {
            $queryParams['meters'] = GetValidatedSetValueList Meters $Meters $allowedMeterValues
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/powerDetails.json" $ApiKey $queryParams

                $spd = [PSCustomObject]@{
                    siteId           = $_site
                    sitePowerDetails = $data.powerDetails
                }
                AddProperties $spd $queryParams

                Write-Output $spd
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeSitePowerFlow
{
    <#
        .SYNOPSIS
        Gets the current site power flow from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Write-SolarEdgeSitePowerFlow
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey
    }

    process {
        foreach ($_site in $Site) {
            try {                
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/currentPowerFlow.json" $ApiKey

                $spf = [PSCustomObject]@{
                    siteId        = $_site
                    sitePowerFlow = $data.siteCurrentPowerFlow
                }

                Write-Output $spf
            }
            catch {
                HandleError $_
            }
        }
    }
}

function Get-SolarEdgeStorageData
{
    <#
        .SYNOPSIS
        Gets site energy data from the SolarEdge monitoring platform.
        .PARAMETER ApiKey
        The SolarEdge API key (32 characters 0-9 or A-Z).
        .PARAMETER Site
        The SolarEdge Site ID (1+ characters 0-9).
        .PARAMETER StartTime
        The start date & time.
        .PARAMETER EndTime
        The end date & time.
        .PARAMETER SerialNumbers
        The list of serial numbers of the batteries to include in the data.
        .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        None
    #>

    [CmdletBinding(PositionalBinding=$false)]

    param (
        [parameter(Mandatory,Position=0)]                   [String]   $ApiKey,
        [parameter(Mandatory,Position=1,ValueFromPipeline)] [String[]] $Site,
        [parameter(Mandatory,Position=2)]                   [DateTime] $StartTime,
        [parameter(Mandatory,Position=3)]                   [DateTime] $EndTime,
                                                            [String[]] $SerialNumbers
    )

    begin {
        $ApiKey = GetValidatedApiKey $ApiKey

        ValidateDateTimePeriod StartTime $StartTime EndTime $EndTime Week

        $queryParams = @{
            startTime = $solarEdgeDateTimeFormat -f $StartTime
            endTime   = $solarEdgeDateTimeFormat -f $EndTime
        }

        if ($PSBoundParameters.ContainsKey('SerialNumbers')) {
            $params['serials'] = $SerialNumbers -join ','
        }
    }

    process {
        foreach ($_site in $Site) {
            try {
                $data = GetApiData "/site/$(GetValidatedSiteID $_site)/storageData.json" $ApiKey $queryParams

                $sd = [PSCustomObject]@{
                    siteId      = $_site
                    storageData = $data.storageData
                }
                AddProperties $sd $queryParams

                Write-Output $sd
            }
            catch {
                HandleError $_
            }
        }
    }
}
