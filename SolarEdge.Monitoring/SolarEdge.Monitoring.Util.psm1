# Copyright (C) 2025-2026 Roger Hunen
#
# Permission to use, copy, modify, and distribute this software and its
# documentation under the terms of the GNU General Public License is hereby 
# granted. No representations are made about the suitability of this software 
# for any purpose. It is provided "as is" without express or implied warranty.
# See the GNU General Public License for more details.

using namespace System.Management.Automation

Set-StrictMode -Version 3.0

$powerFlowElements = 'GRID', 'LOAD', 'PV', 'STORAGE'

#
# Support functions (not exported)
#

function PropertyExistsAndIsNotNull
{
    param (
        [Object] $object,
        [String] $propertyName
    )

    if ($object | Get-Member -Name $propertyName) {
        return ($null -ne $object.$propertyName)
    } else {
        return $false
    }
}

function WriteTable
{
    param (
        [System.Data.DataTable] $table
    )

    $columnWidth = [int[]]::new($table.Columns.Count)
    
    foreach ($row in $table.Rows) {
        for ($i = 0; $i -lt $columnWidth.Length; $i++) {
            if ($columnWidth[$i] -lt $row[$i].Length) {
                $columnWidth[$i] =   $row[$i].Length
            }
        }
    }

    foreach ($row in $table.Rows) {
        $line = ''

        for ($i = 0; $i -lt $columnWidth.Length; $i++) {
            if ($i -ne 0) {
                $line += '  '
            }

            $line += "{0, -$($columnWidth[$i])}" -f $row[$i]
        }

        Write-Output $line
    }
}

#
# Exported functions
#

function Write-SolarEdgeMeterData
{
    <#
        .SYNOPSIS
        Writes the SolarEdge meter data to Output as text.
        .PARAMETER MeterData
        The SolarEdge meter data.
        .PARAMETER OmitHeader
        Omit table headers. The output will contain only timestamps and meter readings.
        .INPUTS
        System.Management.Automation.PSCustomObject
        .LINK
        Get-SolarEdgeMeterData
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $MeterData,
        [Parameter()]                            [switch]           $OmitHeader
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_meterData in $MeterData) {
            if (-not (PropertyExistsAndIsNotNull $_meterData meterData)) {
                throw "Invalid MeterData object (property 'meterData' does not exist or is null)"
            }

            $meterDataTable = $null

            foreach ($_meter in $_meterData.meterData.meters) {
                $dateColumn               = [System.Data.DataColumn]::new('Date')
                $dateColumn.DataType      = [System.String]

                $valueColumn              = [System.Data.DataColumn]::new($_meter.meterType)
                $valueColumn.DataType     = [System.String]
                $valueColumn.DefaultValue = ''

                $meterTable               = [System.Data.DataTable]::new()
                $meterTable.Columns.Add($dateColumn)
                $meterTable.Columns.Add($valueColumn)
                $meterTable.PrimaryKey    = ($dateColumn)

                if (-not $OmitHeader) {
                    [void] $meterTable.Rows.Add('Meter',                  $_meter.meterType)
                    [void] $meterTable.Rows.Add('Meter model',            $_meter.model)
                    [void] $meterTable.Rows.Add('Meter serial number',    $_meter.meterSerialNumber)
                    [void] $meterTable.Rows.Add('Inverter serial number', $_meter.connectedSolaredgeDeviceSN)
                    [void] $meterTable.Rows.Add('--',                     '--')
                }

                foreach ($_meterValue in $_meter.values) {
                    [void] $meterTable.Rows.Add($_meterValue.date, $_meterValue.value.ToString('F1'))
                }

                if ($null -eq $meterDataTable) {
                    $meterDataTable = $meterTable
                } else {
                    $meterDataTable.Merge($meterTable)
                }
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            if (-not $OmitHeader) {
                Write-Output "Site ID      $($_meterData.siteId)"
                Write-Output "Start time   $($_meterData.startTime)"
                Write-Output "End time     $($_meterData.endTime)"
                Write-Output "Time unit    $($_meterData.timeUnit)"
                Write-Output "Energy unit  $($_meterData.meterData.unit)"
                Write-Output '---'
            }

            if ($null -eq $meterDataTable) {
                Write-Output "No meter data available."
            } else {
                WriteTable $meterDataTable
            }
        }
    }
}

function Write-SolarEdgeSiteDataPeriod
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site data period to Output as text.
        .PARAMETER SiteDataPeriod
        The SolarEdge site data period.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteDataPeriod
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteDataPeriod
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteDataPeriod in $SiteDataPeriod) {
            if (-not (PropertyExistsAndIsNotNull $_siteDataPeriod siteDataPeriod)) {
                throw "Invalid SiteDataPeriod object (property 'siteDataPeriod' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            Write-Output "Site ID    : $($_siteDataPeriod.siteId)"
            Write-Output "Start date : $($_siteDataPeriod.siteDataPeriod.startDate)"
            Write-Output "End date   : $($_siteDataPeriod.siteDataPeriod.endDate)"
        }
    }
}

function Write-SolarEdgeSiteDetails
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site details to Output as text.
        .PARAMETER SiteDetails
        The SolarEdge site details.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteDetails
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteDetails
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteDetails in $SiteDetails) {
            if (-not (PropertyExistsAndIsNotNull $_siteDetails siteDetails)) {
                throw "Invalid SiteDetails object (property 'siteDetails' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $details  = $_siteDetails.siteDetails
            $location = $details.location
            $module   = $details.primaryModule
            $isPublic = if ($details.publicSettings.isPublic) {'Yes'} else {'No'}

                Write-Output "Site ID        : $($details.id)"
                Write-Output "Site name      : $($details.name)"
                Write-Output "Account ID     : $($details.accountId)"
                Write-Output "Type           : $($details.type)"
                Write-Output "Peak power     : $($details.peakPower) kW"
                Write-Output "Last update    : $($details.lastUpdateTime)"
                Write-Output "Install date   : $($details.installationDate)"
            if ($details.ptoDate) {
                Write-Output "PTO date       : $($details.ptoDate)"
            }
            if ($details.notes) {
                Write-Output "Notes          : $($details.notes)"
            }
            if ($location.address) {
                Write-Output "Address        : $($location.address), $($location.zip) $($location.city), $($location.country)"
            }
            if ($location.latitude) {
                Write-Output "Coordinates    : $($location.latitude) $($location.longitude)"
            }
            if ($location.timeZone) {
                Write-Output "Timezone       : $($location.timeZone)"
            }
            if ($module.manufacturerName) {
                Write-Output "Primary module : $($module.manufacturerName), $($module.modelName), $($module.maximumPower) Wp, $($module.temperatureCoef) %/C"
            }
            if (PropertyExistsAndIsNotNull $details status) {
                Write-Output "Status         : $($details.status)"
            }
            if (PropertyExistsAndIsNotNull $details alertQuantity) {
                Write-Output "Alert quantity : $($details.alertQuantity)"
            }
            if (PropertyExistsAndIsNotNull $details highestImpact) {
                Write-Output "Highest impact : $($details.highestImpact)"
            }
                Write-Output "Public site    : $($isPublic)"
        }
    }
}

function Write-SolarEdgeSiteEnergy
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site energy data to Output as text.
        .PARAMETER SiteEnergy
        The SolarEdge site energy data.
        .PARAMETER OmitHeader
        Omit table headers. The output will contain only timestamps and energy readings.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteEnergy
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnergy,
        [Parameter()]                            [switch]           $OmitHeader
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnergy in $SiteEnergy) {
            if (-not (PropertyExistsAndIsNotNull $_siteEnergy siteEnergy)) {
                throw "Invalid SiteEnergy object (property 'siteEnergy' does not exist or is null)"
            }

            $dateColumn           = [System.Data.DataColumn]::new('Date')
            $dateColumn.DataType  = [System.String]

            $valueColumn          = [System.Data.DataColumn]::new('Value')
            $valueColumn.DataType = [System.String]

            $energyTable      = [System.Data.DataTable]::new()
            $energyTable.Columns.Add($dateColumn)
            $energyTable.Columns.Add($valueColumn)

            if (-not $OmitHeader) {
                [void] $energyTable.Rows.Add('Date', 'Energy')
                [void] $energyTable.Rows.Add('---',   '--')
            }

            foreach ($_value in $_siteEnergy.siteEnergy.values) {
                [void] $energyTable.Rows.Add($_value.date, $_value.value.ToString('F1'))
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            if (-not $OmitHeader) {
                Write-Output "Site ID      $($_siteEnergy.siteId)"
                Write-Output "Start date   $($_siteEnergy.startDate)"
                Write-Output "End date     $($_siteEnergy.endDate)"
                Write-Output "Time unit    $($_siteEnergy.timeUnit)"
                Write-Output "Energy unit  $($_siteEnergy.siteEnergy.unit)"
                Write-Output '---'
            }

            WriteTable $energyTable
        }
    }
}

function Write-SolarEdgeSiteEnergyDetails
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site energy details to Output as text.
        .PARAMETER EnergyDetails
        The SolarEdge energy details.
        .PARAMETER OmitHeader
        Omit table headers. The output will contain only timestamps and energy readings.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteEnergyDetails
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnergyDetails,
        [Parameter()]                            [switch]           $OmitHeader
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnergyDetails in $SiteEnergyDetails) {
            if (-not (PropertyExistsAndIsNotNull $_siteEnergyDetails siteEnergyDetails)) {
                throw "Invalid SiteEnergyDetails object (property 'siteEnergyDetails' does not exist or is null)"
            }

            $energyTable = $null

            foreach ($_meter in $_siteEnergyDetails.siteEnergyDetails.meters) {
                $meterHasData = $false
                foreach ($_value in $_meter.values) {
                    if (PropertyExistsAndIsNotNull $_value value) {
                        $meterHasData = $true
                        break
                    }
                }

                if ($meterHasData) {
                    $dateColumn               = [System.Data.DataColumn]::new('Date')
                    $dateColumn.DataType      = [System.String]

                    $valueColumn              = [System.Data.DataColumn]::new($_meter.type)
                    $valueColumn.DataType     = [System.String]
                    $valueColumn.DefaultValue = ''

                    $meterTable               = [System.Data.DataTable]::new()
                    $meterTable.Columns.Add($dateColumn)
                    $meterTable.Columns.Add($valueColumn)
                    $meterTable.PrimaryKey    = ($dateColumn)

                    if (-not $OmitHeader) {
                        [void] $meterTable.Rows.Add('Date', $_meter.type)
                        [void] $meterTable.Rows.Add('--',   '--')
                    }

                    foreach ($_value in $_meter.values) {
                        [void] $meterTable.Rows.Add($_value.date, $_value.value.ToString('F1'))
                    }

                    if ($null -eq $energyTable) {
                        $energyTable = $meterTable
                    } else {
                        $energyTable.Merge($meterTable)
                    }
                }
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            if (-Not $OmitHeader) {
                Write-Output "Site ID      $($_siteEnergyDetails.siteId)"
                Write-Output "Start time   $($_siteEnergyDetails.startTime)"
                Write-Output "End time     $($_siteEnergyDetails.endTime)"
                Write-Output "Time unit    $($_siteEnergyDetails.timeUnit)"
                Write-Output "Energy unit  $($_siteEnergyDetails.siteEnergyDetails.unit)"
                Write-Output '---'
            }

            if ($null -eq $energyTable) {
                Write-Output "No energy data available."
            } else {
                WriteTable $energyTable
            }
        }
    }
}

function Write-SolarEdgeSiteEnergySummary
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site energy summary (time frame energy) to Output as text.
        .PARAMETER SiteEnergySummary
        The SolarEdge time frame energy.
        .INPUTS
        System.Management.Automation.PSCustomObject[] (SolarEdgeSiteTimeFrameEnergy)
        .LINK
        Get-SolarEdgeSiteEnergySummary
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnergySummary
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnergySummary in $SiteEnergySummary) {
            if (-not (PropertyExistsAndIsNotNull $_siteEnergySummary siteEnergySummary)) {
                throw "Invalid SiteEnergySummary object (property 'siteEnergySummary' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $energySummary = $_siteEnergySummary.siteEnergySummary

            Write-Output "Site ID               : $($_siteEnergySummary.siteId)"
            Write-Output "Start date            : $($_siteEnergySummary.startDate)"
            Write-Output "End date              : $($_siteEnergySummary.endDate)"
            Write-Output "Time frame energy     : $($energySummary.energy) $($energySummary.unit)"
            Write-Output "Start lifetime energy : $($energySummary.startLifetimeEnergy.energy) $($energySummary.startLifetimeEnergy.unit)"
            Write-Output "End lifetime energy   : $($energySummary.endLifetimeEnergy.energy) $($energySummary.endLifetimeEnergy.unit)"
        }
    }
}

function Write-SolarEdgeSiteEnvBenefits
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site environmental benefits to Output as text.
        .PARAMETER SiteEnvBenefits
        The SolarEdge site environmental benefits.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteEnvBenefits
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnvBenefits
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnvBenefits in $SiteEnvBenefits) {
            if (-not (PropertyExistsAndIsNotNull $_siteEnvBenefits siteEnvBenefits)) {
                throw "Invalid EnvironmentalBenefits object (property 'siteEnvBenefits' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $envBenefits = $_siteEnvBenefits.siteEnvBenefits
            $units       = $envBenefits.gasEmissionSaved.units.ToLower()

            Write-Output ('Site ID             : {0}'        -f $_siteEnvBenefits.siteId)
            Write-Output ('CO2 emission saved  : {0:F0} {1}' -f $envBenefits.gasEmissionSaved.co2, $units)
            Write-Output ('SO2 emission saved  : {0:F0} {1}' -f $envBenefits.gasEmissionSaved.so2, $units)
            Write-Output ('NOx emission saved  : {0:F0} {1}' -f $envBenefits.gasEmissionSaved.nox, $units)
            Write-Output ('Trees planted       : {0:F0}'     -f $envBenefits.treesPlanted)
            Write-Output ('Light bulbs powered : {0:F0}'     -f $envBenefits.lightBulbs)
        }
    }
}

function Write-SolarEdgeSiteInventory
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site inventory to Output as text.
        .PARAMETER SiteInventory
        The SolarEdge site inventory.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteInventory
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteInventory
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteInventory in $SiteInventory) {
            if (-not (PropertyExistsAndIsNotNull $_siteInventory siteInventory)) {
                throw "Invalid SiteInventory object (property 'siteInventory' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $inventory = $_siteInventory.siteInventory

            Write-Output "Site $($_siteInventory.siteId)"
            Write-Output "  Inverters        : $($inventory.inverters.count)"
            Write-Output "  Batteries        : $($inventory.batteries.count)"
            Write-Output "  Gateways         : $($inventory.gateways.count)"
            Write-Output "  Meters           : $($inventory.meters.count)"

            $nitems = 0
            foreach ($_inverter in $inventory.inverters) {
                ++$nitems

                Write-Output "Inverter #$nitems"
                Write-Output "  Name             : $($_inverter.name)"
                Write-Output "  Manufacturer     : $($_inverter.manufacturer)"
                Write-Output "  Model            : $($_inverter.model)"
                Write-Output "  Part number      : $($_inverter.partNumber)"
                Write-Output "  Serial number    : $($_inverter.SN)"
                Write-Output "  CPU version      : $($_inverter.cpuVersion)"
                Write-Output "  DSP1 version     : $($_inverter.dsp1Version)"
                Write-Output "  DSP2 version     : $($_inverter.dsp2Version)"
                Write-Output "  Communication    : $($_inverter.communicationMethod)"
                Write-Output "  Optimizers       : $($_inverter.connectedOptimizers)"
            }

            $nitems = 0
            foreach ($_battery in $inventory.batteries) {
                ++$nitems

                Write-Output "Battery #$nitems (properties TODO)"
            }

            $nitems = 0
            foreach ($_gateway in $inventory.gateways) {
                ++$nitems

                Write-Output "Gateway #$nitems (properties TODO)"
            }

            $nitems = 0
            foreach ($_meter in $inventory.meters) {
                ++$nitems

                Write-Output "Meter #$nitems"
                Write-Output "  Name             : $($_meter.name)"
                Write-Output "  Type             : $($_meter.type)"
                Write-Output "  Form             : $($_meter.form)"
                if ($_meter.form -eq "physical") {
                    Write-Output "  Model            : $($_meter.model)"
                    Write-Output "  Serial number    : $($_meter.SN)"
                    Write-Output "  Firmware version : $($_meter.firmwareVersion)"
                }
                Write-Output "  Connected to     : $($_meter.connectedTo) ($($_meter.connectedSolaredgeDeviceSN))"
            }
        }
    }
}

function Write-SolarEdgeSiteOverview
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site overview to Output as text.
        .PARAMETER SiteOverview
        The SolarEge site overview.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteOverview
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteOverview
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteOverview in $SiteOverview) {
            if (-not (PropertyExistsAndIsNotNull $_siteOverview siteOverview)) {
                throw "Invalid SiteOverview object (property 'siteOverview' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $overview = $_siteOverview.siteOverview

            Write-Output ('Site ID           : {0}'        -f  $_siteOverview.siteId)
            Write-Output ('Last update       : {0}'        -f  $overview.lastUpdatetime)
            Write-Output ('Lifetime energy   : {0:F1} kWh' -f ($overview.lifeTimeData.energy/1000))
            Write-Output ('Last year energy  : {0:F1} kWh' -f ($overview.lastYearData.energy/1000))
            Write-Output ('Last month energy : {0:F1} kWh' -f ($overview.lastMonthData.energy/1000))
            Write-Output ('Last day energy   : {0:F1} kWh' -f ($overview.lastDayData.energy/1000))
            Write-Output ('Current power     : {0:F0} W'   -f  $overview.currentPower.Power)
            Write-Output ('Measured by       : {0}'        -f  $overview.measuredBy)
        }
    }
}

function Write-SolarEdgeSitePower
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site power data to Output as text.
        .PARAMETER SitePower
        The SolarEdge site power data.
        .PARAMETER OmitHeader
        Omit table headers. The output will contain only timestamps and power readings.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSitePower
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SitePower,
        [Parameter()]                            [switch]           $OmitHeader
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_sitePower in $SitePower) {
            if (-not (PropertyExistsAndIsNotNull $SitePower sitePower)) {
                throw "Invalid SitePower object (property 'sitePower' does not exist or is null)"
            }

            $dateColumn           = [System.Data.DataColumn]::new('Date')
            $dateColumn.DataType  = [System.String]

            $valueColumn          = [System.Data.DataColumn]::new('Value')
            $valueColumn.DataType = [System.String]

            $powerTable       = [System.Data.DataTable]::new()
            $powerTable.Columns.Add($dateColumn)
            $powerTable.Columns.Add($valueColumn)

            if (-not $OmitHeader) {
                [void] $powerTable.Rows.Add('Date', 'Power')
                [void] $powerTable.Rows.Add('---',   '--')
            }

            foreach ($_value in $_sitePower.sitePower.values) {
                $value = if (PropertyExistsAndIsNotNull $_value value) { $_value.value.ToString('F1') } else { '0.0' }

                [void] $powerTable.Rows.Add($_value.date, $value)
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            if (-not $OmitHeader) {
                Write-Output "Site ID      $($_sitePower.siteId)"
                Write-Output "Start time   $($_sitePower.startTime)"
                Write-Output "End time     $($_sitePower.endTime)"
                Write-Output 'Time unit    QUARTER_OF_AN_HOUR'
                Write-Output "Energy unit  $($_sitePower.sitePower.unit)"
                Write-Output '---'
            }

            WriteTable $powerTable
        }
    }
}

function Write-SolarEdgeSitePowerDetails
{
    <#
        .SYNOPSIS
        Writes the SolarEdge site power details to Output as text.
        .PARAMETER SitePowerDetails
        The SolarEdge site power details.
        .PARAMETER OmitHeader
        Omit table headers. The output will contain only timestamps and power readings.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSitePowerDetails
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SitePowerDetails,
        [Parameter()]                            [switch]           $OmitHeader
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_sitePowerDetails in $SitePowerDetails) {
            if (-not (PropertyExistsAndIsNotNull $_sitePowerDetails sitePowerDetails)) {
                throw "Invalid PowerDetails object (property 'sitePowerDetails' does not exist or is null)"
            }

            $powerTable = $null

            foreach ($_meter in $_sitePowerDetails.sitePowerDetails.meters) {
                $meterHasData = $false
                foreach ($_value in $_meter.values) {
                    if (PropertyExistsAndIsNotNull $_value value) {
                        $meterHasData = $true
                        break
                    }
                }

                if ($meterHasData) {
                    $dateColumn               = [System.Data.DataColumn]::new('Date')
                    $dateColumn.DataType      = [System.String]

                    $valueColumn              = [System.Data.DataColumn]::new($_meter.type)
                    $valueColumn.DataType     = [System.String]
                    $valueColumn.DefaultValue = ''

                    $meterTable               = [System.Data.DataTable]::new()
                    $meterTable.Columns.Add($dateColumn)
                    $meterTable.Columns.Add($valueColumn)
                    $meterTable.PrimaryKey    = ($dateColumn)

                    if (-not $OmitHeader) {
                        [void] $meterTable.Rows.Add('Date', $_meter.type)
                        [void] $meterTable.Rows.Add('--',   '--')
                    }

                    foreach ($_value in $_meter.values) {
                        $value = if (PropertyExistsAndIsNotNull $_value value) { $_value.value.ToString('F1') } else { '0.0' }

                        [void] $meterTable.Rows.Add($_value.date, $value)
                    }

                    if ($null -eq $powerTable) {
                        $powerTable = $meterTable
                    } else {
                        $powerTable.Merge($meterTable)
                    }
                }
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            if (-Not $OmitHeader) {
                Write-Output "Site ID      $($_sitePowerDetails.siteId)"
                Write-Output "Start time   $($_sitePowerDetails.startTime)"
                Write-Output "End time     $($_sitePowerDetails.endTime)"
                Write-Output "Time unit    $($_sitePowerDetails.sitePowerDetails.timeUnit)"
                Write-Output "Energy unit  $($_sitePowerDetails.sitePowerDetails.unit)"
                Write-Output "---"
            }

            if ($null -eq $powerTable) {
                Write-Output "No power data available."
            } else {
                WriteTable $powerTable
            }
        }
    }
}

function Write-SolarEdgeSitePowerFlow
{
    <#
        .SYNOPSIS
        Writes the current SolarEdge site power flow to Output as text.
        .PARAMETER SitePowerFlow
        The SolarEdge site power flow.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSitePowerFlow
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SitePowerFlow
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_sitePowerFlow in $SitePowerFlow) {
            if (-not (PropertyExistsAndIsNotNull $_sitePowerFlow sitePowerFlow)) {
                throw "Invalid SitePowerFlow object (property 'sitePowerFlow' does not exist or is null)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $powerFlow = $_sitePowerFlow.sitePowerFlow

            Write-Output "Site ID  : $($_sitePowerFlow.siteId)"
            Write-Output "Refresh  : $($powerFlow.updateRefreshRate) seconds"
            Write-Output "---"

            Write-Output "Connections"
            foreach ($_connection in $powerFlow.connections) {
                Write-Output "  $($_connection.from) -> $($_connection.to)"
            }
            
            foreach ($_elementName in $powerFlowElements) {
                if (PropertyExistsAndIsNotNull $powerFlow $_elementName) {
                    $element = $powerFlow.$_elementName

                    Write-Output "Element $($_elementName)"
                    Write-Output "  Status : $($element.status)"

                    if (PropertyExistsAndIsNotNull $element currentPower) {
                        Write-Output "  Power  : $($element.currentPower) $($powerFlow.unit)"
                    }
                }
            }
        }
    }
}
