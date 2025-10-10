# Copyright (C) 2025 Roger Hunen
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

function GetValueFieldWidth
{
    param (
        [Object[]] $objects
    )

    $width = 0;

    foreach ($_object in $objects) {
        if (PropertyHasValue $_object value) {
            $length = $_object.value.ToString('F1').Length
        } else {
            $length = 3
        }

        if ($width -lt $length) {
            $width = $length
        }
    }

    return $width
}

function PropertyExists
{
    param (
        [Object] $object,
        [String] $propertyName
    )

    return ($object | Get-Member -Name $propertyName)
}

function PropertyHasValue
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

#
# Exported functions
#

function Write-SolarEdgeApiInfo
{
    <#
        .SYNOPSIS
        Writes the current and supported SolarEdge Monitoring API versions to
        Output as text.
        .PARAMETER ApiInfo
        The SolarEdge API info.
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeApiInfo
    #>

    param (
        [Parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $ApiInfo
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_apiInfo in $ApiInfo) {
            if (-not (PropertyExists $_apiInfo apiInfo)) {
                throw "Invalid ApiInfo object (property 'apiInfo' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $info              = $_apiInfo.apiInfo
            $currentVersion    = $info.version.release
            $supportedVersions = @()

            foreach ($_supported in $info.supported) {
                $supportedVersions += $_supported.release
            }

            Write-Output "Current API version : $currentVersion"
            Write-Output "Supported versions  : $($supportedVersions -join ', ')"
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
            if (-not (PropertyExists $_siteDataPeriod siteDataPeriod)) {
                throw "Invalid SiteDataPeriod object (property 'siteDataPeriod' does not exist)"
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
            if (-not (PropertyExists $_siteDetails siteDetails)) {
                throw "Invalid SiteDetails object (property 'siteDetails' does not exist)"
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
            if (PropertyExists $details status) {
                Write-Output "Status         : $($details.status)"
            }
            if (PropertyExists $details alertQuantity) {
                Write-Output "Alert quantity : $($details.alertQuantity)"
            }
            if (PropertyExists $details highestImpact) {
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
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteEnergy
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnergy
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnergy in $SiteEnergy) {
            if (-not (PropertyExists $_siteEnergy siteEnergy)) {
                throw "Invalid SiteEnergy object (property 'siteEnergy' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            Write-Output "Site ID    : $($_siteEnergy.siteId)"
            Write-Output "Start date : $($_siteEnergy.startDate)"
            Write-Output "End date   : $($_siteEnergy.endDate)"
            Write-Output "Time unit  : $($_siteEnergy.timeUnit)"
            Write-Output "---"

            $unit       = $_siteEnergy.siteEnergy.unit
            $values     = $_siteEnergy.siteEnergy.values
            $valueWidth = GetValueFieldWidth $values

            foreach ($_value in $values) {
                if (PropertyHasValue $_value value) {
                    Write-Output ("$($_value.date)  {0,$valueWidth} $unit" -f $_value.value.ToString('F1'))
                } else {
                    Write-Output ("$($_value.date)  {0,$valueWidth} $unit" -f '0.0')
                }
            }
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
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSiteEnergyDetails
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SiteEnergyDetails
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_siteEnergyDetails in $SiteEnergyDetails) {
            if (-not (PropertyExists $_siteEnergyDetails siteEnergyDetails)) {
                throw "Invalid SiteEnergyDetails object (property 'siteEnergyDetails' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $energyDetails = $_siteEnergyDetails.siteEnergyDetails

            Write-Output "Site ID      : $($_siteEnergyDetails.siteId)"
            Write-Output "Start time   : $($_siteEnergyDetails.startTime)"
            Write-Output "End time     : $($_siteEnergyDetails.endTime)"
            Write-Output "Time unit    : $($energyDetails.timeUnit)"
            Write-Output "---"

            foreach ($_meter in $energyDetails.meters) {
                $meterHasData = $false
                foreach ($_value in $_meter.values) {
                    if (PropertyHasValue $_value value) {
                        $meterHasData = $true
                        break
                    }
                }

                if (-not $meterHasData) {
                    Write-Output "Meter '$($_meter.type)'"
                    Write-Output "  No data"
                    continue
                }

                $valueWidth = GetValueFieldWidth $_meter.values

                Write-Output "Meter '$($_meter.type)'"
                foreach ($_value in $_meter.values) {
                    $formatString = "  $($_value.date)  {0,$valueWidth} $($energyDetails.unit)"
                    if (PropertyHasValue $_value value) {
                        Write-Output ($formatString -f $_value.value.ToString('F1'))
                    } else {
                        Write-Output ($formatString -f '0.0')
                    }
                }
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
            if (-not (PropertyExists $_siteEnergySummary siteEnergySummary)) {
                throw "Invalid SiteEnergySummary object (property 'siteEnergySummary' does not exist)"
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
        Writes the SolarEdge site current powerflow to Output as text.
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
            if (-not (PropertyExists $_siteEnvBenefits siteEnvBenefits)) {
                throw "Invalid EnvironmentalBenefits object (property 'siteEnvBenefits' does not exist)"
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
            if (-not (PropertyExists $_siteInventory siteInventory)) {
                throw "Invalid SiteInventory object (property 'siteInventory' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $inventory = $_siteInventory.siteInventory

            Write-Output "Summary"
            Write-Output "  Site ID       : $($_siteInventory.siteId)"
            Write-Output "  Inverters     : $($inventory.inverters.count)"
            Write-Output "  Batteries     : $($inventory.batteries.count)"
            Write-Output "  Gateways      : $($inventory.gateways.count)"
            Write-Output "  Meters        : $($inventory.meters.count)"
            Write-Output "  Sensors       : $($inventory.sensors.count)"

            $nitems = 0
            foreach ($_inverter in $inventory.inverters) {
                ++$nitems

                Write-Output "Inverter #$nitems"
                Write-Output "  Name          : $($_inverter.name)"
                Write-Output "  Manufacturer  : $($_inverter.manufacturer)"
                Write-Output "  Model         : $($_inverter.model)"
                Write-Output "  Part number   : $($_inverter.partNumber)"
                Write-Output "  Serial number : $($_inverter.SN)"
                Write-Output "  CPU version   : $($_inverter.cpuVersion)"
                Write-Output "  DSP1 version  : $($_inverter.dsp1Version)"
                Write-Output "  DSP2 version  : $($_inverter.dsp2Version)"
                Write-Output "  Communication : $($_inverter.communicationMethod)"
                Write-Output "  Optimizers    : $($_inverter.connectedOptimizers)"
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

                Write-Output "Meter #$nitems (properties TODO)"
            }

            $nitems = 0
            foreach ($_sensor in $inventory.sensors) {
                ++$nitems

                Write-Output "Sensor #$nitems (properties TODO)"
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
            if (-not (PropertyExists $_siteOverview siteOverview)) {
                throw "Invalid SiteOverview object (property 'siteOverview' does not exist)"
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
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSitePower
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SitePower
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_sitePower in $SitePower) {
            if (-not (PropertyExists $SitePower sitePower)) {
                throw "Invalid SitePower object (property 'sitePower' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            Write-Output "Site ID         : $($_sitePower.siteId)"
            Write-Output "Start date/time : $($_sitePower.startTime)"
            Write-Output "End date/time   : $($_sitePower.endTime)"
            Write-Output "Time unit       : $($_sitePower.sitePower.timeUnit)"
            Write-Output "---"

            $unit       = $_sitePower.sitePower.unit
            $values     = $_sitePower.sitePower.values
            $valueWidth = GetValueFieldWidth $values

            foreach ($_value in $values) {
                if (PropertyHasValue $_value value) {
                    Write-Output ("$($_value.date)  {0,$valueWidth} $unit" -f $_value.value.ToString('F1'))
                } else {
                    Write-Output ("$($_value.date)  {0,$valueWidth} $unit" -f '0.0')
                }
            }
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
        .INPUTS
        System.Management.Automation.PSCustomObject[]
        .LINK
        Get-SolarEdgeSitePowerDetails
    #>

    param (
        [parameter(Mandatory,ValueFromPipeline)] [PSCustomObject[]] $SitePowerDetails
    )

    begin {
        $n = 0
    }

    process {
        foreach ($_sitePowerDetails in $SitePowerDetails) {
            if (-not (PropertyExists $_sitePowerDetails sitePowerDetails)) {
                throw "Invalid PowerDetails object (property 'sitePowerDetails' does not exist)"
            }

            if (++$n -gt 1) {
                Write-Output ''
            }

            $powerDetails = $_sitePowerDetails.sitePowerDetails

            Write-Output "Site ID      : $($_sitePowerDetails.siteId)"
            Write-Output "Start time   : $($_sitePowerDetails.startTime)"
            Write-Output "End time     : $($_sitePowerDetails.endTime)"
            Write-Output "Time unit    : $($powerDetails.timeUnit)"
            Write-Output "---"

            foreach ($_meter in $powerDetails.meters) {
                $meterHasData = $false
                foreach ($_value in $_meter.values) {
                    if (PropertyHasValue $_value value) {
                        $meterHasData = $true
                        break
                    }
                }

                if (-not $meterHasData) {
                    Write-Output "Meter '$($_meter.type)'"
                    Write-Output "  No data"
                    continue
                }

                $valueWidth = GetValueFieldWidth $_meter.values

                Write-Output "Meter '$($_meter.type)'"
                foreach ($_value in $_meter.values) {
                    $formatString = "  $($_value.date)  {0,$valueWidth} $($powerDetails.unit)"
                    if (PropertyHasValue $_value value) {
                        Write-Output ($formatString -f $_value.value.ToString('F1'))
                    } else {
                        Write-Output ($formatString -f '0.0')
                    }
                }
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
            if (-not (PropertyExists $_sitePowerFlow sitePowerFlow)) {
                throw "Invalid SitePowerFlow object (property 'sitePowerFlow' does not exist)"
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
                if (PropertyExists $powerFlow $_elementName) {
                    $element = $powerFlow.$_elementName

                    Write-Output "Element $($_elementName)"
                    Write-Output "  Status : $($element.status)"

                    if (PropertyExists $element currentPower) {
                        Write-Output "  Power  : $($element.currentPower) $($powerFlow.unit)"
                    }
                }
            }
        }
    }
}
