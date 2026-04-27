# Copyright (C) 2025-2026 Roger Hunen
#
# Permission to use, copy, modify, and distribute this software and its
# documentation under the terms of the GNU General Public License is hereby 
# granted. No representations are made about the suitability of this software 
# for any purpose. It is provided "as is" without express or implied warranty.
# See the GNU General Public License for more details.

@{
    GUID              = '7781f6a2-721e-4de1-88f2-f1c4b4a9e31c'
    ModuleVersion     = '2.0.0'
    Author            = 'Roger Hunen'
    CompanyName       = 'Roger Hunen'
    Copyright         = '(c) 2025-2026 Roger Hunen. All rights reserved.'
	Description       = 'Functions to retrieve data from the SolarEdge Monitoring API'
    NestedModules     = @(
        'SolarEdge.Monitoring.Api.psm1',
        'SolarEdge.Monitoring.Export.psm1',
        'SolarEdge.Monitoring.Util.psm1'
    )
    FunctionsToExport = @(
        'Export-SolarEdgeInverterData'
        'Export-SolarEdgeSiteEnergy'
        'Get-SolarEdgeInverterData'
        'Get-SolarEdgeMeterData'
        'Get-SolarEdgeSiteDataPeriod'
        'Get-SolarEdgeSiteDetails'
        'Get-SolarEdgeSiteEnergy'
        'Get-SolarEdgeSiteEnergyDetails'
        'Get-SolarEdgeSiteEnergySummary'
        'Get-SolarEdgeSiteEnvBenefits'
        'Get-SolarEdgeSiteInventory'
        'Get-SolarEdgeSiteList'
        'Get-SolarEdgeSiteOverview'
        'Get-SolarEdgeSitePower'
        'Get-SolarEdgeSitePowerDetails'
        'Get-SolarEdgeSitePowerFlow'
        'Get-SolarEdgeStorageData'
        'Write-SolarEdgeSiteDataPeriod'
        'Write-SolarEdgeSiteDetails'
        'Write-SolarEdgeSiteEnergy'
        'Write-SolarEdgeSiteEnergyDetails'
        'Write-SolarEdgeSiteEnergySummary'
        'Write-SolarEdgeSiteEnvBenefits'
        'Write-SolarEdgeSiteInventory'
        'Write-SolarEdgeSiteOverview'
        'Write-SolarEdgeSitePower'
        'Write-SolarEdgeSitePowerDetails'
        'Write-SolarEdgeSitePowerFlow'
  )
}
