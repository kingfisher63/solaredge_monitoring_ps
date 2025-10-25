# SolarEdge Monitoring API for Windows PowerShell

This repository contains Windows PowerShell modules with functions to query the [SolarEdge Monitoring API](https://knowledge-center.solaredge.com/sites/kc/files/se_monitoring_api.pdf). The functions are grouped into two PowerShell modules

+ SolarEdge.Monitoring
+ SolarEdge.Monitoring.Export

## Module SolarEdge.Monitoring

SolarEdge.Monitoring has two nested modules

+ **SolarEdge.Monitoring.Api** contains functions to query the SolarEdge Monitoring API.
+ **SolarEdge.Monitoring.Util** contains functions to produce human readable text from data retrieved with a *SolarEdge.Monitoring.Api* function.

### Installation

*SolarEdge.Monitoring* has a module manifest and is therefore best installed using the NuGet package (*.nupkg* file).

+ Create a PowerShell repository directory and register it with *Register-PSRepository*.
+ Copy the *.nupkg* file into the repository directory.
+ Install the module into PowerShell with *Install-Module*.

### SolarEdge.Monitoring.Api

This nested module contains functions to query the SolarEdge Monitoring API. Each function returns a *PSCustomObject* object (or an array of objects if multiple sites are queried). The result object has a string property named *siteId* (except for *Get-SolarEdgeApiInfo*) as well as an object property that matches the function name. The object property replaces the top level property returned by the SolarEdge Monitoring API. Query parameters (e.g. *startDate*) are also added as properties to facilitate further processing.

The function *Get-SolarEdgeSiteDataPeriod* will for example return (JSON)
```
{
    "siteId": 1,
    "siteDataPeriod": {
        "startDate": "2019-10-07",
        "endDate": "2025-10-09"
    }
}
```

*SolarEdge.Monitoring.Api* functions use bulk version API endpoints where available to minimize the number of API requests (the Monitoring API limits requests to 300 per day).

#### APIs Supported

| API                         | API Endpoint                                     | Function                        | Notes |
| :-------------------------- | :----------------------------------------------- | :------------------------------ | :---: |
| API Version                 | /version/current                                 | Get-SolaredgeApiInfo            | 1     |
| API Versions Supported      | /version/supported                               | Get-SolaredgeApiInfo            | 1     |
| Equipment Change Log        | /equipment/&lt;site&gt;/&lt;serial&gt;/changeLog | Get-SolarEdgeEquipmentChangeLog |       |
| Inverter Technical Data     | /equipment/&lt;site&gt;/&lt;serial&gt;/data      | Get-SolarEdgeInverterData       |       |
| Inventory                   | /site/&lt;site&gt;/inventory                     | Get-SolarEdgeSiteInventory      |       |
| Get Meter Data              | /site/&lt;site&gt;/meters                        | Get-SolarEdgeMeterData          |       |
| Get Sensor Data             | /equipment/&lt;site&gt;/sensors                  | Get-SolarEdgeSensorData         |       |
| Get Sensor List             | /equipment/&lt;site&gt;/sensors                  | Get-SolarEdgeSensorList         |       |
| Site Data                   | /sites/&lt;site_list&gt;/dataPeriod              | Get-SolarEdgeSiteDataPeriod     |       |
| Site Details                | /site/&lt;site&gt;/details                       | Get-SolarEdgeSiteDetails        |       |
| Site Energy                 | /sites/&lt;site_list&gt;/energy                  | Get-SolarEdgeSiteEnergy         |       |
| Site Energy (details)       | /site/&lt;site&gt;/energyDetails                 | Get-SolarEdgeSiteEnergyDetails  |       |
| Site Energy (time frame)    | /sites/&lt;site_list&gt;/timeframeEnergy         | Get-SolarEdgeSiteEnergySummary  |       |
| Site Environmental Benefits | /site/&lt;site&gt;/envBenefits                   | Get-SolarEdgeSiteEnvBenefits    |       |
| Site List                   | /sites/list                                      | Get-SolarEdgeSiteList           | 2,3   |
| Site Overview               | /sites/&lt;site_list&gt;/overview                | Get-SolarEdgeSiteOverview       |       |
| Site Power                  | /sites/&lt;site_list&gt;/power                   | Get-SolarEdgeSitePower          |       |
| Site Power (details)        | /site/&lt;site&gt;/powerDetails                  | Get-SolarEdgeSitePowerDetails   |       |
| Site Power Flow             | /site/&lt;site&gt;/currentPowerFlow              | Get-SolarEdgeSitePowerFlow      |       |
| Storage Information         | /site/&lt;site&gt;/storageData                   | Get-SolarEdgeStorageData        |       |

(1) Result object does not have a *siteId* property.  
(2) This function returns an array of *siteDetails* style objects.  
(3) Selection parameters are not supported as these seem to be not functional.

#### APIs not supported

| API                  | API Endpoint                      | Reason                                          |
| ---------------------| --------------------------------- | ----------------------------------------------- |
| Account List         | /accounts/list                    | API returns HTTP 403 (forbidden)                |
| Component List       | /equipment/&lt;site&gt;/list      | Inventory API returns more detailed information |
| Installer Logo Image | /site/&lt;site&gt;/installerImage | Possible future addition                        |
| Site Image           | /site/&lt;site&gt;/siteImage      | Possible future addition                        |

### SolarEdge.Monitoring.Util

This nested module contains functions to produce human readable text from selected SolarEdge.Monitoring.Api function output. Pipelining is supported.

#### Functions

The table below lists the *SolarEdge.Monitoring.Api* functions and their *SolarEdge.Monitoring.Util* counterparts (where applicable).

| SolarEdge.Monitoring.Api function | SolarEdge.Monitoring.Util function | Notes |
| --------------------------------- | ---------------------------------- | ----- |
| Get-SolarEdgeApiInfo              | Write-SolarEdgeApiInfo             |       |
| Get-SolarEdgeEquipmentChangeLog   | *not implemented*                  | 1     |
| Get-SolarEdgeInverterData         | *not implemented*                  | 2     |
| Get-SolarEdgeMeterData            | *not implemented*                  | 1     |
| Get-SolarEdgeSensorData           | *not implemented*                  | 1     |
| Get-SolarEdgeSensorList           | *not implemented*                  | 1     |
| Get-SolarEdgeSiteDataPeriod       | Write-SolarEdgeSiteDataPeriod      |       |
| Get-SolarEdgeSiteDetails          | Write-SolarEdgeSiteDetails         |       |
| Get-SolarEdgeSiteEnergy           | Write-SolarEdgeSiteEnergy          |       |
| Get-SolarEdgeSiteEnergyDetails    | Write-SolarEdgeSiteEnergyDetails   |       |
| Get-SolarEdgeSiteEnergySummary    | Write-SolarEdgeSiteEnergySummary   |       |
| Get-SolarEdgeSiteEnvBenefits      | Write-SolarEdgeSiteEnvBenefits     |       |
| Get-SolarEdgeSiteInventory        | Write-SolarEdgeSiteInventory       | 3     |
| Get-SolarEdgeSiteList             | Write-SolarEdgeSiteDetails         |       |
| Get-SolarEdgeSiteOverview         | Write-SolarEdgeSiteOverview        |       |
| Get-SolarEdgeSitePower            | Write-SolarEdgeSitePower           |       |
| Get-SolarEdgeSitePowerDetails     | Write-SolarEdgeSitePowerDetails    |       |
| Get-SolarEdgeSitePowerFlow        | Write-SolarEdgeSitePowerFlow       |       |
| Get-SolarEdgeStorageData          | *not implemented*                  | 1     |

(1) No sample data available. Possible future addition.  
(2) Data is too complex for simple text representation. Use *Export-SolarEdgeInverterData* to export data to a CSV file.  
(3) Only inverter data is fully decoded as no sample data was available for batteries, gateways, meters en sensors.

## Module SolarEdge.Monitoring.Export

This module contains functions to export site energy data and inverter technical data to a CSV file.

### Installation

This module was developed to serve my personal needs and is provided as-is because I think it may be useful to others. Therefore this module does not have a module manifest and must be installed manually.

+ Create a directory named *SolarEdge.Monitoring.Export* in the module directory for the *CurrentUser* scope (see [about_PSModulePath](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_psmodulepath)).
+ Copy the script file *SolarEdge.Monitoring.Export.psm1* into this directory.
+ Restart PowerShell

Output file names are generated automatically, but can be easily customized (modify the file name format variables near the top of the module script).

## Examples

Setup API key and site ID.
```
PS> $key  = 'L4QLVQ1LOKCQX2193VSEICXW61NP6B1O'
PS> $site = 1
```

Use a temporary object.
```
PS> $temp = GetSolarEdgeSiteDataPeriod $key $site
PS> $temp | ConvertTo-Json
{
    "siteId":  1,
    "siteDataPeriod":  {
                           "startDate":  "2019-10-07",
                           "endDate":  "2025-10-10"
                       }
}

PS> Write-SolarEdgeSiteDataPeriod $temp
Site ID    : 1
Start date : 2019-10-07
End date   : 2025-10-10
```

Use pipelining.
```
PS> Get-SolarEdgeSiteDataPeriod $key $site | Write-SolarEdgeSiteDataPeriod
Site ID    : 1
Start date : 2019-10-07
End date   : 2025-10-10

PS> $site | Get-SolarEdgeSiteDataPeriod $key | Write-SolarEdgeSiteDataPeriod
Site ID    : 1
Start date : 2019-10-07
End date   : 2025-10-10

PS> Get-SolarEdgeSiteList $key | Select-Object -ExpandProperty siteId | Get-SolarEdgeSiteDataPeriod $key | Write-SolarEdgeSiteDataPeriod
Site ID    : 1
Start date : 2019-10-07
End date   : 2025-10-10
```

Export monthly site energy for a year.
```
PS> Export-SolarEdgeSiteEnergy $key $site -Year 2024 -TimeUnit MONTH -Verbose
VERBOSE: Time span   : YEAR
VERBOSE: Start date  : 2024-01-01
VERBOSE: End date    : 2025-01-01
VERBOSE: Time unit   : MONTH
VERBOSE: Site ID     : 1
VERBOSE: Output file : 2024 - Example_Site (1) - MONTH.csv

PS> Get-Content '.\2024 - Example_Site (1) - MONTH.csv'
"date","value"
"2024-01-01 00:00:00","137361.0"
"2024-02-01 00:00:00","152399.0"
"2024-03-01 00:00:00","372607.0"
"2024-04-01 00:00:00","437024.0"
"2024-05-01 00:00:00","524664.0"
"2024-06-01 00:00:00","587187.0"
"2024-07-01 00:00:00","608695.0"
"2024-08-01 00:00:00","642817.0"
"2024-09-01 00:00:00","434086.0"
"2024-10-01 00:00:00","304928.0"
"2024-11-01 00:00:00","143920.0"
"2024-12-01 00:00:00","74057.0"
```
