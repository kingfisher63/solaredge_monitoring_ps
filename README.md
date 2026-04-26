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
|:--------------------------- |:------------------------------------------------ |:------------------------------- |:-----:|
| Inverter Technical Data     | /equipment/&lt;site&gt;/&lt;serial&gt;/data      | Get-SolarEdgeInverterData       |       |
| Inventory                   | /site/&lt;site&gt;/inventory                     | Get-SolarEdgeSiteInventory      |       |
| Get Meter Data              | /site/&lt;site&gt;/meters                        | Get-SolarEdgeMeterData          |       |
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
|:-------------------- |:--------------------------------- |:----------------------------------------------- |
| Component List       | /equipment/&lt;site&gt;/list      | Inventory API returns more detailed information |

### SolarEdge.Monitoring.Util

This nested module contains functions to produce human readable text from selected SolarEdge.Monitoring.Api function output. Pipelining is supported.

#### Functions

The table below lists the *SolarEdge.Monitoring.Api* functions and their *SolarEdge.Monitoring.Util* counterparts (where applicable).

| SolarEdge.Monitoring.Api function | SolarEdge.Monitoring.Util function | Notes |
|:--------------------------------- |:---------------------------------- |:-----:|
| Get-SolarEdgeInverterData         | *not implemented*                  | 1     |
| Get-SolarEdgeMeterData            | Write-SolarEdgeMeterData           |       |
| Get-SolarEdgeSiteDataPeriod       | Write-SolarEdgeSiteDataPeriod      |       |
| Get-SolarEdgeSiteDetails          | Write-SolarEdgeSiteDetails         |       |
| Get-SolarEdgeSiteEnergy           | Write-SolarEdgeSiteEnergy          |       |
| Get-SolarEdgeSiteEnergyDetails    | Write-SolarEdgeSiteEnergyDetails   |       |
| Get-SolarEdgeSiteEnergySummary    | Write-SolarEdgeSiteEnergySummary   |       |
| Get-SolarEdgeSiteEnvBenefits      | Write-SolarEdgeSiteEnvBenefits     |       |
| Get-SolarEdgeSiteInventory        | Write-SolarEdgeSiteInventory       | 2     |
| Get-SolarEdgeSiteList             | Write-SolarEdgeSiteDetails         |       |
| Get-SolarEdgeSiteOverview         | Write-SolarEdgeSiteOverview        |       |
| Get-SolarEdgeSitePower            | Write-SolarEdgeSitePower           |       |
| Get-SolarEdgeSitePowerDetails     | Write-SolarEdgeSitePowerDetails    |       |
| Get-SolarEdgeSitePowerFlow        | Write-SolarEdgeSitePowerFlow       |       |
| Get-SolarEdgeStorageData          | *not implemented*                  | 3     |

(1) Data is too complex for simple text representation. Use *Export-SolarEdgeInverterData* to export data to a CSV file.
(2) Batteries and gateways are not decoded (no sample data available).
(3) No sample data available. Possible future addition.

### Module SolarEdge.Monitoring.Export

This nested module contains functions to export SolarEdge.Monitoring.Api output to a CSV file.

| SolarEdge.Monitoring.Api function | SolarEdge.Monitoring.Util function |
|:--------------------------------- |:---------------------------------- |
| Get-SolarEdgeInverterData         | Export-SolarEdgeInverterData       |
| Get-SolarEdgeMeterData            | *not implemented*                  |
| Get-SolarEdgeSiteDataPeriod       | *not implemented*                  |
| Get-SolarEdgeSiteDetails          | *not implemented*                  |
| Get-SolarEdgeSiteEnergy           | Export-SolarEdgeSiteEnergy         |
| Get-SolarEdgeSiteEnergyDetails    | *not implemented*                  |
| Get-SolarEdgeSiteEnergySummary    | *not implemented*                  |
| Get-SolarEdgeSiteEnvBenefits      | *not implemented*                  |
| Get-SolarEdgeSiteInventory        | *not implemented*                  |
| Get-SolarEdgeSiteList             | *not implemented*                  |
| Get-SolarEdgeSiteOverview         | *not implemented*                  |
| Get-SolarEdgeSitePower            | *not implemented*                  |
| Get-SolarEdgeSitePowerDetails     | *not implemented*                  |
| Get-SolarEdgeSitePowerFlow        | *not implemented*                  |
| Get-SolarEdgeStorageData          | *not implemented*                  |

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
