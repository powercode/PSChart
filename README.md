# PSChart

Very basic charting for PowerShell

## Usage

```powershell
Import-Module PSChart
Get-PerforceChangeList ... -TotalCount 1000 | 
    group {  $_.time.ToShortDateString()} | 
    Out-Chart -Title 'Perforce Commits' -XLabel 'Commit Date'
    
Get-PerforceChangeList ... -TotalCount 1000 | 
    group {  $_.time.ToShortDateString()} | 
    Out-Chart -Title 'Perforce Commits' -XLabel 'Commit Date' -OutPath chart.png
```

![sample chart](https://github.com/PowerCode/PSChart/raw/master/img/chart.png "Sample chart")
