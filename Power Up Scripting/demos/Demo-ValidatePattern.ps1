[cmdletbinding()]
Param (
    [Parameter(Position = 0, Mandatory, HelpMessage = "Enter a UNC path like \\server\share")]
    [ValidatePattern('^\\\\[a-zA-Z-\d]+(\\[a-zA-Z-\d\$]+)+$')]
    [ValidateScript({ Test-Path -Path $_ })]
    [string]$Path
)

Write-Host "Getting top level folder size for $Path" -ForegroundColor Yellow
Get-ChildItem $path | Measure-Object -Property Length -Sum

