[cmdletbinding()]
Param (
    [Parameter(
        Position = 0,
        Mandatory,
        HelpMessage = "Enter a UNC path like \\server\share"
    )]
    [ValidatePattern('^\\\\[a-zA-Z-\d]+(\\[a-zA-Z-\d\$]+)+$')]
    [ValidateScript({Test-Path -Path $_ })]
    [string]$Path
)

Write-Host "Getting top level folder size for $Path" -ForegroundColor Yellow

$measure = Get-ChildItem $path | Measure-Object -Property Length -Sum
[PSCustomObject]@{
    Path      = $Path
    FileCount = $measure.Count
    Size      = $measure.Sum
}

