#requires -version 5.1
#requires -module SQLPS

<#
Create a SQL Server database report
This script includes a graphic file which by default should be in the
same directory as this script. Otherwise, edit the $ImageFile variable.

It is recommended that you import the SQLPS module first and then run
this script.

Usage:
c:\scripts\SQLServerReport -computername CHI-SQL01 -includesystem -path c:\work\CHI-SQL01-DB.htm


#>

[cmdletbinding()]

Param(
    [Parameter(Position = 0, HelpMessage = "Enter the name of a SQL server")]
    [ValidateNotNullorEmpty()]
    [Alias("CN")]
    [string]$computername = $env:computername,
    [Parameter(Position = 1, HelpMessage = "Enter the named instance such as Default")]
    [ValidateNotNullorEmpty()]
    [string]$Instance = "SQLExpress",
    [switch]$IncludeSystem,
    [ValidateNotNullorEmpty()]
    [string]$Path = "$env:temp\sqlrpt.htm",
    [switch]$ShowReport
)

$scriptversion = "2.4"

Write-Verbose "Starting $($MyInvocation.Mycommand)"

#define the path to the graphic - this is hard coded now
$graphic = "db.png"
#the default location is the same directory as this script
$imagefile = Join-Path -Path (Split-Path $MyInvocation.InvocationName) -ChildPath $graphic

#define an empty array to hold all of the HTML fragments
$fragments = @("<br><br><br>")

if ($instance -eq 'Default') {
    $serverinstance = $Computername
}
else {
    $serverinstance = "$computername\$instance"
}
Write-Verbose "Querying $serverinstance"

$invokeParams = @{
    Query          = $null
    Database       = 'master'
    ServerInstance = $serverInstance
    ErrorAction    = 'stop'
}

#get uptime
Write-Verbose "Getting SQL Server uptime"

Try {
    #try to connect to the SQL server
    $invokeParams.Query = 'SELECT sqlserver_start_time AS StartTime FROM sys.dm_os_sys_info'
    $starttime = Invoke-Sqlcmd @invokeParams
}
Catch {
    Write-Warning "Can't connect to $computername. $($_.exception.message)"
    #bail out
    Return
}

Write-Verbose "Getting SQL Version"
$invokeParams.query = "Select @@version AS Version,@@ServerName AS Name"
$version = Invoke-Sqlcmd @invokeParams

#create an object
$uptime = New-Object -TypeName PSObject -Property @{
    StartTime = $starttime.Item(0)
    Uptime    = (Get-Date) - $starttime.Item(0)
    Version   = $version.Item(0).replace("`n", "|")
}

$tmp = $uptime | ConvertTo-Html -Fragment -As List
#replace "|" place holder with <br>"
$fragments += $tmp.replace("|", "<br>")

#get services
Write-Verbose "Querying SQL services"
$services = Get-Service -DisplayName *SQL* -ComputerName $computername |
Select-Object -Property Name, Displayname, Status

#add conditional formatting to display stopped services in yellow
[xml]$html = $services | ConvertTo-Html -Fragment

#check each row, skipping the TH header row
for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
    $class = $html.CreateAttribute("class")
    #check the value of the last column and assign a class to the row
    if ($html.table.tr[$i].td[-1] -ne 'Running') {
        $html.table.tr[$i].lastChild.setAttribute("class", "warn") | Out-Null
    }
} #for

$fragments += "<h3>SQL Services</h3>"
$fragments += $html.InnerXML

#get database information
#path to databases
$dbpath = "SQLServer:\SQL\$computername\$instance\databases"
Write-Verbose "Querying database information from $dbpath"

if ($IncludeSystem) {
    Write-Verbose "Including system databases"
    $dbs = Get-ChildItem -Path $dbpath -Force
}
else {
    $dbs = Get-ChildItem -Path $dbpath
}

[xml]$html = $dbs | Select-Object -Property Name,
@{Name = "SizeMB"; Expression = { $_.size } },
@{Name = "DataSpaceMB"; Expression = { $_.DataSpaceUsage / 1KB } },
@{Name = "AvailableMB"; Expression = { $_.SpaceAvailable / 1KB } },
@{Name = "PercentFree"; Expression = { [math]::Round((($_.SpaceAvailable / 1kb) / $_.size) * 100, 2) } } |
Sort-Object -Property PercentFree | ConvertTo-Html -Fragment

for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
    $class = $html.CreateAttribute("class")
    #check the value of the last column and assign a class to the row
    if (($html.table.tr[$i].td[-1] -as [double]) -le 15) {
        $html.table.tr[$i].lastChild.SetAttribute("class", "danger") | Out-Null
    }
    elseif (($html.table.tr[$i].td[-1] -as [double]) -le 25) {
        $class.value = "warn"
        $html.table.tr[$i].lastChild.SetAttribute("class", "warn") | Out-Null
    }
}
$fragments += "<h3>Database Utilization</h3>"
$fragments += $html.InnerXml

$fragments += "<h3>Database Backup</h3>"
$fragments += $dbs | Select-Object -Property Name, Owner, CreateDate, Last*, RecoveryModel | ConvertTo-Html -Fragment

#logins
Write-Verbose "Querying logins"
$dbpath = "SQLServer:\SQL\$computername\$instance\Logins"
$fragments += "<h3>SQL Server Logins</h3>"
$fragments += Get-ChildItem -Path $dbpath | Sort-Object -Property Name |
Select-Object -Property Name, IsDisabled, DateLastModified, LoginType, DefaultDatabase |
ConvertTo-Html -Fragment -As Table

#volume usage
Write-Verbose "Querying system volumes"
$data = Get-CimInstance win32_volume -Filter "drivetype=3" -ComputerName $computername

$drives = foreach ($item in $data) {
    $prophash = [ordered]@{
        Drive       = $item.DriveLetter
        Volume      = $item.DeviceID
        Compressed  = $item.Compressed
        SizeGB      = $item.capacity / 1GB -as [int]
        FreeGB      = "{0:N4}" -f ($item.Freespace / 1GB )
        PercentFree = [math]::Round((($item.Freespace / $item.capacity) * 100), 2)
    }

    #create a new object from the property hash
    New-Object PSObject -Property $prophash
}

[xml]$html = $drives | ConvertTo-Html -Fragment

#check each row, skipping the TH header row
for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
    $class = $html.CreateAttribute("class")
    #check the value of the last column and assign a class to the row
    if (($html.table.tr[$i].td[-1] -as [int]) -le 25) {
        $html.table.tr[$i].lastChild.SetAttribute("class", "danger") | Out-Null
    }
    elseif (($html.table.tr[$i].td[-1] -as [int]) -le 35) {
        $class.value = "warn"
        $html.table.tr[$i].lastChild.SetAttribute("class", "warn") | Out-Null
    }
}

$fragments += "<h3>Volume Utilization</h3>"
$fragments += $html.innerxml

#define the HTML style
Write-Verbose "preparing report"

#encode the graphic file to embed into the HTML
Write-Verbose "Encoding graphic $imagefile"
$ImageBits = [Convert]::ToBase64String((Get-Content $imagefile -Encoding Byte))
$ImageHTML = "<img src=data:image/png;base64,$($ImageBits) alt=utilization />"

#define a here string for the html header
$head = @"
<style>
body { background-color:#FAFAFA;
       font-family:Arial;
       font-size:12pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
tr:nth-child(odd) {background-color: lightgray}
table { margin-left:50px; }
img
{
float:left;
margin: 0px 25px;
}
.danger {background-color: red}
.warn {background-color: yellow}
</style>
$imagehtml
<br><br><br>
<H2>SQL Server Report: $($version.name)</H2>
<br>
"@

#HTML to display at the end of the report
$footer = @"
<br>
<i>
Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: $(Get-Date)<br>
Author&nbsp;&nbsp;: $env:USERDOMAIN\$env:username<br>
Script&nbsp;&nbsp;&nbsp;: $(Convert-Path ($myinvocation.invocationname))<br>
Version: $scriptVersion<br>
Source&nbsp;: $($Env:COMPUTERNAME)<br>
</i>
"@

#create the HTML document
ConvertTo-Html -Head $head -Body $fragments -PostContent $footer |
Out-File -FilePath $path -Encoding ascii

Write-Verbose "Report saved to $path"
if ($ShowReport) {
    #open the finished report
    Write-Verbose "Opening report $path"
    Invoke-Item -Path $path
}

Write-Verbose "Ending $($MyInvocation.Mycommand)"