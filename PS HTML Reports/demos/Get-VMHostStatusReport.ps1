#requires -version 5.1

[cmdletbinding()]
Param(
    [Parameter(Position = 0 , Mandatory, HelpMessage = "Enter the name of a Hyper-V Host")]
    [string]$ComputerName,
    [Parameter(ParameterSetName = "Computername")]
    [PSCredential]$Credential,

    [ValidatePattern("\w+\.(html|htm)$")]
    [string]$Path = ".\HyperVHostStatus.htm"
)

#region get VM Host Status

#dot source the required script with the Get-VMHostStatus function
. $PSScriptRoot\Get-VMHostStatus.ps1

#remove the Path from boundparameters since it isn't part of the parameters
#for Get-VMHostStatus and I want to eventually splat $PSBoundParameters
if ($PSBoundParameters.ContainsKey("Path")) {
    [void]$PSBoundParameters.Remove("Path")
}

#get the VMHost data passing bound parameters
Try {
    $data = Get-VMHostStatus @PSBoundParameters -ErrorAction Stop
}
Catch {
    Throw $_
}

#endregion

#region define the CSS style in the head section

#get sample CSS from https://github.com/jdhitsolutions/SampleCSS
$head = @"
<Title>Hyper-V Host Status</Title>
<style>
body {
    background-color: rgb(233, 223, 223);
    font-family: Monospace;
    font-size: 12pt;
}

td,th {
    border: 0px solid black;
    border-collapse: collapse;
    white-space: pre;
}

th {
    color: white;
    background-color: black;
}

table,tr,td,th {
    padding: 3px;
    margin: 0px;
    white-space: pre;
}

tr:nth-child(odd) {
    background-color: lightgray
}

table {
    margin-left: 25px;
    width: 50%;
}

h2 {
    font-family: Tahoma;
}

.footer {
    color: green;
    margin-left: 25px;
    font-family: Tahoma
    font-size: 7pt;
    font-style: italic;
}
</style>
"@

#endregion

#region get some additional data

$s = New-PSSession @PSBoundParameters

$procdata = Invoke-Command -ScriptBlock { Get-CimInstance Win32_Computersystem -Property 'NumberOfLogicalProcessors', 'NumberOfProcessors' } -Session $s
$hostdetail = Invoke-Command -ScriptBlock { Get-CimInstance Win32_OperatingSystem -Property "Caption", "Version" } -Session $s
$detail = "{0} version {1}" -f $hostdetail.caption, $hostdetail.version
Remove-PSSession $s

#endregion

#region define the pieces of the HTML report as fragments

$fragments = @()
$fragments += "<H1>Hyper-V Host Status</H1>"
$fragments += "<H2 title = '$detail'>$($data.computername)<H2>"

$fragments += "<H3>Memory</H3>"
$fragments += $data | Select-Object -Property *memory*, TotalPctDemand | ConvertTo-Html -Fragment -As List
$fragments += "<H3>Processor</H3>"

$fragments += $data | Select-Object -Property @{Name = "ProcessorCount"; Expression = { $procdata.NumberOfProcessors } },
@{Name = "LogicalProcessorCount"; Expression = { $procdata.NumberOfLogicalProcessors } },
PctProcessorTime, Logical* | ConvertTo-Html -Fragment -As list

$fragments += "<H3>Virtual Machines</H3>"
$fragments += $data | Select-Object -Property *VMs | ConvertTo-Html -Fragment -As List

$fragments += "<H3>Virtual Machine Health</H3>"
$fragments += $data | Select-Object -Property Critical, Healthy | ConvertTo-Html -Fragment -As table

#a future version might include additional network-related values
$fragments += "<H3>Networking</H3>"
$fragments += $data | Select-Object VMSwitchBytesSec, VMSwitchPacketsSec | ConvertTo-Html -Fragment

$fragments += "<H3>Other</H3>"
$fragments += $data | Select-Object Uptime, TotalProcesses, PctFreeDisk | ConvertTo-Html -Fragment -As table

#endregion

#region create an object with footer information so it can be displayed neatly

[xml]$meta = [pscustomobject]@{
    "Report Run"  = (Get-Date)
    Author        = "$env:USERDOMAIN\$env:USERNAME"
    Script        = (Convert-Path $MyInvocation.InvocationName).Replace("\", "/")
    ScriptVersion = '0.9.3'
    Source        = $env:COMPUTERNAME
} | ConvertTo-Html -Fragment -As List

$meta.CreateAttribute("Class") | Out-Null
$meta.table.SetAttribute("class", "footer")

#endregion

#region assemble the final HTML report

ConvertTo-Html -Head $head -Body $fragments -PostContent "<br><br>$($meta.innerxml)" |
Out-File -FilePath $Path -Encoding utf8

Write-Host "See $(Convert-Path $path) for the finished report file." -ForegroundColor green

#endregion