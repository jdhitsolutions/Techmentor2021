#requires -version 5.1


[cmdletbinding()]

Param (
    [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Computername = $env:Computername,
    [ValidateNotNullOrEmpty()]
    [string]$ReportTitle = "System Inventory Report"
)

Begin {

    Write-Verbose "Starting $($myinvocation.MyCommand)"

    #a nested helper function
    Function Get-CSInfo {

        [cmdletbinding()]

        param(
            [Parameter(Position = 0, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string[]]$Computername = $env:Computername
        )

        Begin {
            Write-Verbose "Starting $($myinvocation.mycommand)"
            #define a hashtable of CIM data to retrieve
            $cimGet = @{
                Win32_OperatingSystem = "CSName", "Version", "BuildNumber"
                Win32_ComputerSystem  = "TotalPhysicalMemory", "NumberOfProcessors", "NumberOfLogicalProcessors"
                Win32_Bios            = "SerialNumber", "Manufacturer", "Version"
            }
        }
        Process {
            ForEach ($computer in $Computername) {
                Write-Verbose "Querying $computer for system information"

                $cimParams = @{
                    ErrorAction  = "Stop"
                    ClassName    = ""
                    Property     = ""
                    Computername = $Computer
                }
                Try {
                    $keys = $cimGet.Keys
                    foreach ($key in $keys) {
                        $cimParams.Classname = $key
                        $cimParams.Property = $cimget[$key]
                        New-Variable -Name $key -Value (Get-CimInstance @cimParams)
                    }
                }
                Catch {
                    $msg = ("There was an error getting system information from {0}. {1}" -f $computer, $_.Exception.Message)
                    Write-Warning $msg
                }

                [pscustomobject]@{
                    "Computername" = $Win32_OperatingSystem.CSName
                    "OS Version"   = $Win32_OperatingSystem.version
                    "OS Build"     = $Win32_OperatingSystem.buildnumber
                    "MemoryGB"     = $Win32_Computersystem.totalphysicalmemory / 1GB -as [int]
                    "CPUs"         = $Win32_Computersystem.numberofprocessors
                    "LogicalC PUs" = $Win32_Computersystem.NumberOfLogicalProcessors
                    "BIOS Serial"  = $Win32_bios.serialnumber
                    "BIOS Version" = $win32_bios.Version
                    "BIOS Mfg"     = $Win32_bios.Manufacturer
                }

            } #foreach
        } #Process
        end {
            Write-Verbose "Ending $($myinvocation.mycommand)"
        }
    } #function

    #define variables used later in the script.
    #this must be left justified
    $head = @'
<style>
body { background-color:#FFFFFF;
       font-family:Tahoma;
       font-size:12pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
tr:nth-child(odd) {background-color: lightgray}
table { margin-left:50px; }
</style>
'@

    #initialize
    $fragments = @()

}

Process {
    ForEach ($computer in $Computername) {
        Write-Verbose "Creating a report for $computer"
        $csinfo = Get-CSInfo -Computername $computer

        #if Get-CSinfo worked, then get disk information
        if ($csinfo) {
            Write-Verbose "Getting disk information from $computer"
            $data = Get-CimInstance -Class Win32_LogicalDisk -Filter 'DriveType=3' `
                -Computername $computer

            $drives = foreach ($item in $data) {
                [PSCustomobject]@{
                    Drive       = $item.DeviceID
                    Volume      = $item.VolumeName
                    Compressed  = $item.Compressed
                    SizeGB      = $item.size / 1GB -as [int]
                    FreeGB      = "{0:N4}" -f ($item.Freespace / 1GB)
                    PercentFree = ($item.Freespace / $item.size) * 100
                }

            } #foreach item

            #add as much other reporting code as you'd like foreach computer
        } #if $csinfo

        #did we get CSInfo and $drives?
        If ($CSInfo -AND $drives) {
                    Write-Verbose "Creating HTML code"

                    #create bar chart for disk space
                    Write-Verbose "Creating chart"
                    $Chartfile = .\New-ConditionalDiskChart.ps1 -Data ($Drives | Sort-Object Drive -Descending)
                    Write-Verbose "Created $($chartfile.fullname)"
                    Write-Verbose "Encoding image"
                    $ImageBits = [Convert]::ToBase64String((Get-Content $($Chartfile.fullname) -Encoding Byte))
                    $ImageHTML = "<img src=data:image/png;base64,$($ImageBits) alt='disk utilization' style='left 50px'/>"
                    $fragments += $csinfo | ConvertTo-Html -As LIST -Fragment -PreContent "<h2>$($csinfo.Computername) Computer Info</h2>" | Out-String
                    $fragments += $drives | ConvertTo-Html -Fragment -PreContent "<h2>$($csinfo.Computername) Disk Info</h2>" | Out-String
                    $fragments += "$ImageHTML <br>"

        } #if $csinfo and $drives
        else {
            Write-Host "Skipping report for $computer since not everything could be retrieved" -ForegroundColor Red
        }
    } #foreach

} #process

End {
    #Write the finished report
    If ($fragments.count -gt 0) {
        $fragments += "<I>Report run $(Get-Date)</I>"
        ConvertTo-HTML -head $head -title $reportTitle -PreContent "<h1>System Inventory Report</h1>" -PostContent $fragments
    }
    Write-Verbose "Finished Creating Report"

    Write-Verbose "Ending $($myinvocation.MyCommand)"
} #end

