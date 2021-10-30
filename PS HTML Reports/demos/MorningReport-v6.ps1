#requires -version 5.1


<#
.Synopsis
Create System Report
.Description
Create a system status report with information gathered from WMI. The default output to the pipeline is a collection of custom objects. You can also use -TEXT to write a formatted text report, suitable for sending to a file or printer, or -HTML to create HTML code. You will need to pipe the results to Out-File if you want to save either the text or html output.

.Parameter Computername
The name of the computer to query. The default is the localhost.
.Parameter Credential
The name of an alternate credential, or a saved credential object.
.Parameter Quick
Run a quick report which means no event logs queries. This speeds up the report generation.
.Parameter ReportTitle
The title for your report. This parameter has an alias of 'Title'.
.Parameter Hours
The number of hours to search for errors and warnings. The default is 24.
.Parameter HTML
Create HTML report. You must pipe to Out-File to save the results.
.Parameter ImagePath
The local path to an image file which can be embedded into the report.
Valid file types are PNG, JPG and GIF. The image will be resize to 120x120.
.Parameter Text
Create a formatted text report. You must pipe to Out-File to save the results.
.Example
PS C:\Scripts\> .\MorningReport.ps1 | Export-Clixml ("c:\work\{0:yyyy-MM-dd}_{1}.xml" -f (get-date),$env:computername)
Preparing morning report for SERENITY
...Operating System
...Computer System
...Services
...Logical Disks
...Network Adapters
...System Event Log Error/Warning since 01/09/2013 09:47:26
...Application Event Log Error/Warning since 01/09/2013 09:47:26

Run a morning report and export it to an XML file with a date stamped file name.
.Example
PS C:\Scripts\> .\MorningReport Quark -Text | Out-file c:\work\quark-report.txt

Run a morning report for a remote computer and save the results to an text file.
.Example
PS C:\Scripts\> .\MorningReport -html -hours 30 | Out-file C:\work\MyReport.htm

Run a morning report for the local computer and get last 30 hours of event log information. Save as an HTML report.
.Example
PS C:\Scripts\> get-content computers.txt | .\Morningreport -quick -html | out-file c:\work\morningreport.htm

Get the list of computers and create a single HTML report without the event log information.

.Link
Get-CimInstance
Get-EventLog
ConvertTo-HTML
.Inputs
String
.Outputs
Custom object, text or HTML code
.Notes
Version     : 5.0
Last Updated: 27 December, 2018
Author      : Jeffery Hicks (@JeffHicks)

Originally published at http://jdhitsolutions.com/blog/2013/02/powershell-morning-report-with-credentials

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/


#>

[cmdletbinding(DefaultParameterSetName = "object")]

Param(
    [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string]$Computername = $env:computername,
    [PSCredential]$Credential,
    [ValidateNotNullOrEmpty()]
    [alias("title")]
    [string]$ReportTitle = "Morning System Report",
    [ValidateScript( {$_ -ge 1})]
    [int]$Hours = 24,
    [switch]$Quick,
    [Parameter(ParameterSetName = "HTML")]
    [switch]$HTML,
    [Parameter(ParameterSetName = "HTML")]
    [string]$ImagePath,
    [Parameter(ParameterSetName = "TEXT")]
    [switch]$Text
)

Begin {

    #script internal version number used in output
    [string]$reportVersion = "6.0"

    #where did this file come from
    $meta = [PSCustomObject]@{
        Author        = "$env:userdomain\$env:USERNAME"
        "Report Date" = "$((Get-Date).ToUniversalTime()) UTC"
        Source        = $($MyInvocation.MyCommand).path
        Version       = $reportVersion
        Originated    = $env:computername
    }

    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Version $reportVersion"
    #if  an image path is specified, convert it to Base64
    if ($ImagePath -AND (Test-Path $ImagePath)) {
        Write-Verbose "Inserting image from $ImagePath"
        $ImageBits = [Convert]::ToBase64String((Get-Content $ImagePath -Encoding Byte))
        $ImageFile = Get-Item $ImagePath
        $ImageType = $ImageFile.Extension.Substring(1)
        $ImageHead = "<Img src='data:image/$ImageType;base64,$($ImageBits)' Alt='$($ImageFile.Name)' style='float:left' width='120' height='120' hspace=10>"
    }

    #region define html head
    <#
define some HTML style
here's a source for HTML color codes
http://www.immigration-usa.com/html_colors.html

the code must be left justified
#>
    $head = @"
<style>
h2 {
width:95%;
background-color:#7BA7C7;
font-family:Tahoma;
font-size:12pt;
font-color:Black;
}
body { background-color:#FFFFFF;
       font-family:Tahoma;
       font-size:10pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
.footer {
    font-size:8pt;
    margin-left:0px;
    table-layout:auto;
}
.footer tr:nth-child(odd) {background-color: white}
.footer td,tr {
    border-collapse:collapse;
    border:none;

}
table.footer {width:75%;}

.right {
    text-align: right;
}
</style>
<Title>$ReportTitle</Title>
$ImageHead
<H1> The Morning Report</H1>
<H4>$(Get-Date -DisplayHint Date | Out-String)</H4>
<br>
"@
    #endregion

    #prepare HTML code
    Write-Verbose "Initialize fragment array"
    $fragments = @()

    #Write-Progress parameters
    $progParam = @{
        Activity         = $MyInvocation.MyCommand
        Status           = ""
        CurrentOperation = ""
        PercentComplete  = 0
    }
} #Begin

Process {

    $progParam.Status = $Computername
    $progParam.CurrentOperation = "Connecting to computer"
    Write-Progress @progParam
    #region ping test

    #set a default value for the ping test
    $ok = $False

    If ($computername -eq $env:computername) {
        #local computer so no ping test is necessary
        $OK = $True
        Write-Verbose "Using local computer"
    }
    elseIf (($computername -ne $env:computername) -AND (Test-Connection -ComputerName $computername -quiet -Count 2)) {
        #not local computer and it can be pinged so proceed
        $OK = $True
        Write-Verbose "Querying remote computer $Computername"
    }

    #endregion

    #create a CIMSession
    $csParam = @{
        ErrorAction  = 'stop'
        Computername = $Computername
    }

    if ($Credential) {
        $csParam.add('Credential', $Credential)
    }

    Try {
        Write-Verbose "Creating temporary CIMSession"
        $session = New-CimSession @csParam
    }
    Catch {
        Write-Warning "Failed to create CIMSession to $Computername"
        #Bail out
        Return
    }
    #region define a parameter hashtable
    $paramhash = @{
        Classname   = "Win32_OperatingSystem"
        CIMSession  = $session
        ErrorAction = "Stop"
    }

    #endregion

    If ($OK) {

        Try {
            #get Operating system information from WMI
            Write-Verbose "Querying $($paramhash.classname)"
            $os = Get-CimInstance @paramhash
            #set a variable to indicate WMI can be reached
            $cim = $True
        }
        Catch {
            Write-Warning "Failed to connect to $($computername.ToUpper()). $($_.exception.message)"
        }

        if ($cim) {

            Write-Verbose "Preparing morning report for $($os.CSname)"
            $progParam.CurrentOperation = "Preparing morning report"
            $progParam.PercentComplete = 5
            Write-Progress @progParam
            #region OS Summary
            Write-Verbose "...Operating System"
            $progParam.CurrentOperation = "...Operating system"
            $progParam.PercentComplete += 15
            Write-Progress @progParam

            $osdata = $os | Select-Object @{Name = "Computername"; Expression = {$_.CSName}},
            @{Name = "OS"; Expression = {$_.Caption}},
            @{Name = "ServicePack"; Expression = {$_.CSDVersion}},
            free*memory, totalv*, NumberOfProcesses, LastBootUptime,
            @{Name = "Uptime"; Expression = {(Get-Date) - $_.LastBootupTime}}
            #endregion

            #region Computer system
            Write-Verbose "...Computer System"
            $progParam.CurrentOperation = "...computer system"
            $progParam.PercentComplete += 15
            Write-Progress @progParam

            $paramhash.Classname = "Win32_ComputerSystem"
            $cs = Get-CimInstance @paramhash -Property Status, Manufacturer, Model, SystemType, NumberofLogicalProcessors,NumberofProcessors
            $csdata = $cs | Select-Object Status, Manufacturer, Model, SystemType, Number*
            #endregion

            #region Get Service information
            Write-Verbose "...Services"
            $progParam.CurrentOperation = "...Services"
            $progParam.PercentComplete += 15
            Write-Progress @progParam

            #get all services via WMI and group into a hash table
            $paramhash.Classname = "Win32_Service"
            $cimservices = Get-CimInstance @paramhash
            $services = $cimservices | Group-Object State -AsHashTable -AsString

            #get services set to auto start that are not running
            $failedAutoStart = ($cimservices).where( { ($_.startmode -eq "Auto") -AND ($_.state -ne "Running")} )
            #endregion

            #region Disk Utilization
            Write-Verbose "...Logical Disks"
            $progParam.CurrentOperation = "...disks"
            $progParam.PercentComplete += 15
            Write-Progress @progParam
            $paramhash.Classname = "Win32_LogicalDisk"
            $paramhash.Add("Filter", "Drivetype=3")
            $disks = Get-CimInstance @paramhash
            $diskData = $disks | Select-Object DeviceID,
            @{Name = "SizeGB"; Expression = {$_.size / 1GB -as [int]}},
            @{Name = "FreeGB"; Expression = {"{0:N2}" -f ($_.Freespace / 1GB)}},
            @{Name = "PercentFree"; Expression = {"{0:P2}" -f ($_.Freespace / $_.Size)}}

            #endregion

            #region NetworkAdapters
            Write-Verbose "...Network Adapters"
            $progParam.CurrentOperation = "...network adapters"
            $progParam.PercentComplete += 15
            Write-Progress @progParam
            $paramhash.classname = "Win32_NetworkAdapter"
            $paramhash.filter = "MACAddress Like '%'"
            #get NICS that have a MAC address only
            $nics = Get-CimInstance @paramhash
            $nicdata = $nics | ForEach-Object {
                $tempHash = @{Name = $_.Name; DeviceID = $_.DeviceID; AdapterType = $_.AdapterType; MACAddress = $_.MACAddress}
                #get related configuation information
                $config = $_ | Get-CimAssociatedInstance -ResultClassName "Win32_NetworkadapterConfiguration"

                #add to temporary hash
                $tempHash.Add("IPAddress", $config.IPAddress)
                $tempHash.Add("IPSubnet", $config.IPSubnet)
                $tempHash.Add("DefaultGateway", $config.DefaultIPGateway)
                $tempHash.Add("DHCP", $config.DHCPEnabled)
                #convert lease information if found
                if ($config.DHCPEnabled -AND $config.DHCPLeaseObtained) {
                    $tempHash.Add("DHCPLeaseExpires", ($config.DHCPLeaseExpires))
                    $tempHash.Add("DHCPLeaseObtained", ($config.DHCPLeaseObtained))
                    $tempHash.Add("DHCPServer", $config.DHCPServer)
                }

                New-Object -TypeName PSObject -Property $tempHash

            }
            #endregion

            If ($Quick) {
                Write-Verbose "Skipping event log queries"
            }
            #region Event log queries
            else {
                #Event log errors and warnings in the last $Hours hours
                $last = (Get-Date).AddHours( - $Hours)
                #define a hash table of parameters to splat to Get-Eventlog
                $GetEventLogParam = @{
                    LogName   = "System"
                    EntryType = "Error", "Warning"
                    After     = $last
                }

                #System Log
                Write-Verbose "...System Event Log Error/Warning since $last"
                $progParam.CurrentOperation = "...Eventlogs since $last"
                $progParam.PercentComplete += 15
                Write-Progress @progParam
                #hashtable of optional parameters for Invoke-Command
                $InvokeCommandParam = @{
                    Computername = $Computername
                    ArgumentList = $GetEventLogParam
                    ScriptBlock  = {Param ($params) Get-EventLog @params }
                }

                if ($Credential) {
                    $InvokeCommandParam.Add("Credential", $Credential)
                }

                $syslog = Invoke-Command @InvokeCommandParam

                $syslogdata = $syslog | Select-Object TimeGenerated, EventID, Source, Message

                #Application Log
                #rite-Host "...Application Event Log Error/Warning since $last"
                #update the hashtable
                $GetEventLogParam.LogName = "Application"

                #update invoke-command parameters
                $InvokeCommandParam.ArgumentList = $GetEventLogParam

                $applog = Invoke-Command @InvokeCommandParam
                $applogdata = $applog | Select-Object TimeGenerated, EventID, Source, Message
            }
            #endregion
        } #if wmi is ok

        #write results depending on parameter set
        $footer = "Report v{3} run {0} by {1}\{2}" -f (Get-Date), $env:USERDOMAIN, $env:USERNAME, $reportVersion

        #region Create HTML
        if ($HTML) {
            #add each computer to a navigation menu in the header
            Write-Verbose "Preparing HTML"
            $head += ("<a href=#{0}_Summary>{0}</a> " -f $computername.ToUpper())

            $fragments += ("<H2><a name='{0}_Summary'>{0}: System Summary</a></H2>" -f $computername.ToUpper())
            $fragments += $osdata | ConvertTo-Html -as List -Fragment
            $fragments += $csdata | ConvertTo-Html -as List -Fragment

            #insert navigation bookmarks
            $nav = @"
<br>
<a href=#{0}_Services>{0} Services</a>
<a href='#{0}_NoAutoStart'>{0} Failed Auto Start</a>
<a href='#{0}_Disks'>{0} Disks</a>
<a href='#{0}_Network'>{0} Network</a>
<a href='#{0}_SysLog'>{0} System Log</a>
<a href='#{0}_AppLog'>{0} Application Log</a>
<br>
"@ -f $Computername.ToUpper()

            #add a link to the document top
            $nav += "`n<a href='#' target='_top'>Top</a>"
            $fragments += $nav

            $fragments += "<br clear='All'>"

            $fragments += ConvertTo-Html -Fragment -PreContent ("<H2><a name='{0}_Services'>{0}: Services</a></H2>" -f $computername.ToUpper())
            $services.keys | ForEach-Object {
                $fragments += ConvertTo-Html -Fragment -PreContent "<H3>$_</H3>"
                $fragments += $services.$_ | Select-Object Name, Displayname, StartMode| ConvertTo-HTML -Fragment
                #insert navigation link after each section
                $fragments += $nav
            }

            $fragments += $failedAutoStart | Select-Object Name, Displayname, StartMode, State |
                ConvertTo-Html -Fragment -PreContent ("<h3><a name='{0}_NoAutoStart'>{0}: Failed Auto Start</a></h3>" -f $computername.ToUpper())
            $fragments += $nav

            $fragments += $diskdata | ConvertTo-HTML -Fragment -PreContent ("<H2><a name='{0}_Disks'>{0}: Disk Utilization</a></H2>" -f $computername.ToUpper())
            $fragments += $nav

            #convert nested object array properties to strings
            $fragments += $nicdata | Select-Object Name, DeviceID, DHCP*, AdapterType, MACAddress,
            @{Name = "IPAddress"; Expression = {$_.IPAddress | Out-String}},
            @{Name = "IPSubnet"; Expression = {$_.IPSubnet | Out-String}},
            @{Name = "IPGateway"; Expression = {$_.DefaultGateway | Out-String}}  |
                ConvertTo-HTML -Fragment -PreContent ("<H2><a name='{0}_Network'>{0}: Network Adapters</a></H2>" -f $computername.ToUpper())
            $fragments += $nav

            $fragments += $syslogData | ConvertTo-HTML -Fragment -PreContent ("<H2><a name='{0}_SysLog'>{0}: System Event Log Summary</a></H2>" -f $computername.toUpper())
            $fragments += $nav

            $fragments += $applogData | ConvertTo-HTML -Fragment -PreContent ("<H2><a name='{0}_AppLog'>{0}: Application Event Log Summary</a></H2>" -f $computername.toUpper())
            $fragments += $nav

        }
        #endregion
        #region Create TEXT
        elseif ($TEXT) {
            Write-Verbose "Prepare formatted text"
            $ReportTitle
            "-" * ($ReportTitle.Length)
            "System Summary"
            $osdata | Out-String
            $csdata | Format-List | Out-String
            "Services"
            $services.keys | ForEach-Object {
                $services.$_ | Select-Object Name, Displayname, StartMode, State
            } | Format-List | Out-String
            "Failed Autostart Services"
            $failedAutoStart | Select-Object Name, Displayname, StartMode, State
            "Disk Utilization"
            $diskdata | Format-table -AutoSize | Out-String
            "Network Adapters"
            $nicdata | Format-List | Out-String
            "System Event Log Summary"
            $syslogdata | Format-List | Out-String
            "Application Event Log Summary"
            $applogdata | Format-List | Out-String
            $Footer
        }
        #endregion
        #region Create custom object
        else {
            #Write data to the pipeline as part of a custom object

            New-Object -TypeName PSObject -Property @{
                OperatingSystem = $osdata
                ComputerSystem  = $csdata
                Services        = $services.keys | ForEach-Object {$services.$_ | Select-Object Name, Displayname, StartMode, State}
                FailedAutoStart = $failedAutoStart | Select-Object Name, Displayname, StartMode, State
                Disks           = $diskData
                Network         = $nicData
                SystemLog       = $syslogdata
                ApplicationLog  = $applogdata
                ReportVersion   = $reportVersion
                RunDate         = Get-Date
                RunBy           = "$env:USERDOMAIN\$env:USERNAME"
            }
        }
        #endregion
    } #if OK

    else {
        #can't ping computer so fail
        Write-Warning "Failed to ping $computername"
    }
} #process

End {
    #if HTML finish the report here so that if piping in
    #computer names we get one report for all computers
    If ($HTML) {
        Write-Verbose "Converting to HTML"

        [xml]$footer = $meta | ConvertTo-Html -Fragment -as List
        #insert css tags into this table
        $class = $footer.CreateAttribute("class")
        $footer.table.SetAttribute("class", "footer")
        for ($i = 0; $i -le $footer.table.tr.count - 1; $i++) {
            $class = $footer.CreateAttribute("class")
            $class.value = 'footer'
            [void]$footer.table.tr[$i].attributes.append($class)

            $class = $footer.CreateAttribute("class")
            $class.value = 'right'
            [void]$footer.table.tr[$i].item("td").attributes.append($class)
        }
        $head += "<br><br><hr>"
        ConvertTo-Html -Head $head -Title $ReportTitle -PreContent ($fragments | out-String) -PostContent "<br><i>$($footer.innerxml)<i>"
    }
    #clean up
    if ($session) {
        Write-Verbose "Removing temporary CIM Session"
        $session | Remove-Cimsession
    }

    Write-Verbose "Finished!"
    $progParam.PercentComplete = 100
    $progParam.CurrentOperation = "Completed"
    Write-Progress @progParam

    Write-Verbose "Ending $($myinvocation.mycommand)"
}

#end of script