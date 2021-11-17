#requires -version 3.0

Return "This is a walkthrough demo"

#here's another way to handle conditional formatting
#this can be run as a standalone script
$computername = $env:COMPUTERNAME

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

#You have to have objects
$drives

#embed the style in the html header
# point out the conditional formatting .danger and .warn
$head = @'
<Title>Volume Report</Title>
<style>
body
{
 background-color:#FFFFFF;
 font-family:Tahoma;
 font-size:10pt;
}
td, th
{
 border:1px solid black;
 border-collapse:collapse;
}
th
{
 color:white;
 background-color:black;
}
table, tr, td, th { padding: 5px; margin: 0px;border-spacing:0 }
table { margin-left:50px; }
.danger {background-color: red}
.warn {background-color: yellow}
</style>
'@

#create an xml document from the HTML fragment
[xml]$html = $drives | ConvertTo-Html -Fragment

#explore the html document
$html.table
$html.table.tr
$html.InnerXml

#check each row, skipping the TH header row
for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
    $class = $html.CreateAttribute("class")
    #check the value of the last column and assign a class to the row
    if (($html.table.tr[$i].td[-1] -as [int]) -le 25) {
        $class.value = "danger"
        [void]$html.table.tr[$i].Attributes.Append($class)
    }
    elseif (($html.table.tr[$i].td[-1] -as [int]) -le 35) {
        $class.value = "warn"
        [void]$html.table.tr[$i].Attributes.Append($class)
    }
}

#create the final report from the innerxml which should be html code
$body = @"
<H1>Volume Utilization for $Computername</H1>
$($html.innerxml)
"@

#put it all together
ConvertTo-Html -Head $head -PostContent "<br><i>$(Get-Date)</i>" -Body $body |
Out-File "$env:temp\drives.htm" -Encoding ascii

Invoke-Item "$env:temp\drives.htm"

#here's a sample from home that has the formatting
Invoke-Item .\samples\bovine320-drives.htm
