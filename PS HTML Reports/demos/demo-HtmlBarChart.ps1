#requires -version 5.1

[cmdletbinding()]

Param (
  [Parameter(Position = 0)]
  [ValidateNotNullorEmpty()]
  [string[]]$Computername = @($env:computername),
  [ValidateNotNullorEmpty()]
  [string]$Path = "drivereport.htm"
)

$Title = "Drive Report"

#embed a stylesheet in the html header
$head = @"
<style>
body { background-color:#FFFFCC;
       font-family:Tahoma;
       font-size:10pt; }
td, th { border:1px solid #000033;
         border-collapse:collapse; }
th { color:white;
     background-color:#000033; }
table, tr, td, th { padding: 0px; margin: 0px }
table { margin-left:8px; }
</style>
<Title>$Title</Title>
<br>
"@

#define an array for html fragments
$fragments = @()
$fragments += "<H2>Drive Report</H2>"
#get the drive data
$data = Get-CimInstance -ClassName Win32_logicaldisk -Filter "drivetype=3" -computer $computername

#group data by computername
$groups = $Data | Group-Object -Property SystemName

#this is the graph character
#$g= '&#9608;'  #"|"  #[char]9608
[string]$g = "#9608;"
[string]$g2 = "&$g"

#create html fragments for each computer
#iterate through each group object

ForEach ($computer in $groups) {

  $fragments += "<H3>$($computer.Name)</H3>"

  #define a collection of drives from the group object
  $Drives = $computer.group

  #create an html fragment
  $html = $drives | Select-Object @{Name = "Drive"; Expression = { $_.DeviceID } },
  @{Name = "SizeGB"; Expression = { $_.Size / 1GB -as [int] } },
  @{Name = "UsedGB"; Expression = { "{0:N2}" -f (($_.Size - $_.Freespace) / 1GB) } },
  @{Name = "FreeGB"; Expression = { "{0:N2}" -f ($_.FreeSpace / 1GB) } },
  @{Name = "Usage"; Expression = {
      $UsedPer = (($_.Size - $_.Freespace) / $_.Size) * 100
      $UsedGraph = $g * ($UsedPer / 2)
      $FreeGraph = $g * ((100 - $UsedPer) / 2)
      #I'm using place holders for the < and > characters
      "xopenFont color=Redxclose{0}xopen/FontxclosexopenFont Color=Greenxclose{1}xopen/fontxclose" -f $usedGraph, $FreeGraph
    }
  } | ConvertTo-Html -Fragment

  #fix special character by inserting the & character.
  $html = $html -replace $g, $g2
  #replace the tag place holders. It is a hack but it works.
  $html = $html -replace "xopen", "<"
  $html = $html -replace "xclose", ">"

  #add to fragments
  $Fragments += $html

  #insert a return between each computer
  $fragments += "<br>"

} #foreach computer

#add a footer
$footer = ("<br><I>Report run {0} by {1}\{2}<I>" -f (Get-Date -DisplayHint date), $env:userdomain, $env:username)
$fragments += $footer

#write the result to a file
ConvertTo-Html -Head $head -Body $fragments | Out-File $Path -Encoding ascii