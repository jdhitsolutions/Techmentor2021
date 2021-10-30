
return "This is a walkthough demo script file."

psedit .\demo-htmlsteps.ps1
psedit .\demo-conditionalformatting.ps1

psedit .\SQLServerReport.ps1
#look at previously created report
invoke-item .\sqlreport.htm
#notice the footer

#Additional examples
#embedding graphics
psedit .\New-SystemReport.ps1
.\new-systemreport | out-file $env:temp\sr.htm
invoke-item $env:temp\sr.htm

psedit .\demo-htmlbarchart.ps1
.\demo-htmlbarchart.ps1
invoke-item .\drivereport.htm

#using javascript
psedit .\MorningReport-v6.ps1
.\MorningReport-v6.ps1 -html -ImagePath .\antique-watch.png -verbose | out-file $env:temp\mr.htm
invoke-item $env:temp\mr.htm
psedit $env:temp\mr.htm

psedit .\New-HVHealthReport.ps1
#run this in the console. The RemoteRegistry service must be running
# .\New-HVHealthReport.ps1 -Performance -RecentCreated 180 -LastUsed 180 -Path .\hvhealth.htm

psedit .\Get-VMHostStatusReport.ps1
.\Get-VMHostStatusReport.ps1 -ComputerName $env:COMPUTERNAME
Inovke-item .\HyperVHostStatus.htm