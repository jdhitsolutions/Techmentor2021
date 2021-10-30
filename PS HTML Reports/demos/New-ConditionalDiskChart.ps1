#requires -version 5.1

<#
This was written for this data:

   $data = Get-WmiObject -Class Win32_LogicalDisk -Filter 'DriveType=3' `
   -ComputerName $computer

    $drives = foreach ($item in $data) {
        $prophash = [ordered]@{
        Drive = $item.DeviceID
        Volume = $item.VolumeName
        Compressed = $item.Compressed
        SizeGB = $item.size/1GB -as [int]
        FreeGB = "{0:N4}" -f ($item.Freespace/1GB)
        PercentFree = ($item.Freespace/$item.size) * 100
        }
        New-Object PSObject -Property $prophash
    } #foreach item

#>
Param(
    [string]$Path = "$env:temp\diskchart.png",
    [object]$Data,
    [string]$Legend,
    [string]$LegendText = "My Data",
    [string]$Title,
    [string]$SeriesName = "My Series",
    [switch]$Label,
    [int]$Width = 600,
    [int]$Height = 400,
    [ValidateSet("TopLeft", "TopCenter", "TopRight",
        "MiddleLeft", "MiddleCenter", "MiddleRight",
        "BottomLeft", "BottomCenter", "BottomRight")]
    [String]$Alignment = "TopCenter"
)

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

#Create the chart object
$Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$Chart.Width = $Width
$Chart.Height = $Height

#Create the chart area and set it to be 2D
$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$ChartArea.Area3DStyle.Enable3D = $False
$Chart.ChartAreas.Add($ChartArea)

#Set the title and alignment
if ($Title) {
    [void]$Chart.Titles.Add("$Title")
    $Chart.Titles[0].Alignment = $Alignment
    $Chart.Titles[0].Font = "Verdana,16pt"
}
$SeriesName = $SeriesName
$Chart.Series.Add($SeriesName) | Out-Null
$Chart.Series[$SeriesName].ChartType = "bar"
if ($legend) {
    #add a legend
    $chart.Legends.Add($Legend) | Out-Null
    $chart.Series[$SeriesName].IsVisibleInLegend = $True
    $chart.Series[$SeriesName].LegendText = $LegendText
}

$Data | ForEach-Object {

    # Create the data series and add bar chart
    $point = $Chart.Series[$SeriesName].Points.AddXY($_.drive, $_.percentfree)
    #adjust color
    if (($_.percentFree -as [int]) -le 12) {
        $chart.Series[$SeriesName].points[$point].Color = "red"
    }
    elseif ( ($_.percentFree -as [int]) -le 40 ) {
        $chart.Series[$SeriesName].points[$point].Color = "yellow"
    }
    else {
        $chart.Series[$SeriesName].points[$point].Color = "green"
    }
    #add a label
    if ($label) {
        $chart.Series[$SeriesName].points[$point].Label = "{0:P2}" -f ($_.freeGB / $_.sizeGB)
        $chart.Series[$SeriesName].AxisLabel = $_.drive
    }

} | Out-Null

#Save the chart to a png file
$Chart.SaveImage($Path, "png")
Get-Item $path
