#requires -version 3.0

return "This is a walk through demo"

#demo each region interactively

#these are typical IT Pro examples
$computername = $env:COMPUTERNAME

#region Start with good data

$data = Get-Ciminstance Win32_Service -ComputerName $computername |
    Select-Object Name, Displayname, State, StartMode

#view the data
$data

#endregion

#region A simple report using an external style sheet

#here's my style sheet
#samples at https://github.com/jdhitsolutions/SampleCSS
psedit .\blue.css

#the name of the file to create
$file = Join-Path -path $env:temp -childpath "basic.htm"

#take the data and convert to html using the style sheet
$convParams = @{
    Title       = "$computername Service Report"
    CssUri      = "$(Convert-Path .\blue.css)"
    preContent  = "<h2>$computername</h2>"
    postContent = "<br><I>$(get-date)</I>"
}

$data | ConvertTo-Html @convParams | Out-file $file -Encoding ascii

#view the resulting file
Invoke-Item $file

#view the source
psedit $file

#endregion

#region use fragments

#I want to organize the output
$data | Group-Object -property startmode

$fragments = "<h2>$computername</h2>"

$data | Group-Object -property startmode | ForEach-Object {
    $fragments += "<H3>$($_.Name) ($($_.Count))</H3>"
    $fragments += $_.Group | ConvertTo-Html -Fragment
}

#embed the style sheet
$convParams = @{
    Title       = "$computername Service Status"
    Body        = $fragments
    Head        = "<style>$((Get-Content $(Convert-Path .\blue.css)))</style>"
    postContent = "<br><I>$(get-date)</I>"
}

ConvertTo-html @convParams | Out-File $file

#view the resulting file
Invoke-Item $file
psedit $file

#endregion

#region parse html for conditional formatting

$fragments = "<h2>$computername</h2>"
$fragments += $data | ConvertTo-Html -Fragment

#fragments is just text
$fragments

#find all stopped services that should be running and change
#the font color to red
$revised = $fragments -replace "<td>Stopped</td><td>Auto</td>", "<td style='color:red'>Stopped</td><td>Auto</td>"

#I'm going to use a CSS file with these conditions
#assuming this is being run from the ISE
psedit .\sample.css

$convParams = @{
    Title       = "$computername Service Status"
    Body        = $revised
    Head        = "<style>$((Get-Content $(Convert-Path .\sample.css)))</style>"
    postContent = "<br><I>$(get-date)</I>"
}

ConvertTo-html @convParams | Out-File $file

#view the resulting file
Invoke-Item $file

#endregion
