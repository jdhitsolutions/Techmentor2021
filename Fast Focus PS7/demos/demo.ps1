#requires -version 7.0

Return "This is a walk through demo script file"

#region installing PS 7
# Github
start https://github.com/powershell/powershell

# Microsoft Store
start https://www.microsoft.com/store/productId/9MZ1SNWT0N5D

# Winget https://github.com/microsoft/winget-cli
winget search microsoft.powershell
# PSReleaseTools
Get-PSReleaseCurrent
help Install-PowerShell

#endregion

#region new operators

# ternary
10 -gt 1 ? "OK" : "Try again"
10 -gt 11 ? "OK" : "Try again"

(Test-Path c:\blah) ? (Get-ChildItem C:\blah ): (Write-Warning "Failed to find path")

(Test-Path c:\work) ? (Get-ChildItem C:\work ): (Write-Warning "Failed to find path")

# run in background
Get-Process &

# range of characters
'a'..'f'
'a'..'f' | foreach { New-Item "c:\work\Demo-$_.txt" -Force}

#chain operators
help about_Pipeline_Chain_Operators

Get-Service bits -ov b && "Bits status is $($b.Status)"
Get-Service bits || Write-Warning "Can't find service"
Get-Service foo || Write-Warning "Can't find service"

#endregion

#region auto-prediction

#PreditionSource and InlinePredictionColor
Get-PSReadLineOption
Set-PSReadLineOption -PredictionSource History -Colors @{InlinePrediction = "`e[4;38;5;219m" }
#chords may not conflict with VS Code
Set-PSReadLineKeyHandler -Chord "ctrl+f" -Function ForwardWord

#endregion

#region foreach -parallel

1..10 | ForEach-Object -Parallel {
    Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 10)
    $m = "[$(Get-Date -f 'hh:mm:ss.ffff')] $_"
    $m
}

#not always faster
Measure-Command {
    $a = 1..5000 | ForEach-Object { [math]::Sqrt($_) * 2 }
}

#34ms

Measure-Command {
    $a = 1..5000 | ForEach-Object -Parallel { [math]::Sqrt($_) * 2 }
}


$t = {
    param ([string]$Path)
    Write-Host "[$(Get-Date)] Processing $Path" -ForegroundColor yellow
    Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object -Sum -Property Length -Maximum -Minimum |
    Select-Object @{Name = "Computername"; Expression = { $env:COMPUTERNAME } },
    @{Name = "Path"; Expression = { Convert-Path $path } },
    Count, Sum, Maximum, Minimum
    Write-Host "[$(Get-Date)] Finished processing $Path" -ForegroundColor Yellow
}

#NOT Parallel
Measure-Command {
    $out = "c:\work", "c:\windows", "c:\scripts", "d:\temp", "c:\users\jeff\documents" |
    ForEach-Object -Process { Invoke-Command -ScriptBlock $t -ArgumentList $_ }
}

#Parallel
#harder to pass variables to runspaces
Measure-Command {
    $out = "c:\work", "c:\windows", "c:\scripts", "d:\temp", "c:\users\jeff\documents" |
    ForEach-Object -Parallel {
        $path = $_
        Write-Host "[$(Get-Date)] Processing $Path" -ForegroundColor yellow
        Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Sum -Property Length -Maximum -Minimum |
        Select-Object @{Name = "Computername"; Expression = { $env:COMPUTERNAME } },
        @{Name = "Path"; Expression = { Convert-Path $path } },
        Count, Sum, Maximum, Minimum
        Write-Host "[$(Get-Date)] Finished processing $Path" -ForegroundColor Yellow
    }
}

#endregion

#region ssh remoting

#get ssh working natively before using in PowerShell
#this may not work in VSCode
Enter-PSSession -HostName srv1 -UserName artd -SSHTransport
Enter-PSSession -HostName fred-company-com -UserName jeff -SSHTransport

#endregion

#region cross-platform

Get-Variable is*
$PSEdition

New-PSSession -computername dom1 -credential company\artd
#setting up SSHKeys makes this easier
New-PSSession -HostName srv1 -SSHTransport -UserName artd
New-PSSession -HostName srv2 -SSHTransport -UserName artd
#Fedora host
New-PSSession -HostName fred-company-com -UserName jeff -SSHTransport

Get-PSSession

#be careful with aliases
#Sort is not an alias in Linux
Invoke-Command { Get-Process | Sort-Object WorkingSize | Select-Object -first 5} -session (Get-pssession)

#endregion

# What else did you want to know?