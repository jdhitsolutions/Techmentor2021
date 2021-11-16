#requires -version 7.1

Return "This is a walk through demo script file"

#region installing PS 7
# Github
start https://github.com/powershell/powershell

# Microsoft Store
start https://www.microsoft.com/store/productId/9MZ1SNWT0N5D

# Winget https://github.com/microsoft/winget-cli
winget search microsoft.powershell

# PSReleaseTools
# Install-Module PSReleasetools
Get-PSReleaseCurrent
help Install-PowerShell

#endregion

#region new operators

# ternary
10 -gt 1 ? "OK" : "Try again"
10 -gt 11 ? "OK" : "Try again"

(Test-Path c:\blah) ? (Get-ChildItem C:\blah ) : (Write-Warning "Failed to find path")

(Test-Path c:\work) ? (Get-ChildItem C:\work ) : (Write-Warning "Failed to find path")

# run in background
Get-Process &

# range of characters
'a'..'f'
'a'..'f' | ForEach-Object { New-Item "c:\work\Demo-$_.txt" -Force}

#chain operators
help about_Pipeline_Chain_Operators

Get-Service bits -ov b && "Bits status is $($b.Status)"
Get-Service bits || Write-Warning "Can't find service"
Get-Service foo || Write-Warning "Can't find service"

#endregion

#region auto-prediction

#PredictionSource and InlinePredictionColor
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
# 19 seconds
Measure-Command {
    $out = "c:\work", "c:\windows", "c:\scripts", "d:\temp", "c:\users\jeff\documents" |
    ForEach-Object -Process { Invoke-Command -ScriptBlock $t -ArgumentList $_ }
}

#Parallel
#harder to pass variables to runspaces
# 17 seconds
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

<#
PS demos:\> get-variable is*

Name                           Value
----                           -----
IsCoreCLR                      True
IsLinux                        False
IsMacOS                        False
IsWindows                      True

#>
$PSEdition

New-PSSession -computername dom1 -credential company\artd
#setting up SSHKeys makes this easier
New-PSSession -HostName srv1 -SSHTransport -UserName artd
New-PSSession -HostName srv2 -SSHTransport -UserName artd
#Fedora host
New-PSSession -HostName fred-company-com -UserName jeff -SSHTransport

Get-PSSession

<#
PS C:\> get-pssession

 Id Name            Transport ComputerName    ComputerType    State         ConfigurationName     Availability
 -- ----            --------- ------------    ------------    -----         -----------------     ------------
  1 Runspace1       WSMan     dom1            RemoteMachine   Opened        Microsoft.PowerShell     Available
  3 Runspace2       SSH       srv1            RemoteMachine   Opened        DefaultShell             Available
  5 Runspace4       SSH       srv2            RemoteMachine   Opened        DefaultShell             Available
  7 Runspace6       SSH       fred-company-c… RemoteMachine   Opened        DefaultShell             Available

#>

#be careful with aliases
#Sort is not an alias in Linux
Invoke-Command {
    Get-Process |
    Sort-Object WorkingSize |
    Select-Object -first 5} -session (Get-pssession)

<#

 NPM(K)    PM(M)      WS(M)     CPU(s)      Id  SI ProcessName                        PSComputerName
 ------    -----      -----     ------      --  -- -----------                        --------------
      3     1.49       2.49       0.00    1692   0 cmd                                srv1
      3     1.48       2.73       0.00    2744   0 cmd                                srv2
     24     5.11      12.93       1.14     792   0 svchost                            srv1
     17     3.50       5.53      69.02     884   0 svchost                            srv2
      0     0.00       5.77       0.00     953 940 (sd-pam)                           fred-company-com
     16     9.24      15.16       5.70     800   0 svchost                            srv1
      0     0.00       0.00       0.00    2018   0 kworker/u2:2-flush-253:0           fred-company-com
      0     0.00       0.00       0.00     165   0 kworker/u3:0                       fred-company-com
      0     0.00       4.29       0.01     695 695 low-memory-monitor                 fred-company-com
      0     0.00       1.98       0.00     698 698 mcelog                             fred-company-com
     56    49.62      36.96     731.44     948   0 svchost                            srv2
     17     3.40      11.34       8.20     824   0 svchost                            srv1
     14     3.02      10.78       0.17     864   0 svchost                            srv1
     46    16.21      25.82      34.94     992   0 svchost                            srv2
     11     1.97       7.62       1.55    1216   0 svchost                            dom1
     25     6.22      13.74       2.31     968   0 svchost                            dom1
      7     1.77       4.13       0.00    1020   0 svchost                            srv2
     32    12.43      16.72     389.75    1300   0 svchost                            dom1
     32   126.45      20.68       2.81    1684   0 svchost                            dom1
     11     3.49       9.55       0.06    1656   0 svchost                            dom1

#>
#endregion

# What else did you want to know?