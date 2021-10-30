#requires -version 5.1

Function Export-Eventlog {
    [cmdletbinding()]
    Param (
        [Parameter(Position = 0)]
        [ValidateSet("System", "Application", "Security", "Directory Service", "DNS Server")]
        [string]$Log = "System",
        [ValidateRange(10, 1000)]
        [int]$Count = 100,
        [Parameter(
            Position = 1,
            Mandatory,
            HelpMessage = "What type of export file do you want to create? Valid choices are CSV, XML, CLIXML."
        )]
        [ValidateSet("csv", "xml", "clixml" )]
        [string]$Export,
        [ValidateNotNullorEmpty()]
        [string]$Path = "C:\Work",
        [string]$Computername = $env:Computername
    )

    Write-Verbose "Starting $($MyInvocation.MyCommand)"
    Write-Verbose "Getting last $Count events from $log event log on $computername"
    #base logname
    $base = "{0:yyyyMMdd}_{1}_{2}" -f (Get-Date), $Computername, $Log

    Try {
        $data = Get-EventLog -LogName $log -ComputerName $Computername -Newest $Count -ErrorAction Stop
    }
    Catch {
        Write-Warning "Failed to retrieve $log event log entries from $computername. $($_.Exception.Message)"
    }

    If ($data) {
        Write-Verbose "Exporting results to $($export.ToUpper())"
        Switch ($Export) {
            "csv" {
                $File = Join-Path -Path $Path -ChildPath "$base.csv"
                $data | Export-Csv -Path $File
            }
            "xml" {
                $File = Join-Path -Path $Path -ChildPath "$base.xml"
                ($data | ConvertTo-Xml).Save($File)
            }
            "clixml" {
                $File = Join-Path -Path $Path -ChildPath "$base.xml"
                $data | Export-Clixml -Path $File
            }
        } #switch

        Write-Verbose "Results exported to $File"

    } #if $data

    Write-Verbose "Ending $($MyInvocation.MyCommand)"

}