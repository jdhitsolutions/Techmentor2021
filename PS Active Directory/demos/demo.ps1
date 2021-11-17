#requires -version 5.1
#requires -module ActiveDirectory

return "This is a walk-through demo file"

#run this from a Windows domain member

Clear-Host

#region Basics

#add RSAT Active Directory
# Add-WindowsCapability -name rsat.ActiveDirectory* -online
Get-WindowsCapability -Name rsat.ActiveDirectory* -Online

Import-Module ActiveDirectory
Get-Command -Module ActiveDirectory

#READ THE HELP!!!

#ask for what you want
Get-ADUser Aprils
Get-ADUser Aprils -Properties Title, Department, Description

#discover
#can only use a wildcard like this:
Get-ADUser Aprils -Properties *

Get-ADUser -Filter * -SearchBase "OU=Employees,DC=company,DC=pri" -Properties Department, City -ov a |
Group-Object -property Department

#SF Benefits department moving to Oakland and changing name
$splat = @{
   Name        = "Oakland"
   Description = "Oakland Benefits"
   ManagedBy   = (Get-ADUser artd)
   Path        = "OU=Employees,DC=Company,DC=pri"
   Passthru    = $True
}

New-ADOrganizationalUnit @splat

#get users
$users = Get-ADUser -Filter "Department -eq 'benefits' -AND City -eq 'San Francisco'" -Properties City, Department, Company
$users | Select-Object -property Distinguishedname, Name, City, Department, Company

$users | Move-ADObject -TargetPath "OU=Oakland,OU=Employees,DC=Company,DC=pri" -PassThru |
Set-ADUser -City Oakland -Department "Associates Assistance" -Company "Associated Benefits"

Get-ADUser -Filter * -SearchBase "OU=Oakland,OU=Employees,DC=Company,DC=pri" -Properties City, Department, Company |
Select-Object -property DistinguishedName, Name, Department, City, Company

<#
reset demo

Get-ADuser -filter * -SearchBase "OU=Oakland,OU=Employees,DC=Company,DC=pri" |
Move-ADObject -TargetPath "OU=Accounting,OU=Employees,DC=Company,DC=pri" -PassThru |
Set-ADuser -City 'San Francisco' -Department "Benefits" -Company "Company.com"
Get-ADOrganizationalUnit -filter "Name -eq 'oakland'" |
Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $False -PassThru |
Remove-ADObject -confirm:$False

#>


#endregion

#region FSMO

Get-ADDomain
Get-ADDomain | Select-Object *master, PDC*
Get-ADForest

Function Get-FSMOHolder {
   [cmdletbinding()]
   Param(
      [Parameter(Position = 0, HelpMessage = "Specify a FSMO role or select ALL to display all roles.")]
      [ValidateSet("All", "PDCEmulator", "RIDMaster", "InfrastructureMaster", "SchemaMaster", "DomainNamingMaster")]
      [string[]]$Role = "All",
      [Parameter(HelpMessage = "Specify the distinguished name of a domain")]
      [ValidateNotNullOrEmpty()]
      [string]$Domain = (Get-ADDomain).DistinguishedName
   )

   Try {
      $ADDomain = Get-ADDomain -Identity $Domain -ErrorAction Stop
      $ADForest = $ADDomain | Get-ADForest -ErrorAction Stop
   }
   Catch {
      Throw $_
   }

   if ($ADDomain -AND $ADForest) {
      $fsmo = [PSCustomObject]@{
         Domain               = $ADDomain.Name
         Forest               = $ADForest.Name
         PDCEmulator          = $ADDomain.PDCEmulator
         RIDMaster            = $ADDomain.RIDMaster
         InfrastructureMaster = $ADdomain.InfrastructureMaster
         SchemaMaster         = $ADForest.SchemaMaster
         DomainNamingMaster   = $ADForest.DomainNamingMaster
      }
      if ($Role -eq "All") {
         $fsmo
      }
      else {
         $fsmo | Select-Object -Property $Role
      }
   }
} #end Get-FSMOHolders

Get-FSMOHolder
Get-FSMOHolder -role PDCEmulator,DomainNamingMaster

#endregion

#region Empty OU

#use the AD PSDrive
Get-PSDrive AD
Get-ChildItem 'AD:\DC=Company,DC=Pri'

Get-ADOrganizationalUnit -Filter * | ForEach-Object {
   $ouPath = Join-Path -Path "AD:\" -ChildPath $_.distinguishedName
   #test if the OU has any children other than OUs
   $test = Get-ChildItem -Path $ouPath -Recurse |
   Where-Object ObjectClass -NE 'organizationalunit'
   if (-Not $Test) {
      $_.distinguishedname
   }
}

#You could then decide to remove them,
#but beware of protection from accidental deletion
Set-ADOrganizationalUnit -Identity "OU=Y2kResources,DC=Company,DC=pri" -ProtectedFromAccidentalDeletion $False -PassThru |
Remove-ADObject -WhatIf

#endregion

#region Create new users

#parameters to splat to New-ADUser
$params = @{
   Name              = "Thomas Anderson"
   DisplayName       = "Thomas Anderson"
   SamAccountName    = "tanderson"
   UserPrincipalName = "tanderson@company.com"
   PassThru          = $True
   GivenName         = "Tom"
   Surname           = "Anderson"
   Description       = "the one"
   Title             = "Senior Web Developer"
   Department        = "IT"
   AccountPassword   = (ConvertTo-SecureString -String "P@ssw0rd" -Force -AsPlainText)
   Path              = "OU=IT,OU=Employees,DC=Company,DC=Pri"
   Enabled           = $True
}

#test if user account already exists
Function Test-ADUser {
   [cmdletbinding()]
   [Outputtype("boolean")]
   Param(
      [Parameter(Position = 0, Mandatory,HelpMessage = "Enter a user's SamAccountName")]
      [ValidateNotNullOrEmpty()]
      [string]$Identity,
      [string]$Server,
      [PSCredential]$Credential
   )
   Try {
      [void](Get-ADUser @PSBoundParameters -ErrorAction Stop)
      $True
   }
   Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
      $False
   }
   Catch {
      Throw $_
   }
}

If (-Not (Test-ADUser $params.samaccountname)) {
   Write-Host "Creating new user $($params.name)" -foreground cyan
   #splat the hashtable
   New-ADUser @params
}

# remove-aduser -Identity tanderson

#import

#the column headings match parameter New-ADUser parameter names
Import-Csv .\100NewUsers.csv | Select-Object -First 1

$secure = ConvertTo-SecureString -String "P@ssw0rdXyZ" -AsPlainText -Force

#I'm not taking error handling for duplicate names into account
$newParams = @{
   ChangePasswordAtLogon = $True
   Path                  = "OU=Imported,OU=Employees,DC=company,DC=pri"
   AccountPassword       = $secure
   Enabled               = $True
   PassThru              = $True
}

Import-Csv .\100NewUsers.csv | New-ADUser @newParams #-WhatIf

<#
Get-Aduser -Filter * -SearchBase $newParams.path |
Remove-ADUser -confirm:$false
#>

#endregion

#region Find inactive user accounts

#this demo is only getting the first 25 accounts
$paramHash = @{
   AccountInactive = $True
   Timespan        = (New-TimeSpan -Days 120)
   SearchBase      = "OU=Employees,DC=company,DC=pri"
   UsersOnly       = $True
   ResultSetSize   = "25"
}

Search-ADAccount @paramHash | 
Select-Object Name, LastLogonDate, SamAccountName, DistinguishedName |
Out-Gridview

#endregion

#region Find inactive computer accounts

#definitely look at help for this command

Search-ADAccount -ComputersOnly -AccountInactive

#endregion

#region Find empty groups

#can't use -match in the filter
$paramHash = @{
   filter     = "Members -notlike '*'"
   Properties = "Members", "Created", "Modified", "ManagedBy"
   SearchBase = "DC=company,DC=pri"
}

Get-ADGroup @paramHash |
Select-Object Name, Description,
@{Name = "Location"; Expression = { $_.DistinguishedName.split(",", 2)[1] } },
Group*, Modified, ManagedBy |
Sort-Object Location |
Format-Table -GroupBy Location -Property Name, Description, Group*, Modified, ManagedBy

#filter out User and Builtin
#can't seem to filter on DistinguishedName
$paramHash = @{
   filter     = "Members -notlike '*'"
   Properties = "Members", "Modified", "ManagedBy"
   SearchBase = "DC=company,DC=pri"
}

#formatting to make this nice to read
Get-ADGroup @paramhash |
Where-Object { $_.DistinguishedName -notmatch "CN=(Users)|(BuiltIn)" } |
Sort-Object -Property GroupCategory |
Format-Table -GroupBy GroupCategory -Property DistinguishedName, Name, Modified, ManagedBy

<#
This is the opposite. These are groups with any type of member.
The example is including builtin and default groups.
#>
$data = Get-ADGroup -Filter * -Properties Members, Created, Modified |
Select-Object Name, Description,
@{Name = "Location"; Expression = { $_.DistinguishedName.split(",", 2)[1] } },
Group*, Created, Modified,
@{Name = "MemberCount"; Expression = { $_.Members.count } } |
Sort-Object MemberCount -Descending

#sample
$data[0]

#I renamed properties from Group-Object to make the result easier to understand
$data | Group-Object MemberCount -NoElement |
Select-Object -Property @{Name = "TotalNumberOfGroups"; Expression = { $_.count } },
@{Name = "TotalNumberofGroupMembers"; Expression = { $_.Name } }

<#
TotalNumberOfGroups TotalNumberofGroupMembers
------------------- -------------------------
                  1 8
                  1 6
                  1 5
                  2 4
                  6 3
                  3 2
                  9 1
                 40 0
#>

#endregion

#region Enumerate Nested Group Membership

#show nested groups
psedit .\Get-ADNested.ps1

. .\Get-ADNested.ps1

$group = "Finance-and-Accounting"
Get-ADNested $group | Select-Object Name, Level, ParentGroup, 
@{Name = "Top"; Expression = { $group } }

#list all group members recursively
Get-ADGroupMember -Identity $group -Recursive | 
Select-Object Distinguishedname, samAccountName

#endregion

#region List User Group Memberships

$user = Get-ADUser -Identity "aprils" -Properties *

#this only shows direct membership
$user.MemberOf

psedit .\Get-ADMemberOf.ps1

. .\Get-ADMemberOf.ps1

$user | Get-ADMemberOf -verbose -ov m | 
Select-Object Name, DistinguishedName, GroupCategory -Unique |
Out-GridView

#endregion

#region Password Age Report

#get maximum password age.
#This doesn't take fine tuned password policies into account
$maxDays = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days

#parameters for Get-ADUser
#get enabled accounts with passwords that can expire
$params = @{
   filter     = "Enabled -eq 'true' -AND PasswordNeverExpires -eq 'false'"
   Properties = "PasswordLastSet", "PasswordNeverExpires"
}

#skip user accounts under CN=Users and those with unexpired passwords
Get-ADUser @params |
Where-Object { (-Not $_.PasswordExpired) -and ($_.DistinguishedName -notmatch "CN\=Users") } |
Select-Object DistinguishedName, Name, PasswordLastSet, PasswordNeverExpires,
@{Name = "PasswordAge"; Expression = { (Get-Date) - $_.PasswordLastSet } },
@{Name = "PassExpires"; Expression = { $_.passwordLastSet.addDays($maxDays) } } |
Sort-Object PasswordAge -Descending | Select-Object -First 10

#create an html report
psedit .\PasswordReport.ps1
. .\PasswordReport.ps1

Invoke-Item .\PasswordReport.html

#get an OU
$params = @{
   SearchBase  = "OU=Employees,DC=company,dc=pri"
   FilePath    = ".\employees.html"
   ReportTitle = "Staff Password Report"
   Server      = "DOM1"
   Verbose     = $True
}

.\PasswordReport.ps1 @params | Invoke-Item

#endregion

#region Domain Controller Health

Clear-Host

$dcs = (Get-ADDomain).ReplicaDirectoryServers

#services
#my domain controllers also run DNS
# the legacy way
# Get-Service adws,dns,ntds,kdc -ComputerName $dcs | Select-Object Machinename,Name,Status

$cim = @{
   ClassName    = "Win32_Service"
   filter       = "name='adws' or name='dns' or name='ntds' or name='kdc'"
   ComputerName = $dcs
}
Get-CimInstance @cim | Select-Object SystemName, Name, State

#eventlog
Get-EventLog -List -ComputerName DOM1

#remoting speeds this up
$data = Invoke-Command {
   #ignore errors if nothing is found
   Get-EventLog -LogName 'Active Directory Web Services' -EntryType Error, Warning -Newest 10 -ErrorAction SilentlyContinue
} -computer $dcs

<# demo alternative

$data = Invoke-Command {
Get-EventLog -LogName 'Active Directory Web Services' -Newest 10
} -computer $dcs

#>

#formatted in the console
$data | Sort-Object PSComputername, TimeGenerated -Descending |
Format-Table -GroupBy PSComputername -Property TimeGenerated, EventID, Message -Wrap

#how about a Pester-based health test?

psedit .\ADHealth.tests.ps1

Clear-Host

#make sure I'm using v4.10 of Pester. My test is not compatible with Pester 5.0.
# Install-module Pester -requiredversion 4.10.1 -force -SkipPublisherCheck

Get-Module Pester | Remove-Module
Import-Module Pester -RequiredVersion 4.10.1 -Force

Invoke-Pester .\ADHealth.tests.ps1

#You could automate running the test and taking action on failures

#endregion

#region ADReportingTools

#Looking for reporting?
# Install-Module ADReportingTools
# https://github.com/jdhitsolutions/adreportingtools

Import-Module ADReportingTools
Get-Command ADReportingTools

Get-ADReportingTools
Open-ADReportingToolsHelp

#run these in a PowerShell session to see ANSI
Show-DomainTree
Get-ADBranch "OU=IT,Ou=Employees,DC=Company,DC=pri"

#endregion