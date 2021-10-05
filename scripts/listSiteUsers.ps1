<#
.Synopsis
   List the users with access to the SharePoint site.  Filters can be applied

.Description
   Packages source folder contents into a 7ZIP file, adding a reconciliation 
   file to the 7ZIP file and then encrypting the contents.  Send 
   * this script
   * the 7ZIP package file 
   * plus optional SecretFile ( if using RecipientKey )
   to the target or recipient.
   
.Parameter Tenant
  The SharePoint tenant name.  This is the URL name before ".sharepoint.com"  
   
.Parameter Site
  The SharePoint site name.  This is the names after sites in the URL "https://XXXX.sharepoint.com/sites/"  
   
.Parameter HostDomain
  The host domain name.  This is generally your company domain as follows the @ in your email.
  This is used as part of the filter to identify users, either internal or external.

.Notes
  This script is wriotten for audting of SharePoint site.
 
.Example
   # List users for tenancy "foxshares" and site "FoxLibrary" that have word "example", such as peter@example.org
   .\listSiteUsers.ps1 -Tenant foxshares -Site FoxLibrary -HostDomain foxtale.com -UserFilter example -AdminUser admin@foxtale.com
   
 
 
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String] $Tenant, 

    [Parameter(Mandatory)]
    [String] $Site, 

    [Parameter(Mandatory)]
    [String] $HostDomain, 

    [Alias("Filter")]
    [String] $UserFilter,

    [String] $OutFile,

    [Alias("UserName")]
    [String] $AdminUser

)

Import-Module Microsoft.Online.SharePoint.PowerShell -Scope Local 

If ($PSBoundParameters['Debug']) {
    Write-Host "Function 'listSiteUsers' parameters follow" -ForegroundColor Yellow
    Write-Host "Parameter: Tenant   Value: $Tenant " -ForegroundColor Yellow
    Write-Host "Parameter: Site   Value: $Site " -ForegroundColor Yellow
    Write-Host "Parameter: HostDomain   Value: $HostDomain " -ForegroundColor Yellow
    Write-Host "Parameter: UserFilter   Value: $UserFilter " -ForegroundColor Yellow
    Write-Host "Parameter: OutFile   Value: $OutFile " -ForegroundColor Yellow
    Write-Host "Parameter: AdminUser   Value: $AdminUser " -ForegroundColor Yellow
}

if ($OutFile -eq "") {
    $csvfile = ".\" + $Site + "_Users.csv"
} else {
    $csvfile = $OutFile
}


$AdminSiteURL="https://"+ $Tenant +"-admin.sharepoint.com"
$SiteURL = "https://"+$Tenant + ".sharepoint.com/sites/" + $Site
$pattern = "(^.*"+ $UserFilter + ".*@"+$hostDomain + "$)|(^.*@"+ $UserFilter + ".*$)|(^.*"+ $UserFilter + ".*@"+ $Tenant + ".onmicrosoft.com$)"

Write-Host "Connecting with: $SiteURL "


if ($AdminUser -eq "") {
    $Cred = Get-Credential -Message "SharePoint admin credentials"
} else {
    $Cred = Get-Credential -Message "SharePoint admin credentials" -UserName $AdminUser
}

Connect-SPOService -URL $AdminSiteURL -Credential $Cred


Get-SPOUser -Site $SiteUrl | Where-Object {$_.LoginName -match $pattern -and $_.LoginName -ne "app@sharepoint"} | Select DisplayName,LoginName,KeepAccess | Export-Csv -Force -NoTypeInformation $csvfile
