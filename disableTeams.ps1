<#
Name: disableTeams.ps1
Description: Set a license option to disable teams for users
Author: Austin Vargason
Date ModifieD: 8/15/2018
#>


#connect to MSOL service
Connect-MsolService


#function to disable teams for either a user
function Disable-UserTeams () {
    param (
        [Parameter(Mandatory=$true)]
        [String]$UserPrincipalName
    )

    #Account sku for enteprise pack
    $sku = "sdcountycagov:ENTERPRISEPACK_GOV"

    #service plan to disable
    $serv = "TEAMS_GOV"

    #set the license option
    $licenseOption = New-MsolLicenseOptions -AccountSkuId $sku -DisabledPlans $serv

    #set the license option for the user
    try {
        Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -LicenseOptions $licenseOption -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Could not find MSOL user: $UserPrincipalName"
    }

}

#write the process to the screen
Write-Host "getting user list" -ForegroundColor Cyan

#get the user list to use from O365
$userList = Get-MsolUser -All | where {$_.isLicensed -eq $true}

#start a counter
$i = 0

#get the total count to use
$count = $userList.Count

#foreach user in the list, disable teams
foreach($user in $userList){

    #save the userprincipal name
    $upn = $user.UserPrincipalName

    #if the user has an office 365 gov license, run disable teams
    if (($test | Select -ExpandProperty Licenses | Select -ExpandProperty AccountSkuId) -contains "sdcountycagov:ENTERPRISEPACK_GOV") {
       Disable-UserTeams -UserPrincipalName $upn
    }
    else {
        Write-Host "INFO: User does not have Government License" -ForegroundColor Yellow
    }

    #increase the counter
    $i++

    #write the progress so you can see how things are going
    Write-Progress -Activity "Disabling Teams" -Status "Disabled teams for User: $upn" -PercentComplete (($i / $count) * 100)
}

