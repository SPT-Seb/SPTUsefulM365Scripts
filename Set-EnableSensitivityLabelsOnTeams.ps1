<#
.SYNOPSIS
    Enable Sensitivity Labels for Teams on your Office 365 tenant
.DESCRIPTION
    Azure AD Preview PowerShell Module required,
    source https://docs.microsoft.com/fr-fr/azure/active-directory/users-groups-roles/groups-assign-sensitivity-labels 
.NOTES
    File Name      : Set-EnableSensitivityLabelsOnTeams.ps1
    Author         : Sébastien Paulet (SPT Conseil)
    Prerequisite   : PowerShell V5.1 and AzureADPreview

.EXAMPLE
    .\Set-EnableSensitivityLabelsOnTeams.ps1
#>
Import-Module AzureADPreview
Connect-AzureAD
$Setting = Get-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ).id
$Setting.Values
$Setting["EnableMIPLabels"] = "True"
Set-AzureADDirectorySetting -Id $Setting.Id -DirectorySetting $Setting