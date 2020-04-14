<#
.SYNOPSIS
    Get a simple list containing all email address of all members of a Microsoft Teams team. Even guest ones
.DESCRIPTION
    Exchange Online PowerShell Module required,
    source https://github.com/spgoodman/p365scripts/blob/master/Export-O365Groups.ps1
.NOTES
    File Name      : Get-AllTeamsMembersEmails.ps1
    Author         : Sébastien Paulet (SPT Conseil)
    Prerequisite   : PowerShell V5.1 and Exchange Online PowerShell Module 
.PARAMETER GroupName
    Teams name
.EXAMPLE
    .\Get-AllTeamsMembersEmails.ps1 -GroupName "[Public]aOS Solidarité Calédonie 2020"
#>
param(
 	[Parameter(Mandatory = $true)]
	[String]$GroupName = "[Public]aOS Solidarité Calédonie 2020"
)
Connect-EXOPSSession
$Group = Get-UnifiedGroup -ResultSize Unlimited -Identity $GroupName
$Members = Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType Members -ResultSize Unlimited
$Members | select PrimarySmtpAddress