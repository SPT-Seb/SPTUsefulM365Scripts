<#
.SYNOPSIS
    Sets or removes an Azure Information Protection label for a file, and sets or removes the protection according to the label configuration or custom permissions
.DESCRIPTION
    Warngin : it will update modification date
.NOTES
    File Name      : Set-EnhancedAIPFileLabel.ps1
    Author         : Sébastien Paulet (SPT Conseil)
    Prerequisite   : PowerShell V5.1 and AzureInformationProtection PowerShell Module 
                    (install AzInfoProtection_UL.exe from https://www.microsoft.com/en-us/download/details.aspx?id=53018)
    See also : https://docs.microsoft.com/en-us/powershell/module/azureinformationprotection/set-aipfilelabel?view=azureipps
    Tested on .jpg, .pdf, .docx, .xlsx, .pptx, .zip file formats
    .PARAMETER Path
Specifies a local path, network path, or SharePoint Server URL to the files for which you want to get the label and protection information.
Wildcards are not supported and WebDav locations are not supported.
or SharePoint paths: SharePoint Server 2019, SharePoint Server 2016, and SharePoint Server 2013 are supported.
.PARAMETER LabelId
Specifies the identity (ID) of the label to apply. When a label has sublabels, always specify the ID of just a sublabel and not the parent label.
.PARAMETER CopyLabelFilePath
If you don't know your label ID, apply the label to an existing file and provide here the full path to this file. Its label will be replicated to all files in Path
.PARAMETER ClearLabels
If used, clear labels on all files in Path

.EXAMPLE
    .\Set-EnhancedAIPFileLabel.ps1 -Path "C:\Users\MyUser\MyTenant\MyDocLib\CONFIDENTIEL\" -LabelId "3dbc6dc2-f497-4925-999c-5f959606999d"
    .\Set-EnhancedAIPFileLabel.ps1 -Path "C:\Users\MyUser\MyTenant\MyDocLib\CONFIDENTIEL\" -CopyLabelFilePath "C:\Users\MyUser\MyTenant\MyDocLib\CONFIDENTIEL\AlreadyLabeledFile.pdf"
    .\Set-EnhancedAIPFileLabel.ps1 -Path "C:\Users\MyUser\MyTenant\MyDocLib\CONFIDENTIEL\" -ClearLabels
#>
param(
    [Parameter(Mandatory = $true)]
    [String]$Path,
    [guid]$LabelId,
    [String]$CopyLabelFilePath,
    [switch]$ClearLabels
)
$isConnected = $false
try {Get-AzureADTenantDetail | Out-Null; $isConnected = $true} catch {} 
if (-not $isConnected) {
    Write-Verbose 'Connecting Azure AD'
    Connect-AzureAD
    #pause to ensure bootstrap 
    Start-Sleep -Seconds 5
}
if ($ClearLabels) {
    Write-Verbose 'Removing RMS protection on file(s)'
    Set-AIPFileLabel -Path $Path  -RemoveProtection -PreserveFileDetails
    Write-Verbose 'Removing Labels on file(s)'
    #removing both protection and label with the same command line doesn't work (error message). Splitting in two commands looks OK
    Set-AIPFileLabel -Path $Path  -RemoveLabel -JustificationMessage 'Label removed by Set-EnhancedAIPFileLabel.ps1 script' -PreserveFileDetails
} else { 
    #if you need your Label ID directly, use 
    #Connect-IPPSSession; Get-Label | Format-Table -Property DisplayName, Name, Guid 
    #command from Exchange Online powershell module
    if ($CopyLabelFilePath) {
        $LabelId =  (Get-AIPFileStatus -Path $CopyLabelFilePath).MainLabelId 
        Write-Verbose "Label identified on source file : $LabelId"
    }
    Write-Verbose "Applying label $LabelId on file(s)"
    Set-AIPFileLabel -Path $Path -LabelId $LabelId -PreserveFileDetails
}
if (-not $isConnected) {
    Disconnect-AzureAD
}