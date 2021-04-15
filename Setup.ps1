#Region Prerequisites
<#
    - Need an Azure Account and subscription.
    - Need an Azure Key Vault created.
    - Install below modules.
#>
#EndRegion Prerequisites

#Region Module Setup

$ModuleSetupSplat = @{
    Name = @(
        'ImportExcel',
        'Microsoft.PowerShell.SecretManagement',
        'Microsoft.PowerShell.SecretStore',
        'Az.KeyVault'
    )
    Scope = 'CurrentUser'
    Force = $true
}
Install-Module @ModuleSetupSplat

#EndRegion Module Setup

#Region Register Local SecretStore

Import-Module Az.KeyVault

$AzVaultSplat = @{
    Module = 'Microsoft.PowerShell.SecretStore'
    Name   = 'AzVaultInfo'
}

Register-SecretVault @AzVaultSplat

#EndRegion Register Local SecretStore

#Region Configure Local SecretStore
# Stored using `Export-Clixml` from a Get-Credential
# Username is VaultName, Password is Vault SubscriptionID
$AzKVConfig   = Import-Clixml .\vault.config

$AzVaultNameSecretSplat = @{
    Name   = 'ExcelSecretsDemo.Name'
    Vault  = 'AzVaultInfo'
    Secret = $AzKVConfig.UserName
}
$AzVaultSubscriptionSplat = @{
    Name   = 'ExcelSecretsDemo.SubscriptionID'
    Vault  = 'AzVaultInfo'
    Secret = $AzKVConfig.GetNetworkCredential().Password
}

Set-Secret @AzVaultNameSecretSplat
Set-Secret @AzVaultSubscriptionSplat

#EndRegion Configure Local SecretStore

#Region AzKV Setup

Import-Module Az.KeyVault

$AzVaultSplat = @{
    Module = 'Az.KeyVault'
    Name   = 'AzExcelSecretsDemo'
    VaultParameters = @{
        AZKVaultName   = Get-Secret -Vault AzVaultInfo -Name ExcelSecretsDemo.Name
        SubscriptionID = Get-Secret -Vault AzVaultInfo -Name ExcelSecretsDemo.SubscriptionID
    }
}

Register-SecretVault @AzVaultSplat

#EndRegion AzKV Setup

#Region Configure AzKV
Connect-AzAccount

$AzVaultExcelSecretSplat = @{
    Name   = 'ProcessInfo'
    Vault  = 'AzExcelSecretsDemo'
    Secret = Read-Host -AsSecureString
}

Set-Secret @AzVaultExcelSecretSplat

#EndRegion Configure AzKV

#Region Create Excel Sheet with Protected Data

$ExcelProcessInfoSecretSplat = @{
    Name        = 'ProcessInfo'
    Vault       = 'AzExcelSecretsDemo'
    AsPlainText = $true
}
$Password = Get-Secret @ExcelProcessInfoSecretSplat

Get-Process | Export-Excel -Path 'C:\Temp\ProcessInfo.xlsx' -Password $Password

#EndRegion Create Excel Sheet with Protected Data

#Region Import from protected Excel file

$ExcelProcessInfoSecretSplat = @{
    Name        = 'ProcessInfo'
    Vault       = 'AzExcelSecretsDemo'
    AsPlainText = $true
}
$Password = Get-Secret @ExcelProcessInfoSecretSplat

Import-Excel 'C:\Temp\ProcessInfo.xlsx' -Password $Password

#EndRegion Import from protected Excel file