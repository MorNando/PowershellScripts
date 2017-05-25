import-module *azure*
$cred = get-credential
start-azurerm
Start-AzureRmVM -Name jump