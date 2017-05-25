import-module *azure*
$cred = get-credential
Login-AzureRmAccount -Credential $cred -SubscriptionName 'LABDEVELOPER'
Start-AzureRmVM -Name jump