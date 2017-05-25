import-module *azure*
Login-AzureRmAccount
Select-AzureRmSubscription -SubscriptionName "Microsoft Partner Network"
Start-AzureRmVM -Name jump -ResourceGroupName LABDEVELOPER
start-sleep -Seconds 20
mstsc /v:"jump.ukwest.cloudapp.azure.com"