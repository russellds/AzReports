function CheckAzContext {
    if (-not (Get-AzContext).Account) {
        Connect-AzAccount
    }
}
