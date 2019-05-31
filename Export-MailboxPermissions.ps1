$delegateExport = Get-Mailbox -RecipientTypeDetails SharedMailbox | Where-Object { ( $_.Identity -like '*team*' ) -or ( $_.Identity -like '*mail*' ) } | Sort-Object -Property Name
foreach ($delegate in $delegateExport) {
    Get-MailboxPermission $delegate.Identity |
    Select-Object Identity,User,AccessRights |
    Export-Csv -Delimiter ',' -NoTypeInformation -Path C:\Clients\BizSpace\SharedMailboxPermissions.csv -Append
}
