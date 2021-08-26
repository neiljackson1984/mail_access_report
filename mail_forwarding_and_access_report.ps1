$pathOfOutputReportFile = "mail_forwarding_and_access_report.txt"

$allMailboxPermissions = get-mailbox | get-mailboxpermission 
$allRecipientPermissions = get-mailbox | get-RecipientPermission 
$allAzureAdUsers = Get-AzureADUser
$licensedAzureAdUsers = $allAzureAdUsers | where {$_.AssignedLicenses.Length -gt 0}


"" | Out-File -FilePath $pathOfOutputReportFile
foreach($mailbox in (get-mailbox | Sort-Object -Property PrimarySmtpAddress)){
    $azureAdUserWhoOwnsThisMailbox = $allAzureAdUsers | where {$_.UserPrincipalName -eq $mailbox.PrimarySmtpAddress}
    
    $fullAccessPermissionsToThisMailbox = (
        $allMailboxPermissions | where {
            ($_.Identity -eq $mailbox.Identity) -and 
            ($_.AccessRights.contains( "FullAccess" )) -and
            (! $_.Deny) -and
            # (($allAzureAdUsers.UserPrincipalName).contains($_.User))
            (($licensedAzureAdUsers.UserPrincipalName).contains($_.User))
        }
    )
    
    $recipientPermissionsToThisMailbox = (
        $allRecipientPermissions | where {
            ($_.Identity -eq $mailbox.Identity) -and 
            ($_.AccessRights.contains( "SendAs" )) -and
            ($_.AccessControlType -eq "Allow") -and 
            (($licensedAzureAdUsers.UserPrincipalName).contains($_.Trustee))
        }
    )
    
    
    $azureAdUsersThatHaveFullAccessToThisMailbox = (
        ([system.Array] ($fullAccessPermissionsToThisMailbox | foreach {Get-AzureADUser -ObjectId $_.User} )) +
        ([system.Array] @($azureAdUserWhoOwnsThisMailbox))
    ) | Sort-Object -Property UserPrincipalName
    
    $azureAdUsersThatHaveSendAsPermissionToThisMailbox = (
        ([system.Array] ($recipientPermissionsToThisMailbox | foreach {Get-AzureADUser -ObjectId $_.Trustee} )) +
        ([system.Array] @($azureAdUserWhoOwnsThisMailbox))
    ) | Sort-Object -Property UserPrincipalName
    
    
    $addressesToWhichThisMailboxIsBeingRedirected = (Get-InboxRule -Mailbox $mailbox.Identity).RedirectTo | foreach-object {$_} | sort
    $reportMessage = ""
    $reportMessage += "The mailbox " + $mailbox.PrimarySmtpAddress + " (which is a " + $(if($mailbox.IsShared){"shared"} else {"non-shared"}) +  " mailbox) " 
    
        
    $reportMessage += "`n" + 
    "`t" + "is accessible (full access permission) " + 
    $(
        if($azureAdUsersThatHaveFullAccessToThisMailbox.Length -eq 1){
            "only to " + $azureAdUsersThatHaveFullAccessToThisMailbox[0].UserPrincipalName
        } elseif ($azureAdUsersThatHaveFullAccessToThisMailbox.Length -gt 1){
            "to the following users: " + "`n" +
            [system.String]::Join("`n", ($azureAdUsersThatHaveFullAccessToThisMailbox | sort | foreach {"`t`t" + $_.UserPrincipalName}))
        } else {
            "to nobody."
        } 
    ) + "`n" + 
    "`t" + "and is sendable (send-as permission) " + 
    $(
        if($azureAdUsersThatHaveSendAsPermissionToThisMailbox.Length -eq 1){
            "only to " + $azureAdUsersThatHaveSendAsPermissionToThisMailbox[0].UserPrincipalName
        } elseif ($azureAdUsersThatHaveSendAsPermissionToThisMailbox.Length -gt 1){
            "to the following users: " + "`n" +
            [system.String]::Join("`n", ($azureAdUsersThatHaveSendAsPermissionToThisMailbox | sort | foreach {"`t`t" + $_.UserPrincipalName}))
        } else {
            "to nobody."
        } 
    ) + "`n" + 
    "`t" + "and is being redirected " + 
    $(
        if($addressesToWhichThisMailboxIsBeingRedirected.Length -gt 0){
            "to the following addresses: " + "`n" +
            [system.String]::Join("`n", ($addressesToWhichThisMailboxIsBeingRedirected | foreach {"`t`t" + $_})) 
        } else {
            "nowhere."         
        }
    ) + "`n"
        
    write-output($reportMessage)
    $reportMessage | Out-File -Append -FilePath $pathOfOutputReportFile
    
    
    #pause to avoid hitting the Office365 call quota.
    # Start-Sleep -Milliseconds 4000
        
}



