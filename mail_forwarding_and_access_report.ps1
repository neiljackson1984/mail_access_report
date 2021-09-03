$reportTime = (Get-Date )
$defaultDomainName = (Get-AzureAdDomain | where-object {$_.IsDefault}).Name

$pathOfOutputReportFile = "mail_forwarding_and_access_report_$defaultDomainName_$('{0:yyyy-MM-dd_HH-mm}' -f $reportTime).txt"

$allMailboxPermissions = get-mailbox | get-mailboxpermission 
$allRecipientPermissions = get-mailbox | get-RecipientPermission 
$allAzureAdUsers = Get-AzureADUser
$licensedAzureAdUsers = $allAzureAdUsers | where {$_.AssignedLicenses.Length -gt 0}


"" | Out-File -FilePath $pathOfOutputReportFile
$reportTime = (Get-Date )


(
@"
EMAIL FORWARDING AND ACCESS REPORT
$defaultDomainName
Prepared $('{0:yyyy/MM/dd HH:mm}' -f $reportTime)

"@
) | Out-File -Append -FilePath $pathOfOutputReportFile


function convertRedirectEntryToAddress($redirectEntry){
    # redirectEntry is expected to be a string like the members
    # of an InboxRule's RedirectTo property.
    $pattern='^"([^"]*)"\s*\[([^:]*):(.*)\]$'
    #when we apply $pattern to redirectEntry, the matching groups will pull 
    # out the part between quotes, the protocol name, and the address, respectively.
    
    
    #example $redirectEntry=='"John Doe" [SMTP:jdoe@acme.com]'
    #the matching groups will be:
    #$1: John Doe
    #$2: SMTP
    #$3: jdoe@acme.com    
    
    
    #example $redirectEntry=='"John Doe" [EX:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=39e06be27d4a4e3e813d7ea40b95fa3f-jdoe]'

    #the matching groups will be:
    #$1: John Doe
    #$2: EX
    #$3: /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=39e06be27d4a4e3e813d7ea40b95fa3f-jdoe
    
    #in either case, we can attempt to "convert the address to SMTP by attempting to do (Get-Recipient $3).primarySMTPAddress
    # if this throws a "couldn't find" error, then we return $3 as is.
    $matches = $null
    $result = $redirectEntry -match $pattern
    
    # if(!$result){
        # #this is unexpected, and would constitute an error.
        # # return "NOMATCH " + ([String] $redirectEntry)
        # return $redirectEntry
    # }
       

    if($result){
        $resolvedSMTPAddress = (Get-Recipient -Identity $matches[3]).PrimarySmtpAddress  2>$null
        if($resolvedSMTPAddress){
            $resolvedSMTPAddress
        } elseif($matches[3]) {
            $matches[3]
        } else {
            $redirectEntry
        }
    } else {
        $redirectEntry
    }
        

}



foreach($mailbox in (get-mailbox | Sort-Object -Property PrimarySmtpAddress)){
# foreach($mailbox in (get-mailbox | Sort-Object -Property PrimarySmtpAddress | Where-Object { -not $_.Identity.Contains("-snapshot20210502")  })){
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
    
    $addressesToWhichThisMailboxIsBeingRedirected = (
        @( 
            ( Get-InboxRule -Mailbox $mailbox.Identity 
            ) | foreach-object{ 
                $_.RedirectTo; 
                $_.ForwardTo; 
                $_.ForwardAsAttachmentTo;  
            } |  Where-Object {
                $_
                # we need the "|  Where-Object {$_}" in order to remove nulls from the pipeline, which happens when RedirectTo is, essentially, an empty list.
            } | foreach-object { convertRedirectEntryToAddress $_ }
        ) + @(
            Get-TransportRule | where-object {
                $_.State -eq "Enabled"
            } | where-object {
                $sentToList = $_.SentTo;
                @(
                    &{ $sentToList | foreach-object {(get-mailbox -identity $_).Identity} | where-object {$_} }
                ).Contains($mailbox.Identity  )
            } | foreach-object {
                $_.CopyTo;
                $_.BlindCopyTo;
                $_.RedirectMessageTo;
            }
        )
        # This is probably not a comprehensive way to detect all possible forwarding due to mail flow rules,
        # but it serves the immediate purpose.
    )
    
    
    #TODO: we should also look at the mailbox's ForwardingAddress and ForwardingSMTPAddress properties as possible sources of addresses to which this mailbox is being redirected.
    
    
    $addressesToWhichThisMailboxIsBeingRedirected = $addressesToWhichThisMailboxIsBeingRedirected | sort
    
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



