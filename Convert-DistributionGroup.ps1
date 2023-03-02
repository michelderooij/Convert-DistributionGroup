<#
    .SYNOPSIS
    Convert-DistributionGroup

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 1.01, March 2nd, 2023

    .DESCRIPTION
    Script to support converting synchronized Distribution Groups to cloud-only Distribution Groups

    The script has three operating modes, supporting the phases of the conversion:
 
    1) CLONE
       In Exchange Online, create clone objects from the existing (sychronized) Distribution Groups, including Recipient Permissions
       The following properties will be prefixed (default Clone) to prevent conflicts: Name, DisplayName, Alias, PrimarySmtpAddress
       The property specified by SyncAttribute (default CustomAttribute15) will be cleared, as its designated to put things in sync scope
       The property specified by legacyExchangeAttribute is used to store the current legacyExchangeDN
       The property specified by EmailAddressesAttribute is used to store the current EmailAddresses

    2) CONTACT
       In Exchange On-premises, convert the Distribution Groups to Contacts for addresslist representation etc.
       The property specified by SyncAttribute (default CustomAttribute15) will be cleared to put contact out of sync scope

    3) CONVERT
       In Exchange Online, remove the Distribution Group after checking if the clone is present.
       The following properties will be stripped from the prefix: Name, DisplayName, Alias, PrimarySmtpAddress
       The property specified by EmailAddressesAttribute is used to stamp the original EmailAddresses
       The property specified by legacyExchangeAttribute is used to add an X500 proxy to EmailAddresses for resolving old communications
    
    The script also has a RESTORE mode for restoring accidentally converted Distribution Groups, using the information saved in the backup XML files.

    .LINK
    http://eightwone.com

    .NOTES
    Requires Exchange Management Shell access using sufficient permissions in either Exchange Online or Exchange on-premises, depending on the phase of the conversion.
    To do: In addition to using an extension attribute for AD Connect synchronization scoping, add parameter to relocate objects to another Organizational Unit.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release

    .PARAMETER Identity
    Name of the Distribution Group to convert.

    .PARAMETER Mode
    Instructs the script which operating mode to use:
    - Clone   : In Exchange Online, create a clone of an existing synchronized Distribution Group
    - Contact : In Exchange On-Premises, convert a Distribution Group to a Mail-Enabled Contact
    - Convert : In Exchange Online, modify the clone and remove prefixes temporary used for name, primarySmtpAddress etc.
    - Restore : In Exchange On-Premises, restore a Distribution Group using the information stored during conversion to contact 

    .PARAMETER ExportFolder
    Specifies the location of the XML files used for saving/restoring during Contact conversion or Restore of Distribution Group
    Default is the current folder.

    .PARAMETER Prefix
    Specifies the prefix to use when creating a clone of a Distribution Group. Default is 'Clone-'.

    .PARAMETER SyncAttribute
    Specifies the attribute which holds a value when included in AD Connect synchronization, and which needs to be cleared to move
    objects out of synchronization scope. Default is CustomAttribute15.

    .PARAMETER legacyExchangeAttribute
    Specifies which extension attribute can be used for temporary preserving the current value of legacyExchangeDN.
    Defaults to extensionCustomAttribute4.

    .PARAMETER EmailAddressesAttribute
    Specifies which extension attribute can be used for temporary preserving the current value of EmailAddresses.
    Defaults to extensionCustomAttribute5.

    .EXAMPLE
    .\Convert-DistributionGroup.ps1 -Identity ServiceDesk -Mode Clone  
    In Exchange Online, create a clone of the distribution group ServiceDesk with all its properties, using prefix Clone- for attributes which cannot be
    set yet due to conflicts with the original distribution group.

    .\Convert-DistributionGroup.ps1 -Identity ServiceDesk -Mode Contact
    In Exchange on-premises, convert the distribution group ServiceDesk to a mail-enabled contact with all possible properties intact.

    .\Convert-DistributionGroup.ps1 -Identity ServiceDesk -Mode Convert
    In Exchange Online, modify the clone of the distribution group ServiceDesk removing any prefixed attributes. This needs to take place after
    AD Connect synchronization ran, removing the original distribution group since it should be out of its scope now.

    .\Convert-DistributionGroup.ps1 -Identity ServiceDesk -Mode Restore
    In Exchange on-premises, restore the distribution group ServiceDesk with all its properties using the information stored during conversion to mail-enabled contact.

#>
[cmdletbinding(
    SupportsShouldProcess= $true,
    ConfirmImpact= 'High'
)]
param (
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true)] 
    [string]$Identity,
    [parameter( Mandatory= $false)] 
    [ValidateSet( 'Clone', 'Convert', 'Contact','Restore')]
    [string]$Mode='Clone',
    [parameter( Mandatory= $false)] 
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$ExportFolder='.',
    [parameter( Mandatory= $false)] 
    [string]$Prefix='Clone-',
    [parameter( Mandatory= $false)] 
    [string]$SyncAttribute='CustomAttribute15',
    [parameter( Mandatory= $false)] 
    [string]$legacyExchangeAttribute='ExtensionCustomAttribute4',
    [parameter( Mandatory= $false)] 
    [string]$EmailAddressesAttribute='ExtensionCustomAttribute5'
)

Begin {

    If(!( Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue)) {
        Throw( 'Exchange cmdlets not available, connect to Exchange Online first.') 
    }
}

Process {

    Switch( $Mode) {
        'Clone' {

            $CurrentDG= Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
            If( !( $CurrentDG)) { 
                Throw ('ERR: Distribution Group {0} not found' -f $Identity)
            }

            $DGInfoFile= Join-Path $ExportFolder ('DG-{0}-Info.xml' -f ($CurrentDG.Name -replace '[\\/]','_'))

            $CloneName= '{0}{1}' -f $Prefix, $CurrentDG.Name
            $CloneDisplayName= '{0}{1}' -f $Prefix, $CurrentDG.DisplayName
            $CloneSmtpAddress= '{0}{1}' -f $Prefix, $CurrentDG.PrimarySmtpAddress
            $CloneAlias= '{0}{1}' -f $Prefix, $CurrentDG.Alias

            $Members= $CurrentDG | Get-DistributionGroupMember
            $SendAs= $CurrentDG | Get-RecipientPermission | Where {$_.AccessRights -like 'SendAs'}
            @{ 'Info'= $CurrentDG; 'Members'= $Members; 'SendAs'= $SendAs} | Export-CliXml $DGInfoFile

            $CloneDGParam= @{
                Name= $CloneName
                DisplayName= $CloneDisplayName
                PrimarySmtpAddress= $CloneSmtpAddress
                Alias= $CloneAlias
                Members= ($Members).Name
            }

            If( Get-DistributionGroup -Identity $CloneDGParam.Name -ErrorAction SilentlyContinue) {
                Write-Warning ('Distribution Group {0} already exists, reconfiguring' -f $CloneDGParam.Name)
            }
            Else {
                New-DistributionGroup @CloneDGParam
            }

            While(!( Get-DistributionGroup -Identity $CloneName)) {
                Write-Host ('Waiting for Distribution Group clone creation')
                Start-Sleep -Seconds 5
            }

            $CloneDGSetParam= @{
                Identity= $CloneName
                SimpleDisplayName= $CurrentDG.SimpleDisplayName
                AcceptMessagesOnlyFromSendersOrMembers= $CurrentDG.AcceptMessagesOnlyFromSendersOrMembers
                RejectMessagesFromSendersOrMembers= $CurrentDG.RejectMessagesFromSendersOrMembers
                HiddenFromAddressListsEnabled= $True
                BypassSecurityGroupManagerCheck= $True
                ManagedBy= $CurrentDG.ManagedBy
                BypassModerationFromSendersOrMembers= $CurrentDG.BypassModerationFromSendersOrMembers
                BypassNestedModerationEnabled= $CurrentDG.BypassNestedModerationEnabled
                CustomAttribute1= $CurrentDG.CustomAttribute1
                CustomAttribute2= $CurrentDG.CustomAttribute2
                CustomAttribute3= $CurrentDG.CustomAttribute3
                CustomAttribute4= $CurrentDG.CustomAttribute4
                CustomAttribute5= $CurrentDG.CustomAttribute5
                CustomAttribute6= $CurrentDG.CustomAttribute6
                CustomAttribute7= $CurrentDG.CustomAttribute7
                CustomAttribute8= $CurrentDG.CustomAttribute8
                CustomAttribute9= $CurrentDG.CustomAttribute9
                CustomAttribute10= $CurrentDG.CustomAttribute10
                CustomAttribute11= $CurrentDG.CustomAttribute11
                CustomAttribute12= $CurrentDG.CustomAttribute12
                CustomAttribute13= $CurrentDG.CustomAttribute13
                CustomAttribute14= $CurrentDG.CustomAttribute14
                CustomAttribute15= $CurrentDG.CustomAttribute15
                ExtensionCustomAttribute1= $CurrentDG.ExtensionCustomAttribute1
                ExtensionCustomAttribute2= $CurrentDG.ExtensionCustomAttribute2
                ExtensionCustomAttribute3= $CurrentDG.ExtensionCustomAttribute3
                ExtensionCustomAttribute4= $CurrentDG.ExtensionCustomAttribute4
                ExtensionCustomAttribute5= $CurrentDG.ExtensionCustomAttribute5
                GrantSendOnBehalfTo= $CurrentDG.GrantSendOnBehalfTo
                MailTip= $CurrentDG.MailTip
                MailTipTranslations= $CurrentDG.MailTipTranslations
                MemberDepartRestriction= $CurrentDG.MemberDepartRestriction
                MemberJoinRestriction= $CurrentDG.MemberJoinRestriction
                ModeratedBy= $CurrentDG.ModeratedBy
                ModerationEnabled= $CurrentDG.ModerationEnabled
                ReportToManagerEnabled= $CurrentDG.ReportToManagerEnabled
                ReportToOriginatorEnabled= $CurrentDG.ReportToOriginatorEnabled
                RequireSenderAuthenticationEnabled= $CurrentDG.RequireSenderAuthenticationEnabled
                SendModerationNotifications= $CurrentDG.SendModerationNotifications
                SendOofMessageToOriginatorEnabled= $CurrentDG.SendOofMessageToOriginatorEnabled
            }
            Set-DistributionGroup @CloneDGSetParam

            # CustomAttribute15 used to bring DGs in-scope of ADConnect, so clear it
            # Use ExtensionCustomAttribute5 and 4 for storing current EmailAddresses and legacyExchangeDN
            $CloneDGSetParam= @{
                Identity= $CloneName
                $EmailAddressesAttribute= $CurrentDG.EmailAddresses
                $legacyExchangeAttribute= $CurrentDG.LegacyExchangeDN
            }
            If( !( [string]::IsNullOrEmpty($SyncAttribute.IsPresent))) {
                $CloneDGSetParam.$SyncAttribute= $null
            }
            Set-DistributionGroup @CloneDGSetParam

            $SendAs | ForEach-Object { Add-RecipientPermission -Identity $CloneName -Trustee $_.Trustee -AccessRights 'SendAs' -Confirm:$false }
             
        }
        'Contact' {

            $CurrentDG= Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
            If( !( $CurrentDG)) { 
                Throw ('ERR: Distribution Group {0} not found' -f $Identity)
            }
            $CurrentContact= Get-MailContact -Identity $Identity -ErrorAction SilentlyContinue
            If( $CurrentContact) { 
                Throw ('ERR: MailContact {0} already exists' -f $Identity)
            }

            $DGInfoFile= Join-Path $ExportFolder ('DG-{0}-Info.xml' -f ($Identity -replace '[\\/]','_'))

            $Members= $CurrentDG | Get-DistributionGroupMember
            $SendAs= $CurrentDG | Get-ADPermission | Where {$_.AccessRights -like 'SendAs'}
            @{ 'Info'= $CurrentDG; 'Members'= $Members; 'SendAs'= $SendAs} | Export-CliXml $DGInfoFile

            Remove-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
            While( Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue) {
                Write-Host ('Waiting for Distribution Group {0} removal' -f $Identity)
                Start-Sleep 5
            }

            $TargetAddress= ($CurrentDG.EmailAddresses | Where {$_ -like '*.mail.onmicrosoft.com'} | Select -first 1).proxyAddressString
            Write-Host ('Creating MailContact{0} using targetAddress {1}' -f $Identity, $TargetAddress)

            $MailContactParam= @{
                Name= $CurrentDG.Name
                DisplayName= $CurrentDG.DisplayName
                PrimarySmtpAddress= $CurrentDG.PrimarySmtpAddress
                Alias= $CurrentDG.Alias
                OrganizationalUnit= $CurrentDG.OrganizationalUnit
                ExternalEmailAddress= $TargetAddress
            }
            New-MailContact @MailContactParam

            While(!( Get-MailContact -Identity $Identity)) {
                Write-Host ('Waiting for MailContact {0} creation' -f $Identity)
                Start-Sleep -Seconds 5
            }

            $MailContactSetParam= @{
                Identity= $Identity
                EmailAddressPolicyEnabled= $False
            }
            Set-MailContact @MailContactSetParam

            $MailContactSetParam= @{
                Identity= $Identity
                EmailAddresses= ($CurrentDG.EmailAddresses) + ('X500:{0}' -f $CurrentDG.legacyExchangeDN)
                SimpleDisplayName= $CurrentDG.SimpleDisplayName
                HiddenFromAddressListsEnabled= $False
                CustomAttribute1= $CurrentDG.CustomAttribute1
                CustomAttribute2= $CurrentDG.CustomAttribute2
                CustomAttribute3= $CurrentDG.CustomAttribute3
                CustomAttribute4= $CurrentDG.CustomAttribute4
                CustomAttribute5= $CurrentDG.CustomAttribute5
                CustomAttribute6= $CurrentDG.CustomAttribute6
                CustomAttribute7= $CurrentDG.CustomAttribute7
                CustomAttribute8= $CurrentDG.CustomAttribute8
                CustomAttribute9= $CurrentDG.CustomAttribute9
                CustomAttribute10= $CurrentDG.CustomAttribute10
                CustomAttribute11= $CurrentDG.CustomAttribute11
                CustomAttribute12= $CurrentDG.CustomAttribute12
                CustomAttribute13= $CurrentDG.CustomAttribute13
                CustomAttribute14= $CurrentDG.CustomAttribute14
                CustomAttribute15= $CurrentDG.CustomAttribute15
                ExtensionCustomAttribute1= $CurrentDG.ExtensionCustomAttribute1
                ExtensionCustomAttribute2= $CurrentDG.ExtensionCustomAttribute2
                ExtensionCustomAttribute3= $CurrentDG.ExtensionCustomAttribute3
                ExtensionCustomAttribute4= $CurrentDG.ExtensionCustomAttribute4
                ExtensionCustomAttribute5= $CurrentDG.ExtensionCustomAttribute5
                MailTip= $CurrentDG.MailTip
                MailTipTranslations= $CurrentDG.MailTipTranslations
            }
            If( [string]::IsNullOrEmpty($CurrentDG.$SyncAttribute.IsPresent)) {
                $MailContactSetParam.$SyncAttribute= $null
            }
            Set-MailContact @MailContactSetParam

        }
        'Convert' {

            $CloneIdentity= '{0}{1}' -f $Prefix, $Identity

            $CloneDG= Get-DistributionGroup -Identity $CloneIdentity -ErrorAction SilentlyContinue
            If( !( $CloneDG)) { 
                Throw ('ERR: Distribution Group clone {0} not found' -f $CloneIdentity)
            }

            $OrigName= $CloneDG.Name -Replace ('^{0}' -f $Prefix), ''
            $OrigDisplayName= $CloneDG.DisplayName -Replace ('^{0}' -f $Prefix), ''
            $OrigSmtpAddress= $CloneDG.PrimarySmtpAddress -Replace ('^{0}' -f $Prefix), ''
            $OrigAlias= $CloneDG.Alias -Replace ('^{0}' -f $Prefix), ''

            If( Get-DistributionGroup -Identity $OrigName -ErrorAction SilentlyContinue) {
                Throw ('ERR: Original Distribution Group {0}, move it out of sync scope first' -f $OrigName)
            }

            $CloneRestoreParam= @{
                Identity= $CloneIdentity
                Name= $OrigName
                DisplayName= $OrigDisplayName
                PrimarySmtpAddress= $OrigSmtpAddress
                Alias= $OrigAlias
                HiddenFromAddressListsEnabled= $False
                BypassSecurityGroupManagerCheck= $True
            }
            Set-DistributionGroup @CloneRestoreParam

            # Cannot use primarySmtpAddress & EmailAddresses in one call. Also, clear customAttr storing original EmailAddresses/legacyExchangeDN
            $CloneRestoreParam= @{
                Identity= $OrigName
                BypassSecurityGroupManagerCheck= $True
                EmailAddresses= ($CloneDG.$EmailAddressesAttribute) + ('X500:{0}' -f $CloneDG.$legacyExchangeAttribute)
                $EmailAddressesAttribute=$null
                $legacyExchangeAttribute=$null
            }
            Set-DistributionGroup @CloneRestoreParam

        }
        'Restore' {

            $DGInfoFile= Join-Path $ExportFolder ('DG-{0}-Info.xml' -f ($Identity -replace '[\\/]','_'))

            $RestoreData= Import-CliXMl $DGInfoFile
            $CurrentDG= $RestoreData['Info']
            $Members= $RestoreData['Members']
            $SendAs= $RestoreData['SendAs']

            $RestoreDGParam= @{
                Name= $Identity
                DisplayName= $CurrentDG.DisplayName
                PrimarySmtpAddress= $CurrentDG.primarySmtpAddress
                Alias= $CurrentDG.Alias
                Members= ($Members).Name
            }

            If( Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue) {
                Write-Warning ('Distribution Group {0} already exists, reconfiguring' -f $Identity)
            }
            Else {
                New-DistributionGroup @RestoreDGParam
            }

            While(!( Get-DistributionGroup -Identity $Identity)) {
                Write-Host ('Waiting for Distribution Group restoration')
                Start-Sleep -Seconds 5
            }

            $RestoreDGSetParam= @{
               Identity= $Identity
                EmailAddresses= ($CurrentDG.EmailAddresses) + ('X500:{0}' -f $CurrentDG.legacyExchangeDN)
                SimpleDisplayName= $CurrentDG.SimpleDisplayName
                AcceptMessagesOnlyFromSendersOrMembers= $CurrentDG.AcceptMessagesOnlyFromSendersOrMembers
                RejectMessagesFromSendersOrMembers= $CurrentDG.RejectMessagesFromSendersOrMembers
                HiddenFromAddressListsEnabled= $True
                BypassSecurityGroupManagerCheck= $True
                ManagedBy= $CurrentDG.ManagedBy
                BypassModerationFromSendersOrMembers= $CurrentDG.BypassModerationFromSendersOrMembers
                BypassNestedModerationEnabled= $CurrentDG.BypassNestedModerationEnabled
                CustomAttribute1= $CurrentDG.CustomAttribute1
                CustomAttribute2= $CurrentDG.CustomAttribute2
                CustomAttribute3= $CurrentDG.CustomAttribute3
                CustomAttribute4= $CurrentDG.CustomAttribute4
                CustomAttribute5= $CurrentDG.CustomAttribute5
                CustomAttribute6= $CurrentDG.CustomAttribute6
                CustomAttribute7= $CurrentDG.CustomAttribute7
                CustomAttribute8= $CurrentDG.CustomAttribute8
                CustomAttribute9= $CurrentDG.CustomAttribute9
                CustomAttribute10= $CurrentDG.CustomAttribute10
                CustomAttribute11= $CurrentDG.CustomAttribute11
                CustomAttribute12= $CurrentDG.CustomAttribute12
                CustomAttribute13= $CurrentDG.CustomAttribute13
                CustomAttribute14= $CurrentDG.CustomAttribute14
                CustomAttribute15= $CurrentDG.CustomAttribute15
                ExtensionCustomAttribute1= $CurrentDG.ExtensionCustomAttribute1
                ExtensionCustomAttribute2= $CurrentDG.ExtensionCustomAttribute2
                ExtensionCustomAttribute3= $CurrentDG.ExtensionCustomAttribute3
                ExtensionCustomAttribute4= $CurrentDG.ExtensionCustomAttribute4
                ExtensionCustomAttribute5= $CurrentDG.ExtensionCustomAttribute5
                GrantSendOnBehalfTo= $CurrentDG.GrantSendOnBehalfTo
                MailTip= $CurrentDG.MailTip
                MailTipTranslations= $CurrentDG.MailTipTranslations
                MemberDepartRestriction= $CurrentDG.MemberDepartRestriction
                MemberJoinRestriction= $CurrentDG.MemberJoinRestriction
                ModeratedBy= $CurrentDG.ModeratedBy
                ModerationEnabled= $CurrentDG.ModerationEnabled
                ReportToManagerEnabled= $CurrentDG.ReportToManagerEnabled
                ReportToOriginatorEnabled= $CurrentDG.ReportToOriginatorEnabled
                RequireSenderAuthenticationEnabled= $CurrentDG.RequireSenderAuthenticationEnabled
                SendModerationNotifications= $CurrentDG.SendModerationNotifications
                SendOofMessageToOriginatorEnabled= $CurrentDG.SendOofMessageToOriginatorEnabled
            }
            Set-DistributionGroup @RestoreDGSetParam
        }

    }



}

End {

}
