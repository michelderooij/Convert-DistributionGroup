# Convert-DistributionGroup
Script to support converting synchronized Distribution Groups to cloud-only Distribution Groups

Convert-DistributionGroup.ps1 [-Name] <string> [-Mode <string>] [-ExportFolder <string>] [-Prefix <string>] [-SyncAttribute <string>] [-legacyExchangeAttribute <string>] [-EmailAddressesAttribute <string>] [<CommonParameters>]

The script has three operating modes, depending on the phase of the conversion:
1) CLONE
   In Exchange Online, create clone objects from the existing (sychronized) Distribution Groups, including Recipient Permissions
   The following properties will be prefixed (default Clone) to prevent conflicts: Name, DisplayName, Alias, PrimarySmtpAddress
   The property specified by SyncAttribute (default CustomAttribute15) will be cleared (in the clone)
   The property specified by legacyExchangeAttribute is used to store the current legacyExchangeDN
   The property specified by EmailAddressesAttribute is used to store the current EmailAddresses
2) CONTACT
   In Exchange On-premises, convert the Distribution Groups to Contacts for representation (addresslists) etc.
   The property specified by SyncAttribute (default CustomAttribute15) will be cleared, with the intention to put contacts out of sync scope
3) CONVERT
   In Exchange Online, remove the Distribution Group after checking if the clone is present.
   The following properties will be stripped from the prefix: Name, DisplayName, Alias, PrimarySmtpAddress
   The property specified by EmailAddressesAttribute is used to stamp the original EmailAddresses
   The property specified by legacyExchangeAttribute is used to add an X500 proxy to EmailAddresses for resolving communications using name cache or old messages
   
Use mode RESTORE for restoring accidentally converted Distribution Groups using the information in the XML save files.


### About

For more information on this script, as well as usage and examples, see
the related blog article, [Exchange v15 Unattended Setup](https://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script/).

## License

This project is licensed under the MIT License - see the LICENSE for details.
