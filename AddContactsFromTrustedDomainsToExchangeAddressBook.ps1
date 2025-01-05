# Скрипт добавления в адресную книгу Exchange почтовых ящиков из доменов с которыми есть доверенные отношения.

# Задача:
# Имеется несколько доменов AD. Каждый в своем лесу. Между ними настроены трасты.
# В адресной книге Exchange каждого домена необходимо иметь адреса пользователей из соседних доменов. 
# Вариант с Microsoft Identity Manager не рассматривается по финансовым мотивам.

# Решение: 
# Получаем пользователей и группы с почтовыми адресами из каждого домена с которым есть двухсторонние доверенные отношения.
# На основе полученных данных создаем контакты в текущем домене и загружаем их в Exchange.
# Поддерживаем полученные контакты в актуальном состоянии.

# Использование:
# 

# TODO: Доделать описание. Прокомментировать код.


# +------------------------------------------------------------------+
# |      AddContactsFromTrustedDomainsToExchangeAddressBook.ps1      |
# +------------------------------------------------------------------+


$AttributeListBase =   @("telephoneNumber",
                        "othertelephone",
                        "ipPhone",
                        "otherIpPhone",
                        "mobile",
                        "otherMobile",
                        "facsimileTelephoneNumber",
                        "otherfacsimileTelephoneNumber",
                        "homePhone",
                        "otherHomePhone",
                        "postOfficeBox",
                        "postalCode",
                        "c",
                        "co",
			"countryCode",
                        "st",
                        "l",
                        "streetAddress",
                        "physicalDeliveryOfficeName",
                        "employeeID",
                        "employeeType",
			"thumbnailPhoto",
                        "displayName",
                        "sn",
                        "givenName",
                        "description",
                        "info",
                        "wWWHomePage",
                        "url",
                        "company",
                        "department",
                        "title")
$AttributeListUser = $AttributeListBase + "proxyAddresses" + "mail"
$AttributeListContact = $AttributeListBase + "proxyAddresses" + "extensionAttribute1"
$OUForContacts = "Contacts"
$TrustedDomains = Get-ADTrust -Filter {Direction -eq "BiDirectional"} -Properties flatname, target
$PDC = (Get-ADDomainController -Discover -Service PrimaryDC).HostName.Value
$Exchange = (Get-ADComputer -Filter {servicePrincipalName -like 'exchangeMDB*'} -ResultSetSize 1).DNSHostName
$ExchangePowershell = "http://" + $Exchange + "/powershell/"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangePowershell -Authentication Kerberos
Import-PSSession $session -DisableNameChecking -AllowClobber

$RootOU = Get-ADOrganizationalUnit -Filter {Name -eq $OUForContacts} -Server $PDC
if ($RootOU -eq $null) {
	$RootOU = New-ADOrganizationalUnit -Name $OUForContacts -PassThru -Server $PDC
}
foreach ($TrustedDomain in $TrustedDomains) {
	$Users = @()
	$Contacts = @()
	If (Test-Connection $TrustedDomain.target -Count 1 -Quiet) {
		$TrustedDomainOU = Get-ADOrganizationalUnit -Filter {Name -eq $TrustedDomain.flatname} -SearchBase $RootOU -Server $PDC
		if ($TrustedDomainOU -eq $null) {
			$TrustedDomainOU = New-ADOrganizationalUnit -Name $TrustedDomain.flatname -Path $RootOU -PassThru -Server $PDC
		}
		$TrustedDomainExchange = Get-ADComputer -Filter {servicePrincipalName -like 'exchangeMDB*'} -Server $TrustedDomain.target -ResultSetSize 1
		If ($TrustedDomainExchange.DNSHostName -ne $null) {
			$ExchangeangeExistsInTrustedDomain = $True
			$FilterForUsers = "objectClass -ne ""contact"" -and msExchHideFromAddressLists -notlike ""*"" -and proxyAddresses -ne ""null"" -and -not UserAccountControl -band 2"
		} else {
			$ExchangeangeExistsInTrustedDomain = $False
			$FilterForUsers = "objectClass -ne ""contact"" -and mail -ne ""null"" -and -not UserAccountControl -band 2"
		}
		$Users = Get-ADObject -Filter $FilterForUsers -Properties $AttributeListUser -Server $TrustedDomain.target -ResultPageSize 1000
		$Contacts = Get-ADObject -Filter {objectClass -eq "contact"} -SearchBase $TrustedDomainOU -Properties $AttributeListContact -ResultPageSize 1000
		foreach ($User in $Users) {
			if ($ExchangeangeExistsInTrustedDomain -eq $True) {
				$Mail = ($User.proxyAddresses -cmatch "^SMTP:").Split(":")[1]
			} else {
				$Mail = $User.mail
            }
			$ContactExists = $False
			foreach ($Contact in $Contacts) {
				if ($User.objectGUID -eq $Contact.extensionAttribute1) {
					$Update = @{}
					$Add = @{}
					$Remove = @()

					foreach ($Attribute in $AttributeListBase) {
						if (($User.$Attribute.Count -gt 0) -and ($Contact.$Attribute.Count -gt 0)) {
							if ([bool](Compare-Object $User.$Attribute $Contact.$Attribute)) {
								if ($User.$Attribute.GetType().Name -eq "ADPropertyValueCollection") {
									$SU = @()
									foreach ($V in $User.$Attribute.Value) {
										$SU = $SU + $V
									}
									$Update = $Update + @{$Attribute = $SU}
								} Else {
									$Update = $Update + @{$Attribute = $User.$Attribute}
								}
							}
						}
							
						if (($User.$Attribute.Count -gt 0) -and ($Contact.$Attribute.Count -eq 0)) {
							if ($User.$Attribute.GetType().Name -eq "ADPropertyValueCollection") {
								$SU = @()
								foreach ($V in $User.$Attribute.Value) {
									$SU = $SU + $V
								}
								$Add = $Add + @{$Attribute = $SU}
							} Else {
								$Add = $Add + @{$Attribute = $User.$Attribute}
							}
						}
							
						if (($User.$Attribute.Count -eq 0) -and ($Contact.$Attribute.Count -gt 0)) {
							$Remove = $Remove + $Attribute
						}

					}
					if ($Remove.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Clear $Remove -Confirm:$false -Server $PDC
					}
					if ($Add.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Add $Add -Confirm:$false -Server $PDC
					}
					if ($Update.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Replace $Update -Confirm:$false -Server $PDC
					}

					
					$CMail = ($Contact.proxyAddresses -cmatch "^SMTP:").Split(":")[1]
					$alc = $Mail.Split("@")[0] + "." +$TrustedDomain.target
					if ($Mail -ne $CMail) {
						Set-MailContact -Identity $Contact.distinguishedName -Alias $alc -ExternalEmailAddress $Mail -PrimarySmtpAddress $Mail -Confirm:$false -DomainController $PDC
					}
					
					
					if ($User.displayName -ne $Contact.displayName) {
						Rename-ADObject -Identity $Contact.distinguishedName -NewName $User.displayName -Confirm:$false -Server $PDC
					}
					
					$ContactExists = $True
					$Contacts = $Contacts -ne $Contact					
					break
				}
			}
			if ($ContactExists -eq $False) {
				$AttributesForNewContact = @{}
				$AttributesForNewContact = $AttributesForNewContact + @{extensionAttribute1 = $User.objectGUID}
				foreach ($Attribute in $AttributeListBase) {
					if ($User.$Attribute.Count -gt 0) {
						$AttributesForNewContact = $AttributesForNewContact + @{$Attribute = $User.$Attribute}
					}
				}
				$al = $Mail.Split("@")[0] + "." +$TrustedDomain.target
				$NewContact = New-ADObject -Type "Contact" -Name $User.displayname -otherAttributes $AttributesForNewContact -Path $TrustedDomainOU -PassThru -Confirm:$false -Server $PDC

				$NewMailContact = Enable-MailContact -Identity $NewContact.distinguishedName -Alias $al -ExternalEmailAddress $Mail -PrimarySmtpAddress $Mail -DomainController $PDC
			}
		}
		foreach ($Contact in $Contacts) {
			$Contact.distinguishedName
			Remove-MailContact -Identity $Contact.distinguishedName -Confirm:$false -DomainController $PDC
		}
	}
}

Start-Sleep -Seconds 15

Get-AddressList | Update-AddressList
Start-Sleep -Seconds 5
Get-GlobalAddressList | Update-GlobalAddressList
Start-Sleep -Seconds 5
Get-OfflineAddressBook | Update-OfflineAddressBook
Start-Sleep -Seconds 5

Remove-PSSession $Session
