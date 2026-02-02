# Скрипт добавления в адресную книгу Exchange почтовых ящиков из доменов с которыми есть доверенные отношения.

# Задача:
# Имеется несколько доменов AD. Каждый в своем лесу. Между ними настроены трасты.
# В адресной книге Exchange каждого домена необходимо иметь адреса пользователей из соседних доменов. 
# Вариант с Microsoft Identity Manager не рассматривается по финансовым мотивам.

# Решение: 
# Получаем пользователей и группы с почтовыми адресами из каждого домена с которым есть двухсторонние доверенные отношения.
# На основе полученных данных создаем контакты в текущем домене и загружаем их в Exchange.
# Поддерживаем полученные контакты в актуальном состоянии.
# Из доменов в которых найден сервер Exchange будем брать все объекты у которых есть почтовый адрес в "proxyAddresses"
# Таким образом сюда попадут не только сами пользователи, но и группы рассылки (в том числе и динамические)
# Из доменов без Exchange возьмем объекты у которых заполнен артибут "mail"
# Для сопоставления контакта пользователю будем использовать "objectGUID" пользователя записав его в "extensionAttribute1" контакта

# Использование:
# Для работы скрипта необходимы права для изменения учетных записей ActiveDirectory и работы с Exchange


# +------------------------------------------------------------------+
# |      AddContactsFromTrustedDomainsToExchangeAddressBook.ps1      |
# +------------------------------------------------------------------+

# $AttributeListBase - базовый массив атрибутов которые будем синхронизировать 

$AttributeListBase =	@("telephoneNumber",
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
						
# $AttributeListUser - массив атрибутов которые будем считывать у пользователей доверенных доменов

$AttributeListUser = $AttributeListBase + "proxyAddresses" + "mail"

# $AttributeListContact - массив атрибутов которые будем считывать у контактов текущего домена

$AttributeListContact = $AttributeListBase + "proxyAddresses" + "extensionAttribute1"

# $OUForContacts - имя базового контейнера в котором будем создавать подконтейнеры с контактами для каждого доверенного домена 

$OUForContacts = "Contacts"

# В $TrustedDomains получаем список доменов с которыми есть двухсторонние доверенные отношения
# Свойство "flatname" - плоское имя домена. А "target" - полное

$TrustedDomains = Get-ADTrust -Filter {Direction -eq "BiDirectional"} -Properties flatname, target

# В $PDC получаем имя контроллера в текущем домене с ролью "Эмулятор PDC". Создавать контакты будем на нем

$PDC = (Get-ADDomainController -Discover -Service PrimaryDC).HostName.Value

# $Exchange - имя почтового сервера
# Для его получения находим все компьютеры домена у которых есть SPN содержащий "exchangeMDB" и берем первый из них
# Т.е. находим все почтовые сервера с ролью "Mailbox". В инсталляциях с одним-двумя почтовыми серверами этого вполне достаточно

$Exchange = (Get-ADComputer -Filter {servicePrincipalName -like 'exchangeMDB*'} -ResultSetSize 1).DNSHostName

# $ExchangePowershell - путь для модуля ExchangePowerShell

$ExchangePowershell = "http://" + $Exchange + "/powershell/"

# Создаем сессию с почтовым сервером
# https://learn.microsoft.com/ru-ru/powershell/exchange/connect-to-exchange-servers-using-remote-powershell?view=exchange-ps#connect-to-a-remote-exchange-server

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangePowershell -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber

# Получаем в $RootOU контейнер указанный в $OUForContacts

$RootOU = Get-ADOrganizationalUnit -Filter {Name -eq $OUForContacts} -Server $PDC

# Если контейнер отсутствует, то создаем

If ($RootOU -eq $null) {
	$RootOU = New-ADOrganizationalUnit -Name $OUForContacts -PassThru -Server $PDC
}

# Основной цикл скрипта
# Перебираем все домены из списка $TrustedDomains

ForEach ($TrustedDomain in $TrustedDomains) {
	
# $Users - массив в который будем получать пользователей из доверенного домена

	$Users = @()

# $Contacts - массив в который будем получать контакты для этого домена

	$Contacts = @()

# Если домен отвечает на пинги работаем с ним, если нет - переходим к следующему

	If (Test-Connection $TrustedDomain.target -Count 1 -Quiet) {

# Получаем в $TrustedDomainOU контейнер для контактов
# Контейнер будет располагаться в контейнере $OUForContacts и иметь имя эквивалентное плоскому имени домена, который обрабатывается в данной итерации

		$TrustedDomainOU = Get-ADOrganizationalUnit -Filter {Name -eq $TrustedDomain.flatname} -SearchBase $RootOU -Server $PDC
		
# Если такого контейнера нету, то создаем его

		If ($TrustedDomainOU -eq $null) {
			$TrustedDomainOU = New-ADOrganizationalUnit -Name $TrustedDomain.flatname -Path $RootOU -PassThru -Server $PDC
		}

# Пробуем найти сервер Exchange в обрабатываемом домене, аналогично тому как получали имя почтового сервера в текущем

		$TrustedDomainExchange = Get-ADComputer -Filter {servicePrincipalName -like 'exchangeMDB*'} -Server $TrustedDomain.target -ResultSetSize 1

# Если нашли, то устанавливаем для переменной $ExchangeangeExistsInTrustedDomain значение "истина" и формируем фильтр поиска почтовых ящиков 
# Значение $ExchangeangeExistsInTrustedDomain будет учитываться при выборе из какого атрибута брать почтовый адрес для контакта
# Если в обрабатываемом домене есть Exchange то почтовый адрес будем брать из атрибута "proxyAddresses", иначе из "mail"
# Фильтр для доменов с Exchange содержит следующие условия:
# "objectClass -ne "contact"" - объект не является контактом
# "msExchHideFromAddressLists -notlike "*"" - у объекта не установлен флаг "Скрыть из адресной книги"
# "proxyAddresses -ne "null"" - у объекта есть почтовый адрес
# "-not UserAccountControl -band 2" - объект не является отключенной учетной записью

		If ($TrustedDomainExchange.DNSHostName -ne $null) {
			$ExchangeangeExistsInTrustedDomain = $True
			$FilterForUsers = "objectClass -ne ""contact"" -and msExchHideFromAddressLists -notlike ""*"" -and proxyAddresses -ne ""null"" -and -not UserAccountControl -band 2"
		} else {

# Если не нашли, то устанавливаем для переменной $ExchangeangeExistsInTrustedDomain значение "ложь" и формируем фильтр поиска почтовых ящиков
# Фильтр для доменов без Exchange содержит следующие условия:
# "objectClass -ne "contact"" - объект не является контактом
# "mail -ne "null"" - у объекта есть почтовый адрес
# "-not UserAccountControl -band 2" - объект не является отключенной учетной записью

			$ExchangeangeExistsInTrustedDomain = $False
			$FilterForUsers = "objectClass -ne ""contact"" -and mail -ne ""null"" -and -not UserAccountControl -band 2"
		}

# Заполняем массив пользователей из обрабатываемого домена
		
		$Users = Get-ADObject -Filter $FilterForUsers -Properties $AttributeListUser -Server $TrustedDomain.target -ResultPageSize 1000

# Заполняем массив контактов соответствующих домену

		$Contacts = Get-ADObject -Filter {objectClass -eq "contact"} -SearchBase $TrustedDomainOU -Properties $AttributeListContact -ResultPageSize 1000

# Начинаем обработку полученного массива пользователей

		ForEach ($User in $Users) {

# Если в обрабатываемом домене есть Exchange то в $Mail заносим адрес указанный в качестве первичного почтового адреса а атрибуте "proxyAddresses"
# Атрибут "proxyAddresses" может содержать несколько строк вида "smtp:user@consoto.com". У первичного адреса "SMTP" записано заглавными буквами
# Поэтому для $Mail находим строку начинающуюся с "SMTP" и берем символы после ":"
# Если Exchange в домене нет то присваиваем $Mail значение атрибута "mail"

			If ($ExchangeangeExistsInTrustedDomain -eq $True) {
				$Mail = ($User.proxyAddresses -cmatch "^SMTP:").Split(":")[1]
			} else {
				$Mail = $User.mail
            }

# $MailAlias - псевдоним для контакта. Должен быть по возможности уникальным.
# Поэтому для него берем имя почтового ящика из $Mail и через точку добавляем к нему имя обрабатываемого домена
# Т.е. для user@consoto.com в случае если почтовый домен равен имени домена AD псевдоним будет user.consoto.com

			$MailAlias = $Mail.Split("@")[0] + "." +$TrustedDomain.target

# Переменная $ContactExists указывает найден ли контакт соответствующий пользователю

			$ContactExists = $False

# Начинаем сравнение обрабатываемого пользователя со списком контактов
# Для идентификации используется значение атрибута "objectGUID" пользователя, которое записывается в атрибут "extensionAttribute1" контакта при его создании

			ForEach ($Contact in $Contacts) {

# Если атрибут "objectGUID" пользователя равен атрибуту "extensionAttribute1" контакта, то значит контакт для этого пользователя есть
# Теперь необходимо сравнить их атрибуты и внести правки при необходимости

				If ($User.objectGUID -eq $Contact.extensionAttribute1) {

# $UpdateAttribute - хэш-таблица с изменяемыми атрибутами
# $AddAttribute - хэш-таблица с добавляемыми атрибутами
# $RemoveAttribute - массив атрибутов для очистки

					$UpdateAttribute = @{}
					$AddAttribute = @{}
					$RemoveAttribute = @()

# Перебираем атрибуты из базового списка $AttributeListBase

					ForEach ($Attribute in $AttributeListBase) {

# Если атрибут заполнен и у контакта и у пользователя, то сравниваем их. При расхождении заносим в $UpdateAttribute значение из атрибута пользователя
# Т.к. атрибут может быть не только строкового или числового типа, а в том числе и многостроковым, то сравниваем через Compare-Object
# Атрибут многострокового типа, по факту хэш-таблицу, нельзя передать напрямую командлету Set-ADObject
# Тип "ADPropertyValueCollection" как раз многостроковый (хэш-таблица). Его обрабатываем отдельно
# Для этого его значениями заполняем массив $ValueCollectionUpdate
# Данные из $UpdateAttribute будут использованы для обновления контакта

						If (($User.$Attribute.Count -gt 0) -and ($Contact.$Attribute.Count -gt 0)) {
							If ([bool](Compare-Object $User.$Attribute $Contact.$Attribute)) {
								If ($User.$Attribute.GetType().Name -eq "ADPropertyValueCollection") {
									$ValueCollectionUpdate = @()
									ForEach ($Value in $User.$Attribute.Value) {
										$ValueCollectionUpdate = $ValueCollectionUpdate + $Value
									}
									$UpdateAttribute = $UpdateAttribute + @{$Attribute = $ValueCollectionUpdate}
								} Else {
									$UpdateAttribute = $UpdateAttribute + @{$Attribute = $User.$Attribute}
								}
							}
						}

# Если атрибут заполнен у пользователя, но не заполнен у контакта, то заносим его в $AddAttribute для добавления значений к контакту
# Обработка аналогична обновлению контакта
							
						If (($User.$Attribute.Count -gt 0) -and ($Contact.$Attribute.Count -eq 0)) {
							If ($User.$Attribute.GetType().Name -eq "ADPropertyValueCollection") {
								$ValueCollectionAdd = @()
								ForEach ($Value in $User.$Attribute.Value) {
									$ValueCollectionAdd = $ValueCollectionAdd + $Value
								}
								$AddAttribute = $AddAttribute + @{$Attribute = $ValueCollectionAdd}
							} Else {
								$AddAttribute = $AddAttribute + @{$Attribute = $User.$Attribute}
							}
						}

# Если атрибут заполнен у контакта, но не заполнен у пользователя, то заносим его в $RemoveAttribute для последующего удаления
							
						If (($User.$Attribute.Count -eq 0) -and ($Contact.$Attribute.Count -gt 0)) {
							$RemoveAttribute = $RemoveAttribute + $Attribute
						}
					}

# Перебор основных атрибутов закончен
# Изменяем контакт 
# Проверяем что переменные заполняемые выше не пусты и производим удаление/добавление/обновление атрибутов

					If ($RemoveAttribute.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Clear $RemoveAttribute -Confirm:$false -Server $PDC
					}
					If ($AddAttribute.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Add $AddAttribute -Confirm:$false -Server $PDC
					}
					If ($UpdateAttribute.Count -gt 0) {
						Set-ADObject -Identity $Contact.distinguishedName -Replace $UpdateAttribute -Confirm:$false -Server $PDC
					}

# В переменную $ContactMail получаем почтовый адрес контакта

					$ContactMail = ($Contact.proxyAddresses -cmatch "^SMTP:").Split(":")[1]

# Проверяем почтовые адреса контакта и пользователя
# Если они не совпадают то меняем адрес у контакта

					If ($Mail -ne $ContactMail) {
						Set-MailContact -Identity $Contact.distinguishedName -Alias $MailAlias -ExternalEmailAddress $Mail -PrimarySmtpAddress $Mail -Confirm:$false -DomainController $PDC
					}

# Переименовываем контакт если был переименован соответствующий ему пользователь					
					
					If ($User.displayName -ne $Contact.displayName) {
						Rename-ADObject -Identity $Contact.distinguishedName -NewName $User.displayName -Confirm:$false -Server $PDC
					}

# Т.к. контакт найден присваиваем $ContactExists значение "истина"
					
					$ContactExists = $True

# Исключаем контакт из массива и прерываем перебор
# Таким образом, в конце концов в массиве останутся только контакты для которых не нашлись соответствующие им пользователи

					$Contacts = $Contacts -ne $Contact					
					break
				}
			}

# Закончили перебор контактов
# Если $ContactExists равен "ложь", т.е. соответствующий пользователю контакт не найден, то создаем его
			
			If ($ContactExists -eq $False) {

# $AttributesForNewContact - хэш-таблица для атрибутов нового контакта

				$AttributesForNewContact = @{}

# Добавляем в нее атрибут "extensionAttribute1" в который занесем значение атрибута "objectGUID" пользователя
# Данная связка будет использоваться для сопоставления пары "контакт-пользователь"

				$AttributesForNewContact = $AttributesForNewContact + @{extensionAttribute1 = $User.objectGUID}

# Заполняем $AttributesForNewContact атрибутами пользователя

				ForEach ($Attribute in $AttributeListBase) {
					If ($User.$Attribute.Count -gt 0) {
						$AttributesForNewContact = $AttributesForNewContact + @{$Attribute = $User.$Attribute}
					}
				}

# Создаем контакт в Active Directory

				$NewContact = $null
				$NewContact = New-ADObject -Type "Contact" -Name $User.displayname -otherAttributes $AttributesForNewContact -Path $TrustedDomainOU -PassThru -Confirm:$false -Server $PDC

# Создаем на основе этого контакта почтовый контакт на сервере Exchange

				Enable-MailContact -Identity $NewContact.distinguishedName -Alias $MailAlias -ExternalEmailAddress $Mail -PrimarySmtpAddress $Mail -DomainController $PDC
			}
		}

# Закончили перебор пользователей
# Т.к. в массиве $Contacts остались только контакты которым не соответствует ни один пользователь, то удаляем их
		
		ForEach ($Contact in $Contacts) {
			Remove-MailContact -Identity $Contact.distinguishedName -Confirm:$false -DomainController $PDC
		}
	}
}

# Закончили перебор доверенных доменов
# Делаем паузу в 15 секунд

Start-Sleep -Seconds 15

# Обновляем адресные книги на сервере

Get-AddressList | Update-AddressList
Start-Sleep -Seconds 5
Get-GlobalAddressList | Update-GlobalAddressList
Start-Sleep -Seconds 5
Get-OfflineAddressBook | Update-OfflineAddressBook
Start-Sleep -Seconds 5

# Закрываем сессию с почтовым сервером

Remove-PSSession $Session
