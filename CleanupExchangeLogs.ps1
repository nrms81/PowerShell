
# Скрипт очистки логов Exchange 2013 и новее

# Задача:
# Логи Exchange и IIS со временем могут занять все свободное место на диске 

# Решение: 
# Будем удалять логи которые старше количества дней указанных в переменной $Days
# По факту это чуть-чуть переработанный скрипт отсюда:
# https://winitpro.ru/index.php/2021/03/15/ochistka-i-peremeshhenie-logov-v-exchange/
# Основное отличие - папки с логами будем находить сами, не используя захардкоденные пути


# +------------------------------------------------------------------+
# |                      CleanupExchangeLogs.ps1                     |
# +------------------------------------------------------------------+

# Подключаем модуль для работы с IIS. Понадобится для поиска папок с логами расположенных на нем сайтов

Import-Module WebAdministration

# $Days - количество дней за которое будем хранить логи. Все что старше - удаляем

$Days = 15

# Массив содержит пути к папкам логов Exchange
# env:ExchangeInstallPath - путь к папке с установленным Exchange 
# По дефолту C:\Program Files\Microsoft\Exchange Server\V15, но может быть изменено при установке

$DirectoriesWithLogs = 	@("${env:ExchangeInstallPath}Logging\",
						"${env:ExchangeInstallPath}Bin\Search\Ceres\Diagnostics\ETLTraces\",
						"${env:ExchangeInstallPath}Bin\Search\Ceres\Diagnostics\Logs\",
						"${env:ExchangeInstallPath}TransportRoles\Logs\")

# $Sites - получаем сайты запущенные на IIS

$Sites = (Get-Website)

# Перебираем полученные сайты

ForEach ($Site in $Sites) {

# $LogDirectory - путь к папке логов сайта
# Путь может содержать (и по умолчанию содержит) переменные окружения вида %SYSTEMDRIVE%
# Нужно их привести к виду env:SYSTEMDRIVE
# Для этого находим подстроку расположенную между двумя %
# Далее, если такая подстрока есть, получаем на ее основе переменную нужного формата и заменяем в $LogDirectory

	$LogDirectory = $Site.LogFile.Directory
	$StartIndex = $LogDirectory.IndexOf("%") + 1
	$Length = $LogDirectory.LastIndexOf("%") - $StartIndex
	If (($StartIndex -gt 0) -and ($Length -gt 0)) {
		$Env = $LogDirectory.Substring($StartIndex, $length)
		$EnvOld = "%"+$Env+"%"
		$EnvNew = (Get-ChildItem -Path Env:$Env).value
		$LogDirectory = $LogDirectory -replace $EnvOld, $EnvNew
	}

# Если полученный путь отсутствует в $DirectoriesWithLogs, то добавляем его

	If (!($LogDirectory -in $DirectoriesWithLogs)) {
		$DirectoriesWithLogs += $LogDirectory
	}
}

# Для каждой папки из массива $DirectoriesWithLogs проверяем ее существование
# Если папка существует то ищем в ней и подпапках файлы *.log *.blg *.etl
# И если они старше чем количество дней в переменной $Days - безжалостно удаляем

ForEach ($DirectoryWithLogs in $DirectoriesWithLogs) {
	If (Test-Path $DirectoryWithLogs) {
		Get-ChildItem $DirectoryWithLogs -Recurse -File | Where-Object { $_.Name -like "*.log" -or $_.Name -like "*.blg" -or $_.Name -like "*.etl" } | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$Days) | Remove-Item -ErrorAction SilentlyContinue
	}
}
