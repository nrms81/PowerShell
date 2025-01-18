
# Скрипт очистки кеша 1С для текущего пользователя


# +------------------------------------------------------------------+
# |                 CleanupCashe1CForCurrentUser.ps1                 |
# +------------------------------------------------------------------+

# В переменную $Procs получаем список процессов содержащих в названии "1cv8"
# Т.е. находим запущенные экземпляры 1С 

$Procs = Get-Process -Name "1cv8*"

# Если нашли, то принудительно их завершаем

If ($Procs.count -gt 0) {
	Stop-Process -InputObject $Procs -Force
}

# $Paths - массив с путями к папкам AppData\Roaming и AppData\Local

$Paths = 	@($Env:APPDATA, 
			$Env:LOCALAPPDATA)

# К каждой из папок в массиве $Paths добавляем "\1C\1cv8\"
# Проверяем существует ли такой путь
# И если да, то удаляем в нем все папки соответствующие шаблону '........-....-....-....-............'

ForEach ($Path in $Paths) {
	$P = $Path+'\1C\1cv8\'
    If (Test-Path $P) {
		Get-ChildItem -Path $P -Directory | Where-Object { $_.Name -match '........-....-....-....-............' } | Remove-Item -ErrorAction SilentlyContinue -Force -Recurse
	}
	$P = $Null
}
