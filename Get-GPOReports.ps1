###########################################################################
# Имя скрипта:
# 	Get-GPOReports.ps1
# Назначение:
# 	Получение HTML-отчетов по групповым политикам из AD
# Описание:
#	В скрипте заданы следующие параметры:
#		Список доменов, из которых выгружать отчеты (NetBIOS и DNS-имена)
#		Путь в файловой системе, куда выгружать отчеты
#	Для каждого из доменов производится запуск команд:
#		Получение имен всех политик данного домена
#		Выгрузка отчета по каждой из политик в формате HTML
#	Для каждого из доменов выгрузка осуществляется в отдельную директорию.
#	В процессе работы скрипта ведется журнал.
###########################################################################

########################################################################### 
# Параметры и переменные
###########################################################################
$LoggingPreference="Continue"
$Date = Get-Date
$LogFilename = "Get-GPOReports{0}{1:d2}{2:d2}" -f $Date.year,$Date.month,$Date.day
$PathToReports = "C:\Temp\GPO\"
$Domains = @{
	example="example.com"
    test="test.org"
    }
########################################################################### 
# Функции
###########################################################################

Function Write-Log {
# Назначение: 
# 	Ведение журнала
    [cmdletbinding()]
    Param(
    [Parameter(Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$Message,
    [Parameter(Position=1)]
    [string]$LogFilePath="C:\Logs\$LogFilename.txt"
    )
    # Передача сообщения для Write-Verbose, если -Verbose был обнаружен
    Write-Verbose -Message $Message
    # Записывает в журнал, если переменная $LoggingPreference установлена в Continue
    if ($LoggingPreference -eq "Continue"){
        $TimeStamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
        Write-Output "$TimeStamp $Message" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
}
# Add-Type -Path "C:\Program Files\SDM Software\SDM Software GPMC Cmdlets\SDMGPOCmdlets.dll"
# import-module SDM-GPMC
Write-Log -Message "==============================Начало работы скрипта=============================="
Write-Log -Message "Загрузка командлетов PowerShell, необходимых для управления групповыми политиками через API GPMC"
Trap {Write-Log -Message "Ошибка: $_";Continue}
Add-PSSnapin SDMSoftware.PowerShell.GPMC -ErrorAction Stop
Set-Location $PathToReports

# Удаление старых отчетов
Write-Log -Message "Начало удаления старых отчетов из каталога $PathToReports"
$FilesToDelete = Get-Childitem $PathToReports -Filter *.html -Recurse
Write-Log -Message "Количество отчетов для удаления: $($FilesToDelete.Count)"
$FilesToDelete | ForEach-Object {
	Write-Log -Message "Удаление отчета: $($_.FullName)"
	Trap {Write-Log -Message "Ошибка: $_";Continue}
	Remove-Item $_.FullName -ErrorAction Stop 
	}
$ReportsInFolderAfterDelete = Get-Childitem $PathToReports -Filter *.html -Recurse
Write-Log -Message "Количество отчетов в папке $($PathToReports): $($ReportsInFolderAfterDelete.Count)"
Write-Log -Message "Удаление завершено"
Write-Log -Message "Начало получения отчетов GPMC в формате HTML"
Write-Log -Message "Количество доменов, из которых будут получены отчеты: $($Domains.Count)"
$Domains.GetEnumerator()| Foreach-Object {
    $NETBIOSdomainName = $_.Key.Tostring()
    $DNSdomainName = $_.Value.Tostring()
    Trap {Write-Log -Message "Ошибка: $($NETBIOSdomainName) $_";Continue}
	$GPOsToGet = $null
	$GPOsToGet = Get-SDMgpo -Name * -DomainName $DNSdomainName
	If ($?){
		Write-Log -Message "Получение политик из домена $($NETBIOSdomainName): $($GPOsToGet.Count)"
		$GPOsToGet | Foreach-Object {
        	Write-Log -Message "Получение отчета по политике: $($_.Name.Tostring())"
			$ReportName = $PathToReports + $NETBIOSdomainName + "\"+ $_.Name.Tostring() +".html"
        	$CurrentReport = $_.Name.Tostring()
        	Trap {Write-Log -Message "Ошибка: Не удалось получить отчет по политике $CurrentReport из домена $($NETBIOSdomainName) - $_ ";Continue}
			Out-SDMgpsettingsreport -Name $_.Name -FileName $ReportName -ReportHTML -DomainName $DNSdomainName -ErrorAction Stop
    	}
	}
}
$ReportsInFolderAfterGet = Get-Childitem $PathToReports -Filter *.html -Recurse
Write-Log -Message "Количество отчетов в папке $($PathToReports): $($ReportsInFolderAfterGet.Count)"
Write-Log -Message "Выполнение скрипта Parse-GPOReports.ps1"
$ToRun = "& '.\Parse-GPOReports.ps1'"
Invoke-Expression -Command $ToRun
Write-Log -Message "Выполнение скрипта Rotate-Logs.ps1"
$ToRun = "& '.\Rotate-Logs.ps1'"
Invoke-Expression -Command $ToRun
Write-Log -Message "==============================Завершение работы скрипта=========================="
