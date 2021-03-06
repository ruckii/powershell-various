###########################################################################
# Имя скрипта:
# 	Parse-SecurityLogs.ps1
# Назначение:
# 	Обработка и заливка в базу данных архивных журналов событий.
# Описание работы:
#	1. Скрипт рекурсивно обращается к архивным файлам журналов событий и преобразует их в формат XML с помощью утилиты WEVTUTIL.
#	2. Каждый XML-файл обрабатывается, нужные данные сохраняются в файле, пригодном для импорта в БД с помощью утилиты BCP.
#	3. После обработки всех архивных журналов - данные загружаются в БД.
#	4. Удаляются временные файлы. 
#	5. В процессе работы скрипта ведется журнал.
###########################################################################

########################################################################### 
# Параметры и переменные
###########################################################################
$Date = Get-Date
# Путь к папке с журналами событий
$PathToEventlogs = "C:\Windows\System32\winevt\Logs\"
# Путь к папке с обработанными журналами событий
$PathToRenderedEventlogs = "C:\Temp\"
# Промежуточный файл событий
$PathToEventsCSV = "C:\Temp\EventsRendered{0}{1:d2}{2:d2}{3:d2}.csv" -f $Date.Year,$Date.Month,$Date.Day,$Date.Hour
# Параметры журнала
$LogFilename = "Parse-SecurityLogs{0}{1:d2}{2:d2}.txt" -f $Date.Year,$Date.Month,$Date.Day
$LoggingPreference="Continue"
# Путь к журналам работы скрипта
$PathToLogs = "C:\Logs\"
# Удалять файлы старше чем:
$OlderThanDays = 30
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
    [string]$LogFilePath="C:\Logs\$LogFilename"
    )
	# Передача сообщения для Write-Verbose, если -Verbose был обнаружен
    Write-Verbose -Message $Message
    # Записывает в журнал, если переменная $LoggingPreference установлена в Continue
    if ($LoggingPreference -eq "Continue"){
        $TimeStamp = Get-Date -format "yyyy.MM.dd HH:mm:ss"
        Write-Output "$TimeStamp $Message" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
}

Function Delete-OldLogs {
# Назначение: 
# 	Удаление журналов работы старше задаваемого периода времени
	Write-Log -Message "Начало удаления старых журналов из каталога $PathToLogs"
	$FilesToDelete = Get-Childitem $PathToLogs -Filter *.txt | Where-Object {$_.CreationTime -lt (Get-Date).AddDays(-$OlderThanDays)}
	if ($FilesToDelete.Count -gt 0){
    	Write-Log -Message "Количество журналов для удаления: $($FilesToDelete.Count)"
		$FilesToDelete | Foreach {
			Write-Log -Message "Удаление журнала: $($_.FullName)"
			Trap {Write-Log -Message "Ошибка: $_";Continue}
			Remove-Item $_.FullName -ErrorAction Stop 
		}   
    } else {
		Write-Log -Message "Нет журналов для удаления."
	}
}
########################################################################### 
# Основной код
###########################################################################
Write-Log -Message "==============================Начало работы скрипта=============================="
# Парсинг журналов событий
[String]$ElapsedTime = "---"
$Archive = Get-Childitem $PathToEventlogs -Filter "Archive-ForwardedEvents-*.evtx" 
$Archive | Foreach-Object {
	Trap {Write-Log -Message "Ошибка: $_";Continue}
	Write-Log -Message "Обработка журнала [$($i) из $($Archive.Count) | $($ElapsedTime) мин]: $($_.FullName)"
	$Progress = Measure-Command {
	# переименование в XML и генерация пути преобразованного файла
	$CurrentEventlog = $_.FullName
	$RenderedEventlog = Join-Path -Path $PathToRenderedEventlogs -ChildPath $_.Name
	$RenderedEventlog = [System.IO.Path]::ChangeExtension($RenderedEventlog, "xml")
	Start-Process 'cmd' -ArgumentList "/c wevtutil /F:RenderedXML qe /lf:true ""$CurrentEventlog"" > ""$RenderedEventlog""" -Wait -NoNewWindow -ErrorAction Stop
	[xml]$xml = "<Events>" + (Get-Content $RenderedEventlog) + "</Events>"
	$Rows = New-Object Text.StringBuilder
	$Events = $xml.Events.Event
	$Events | Foreach-Object {
		$EventRecordID = $_.System.EventRecordId
		$EventID = $_.System.EventId.InnerText
		$LevelID = $_.System.Level
		$LevelName = $_.RenderingInfo.Level
		$TaskID = $_.System.Task
		$TaskName = $_.RenderingInfo.Task
		$UserID = $_.System.Security.UserID
		$ProviderName = $_.System.Provider.Name
		$Channel = $_.System.Channel
		$Computer = $_.System.Computer
		[Datetime]$TimeCreated = $_.System.TimeCreated.SystemTime
		$Keywords = $_.RenderingInfo.Keywords.Keyword -join ","
		$Message = $_.RenderingInfo.Message
		$Data = $_.EventData.Data -join " "
		[Void]$Rows.append("|")
		[Void]$Rows.Append($EventRecordID)
		[Void]$Rows.append("|")
		[Void]$Rows.append($EventID)
		[Void]$Rows.append("|")
		[Void]$Rows.append($LevelID)
		[Void]$Rows.append("|")
		[Void]$Rows.append($LevelName)
		[Void]$Rows.append("|")
		[Void]$Rows.append($TaskID)
		[Void]$Rows.append("|")
		[Void]$Rows.append($TaskName)
		[Void]$Rows.append("|")
		[Void]$Rows.append($UserID)
		[Void]$Rows.append("|")
		[Void]$Rows.append($ProviderName)
		[Void]$Rows.append("|")
		[Void]$Rows.append($Channel)
		[Void]$Rows.append("|")
		[Void]$Rows.append($Computer)
		[Void]$Rows.append("|")
		[Void]$Rows.append([String]$TimeCreated)
		[Void]$Rows.append("|")
		[Void]$Rows.append($Keywords)
		[Void]$Rows.append("|")
		[Void]$Rows.append($Message)
		[Void]$Rows.append("|")
		[Void]$Rows.append($Data)
		[Void]$Rows.append("#").AppendLine()
	}
	$Rows.ToString(0,$Rows.Length-2) | Out-File -FilePath $PathToEventsCSV -Append -ErrorAction Stop # удаление лишнего перевода строки
	$i++ # счетчик файлов
	}
	[Int]$ElapsedTime = $Progress.TotalMinutes*($Archive.Count-$i)
	Remove-Item $CurrentEventlog -Force -ErrorAction Stop
	Remove-Item $RenderedEventlog -Force -ErrorAction Stop
}
Trap {Write-Log -Message "Ошибка: $_";Continue}
Write-Log -Message "Начало загрузки данных в БД"	
Start-Process 'bcp.exe' -ArgumentList "DBNAME.dbo.Events in ""$PathToEventsCSV"" -T -t""|"" -w -r""#\n"" -b 10000 -h ""TABLOCK""" -Wait -RedirectStandardOutput "C:\Temp\bcp.log" -ErrorAction Stop
$BCP = Get-Content "C:\Temp\bcp.log"
Add-Content "C:\Logs\$LogFilename" $BCP -Encoding UTF8
if ($BCP -like '*rows per sec.)*')
	{
		Write-Log -Message "Загрузка файла $($PathToEventsCSV) в БД завершена"
		Remove-Item $PathToEventsCSV -ErrorAction Stop
	} 
	else {
		Write-Log -Message "Ошибка при загрузке в БД файла $($PathToEventsCSV)"
	}
Delete-OldLogs
Write-Log -Message "==============================Завершение работы скрипта=============================="
  