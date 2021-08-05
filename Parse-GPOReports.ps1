###########################################################################
# Имя скрипта:
# 	Parse-GPOReports.ps1
# Назначение:
# 	Разбор HTML-отчетов по групповым политикам, выгруженным из AD и заливка в БД
# Описание:
#	Скрипт рекурсивно обращается к файлам выгруженных отчетов и получает необходимые данные с помощью:
#		Обращения к конкретным параметрам с помощью XPath(XML Path Language)
#		Рекурсивного обхода дерева XML Document Object Model (DOM) и формирования списка Policy/Setting
#	Полученные данные объединяются в промежуточных файлах и импортируются в соответствующие 
#	таблицы базы данных с помощью команды BULK INSERT.
#	В процессе работы скрипта ведется журнал.
###########################################################################

########################################################################### 
# Параметры и переменные
###########################################################################
# Путь к папке с отчетами
$PathToReports = "C:\GPO\"
# Промежуточный файл с основными параметрами политик
$PathToGPOGeneralDetailsADCSV = "C:\GPO\GPO-GeneralDetails.csv"
# Промежуточный файл с настройками политик
$PathToGPOParametersADCSV = "C:\GPO\GPO-Parameters.csv"
# Путь к базе данных
$PathToDB = "SQLSERVER:\SQL\localhost\SQLEXPRESS\Databases\DBNAME"
# Параметры журнала
$Date = Get-Date
$LogFilename = "Parse-GPOReports{0}{1:d2}{2:d2}" -f $Date.year,$Date.month,$Date.day
$LoggingPreference="Continue"
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
    [string]$LogFilePath="C:\GPO\Logs\$LogFilename.txt"
    )
    # Передача сообщения для Write-Verbose, если -Verbose был обнаружен
    Write-Verbose -Message $Message
    # Записывает в журнал, если переменная $LoggingPreference установлена в Continue
    if ($LoggingPreference -eq "Continue"){
        $TimeStamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
        Write-Output "$TimeStamp $Message" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
}

Function parseGPOGeneralDetails {
# Назначение: 
# 	Парсинг основных параметров групповых политик и сохранение в промежуточный файл
	param($filename)
	[xml]$xml = Get-Content $filename
	$ID = [GUID]$xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[7]/td[2]").InnerText # ID
	$Name = $xml.SelectSingleNode("//td[@class='gponame']").InnerText # Name
	$Domain = $xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[1]/td[2]").InnerText # Domain
	$Created = Get-Date ($xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[3]/td[2]").InnerText) # Created
	$Modified = Get-Date ($xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[4]/td[2]").InnerText) # Modified
	$Collected = Get-Date ($xml.SelectSingleNode("//table[@class='title']/tr[2]/td[1]").InnerText -replace "Data collected on: ", "") # Collected
	$UserRev = $xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[5]/td[2]").InnerText # UserRev
	$ComputerRev = $xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[6]/td[2]").InnerText # ComputerRev
	$Status = $xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[8]/td[2]").InnerText # Status
	#($format in "N","D","B","P") - типы форматов для GUID
	$Rows = $ID.ToString("D").ToLower() + "|" + $Name + "|" + $Domain + "|" + $Created + "|" + $Modified + "|" + $Collected + "|" + $UserRev + "|" + $ComputerRev + "|" + $Status + "|1"
	$Details = $Details + $Rows
	$Details | Out-File -FilePath $PathToGPOGeneralDetailsADCSV -Append
}

Function parsePath {
# Назначение: 
# 	Парсинг настроек GPO и сохранение в промежуточный файл
	param($node)
	$GPOparams = $node.OuterXML
	$ID = [GUID]$xml.SelectSingleNode("//div[@class='gposummary']/div[@class='container']/div/div/*/tr[7]/td[2]").InnerText # ID
	$Collected = Get-Date ($xml.SelectSingleNode("//table[@class='title']/tr[2]/td[1]").InnerText -replace "Data collected on: ", "") # Collected
	$currentNode = $node.ParentNode
	for ($i = 0; $i -lt 10; $i++) {
    	if ($currentNode.ParentNode.NodeType.value__ -ne 9 ) {
			if ($currentNode.ParentNode.PreviousSibling.name -ne "head"){
				switch ($currentNode.class){
					he4i {$path = $currentNode.ParentNode.PreviousSibling.InnerText}
					container {
						if ($currentNode.ParentNode.PreviousSibling.class -ne "title"){
							$path = $currentNode.ParentNode.PreviousSibling.InnerText + ">" + $path
						}
					}

				}
			}
			$currentNode = $currentNode.ParentNode
		}
    }
	$Rows = $ID.ToString("D").ToLower() + "||" + $Collected + "||" + $path  + "||" + $GPOparams + "||1"
	$Details = $Details + $Rows
	$Details | Out-File -FilePath $PathToGPOParametersADCSV -Append
}

Function parseGPOParameters {
# Назначение: 
# 	Получение всех таблиц и обработка
	param($filename)
	[xml]$xml = Get-Content $filename
	# Замена гиперссылок со скриптами на текст
	$Explains = $xml.SelectNodes("//a[@class='explainlink']") # поиск ссылок <a class="explainlink">text</a>
	$Explains | Foreach-Object {
		$Explain = $_ # текущая ссылка 
		$ExplainNew = $Explain.FirstChild # текст текущей ссылки
		#$Explain.ParentNode.RemoveChild($Explain)
		$Explain.ParentNode.ReplaceChild($ExplainNew, $Explain) # замена
	}
	$Tables = $xml.SelectNodes("//table")
	$Tables | Foreach-Object {
		$TableXML = $_
		$TableClass = $_.Attributes.GetNamedItem("class").Value
		#$TableClass
		switch ($TableClass){
			title {}
			info {parsePath($TableXML)}
			#info3 {$TableXML.OuterXML}
			info3 {parsePath($TableXML)}
			subtable {if ($TableXML.ParentNode.Name -ne "td"){parsePath($TableXML)}} # Если таблица вложенная - не выводить, так как она будет выведена из родителя
			subtable3 {if ($TableXML.ParentNode.Name -ne "td"){parsePath($TableXML)}} # Если таблица вложенная - не выводить, так как она будет выведена из родителя
			subtable_frame {if ($TableXML.ParentNode.Name -ne "td"){parsePath($TableXML)}} # Если таблица вложенная - не выводить, так как она будет выведена из родителя
			default {Write-Host "Default"}
		}
	}
}	
Write-Log -Message "Начало работы скрипта"
Write-Log -Message "Удаление промежуточного файла $PathToGPOParametersADCSV"
# Удаление старого и создание нового файла CSV 
Trap {Write-Log -Message "Ошибка: $_";Continue}
Remove-Item $PathToGPOParametersADCSV -ErrorAction Stop
Write-Log -Message "Создание нового промежуточного файла $PathToGPOParametersADCSV"
New-Item $PathToGPOParametersADCSV -type File -ErrorAction Stop
Write-Log -Message "Удаление промежуточного файла $PathToGPOGeneralDetailsADCSV"
Remove-Item $PathToGPOGeneralDetailsADCSV -ErrorAction Stop
Write-Log -Message "Создание нового промежуточного файла $PathToGPOGeneralDetailsADCSV"
New-Item $PathToGPOGeneralDetailsADCSV -type File -ErrorAction Stop

# Парсинг общих параметров 
Get-Childitem $PathToReports -Filter *.html -Recurse | Foreach-Object {
	Trap {Write-Log -Message "Ошибка: $_";Continue}
	Write-Log -Message "Обработка отчета: $($_.FullName)"
	parseGPOGeneralDetails $_.FullName -ErrorAction Stop
	parseGPOParameters $_.FullName -ErrorAction Stop
}

# Заливка в БД
Trap {Write-Log -Message "Ошибка: $_";Continue}
Write-Log -Message "Загрузка данных в БД"
$ToRun = "& '.\Initialize-SQLProvider.ps1'"
Invoke-Expression -Command $ToRun -ErrorAction Stop
Set-Location $PathToDB -ErrorAction Stop
Write-Log -Message "Очистка признака текущих политик в таблице GPOGeneralDetails"
Invoke-Sqlcmd -Query "UPDATE GPOGeneralDetails SET Last = '0' WHERE Last = '1'" -ErrorAction Stop
Write-Log -Message "Загрузка текущих политик в таблицу GPOGeneralDetails"
Invoke-Sqlcmd -Query "BULK INSERT GPOGeneralDetails FROM 'C:\GPO\GPO-GeneralDetails.csv' WITH (FIELDTERMINATOR = '|', ROWTERMINATOR = '\n', DATAFILETYPE = 'widechar', FIRSTROW = 0)" -ErrorAction Stop
Write-Log -Message "Очистка признака текущих политик в таблице GPOParametersAD"
Invoke-Sqlcmd -Query "UPDATE GPOParameters SET Last = '0' WHERE Last = '1'" -ConnectionTimeout 60 -QueryTimeout 300 -ErrorAction Stop
Write-Log -Message "Загрузка текущих политик в таблицу GPOParameters"
Invoke-Sqlcmd -Query "BULK INSERT GPOParameters FROM 'C:\GPO\GPO-Parameters.csv' WITH (FIELDTERMINATOR = '||', ROWTERMINATOR = '\n', DATAFILETYPE = 'widechar', FIRSTROW = 0)" -ConnectionTimeout 60 -QueryTimeout 300 -ErrorAction Stop
Write-Log -Message "Загрузка данных в БД завершена"
Write-Log -Message "Завершение работы скрипта"
