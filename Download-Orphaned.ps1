# CAUTION:) Just example, cyrillic comments, tech debts, PScore7 required
<#
Экспорт документов GSuite в форматы MS Office (часть кода уже есть, для Excel файлов)
Скачивание документов GSuite альтернативным способом для обхода ограничения в 10МБ (с помощью создания прямых ссылок)?
Фильтрация объектов, которые нельзя экспортировать (некоторые файлы GSuite, временные файлы с именем "~*")
Многопоточность (отдельные Job по пользователям)
Логирование
Валидация по контрольным суммам
Фильтрация недопустимых символов в именах папок и файлов
Исправление - элементы пути, содержащие символ-разделитель "\" - "Фамилия\имя"
Исправление - имена файлов без расширения в имени, но с известным расширением в атрибуте ItemFileExtension и типом ItemMimeType

MS Sharepoint Naming remark:
Unsupported files and characters
We automatically process file and folder names to ensure they are accepted by Microsoft 365:
Files larger than 15 GB are not migrated.
Files with a size of 0 bytes (zero-byte files) are not migrated.
The following characters in file or folder names are removed: " * : < > ? / \ |
Leading tildes (~) are removed.
Leading or trailing whitespace is removed.
Leading or trailing periods (.) are removed.
See Invalid file names and file types for all other limitations.
In some possible circumstances with older sites, any file or folder ending in _files could fail. If you experience these errors, contact Support.
Microsoft currently has no file type limitations, meaning you can upload data with any file extension. For more info, see Types of files that cannot be added to a list or library
Character limits for files and folders
Filenames can have up to 256 characters.
Folder names may have up to 250 characters.
Total path length for folder and filename combinations can have up to 400 characters.
#>
# Параметры:
$timestamp = [int](Get-Date -UFormat %s)

$Code = {
    param ($sub)
    $logPath = "F:\Logs\" # Директория верхнего уровня
    $targetRootDirectory = "O:\.Orphaned" # Директория верхнего уровня, содержащая личные директории пользователей
    #$timestamp = [int](Get-Date -UFormat %s)
    $user = $sub # Имя пользователя (имя каталога пользователя, в который будут синхронизироваться файлы) 
    # Вычисляемые параметры:
    $date = Get-Date
    $logFilename = "DownloadOrphaned-{3}-{0}{1:d2}{2:d2}.log" -f $date.year, $date.month, $date.day, $sub
    $logFilePath = Join-Path -Path $logPath -ChildPath $logFilename
    $loggingPreference = "Continue"

    $issuer = 'gdrive-audit-account@gdrive-project-346546.iam.gserviceaccount.com'
    $key = Get-Content ".\jwt\gdrive-audit.pem" -AsByteStream
    $scope = "https://www.googleapis.com/auth/drive.readonly"
    $sqlServerName = "sql.example.com"


    Function Write-Log {
        # Назначение: 
        # 	Ведение журнала
        [cmdletbinding()]
        Param(
            [Parameter(Position = 0)]
            [ValidateNotNullOrEmpty()]
            [string]$Message
        )
        # Передача сообщения для Write-Verbose, если -Verbose был обнаружен
        Write-Verbose -Message $Message
        # Записывает в журнал, если переменная $LoggingPreference установлена в Continue
        if ($loggingPreference -eq "Continue") {
            $timeStamp = Get-Date -format "yyyy-MM-dd HH:mm:ss"
            Write-Output "$timeStamp $message" | Out-File -FilePath $logFilePath -Append -Encoding UTF8
        }
    }

    function Request-Token {
        # Входные параметры Issuer, Key, Sub, Scope
        param (
            [Parameter(Mandatory = $true)]
            [string]$issuer,
            [Parameter(Mandatory = $true)]
            [array]$key,
            [Parameter(Mandatory = $true)]
            [string]$sub,
            [Parameter(Mandatory = $true)]
            [string]$scope
        )
        $iat = [int](Get-Date -UFormat %s)  # The time the assertion was issued, specified as seconds since 00:00:00 UTC, January 1, 1970.
        $exp = $iat + 3600                  # The expiration time of the assertion, specified as seconds since 00:00:00 UTC, January 1, 1970. This value has a maximum of 1 hour after the issued time.
        $Global:expiredTimestamp = $exp #- 3580 debug
        $payloadClaims = @{
            sub   = $sub
            scope = $scope
            aud   = "https://oauth2.googleapis.com/token"
            iat   = $iat
        }
        $jwt = New-JWT -Algorithm 'RS256' -Type 'JWT' -Issuer $issuer -SecretKey $key -ExpiryTimestamp $exp -PayloadClaims $payloadClaims
        $grantType = "urn:ietf:params:oauth:grant-type:jwt-bearer"
        $requestUri = "https://oauth2.googleapis.com/token"
        $requestBody = "grant_type=$grantType&assertion=$jwt"

        $requestResponse = Invoke-RestMethod -Method Post -Uri $requestUri -ContentType "application/x-www-form-urlencoded" -Body $requestBody -SkipHttpErrorCheck
        if ($requestResponse.error) {
            Write-Log -Message "Error requesting access_token: $($requestResponse.error) $($requestResponse.error_description)"
            Write-Log -Message "Script stopped (Error)"
            Write-Output "unsuccessful"
            Exit
        }
        if ($requestResponse.access_token) {
            Write-Log -Message "Token received: $($requestResponse.access_token)"
            $secureStringAccessToken = ConvertTo-SecureString $requestResponse.access_token -AsPlainText -Force
        }
        # Cleaning vars  
        Remove-Variable -Name "iat"
        Remove-Variable -Name "exp"
        Remove-Variable -Name "payloadClaims"
        Remove-Variable -Name "jwt"
        Remove-Variable -Name "grantType"
        Remove-Variable -Name "requestUri"
        Remove-Variable -Name "requestBody"
        Remove-Variable -Name "requestResponse"
        return $secureStringAccessToken
    }

    function Get-GDriveItemsList {
        param (
            [int]$timestamp = [int](Get-Date -UFormat %s)
        )
        # получение списка файлов пользователя
        $sqlParametersUserId = "UserId='$($user)'"
        $itemsOrphaned = Invoke-Sqlcmd -Query "SELECT * FROM [GDriveMigration].[dbo].[GDriveReports] WHERE [ItemParents] IS NULL AND [ItemOwned] = 1 AND [ItemTrashed] = 0 AND NOT (ItemNamePath = N'Мой Диск' OR ItemNamePath = N'My Drive') AND UserId = `$(UserId) AND ItemMimeType != 'application/vnd.google-apps.folder' ORDER BY ItemNamePath" -ServerInstance $sqlServerName -Variable $sqlParametersUserId
        $itemsCount = $itemsOrphaned.Count
        Write-Log -Message "Получен список файлов из БД: $itemsCount"
        # запуск скачивания списка файлов
        Save-GDriveItems -GDriveItems $itemsOrphaned
    }

    function Repair-Name {
        param (
            [string]$itemName
        )
        # очистка имени файла
        $itemName = $itemName.Split([IO.Path]::GetInvalidFileNameChars()) -join ''
        $itemNamePrevious = $itemName #prev
        do {
            $itemNameCleaned = $itemNamePrevious.Trim().Trim(".") # next
            if ($itemNamePrevious.Length -eq $itemNameCleaned.Length) {
                break
            }
            else {
                $itemNamePrevious = $itemNameCleaned
            }
        } while ($true)
        return $itemNameCleaned
    }

    function Save-GDriveItems {
        param (
            $GDriveItems
        )
        $currentItem = 0
        $itemsCount = $itemsOrphaned.Count
        foreach ($GDriveItem in $GDriveItems) {
            # формирование полного пути к файлу
            $currentItem += 1
            $itemId = $GDriveItem.ItemId
            $itemName = $GDriveItem.ItemName
            # Логирование и пропуск временных файлов (начинаются с символа "~")
            #$itemName = "~My temp file.tmp"
            if ($itemName.StartsWith("~")) {
                Write-Log -Message "[$currentItem/$itemsCount] SKIP: пропуск временного файла: $itemName"
                continue
            }
            $itemNameCleaned = Repair-Name -itemName $itemName
            #Write-Log -Message "Нормализованный путь и имя файла: $itemNamePathCleaned"
            $userId = $GDriveItem.UserId
            $mimeType = $GDriveItem.ItemMimeType
            $specialFolder = ".Orphaned" # имя специальной папки для файлов
            $itemPath = Join-Path -Path $targetRootDirectory -ChildPath $userId -AdditionalChildPath $specialFolder, $itemNameCleaned
            #Write-Log -Message "Полный путь и имя сохраняемого файла: $itemPath"
            $itemParentPath = Split-Path -Path $itemPath # путь к файлу, не включая имя
            # создаём полный путь, если отсутствует
            if (!(Test-Path $itemParentPath)) {
                New-Item -ItemType Directory -Force -Path $itemParentPath | Out-Null
                #Write-Log -Message "Создание несуществующего пути файла: $itemParentPath"
            }
            
            # Формирование URL и имени файла в зависимости от типа файлов
            # Экспортируются только файлы:
            # G document --> docx
            # G spreadsheet --> xlsx
            # G presentation --> pptx
            # G drawing --> svg
            # G app script --> json
            # Файлы, для которых невозможен экспорт, логируются (G form, G map, G shortcut)
            # Остальные файлы скачиваются обычным запросом
            switch ($mimeType) {
                "application/vnd.google-apps.document" {
                    #docx application/vnd.openxmlformats-officedocument.wordprocessingml.document
                    if (!$itemPath.EndsWith(".docx")) {
                        $itemPath = $itemPath + ".docx"
                        Write-Log -Message "EXPORT: .docx"
                    }
                    $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
                "application/vnd.google-apps.spreadsheet" { 
                    #xslx application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                    if (!$itemPath.EndsWith(".xlsx")) {
                        $itemPath = $itemPath + ".xlsx"
                        Write-Log -Message "EXPORT: .xlsx"
                    }
                    $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
                "application/vnd.google-apps.presentation" {
                    #pptx application/vnd.openxmlformats-officedocument.presentationml.presentation
                    if (!$itemPath.EndsWith(".pptx")) {
                        $itemPath = $itemPath + ".pptx"
                        Write-Log -Message "EXPORT: .pptx"
                    }
                    $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "/export?mimeType=application/vnd.openxmlformats-officedocument.presentationml.presentation"
                }
                "application/vnd.google-apps.drawing" {
                    #svg image/svg+xml
                    if (!$itemPath.EndsWith(".svg")) {
                        $itemPath = $itemPath + ".svg"
                        Write-Log -Message "EXPORT: .svg"
                    }
                    $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "/export?mimeType=image/svg+xml"
                }
                "application/vnd.google-apps.script" {
                    #json application/vnd.google-apps.script+json
                    if (!$itemPath.EndsWith(".json")) {
                        $itemPath = $itemPath + ".json"
                        Write-Log -Message "EXPORT: .json"
                    }
                    $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "/export?mimeType=application/vnd.google-apps.script+json"
                }
                "application/vnd.google-apps.form" {
                    Write-Log -Message "[$currentItem/$itemsCount] SKIP: пропуск файла Google Forms ($itemId): $itemPath"
                    continue
                }
                "application/vnd.google-apps.map" {
                    Write-Log -Message "[$currentItem/$itemsCount] SKIP: пропуск файла Google Maps ($itemId): $itemPath"
                    continue
                }
                "application/vnd.google-apps.shortcut" {
                    Write-Log -Message "[$currentItem/$itemsCount] SKIP: пропуск файла-ярлыка Google Shortcut ($itemId): $itemPath"
                    continue
                }
                Default { $fileContentURL = "https://www.googleapis.com/drive/v3/files/" + $itemId + "?alt=media" }
            }
            
            $currentTimestamp = [int](Get-Date -UFormat %s)
            if (($expiredTimestamp - $currentTimestamp) -lt 60) {
                Write-Log -Message "INFO: запрос нового токена"
                $accessToken = Request-Token -issuer $issuer -key $key -sub $sub -scope $scope
            }
            # проверка дубликата имени файла
            if (Test-Path $itemPath) {
                $itemExtension = Split-Path -Path $itemPath -Extension
                $itemFullPathAndName = Join-Path -Path $(Split-Path -Path $itemPath -Parent) -ChildPath $(Split-Path -Path $itemPath -LeafBase)
                $itemPath = $itemFullPathAndName + "_" + $itemId + $itemExtension
                Write-Log -Message "RESOLVE: обнаружен дубликат, добавлен ItemId к имени файла"
            }

            try {
                $requestResponse = Invoke-RestMethod -Uri $fileContentURL -Authentication OAuth -Token $accessToken -OutFile $itemPath #-SkipHttpErrorCheck 
                if ($requestResponse.error) {
                    Write-Log -Message "ERROR: $($requestResponse.error) $($requestResponse.error_description)"
                    Write-Log -Message "Script stopped (Error)"
                    Write-Output "unsuccessful"
                    Exit
                }
                Write-Log -Message "[$currentItem/$itemsCount] SUCCESS: скачан файл ($itemId):$itemPath"

            }
            catch {
                Write-Log -Message "[$currentItem/$itemsCount] ERROR: не удалось скачать файл ($itemId):$itemPath"
            }
        }
    }
    Write-Log -Message "Скрипт запущен"
    $accessToken = Request-Token -issuer $issuer -key $key -sub $sub -scope $scope
    Get-GDriveItemsList
    Remove-Variable -Name "accessToken"
    Write-Log -Message "Скрипт остановлен (Выполнен)"
}

<#
#$InvalidPathChars = [IO.Path]::GetInvalidPathChars()
# |╔╗╚╝║═  
#$InvalidFileNameChars = [IO.Path]::GetInvalidFileNameChars()
# "<>|╔╗╚╝║═*?\/
#The following characters in file or folder names are removed: " * : < > ? / \ | GetInvalidFileNameChars()
#Leading tildes (~) are removed Skip/Log
#Leading or trailing whitespace is removed Trim()
#Leading or trailing periods (.) are removed Trim(".")

$itemNamePath = "Мой диск  .  .\.ТЕЛЕ?*??ФОНЫ .. .  .\=5F27.02.2020?=\                         ..   ..  ..                         Apple iPhone XR 256GB black, "
$itemName = "                         ..   ..  ..                         Apple iPhone XR 256GB black, "
Clean-NamePath -itemNamePath $itemNamePath -itemName $itemName
$itemNamePath.Length
#$test = "Мой ди**ск\ТЕЛЕФ???ОНЫ\                                                    Apple iPhone XR 256GB black, "
#$testArray = $test.Split("\")
#$testArray.Count
#>

# preparing jobs
$users = Invoke-Sqlcmd -Query "SELECT [UserId] FROM [GDriveMigration].[dbo].[Progress] WHERE [DownloadOrphaned] = 0 ORDER BY [UserId]" -ServerInstance $sqlServerName
$users | ForEach-Object {
    Start-ThreadJob -Name $_.UserId -ArgumentList $_.UserId -ScriptBlock $Code -ThrottleLimit 10
} | Out-Null
