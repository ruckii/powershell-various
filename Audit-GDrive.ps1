# CAUTION:) Just example, cyrillic comments, tech debts, PScore7 required
# Получение метаданных файлов и работа с Google Drive API (G Suite)
# Для работы скрипта требуется настройка доступа к Google Drive API 
# с помощью Service Account делегированными правами Domain-Wide:
# https://developers.google.com/identity/protocols/oauth2/service-account
# В результате настройки должны быть получены и указаны в функции Auth следующие параметры (пример):
# $Issuer = 'gdrive-audit-account@gdrive-project-20210510.iam.gserviceaccount.com'
# $rsaPrivateKey = Get-Content "./jwt/gdrive-audit.pem" -AsByteStream
# sub   = "user1@example.com"
# scope = "https://www.googleapis.com/auth/drive.readonly"
# aud   = "https://oauth2.googleapis.com/token"
# С помощью данных параметров в функции Auth производится получение токена доступа, 
# который используется в последующих вызовах сервиса Google Drive API. Токен имеет срок действия, 
# при истечении которого необходимо повторно вызвать функцию Auth для получения нового токена.
# Это нужно будет автоматизировать (отслеживать срок действия токена или отлавливать ошибку авторизации).
# По умолчанию время жизни токена 1 час: "expires_in": 3600

<#
Drive API, v3 Scopes
https://www.googleapis.com/auth/drive	See, edit, create, and delete all of your Google Drive files
https://www.googleapis.com/auth/drive.appdata	View and manage its own configuration data in your Google Drive
https://www.googleapis.com/auth/drive.file	View and manage Google Drive files and folders that you have opened or created with this app
https://www.googleapis.com/auth/drive.metadata	View and manage metadata of files in your Google Drive
https://www.googleapis.com/auth/drive.metadata.readonly	View metadata for files in your Google Drive
https://www.googleapis.com/auth/drive.photos.readonly	View the photos, videos and albums in your Google Photos
https://www.googleapis.com/auth/drive.readonly	See and download all your Google Drive files
https://www.googleapis.com/auth/drive.scripts	Modify your Google Apps Script scripts' behavior
#>

Import-Module 'powershell-jwt'

# Параметры:
$sqlServerName = "sql.example.com"
$databaseName = "GDrive-OneDrive"
$tableName = "GDriveAudit"
$issuer = 'gdrive-audit-account@gdrive-project-8753764.iam.gserviceaccount.com'
$key = Get-Content ".\jwt\gdrive-audit.pem" -AsByteStream
$sub = "user1@example.com"
$scope = "https://www.googleapis.com/auth/drive.readonly"

# Функции:
Function Get-ExpiryTimeStamp {
    param (
        [int]$ValidForSeconds = 3600
    )
    $exp = [int](Get-Date -UFormat %s) + $ValidForSeconds # Grab Unix Epoch Timestamp and add desired expiration.
    $exp
}

# Функция аутентификации (запрос токена доступа)
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
    $requestResponse = Invoke-RestMethod -Method Post -Uri $requestUri -ContentType "application/x-www-form-urlencoded" -Body $requestBody
    $accessToken = $requestResponse.access_token
    $headers = @{"Authorization" = "Bearer $accessToken" }
    return $headers
}

function Test-Token {
    param (
        [int]$exp
    )
    $currentTimestamp = [int](Get-Date -UFormat %s)
    $remainingSeconds = ($exp -le $currentTimestamp) ? 0 : $exp - $currentTimestamp
    if ($remainingSeconds -eq 0) {
        Write-Host "Remaining [seconds]: 0"
        return $false
    }
    else {
        Write-Host "Remaining [seconds]: " $remainingSeconds
        return $true
    }
}

# Функция рекурсивного получения списка директорий 
# Если список директорий нельзя выгрузить одной командой - используется пейджинг (постраничная отдача результата).
# Признак окончания выгрузки - отсутствие в очередной отданной странице параметра nextPageToken
function Get-GDriveItemsMetadata {
    param (
        [string]$Token,
        [array]$Items
    )
    # При первом запуске (отсутствуют входные параметры) 
    $URL = ""
    if ($null -eq $Items) {
        $newItems = @()
    }
    # формирование запроса
    if ($Token) {
        # если есть ещё элементы - для их запроса используется соответствующий параметр pageToken
        $URL = "https://www.googleapis.com/drive/v3/files?fields=*&pageSize=1000&pageToken=" + $Token
    }
    else {
        # если первый запуск (отсутствует pageToken)
        $URL = "https://www.googleapis.com/drive/v3/files?fields=*&pageSize=1000"
    }
    #выполнить запрос
    $answer = Invoke-RestMethod -Uri $URL -Method Get -Headers $headers
    $answer.files.Count
    Write-Host "Load: $(Get-Date)"
    #добавить к массиву результат
    $newItems = $answer.files
    $newToken = $answer.nextPageToken
    $allItems = $allItems + $newItems
    if ($null -eq $newToken) {
        # если в очередном ответе отсутствует nextPageToken - конец
        return $allItems
    }
    Get-GDriveItemsMetadata -Token $newToken -Items $newItems # рекурсивный вызов
}

Function Write-ToMSSQL {
    param (
        [Parameter(Mandatory = $true)]
        [array]$GDriveItems
    )
    for ($i = 0; $i -lt $GDriveItems.Count; $i += 1000) {
        Write-Host "Loop: $i"
        $sqlQuery = "INSERT INTO [dbo].[$tableName] (metadata) Values (N'{0}')" -f (($GDriveItems[$i..($i + 999)] | ForEach-Object { $_ | ConvertTo-Json -EscapeHandling EscapeHtml -Compress }) -join "')`r`n,(N'")
        Invoke-Sqlcmd -Database $databaseName -Query $sqlQuery -ServerInstance $sqlServerName -ErrorAction Stop
    }
}

# Вызов функций:
# Получение токена
$headers = Request-Token -issuer $issuer -key $key -sub $sub -scope $scope
# Вывод токена авторизации
$headers.Authorization

# Получение списка директорий
$GDriveItems = Get-GDriveItemsMetadata
#$GDriveItems.Count
#$GDriveItems.GetType()
# Запись в БД
Write-ToMSSQL -GDriveItems $GDriveItems

<# Jobs example
$Code = {
    param ($sub)
    $headers = $Using:Request-Token -issuer $issuer -key $key -sub $sub -scope $scope
    Write-Output "User: $sub Header: $($headers.Authorization)"
}

$jobs = @()
("user1@example.com","user2@example.com","user3@example.com") | ForEach-Object { $jobs += Start-Job -ArgumentList $_ -ScriptBlock $Code }
#>
