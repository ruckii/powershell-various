########################
# Provision-Mailboxes.ps1
########################
#  v.1.1 (19.01.2018)
########################

#################################################
# Параметры
#################################################
[String]$LogPath = "C:\Logs"                                                           # Путь к журналам работы скрипта
[String]$LogFile = "Provision-Mailboxes"                                               # Имя журнала работы
[system.URI]$CAS = "https://ems.example.com/ews/exchange.asmx"                         # EWS URL

###################################################################################################
# Функция логирования
###################################################################################################
Function Write-Log {
# Указание обязательных параметров
param
    (
    [Parameter(Mandatory=$true)][string]$Message,
    [Parameter(Mandatory=$false)][string]$Log,
    [Parameter(Mandatory=$false)][string]$Type,
    [Parameter(Mandatory=$false)][switch]$NoSilent
    )

# Формирование строки сообщения	
$CompleteMessage = "UTC: " + ('{0:dd.MM.yyyy HH:mm:ss}' -f (Get-Date).ToUniversalTime()) + "`t" + $Type.toLower() + "`t" + $Message
# Вывод сообщения в лог-файл
switch ($Type.toLower())
    {
    "info"
        {
        if ($NoSilent) {Write-Host $CompleteMessage}
        Out-File -FilePath $Log -InputObject $CompleteMessage -Append -Encoding unicode
        break
        }
    "error"
        {
        if ($NoSilent) {Write-Host $CompleteMessage -ForegroundColor Red}
        Out-File -FilePath $Log -InputObject $CompleteMessage -Append -Encoding unicode
        break
        }
    "warning"
        {
        if ($NoSilent) {Write-Host $CompleteMessage -ForegroundColor Yellow}
        Out-File -FilePath $Log -InputObject $CompleteMessage -Append -Encoding unicode
        break
        }
    "completed"
        {
        if ($NoSilent) {Write-Host $CompleteMessage -ForegroundColor Green}
        Out-File -FilePath $Log -InputObject $CompleteMessage -Append -Encoding unicode
        break
        }
    default
        {
        $Type = "info"
        $CompleteMessage = "UTC: " + ('{0:dd.MM.yyyy HH:mm:ss}' -f (Get-Date).ToUniversalTime()) + "`t" + $Type + "`t" + $Message
        if ($NoSilent) {Write-Host $CompleteMessage}
        Out-File -FilePath $Log -InputObject $CompleteMessage -Append -Encoding unicode
        }
    }
}

###################################################################################################
# Функция открытия лога
###################################################################################################
Function Start-Logging {
    # Проверка директории хранения лог-файлов
    if (-not(Test-Path $LogPath -PathType Container)) {New-Item $LogPath -ItemType Directory | Out-Null}
    # Имя лог-файла создается на основе текущей даты и времени
    $LogFile = $LogFile+"-"+(Get-Date -Format "yyyyMMddHHmm")+".log"
    $Script:Log = $LogPath+"`\"+$LogFile
    # Старт журналирования, каждый раз создается новый журнал
    Write-Log "**********************" -Log $Log
    Write-Log "Windows PowerShell Transcript Start" -Log $Log
    Write-Log ("Start time: " + (Get-Date -Format "dd.MM.yyyy HH:mm:ss")) -Log $Log
    Write-Log ("Username: " + $env:USERNAME) -Log $Log
    Write-Log ("Machine: " + $env:COMPUTERNAME + " `(" + [System.Environment]::OSVersion.VersionString + "`)") -Log $Log
    Write-Log "**********************" -Log $Log
    Write-Log ("Transcript started, output file is " + $Log) -Log $Log
    Write-Log "**********************" -Log $Log
}

###################################################################################################
# Функция закрытия лога
###################################################################################################
Function Stop-Logging {
    param
    (
    [Parameter(Mandatory=$true)][string]$Log
    )
    # Закрытие лог-файлов
    Write-log "**********************" -Log $Log
    Write-log "Windows PowerShell Transcript End" -Log $Log
    Write-log ("End time: " + (Get-Date -Format "dd.MM.yyyy HH:mm:ss")) -Log $Log
}

###################################################################################################
# Функция подключения к EWS
###################################################################################################
Function Connect-EWS {
    # Загрузка модуля EMC
    # Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

    # Подключение к EWS
    [Reflection.Assembly]::LoadFrom(".\Microsoft.Exchange.WebServices.dll")
    $Script:EWS = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    $EWS.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.WebCredentials("user","password","domain")
    #$EWS.UseDefaultCredentials = $true
    # CAS URL Option 1 Autodiscover
    #$EWS.AutodiscoverUrl("IvanovII@domain.org",{$true})
    # CAS URL Option 2 Hardcoded
    $EWS.Url = $CAS
    Write-Log ("Подключение к серверу Exchange (Exchange Web Services): " + $EWS.Url) -Log $Log
}

###################################################################################################
# Функция подключения к EMS (PS Remote)
###################################################################################################
Function Connect-EMS {
    # Загрузка модуля EMS
    # Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2016
    $pw = ConvertTo-SecureString -AsPlainText -Force -String "password"
    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "domain\user",$pw
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ems.example.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session
    Write-Log ("Подключение к серверу Exchange (Exchange Management Shell): http://ems.example.com/PowerShell/") -Log $Log
}


###################################################################################################
# Функция обработки писем и миграции (скрипт обрабатывает только один запрос за каждый запуск)
###################################################################################################
# 1. Получение списка писем (непрочитанные с темой [CGP])
# 2. Для первого письма в списке:
# 2.1 Получение адреса отправителя и сохранение его LegacyExchangeDN
# 2.2 Удаление ящика, присвоение SMTP-адреса
# 2.3 Присвоение X500-адреса = сохраненному ранее LegacyExchangeDN
# 3. Пометка запроса флагом завершения (для исключения его повторной обработки)
Function Start-Request {
# Подключение к почтовому ящику migrator
    $emailAddress = "user@example.com"
    Write-Log ("Подключение к почтовому ящику: " + $emailAddress) -Log $Log
    $mailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($emailAddress)
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $mailbox)
    $scanFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS,$folderId)
    Write-Log ("Подключение к папке: " + $scanFolder.DisplayName) -Log $Log
# 1. Получение списка писем (непрочитанные с темой [CGP])
    $searchQuery = "isflagged:false AND subject:[CGP]"
    $itemsInbox = $scanFolder.FindItems($searchQuery,1000)
    Write-Log ("Найдено новых запросов на миграцию: " + $itemsInbox.TotalCount) -Log $Log
    if ($itemsInbox.TotalCount -ne 0) {
        $firstRequest = $itemsInbox.Items[0]
    } else {
        Write-Log ("Завершение работы скрипта") -Log $Log
        Stop-Logging -Log $Log; Exit
    }
    
# Пометка запроса флагом к исполнению (для исключения его повторной обработки)
    Write-Log "Пометка запроса флагом к исполнению (для исключения его повторной обработки)" -Log $Log
    #$flag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common, 0x8530,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String) 
    $flagIcon = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1095,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $firstRequest.Flag.FlagStatus = [Microsoft.Exchange.WebServices.Data.ItemFlagStatus]::Flagged # Complete #  Flagged
    $firstRequest.Flag.StartDate = (Get-Date)
    $firstRequest.Flag.DueDate = (Get-Date).AddMinutes(10)
    $firstRequest.SetExtendedProperty($flagIcon,6)
    $firstRequest.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)   
# 2. Для первого письма в списке:
# 2.1 Получение адреса отправителя и сохранение его LegacyExchangeDN
    $targetAccount = $firstRequest.From
    Write-Log ("Запрос на миграцию учётной записи: " + $targetAccount) -Log $Log
    $targetUser = Get-User -Identity $targetAccount.Address
    $DN = $targetUser.DistinguishedName
    Write-Log ("DN учётной записи: " + $DN) -Log $Log
    $LegacyExchangeDN = $targetUser.LegacyExchangeDN
    Write-Log ("Текущий LegacyExchangeDN учётной записи: " + $LegacyExchangeDN) -Log $Log
# 2.2 Удаление ящика, присвоение SMTP-адреса
    Write-Log ("Производится перенастройка учётной записи") -Log $Log
    Write-Log ("Отключение ящика от учётной записи") -Log $Log
    Try {
         Disable-Mailbox -Identity $targetAccount.Address -Confirm:$true #-WhatIf
    } Catch {
         Write-Log ("Отключение ящика от учётной записи не выполнено: " + $_.Exception.Message) -Log $Log -Type error
         Stop-Logging -Log $Log; Exit
    }
    Write-Log ("Присвоение учётной записи адреса SMTP") -Log $Log
    Try {
         Enable-MailUser -Identity $DN -UsePreferMessageFormat $false -MessageFormat Mime -MessageBodyFormat TextAndHtml -ExternalEmailAddress $targetAccount.Address -Confirm:$true #-WhatIf
    } Catch {
         Write-Log ("Присвоение учётной записи адреса SMTP не выполнено: " + $_.Exception.Message) -Log $Log -Type error
         Stop-Logging -Log $Log; Exit
    }    
# 2.3 Присвоение X500-адреса = сохраненному ранее LegacyExchangeDN
    Write-Log ("Присвоение учётной записи адреса X500") -Log $Log
    Try {
         Set-MailUser -Identity $DN -EmailAddresses @{Add="X500:" + $LegacyExchangeDN} -UseMapiRichTextFormat Never -Confirm:$true #-WhatIf
    } Catch {
         Write-Log ("Присвоение учётной записи адреса X500 не выполнено: " + $_.Exception.Message) -Log $Log -Type error
         Stop-Logging -Log $Log; Exit
    }
  
# 3. Пометка запроса флагом завершения (для исключения его повторной обработки)
    Write-Log "Пометка запроса флагом завершения (для исключения его повторной обработки)" -Log $Log
    $flag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common, 0x8530,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
    $flagIcon = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1095,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
    $firstRequest.Flag.FlagStatus = [Microsoft.Exchange.WebServices.Data.ItemFlagStatus]::Complete # Complete #  Flagged
    $firstRequest.Flag.CompleteDate = (Get-Date)
    $firstRequest.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
    Write-Log "Обработка запроса завершена" -Log $Log
}


###################################################################################################
# Выполнение функций
###################################################################################################

Start-Logging
Connect-EWS
Connect-EMS
Start-Request
Stop-Logging -Log