<#
.SYNOPSIS
Remote printer driver installation script with test print functionality and enhanced timeout handling
#>

#Requires -Version 5.0
#Requires -RunAsAdministrator

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Запустите скрипт от имени администратора" -Category AuthenticationError
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms

# Параметры таймаутов
$infTimeout = 120    # 2 минуты для INF
$msiTimeout = 600    # 10 минут для MSI
$exeTimeout = 600    # 10 минут для EXE
$printTimeout = 300  # 5 минут для печати

$computer = Read-Host "Введите имя целевого компьютера"

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Выберите архив с драйверами (ZIP)"
$dialog.Filter = "ZIP files (*.zip)|*.zip"
$dialog.Multiselect = $false

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Файл не выбран. Выход."
    exit 2
}
$localDriverPath = $dialog.FileName

$remoteTempDir = "\\$computer\C$\Temp\Drivers\"
$remoteUnpackDir = "C:\Temp\Drivers\"

try {
    $robocopyLog = robocopy (Split-Path $localDriverPath) $remoteTempDir (Split-Path $localDriverPath -Leaf) /Z /J /NJH /NJS /R:3 /W:5
    if ($LASTEXITCODE -gt 7) { throw "Robocopy error: $LASTEXITCODE" }
}
catch {
    Write-Error "Ошибка копирования файлов: $_"
    exit 3
}

Invoke-Command -ComputerName $computer -ScriptBlock {
    param($remoteUnpackDir, $infTimeout, $msiTimeout, $exeTimeout, $printTimeout)
    
    $originalEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(1251)

    if (-not (Test-Path $remoteUnpackDir)) {
        New-Item -Path $remoteUnpackDir -ItemType Directory -Force | Out-Null
    }

    try {
        $zipFile = Get-Item "$remoteUnpackDir*.zip" -ErrorAction Stop
        Expand-Archive -Path $zipFile.FullName -DestinationPath $remoteUnpackDir -Force
    }
    catch {
        Write-Error "Ошибка распаковки архива: $_"
        exit 4
    }

    Get-ChildItem -Path $remoteUnpackDir -Recurse -File | ForEach-Object {
        switch ($_.Extension.ToLower()) {
            '.inf' {
                try {
                    $infProcess = Start-Process "pnputil.exe" -ArgumentList "/add-driver $($_.FullName) /install /subdirs" -PassThru -NoNewWindow
                    $startTime = Get-Date
                    
                    while (!$infProcess.HasExited) {
                        if ((Get-Date - $startTime).TotalSeconds -gt $infTimeout) {
                            Stop-Process -Id $infProcess.Id -Force -ErrorAction SilentlyContinue
                            Write-Warning "Превышено время установки INF ($infTimeout сек): $($_.Name)"
                            break
                        }
                        Start-Sleep -Seconds 5
                    }
                    
                    if ($infProcess.ExitCode -ne 0) {
                        Write-Warning "Ошибка установки INF (код $($infProcess.ExitCode)): $($_.Name)"
                    } else {
                        Write-Host "Установлен INF: $($_.Name)" -ForegroundColor DarkGray
                    }
                }
                catch {
                    Write-Warning "Ошибка установки INF: $_"
                }
            }
            '.msi' {
                try {
                    $msiProcess = Start-Process "msiexec.exe" -ArgumentList "/i `"$($_.FullName)`" /qn /norestart" -PassThru -NoNewWindow
                    $startTime = Get-Date
                    
                    while (!$msiProcess.HasExited) {
                        if ((Get-Date - $startTime).TotalSeconds -gt $msiTimeout) {
                            Stop-Process -Id $msiProcess.Id -Force -ErrorAction SilentlyContinue
                            Write-Warning "Превышено время установки MSI ($msiTimeout сек): $($_.Name)"
                            break
                        }
                        Start-Sleep -Seconds 5
                    }
                    
                    if ($msiProcess.ExitCode -ne 0) {
                        Write-Warning "Ошибка установки MSI (код $($msiProcess.ExitCode)): $($_.Name)"
                    } else {
                        Write-Host "Установлен MSI: $($_.Name)" -ForegroundColor DarkGray
                    }
                }
                catch {
                    Write-Warning "Ошибка установки MSI: $_"
                }
            }
            '.exe' {
                try {
                    $exeProcess = Start-Process $_.FullName -ArgumentList "/S /quiet /norestart" -PassThru -NoNewWindow
                    $startTime = Get-Date
                    
                    while (!$exeProcess.HasExited) {
                        if ((Get-Date - $startTime).TotalSeconds -gt $exeTimeout) {
                            Stop-Process -Id $exeProcess.Id -Force -ErrorAction SilentlyContinue
                            Write-Warning "Превышено время установки EXE ($exeTimeout сек): $($_.Name)"
                            break
                        }
                        Start-Sleep -Seconds 5
                    }
                    
                    if ($exeProcess.ExitCode -ne 0) {
                        Write-Warning "Ошибка установки EXE (код $($exeProcess.ExitCode)): $($_.Name)"
                    } else {
                        Write-Host "Запущен EXE: $($_.Name)" -ForegroundColor DarkGray
                    }
                }
                catch {
                    Write-Warning "Ошибка запуска EXE: $_"
                }
            }
        }
    }

    Write-Host "`nСписок установленных принтеров:" -ForegroundColor Cyan
    $printers = Get-Printer -ErrorAction SilentlyContinue
    if ($printers) {
        $printers | Format-Table -Property Name, DriverName, PortName, Shared, Type -AutoSize | Out-String -Width 4096
    }
    else {
        Write-Host "Принтеры не обнаружены" -ForegroundColor Yellow
    }

    if ($printers) {
        Write-Host "`nОтправка пробной печати..." -ForegroundColor Cyan
        $testFile = "$env:TEMP\testprint.txt"
        "Тестовая печать`nУспешная установка драйверов!`nДата: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File $testFile -Encoding Unicode

        foreach ($printer in $printers) {
            try {
                $printProcess = Start-Process notepad.exe -ArgumentList "/p `"$testFile`"" -PassThru -NoNewWindow
                $startTime = Get-Date
                
                while (!$printProcess.HasExited) {
                    if ((Get-Date - $startTime).TotalSeconds -gt $printTimeout) {
                        Stop-Process -Id $printProcess.Id -Force -ErrorAction SilentlyContinue
                        Write-Host "[✗] Печать на $($printer.Name) прервана по таймауту ($printTimeout сек)" -ForegroundColor Red
                        break
                    }
                    Start-Sleep -Seconds 5
                }
                
                if ($printProcess.ExitCode -eq 0) {
                    Write-Host "[✓] Печать на $($printer.Name)" -ForegroundColor Green
                }
                else {
                    Write-Host "[✗] Ошибка печати на $($printer.Name) (код $($printProcess.ExitCode))" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "[✗] Ошибка печати на $($printer.Name): $_" -ForegroundColor Red
            }
        }
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    }

    [Console]::OutputEncoding = $originalEncoding
    Remove-Item -Path "$remoteUnpackDir*" -Recurse -Force -ErrorAction SilentlyContinue

} -ArgumentList $remoteUnpackDir, $infTimeout, $msiTimeout, $exeTimeout, $printTimeout -ErrorAction Stop

Write-Host "`nПроцесс завершён успешно!`n" -ForegroundColor Green
