<#
.SYNOPSIS
Remote printer driver installation script with test print functionality and stuck process handling
#>

#Requires -Version 5.0
#Requires -RunAsAdministrator

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Запустите скрипт от имени администратора" -Category AuthenticationError
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms

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
    param($remoteUnpackDir)
    
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
                    pnputil.exe /add-driver $_.FullName /install /subdirs | Out-Null
                    Write-Host "Установлен INF: $($_.Name)" -ForegroundColor DarkGray
                }
                catch {
                    Write-Warning "Ошибка установки INF: $_"
                }
            }
            '.msi' {
                try {
                    $process = Start-Process msiexec.exe -ArgumentList "/i `"$($_.FullName)`" /qn /norestart" -PassThru -NoNewWindow
                    try {
                        $process | Wait-Process -Timeout 600 -ErrorAction Stop
                        if ($process.ExitCode -ne 0) {
                            Write-Warning "Ошибка установки MSI: $($_.Name), код $($process.ExitCode)"
                        } else {
                            Write-Host "Установлен MSI: $($_.Name)" -ForegroundColor DarkGray
                        }
                    }
                    catch {
                        $process | Stop-Process -Force
                        Write-Warning "Установка MSI $($_.Name) прервана по таймауту (10 минут)"
                    }
                }
                catch {
                    Write-Warning "Ошибка запуска MSI: $_"
                }
            }
            '.exe' {
                try {
                    $process = Start-Process $_.FullName -ArgumentList "/S /quiet /norestart" -PassThru -NoNewWindow
                    try {
                        $process | Wait-Process -Timeout 600 -ErrorAction Stop
                        if ($process.ExitCode -ne 0) {
                            Write-Warning "Ошибка запуска EXE: $($_.Name), код $($process.ExitCode)"
                        } else {
                            Write-Host "Запущен EXE: $($_.Name)" -ForegroundColor DarkGray
                        }
                    }
                    catch {
                        $process | Stop-Process -Force
                        Write-Warning "Запуск EXE $($_.Name) прерван по таймауту (10 минут)"
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
                try {
                    $printProcess | Wait-Process -Timeout 300 -ErrorAction Stop
                    if ($printProcess.ExitCode -eq 0) {
                        Write-Host "[✓] Печать на $($printer.Name)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "[✗] Ошибка печати на $($printer.Name) (код $($printProcess.ExitCode))" -ForegroundColor Red
                    }
                }
                catch {
                    $printProcess | Stop-Process -Force
                    Write-Host "[✗] Печать на $($printer.Name) прервана по таймауту (5 минут)" -ForegroundColor Red
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

} -ArgumentList $remoteUnpackDir -ErrorAction Stop

Write-Host "`nПроцесс завершён успешно!`n" -ForegroundColor Green
