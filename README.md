# Remote Printer Driver Installer 🖨️

![PowerShell Version](https://img.shields.io/badge/PowerShell-5.0+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

Скрипт для автоматизированной установки драйверов принтеров и сканеров на удалённые компьютеры с пробной печатью.

## 📥 Скачать
```powershell
Invoke-WebRequest -Uri "https://example.com/Install-PrinterDrivers.ps1" -OutFile "Install-PrinterDrivers.ps1"
```

🌟 Особенности
Поддержка всех форматов драйверов (INF, MSI, EXE)

Графический интерфейс выбора архива

Автоматическая распаковка ZIP

Проверка установки через тестовую печать

Корректная обработка русской кодировки

Автоочистка временных файлов

🛠 Требования
Целевой компьютер:

Windows 7+

PowerShell 5.0+

Включён PSRemoting (Enable-PSRemoting)

Доступ к C$ share

Локальный компьютер:

.NET Framework 3.5+ (для GUI)

Права администратора

🚀 Быстрый старт
1. Запуск скрипта
# С правами администратора
Start-Process powershell -Verb RunAs -ArgumentList "-File .\Install-PrinterDrivers.ps1"

2. Пример работы
```
Введите имя целевого компьютера: SRV-PRINT01
[✓] Архив drivers.zip успешно скопирован
[✓] Драйверы распакованы в C:\Temp\Drivers\
[✓] Установлено 3 компонента

Список принтеров:
Name           Driver         Port       Status
----           ------         ----       ------
HP-LJ-4250     Universal      USB001     Ready
Xerox-C220     PCL6           IP_10.0.1.5 Offline

Отправка тестовой печати...
[✓] HP-LJ-4250: документ отправлен
[✗] Xerox-C220: принтер недоступен
```

📚 Подробная документация
🔧 Параметры установки
```
Тип файла	Команда установки	Параметры
.inf	pnputil /add-driver	/install /subdirs
.msi	msiexec /i	/qn /norestart
.exe	Запуск напрямую	/S /quiet
```

🔄 Логирование ошибок
Коды возврата:

0 - Успех

1 - Ошибка прав

2 - Файл не выбран

3 - Ошибка копирования

4 - Ошибка распаковки

🖨 Настройка печати
# Для кастомного текста печати
```
$testContent = @"
Компания: ООО «Пример»
Тестовая страница
Дата: $(Get-Date -Format 'dd.MM.yyyy')
"@
```

⚠️ Важно!
Проверьте политики выполнения:
```
powershell
Get-ExecutionPolicy -List
Для Linux-хостов используйте PowerShell Core
```

При проблемах с кодировкой:

powershell
Copy
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
📜 Лицензия
MIT License. Полный текст доступен в файле LICENSE.