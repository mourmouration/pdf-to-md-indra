<#
    pdf_to_md_indra.ps1
    Пакетная конвертация целой папки (рекурсивно) в Markdown.
    Версия на движке Indra OCR (https://github.com/datalab-to/indra) вместо
    OCRmyPDF + Tesseract.

    Обрабатывает три типа документов:

      1) PDF – Indra OCR сразу выдаёт Markdown:
           python indra_batch.py <задания> <отчёт>
         Режим auto: страницы с годным текстовым слоем читаются из него,
         остальные (сканы) распознаются. Промежуточный *_ocr.pdf не создаётся,
         исходный PDF остаётся на месте и никуда не переносится.

      2) Изображения (.png, .jpg, .jpeg, .tif, .tiff, .bmp, .webp) – тем же
         путём через Indra OCR. Многостраничный TIFF – по странице на кадр.

      3) Word / Excel / PowerPoint и прочие офисные форматы – без OCR:
           python -m markitdown <src> -o <md>
         Поддерживаются: .docx, .xlsx, .xls, .pptx, .html, .htm, .csv,
         .txt, .epub. Форматы .doc и .ppt пропускаются с предупреждением.

    Все файлы для Indra распознаются ОДНИМ запуском Python: модели
    загружаются один раз, страницы всех файлов распределяются по ядрам.

    Папка результатов локальная – в каждой подпапке, где нашлись файлы:
      <обработанная папка>\0.1. MD\...

    Совместимость с результатами прежней версии (OCRmyPDF):
      - папки "0.1. MD", "0.2. Not ocr", "MD", "Not ocr" исключаются из поиска;
      - X_ocr.pdf рядом с X.pdf пропускается (распознаётся исходник X.pdf);
      - X_ocr.pdf без исходника распознаётся, а результат пишется в X.md –
        уже готовый X.md от прежней версии не переделывается, пока X_ocr.pdf
        не изменится (или пока не задан ключ -Rebuild).

    Параметры:
      -FolderPath     папка для обработки (если не задана – будет запрошена)
      -Cores N        сколько ядер использовать (по умолчанию все)
      -ForceOcr       распознавать все страницы PDF, игнорируя текстовый слой
      -KeepFurniture  сохранять колонтитулы и номера страниц в .md
      -Rebuild        пересоздать .md для PDF и изображений, даже если они
                      актуальны (ручные правки в этих .md будут перезаписаны)
#>

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$FolderPath,
    [int]$Cores = 0,
    [switch]$ForceOcr,
    [switch]$KeepFurniture,
    [switch]$Rebuild
)

# Намеренно НЕ "Stop": Python пишет в stderr и при нормальной работе.
$ErrorActionPreference = "Continue"

if (-not $FolderPath) {
    $FolderPath = Read-Host "Введите путь к папке с файлами"
}

$FolderPath = $FolderPath.Trim().Trim('"').TrimEnd('\', '/')

if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
    Write-Host "Папка не найдена: $FolderPath" -ForegroundColor Red
    exit 1
}

$FolderPath = (Resolve-Path -LiteralPath $FolderPath).ProviderPath

$helperScript = Join-Path $PSScriptRoot "indra_batch.py"
if (-not (Test-Path -LiteralPath $helperScript)) {
    Write-Host "Не найден indra_batch.py рядом со скриптом: $helperScript" -ForegroundColor Red
    exit 1
}

# --- Вспомогательные функции ---

function Test-CommandExists {
    param([string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

# Возвращает интерпретатор, в котором модуль реально импортируется, либо $null.
function Resolve-PythonForModule {
    param([string]$Module, [string[]]$Candidates)
    foreach ($candidate in $Candidates) {
        if (-not (Test-CommandExists $candidate)) { continue }
        try {
            $null = & $candidate -c "import $Module" 2>&1
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch { }
    }
    return $null
}

function Invoke-ExternalTool {
    param([string]$Exe, [string[]]$ToolArgs, [int[]]$OkExitCodes = @(0))

    $captured = @()
    $code = -1
    try {
        $captured = & $Exe @ToolArgs 2>&1
        $code = $LASTEXITCODE
    } catch {
        $captured = @($_.Exception.Message)
        $code = -1
    }

    $lines = @()
    foreach ($item in $captured) {
        if ($null -ne $item) {
            $text = $item.ToString()
            if ($text.Trim() -ne "" -and $text -ne "System.Management.Automation.RemoteException") { $lines += $text }
        }
    }
    $tail = @($lines | Select-Object -Last 8)

    return [PSCustomObject]@{
        ExitCode = $code
        Ok       = ($OkExitCodes -contains $code)
        Details  = ($tail -join "; ")
    }
}

function Test-OutputIsStale {
    param([string]$OutputPath, [string]$InputPath)
    if (-not (Test-Path -LiteralPath $OutputPath)) { return $true }
    if (-not (Test-Path -LiteralPath $InputPath))  { return $false }
    $outputItem = Get-Item -LiteralPath $OutputPath
    $sourceItem = Get-Item -LiteralPath $InputPath
    return ($sourceItem.LastWriteTimeUtc -gt $outputItem.LastWriteTimeUtc)
}

function New-TempOutputPath {
    param([string]$FinalPath)
    $dir       = Split-Path -Parent $FinalPath
    $leaf      = Split-Path -Leaf $FinalPath
    $extension = [System.IO.Path]::GetExtension($leaf)
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    $stamp     = "{0}-{1}" -f $PID, (Get-Random -Maximum 1000000)
    return (Join-Path $dir ("~$baseName.partial-$stamp$extension"))
}

function Remove-TempOutput {
    param([string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Test-MarkdownHasContent {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $item = Get-Item -LiteralPath $Path
    if ($item.Length -eq 0) { return $false }
    if ($item.Length -gt 512) { return $true }
    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($null -eq $content) { return $false }
    return ($content.Trim().Length -gt 0)
}

$utf8NoBom = New-Object System.Text.UTF8Encoding $false

# --- Настройки ---

$mdFolderName = "0.1. MD"

$indraExtensions       = @(".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp")
$directExtensions      = @(".docx", ".xlsx", ".xls", ".pptx", ".html", ".htm", ".csv", ".txt", ".epub")
$unsupportedExtensions = @(".doc", ".ppt")
$allExtensions         = $indraExtensions + $directExtensions + $unsupportedExtensions

$reservedFolderNames = @($mdFolderName, "0.2. Not ocr", "MD", "Not ocr")

function Test-InsideReservedFolder {
    param([string]$FullName, [string]$RootPath, [string[]]$ReservedNames)
    $relative = $FullName.Substring($RootPath.Length).TrimStart('\', '/')
    $segments = $relative -split '[\\/]'
    foreach ($segment in $segments) {
        foreach ($reserved in $ReservedNames) {
            if ($segment -ieq $reserved) { return $true }
        }
    }
    return $false
}

# --- Сбор файлов ---

$allFiles = @(
    Get-ChildItem -LiteralPath $FolderPath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
            $null -ne $_.FullName -and
            $allExtensions -contains $_.Extension.ToLower() -and
            -not $_.Name.StartsWith("~") -and
            -not (Test-InsideReservedFolder -FullName $_.FullName -RootPath $FolderPath -ReservedNames $reservedFolderNames)
        }
)

# Для каждого файла определяем "логическое" имя документа. У X_ocr.pdf от
# прежней версии это X (если исходника X.pdf рядом нет); если исходник рядом
# есть – X_ocr.pdf дублирует его и пропускается.
$sourceFiles = @()
$countOcrDuplicates = 0
foreach ($file in $allFiles) {
    $logicalBase = $file.BaseName
    if ($file.Name -like "*_ocr.pdf") {
        $originalBase = $file.BaseName -replace '_ocr$', ''
        $originalPath = Join-Path $file.DirectoryName ($originalBase + ".pdf")
        if (Test-Path -LiteralPath $originalPath) {
            $countOcrDuplicates++
            continue
        }
        $logicalBase = $originalBase
    }
    $sourceFiles += [PSCustomObject]@{
        File        = $file
        LogicalBase = $logicalBase
        LogicalName = $logicalBase + $file.Extension
    }
}

# --- Карта конфликтов имён (Договор.pdf + Договор.docx -> Договор.pdf.md / Договор.docx.md) ---

$nameGroups = @{}
foreach ($entry in $sourceFiles) {
    $key = ($entry.File.DirectoryName.ToLower() + "|" + $entry.LogicalBase.ToLower())
    if ($nameGroups.ContainsKey($key)) { $nameGroups[$key]++ } else { $nameGroups[$key] = 1 }
}

function Get-MdTargetPath {
    param($Entry)
    $mdDir = Join-Path $Entry.File.DirectoryName $mdFolderName
    $key = ($Entry.File.DirectoryName.ToLower() + "|" + $Entry.LogicalBase.ToLower())
    if ($nameGroups[$key] -gt 1) {
        return (Join-Path $mdDir ($Entry.LogicalName + ".md"))
    }
    return (Join-Path $mdDir ($Entry.LogicalBase + ".md"))
}

Write-Host "Найдено файлов для обработки: $($sourceFiles.Count)" -ForegroundColor Yellow
if ($countOcrDuplicates -gt 0) {
    Write-Host "Пропущено *_ocr.pdf прежней версии (рядом есть исходник): $countOcrDuplicates" -ForegroundColor DarkGray
}
Write-Host ""

# --- Счётчики ---

$countIndraDone    = 0
$countIndraUpdated = 0
$countIndraSkip    = 0
$countDirectDone   = 0
$countDirectUpdated = 0
$countDirectSkip   = 0
$countUnsupported  = 0
$countEmpty        = 0
$pageErrorList     = @()
$errorList         = @()

function Add-Error {
    param([string]$FileName, [string]$Stage, [string]$Reason)
    $script:errorList += [PSCustomObject]@{
        File   = $FileName
        Stage  = $Stage
        Reason = $Reason
    }
    Write-Host "  Ошибка ($Stage): $FileName – $Reason" -ForegroundColor Red
}

function New-MdFolder {
    param([string]$MdPath, [string]$DisplayName, [string]$Stage)
    $mdDir = Split-Path -Parent $MdPath
    try {
        New-Item -ItemType Directory -Path $mdDir -Force -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Add-Error -FileName $DisplayName -Stage $Stage -Reason "не удалось создать папку результатов: $($_.Exception.Message)"
        return $false
    }
}

# --- Разбор файлов по движкам ---

$directQueue = @()
$indraQueue  = @()

foreach ($entry in $sourceFiles) {
    $file = $entry.File
    $extension = $file.Extension.ToLower()

    if ($unsupportedExtensions -contains $extension) {
        Write-Host "Пропущен ($extension не поддерживается, пересохраните в новом формате): $($file.Name)" -ForegroundColor DarkYellow
        $countUnsupported++
        continue
    }

    $mdFile = Get-MdTargetPath -Entry $entry
    $isIndra = $indraExtensions -contains $extension
    $stale = Test-OutputIsStale -OutputPath $mdFile -InputPath $file.FullName
    if ($isIndra -and $Rebuild) { $stale = $true }

    if (-not $stale) {
        Write-Host "MD уже есть: $($file.Name)" -ForegroundColor DarkGray
        if ($isIndra) { $countIndraSkip++ } else { $countDirectSkip++ }
        continue
    }

    $item = [PSCustomObject]@{
        Entry    = $entry
        MdFile   = $mdFile
        IsUpdate = (Test-Path -LiteralPath $mdFile)
    }
    if ($isIndra) { $indraQueue += $item } else { $directQueue += $item }
}

# --- 1. Word / Excel / PowerPoint и прочее: markitdown ---

if ($directQueue.Count -gt 0) {
    $pythonForMd = Resolve-PythonForModule -Module "markitdown" -Candidates @("python", "python3", "py")
    $env:PYTHONWARNINGS = "ignore::RuntimeWarning:pydub.utils"

    if (-not $pythonForMd) {
        foreach ($item in $directQueue) {
            Add-Error -FileName $item.Entry.File.Name -Stage "MD" -Reason "markitdown не установлен: python -m pip install --upgrade ""markitdown[pdf,docx,pptx,xlsx,xls]"""
        }
    } else {
        foreach ($item in $directQueue) {
            $file = $item.Entry.File
            if ($item.IsUpdate) {
                Write-Host "MD (исходник изменился, пересоздаю): $($file.Name)" -ForegroundColor Cyan
            } else {
                Write-Host "MD: $($file.Name)" -ForegroundColor Cyan
            }
            if (-not (New-MdFolder -MdPath $item.MdFile -DisplayName $file.Name -Stage "MD")) { continue }

            $tempMd = New-TempOutputPath -FinalPath $item.MdFile
            $result = Invoke-ExternalTool -Exe $pythonForMd -ToolArgs @("-m", "markitdown", $file.FullName, "-o", $tempMd)

            if (-not $result.Ok -or -not (Test-Path -LiteralPath $tempMd)) {
                Remove-TempOutput -Path $tempMd
                $reason = "markitdown завершился с кодом $($result.ExitCode)"
                if ($result.Details) { $reason += " – $($result.Details)" }
                Add-Error -FileName $file.Name -Stage "MD" -Reason $reason
                continue
            }

            $hasContent = Test-MarkdownHasContent -Path $tempMd
            try {
                Move-Item -LiteralPath $tempMd -Destination $item.MdFile -Force -ErrorAction Stop
            } catch {
                Remove-TempOutput -Path $tempMd
                Add-Error -FileName $file.Name -Stage "MD" -Reason "не удалось сохранить результат: $($_.Exception.Message)"
                continue
            }

            if (-not $hasContent) {
                Write-Host "  Результат пуст – текст извлечь не удалось: $($file.Name)" -ForegroundColor DarkYellow
                $countEmpty++
            } elseif ($item.IsUpdate) {
                $countDirectUpdated++
            } else {
                $countDirectDone++
            }
        }
    }
}

# --- 2. PDF и изображения: Indra OCR одним запуском ---

if ($indraQueue.Count -gt 0) {
    Write-Host ""
    Write-Host "Indra OCR: файлов к распознаванию – $($indraQueue.Count)" -ForegroundColor Cyan

    $pythonForIndra = Resolve-PythonForModule -Module "indra_ocr" -Candidates @("python", "python3", "py")

    if (-not $pythonForIndra) {
        foreach ($item in $indraQueue) {
            Add-Error -FileName $item.Entry.File.Name -Stage "Indra" -Reason "indra-ocr не установлен: python -m pip install --upgrade indra-ocr"
        }
    } else {
        # Задания: исходник -> временный .md рядом с итоговым
        $jobs = @()
        $jobMap = @{}
        foreach ($item in $indraQueue) {
            $file = $item.Entry.File
            if (-not (New-MdFolder -MdPath $item.MdFile -DisplayName $file.Name -Stage "Indra")) { continue }
            $tempMd = New-TempOutputPath -FinalPath $item.MdFile
            $item | Add-Member -NotePropertyName TempMd -NotePropertyValue $tempMd
            $jobs += [PSCustomObject]@{ src = $file.FullName; out = $tempMd }
            $jobMap[$file.FullName] = $item
        }

        if ($jobs.Count -gt 0) {
            $stamp      = "{0}-{1}" -f $PID, (Get-Random -Maximum 1000000)
            $jobsFile   = Join-Path $env:TEMP "indra_jobs_$stamp.json"
            $resultFile = Join-Path $env:TEMP "indra_result_$stamp.json"

            $json = ConvertTo-Json -InputObject @($jobs) -Depth 3
            [System.IO.File]::WriteAllText($jobsFile, $json, $utf8NoBom)

            $helperArgs = @($helperScript, $jobsFile, $resultFile)
            if ($Cores -gt 0)   { $helperArgs += @("--cores", "$Cores") }
            if ($ForceOcr)      { $helperArgs += @("--mode", "ocr") }
            if ($KeepFurniture) { $helperArgs += "--keep-furniture" }

            $env:HF_HUB_DISABLE_SYMLINKS_WARNING = "1"
            $env:PYTHONUTF8 = "1"

            # Вывод не перехватывается: прогресс идёт прямо в консоль.
            # При первом запуске скачиваются модели (~76 МБ).
            & $pythonForIndra @helperArgs
            $helperExit = $LASTEXITCODE

            $results = @()
            if (Test-Path -LiteralPath $resultFile) {
                try {
                    $raw = [System.IO.File]::ReadAllText($resultFile, [System.Text.Encoding]::UTF8)
                    $results = @($raw | ConvertFrom-Json)
                } catch {
                    $results = @()
                }
            }

            $handled = @{}
            foreach ($r in $results) {
                $item = $jobMap[$r.src]
                if ($null -eq $item) { continue }
                $handled[$r.src] = $true
                $file = $item.Entry.File

                if ($r.status -eq "error" -or -not (Test-Path -LiteralPath $item.TempMd)) {
                    Remove-TempOutput -Path $item.TempMd
                    $reason = $r.reason
                    if (-not $reason) { $reason = "результат не получен" }
                    Add-Error -FileName $file.Name -Stage "Indra" -Reason $reason
                    continue
                }

                try {
                    Move-Item -LiteralPath $item.TempMd -Destination $item.MdFile -Force -ErrorAction Stop
                } catch {
                    Remove-TempOutput -Path $item.TempMd
                    Add-Error -FileName $file.Name -Stage "Indra" -Reason "не удалось сохранить результат: $($_.Exception.Message)"
                    continue
                }

                foreach ($pe in @($r.page_errors)) {
                    if ($pe) { $script:pageErrorList += [PSCustomObject]@{ File = $file.Name; Detail = $pe } }
                }

                if ($r.status -eq "empty") {
                    Write-Host "  Результат пуст – распознать текст не удалось: $($file.Name)" -ForegroundColor DarkYellow
                    $countEmpty++
                } elseif ($item.IsUpdate) {
                    $countIndraUpdated++
                } else {
                    $countIndraDone++
                }
            }

            # Файлы без записи в отчёте: помощник аварийно завершился
            foreach ($item in $indraQueue) {
                if (-not $item.TempMd) { continue }
                $src = $item.Entry.File.FullName
                if ($handled.ContainsKey($src)) { continue }
                Remove-TempOutput -Path $item.TempMd
                Add-Error -FileName $item.Entry.File.Name -Stage "Indra" -Reason "Indra OCR аварийно завершился (код $helperExit) – см. сообщения выше"
            }

            Remove-TempOutput -Path $jobsFile
            Remove-TempOutput -Path $resultFile
        }
    }
}

# --- Итоги ---

Write-Host ""
Write-Host "Готово." -ForegroundColor Green
Write-Host "PDF и изображения (Indra OCR):"
Write-Host "  MD выполнен:               $countIndraDone"
Write-Host "  MD обновлён:               $countIndraUpdated"
Write-Host "  MD пропущен (есть):        $countIndraSkip"
Write-Host "Word/Excel/PowerPoint/прочее (markitdown):"
Write-Host "  MD выполнен:               $countDirectDone"
Write-Host "  MD обновлён:               $countDirectUpdated"
Write-Host "  MD пропущен (есть):        $countDirectSkip"
Write-Host "  Пропущено (.doc/.ppt):     $countUnsupported"
Write-Host ""

if ($countEmpty -gt 0) {
    Write-Host "Пустых результатов (текст извлечь не удалось): $countEmpty" -ForegroundColor DarkYellow
}

if ($pageErrorList.Count -gt 0) {
    Write-Host ""
    Write-Host "Страницы, не прочитанные внутри обработанных файлов: $($pageErrorList.Count)" -ForegroundColor DarkYellow
    foreach ($pe in $pageErrorList) {
        Write-Host ("  {0}: {1}" -f $pe.File, $pe.Detail) -ForegroundColor DarkGray
    }
}

if ($errorList.Count -gt 0) {
    Write-Host ""
    Write-Host "Ошибок: $($errorList.Count)" -ForegroundColor Red
    foreach ($item in $errorList) {
        Write-Host ("  [{0}] {1}" -f $item.Stage, $item.File) -ForegroundColor Red
        Write-Host ("      {0}" -f $item.Reason) -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "Исправьте причину и запустите скрипт повторно – готовые файлы не переделываются."
} else {
    Write-Host "Ошибок: 0" -ForegroundColor Green
}

Write-Host ""
Write-Host "MD-файлы: в подпапке '$mdFolderName' внутри КАЖДОЙ обработанной папки"
Write-Host "Исходные файлы остались на своих местах."
