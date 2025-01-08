# Проверяем наличие установленного модуля MS Office
if (-not (Get-Module -ListAvailable | Where-Object { $_.Name -eq "MSOnline" })) {
    Write-Host "Не найден модуль MS Online. Установите его командой Install-Module MSOnline"
}

Add-Type -AssemblyName "Microsoft.Office.Interop.PowerPoint"
$powerpoint = New-Object -ComObject PowerPoint.Application

# Получаем текущий каталог
$currentDir = Get-Location

# Ищем все файлы .pptx в текущем каталоге
$pptFiles = Get-ChildItem -Path $currentDir -Filter *.pptx

foreach ($file in $pptFiles) {
    # Открываем файл презентации
    $presentation = $powerpoint.Presentations.Open($file.FullName)
    
    # Формируем имя выходного файла PDF
    $outputFile = Join-Path $currentDir "$($file.BaseName).pdf"
    
    # Конвертируем презентацию в PDF
    $presentation.SaveAs($outputFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    
    # Закрываем презентацию
    $presentation.Close()
    
    Write-Host "Конвертирован файл: $($file.Name)"
}

# Завершаем работу приложения PowerPoint
$powerpoint.Quit()