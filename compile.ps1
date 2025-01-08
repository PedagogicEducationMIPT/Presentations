# Удаляем все предыдущие PDF-файлы в текущем каталоге и его подкаталогах
Get-ChildItem -Path $currentDir -Filter *.pdf -Recurse | Remove-Item -Force
# Ищем все файлы .pptx в текущем каталоге
$pptFiles = Get-ChildItem -Path $(Get-Location) -Filter *.pptx  -Recurse 

Add-Type -AssemblyName "Microsoft.Office.Interop.PowerPoint"
$powerpoint = New-Object -ComObject PowerPoint.Application
foreach ($file in $pptFiles) {
    # Открываем файл презентации
    $presentation = $powerpoint.Presentations.Open($file.FullName)
    # Конвертируем презентацию в PDF
    $outputFile = $file.FullName.Replace(".pptx", ".pdf")
    $presentation.SaveAs($outputFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    # Закрываем презентацию
    $presentation.Close()
}
# Завершаем работу приложения PowerPoint
$powerpoint.Quit()