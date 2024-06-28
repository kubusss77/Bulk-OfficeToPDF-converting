# Script to batch convert Office365 files PDFs using OfficeToPDF.exe
# (c) 2024 Jakub Szczepa

$dir_input = "docx"
$dir_output = "pdf"
$extension = "docx"



Set-Location -Path $dir_input
$files = Get-ChildItem -Filter "*.$extension" -File
Set-Location -Path ..

foreach ($file in $files) {
    try {
        & .\OfficeToPDF.exe $file.FullName $dir_output\$($file.BaseName).pdf
        Write-Host "Successfully converted $($file.Name)"
    } catch {
        Write-Host "Failed to convert $($file.Name)"
    }
}

Write-Host "Conversion completed"