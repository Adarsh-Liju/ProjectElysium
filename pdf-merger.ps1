# FULL PATH OF PDFs
$documents_path = 'C:\Users\adars\Desktop\Project_Elysium'

$word_app = New-Object -ComObject Word.Application

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {

    $document = $word_app.Documents.Open($_.FullName)

    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"

    Write-Output "Saving to: $pdf_filename"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)

    $document.Close()
}

$word_app.Quit()
Write-Output "Done"
Write-Output "Merging PDFs"
Merge-PDF -InputFile $documents_path -OutputFile $documents_path\COMBINED.pdf
