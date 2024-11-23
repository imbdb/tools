# tools

[merge-pdfs.html](https://imbdb.github.io/tools/merge-pdfs.html) - Merge multiple pdfs into onw

[pdf-page-extractor.html](https://imbdb.github.io/tools/pdf-page-extractor.html) - Extract some pages from a pdf file using page numbers

[pdf-page-remover.html](https://imbdb.github.io/tools/pdf-page-remover.html) - Remove particular pages from a pdf file

[title-extractor.html](https://imbdb.github.io/tools/title-extractor.html) - Extract titles from a pdf file

[pdf-pagenumber.html](https://imbdb.github.io/tools/pdf-pagenumber.html) - Add Page number to pdf file

## Convert PPTx to PDF in current directory

### Windows

```powershell
Get-ChildItem *.pptx | ForEach-Object { $powerPoint = New-Object -ComObject PowerPoint.Application; $presentation = $powerPoint.Presentations.Open($_.FullName); $pdfPath = [System.IO.Path]::ChangeExtension($_.FullName, "pdf"); $presentation.SaveAs($pdfPath, 32); $presentation.Close(); $powerPoint.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint); }
```

### Linux/Mac

```shell
for file in *.pptx; do libreoffice --headless --convert-to pdf "$file"; done
```

## Add prefix to every pdf name in current directory

```cmd
for file in *.pdf; do
    if [[ ! $file =~ ^prefix_ ]]; then
        mv "$file" "prefix_$file"
    fi
done
```
