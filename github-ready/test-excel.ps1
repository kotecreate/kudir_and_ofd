$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dir=Join-Path $PSScriptRoot 'test-output'
[void][IO.Directory]::CreateDirectory($dir)
$path=Join-Path $dir 'fixture.xlsx'
$stream=[IO.File]::Open($path,[IO.FileMode]::Create)
$zip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
try {
    $parts=@{
        '[Content_Types].xml'='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels'='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'='<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Данные" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels'='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml'='<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Дата</t></is></c><c r="B1" t="inlineStr"><is><t>Доход</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>03.09.2026</t></is></c><c r="B2"><v>150.25</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>04.09.2026</t></is></c><c r="B3"><v>49.75</v></c></row></sheetData></worksheet>'
    }
    foreach ($key in $parts.Keys) {
        $entry=$zip.CreateEntry($key); $writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
        try { $writer.Write($parts[$key]) } finally { $writer.Dispose() }
    }
} finally { $zip.Dispose(); $stream.Dispose() }
$before=(Get-FileHash -LiteralPath $path).Hash
$names=Read-Excel $path
if ($names.Count -ne 1 -or $names[0] -ne 'Данные') { throw 'Wrong worksheet names' }
$data=Read-Excel $path 'Данные'
$records=Extract-Records $data A B '' 2 3
if ($records.Records.Count -ne 2 -or $records.Records[0].Amount -ne [decimal]150.25) { throw 'Wrong Excel values' }
$after=(Get-FileHash -LiteralPath $path).Hash
if ($before -ne $after) { throw 'Input changed' }
Write-Output 'PASS: actual Excel read, sheet selection, decimal values, input unchanged'
