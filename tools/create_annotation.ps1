$ErrorActionPreference = "Stop"

function Escape-Xml([string]$text) {
  if ($null -eq $text) { return "" }
  return [System.Security.SecurityElement]::Escape($text)
}

function Add-Paragraph([System.Collections.Generic.List[string]]$parts, [string]$text, [string]$style = "Normal") {
  $parts.Add("<w:p><w:pPr><w:pStyle w:val=`"$style`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$(Escape-Xml $text)</w:t></w:r></w:p>")
}

$root = Resolve-Path "."
$docDir = Join-Path $root "doc"
$outPath = Join-Path $docDir "3_Анотація_зелена-енергетика.docx"
$tmp = Join-Path $env:TEMP ("annotation-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "word") | Out-Null

$parts = New-Object System.Collections.Generic.List[string]

Add-Paragraph $parts "АНОТАЦІЯ" "H1"
Add-Paragraph $parts "Пояснювальна записка: орієнтовно 25–30 с., рисунки, приклади коду, використані джерела."
Add-Paragraph $parts "Тема роботи: «Розроблення сайту компанії у сфері зеленої енергетики»."
Add-Paragraph $parts "Метою курсового проєкту є розроблення інформаційного вебсайту компанії, що працює у сфері зеленої енергетики, який забезпечує зручне ознайомлення користувачів із напрямами діяльності організації, її місією, командою фахівців, контактними даними та можливостями зв’язку."
Add-Paragraph $parts "У курсовому проєкті створено багатосторінковий вебресурс, до складу якого входять головна сторінка «index.html», сторінка «about.html», сторінка «team.html» і сторінка «contact.html». На головній сторінці розміщено перший екран із навігацією та закликом до взаємодії, блок партнерів, промо-секцію з напрямами зеленої енергетики, інформацію про компанію, секцію експертів, відгуки клієнтів, відповіді на поширені запитання, блоговий блок і форму підписки."
Add-Paragraph $parts "Під час розроблення було сформовано єдину систему стилізації у файлі «styles/style.css», у якій описано базові правила оформлення, допоміжні класи, структуру секцій, кнопки, форми, картки, футер, контактну сторінку та сторінку команди. Для покращення взаємодії користувача із сайтом у файлі «script/app.js» реалізовано кнопку прокрутки сторінки вгору та модальне вікно з контактною формою."
Add-Paragraph $parts "Для створення вебсайту використано технології HTML, CSS і JavaScript. У проєкті підключено шрифт «Open Sans» через сервіс Google Fonts, а також бібліотеку «Font Awesome» для використання іконок в інтерфейсі. Графічні матеріали організовано в папці «images», де розміщено логотип, фонові зображення, іконки, фотографії експертів, зображення клієнтів, партнерів і матеріали блогового блоку."
Add-Paragraph $parts "У пояснювальній записці розглянуто предметну область проєкту, призначення вебсайту, структуру основних сторінок, особливості HTML-розмітки, CSS-оформлення, реалізацію інтерактивних компонентів і підходи до тестування. Наведені приклади коду демонструють побудову навігації, інформаційних секцій, контактної форми, модального вікна, кнопки прокрутки та стилів окремих елементів інтерфейсу."
Add-Paragraph $parts "Ключові слова: зелена енергетика, вебсайт, вебсторінка, HTML, CSS, JavaScript, інтерфейс користувача, модальне вікно, контактна форма, Font Awesome, Google Fonts, веброзробка."

$documentBody = ($parts -join "`n")

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    $documentBody
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="160" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/><w:ind w:firstLine="708"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="H1">
    <w:name w:val="Heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="240"/><w:jc w:val="center"/><w:keepNext/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:b/><w:sz w:val="30"/></w:rPr>
  </w:style>
</w:styles>
"@

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"@

$rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $tmp "[Content_Types].xml"), $contentTypes, $utf8)
[System.IO.File]::WriteAllText((Join-Path $tmp "_rels/.rels"), $rels, $utf8)
[System.IO.File]::WriteAllText((Join-Path $tmp "word/document.xml"), $documentXml, $utf8)
[System.IO.File]::WriteAllText((Join-Path $tmp "word/styles.xml"), $stylesXml, $utf8)

if (Test-Path $outPath) {
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $outPath = Join-Path $docDir "3_Анотація_зелена-енергетика_$stamp.docx"
}

$zipPath = [System.IO.Path]::ChangeExtension($outPath, ".zip")
if (Test-Path $zipPath) {
  Remove-Item -Path $zipPath -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$fs = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::Create)
$archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $tmp "[Content_Types].xml"), "[Content_Types].xml") | Out-Null
[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $tmp "_rels/.rels"), "_rels/.rels") | Out-Null
[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $tmp "word/document.xml"), "word/document.xml") | Out-Null
[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $tmp "word/styles.xml"), "word/styles.xml") | Out-Null
$archive.Dispose()
$fs.Dispose()
Move-Item -Path $zipPath -Destination $outPath -Force
Remove-Item -Path $tmp -Recurse -Force

Write-Output $outPath

