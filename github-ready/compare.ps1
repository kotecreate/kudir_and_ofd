param([switch]$NoGui)
$ErrorActionPreference = 'Stop'
$script:Ru = [Globalization.CultureInfo]::GetCultureInfo('ru-RU')

function Money($value) {
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $null }
    if ($value -is [Runtime.InteropServices.ErrorWrapper] -or ($value -is [int] -and $value -ge -2146826288 -and $value -le -2146826240)) { throw 'В ячейке ошибка Excel. Исправьте формулу в исходном файле.' }
    if ($value -is [double] -or $value -is [decimal] -or $value -is [int]) { return [decimal]::Round([decimal]$value, 2, [MidpointRounding]::AwayFromZero) }
    $s = ([string]$value).Trim().Replace([char]0xA0,' ').Replace([char]0x202F,' ')
    $s = $s -replace '(?i)\s*(руб\.?|₽)\s*$', ''
    $s = $s.Replace(' ','')
    if ($s -match '^\((.+)\)$') { $s = '-' + $Matches[1] }
    if ($s -notmatch '^[+-]?\d+([,.]\d{1,2})?$') { throw "Не удалось прочитать сумму: $value" }
    return [decimal]::Parse($s.Replace(',','.'), [Globalization.CultureInfo]::InvariantCulture)
}

function Read-Date($value) {
    if ($null -eq $value) { throw 'Нет даты' }
    if ($value -is [datetime]) { return $value.Date }
    if ($value -is [double] -or $value -is [decimal] -or $value -is [int]) {
        $n = [double]$value
        if ($n -lt 1 -or $n -gt 109574) { throw "Некорректная дата Excel: $value" }
        return [datetime]::FromOADate($n).Date
    }
    $s = [string]$value
    $found = [regex]::Matches($s, '(?<!\d)(\d{1,2}[./-]\d{1,2}[./-]\d{4}|\d{4}-\d{2}-\d{2})(?!\d)')
    if ($found.Count -ne 1) { throw "Нужна одна дата с годом из четырёх цифр: $value" }
    foreach ($fmt in @('d.M.yyyy','d/M/yyyy','d-M-yyyy','yyyy-MM-dd')) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($found[0].Value,$fmt,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$d)) { return $d.Date }
    }
    throw "Не удалось прочитать дату: $value"
}

function Column-Name([int]$n) {
    $s = ''
    while ($n -gt 0) { $n--; $s = [string][char](65 + ($n % 26)) + $s; $n = [int][math]::Floor($n / 26) }
    return $s
}

function Release-Com($obj) { if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } }

function Read-Excel([string]$path, [string]$sheetName = '') {
    $excel=$books=$book=$sheets=$sheet=$range=$null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $books = $excel.Workbooks
        $book = $books.Open($path, 0, $true)
        $sheets = $book.Worksheets
        if (-not $sheetName) {
            $names = @()
            for ($i=1; $i -le $sheets.Count; $i++) {
                $sheet = $sheets.Item($i)
                $names += [string]$sheet.Name
                Release-Com $sheet; $sheet=$null
            }
            return ,$names
        }
        $sheet = $sheets.Item($sheetName)
        $range = $sheet.UsedRange
        $rowsObj=$range.Rows; $colsObj=$range.Columns
        try { $rowCount=[int]$rowsObj.Count; $colCount=[int]$colsObj.Count } finally { Release-Com $rowsObj; Release-Com $colsObj }
        if ([long]$rowCount * $colCount -gt 3000000) { throw 'Слишком большая область листа (более 3 млн ячеек). Сохраните нужную таблицу в отдельный файл без лишних пустых строк и столбцов.' }
        $values = $range.Value2
        $firstRow = [int]$range.Row; $firstCol = [int]$range.Column
        $columnNames=@{}; for ($c=1; $c -le $colCount; $c++) { $columnNames[$c]=Column-Name ($firstCol+$c-1) }
        $rows = [Collections.Generic.List[object]]::new()
        for ($r=1; $r -le $rowCount; $r++) {
            $cells = @{}
            for ($c=1; $c -le $colCount; $c++) {
                $v = if ($rowCount -eq 1 -and $colCount -eq 1) { $values } else { $values[$r,$c] }
                if ($null -ne $v -and [string]$v -ne '') { $cells[$columnNames[$c]] = $v }
            }
            if ($cells.Count -gt 0) { $rows.Add([pscustomobject]@{ Number=$firstRow+$r-1; Cells=$cells }) }
        }
        return [pscustomobject]@{ Path=$path; Sheet=$sheetName; Rows=$rows.ToArray(); First=$firstRow; Last=$firstRow+$rowCount-1; FirstCol=$firstCol; ColCount=$colCount; Date1904=[bool]$book.Date1904 }
    } catch {
        throw "Не удалось прочитать Excel. Закройте этот файл в Excel и повторите. Файл не должен иметь пароль. Причина: $($_.Exception.Message)"
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        foreach ($o in @($range,$sheet,$sheets,$book,$books,$excel)) { Release-Com $o }
    }
}

function Extract-Records($data, [string]$dateCol, [string]$amountCol, [string]$refundCol, [int]$first, [int]$last) {
    if ($dateCol -eq $amountCol -or ($refundCol -and ($refundCol -eq $amountCol -or $refundCol -eq $dateCol))) { throw 'Дата, сумма и возврат должны быть в разных столбцах.' }
    if ($first -gt $last) { throw 'Первая строка должна быть не больше последней.' }
    $records=[Collections.Generic.List[object]]::new(); $audit=[Collections.Generic.List[object]]::new()
    foreach ($row in $data.Rows) {
        $reason=''; $level='Справочно'; $dateValue=$row.Cells[$dateCol]; $amountValue=$row.Cells[$amountCol]; $refundValue=$null
        if ($refundCol) { $refundValue=$row.Cells[$refundCol] }
        if ($row.Number -lt $first -or $row.Number -gt $last) { $reason='Вне выбранных строк' }
        else {
            try {
                $rowText = ($row.Cells.Values | ForEach-Object { [string]$_ }) -join ' | '
                if ($rowText -match '(?i)(^|\|\s*)(итого|всего|остаток на|нарастающим итогом)(\s|:|$)') { throw 'Строка итогов или остатка, не включена' }
                $dValue=$dateValue
                if ($data.Date1904 -and ($dValue -is [double] -or $dValue -is [decimal] -or $dValue -is [int])) { $dValue=[double]$dValue+1462 }
                $date = Read-Date $dValue
                $amount = Money $amountValue
                $refund = if ($refundCol) { Money $refundValue } else { $null }
                if ($null -eq $amount -and $null -eq $refund) { throw 'Нет суммы в выбранных столбцах' }
                if ($null -eq $amount) { $amount=[decimal]0 }
                if ($null -ne $refund) { $amount -= [math]::Abs($refund) }
                $records.Add([pscustomobject]@{ Date=$date.ToString('yyyy-MM-dd'); Amount=[decimal]$amount; Row=$row.Number })
                continue
            } catch {
                $reason=$_.Exception.Message; $level='Проверить'
                $text = ($row.Cells.Values | ForEach-Object { [string]$_ }) -join ' | '
                if ($text -match '(?i)(^|\|\s*)(итого|всего|остаток на|нарастающим итогом)(\s|:|$)') { $reason='Строка итогов или остатка, не включена'; $level='Справочно' }
            }
        }
        $audit.Add([pscustomobject]@{ File=$data.Path; Sheet=$data.Sheet; Row=$row.Number; Level=$level; Reason=$reason; Date=[string]$dateValue; Amount=[string]$amountValue; Refund=[string]$refundValue })
    }
    if ($records.Count -eq 0) { throw "На листе «$($data.Sheet)» не найдено ни одной записи. Проверьте столбцы и номера строк." }
    return [pscustomobject]@{ Records=$records.ToArray(); Audit=$audit.ToArray(); Data=$data; DateCol=$dateCol; AmountCol=$amountCol; RefundCol=$refundCol; First=$first; Last=$last }
}

function Compare-Records($k, $o) {
    $kg=@{}; $og=@{}; $kr=@{}; $orows=@{}
    foreach ($r in $k) { if (-not $kg.ContainsKey($r.Date)) { $kg[$r.Date]=[decimal]0; $kr[$r.Date]=[Collections.Generic.List[int]]::new() }; $kg[$r.Date]+=$r.Amount; $kr[$r.Date].Add($r.Row) }
    foreach ($r in $o) { if (-not $og.ContainsKey($r.Date)) { $og[$r.Date]=[decimal]0; $orows[$r.Date]=[Collections.Generic.List[int]]::new() }; $og[$r.Date]+=$r.Amount; $orows[$r.Date].Add($r.Row) }
    $result=[Collections.Generic.List[object]]::new()
    foreach ($date in (@(@($kg.Keys)+@($og.Keys)) | Sort-Object -Unique)) {
        $kv=[decimal]$kg[$date]; $ov=[decimal]$og[$date]; $diff=$ov-$kv
        $status = if (-not $kg.ContainsKey($date)) { 'Нет записей КУДиР за дату' } elseif (-not $og.ContainsKey($date)) { 'Нет записей ОФД за дату' } elseif ($diff -gt 0) { 'В КУДиР меньше: проверить пропуск' } elseif ($diff -lt 0) { 'В КУДиР больше: проверить причину' } else { 'Суммы за день совпали' }
        $result.Add([pscustomobject]@{ Date=$date; Kudir=$kv; Ofd=$ov; Difference=$diff; Status=$status; KRows=($kr[$date] -join ', '); ORows=($orows[$date] -join ', ') })
    }
    return ,$result.ToArray()
}

function Html($s) { return [Net.WebUtility]::HtmlEncode([string]$s) }
function Fmt($n) { return ([decimal]$n).ToString('N2',$script:Ru) }
function Write-CsvSafe($path, $rows, [string[]]$headers) {
    $lines=[Collections.Generic.List[string]]::new()
    $lines.Add(($headers -join ';'))
    foreach ($r in $rows) {
        $parts = foreach ($h in $headers) {
            $v=$r.$h
            if ($v -is [decimal] -or $v -is [double]) { $s=$v.ToString('0.00',$script:Ru) }
            else { $s=[string]$v; if ($s -match '^\s*[=+@-]') { $s="'"+$s } }
            '"' + $s.Replace('"','""') + '"'
        }
        $lines.Add(($parts -join ';'))
    }
    [IO.File]::WriteAllLines($path,$lines,[Text.UTF8Encoding]::new($true))
}

function Write-Report($k, $o, [string]$folder) {
    $days = Compare-Records $k.Records $o.Records
    $kt=[decimal]0; $ot=[decimal]0
    foreach ($r in $k.Records) { $kt+=$r.Amount }; foreach ($r in $o.Records) { $ot+=$r.Amount }
    $audit=@($k.Audit)+@($o.Audit)
    $errors=@($audit | Where-Object Level -eq 'Проверить').Count
    $mismatch=@($days | Where-Object { $_.Difference -ne 0 -or $_.Status -ne 'Суммы за день совпали' })
    $body=[Text.StringBuilder]::new()
    [void]$body.Append('<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Сверка КУДиР и ОФД</title><style>body{font:16px Segoe UI,Arial;margin:32px;color:#172b3a;max-width:1300px}h1{font-size:30px}h2{margin-top:32px}table{border-collapse:collapse;width:100%;font-size:14px}th,td{padding:10px;border:1px solid #cbd5df;text-align:left;vertical-align:top}th{background:#e9f0f5}td.num{text-align:right;white-space:nowrap}.warn{padding:18px;background:#fff0ce;border-left:5px solid #b57700}.card{padding:18px;background:#edf5fa}.muted{color:#4e6270}details{margin:20px 0}a{color:#125d92}@media print{body{margin:12px}details{display:block}}</style><h1>Сверка КУДиР и ОФД</h1>')
    $state=if ($errors) { 'Сверка неполная: есть непрочитанные строки' } elseif ($mismatch.Count) { 'Есть расхождения по датам' } else { 'Суммы по выбранным датам совпали' }
    [void]$body.Append("<h2>$(Html $state)</h2><div class='card'>КУДиР: <b>$(Fmt $kt) ₽</b><br>ОФД: <b>$(Fmt $ot) ₽</b><br>ОФД минус КУДиР: <b>$(Fmt ($ot-$kt)) ₽</b><br>Дат для проверки: $($mismatch.Count). Непрочитанных строк: $errors.</div>")
    [void]$body.Append('<p>Итоги рассчитаны по принятым строкам выбранных листов. Печатные строки «Итого» отдельно не сверяются. Положительная разница означает, что в КУДиР меньше; отрицательная — больше.</p><p class="warn">Это сверка сумм по дням. Отсутствие записи или разница — повод проверить документы, а не доказательство пропущенного дохода. Совпадение сумм не подтверждает совпадение каждого чека. Сравнивайте одну организацию, один период и одинаковый состав поступлений; кассы, возвраты и даты отражения должны учитываться одинаково.</p>')
    foreach ($source in @(@{Name='КУДиР';Value=$k},@{Name='ОФД';Value=$o})) {
        $v=$source.Value; $dates=@($v.Records.Date | Sort-Object)
        [void]$body.Append("<p><b>$($source.Name)</b>: $(Html $v.Data.Path)<br>Лист: $(Html $v.Data.Sheet). Строки: $($v.First)–$($v.Last). Дата: $($v.DateCol), сумма: $($v.AmountCol), возвраты: $(Html $v.RefundCol).<br>Принято: $($v.Records.Count) строк; даты: $($dates[0]) — $($dates[-1]).</p>")
    }
    [void]$body.Append('<p>Если выбран отдельный столбец возвратов, из суммы вычтен модуль возврата. Уже уменьшенную на возвраты сумму нельзя уменьшать повторно. Числа округлены до копеек в каждой строке. Скрытые и отфильтрованные строки внутри диапазона тоже учитываются.</p>')
    [void]$body.Append('<h2>Даты для проверки</h2><table><tr><th>Дата</th><th>КУДиР, ₽</th><th>ОФД, ₽</th><th>Разница, ₽</th><th>Результат</th><th>Строки КУДиР</th><th>Строки ОФД</th></tr>')
    foreach ($d in $mismatch) { [void]$body.Append("<tr><td>$($d.Date)</td><td class='num'>$(Fmt $d.Kudir)</td><td class='num'>$(Fmt $d.Ofd)</td><td class='num'>$(Fmt $d.Difference)</td><td>$(Html $d.Status)</td><td>$(Html $d.KRows)</td><td>$(Html $d.ORows)</td></tr>") }
    [void]$body.Append('</table>')
    if (-not $mismatch.Count) { [void]$body.Append('<p>Расхождений среди прочитанных строк нет.</p>') }
    [void]$body.Append('<h2>Неучтённые строки внутри выбранного диапазона</h2><table><tr><th>Файл / лист</th><th>Строка</th><th>Причина</th><th>Дата в файле</th><th>Сумма в файле</th></tr>')
    foreach ($a in ($audit | Where-Object Reason -ne 'Вне выбранных строк')) { [void]$body.Append("<tr><td>$(Html ([IO.Path]::GetFileName($a.File))) / $(Html $a.Sheet)</td><td>$($a.Row)</td><td>$(Html $a.Reason)</td><td>$(Html $a.Date)</td><td>$(Html $a.Amount)</td></tr>") }
    [void]$body.Append('</table><h2>Файлы для подробной проверки</h2><p><a href="days.csv">Все дни</a> · <a href="records.csv">Все принятые строки</a> · <a href="skipped.csv">Все исключённые строки, включая строки вне диапазона</a></p><p class="muted">CSV открываются в Excel. Исходные файлы не изменены. Сверка использует значения, прочитанные Excel; перед запуском пересчитайте и сохраните исходные формулы в Excel. Другие листы автоматически не объединяются.</p></html>')
    [void][IO.Directory]::CreateDirectory($folder)
    [IO.File]::WriteAllText((Join-Path $folder 'Отчёт.html'),$body.ToString(),[Text.UTF8Encoding]::new($true))
    $csvDays=foreach ($d in $days) { [pscustomobject]@{'Дата'=$d.Date;'КУДиР'=$d.Kudir;'ОФД'=$d.Ofd;'ОФД минус КУДиР'=$d.Difference;'Результат'=$d.Status;'Строки КУДиР'=$d.KRows;'Строки ОФД'=$d.ORows} }
    Write-CsvSafe (Join-Path $folder 'days.csv') $csvDays @('Дата','КУДиР','ОФД','ОФД минус КУДиР','Результат','Строки КУДиР','Строки ОФД')
    $csvRecords=foreach ($src in @(@{Name='КУДиР';Value=$k},@{Name='ОФД';Value=$o})) { foreach ($r in $src.Value.Records) { [pscustomobject]@{'Источник'=$src.Name;'Файл'=$src.Value.Data.Path;'Лист'=$src.Value.Data.Sheet;'Строка'=$r.Row;'Дата'=$r.Date;'Сумма'=$r.Amount} } }
    Write-CsvSafe (Join-Path $folder 'records.csv') $csvRecords @('Источник','Файл','Лист','Строка','Дата','Сумма')
    $csvAudit=foreach ($a in $audit) { [pscustomobject]@{'Файл'=$a.File;'Лист'=$a.Sheet;'Строка'=$a.Row;'Важность'=$a.Level;'Причина'=$a.Reason;'Дата'=$a.Date;'Сумма'=$a.Amount;'Возврат'=$a.Refund} }
    Write-CsvSafe (Join-Path $folder 'skipped.csv') $csvAudit @('Файл','Лист','Строка','Важность','Причина','Дата','Сумма','Возврат')
    return (Join-Path $folder 'Отчёт.html')
}

. (Join-Path $PSScriptRoot 'auto-compare.ps1')
if ($NoGui) { return }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$form=New-Object Windows.Forms.Form
$form.Text='Сверка КУДиР и ОФД'; $form.Size=New-Object Drawing.Size(1100,800); $form.MinimumSize=New-Object Drawing.Size(1000,740); $form.StartPosition='CenterScreen'; $form.Font=New-Object Drawing.Font('Segoe UI',10)

function Add-Label($parent,$text,$x,$y,$w=400) { $c=New-Object Windows.Forms.Label; $c.Text=$text; $c.SetBounds($x,$y,$w,26); $parent.Controls.Add($c); return $c }
function Add-Button($parent,$text,$x,$y,$w=170) { $c=New-Object Windows.Forms.Button; $c.Text=$text; $c.SetBounds($x,$y,$w,34); $parent.Controls.Add($c); return $c }
function Add-Combo($parent,$x,$y,$w=180) { $c=New-Object Windows.Forms.ComboBox; $c.DropDownStyle='DropDownList'; $c.SetBounds($x,$y,$w,30); $parent.Controls.Add($c); return $c }

[void](Add-Label $form 'Выберите два файла. Для КУДиР по кварталам и сверки ОФД нажмите «Автосверка». Для иных таблиц — ручной режим.' 16 12 1040)
$tabs=New-Object Windows.Forms.TabControl; $tabs.SetBounds(16,46,1050,620); $tabs.Anchor='Top,Bottom,Left,Right'; $form.Controls.Add($tabs)
$script:panels=@()
foreach ($name in @('КУДиР','ОФД')) {
    $tab=New-Object Windows.Forms.TabPage; $tab.Text=$name; $tabs.TabPages.Add($tab)
    $choose=Add-Button $tab "Выбрать файл $name" 12 12 210
    $path=Add-Label $tab 'Файл не выбран' 234 18 780; $path.AutoEllipsis=$true
    [void](Add-Label $tab 'Лист' 12 57 60); $sheetBox=Add-Combo $tab 74 54 430
    $load=Add-Button $tab 'Показать лист' 520 52 190
    [void](Add-Label $tab 'Столбец даты' 12 98 180); [void](Add-Label $tab 'Столбец суммы дохода' 212 98 220); [void](Add-Label $tab 'Возвраты (если отдельно)' 452 98 255)
    $dc=Add-Combo $tab 12 126 180; $ac=Add-Combo $tab 212 126 220; $rc=Add-Combo $tab 452 126 240
    [void](Add-Label $tab 'Первая строка данных' 12 164 190); [void](Add-Label $tab 'Последняя строка данных' 242 164 240)
    $start=New-Object Windows.Forms.NumericUpDown; $start.Minimum=1; $start.Maximum=1048576; $start.SetBounds(12,193,190,30); $tab.Controls.Add($start)
    $end=New-Object Windows.Forms.NumericUpDown; $end.Minimum=1; $end.Maximum=1048576; $end.SetBounds(242,193,190,30); $tab.Controls.Add($end)
    $hint=Add-Label $tab 'Выберите столбцы по буквам Excel. Даты в просмотре могут быть числами — это нормально.' 12 238 1010
    $grid=New-Object Windows.Forms.DataGridView; $grid.SetBounds(12,270,1012,288); $grid.Anchor='Top,Bottom,Left,Right'; $grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.RowHeadersVisible=$false; $grid.AutoSizeColumnsMode='None'; $grid.SelectionMode='FullRowSelect'; $tab.Controls.Add($grid)
    $p=@{Name=$name;Tab=$tab;Choose=$choose;Path=$path;Sheet=$sheetBox;Load=$load;Date=$dc;Amount=$ac;Refund=$rc;First=$start;Last=$end;Grid=$grid;Hint=$hint;Data=$null;File=''}
    $choose.Tag=$p; $load.Tag=$p; $sheetBox.Tag=$p; $script:panels+=,$p
    $sheetBox.Add_SelectedIndexChanged({ $p=$this.Tag; $p.Data=$null; $p.Grid.DataSource=$null })
    $choose.Add_Click({
        $p=$this.Tag; $dialog=New-Object Windows.Forms.OpenFileDialog; $dialog.Filter='Файлы Excel (*.xlsx;*.xls)|*.xlsx;*.xls'; $dialog.Title="Выберите файл $($p.Name)"
        if ($dialog.ShowDialog() -ne 'OK') { return }
        $form.UseWaitCursor=$true
        try {
            $names=Read-Excel $dialog.FileName
            $p.Data=$null; $p.File=$dialog.FileName; $p.Path.Text=$dialog.FileName
            $p.Sheet.Items.Clear(); foreach ($n in $names) { [void]$p.Sheet.Items.Add($n) }; if ($p.Sheet.Items.Count) { $p.Sheet.SelectedIndex=0 }
            $p.Load.PerformClick()
        } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Не удалось открыть файл') } finally { $form.UseWaitCursor=$false; $dialog.Dispose() }
    })
    $load.Add_Click({
        $p=$this.Tag; if (-not $p.File -or $p.Sheet.SelectedIndex -lt 0) { return }
        $form.UseWaitCursor=$true
        try {
            $p.Data=$null
            $data=Read-Excel $p.File ([string]$p.Sheet.SelectedItem)
            foreach ($box in @($p.Date,$p.Amount,$p.Refund)) { $box.Items.Clear() }
            [void]$p.Refund.Items.Add('Не вычитать')
            $table=New-Object Data.DataTable; [void]$table.Columns.Add('Строка',[int])
            for ($c=$data.FirstCol; $c -lt $data.FirstCol+$data.ColCount; $c++) {
                $letter=Column-Name $c
                foreach ($box in @($p.Date,$p.Amount,$p.Refund)) { [void]$box.Items.Add($letter) }
                [void]$table.Columns.Add($letter,[string])
            }
            foreach ($r in ($data.Rows | Select-Object -First 300)) {
                $line=$table.NewRow(); $line['Строка']=$r.Number
                foreach ($key in $r.Cells.Keys) { $line[$key]=[string]$r.Cells[$key] }; $table.Rows.Add($line)
            }
            $p.Grid.DataSource=$table; foreach ($col in $p.Grid.Columns) { $col.SortMode='NotSortable'; $col.Width=155 }; $p.Grid.Columns[0].Width=75
            $p.First.Value=$data.First; $p.Last.Value=$data.Last; $p.Refund.SelectedIndex=0; $p.Data=$data
            $p.Hint.Text="Показаны первые 300 непустых строк. Лист: $($data.First)–$($data.Last). Укажите столбцы и границы данных."
        } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Не удалось прочитать лист') } finally { $form.UseWaitCursor=$false }
    })
}
$auto=Add-Button $form 'Автосверка всех кварталов' 16 684 280; $auto.Anchor='Bottom,Left'
$run=Add-Button $form 'Ручная сверка листов' 308 684 235; $run.Anchor='Bottom,Left'
$help=Add-Button $form 'Инструкция' 555 684 145; $help.Anchor='Bottom,Left'
$status=Add-Label $form 'Исходные файлы не изменяются.' 710 690 350; $status.Anchor='Bottom,Left,Right'
$auto.Add_Click({
    $form.UseWaitCursor=$true; $auto.Enabled=$false; $run.Enabled=$false
    try {
        if (-not $script:panels[0].File -or -not $script:panels[1].File) { throw 'Выберите файл КУДиР на первой вкладке и файл ОФД на второй вкладке.' }
        $status.Text='Чтение всех кварталов и чеков...'; [Windows.Forms.Application]::DoEvents()
        $inputs=Read-AutoInputs $script:panels[0].File $script:panels[1].File
        $result=Get-AutoComparison $inputs.K $inputs.O
        $dialog=New-Object Windows.Forms.FolderBrowserDialog; $dialog.Description='Выберите папку для отчёта автосверки'; $dialog.SelectedPath=[Environment]::GetFolderPath('MyDocuments')
        if ($dialog.ShowDialog() -ne 'OK') { return }
        $folder=Join-Path $dialog.SelectedPath ('Автосверка_'+(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,6))
        $report=Write-AutoReport $result $folder
        $status.Text='Готово. Отчёт сохранён.'; Start-Process $report
    } catch { $status.Text='Проверьте сообщение об ошибке.'; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Автосверка') }
    finally { $form.UseWaitCursor=$false; $auto.Enabled=$true; $run.Enabled=$true }
})
$help.Add_Click({ Start-Process (Join-Path $PSScriptRoot 'Инструкция.html') })
$run.Add_Click({
    $form.UseWaitCursor=$true; $run.Enabled=$false
    try {
        $sources=@()
        foreach ($p in $script:panels) {
            if (-not $p.Data) { throw "Выберите файл и нажмите «Показать лист» на вкладке $($p.Name)." }
            if ($p.Date.SelectedIndex -lt 0 -or $p.Amount.SelectedIndex -lt 0) { throw "Выберите столбцы даты и суммы на вкладке $($p.Name)." }
            $refund=if ($p.Refund.SelectedIndex -gt 0) { [string]$p.Refund.SelectedItem } else { '' }
            $sources+=Extract-Records $p.Data ([string]$p.Date.SelectedItem) ([string]$p.Amount.SelectedItem) $refund ([int]$p.First.Value) ([int]$p.Last.Value)
        }
        $dialog=New-Object Windows.Forms.FolderBrowserDialog; $dialog.Description='Выберите папку для отчёта. Внутри будет создана новая папка сверки.'; $dialog.SelectedPath=[Environment]::GetFolderPath('MyDocuments')
        if ($dialog.ShowDialog() -ne 'OK') { return }
        $folder=Join-Path $dialog.SelectedPath ('Сверка_'+(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,6))
        $report=Write-Report $sources[0] $sources[1] $folder
        $status.Text='Готово. Отчёт сохранён и открыт в браузере.'
        Start-Process $report
    } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Проверьте настройки') } finally { $form.UseWaitCursor=$false; $run.Enabled=$true }
})
if ($env:KUDIR_UI_CHECK -eq '1') {
    $form.Show()
    [Windows.Forms.Application]::DoEvents()
    $p=$script:panels[0]
    $p.File=Join-Path $PSScriptRoot 'test-output/fixture.xlsx'
    $p.Path.Text=$p.File
    [void]$p.Sheet.Items.Add('Данные'); $p.Sheet.SelectedIndex=0
    $p.Load.PerformClick()
    $p.Date.SelectedItem='A'; $p.Amount.SelectedItem='B'; $p.First.Value=2
    [Windows.Forms.Application]::DoEvents()
    $bitmap=New-Object Drawing.Bitmap($form.Width,$form.Height)
    $form.DrawToBitmap($bitmap,(New-Object Drawing.Rectangle(0,0,$form.Width,$form.Height)))
    $bitmap.Save((Join-Path $PSScriptRoot 'test-output/window.png'))
    $bitmap.Dispose()
    $form.Close()
} else { [void]$form.ShowDialog() }
$form.Dispose()
