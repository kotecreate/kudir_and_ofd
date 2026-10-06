# Automatic reconciliation for the verified KUDIR quarterly and OFD document layouts.
function Sum-Field($items,[string]$field) {
    $total=[decimal]0; foreach ($item in $items) { $total += [decimal]$item.$field }; return $total
}
function Auto-Table($rows,[string[]]$columns) {
    $b=[Text.StringBuilder]::new(); [void]$b.Append('<div class="table-scroll" tabindex="0" role="region" aria-label="Таблица, при необходимости прокрутите вправо"><table><thead><tr>')
    foreach ($c in $columns) { $cls=if ($c -match '^(Дата|Месяц|КУДиР$|ОФД$|Разница$|Сумма$|Наличные$|Безналичные$|Печатный итог$|По операциям$)') { 'nowrap' } else { '' }; [void]$b.Append('<th class="'+$cls+'">'+(Html $c)+'</th>') }; [void]$b.Append('</tr></thead><tbody>')
    foreach ($r in $rows) {
        [void]$b.Append('<tr>')
        foreach ($c in $columns) { $v=$r.$c; $cls=''; if ($v -is [decimal] -or $v -is [double] -or $v -is [int] -or $v -is [long]) { $v=Fmt $v; $cls='money' } elseif ($c -match '^(Дата|Месяц)') { $cls='nowrap' }; [void]$b.Append('<td class="'+$cls+'">'+(Html $v)+'</td>') }
        [void]$b.Append('</tr>')
    }
    [void]$b.Append('</tbody></table></div>'); return $b.ToString()
}

function Display-Date([string]$date) { return [datetime]::ParseExact($date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).ToString('dd.MM.yyyy') }

function Get-BankExplanation($r,$day) {
    $date=$day.'Дата продажи'; $diff=[decimal]$day.Разница
    $kr=@($r.K | Where-Object { $_.Date -eq $date -and $_.Kind -in @('Карты','QR') } | Sort-Object Kind,DocDate,Ref)
    $qr=@($kr | Where-Object Kind -eq 'QR'); $cards=@($kr | Where-Object Kind -eq 'Карты')
    $orr=@($r.O | Where-Object { $_.Date -eq $date -and $_.Sign -ne 0 -and $_.Bank -ne 0 })
    $special=@($r.O | Where-Object { $_.Date -eq $date -and $_.Sign -eq 0 -and $_.Bank -ne 0 })
    $first=''; $hint=''; $conditional=$null; $target=''
    if ($special.Count) {
        $first='Сначала проверьте типы отдельных чеков ОФД'
        $hint='На этот день есть «Расход» или «Возврат расхода». Они показаны ниже отдельно и не включены в продажи. Сверьте назначение этих чеков с кассовыми документами. Пока оно не выяснено, менять КУДиР только на сумму разницы нельзя.'
    } elseif ($cards.Count -eq 1 -and ([decimal]$day.ОФД-(Sum-Field $qr Amount)) -ge 0) {
        $target=$cards[0].Ref
        $conditional=[decimal]$day.ОФД-(Sum-Field $qr Amount)
        $first="Начните с расшифровки зачисления: $target, сумма $(Fmt $cards[0].Amount) ₽"
        $hint="Это единственное зачисление по картам за этот день продажи. Дата документа в КУДиР — $(Display-Date $cards[0].DocDate), дата продажи из описания — $(Display-Date $date). Откройте лист «$($cards[0].Sheet)»: сумма в D$($cards[0].Row), дата документа в B$($cards[0].Row), назначение в C$($cards[0].Row). Проверьте банковскую расшифровку этого зачисления и QR-поступления ниже. Это место для начала проверки, а не доказанная ошибочная строка."
    } elseif ($kr.Count -eq 0) {
        $first='За эту дату продажи не найдено подходящих записей КУДиР'
        $hint='Проверьте показанные ниже чеки и банковские зачисления следующих рабочих дней. Программа уже искала дату продажи в назначении всех кварталов файла. Возможны отсутствующая запись, другое назначение платежа или зачисление, которого ещё нет в выгрузке.'
    } else {
        $first='Проверьте состав записей КУДиР за день продажи'
        $hint='Ниже перечислены отдельные QR-поступления и все зачисления по картам с обеими датами. Сверьте их с банковской расшифровкой. По сумме за день нельзя доказать, какая из этих строк ошибочна; программа не распределяет разницу между ними произвольно.'
    }
    return [pscustomobject]@{Date=$date;K=$kr;QR=$qr;Cards=$cards;O=$orr;Special=$special;First=$first;Hint=$hint;Conditional=$conditional;Target=$target}
}

function Write-BankDetails($r,$days) {
    $b=[Text.StringBuilder]::new()
    [void]$b.Append('<h2>Разбор каждой разницы</h2><p>Дата продажи — когда оплатил покупатель. Дата документа КУДиР — когда отражено поступление. При оплате картой они могут различаться. Программа уже использует дату продажи из назначения платежа; переносить запись вручную для этой сверки не нужно.</p>')
    foreach ($day in $days) {
        $e=Get-BankExplanation $r $day; $date=$e.Date
        $direction=if ($day.Разница -lt 0) { 'В КУДиР больше' } else { 'В КУДиР меньше' }
        [void]$b.Append("<section class='case' id='day-$date'><h3>Продажи $(Display-Date $date): $direction на <span class='nowrap'>$(Fmt ([math]::Abs($day.Разница))) ₽</span></h3><div class='action'><b>$(Html $e.First)</b><p>$(Html $e.Hint)</p></div>")
        $parts=@($e.K | ForEach-Object { (Fmt $_.Amount)+' ₽' })
        $equation=if ($parts.Count) { ($parts -join ' + ')+' = '+(Fmt $day.КУДиР)+' ₽' } else { 'Записей нет; в расчёте 0,00 ₽' }
        [void]$b.Append("<p><b>Что сравнили:</b> КУДиР: $(Html $equation). ОФД: $(Fmt $day.ОФД) ₽ безналичными по $($e.O.Count) чекам с ненулевой безналичной суммой. Наличные и служебные строки сюда не входят.</p>")
        if ($day.Разница -lt 0) { [void]$b.Append('<p>Отдельного чека ОФД на сумму разницы может не быть: превышение может находиться внутри общей банковской записи или складываться из нескольких операций.</p>') }
        [void]$b.Append('<h4>Записи КУДиР, включённые именно в этот день продажи</h4>')
        $krows=@($e.K | ForEach-Object {
            $why=if ($_.Kind -eq 'Карты') { 'Дата продажи взята из «за '+$_.Date+'» в назначении. Дата документа не использовалась для группировки.' } else { 'Дата операции взята после «от» в описании QR.' }
            [pscustomobject]@{'Где открыть'="Лист «$($_.Sheet)», строка $($_.Row), сумма D$($_.Row)";'Вид'=$_.Kind;'Дата документа'=(Display-Date $_.DocDate);'Дата продажи'=(Display-Date $_.Date);'Сумма'=$_.Amount;'Почему здесь'=$why}
        })
        [void]$b.Append((Auto-Table $krows @('Где открыть','Вид','Дата документа','Дата продажи','Сумма','Почему здесь')))
        if ($null -ne $e.Conditional) {
            $qs=Sum-Field $e.QR Amount
            [void]$b.Append("<p class='conditional'><b>Ориентир, если QR-поступления верны и относятся к этим чекам:</b> из суммы ОФД $(Fmt $day.ОФД) ₽ вычитаем QR $(Fmt $qs) ₽. На зачисление по картам остаётся $(Fmt $e.Conditional) ₽; в $(Html $e.Target) записано $(Fmt $e.Cards[0].Amount) ₽. Это условный расчёт для проверки банковской расшифровки, <b>не готовая сумма для исправления</b>. Выгрузка ОФД не разделяет безналичные оплаты на карты и QR, поэтому подтвердить это распределение по ней нельзя.</p>")
        }
        [void]$b.Append('<details><summary>Открыть полные назначения платежей КУДиР</summary>')
        $desc=@($e.K | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Описание'=$_.Description} })
        [void]$b.Append((Auto-Table $desc @('Строка','Описание'))); [void]$b.Append('</details>')
        if ($e.Special.Count) {
            [void]$b.Append('<h4>Отдельные операции ОФД: проверить первыми</h4>')
            $special=@($e.Special | ForEach-Object { [pscustomobject]@{'Где открыть'=$_.Ref;'Операция'=$_.Kind;'Безналичные'=$_.Bank;'ФН / ФД'=$_.Id} })
            [void]$b.Append((Auto-Table $special @('Где открыть','Операция','Безналичные','ФН / ФД')))
            [void]$b.Append('<p>В исходном ОФД тип операции находится в Q, безналичная сумма — в T. Совпадение этой суммы с разницей само по себе не подтверждает, как нужно исправить учёт.</p>')
        }
        [void]$b.Append("<details><summary>Показать $($e.O.Count) чеков ОФД, вошедших в безналичную сумму за день</summary><p>Открывайте указанные строки листа ОФД; сравнивайте столбец <b>T «Сумма безналичными»</b>. Столбец R — весь чек и может включать наличные. Для возврата прихода безналичная сумма вычитается.</p>")
        $orows=@($e.O | ForEach-Object { [pscustomobject]@{'Строка ОФД'=$_.Ref;'Операция'=$_.Kind;'Безналичные'=$_.Bank;'Учтено в сумме'=([decimal]$_.Bank*$_.Sign);'ФН / ФД'=$_.Id} })
        [void]$b.Append((Auto-Table $orows @('Строка ОФД','Операция','Безналичные','Учтено в сумме','ФН / ФД')))
        [void]$b.Append('</details><p><a href="#bank-overview">Вернуться к списку дат ↑</a></p></section>')
    }
    return $b.ToString()
}

function Read-AutoInputs([string]$kpath,[string]$opath) {
    if ([IO.Path]::GetFullPath($kpath) -eq [IO.Path]::GetFullPath($opath)) { throw 'Для КУДиР и ОФД выбран один и тот же файл.' }
    $knames=Read-Excel $kpath
    $quarters=@($knames | Where-Object { $_ -match '^Квартал-\d+$' })
    if (-not $quarters.Count) { throw 'Автосверка ожидает листы Квартал-1, Квартал-2 и т. д. Для другого формата используйте ручную сверку.' }
    $kdata=@(); foreach ($name in $quarters) { $kdata+=Read-Excel $kpath $name }
    $onames=Read-Excel $opath
    if ($onames.Count -ne 1) { throw 'Автосверка ожидает один лист ОФД. Для другого формата используйте ручную сверку.' }
    $odata=Read-Excel $opath $onames[0]
    return [pscustomobject]@{ K=$kdata; O=$odata }
}

function Get-AutoComparison($kdata,$odata) {
    $headers=@($odata.Rows | Where-Object { $_.Cells['K'] -eq 'Дата создания документа' -and $_.Cells['Q'] -eq 'Тип операции' -and $_.Cells['R'] -eq 'Сумма документа, руб.' -and $_.Cells['S'] -eq 'Сумма наличными, руб.' -and $_.Cells['T'] -eq 'Сумма безналичными, руб.' })
    if ($headers.Count -ne 1) { throw 'Не распознаны столбцы ОФД K, Q, R, S, T. Используйте ручную сверку для другого формата.' }
    $issues=[Collections.Generic.List[object]]::new(); $audit=[Collections.Generic.List[object]]::new()
    $k=[Collections.Generic.List[object]]::new(); $o=[Collections.Generic.List[object]]::new(); $controls=[Collections.Generic.List[object]]::new()
    foreach ($data in $kdata) {
        $header=@($data.Rows | Where-Object { $_.Cells['B'] -like '*Дата и номер первичного документа*' -and $_.Cells['D'] -like 'Доходы*' })
        if ($header.Count -ne 1) { throw "Не распознаны столбцы КУДиР B, C, D на листе $($data.Sheet)." }
        foreach ($row in $data.Rows) {
            $ref="$($data.Sheet)!$($row.Number)"
            if ($row.Number -le $header[0].Number+1) { continue }
            if ([string]$row.Cells['A'] -like 'Итого*') {
                $controls.Add([pscustomobject]@{Source='КУДиР';Ref=$ref;Label=[string]$row.Cells['A'];Value=(Money $row.Cells['D'])})
                continue
            }
            try {
                if (-not $row.Cells.ContainsKey('B') -and -not $row.Cells.ContainsKey('D')) { continue }
                $doc=(Read-Date $row.Cells['B']).ToString('yyyy-MM-dd')
                $amount=Money $row.Cells['D']; if ($null -eq $amount) { throw 'Пустая сумма дохода' }
                $desc=[string]$row.Cells['C']; $date=$doc; $kind='Не распознано'; $from=''; $to=''
                if ($desc -match '^Оплата наличными') {
                    $ds=[regex]::Matches($desc,'\d{2}\.\d{2}\.\d{4}')
                    if ($ds.Count -ne 2) { throw 'У наличной суммы не найден точный период из двух дат' }
                    $from=(Read-Date $ds[0].Value).ToString('yyyy-MM-dd'); $to=(Read-Date $ds[1].Value).ToString('yyyy-MM-dd')
                    if ($from -gt $to -or $from.Substring(0,7) -ne $to.Substring(0,7)) { throw 'Наличный период должен лежать внутри одного месяца' }
                    $date=$to; $kind='Наличные'
                } elseif ($desc -match 'по терминалу.+?\sза\s(\d{4}-\d{2}-\d{2})') {
                    $date=(Read-Date $Matches[1]).ToString('yyyy-MM-dd'); $kind='Карты'
                } elseif ($desc -match '^Зачисление по QR коду') {
                    $ds=[regex]::Matches($desc,'\d{2}\.\d{2}\.\d{4}')
                    if ($ds.Count -ne 1) { throw 'У QR-поступления не найдена одна дата операции' }
                    $date=(Read-Date $ds[0].Value).ToString('yyyy-MM-dd'); $kind='QR'
                } elseif ($desc -match '^Выплата Гарантии качества') { $kind='Прочее поступление' }
                else { $issues.Add([pscustomobject]@{Источник=$ref;Причина='Не распознан вид поступления; исключено из сопоставимых продаж';Описание=$desc}) }
                $k.Add([pscustomobject]@{Date=$date;DocDate=$doc;Amount=$amount;Kind=$kind;Ref=$ref;Sheet=$data.Sheet;Row=$row.Number;Description=$desc;From=$from;To=$to})
            } catch { $issues.Add([pscustomobject]@{Источник=$ref;Причина=$_.Exception.Message;Описание=[string]$row.Cells['C']}) }
        }
    }
    $seen=@{}
    foreach ($row in $odata.Rows) {
        if ($row.Number -le $headers[0].Number) { continue }
        $ref="$($odata.Sheet)!$($row.Number)"; $type=[string]$row.Cells['N']; $op=[string]$row.Cells['Q']
        if ($op -like 'Итого*') {
            if ($op -like 'Итого всего*') { $controls.Add([pscustomobject]@{Source='ОФД';Ref=$ref;Label=$op;Value=(Money $row.Cells['R'])}) }
            continue
        }
        if ($type -notmatch '^Кассовый чек') {
            $audit.Add([pscustomobject]@{Источник=$ref;Причина='Служебный документ ОФД';Описание=$type})
            if ((Money $row.Cells['R']) -ne 0) { $issues.Add([pscustomobject]@{Источник=$ref;Причина='Ненулевая сумма у служебного документа';Описание=$type}) }
            continue
        }
        try {
            $dateValue=$row.Cells['K']; if ($odata.Date1904) { $dateValue=[double]$dateValue+1462 }
            $date=(Read-Date $dateValue).ToString('yyyy-MM-dd')
            $amount=Money $row.Cells['R']; $cash=Money $row.Cells['S']; $bank=Money $row.Cells['T']
            if ($null -eq $amount -or $null -eq $cash -or $null -eq $bank) { throw 'Пустая сумма документа или вида оплаты' }
            if ($amount -lt 0 -or $cash -lt 0 -or $bank -lt 0) { throw 'Автосверка ожидает положительные суммы и знак из типа операции' }
            if ($cash+$bank -ne $amount) { throw 'Сумма документа отличается от наличных плюс безналичных; есть другие виды оплаты' }
            $id="$($row.Cells['D'])/$($row.Cells['J'])"
            if ($id -eq '/' -or $seen.ContainsKey($id)) { throw "Пустой или повторный ключ ФН/ФД: $id" }; $seen[$id]=$true
            $sign=if ($op -eq 'Приход') { 1 } elseif ($op -eq 'Возврат прихода') { -1 } else { 0 }
            if ($op -notin @('Приход','Возврат прихода','Расход','Возврат расхода')) { throw "Неизвестный тип операции: $op" }
            if ($type -ne 'Кассовый чек') { $issues.Add([pscustomobject]@{Источник=$ref;Причина='Чек коррекции: проверить отдельно';Описание=$type}) }
            $o.Add([pscustomobject]@{Date=$date;Amount=$amount;Cash=$cash;Bank=$bank;Kind=$op;Sign=$sign;Ref=$ref;Row=$row.Number;Id=$id})
        } catch { $issues.Add([pscustomobject]@{Источник=$ref;Причина=$_.Exception.Message;Описание=$op}) }
    }
    if (-not $k.Count -or -not $o.Count) { throw 'В одном из файлов не найдены операции.' }
    $ordered=@($o.Date | Sort-Object); $start=$ordered[0]; $end=$ordered[-1]
    # Reconcile using the export's stated period, not merely its first/last sale.
    $periodCells=@($odata.Rows | Where-Object { $_.Number -lt $headers[0].Number } | ForEach-Object { $_.Cells.Values })
    $periodFrom=@($periodCells | Where-Object { $_ -is [string] -and $_ -match '^с \d{2}\.\d{2}\.\d{4}$' })
    $periodTo=@($periodCells | Where-Object { $_ -is [string] -and $_ -match '^по \d{2}\.\d{2}\.\d{4}$' })
    if ($periodFrom.Count -eq 1 -and $periodTo.Count -eq 1) { $start=(Read-Date $periodFrom[0]).ToString('yyyy-MM-dd'); $end=(Read-Date $periodTo[0]).ToString('yyyy-MM-dd') }
    if ($start -gt $end -or @($o | Where-Object { $_.Date -lt $start -or $_.Date -gt $end }).Count) { throw 'Даты чеков не соответствуют периоду отчёта ОФД.' }
    $inPeriod=@($k | Where-Object { $_.Date -ge $start -and $_.Date -le $end })
    $sales=@($inPeriod | Where-Object Kind -in @('Наличные','Карты','QR'))
    $ksBank=@($sales | Where-Object Kind -in @('Карты','QR')); $osSales=@($o | Where-Object Sign -ne 0)
    $kRecords=@($ksBank | ForEach-Object { [pscustomobject]@{Date=$_.Date;Amount=$_.Amount;Row=$_.Row} })
    $oRecords=@($osSales | Where-Object Bank -ne 0 | ForEach-Object { [pscustomobject]@{Date=$_.Date;Amount=($_.Bank*$_.Sign);Row=$_.Row} })
    # Zero-only dates do not indicate a missing bank payment.
    $bankDays=Compare-Records $kRecords $oRecords
    $bankRows=@(foreach ($d in $bankDays) {
        $refs=@($ksBank | Where-Object Date -eq $d.Date | ForEach-Object Ref)
        $expenseReturn=Sum-Field @($o | Where-Object { $_.Date -eq $d.Date -and $_.Kind -eq 'Возврат расхода' }) Bank
        $note=if ($d.Difference -gt 0 -and $d.Difference -eq $expenseReturn) { 'Разница равна «Возврату расхода». Проверьте тип чека; пропуск дохода не подтверждён.' } elseif ($expenseReturn -ne 0) { 'За дату есть «Возврат расхода»: проверить вместе с расхождением.' } elseif ($d.Difference -gt 0) { 'В КУДиР меньше: проверить документы.' } elseif ($d.Difference -lt 0) { 'В КУДиР больше: проверить документы.' } else { 'Суммы совпали.' }
        [pscustomobject]@{'Дата продажи'=$d.Date;'КУДиР'=$d.Kudir;'ОФД'=$d.Ofd;'Разница'=$d.Difference;'Возврат расхода ОФД'=$expenseReturn;'Комментарий'=$note;'Строки КУДиР'=($refs -join ', ');'Строки ОФД'=$d.ORows}
    })
    $cashRows=@(); $cashMonths=@(@($sales | Where-Object Kind -eq 'Наличные' | ForEach-Object { $_.Date.Substring(0,7) })+@($osSales | ForEach-Object { $_.Date.Substring(0,7) })) | Sort-Object -Unique
    foreach ($month in $cashMonths) {
        $kr=@($sales | Where-Object { $_.Kind -eq 'Наличные' -and $_.Date.StartsWith($month) }); $orr=@($osSales | Where-Object { $_.Date.StartsWith($month) })
        $kv=Sum-Field $kr Amount; $ov=[decimal]0; foreach ($r in $orr) { $ov+=$r.Cash*$r.Sign }
        $ranges=@($kr | Sort-Object From)
        for ($i=0; $i -lt $ranges.Count; $i++) {
            if ($ranges[$i].From -lt $start -or $ranges[$i].To -gt $end -or ($i -gt 0 -and $ranges[$i].From -le $ranges[$i-1].To)) {
                $issues.Add([pscustomobject]@{Источник=$ranges[$i].Ref;Причина='Наличный период пересекает границу отчёта или другой наличный период';Описание=$ranges[$i].Description})
            }
        }
        $cashRows += [pscustomobject]@{'Месяц'=$month;'КУДиР'=$kv;'ОФД'=$ov;'Разница'=($ov-$kv);'Периоды КУДиР'=(($kr | ForEach-Object { "$($_.From) — $($_.To)" }) -join '; ');'Строки КУДиР'=(($kr.Ref) -join ', ')}
    }
    $ofSales=[decimal]0; foreach ($r in $osSales) { $ofSales+=$r.Amount*$r.Sign }
    $kLedger=Sum-Field @($k | Where-Object { $_.DocDate -ge $start -and $_.DocDate -le $end }) Amount
    $other=@($k | Where-Object Kind -in @('Прочее поступление','Не распознано'))
    $shifted=@($sales | Where-Object { $_.DocDate -lt $start -or $_.DocDate -gt $end })
    foreach ($item in $k) {
        if ($item.Kind -eq 'Наличные' -and $item.From -le $end -and $item.To -ge $start -and ($item.From -lt $start -or $item.To -gt $end)) {
            $issues.Add([pscustomobject]@{Источник=$item.Ref;Причина='Наличная сумма охватывает даты за пределами отчёта ОФД; её нельзя делить пропорционально';Описание=$item.Description})
        }
    }
    foreach ($control in $controls) {
        $calc=$null
        if ($control.Source -eq 'ОФД') { $calc=Sum-Field $o Amount }
        elseif ($control.Label -match 'за год') { $calc=Sum-Field $k Amount }
        elseif ($control.Label -match 'за 9 месяцев') { $calc=Sum-Field @($k | Where-Object Sheet -in @('Квартал-1','Квартал-2','Квартал-3')) Amount }
        elseif ($control.Label -match 'за полугодие') { $calc=Sum-Field @($k | Where-Object Sheet -in @('Квартал-1','Квартал-2')) Amount }
        elseif ($control.Label -match 'квартал') { $sheetName=$control.Ref.Split('!')[0]; $calc=Sum-Field @($k | Where-Object Sheet -eq $sheetName) Amount }
        $delta=$null; if ($null -ne $calc -and $null -ne $control.Value) { $delta=$control.Value-$calc }
        $control | Add-Member -NotePropertyName Calculated -NotePropertyValue $calc
        $control | Add-Member -NotePropertyName Difference -NotePropertyValue $delta
        if ($null -ne $delta -and $delta -ne 0) { $issues.Add([pscustomobject]@{Источник=$control.Ref;Причина='Печатный итог отличается от суммы прочитанных операций';Описание="Разница: $(Fmt $delta)"}) }
    }
    return [pscustomobject]@{Start=$start;End=$end;K=$k.ToArray();O=$o.ToArray();KSales=(Sum-Field $sales Amount);OSales=$ofSales;KLedger=$kLedger;KAll=(Sum-Field $k Amount);OAll=(Sum-Field $o Amount);Difference=($ofSales-(Sum-Field $sales Amount));Bank=$bankRows;Cash=$cashRows;Other=$other;Special=@($o | Where-Object Sign -eq 0);Shifted=$shifted;Issues=$issues.ToArray();Audit=$audit.ToArray();Controls=$controls.ToArray();KPath=$kdata[0].Path;OPath=$odata.Path}
}

function Write-AutoReport($r,[string]$folder) {
    [void][IO.Directory]::CreateDirectory($folder)
    $bankDiff=@($r.Bank | Where-Object Разница -ne 0); $cashDiff=@($r.Cash | Where-Object Разница -ne 0)
    $b=[Text.StringBuilder]::new()
    [void]$b.Append('<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Сверка реальных файлов КУДиР и ОФД</title><style>body{font:16px/1.5 Segoe UI,Arial;color:#1b3040;max-width:1400px;margin:32px auto;padding:0 20px}h1{font-size:30px}h2{margin-top:30px}.box{background:#edf4fa;padding:20px}.warn{background:#fff0cf;padding:18px}table{border-collapse:collapse;width:100%;margin:14px 0;font-size:14px}th,td{border:1px solid #cad5df;padding:9px;text-align:left;vertical-align:top;overflow-wrap:anywhere}th{background:#edf3f8}a{color:#12639a}summary{cursor:pointer;font-weight:600}</style><h1>Сверка КУДиР и ОФД</h1>')
    [void]$b.Append("<p>Период ОФД: <b>$($r.Start) — $($r.End)</b>. Прочитано записей КУДиР: $($r.K.Count); чеков ОФД: $($r.O.Count).</p>")
    if ($r.Issues.Count) { [void]$b.Append("<p class='warn'><b>Сверка требует проверки: $($r.Issues.Count) замечаний.</b> Суммы ниже предварительные; см. раздел замечаний.</p>") }
    [void]$b.Append("<div class='box'><b>Сопоставимые продажи:</b><br>ОФД (приход минус возврат прихода): <b>$(Fmt $r.OSales) ₽</b><br>КУДиР (наличные, карты и QR, по датам продаж): <b>$(Fmt $r.KSales) ₽</b><br>ОФД минус КУДиР: <b>$(Fmt $r.Difference) ₽</b><br>Безналичные: $($bankDiff.Count) дат с расхождениями. Наличные: $($cashDiff.Count) месяцев с расхождениями.</div>")
    [void]$b.Append('<p>Положительная разница: в выбранных записях КУДиР меньше. Отрицательная: в КУДиР больше. Это список для проверки, а не доказательство отсутствия конкретного чека.</p><p><a href="#bank-overview">Проверить безналичные</a> · <a href="#cash-overview">Проверить наличные</a></p><details><summary>Почему общие суммы файлов отличаются от сопоставимых продаж</summary>')
    $summary=@(
        [pscustomobject]@{'Показатель'='ОФД: все денежные чеки, включая расходные операции';'Сумма'=$r.OAll},
        [pscustomobject]@{'Показатель'='КУДиР: все кварталы файла по дате документа';'Сумма'=$r.KAll},
        [pscustomobject]@{'Показатель'='Разница всех денежных чеков ОФД и всех кварталов КУДиР';'Сумма'=($r.OAll-$r.KAll)},
        [pscustomobject]@{'Показатель'='КУДиР: доходы с датой документа внутри периода ОФД';'Сумма'=$r.KLedger},
        [pscustomobject]@{'Показатель'='ОФД: расход / возврат расхода, отдельно от продаж';'Сумма'=(Sum-Field $r.Special Amount)},
        [pscustomobject]@{'Показатель'='КУДиР: прочие / нераспознанные поступления, отдельно';'Сумма'=(Sum-Field $r.Other Amount)},
        [pscustomobject]@{'Показатель'='КУДиР: продажи периода, зачисленные за его границами';'Сумма'=(Sum-Field $r.Shifted Amount)}
    )
    [void]$b.Append((Auto-Table $summary @('Показатель','Сумма')))
    [void]$b.Append('<p>Общие суммы файлов имеют разный состав и могут охватывать разные даты зачисления. Для сравнения продаж программа объединяет кварталы, берёт дату после «за» в описании зачисления по картам, дату после «от» у QR и сравнивает наличные по месяцам. Оригинальная дата документа сохраняется. Печатные итоги ниже показаны отдельно; они не складываются с операциями.</p></details>')
    [void]$b.Append('<style>.table-scroll{overflow-x:auto;max-width:100%;margin:14px 0}.table-scroll table{margin:0}th,td{overflow-wrap:normal;word-break:normal}th{min-width:90px}td{overflow-wrap:break-word}.nowrap,.money{white-space:nowrap!important;overflow-wrap:normal!important;word-break:normal!important}.money{text-align:right;font-variant-numeric:tabular-nums;min-width:105px}table.bank{min-width:850px}.bank th:nth-child(-n+4){min-width:110px;white-space:nowrap}.bank td:nth-child(5){min-width:200px}.case{border:1px solid #c8d6e1;border-radius:10px;padding:22px;margin:24px 0;scroll-margin-top:18px}.case h3{font-size:22px;margin-top:0}.action{background:#eaf3fc;padding:16px;border-left:4px solid #246eaa}.action p{margin-bottom:0}.conditional{background:#fff4dc;padding:16px}details{margin:18px 0}summary{padding:10px;background:#f0f4f7}h4{font-size:17px}@media print{.table-scroll{overflow:visible}table.bank{min-width:0}.case{break-before:page}body{max-width:none;padding:0;font-size:12px}th,td{padding:5px}}</style><h2 id="bank-overview">Безналичные: даты для проверки</h2><p>Начните с «Разобрать разницу»: там указаны обе даты, точные ячейки и порядок проверки. На узком экране таблицу можно прокрутить вправо.</p><div class="table-scroll" tabindex="0" role="region" aria-label="Даты для проверки"><table class="bank"><thead><tr><th>Дата продажи</th><th>КУДиР, ₽</th><th>ОФД, ₽</th><th>Разница, ₽</th><th>Что это значит</th><th>Что проверить</th></tr></thead><tbody>')
    foreach ($d in $bankDiff) {
        $direction=if ($d.Разница -lt 0) { 'В КУДиР больше' } else { 'В КУДиР меньше' }
        [void]$b.Append("<tr><td class='nowrap'>$(Display-Date $d.'Дата продажи')</td><td class='money'>$(Fmt $d.КУДиР)</td><td class='money'>$(Fmt $d.ОФД)</td><td class='money'>$(Fmt $d.Разница)</td><td>$direction на <span class='nowrap'>$(Fmt ([math]::Abs($d.Разница))) ₽</span></td><td><a href='#day-$($d.'Дата продажи')'>Разобрать разницу →</a></td></tr>")
    }
    [void]$b.Append('</tbody></table></div>')
    [void]$b.Append((Write-BankDetails $r $bankDiff))
    [void]$b.Append('<h2 id="cash-overview">Наличные: сверка по месяцам</h2><p>Наличные в КУДиР записаны за период общей суммой. Поэтому день возможного пропуска внутри месяца установить по этому файлу нельзя.</p>')
    [void]$b.Append((Auto-Table $r.Cash @('Месяц','КУДиР','ОФД','Разница','Периоды КУДиР','Строки КУДиР')))
    [void]$b.Append('<h2>Зачисления за границей периода ОФД, учтённые по дате продажи</h2>')
    $shift=@($r.Shifted | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Дата документа'=$_.DocDate;'Дата продажи'=$_.Date;'Сумма'=$_.Amount;'Описание'=$_.Description} })
    [void]$b.Append((Auto-Table $shift @('Строка','Дата документа','Дата продажи','Сумма','Описание')))
    [void]$b.Append('<h2>Прочие поступления КУДиР</h2><p>Эти записи исключены из сопоставимых продаж и включены в общую сумму КУДиР. Проверьте их состав.</p>')
    $other=@($r.Other | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Дата'=$_.DocDate;'Сумма'=$_.Amount;'Описание'=$_.Description} })
    [void]$b.Append((Auto-Table $other @('Строка','Дата','Сумма','Описание')))
    [void]$b.Append("<h2>Расходные операции ОФД — отдельно</h2><p>Найдено $($r.Special.Count) чеков, сумма $(Fmt (Sum-Field $r.Special Amount)) ₽. Типы взяты буквально из столбца Q: «Расход» и «Возврат расхода» не переименовываются в возврат продажи. Они не включены в сумму продаж. Если это ошибочно оформленные чеки, причину расхождений нужно разбирать с их учётом.</p>")
    $special=@($r.Special | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Дата'=$_.Date;'Операция'=$_.Kind;'Сумма'=$_.Amount;'Наличные'=$_.Cash;'Безналичные'=$_.Bank;'ФН / ФД'=$_.Id} })
    [void]$b.Append((Auto-Table $special @('Строка','Дата','Операция','Сумма','Наличные','Безналичные','ФН / ФД')))
    [void]$b.Append('<h2>Печатные итоги исходных файлов</h2>')
    $controls=@($r.Controls | ForEach-Object { [pscustomobject]@{'Источник'=$_.Source;'Строка'=$_.Ref;'Название'=$_.Label;'Печатный итог'=$(if ($null -eq $_.Value) { 'Нет сохранённого значения' } else { $_.Value });'По операциям'=$_.Calculated;'Разница'=$_.Difference} })
    [void]$b.Append((Auto-Table $controls @('Источник','Строка','Название','Печатный итог','По операциям','Разница')))
    if ($r.Issues.Count) { [void]$b.Append('<h2>Замечания к исходным данным</h2>'); [void]$b.Append((Auto-Table $r.Issues @('Источник','Причина','Описание'))) }
    [void]$b.Append('<h2>Данные для проверки в Excel</h2><p><a href="bank.csv">Все дни безналичных</a> · <a href="cash.csv">Наличные</a> · <a href="kudir.csv">Все записи КУДиР</a> · <a href="ofd.csv">Все чеки ОФД</a> · <a href="audit.csv">Исключённые служебные строки и замечания</a></p>')
    [void]$b.Append("<p>КУДиР: $(Html $r.KPath)<br>ОФД: $(Html $r.OPath)</p><p>Файлы прочитаны без изменения. Фильтры и скрытые строки не исключают данные. Копейки не игнорируются. Совпадение сумм не доказывает совпадение всех чеков. Программа не меняет бухгалтерские даты и записи.</p></html>")
    [IO.File]::WriteAllText((Join-Path $folder 'Отчёт.html'),$b.ToString(),[Text.UTF8Encoding]::new($true))
    Write-CsvSafe (Join-Path $folder 'bank.csv') $r.Bank @('Дата продажи','КУДиР','ОФД','Разница','Возврат расхода ОФД','Комментарий','Строки КУДиР','Строки ОФД')
    Write-CsvSafe (Join-Path $folder 'cash.csv') $r.Cash @('Месяц','КУДиР','ОФД','Разница','Периоды КУДиР','Строки КУДиР')
    $kcsv=@($r.K | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Дата документа'=$_.DocDate;'Дата для сверки'=$_.Date;'Вид'=$_.Kind;'Сумма'=$_.Amount;'Описание'=$_.Description} })
    $ocsv=@($r.O | ForEach-Object { [pscustomobject]@{'Строка'=$_.Ref;'Дата'=$_.Date;'Операция'=$_.Kind;'Сумма'=$_.Amount;'Наличные'=$_.Cash;'Безналичные'=$_.Bank;'ФН / ФД'=$_.Id;'Включён в продажи'=($_.Sign -ne 0)} })
    Write-CsvSafe (Join-Path $folder 'kudir.csv') $kcsv @('Строка','Дата документа','Дата для сверки','Вид','Сумма','Описание')
    Write-CsvSafe (Join-Path $folder 'ofd.csv') $ocsv @('Строка','Дата','Операция','Сумма','Наличные','Безналичные','ФН / ФД','Включён в продажи')
    Write-CsvSafe (Join-Path $folder 'audit.csv') (@($r.Audit)+@($r.Issues)) @('Источник','Причина','Описание')
    return (Join-Path $folder 'Отчёт.html')
}
