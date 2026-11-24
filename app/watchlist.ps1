# Dot-sourced by update.ps1. Does not alter strategy signals.
function Get-WatchHtml {
 $groups=ConvertFrom-Json -InputObject (Get-Content (Join-Path $PSScriptRoot 'watchlist.json') -Raw -Encoding UTF8)
 if($groups.Count -ne 6){throw '觀察清單應包含六個群組'}
 $codes=@($groups | ForEach-Object {$_.symbols} | Sort-Object -Unique)
 $cacheDir=Join-Path $PSScriptRoot 'data/watchlist';$null=New-Item -ItemType Directory -Force $cacheDir
 $tz=[TimeZoneInfo]::FindSystemTimeZoneById('Taipei Standard Time')
 $now=[TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow,$tz)
 $calendar=@();if(Test-Path 'data/TWII.csv'){$calendar=@(Import-Csv 'data/TWII.csv' | Select-Object -ExpandProperty Date)}
 $pool=[runspacefactory]::CreateRunspacePool(1,6);$pool.Open();$jobs=@()
 $worker={param($code,$cacheDir)
  $ErrorActionPreference='Stop';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
  $cache=Join-Path $cacheDir ($code+'.json');$old=$null
  if(Test-Path $cache){try{$old=Get-Content $cache -Raw -Encoding UTF8 | ConvertFrom-Json}catch{}}
  $merged=@{};if($old){foreach($row in $old.Rows){if($row.Date -and $row.Close -gt 0){$merged[[string]$row.Date]=$row}}}
  $result=$null;$errorText='無有效行情（請檢查代號或資料源）'
  $tz=[TimeZoneInfo]::FindSystemTimeZoneById('Taipei Standard Time');$now=[TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow,$tz)
  if($code -notmatch '^\d{4,6}[A-Z]?$'){$errorText='代號格式異常'}else{
   $suffixes=@('.TW','.TWO');if($old -and $old.Symbol -like '*.TWO'){$suffixes=@('.TWO','.TW')}
   foreach($suffix in $suffixes){
    $symbol=$code+$suffix;$enc=[uri]::EscapeDataString($symbol)
    foreach($server in @('query1.finance.yahoo.com','query2.finance.yahoo.com')){
     try{
      $res=Invoke-RestMethod -Uri "https://$server/v8/finance/chart/${enc}?range=6mo&interval=1d" -UserAgent 'Mozilla/5.0' -TimeoutSec 9
      $r=$res.chart.result[0];if(-not $r.timestamp.Count){continue};$q=$r.indicators.quote[0];$rows=@()
      for($i=0;$i -lt $r.timestamp.Count;$i++){
       $d=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$r.timestamp[$i]).UtcDateTime,$tz).Date
       if($d -gt $now.Date -or ($d -eq $now.Date -and $now.TimeOfDay -lt [timespan]'13:35:00') -or $q.close[$i] -le 0){continue}
       $rows+=[pscustomobject]@{Date=$d.ToString('yyyy-MM-dd');Close=[double]$q.close[$i]}
      }
      if($rows.Count -eq 0){continue}
      foreach($row in $rows){$merged[[string]$row.Date]=$row}
      $livePrice=$null;$liveTime='';$liveDate=''
      if($r.meta.regularMarketPrice -gt 0 -and $r.meta.regularMarketTime -gt 0){
       $livePrice=[double]$r.meta.regularMarketPrice
       $lt=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$r.meta.regularMarketTime).UtcDateTime,$tz)
       $liveTime=$lt.ToString('yyyy-MM-dd HH:mm:ss');$liveDate=$lt.ToString('yyyy-MM-dd')
      }
      $result=[pscustomobject]@{Code=$code;Symbol=$symbol;Name=$r.meta.shortName;LivePrice=$livePrice;LiveTime=$liveTime;LiveDate=$liveDate;Rows=@($merged.Values | Sort-Object Date);Stale=$false;Error='';Retrieved=$now.ToString('yyyy-MM-dd HH:mm:ss')}
      $result | ConvertTo-Json -Depth 6 | Set-Content "$cache.tmp" -Encoding UTF8;Move-Item "$cache.tmp" $cache -Force;break
     }catch{}
    };if($result){break}
   }
  }
  if(-not $result){if($old){$old.Stale=$true;$old.Error='更新失敗，沿用快取';$result=$old}else{$result=[pscustomobject]@{Code=$code;Symbol='';Name='';Rows=@();Stale=$true;Error=$errorText;Retrieved=''}}}
  return $result
 }
 try{
  foreach($code in $codes){$ps=[powershell]::Create();$ps.RunspacePool=$pool;$null=$ps.AddScript($worker.ToString()).AddArgument($code).AddArgument($cacheDir);$jobs+=@{Code=$code;PS=$ps;Handle=$ps.BeginInvoke()}}
  $results=@{};foreach($job in $jobs){try{$results[$job.Code]=@($job.PS.EndInvoke($job.Handle))[-1]}catch{$results[$job.Code]=[pscustomobject]@{Code=$job.Code;Rows=@();Stale=$true;Error='更新失敗'}}finally{$job.PS.Dispose()}}
 }finally{$pool.Close();$pool.Dispose()}
 $allDates=@(@($calendar)+@($results.Values | ForEach-Object {$_.Rows | ForEach-Object {$_.Date}}) | Where-Object {$_} | Sort-Object -Unique | Select-Object -Last 60)
 $dates=@($allDates | Select-Object -Last 11)
 foreach($code in $codes){
  $r=$results[$code];if(-not $r){continue}
  $map=@{};foreach($row in $r.Rows){if($row.Date -and $row.Close -gt 0){$map[[string]$row.Date]=$row}}
  $missing=@($dates | Where-Object {-not $map.ContainsKey($_)})
  if($missing.Count -eq 0){continue}
  $repaired=0
  foreach($month in @($missing | ForEach-Object {$_.Substring(0,7)} | Sort-Object -Unique)){
   $ym=$month.Replace('-','')+'01'
   try{
    $official=Invoke-RestMethod -Uri "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=$ym&stockNo=$code&response=json" -UserAgent 'Mozilla/5.0' -TimeoutSec 8
    if($official.stat -ne 'OK'){continue}
    foreach($record in $official.data){
     $parts=([string]$record[0]).Split('/');if($parts.Count -ne 3){continue}
     $d=('{0:D4}-{1:D2}-{2:D2}' -f ([int]$parts[0]+1911),[int]$parts[1],[int]$parts[2])
     if($d -notin $missing -or $map.ContainsKey($d)){continue}
     if($d -eq $now.ToString('yyyy-MM-dd') -and $now.TimeOfDay -lt [timespan]'13:35:00'){continue}
     $close=0.0;if(-not [double]::TryParse(([string]$record[6]).Replace(',',''),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$close) -or $close -le 0){continue}
     $map[$d]=[pscustomobject]@{Date=$d;Close=$close};$repaired++
    }
   }catch{}
   Start-Sleep -Milliseconds 350
  }
  if($repaired -gt 0){
   $r.Rows=@($map.Values | Sort-Object Date)
   $r | Add-Member -NotePropertyName Repaired -NotePropertyValue $repaired -Force
   $cache=Join-Path $cacheDir ($code+'.json')
   $r | ConvertTo-Json -Depth 6 | Set-Content "$cache.tmp" -Encoding UTF8;Move-Item "$cache.tmp" $cache -Force
  }
 }
 $summary=Get-GroupPerformance $groups $results $allDates
 $visible=@($dates | Select-Object -Last 10);$visible=@($visible | Sort-Object)
 $html=$summary+'<div class="watchhead"><h2>觀察清單</h2><p>預設依績效評分降冪排序（3 日累積＋10 日累積；非實際期間報酬）；點擊日期、即時或任一累積欄可獨立排序。即時欄為最近報價及相對前一交易日收盤漲跌幅（可能延遲），休市沿用最近報價；3 日及 10 日累積僅使用已完成交易日收盤價，不包含即時報價；缺值排最後。</p></div><div class="watchgroups">'
 foreach($group in $groups){
  $html+="<section class='watchgroup'><h3>$(E $group.section)－$(E $group.name)</h3><div class='watchscroll'><table class='watchtable'><thead><tr><th>代號</th>"
  foreach($d in $visible){$html+="<th><button type='button' class='watchsort' data-key='d$d'>$(E $d)</button></th>"};$html+='<th><button type="button" class="watchsort" data-key="three">3 日累積</button></th><th><button type="button" class="watchsort" data-key="ten">10 日累積</button></th><th><button type="button" class="watchsort" data-key="score">績效評分</button></th><th>狀態</th><th><button type="button" class="watchsort" data-key="live">即時</button></th></tr></thead><tbody>'
  foreach($code in $group.symbols){
   $r=$results[$code];$map=@{};foreach($row in $r.Rows){if($row.Date){$map[$row.Date]=[double]$row.Close}}
   $daily=$null;$three=$null;$ten=$null;$cells='';foreach($d in $visible){
    $idx=[array]::IndexOf($dates,$d);$pct=$null;$close=$null;if($map.ContainsKey($d)){$close=$map[$d];if($idx -gt 0 -and $map.ContainsKey($dates[($idx-1)]) -and $map[$dates[($idx-1)]] -gt 0){$pct=100*($close/$map[$dates[($idx-1)]]-1)}}
    $sortValue='';if($null -ne $pct){$sortValue=$pct.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
    if($d -eq $visible[-1]){$daily=$pct};$cells+="<td data-key='d$d' data-value='$sortValue'>$(F $close)<br>$(Format-WatchPct $pct)</td>"
   }
   $lastFour=@($dates | Select-Object -Last 4)
   if($lastFour.Count -eq 4 -and @($lastFour | Where-Object {$map.ContainsKey($_)}).Count -eq 4 -and $map[$lastFour[0]] -gt 0){$three=100*($map[$lastFour[-1]]/$map[$lastFour[0]]-1)}
   if($dates.Count -eq 11 -and @($dates | Where-Object {$map.ContainsKey($_)}).Count -eq 11 -and $map[$dates[0]] -gt 0){$ten=100*($map[$dates[-1]]/$map[$dates[0]]-1)}
   $ts='';if($null -ne $ten){$ts=$ten.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
   $ds='';$fs='';if($null -ne $daily){$ds=$daily.ToString('R',[Globalization.CultureInfo]::InvariantCulture)};if($null -ne $three){$fs=$three.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
   $score=$null;$ss='';if($null -ne $three -and $null -ne $ten){$score=$three+$ten;$ss=$score.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
   $status='正常';if($r.Stale){$status=$r.Error}elseif($visible.Count -eq 0 -or -not $map.ContainsKey($visible[-1])){$status='最新交易日缺值'}elseif($null -eq $ten){$status='缺收盤：'+(@($dates | Where-Object {-not $map.ContainsKey($_)}) -join '、');if($dates.Count -lt 11){$status='不足十一個交易日'}}
   if($status -eq '正常' -and $r.Repaired -gt 0){$status='已回補 '+$r.Repaired+' 日'}
   $livePct=$null;$liveSort='';$livePrice=$r.LivePrice;$liveTime=$r.LiveTime
   if($livePrice -gt 0 -and $r.LiveDate){
    $prior=@($dates | Where-Object {$_ -lt $r.LiveDate} | Select-Object -Last 1)
    if($prior.Count -eq 1 -and $map.ContainsKey($prior[0]) -and $map[$prior[0]] -gt 0){$livePct=100*($livePrice/$map[$prior[0]]-1);$liveSort=$livePct.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
   }
   $liveCell="<td data-key='live' data-value='$liveSort' title='$(E $liveTime)（台北時間）'>$(F $livePrice)<br>$(Format-WatchPct $livePct)<br><small>$(E $liveTime)</small></td>"
   $name=E $r.Name;$html+="<tr data-daily='$ds' data-three='$fs'><td><strong>$(E $code)</strong><br><small>$name</small></td>$cells<td data-key='three' data-value='$fs'>$(Format-WatchPct $three)</td><td data-key='ten' data-value='$ts'>$(Format-WatchPct $ten)</td><td data-key='score' data-value='$ss'>$(Format-WatchPct $score)</td><td>$(E $status)</td>$liveCell</tr>"
  };$html+='</tbody></table></div></section>'
 };$html+='</div>';return $html
}
function Format-WatchPct($value){if($null -eq $value){return '—'};$cls='flat';if($value -gt 0){$cls='up'}elseif($value -lt 0){$cls='down'};return ('<span class="{0}">{1:+0.00;-0.00;0.00}%</span>' -f $cls,$value)}
function Get-GroupPerformance($groups,$results,$days){
 $stats=@()
 foreach($group in $groups){
  $maps=@{};foreach($code in $group.symbols){$map=@{};foreach($row in $results[$code].Rows){if($row.Date -and $row.Close -gt 0){$map[[string]$row.Date]=[double]$row.Close}};$maps[$code]=$map}
  $series=@();for($i=10;$i -lt $days.Count;$i++){
   $values=@{};foreach($code in $group.symbols){$m=$maps[$code];$valid=$true;for($j=$i-10;$j -le $i;$j++){if(-not $m.ContainsKey($days[$j])){$valid=$false;break}};if($valid){$values[$code]=100*($m[$days[$i]]/$m[$days[$i-10]]-1)}}
   $series+= [pscustomobject]@{Date=$days[$i];Values=$values}
  }
  $value=$null;$prior=$null;$delta=$null;$count=0;$pairCount=0;$direction=0;$streak=0
  if($series.Count -gt 0){$latest=$series[-1].Values;$count=$latest.Count;if($count -gt 0){$value=($latest.Values | Measure-Object -Average).Average}}
  for($i=$series.Count-1;$i -ge 1;$i--){
   $a=$series[$i].Values;$b=$series[$i-1].Values;$common=@($a.Keys | Where-Object {$b.ContainsKey($_)});if($common.Count -eq 0){break}
   $av=($common | ForEach-Object {$a[$_]} | Measure-Object -Average).Average
   $bv=($common | ForEach-Object {$b[$_]} | Measure-Object -Average).Average
   $diff=[math]::Round(($av-$bv),8);$dir=[math]::Sign($diff)
   if($i -eq $series.Count-1){$prior=$bv;$delta=$diff;$pairCount=$common.Count;$direction=$dir}
   if($dir -eq 0 -or $dir -ne $direction){break};$streak++
  }
  $stats+= [pscustomobject]@{Name=($group.section+'－'+$group.name);Value=$value;Prior=$prior;Delta=$delta;Count=$count;Total=$group.symbols.Count;PairCount=$pairCount;Direction=$direction;Streak=$streak}
 }
 $max=1.0;foreach($x in $stats){if($null -ne $x.Value){$max=[math]::Max($max,[math]::Abs($x.Value))}}
 $dateText='無資料';if($days.Count){$dateText=$days[-1]}
 $h="<section class='groupoverview'><h2>群組績效總覽</h2><p class='groupnote'>截至 $(E $dateText) ｜ 10 日累積平均；紅色＝較前一交易日增強，綠色＝放緩，灰色＝持平或無法比較。即時行情不計入。</p>"
 foreach($x in @($stats | Sort-Object -Property @{Expression={if($null -eq $_.Value){[double]::NegativeInfinity}else{[double]$_.Value}};Descending=$true}, @{Expression={$_.Name};Descending=$false})){
  $color='#78828f';if($x.Direction -gt 0){$color='#b42332'}elseif($x.Direction -lt 0){$color='#167348'}
  $trend='';if($x.Streak -ge 2){$word='增強';if($x.Direction -lt 0){$word='放緩'};$trend="<strong style='color:$color'>連續$word$($x.Streak)日</strong>"}
  $width=0;$left=50;if($null -ne $x.Value){$width=48*[math]::Abs($x.Value)/$max;if($x.Value -lt 0){$left=50-$width}}
  $w=$width.ToString('0.###',[Globalization.CultureInfo]::InvariantCulture);$l=$left.ToString('0.###',[Globalization.CultureInfo]::InvariantCulture)
  $v='—';if($null -ne $x.Value){$v=('{0:+0.00;-0.00;0.00}%' -f $x.Value)}
  $detail='無法比較前一交易日';if($null -ne $x.Delta){$detail=('前一日（相同標的） {0:+0.00;-0.00;0.00}% ｜ 變化 {1:+0.00;-0.00;0.00} 個百分點 ｜ 比較 {2} 檔' -f $x.Prior,$x.Delta,$x.PairCount)}
  $h+="<div class='gitem'><div class='gtitle'><b>$(E $x.Name)</b><span>$trend</span></div><div class='gline'><div class='gtrack'><i style='left:${l}%;width:${w}%;background:$color'></i></div><b style='color:$color'>$v</b></div><small>有效 $($x.Count)/$($x.Total) 檔 ｜ $detail</small></div>"
 }
 return $h+'</section>'
}
