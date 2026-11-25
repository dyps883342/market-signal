param([switch]$Silent)
$ErrorActionPreference='Stop'
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
Set-Location -LiteralPath $PSScriptRoot
$cfg=Get-Content strategy.json -Raw | ConvertFrom-Json
$today=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow,'Taipei Standard Time').Date
$null=New-Item -ItemType Directory -Force data
$mutex=New-Object System.Threading.Mutex($false,('MarketSignal_'+[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PSScriptRoot)).Replace('/','_').Replace('+','-')))
if(-not $mutex.WaitOne(0)){if(-not $Silent){Write-Host 'Update already running.'};exit 0}
try {
function Get-Rows($symbol,$zone,$file){
 $old=@{}; if(Test-Path $file){Import-Csv $file | ForEach-Object {$old[$_.Date]=$_}}
 $start=$today.AddMonths(-[int]$cfg.initial_history_months)
 if($old.Count){$start=([datetime]($old.Keys | Sort-Object | Select-Object -Last 1)).AddDays(-7)}
 $epoch=[datetime]'1970-01-01'; $p1=[long]($start-$epoch).TotalSeconds; $p2=[long]($today.AddDays(1)-$epoch).TotalSeconds
 $encoded=[uri]::EscapeDataString($symbol); $r=$null
 foreach($hostName in @('query1.finance.yahoo.com','query2.finance.yahoo.com')){try{$res=Invoke-RestMethod -Uri "https://$hostName/v8/finance/chart/${encoded}?period1=$p1&period2=$p2&interval=1d" -UserAgent 'Mozilla/5.0' -TimeoutSec 25; $r=$res.chart.result[0]; if($r.timestamp.Count -gt 0){break}}catch{$r=$null}}
 if($null -eq $r -or $r.timestamp.Count -eq 0){throw 'Yahoo Finance 未回傳有效行情，請稍後再試。'}
 $q=$r.indicators.quote[0]; $added=0
 # Quote display is separate from the closed-day strategy series.
 $tz=[TimeZoneInfo]::FindSystemTimeZoneById($zone)
 $nowUtc=[DateTime]::UtcNow
 $localNow=[TimeZoneInfo]::ConvertTimeFromUtc($nowUtc,$tz)
 $lastOpen=$null
 for($k=0;$k -lt $r.timestamp.Count;$k++){
  if($q.open[$k] -gt 0){
   $candidate=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$r.timestamp[$k]).UtcDateTime,$tz)
   if($candidate -le $localNow){$lastOpen=$candidate}
  }
 }
 $meta=$r.meta; $quoteTime=$null; $quotePrice=$null; $isOpen=$false
 if($meta.regularMarketTime -gt 0 -and $meta.regularMarketPrice -gt 0){
  $quoteTime=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$meta.regularMarketTime).UtcDateTime,$tz)
  $quotePrice=[double]$meta.regularMarketPrice
 }
 $session=$meta.currentTradingPeriod.regular
 if($session.start -gt 0 -and $session.end -gt $session.start){
  $sessionStart=[DateTimeOffset]::FromUnixTimeSeconds([long]$session.start).UtcDateTime
  $sessionEnd=[DateTimeOffset]::FromUnixTimeSeconds([long]$session.end).UtcDateTime
  $isOpen=($nowUtc -ge $sessionStart -and $nowUtc -lt $sessionEnd -and $localNow.DayOfWeek -notin @('Saturday','Sunday'))
 }
 $script:QuoteInfo=[pscustomobject]@{Price=$quotePrice;Time=$quoteTime;LastOpen=$lastOpen;IsOpen=$isOpen;LocalDate=$localNow.ToString('yyyy-MM-dd');Retrieved=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($nowUtc,'Taipei Standard Time')}

 for($i=0;$i -lt $r.timestamp.Count;$i++){
  $utc=[DateTimeOffset]::FromUnixTimeSeconds([long]$r.timestamp[$i]).UtcDateTime
  $d=[TimeZoneInfo]::ConvertTimeFromUtc($utc,[TimeZoneInfo]::FindSystemTimeZoneById($zone)).ToString('yyyy-MM-dd')
  if($q.close[$i] -le 0 -or $q.open[$i] -le 0){continue}
  $localDate=$localNow.ToString('yyyy-MM-dd')
  if($d -gt $localDate){continue}
  if($d -eq $localDate){
   $closeAt=[timespan]'16:00:00';if($zone -eq 'Taipei Standard Time'){$closeAt=[timespan]'13:30:00'}
   $closed=($localNow.TimeOfDay -ge $closeAt.Add([timespan]'00:05:00'))
   if($session.end -gt 0 -and $session.start -gt 0){
    $sd=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$session.start).UtcDateTime,$tz).ToString('yyyy-MM-dd')
    if($sd -eq $d){$closed=($nowUtc -ge [DateTimeOffset]::FromUnixTimeSeconds([long]$session.end).UtcDateTime.AddMinutes(5))}
   }
   if(-not $closed){continue}
  }
  $old[$d]=[pscustomobject]@{Date=$d;Open=[double]$q.open[$i];High=[double]$q.high[$i];Low=[double]$q.low[$i];Close=[double]$q.close[$i]};$added++
 }
 # The daily candle retains intraday extremes even if price has since reversed.
 $touch=$null
 for($i=0;$i -lt $r.timestamp.Count;$i++){
  if($q.high[$i] -le 0 -or $q.low[$i] -le 0){continue}
  $date=[TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds([long]$r.timestamp[$i]).UtcDateTime,$tz).ToString('yyyy-MM-dd')
  if($date -gt $localNow.ToString('yyyy-MM-dd')){continue}
  $touch=[pscustomobject]@{Date=$date;High=[double]$q.high[$i];Low=[double]$q.low[$i]}
 }
 if($touch){
  $touchFile="$file.touch.json"
  if(Test-Path $touchFile){try{$saved=Get-Content $touchFile -Raw -Encoding UTF8 | ConvertFrom-Json;if($saved.Date -eq $touch.Date){$touch.High=[math]::Max($touch.High,[double]$saved.High);$touch.Low=[math]::Min($touch.Low,[double]$saved.Low)}}catch{}}
  $touch | ConvertTo-Json -Compress | Set-Content "$touchFile.tmp" -Encoding UTF8
  Move-Item "$touchFile.tmp" $touchFile -Force
 }
 $script:QuoteInfo | Add-Member -NotePropertyName Touch -NotePropertyValue $touch -Force
 if($old.Count -eq 0){throw '沒有有效的已收盤行情。'}
 $rows=@($old.Values | Sort-Object Date);$rows | Export-Csv "$file.tmp" -NoTypeInformation -Encoding UTF8;Move-Item "$file.tmp" $file -Force
 return $rows
}
function Get-Signals($rows){
 $high=[double]$rows[0].Close;$low=$high;$phase='追蹤高點';$bought=@{};$sold=0;$reset=[double]$cfg.cycle_reset_threshold/100;$step=[double]$cfg.sell_interval/100;$ratio=[double]$cfg.sell_ratio/100
 foreach($row in $rows){
  $price=[double]$row.Close;$action='觀察';$buy=0
  if($phase -eq '反彈' -and $price -le $high*(1-$reset)+1e-8){$phase='修正';$low=$price;$bought=@{};$sold=0}
  if($phase -eq '追蹤高點'){$high=[math]::Max($high,$price);if($price -le $high*(1-$reset)+1e-8){$phase='修正';$low=$price}}
  if($phase -eq '修正'){
   $low=[math]::Min($low,$price)
   if($price -ge $low*(1+$step)-1e-8){$phase='反彈';$high=$price}
   else{for($j=0;$j -lt $cfg.buy_thresholds.Count;$j++){if(-not $bought.ContainsKey($j) -and $price -le $high*(1-$cfg.buy_thresholds[$j]/100)+1e-8){$bought[$j]=$true;$buy+=$cfg.buy_weights[$j]}};if($buy -gt 0){$action="加碼 $buy%"}}
  }
  if($phase -eq '反彈'){$high=[math]::Max($high,$price);$n=[int][math]::Floor(($price/$low-1+1e-10)/$step);if($n -gt $sold){$pct=100*(1-[math]::Pow(1-$ratio,$n-$sold));$action=('出場 {0:0.##}%' -f $pct);$sold=$n}}
  $nextWeight=$null;$nextBuy=$null;if($phase -eq '修正'){for($j=0;$j -lt $cfg.buy_thresholds.Count;$j++){if(-not $bought.ContainsKey($j)){$nextBuy=$high*(1-$cfg.buy_thresholds[$j]/100);$nextWeight=$cfg.buy_weights[$j];break}}}else{$nextBuy=$high*(1-$reset);$nextWeight=$cfg.buy_weights[0]}
  $nextSell=$null;$shownLow=$null;if($phase -ne '追蹤高點'){$nextSell=$low*(1+($sold+1)*$step);$shownLow=$low}
  [pscustomobject]@{Date=$row.Date;Open=[double]$row.Open;Close=$price;Action=$action;Phase=$phase;High=$high;Low=$shownLow;NextBuy=$nextBuy;NextBuyWeight=$nextWeight;NextSell=$nextSell;SoldSteps=$sold;BoughtSteps=@($bought.Keys)}
 }
}
# Preview uses prior confirmed close state; never writes a simulated trade into the formal history.
function Get-TouchPreview($signals,$touch){
 if(-not $touch){return '盤中高低點未提供'}
 $prior=@($signals | Where-Object {$_.Date -lt $touch.Date})
 if($prior.Count -eq 0){return '歷史基準不足'}
 $p=$prior[-1];$hi=[double]$touch.High;$lo=[double]$touch.Low
 $baseHigh=[double]$p.High;$baseLow=$p.Low
 $buy=0;$sell=0;$step=[double]$cfg.sell_interval/100
 # Only levels anchored to confirmed prior state are evaluated; daily OHLC cannot establish ordering.
 if($p.Phase -eq '修正'){
  for($j=0;$j -lt $cfg.buy_thresholds.Count;$j++){if($j -notin $p.BoughtSteps -and $lo -le $baseHigh*(1-$cfg.buy_thresholds[$j]/100)+1e-8){$buy+=$cfg.buy_weights[$j]}}
 }else{
  $baseHigh=[math]::Max($baseHigh,$hi)
  for($j=0;$j -lt $cfg.buy_thresholds.Count;$j++){if($lo -le $baseHigh*(1-$cfg.buy_thresholds[$j]/100)+1e-8){$buy+=$cfg.buy_weights[$j]}}
 }
 if($p.Phase -ne '追蹤高點' -and $baseLow -gt 0){
  if($p.Phase -eq '修正'){$baseLow=[math]::Min([double]$baseLow,$lo)}
  $n=[int][math]::Floor(($hi/$baseLow-1+1e-10)/$step)
  if($n -gt $p.SoldSteps){$sell=100*(1-[math]::Pow(1-[double]$cfg.sell_ratio/100,$n-$p.SoldSteps))}
 }
 $msg=@()
 if($buy -gt 0){$msg+="加碼 $buy%"}
 if($sell -gt 0){$msg+=('出場 {0:0.##}%' -f $sell)}
 if($msg.Count -eq 0){$msg+='觀察'}
 $note='';if($buy -gt 0 -or ($p.Phase -eq '修正' -and $lo -lt $p.Low -and $sell -gt 0)){$note='（日內高低先後未確認，僅觸及提醒）'}
 return (($msg -join ' ／ ')+$note)
}
function F($v){if($null -eq $v){return '—'};return ('{0:N2}' -f [double]$v)}
function E($v){return [System.Net.WebUtility]::HtmlEncode([string]$v)}
$cards=@();foreach($m in @(@{Name='台股加權指數';Symbol='^TWII';Zone='Taipei Standard Time';Id='TWII'},@{Name='S&P 500';Symbol='^GSPC';Zone='Eastern Standard Time';Id='GSPC'})){
 try{$rows=@(Get-Rows $m.Symbol $m.Zone "data/$($m.Id).csv");$s=@(Get-Signals $rows);$s | Export-Csv "data/$($m.Id)_signals.csv" -NoTypeInformation -Encoding UTF8;$r=$s[-1];$last=$s | Where-Object {$_.Action -ne '觀察'} | Select-Object -Last 1
 $qi=$script:QuoteInfo; $zoneLabel='台北時間';$displayTime=$null;if($qi.Time){$displayTime=[TimeZoneInfo]::ConvertTime($qi.Time,[TimeZoneInfo]::FindSystemTimeZoneById($m.Zone),[TimeZoneInfo]::FindSystemTimeZoneById('Taipei Standard Time'))}
 $quoteText='最新報價未提供';if($null -ne $qi.Price -and $null -ne $qi.Time){$quoteText="最新行情 $(F $qi.Price) 點 ｜ $($displayTime.ToString('yyyy-MM-dd HH:mm:ss'))（$zoneLabel）"}
 $marketText='非交易時段';if($qi.IsOpen){$marketText='交易時段（報價可能延遲）'}elseif($null -ne $qi.LastOpen){$marketText="最近一次開盤：$($qi.LastOpen.ToString('yyyy-MM-dd HH:mm'))（$zoneLabel）"}
 $retrievedText="取得時間：$($qi.Retrieved.ToString('yyyy-MM-dd HH:mm:ss'))（台北時間）"
 $touch=$qi.Touch
 $preview='非交易時段';$touchLine='盤中提醒僅在交易時段提供';$touchHtml=''
 if($qi.IsOpen){
  if($touch -and $touch.Date -eq $qi.LocalDate -and $touch.Date -gt $r.Date -and $qi.Time -and $qi.Time.ToString('yyyy-MM-dd') -eq $qi.LocalDate){
   $preview=Get-TouchPreview $s $touch
   $touchLine="$($touch.Date) ｜ 最高 $(F $touch.High) ｜ 最低 $(F $touch.Low)"
  }else{
   $preview='尚未取得今日盤中行情'
   $touchLine=$qi.LocalDate
  }
 }
 $touchHtml="<div class='signal'><small>$(E $qi.LocalDate) 盤中觸及提醒（非成交）</small><div class='signalrow'><h1>$(E $preview)</h1><div class='quote'>$(E $quoteText)</div></div><small>$(E $touchLine)</small></div>"
 $formalLabel="最近已確認交易日策略（$($r.Date)）"
 $buyLabel='—';if($null -ne $r.NextBuy){$buyLabel="$(F $r.NextBuy) 點（加碼 $($r.NextBuyWeight)%）"}
 $sellLabel='—';if($null -ne $r.NextSell){$sellLabel="$(F $r.NextSell) 點（出場 $($cfg.sell_ratio)%）"}
 $phaseText=$r.Phase;if($phaseText -eq '修正'){$phaseText='下跌／加碼階段'}elseif($phaseText -eq '反彈'){$phaseText='反彈／減碼階段'}
 $lastText='尚無觸發';if($last){$lastText="$(E $last.Action) ｜ $(F $last.Close) 點 ｜ $($last.Date)"}
 $table=(@($s | Select-Object -Last 15 | Sort-Object Date -Descending | ForEach-Object {"<tr><td>$($_.Date)</td><td>$(F $_.Open)</td><td>$(F $_.Close)</td><td>$(E $_.Action)</td></tr>"}) -join '')
 $cards+=@"
<section><h2>$(E $m.Name)</h2><small>資料日 $($r.Date) ｜ $($s.Count) 筆交易日</small>$touchHtml<div class="last"><small>$(E $formalLabel)</small><h2 style="font-size:20px;margin:5px 0;color:#526575">$(E $r.Action)</h2></div><h3>上次觸發（含最新一筆）</h3><div class="last">$lastText</div><dl><dt>開盤</dt><dd>$(F $r.Open)</dd><dt>收盤</dt><dd>$(F $r.Close)</dd><dt>策略階段</dt><dd>$(E $phaseText)</dd><dt>追蹤高點</dt><dd>$(F $r.High)</dd><dt>波段低點</dt><dd>$(F $r.Low)</dd><dt>下一加碼點</dt><dd>$(E $buyLabel)</dd><dt>下一出場點</dt><dd>$(E $sellLabel)</dd></dl><table><thead><tr><th>日期</th><th>開盤</th><th>收盤</th><th>訊號</th></tr></thead><tbody>$table</tbody></table></section>
"@
 }catch{$cards+="<section><h2>$(E $m.Name)</h2><h1>更新失敗</h1><p>$(E $_.Exception.Message)</p><p>不提供新訊號；已保存的歷史行情不會刪除。</p></section>"}
}
. (Join-Path $PSScriptRoot 'watchlist.ps1')
try{$watchHtml=Get-WatchHtml}catch{$watchHtml="<section><h2>觀察清單更新失敗</h2><p>$(E $_.Exception.Message)</p></section>"}
$schedule='每日排程尚未建立' 
if(-not $Silent){try{$taskName='MarketSignal-'+([BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($PSScriptRoot))).Replace('-','').Substring(0,16));$arg='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$PSCommandPath+'" -Silent';$act=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg;$tr=New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($cfg.update_time));$settings=New-ScheduledTaskSettingsSet -StartWhenAvailable;Register-ScheduledTask -TaskName $taskName -Action $act -Trigger $tr -Settings $settings -Force | Out-Null;$schedule='每日排程已建立（'+$cfg.update_time+'，使用本機時間）'}catch{$schedule='排程建立失敗；仍可雙擊手動更新。'}}else{$schedule='排程更新模式'}
if($env:GITHUB_ACTIONS -eq 'true'){$schedule='GitHub Actions 手動更新'}
$page=@"
<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>每日策略監控 V3.1</title><style>body{font:16px system-ui,sans-serif;background:#edf1f5;color:#243447;margin:24px}main{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(320px,100%),1fr));gap:20px}section{min-width:0;background:white;border:1px solid #d4dce5;border-radius:12px;padding:22px}h1{font-size:36px;margin:6px 0;color:#14665d}h2{margin:0 0 8px}h3{font-size:15px}small{color:#617080}.signal,.last{background:#edf5f3;padding:14px;border-radius:8px;margin-top:14px}.signalrow{display:flex;align-items:center;gap:16px;flex-wrap:wrap}.quote{font-size:13px;line-height:1.65;flex:1;min-width:175px}.quote small{font-size:12px}dl{display:grid;grid-template-columns:1fr 1fr;line-height:1.8}dd{text-align:right;margin:0;font-variant-numeric:tabular-nums}table{width:100%;border-collapse:collapse;font-size:13px}td,th{padding:7px 3px;text-align:right;border-bottom:1px solid #e1e6ec}td:first-child,th:first-child{text-align:left}footer{font-size:13px;color:#526575;margin-top:20px} .watchhead{margin-top:32px}.watchhead label{display:inline-block;margin:8px 0}.watchhead select{font:inherit;padding:5px}.watchhead p{font-size:13px;color:#617080}.watchsection{margin:24px 0 12px}.watchgroups{display:grid;grid-template-columns:1fr;gap:16px;align-items:start}@media(max-width:1100px){.watchgroups{grid-template-columns:1fr}}.watchgroup{margin:0;min-width:0}.watchgroup h3{margin:0 0 12px}.watchscroll{overflow-x:auto}.watchtable{min-width:1580px}.watchtable td,.watchtable th{padding:9px 10px;white-space:nowrap}.watchtable td:first-child,.watchtable th:first-child{position:sticky;left:0;background:white;z-index:1;white-space:normal;min-width:110px}.watchsort{font:inherit;color:inherit;background:none;border:0;padding:4px 0;cursor:pointer;white-space:nowrap}.watchsort:hover{text-decoration:underline}.watchsort::after{content:attr(data-arrow);font-size:10px}.watchsort[aria-pressed=true]{font-weight:700}.up{color:#b42332;font-weight:600}.down{color:#167348;font-weight:600}.flat{color:#617080}.groupoverview{margin-top:24px}.groupnote{font-size:13px;color:#617080}.gitem{padding:12px 0;border-bottom:1px solid #e1e6ec}.gtitle,.gline{display:flex;align-items:center;justify-content:space-between;gap:12px}.gline{margin:7px 0}.gline>b{min-width:85px;text-align:right;font-variant-numeric:tabular-nums}.gtrack{position:relative;flex:1;height:20px;background:#f2f4f7;border-radius:4px}.gtrack:before{content:"";position:absolute;left:50%;height:100%;border-left:1px solid #aab4bf}.gtrack i{position:absolute;top:2px;height:16px;border-radius:3px}.gtitle span{font-size:13px}</style></head><body><h2>每日投資策略監控 V4.5</h2><p>更新時間：$([TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow,'Taipei Standard Time').ToString('yyyy-MM-dd HH:mm:ss'))（台北時間） ｜ 訪問即更新</p><div style="margin:0 0 16px"><button id="manual-refresh" type="button" style="background:#14665d;color:white;padding:12px 18px;border:0;border-radius:7px;font:inherit;font-weight:600;cursor:pointer">更新行情</button> <button id="manual-auth" type="button" style="padding:10px;font:inherit">設定授權</button><div id="refresh-status" role="status" aria-live="polite" style="margin:8px 0;font-size:14px">開啟網頁即更新行情，亦可隨時手動更新；無定時排程。</div><div id="auth-panel" hidden style="padding:12px;background:white;border-radius:8px"><label for="refresh-token">GitHub Fine-grained token（僅限 market-signal，Actions 讀寫）</label><div style="display:flex;flex-wrap:wrap;gap:8px;margin:8px 0"><input id="refresh-token" type="password" autocomplete="off" spellcheck="false" style="flex:1;min-width:150px;font:inherit;padding:8px" placeholder="輸入授權金鑰"><button id="auth-use" type="button">確認授權</button><button id="auth-clear" type="button">清除授權</button></div><small><label><input type="checkbox" id="auth-remember" checked> 在此私人裝置記住授權</label><br>金鑰若選擇記住，會儲存在此瀏覽器的本機儲存空間（未加密），不會寫入公開儲存庫。請勿在共用裝置使用；可隨時清除授權。<a href="https://github.com/settings/personal-access-tokens/new" target="_blank" rel="noopener noreferrer">建立限定權限的金鑰</a></small></div></div><main>$($cards -join '')</main>$watchHtml<script>(function(){
const btn=document.getElementById('manual-refresh'),auth=document.getElementById('manual-auth'),panel=document.getElementById('auth-panel'),input=document.getElementById('refresh-token'),status=document.getElementById('refresh-status');
const base='https://api.github.com/repos/dyps883342/market-signal/actions';
const storageKey='marketsignal.market-signal.actions.token.v1',pendingKey='marketsignal.market-signal.pending.v1';
let token='',busy=false,localPending=null;
try{token=localStorage.getItem(storageKey)||'';}catch(e){}
const say=t=>{status.textContent=t;};
function getPending(){try{const p=JSON.parse(localStorage.getItem(pendingKey)||'null');if(p&&p.id&&Date.now()-p.time<3600000)return p;}catch(e){}return localPending&&Date.now()-localPending.time<3600000?localPending:null;}
function savePending(p){localPending=p;try{if(p)localStorage.setItem(pendingKey,JSON.stringify(p));else localStorage.removeItem(pendingKey);}catch(e){}}
function clearPending(id){if(getPending()?.id===id)savePending(null);}
function controls(v){busy=v;btn.disabled=v;auth.disabled=v;document.getElementById('auth-clear').disabled=v;document.getElementById('auth-use').disabled=v;}
auth.onclick=()=>{panel.hidden=!panel.hidden;if(!panel.hidden)input.focus();};
document.getElementById('auth-use').onclick=()=>{const v=input.value.trim();input.value='';if(!v.startsWith('github_pat_')){say('請使用限定此儲存庫的 Fine-grained token。');return;}token=v;panel.hidden=true;let saved=false;try{if(document.getElementById('auth-remember').checked){localStorage.setItem(storageKey,v);saved=true;}else{localStorage.removeItem(storageKey);}}catch(e){}say(saved?'已記住授權，正在更新行情。':'本頁已設定授權，正在更新行情。');refresh();};
document.getElementById('auth-clear').onclick=()=>{token='';input.value='';try{localStorage.removeItem(storageKey);}catch(e){say('無法清除瀏覽器儲存，請從瀏覽器設定清除網站資料。');return;}say('已清除此瀏覽器授權。');};
async function api(path,options){const h={Accept:'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28',Authorization:'Bearer '+token};if(options&&options.body)h['Content-Type']='application/json';const r=await fetch(base+path,Object.assign({cache:'no-store'},options||{},{headers:h}));if(!r.ok){const e=new Error('GitHub HTTP '+r.status);e.httpStatus=r.status;throw e;}return r;}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function acquireRequest(){
 const existing=getPending();if(existing)return existing;
 const p={id:crypto.randomUUID(),time:Date.now()};
 savePending(p);
 try{say('正在送出更新請求…');await api('/workflows/market.yml/dispatches',{method:'POST',body:JSON.stringify({ref:'main',inputs:{request_id:p.id}})});return p;}
 catch(e){if(e.httpStatus&&e.httpStatus>=400&&e.httpStatus<500)clearPending(p.id);throw e;}
}
async function refresh(){
 if(busy)return;
 if(!token){panel.hidden=false;say('首次使用請先設定授權，完成後即自動更新。');return;}
 controls(true);
 let p=null;
 try{
 p=navigator.locks&&navigator.locks.request?await navigator.locks.request('marketsignal-refresh-dispatch',acquireRequest):await acquireRequest();
 const title='MarketSignal manual '+p.id;
 say('已送出或已有更新請求，等待 GitHub 啟動…');
 for(let i=0;i<150;i++){
 await sleep(8000);
 const r=await api('/runs?event=workflow_dispatch&branch=main&per_page=30');
 const runs=(await r.json()).workflow_runs||[];
 const run=runs.find(x=>x.display_title===title);
 if(!run){say('更新請求已送出，等待執行紀錄…');continue;}
 if(run.status!=='completed'){say('更新中：'+(run.status==='in_progress'?'取得行情及部署':'等待執行'));continue;}
 clearPending(p.id);
 if(run.conclusion!=='success')throw new Error('更新未成功：'+run.conclusion+'；可手動重試。');
 say('部署完成，正在重新載入最新網站…');
 const url=new URL(location.href);url.searchParams.set('ms_updated',p.id);url.searchParams.set('refresh',String(Date.now()));location.replace(url.href);return;
 }
 say('更新請求尚未確認完成；再次開啟或按更新會優先追蹤同一請求。');
 }catch(e){if(e.httpStatus===401||e.httpStatus===403){say('授權無效或權限不足，請重新設定授權：'+e.message);panel.hidden=false;}else{say('更新異常：'+e.message+'；可查看 GitHub Actions。');}}
 finally{controls(false);}
}
btn.onclick=refresh;
const pageUrl=new URL(location.href);
if(pageUrl.searchParams.has('ms_updated')){
 pageUrl.searchParams.delete('ms_updated');pageUrl.searchParams.delete('refresh');history.replaceState(null,'',pageUrl.href);say('本次更新已完成。');
}else if(token){refresh();}else{say('首次使用請設定授權，之後開啟網頁將自動更新。');}
})();</script><script>(function(){document.querySelectorAll('.watchtable').forEach(function(table){const body=table.tBodies[0];let active='score',descending=true;const buttons=Array.from(table.querySelectorAll('.watchsort'));function sort(){const rows=Array.from(body.rows);function val(row){const cell=Array.from(row.cells).find(c=>c.dataset.key===active);return !cell||!cell.dataset.value?null:Number(cell.dataset.value);}rows.sort((a,b)=>{const x=val(a),y=val(b);if(x===null)return y===null?0:1;if(y===null)return -1;return descending?y-x:x-y;});rows.forEach(r=>body.appendChild(r));buttons.forEach(b=>{const selected=b.dataset.key===active;b.setAttribute('aria-pressed',String(selected));b.closest('th').setAttribute('aria-sort',selected?(descending?'descending':'ascending'):'none');b.dataset.arrow=selected?(descending?' ▼':' ▲'):'';});}buttons.forEach(b=>b.addEventListener('click',()=>{const key=b.dataset.key;if(key===active)descending=!descending;else{active=key;descending=true;}sort();}));sort();});})();</script><footer>正式策略採各市場最近已完成交易日收盤價；盤中觸及提醒依最新日K最高最低點試算，不寫入正式交易紀錄。日內高低順序不明，可能同時觸及不同方向，請核對實際行情。交易時段依資料源提供的時段判斷，報價可能延遲。加碼按本輪起始現金；出場按當下持有股數。跨多個門檻合併計算。同一資料日期與動作不可重複執行。歷史訊號不代表成交，休市時沿用最近交易日，請核對日期。半年初始化可能無法辨識更早波段。Yahoo Finance 為非保證服務。</footer></body></html>
"@
$page | Set-Content dashboard.html -Encoding UTF8
if(-not $Silent){Start-Process (Join-Path $PSScriptRoot 'dashboard.html')}
}finally{$mutex.ReleaseMutex();$mutex.Dispose()}
