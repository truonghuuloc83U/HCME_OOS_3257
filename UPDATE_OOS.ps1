$ErrorActionPreference='Stop'
$folder=Split-Path -Parent $MyInvocation.MyCommand.Path
$input=Join-Path $folder 'DATA_INPUT.xlsx'
$htmlPath=Join-Path $folder 'OOS_Dashboard.html'
$excel=$null;$wb=$null
function Norm([string]$s){if(!$s){return ''};$d=$s.Normalize([Text.NormalizationForm]::FormD);$b=New-Object Text.StringBuilder;foreach($c in $d.ToCharArray()){if([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)-ne [Globalization.UnicodeCategory]::NonSpacingMark){[void]$b.Append($c)}};return (($b.ToString().Replace([char]0x0111,'d').Replace([char]0x0110,'D') -replace '[^a-zA-Z0-9]+',' ').Trim().ToLower())}
function HMap($a,$r,$cols){$h=@{};for($c=1;$c-le$cols;$c++){$k=Norm ([string]$a[$r,$c]);if($k){$h[$k]=$c}};return $h}
function FindCol($h,$names){foreach($n in $names){if($h.ContainsKey($n)){return [int]$h[$n]}};foreach($k in $h.Keys){foreach($n in $names){if($k.Contains($n)){return [int]$h[$k]}}};return 0}
function ToNum($v){if($null-eq$v-or$v-eq''){return 0};try{return [double]$v}catch{return 0}}
try{
 if(!(Test-Path $input)){throw 'DATA_INPUT.xlsx not found'}
 if(!(Test-Path $htmlPath)){throw 'OOS_Dashboard.html not found'}
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$wb=$excel.Workbooks.Open($input,0,$true)
 $bestO=$null;$bestR=$null
 foreach($ws in @($wb.Worksheets)){
  $arr=$ws.UsedRange.Value2;$rows=$ws.UsedRange.Rows.Count;$cols=$ws.UsedRange.Columns.Count
  for($rr=1;$rr-le[Math]::Min(12,$rows);$rr++){
   $hm=HMap $arr $rr ([Math]::Min(60,$cols));$os=0;$rs=0
   foreach($x in @('ma ch','ten san pham','so luong ke hoach','so luong thuc te hien tai','oos')){if((FindCol $hm @($x))-gt0){$os++}}
   foreach($x in @('ma ch','ten ch','ct tham gia','loai hinh ch','ten npp')){if((FindCol $hm @($x))-gt0){$rs++}}
   if($os-ge4-and($null-eq$bestO-or$os-gt$bestO.Score)){$bestO=[pscustomobject]@{A=$arr;H=$hm;R=$rr;Score=$os;Name=$ws.Name}}
   if($rs-ge4-and($null-eq$bestR-or$rs-gt$bestR.Score)){$bestR=[pscustomobject]@{A=$arr;H=$hm;R=$rr;Score=$rs;Name=$ws.Name}}
  }
 }
 if($null-eq$bestO-or$null-eq$bestR){throw 'Required sheets not detected'}
 Write-Host ('OOS: '+$bestO.Name+' | RESULT: '+$bestR.Name) -ForegroundColor Cyan
 $ra=$bestR.A;$rh=$bestR.H;$rcStore=FindCol $rh @('ma ch');$rcNpp=FindCol $rh @('ten npp','nha phan phoi');$rcSet=FindCol $rh @('ct tham gia','ten muc tb');$rcProg=FindCol $rh @('loai hinh ch','chuong trinh trung bay')
 $lookup=@{}
 for($r=$bestR.R+1;$r-le$ra.GetLength(0);$r++){
  $sid=[string]$ra[$r,$rcStore];if(!$sid){continue}
  $npp=if($rcNpp-gt0){[string]$ra[$r,$rcNpp]}else{''};$set=if($rcSet-gt0){[string]$ra[$r,$rcSet]}else{''};$prog=if($rcProg-gt0){[string]$ra[$r,$rcProg]}else{''}
  if((Norm $prog)-eq'perfectstore'){$prog='PS'}
  $lookup[$npp+'|'+$sid+'|'+$set]=$prog
 }
 $a=$bestO.A;$h=$bestO.H
 $cStore=FindCol $h @('ma ch','store id');$cName=FindCol $h @('ten ch','store name');$cNpp=FindCol $h @('nha phan phoi','ten npp');$cSet=FindCol $h @('ten muc tb','ct tham gia');$cBu=FindCol $h @('bu');$cCat=FindCol $h @('nganh hang','category');$cCode=FindCol $h @('productcode unilever','ma sku');$cProd=FindCol $h @('ten san pham','product name');$cTar=FindCol $h @('so luong ke hoach','target');$cAct=FindCol $h @('so luong thuc te hien tai','actual');$cOos=FindCol $h @('so luong oos hien tai','oos');$cVal=FindCol $h @('gia tri hang thieu','oos value')
 # Read optional store-to-salesman mapping sheet
 $salesMap=@{}
 foreach($mws in @($wb.Worksheets)){
  if((Norm $mws.Name)-eq'mapping ch salesman'){
   $ma=$mws.UsedRange.Value2;$mr=$mws.UsedRange.Rows.Count;$mc=$mws.UsedRange.Columns.Count
   for($mhr=1;$mhr-le[Math]::Min(12,$mr);$mhr++){
    $mh=HMap $ma $mhr ([Math]::Min(30,$mc));$mStore=FindCol $mh @('ma ch');$mCode=FindCol $mh @('ma salesman');$mName=FindCol $mh @('ten salesman');$mRoute=FindCol $mh @('tuyen route','tuyen','route');$mStatus=FindCol $mh @('trang thai')
    if($mStore-gt0-and$mName-gt0){
     for($rr=$mhr+1;$rr-le$mr;$rr++){
      $msid=[string]$ma[$rr,$mStore];if(!$msid){continue}
      $active=if($mStatus-gt0){[string]$ma[$rr,$mStatus]}else{''}
      if($active-and(Norm $active)-ne'dang phu trach'){continue}
      $salesMap[$msid]=[pscustomobject]@{code=if($mCode-gt0){[string]$ma[$rr,$mCode]}else{''};name=[string]$ma[$rr,$mName];route=if($mRoute-gt0){[string]$ma[$rr,$mRoute]}else{''}}
     }
     break
    }
   }
  }
 }
 Write-Host ('SALESMAN MAPPING: '+$salesMap.Count+' stores') -ForegroundColor Cyan
 $list=New-Object Collections.Generic.List[object]
 for($r=$bestO.R+1;$r-le$a.GetLength(0);$r++){
  $sid=[string]$a[$r,$cStore];if(!$sid){continue}
  $npp=if($cNpp-gt0){[string]$a[$r,$cNpp]}else{''};$set=if($cSet-gt0){[string]$a[$r,$cSet]}else{''}
  $target=ToNum ($a[$r,$cTar]);$actual=ToNum ($a[$r,$cAct]);$oos=ToNum ($a[$r,$cOos]);$value=if($cVal-gt0){ToNum ($a[$r,$cVal])}else{0}
  $obj=[pscustomobject]@{npp=$npp;storeId=$sid;store=if($cName-gt0){[string]$a[$r,$cName]}else{$sid};program=$set;displayProgram=$lookup[$npp+'|'+$sid+'|'+$set];bu=if($cBu-gt0){[string]$a[$r,$cBu]}else{''};cat=if($cCat-gt0){[string]$a[$r,$cCat]}else{''};code=if($cCode-gt0){[string]$a[$r,$cCode]}else{''};product=[string]$a[$r,$cProd];salesmanCode=if($salesMap.ContainsKey($sid)){$salesMap[$sid].code}else{''};salesman=if($salesMap.ContainsKey($sid)){$salesMap[$sid].name}else{''};route=if($salesMap.ContainsKey($sid)){$salesMap[$sid].route}else{''};target=$target;actual=$actual;oos=$oos;value=$value}
  [void]$list.Add($obj)
 }
 $json=ConvertTo-Json $list -Depth 4 -Compress;$payload='/*DATA_START*/window.OOS_DATA='+$json+';/*DATA_END*/'
 $html=[IO.File]::ReadAllText($htmlPath,[Text.Encoding]::UTF8)
 if($html -notmatch '/\*DATA_START\*/'){throw 'HTML data marker not found'}
 $updated=[regex]::Replace($html,'(?s)/\*DATA_START\*/.*?/\*DATA_END\*/',[Text.RegularExpressions.MatchEvaluator]{param($m)$payload},1)
 [IO.File]::WriteAllText($htmlPath,$updated,(New-Object Text.UTF8Encoding($false)))
 Write-Host ('DONE: '+$list.Count+' rows embedded into OOS_Dashboard.html') -ForegroundColor Green
}catch{Write-Host ('ERROR: '+$_.Exception.Message) -ForegroundColor Red;Write-Host ('LINE: '+$_.InvocationInfo.ScriptLineNumber) -ForegroundColor Yellow;exit 1}finally{if($wb){$wb.Close($false)};if($excel){$excel.Quit()}}
