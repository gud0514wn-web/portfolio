/**
 * 자산관리 HTML용 시세 프록시 (Google Apps Script)
 * 배포: 배포 > 새 배포 > 웹 앱 > 실행 사용자: 나 > 액세스 권한: 모든 사용자
 * 배포 후 /exec URL을 자산관리 HTML의 설정에 입력하세요.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    var action = String(p.action || '');
    if (action.indexOf('sync') === 0) {
      out = handleSyncAction_(action, p);
    } else {
      out = { ok: true, timestamp: new Date().toISOString(), prices: {}, fx: null, fxSource: null, fxTimestamp: null, errors: {} };
      var us = splitList_(p.us), kr = splitList_(p.kr), crypto = splitList_(p.crypto);
      fetchUS_(us, out); fetchKR_(kr, out); fetchCrypto_(crypto, out);
      if (p.fx === '1') fetchFx_(out);
    }
  } catch (err) {
    out = { ok:false, timestamp:new Date().toISOString(), message:String(err && err.message ? err.message : err) };
  }
  return output_(out, p.callback);
}
function output_(obj, callback) {
  var body = JSON.stringify(obj), cb = String(callback || '');
  if (/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}


var SYNC_SPREADSHEET_PROPERTY='PORTFOLIO_SYNC_SPREADSHEET_ID';
var SYNC_ASSET_SHEET='Assets';
var SYNC_META_SHEET='Meta';
var SYNC_HISTORY_SHEET='History';

function handleSyncAction_(action,p){
  var key=String(p.key||'').trim();
  if(!key||key.length<6)throw new Error('동기화 코드가 올바르지 않습니다.');
  var ss=getSyncSpreadsheet_(),ash=ss.getSheetByName(SYNC_ASSET_SHEET),msh=ss.getSheetByName(SYNC_META_SHEET),hsh=ss.getSheetByName(SYNC_HISTORY_SHEET);
  if(action==='syncInit')return{ok:true,updatedAt:new Date().toISOString(),assetCount:countAssetsForKey_(ash,key),spreadsheetUrl:ss.getUrl()};
  var lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    if(action==='syncGet'){
      var assets=readAssetsForKey_(ash,key),meta=readMetaForKey_(msh,key);
      return{ok:true,assets:assets,accounts:meta.accounts,goalAssetTarget:Number(meta.goalAssetTarget||0),updatedAt:meta.updatedAt||latestAssetUpdatedAt_(ash,key),spreadsheetUrl:ss.getUrl()};
    }
    if(action==='syncUpsert'){upsertAsset_(ash,key,JSON.parse(String(p.data||'{}')));return{ok:true,updatedAt:new Date().toISOString()}}
    if(action==='syncBatchUpsert'){
      var list=JSON.parse(String(p.data||'[]'));if(!Array.isArray(list))throw new Error('자산 데이터 형식 오류');
      list.forEach(function(a){upsertAsset_(ash,key,a)});return{ok:true,updatedAt:new Date().toISOString(),count:list.length}
    }
    if(action==='syncDelete'){deleteAsset_(ash,key,String(p.id||''));return{ok:true,updatedAt:new Date().toISOString()}}
    if(action==='syncClear'){clearKeyAssets_(ash,key);return{ok:true,updatedAt:new Date().toISOString()}}
    if(action==='syncMeta'){
      var acc=[];try{acc=JSON.parse(String(p.accounts||'[]'))}catch(e){}if(!Array.isArray(acc))acc=[];
      var goalAssetTarget=Number(p.goalAssetTarget||0);
      upsertMeta_(msh,key,acc,goalAssetTarget);return{ok:true,updatedAt:new Date().toISOString()}
    }
    if(action==='syncHistoryUpsert'){
      var snap=JSON.parse(String(p.data||'{}'));
      var result=upsertDailyHistory_(hsh,key,snap);
      return{ok:true,saved:result.saved,keptValue:result.keptValue,updatedAt:result.updatedAt}
    }
    if(action==='syncHistoryGet'){
      return{ok:true,history:readHistorySummaries_(hsh,key),updatedAt:new Date().toISOString()}
    }
    if(action==='syncHistoryDetail'){
      return{ok:true,date:String(p.date||''),assets:readHistoryDetail_(hsh,key,String(p.date||'')),updatedAt:new Date().toISOString()}
    }
    throw new Error('지원하지 않는 동기화 작업입니다: '+action);
  }finally{lock.releaseLock()}
}

/**
 * 최초 1회만 Apps Script 편집기에서 직접 실행하세요.
 * Google Sheets 생성 권한 승인을 받은 뒤 동기화용 스프레드시트를 준비합니다.
 */
function setupPortfolioSync() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SYNC_SPREADSHEET_PROPERTY);
  var ss = null;

  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) {}
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Portfolio Cloud Sync');
    props.setProperty(SYNC_SPREADSHEET_PROPERTY, ss.getId());
  }

  ensureSheet_(ss, SYNC_ASSET_SHEET, [
    'syncKey','id','account','name','qty','avg','purchaseFx','cost','price',
    'currency','market','ticker','targetWeight','memo','updatedAt'
  ]);
  ensureSheet_(ss, SYNC_META_SHEET, ['syncKey','accountsJson','updatedAt']);
  ensureSheet_(ss, SYNC_HISTORY_SHEET, ['syncKey','date','totalValue','totalCost','totalPL','returnPct','fx','assetsJson','updatedAt']);
  migrateKrTickers_(ss.getSheetByName(SYNC_ASSET_SHEET));

  Logger.log('Portfolio Cloud Sync 준비 완료: ' + ss.getUrl());
  return ss.getUrl();
}

function getSyncSpreadsheet_(){
  var props=PropertiesService.getScriptProperties();
  var id=props.getProperty(SYNC_SPREADSHEET_PROPERTY);

  if(!id){
    throw new Error('초기 설정이 필요합니다. Apps Script 편집기에서 setupPortfolioSync 함수를 1회 실행해 Google Sheets 권한을 승인한 뒤 다시 시도하세요.');
  }

  var ss;
  try{
    ss=SpreadsheetApp.openById(id);
  }catch(e){
    throw new Error('동기화용 스프레드시트를 열 수 없습니다. Apps Script 편집기에서 setupPortfolioSync 함수를 다시 실행해 주세요.');
  }

  ensureSheet_(ss,SYNC_ASSET_SHEET,['syncKey','id','account','name','qty','avg','purchaseFx','cost','price','currency','market','ticker','targetWeight','memo','updatedAt']);
  ensureSheet_(ss,SYNC_META_SHEET,['syncKey','accountsJson','updatedAt']);
  ensureSheet_(ss,SYNC_HISTORY_SHEET,['syncKey','date','totalValue','totalCost','totalPL','returnPct','fx','assetsJson','updatedAt']);
  return ss;
}
function ensureSheet_(ss,name,headers){
  var sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1)}
  if(name===SYNC_ASSET_SHEET){
    sh.getRange('L:L').setNumberFormat('@');
  }
  if(name===SYNC_HISTORY_SHEET){
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  }
  return sh;
}

function migrateKrTickers_(sh){
  if(!sh)return;
  var last=sh.getLastRow();
  if(last<2)return;
  var vals=sh.getRange(2,11,last-1,2).getValues(); // K=market, L=ticker
  var changed=false;
  for(var i=0;i<vals.length;i++){
    var market=String(vals[i][0]||'');
    var ticker=String(vals[i][1]||'');
    if(market==='KR'){
      var normalized=normalizeKrCode_(ticker);
      if(normalized!==ticker){
        vals[i][1]=normalized;
        changed=true;
      }
    }
  }
  if(changed)sh.getRange(2,11,vals.length,2).setValues(vals);
  sh.getRange('L:L').setNumberFormat('@');
}

function countAssetsForKey_(sh,key){var last=sh.getLastRow();if(last<2)return 0;var v=sh.getRange(2,1,last-1,1).getValues(),n=0;v.forEach(function(r){if(String(r[0])===key)n++});return n}
function assetToRow_(key,a){
  var market=String(a.market||'MANUAL');
  var ticker=String(a.ticker||'');
  if(market==='KR')ticker=normalizeKrCode_(ticker);
  return[key,Number(a.id||0),String(a.account||''),String(a.name||''),Number(a.qty||0),Number(a.avg||0),a.purchaseFx==null||a.purchaseFx===''?'':Number(a.purchaseFx),Number(a.cost||0),Number(a.price||0),String(a.currency||'KRW'),market,ticker,a.targetWeight==null||a.targetWeight===''?'':Number(a.targetWeight),String(a.memo||''),new Date().toISOString()]
}
function rowToAsset_(r){
  var market=String(r[10]||'MANUAL');
  var ticker=String(r[11]||'');
  if(market==='KR')ticker=normalizeKrCode_(ticker);
  return{id:Number(r[1]||0),account:String(r[2]||''),name:String(r[3]||''),qty:Number(r[4]||0),avg:Number(r[5]||0),purchaseFx:r[6]===''?null:Number(r[6]),cost:Number(r[7]||0),price:Number(r[8]||0),currency:String(r[9]||'KRW'),market:market,ticker:ticker,targetWeight:r[12]===''?'':Number(r[12]),memo:String(r[13]||'')}
}
function findAssetRow_(sh,key,id){var last=sh.getLastRow();if(last<2)return-1;var v=sh.getRange(2,1,last-1,2).getValues();for(var i=0;i<v.length;i++)if(String(v[i][0])===key&&String(v[i][1])===String(id))return i+2;return-1}
function upsertAsset_(sh,key,a){if(!a||!a.id)throw new Error('자산 ID가 없습니다.');var row=assetToRow_(key,a),n=findAssetRow_(sh,key,a.id);if(n>0)sh.getRange(n,1,1,row.length).setValues([row]);else sh.appendRow(row)}
function deleteAsset_(sh,key,id){var n=findAssetRow_(sh,key,id);if(n>0)sh.deleteRow(n)}
function clearKeyAssets_(sh,key){var last=sh.getLastRow();if(last<2)return;var v=sh.getRange(2,1,last-1,1).getValues();for(var i=v.length-1;i>=0;i--)if(String(v[i][0])===key)sh.deleteRow(i+2)}
function readAssetsForKey_(sh,key){var last=sh.getLastRow();if(last<2)return[];var v=sh.getRange(2,1,last-1,15).getValues(),o=[];v.forEach(function(r){if(String(r[0])===key)o.push(rowToAsset_(r))});return o}
function latestAssetUpdatedAt_(sh,key){var last=sh.getLastRow();if(last<2)return'';var v=sh.getRange(2,1,last-1,15).getValues(),x='';v.forEach(function(r){if(String(r[0])===key&&String(r[14]||'')>x)x=String(r[14]||'')});return x}
function findMetaRow_(sh,key){var last=sh.getLastRow();if(last<2)return-1;var v=sh.getRange(2,1,last-1,1).getValues();for(var i=0;i<v.length;i++)if(String(v[i][0])===key)return i+2;return-1}
function upsertMeta_(sh,key,accounts,goalAssetTarget){
  var payload={accounts:Array.isArray(accounts)?accounts:[],goalAssetTarget:Number(goalAssetTarget||0)};
  var row=[key,JSON.stringify(payload),new Date().toISOString()],n=findMetaRow_(sh,key);
  if(n>0)sh.getRange(n,1,1,3).setValues([row]);else sh.appendRow(row)
}
function readMetaForKey_(sh,key){
  var n=findMetaRow_(sh,key);
  if(n<0)return{accounts:[],goalAssetTarget:0,updatedAt:''};
  var r=sh.getRange(n,1,1,3).getValues()[0],raw=null,accounts=[],goalAssetTarget=0;
  try{raw=JSON.parse(String(r[1]||'[]'))}catch(e){}
  if(Array.isArray(raw)){
    accounts=raw;
  }else if(raw&&typeof raw==='object'){
    accounts=Array.isArray(raw.accounts)?raw.accounts:[];
    goalAssetTarget=Number(raw.goalAssetTarget||0);
  }
  return{accounts:accounts,goalAssetTarget:goalAssetTarget,updatedAt:String(r[2]||'')}
}


function findHistoryRow_(sh,key,date){
  var last=sh.getLastRow();if(last<2)return-1;
  var vals=sh.getRange(2,1,last-1,2).getValues();
  for(var i=0;i<vals.length;i++)if(String(vals[i][0])===key&&String(vals[i][1])===date)return i+2;
  return-1;
}
function upsertDailyHistory_(sh,key,snap){
  var date=String(snap.date||'').trim();if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('자산변화 날짜 형식 오류');
  var totalValue=Number(snap.totalValue||0);if(!(totalValue>=0))throw new Error('총자산 값 오류');
  var now=String(snap.capturedAt||new Date().toISOString());
  var rowNo=findHistoryRow_(sh,key,date),existing=-1;
  if(rowNo>0)existing=Number(sh.getRange(rowNo,3).getValue()||0);
  if(rowNo>0&&existing>=totalValue)return{saved:false,keptValue:existing,updatedAt:String(sh.getRange(rowNo,9).getValue()||now)};
  var assets=Array.isArray(snap.assets)?snap.assets:[];
  var row=[key,date,totalValue,Number(snap.totalCost||0),Number(snap.totalPL||0),Number(snap.returnPct||0),Number(snap.fx||0),JSON.stringify(assets),now];
  if(rowNo>0)sh.getRange(rowNo,1,1,row.length).setValues([row]);else sh.appendRow(row);
  sh.getRange('B:B').setNumberFormat('@');sh.getRange('H:H').setNumberFormat('@');
  return{saved:true,keptValue:totalValue,updatedAt:now};
}
function readHistorySummaries_(sh,key){
  var last=sh.getLastRow();if(last<2)return[];
  var vals=sh.getRange(2,1,last-1,9).getValues(),out=[];
  vals.forEach(function(r){if(String(r[0])===key)out.push({date:String(r[1]||''),totalValue:Number(r[2]||0),totalCost:Number(r[3]||0),totalPL:Number(r[4]||0),returnPct:Number(r[5]||0),fx:Number(r[6]||0),updatedAt:String(r[8]||'')})});
  out.sort(function(a,b){return String(a.date).localeCompare(String(b.date))});return out;
}
function readHistoryDetail_(sh,key,date){
  var rowNo=findHistoryRow_(sh,key,date);if(rowNo<0)return[];
  var raw=String(sh.getRange(rowNo,8).getValue()||'[]'),arr=[];try{arr=JSON.parse(raw)}catch(e){};
  if(!Array.isArray(arr))arr=[];return arr;
}

function splitList_(s) {
  if (!s) return [];
  var seen = {};
  return String(s).split(',').map(function(x){ return x.trim().toUpperCase(); })
    .filter(function(x){ if (!x || seen[x]) return false; seen[x] = true; return true; });
}

function fetchUS_(symbols, out) {
  if (!symbols.length) return;

  function makeReq_(host, sym) {
    return {
      url: 'https://' + host + '/v8/finance/chart/' + encodeURIComponent(sym) +
           '?range=1d&interval=1m&includePrePost=true',
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
        'Accept':'application/json,text/plain,*/*',
        'Referer':'https://finance.yahoo.com/'
      }
    };
  }

  function parseYahoo_(r) {
    var code = r.getResponseCode();
    if (code < 200 || code >= 300) throw new Error('HTTP ' + code);

    var j = JSON.parse(r.getContentText());
    var result = j && j.chart && j.chart.result && j.chart.result[0];
    if (!result) {
      var err = j && j.chart && j.chart.error;
      throw new Error(err && err.description ? err.description : 'Yahoo 응답 없음');
    }

    var meta = result.meta || {};
    var ts = result.timestamp || [];
    var quote = result.indicators && result.indicators.quote &&
                result.indicators.quote[0] || {};
    var closes = quote.close || [];

    var latestPrice = null;
    var latestTs = null;

    // includePrePost=true 이므로 프리/정규/애프터 중 가장 최근 유효 1분봉 가격을 사용.
    for (var k = closes.length - 1; k >= 0; k--) {
      var p = Number(closes[k]);
      if (p > 0) {
        latestPrice = p;
        latestTs = ts[k] || null;
        break;
      }
    }

    // 1분봉이 비어 있는 경우 Yahoo meta 값으로 폴백.
    if (!(latestPrice > 0)) {
      latestPrice = Number(meta.postMarketPrice) ||
                    Number(meta.preMarketPrice) ||
                    Number(meta.regularMarketPrice);
    }

    if (!(latestPrice > 0)) throw new Error('유효한 가격 없음');

    var session = 'LATEST';
    var ctp = meta.currentTradingPeriod || {};
    var checkTs = latestTs || Math.floor(Date.now()/1000);

    function inPeriod_(period, t) {
      return period && Number(period.start) && Number(period.end) &&
             t >= Number(period.start) && t <= Number(period.end);
    }

    if (inPeriod_(ctp.pre, checkTs)) session = 'PRE';
    else if (inPeriod_(ctp.regular, checkTs)) session = 'REGULAR';
    else if (inPeriod_(ctp.post, checkTs)) session = 'POST';

    var previousClose =
      Number(meta.chartPreviousClose) ||
      Number(meta.previousClose) ||
      Number(meta.regularMarketPreviousClose) || null;

    return {
      price: latestPrice,
      previousClose: previousClose,
      currency: meta.currency || 'USD',
      source: 'Yahoo Finance',
      session: session,
      timestamp: latestTs
    };
  }

  // 1차: query1
  var reqs1 = symbols.map(function(sym){ return makeReq_('query1.finance.yahoo.com', sym); });
  var res1 = UrlFetchApp.fetchAll(reqs1);

  var failed = [];
  res1.forEach(function(r, i){
    var sym = symbols[i];
    try {
      out.prices['US:' + sym] = parseYahoo_(r);
    } catch(err) {
      failed.push(sym);
    }
  });

  // 2차: query2 재시도
  if (failed.length) {
    var reqs2 = failed.map(function(sym){ return makeReq_('query2.finance.yahoo.com', sym); });
    var res2 = UrlFetchApp.fetchAll(reqs2);

    res2.forEach(function(r, i){
      var sym = failed[i];
      try {
        out.prices['US:' + sym] = parseYahoo_(r);
      } catch(err) {
        out.errors['US:' + sym] = 'Yahoo Finance 조회 실패: ' +
          String(err && err.message ? err.message : err);
      }
    });
  }

  // 드물게 chartPreviousClose가 빠지는 경우 5일 일봉으로 보완.
  var prevMissing = symbols.filter(function(sym){
    var q = out.prices['US:' + sym];
    return q && Number(q.price) > 0 && !(Number(q.previousClose) > 0);
  });
  if (prevMissing.length) {
    try {
      var dailyReqs = prevMissing.map(function(sym){
        return makeReq_('query1.finance.yahoo.com', sym);
      }).map(function(req){
        req.url = req.url.replace('range=1d&interval=1m&includePrePost=true','range=5d&interval=1d&includePrePost=false');
        return req;
      });
      var dailyRes = UrlFetchApp.fetchAll(dailyReqs);
      dailyRes.forEach(function(r,i){
        try{
          if(r.getResponseCode()<200||r.getResponseCode()>=300)return;
          var j=JSON.parse(r.getContentText());
          var result=j&&j.chart&&j.chart.result&&j.chart.result[0];
          if(!result)return;
          var meta=result.meta||{};
          var prev=Number(meta.chartPreviousClose)||Number(meta.previousClose)||null;
          if(!(prev>0)){
            var closes=result.indicators&&result.indicators.quote&&result.indicators.quote[0]&&result.indicators.quote[0].close||[];
            var good=closes.map(Number).filter(function(v){return v>0});
            if(good.length>=2) prev=good[good.length-2];
          }
          if(prev>0&&out.prices['US:'+prevMissing[i]]) out.prices['US:'+prevMissing[i]].previousClose=prev;
        }catch(_e){}
      });
    }catch(_e2){}
  }

}

function normalizeKrCode_(code) {
  var s = String(code == null ? '' : code).trim();
  if (/^\d+$/.test(s) && s.length < 6) s = ('000000' + s).slice(-6);
  return s;
}


function signedKrChange_(obj) {
  if (!obj) return null;
  var raw = obj.compareToPreviousClosePrice;
  if (raw == null || raw === '') raw = obj.changePrice;
  if (raw == null || raw === '') raw = obj.cv;
  if (raw == null || raw === '') return null;

  var v = Number(String(raw).replace(/,/g,''));
  if (!isFinite(v)) return null;

  // Naver basic often provides direction separately.
  var code = String(
    (obj.compareToPreviousPrice && (obj.compareToPreviousPrice.code || obj.compareToPreviousPrice.name)) ||
    obj.compareToPreviousPriceCode || obj.rf || ''
  ).toUpperCase();

  // 2/5 or RISE/UP = 상승, 4/3 or FALL/DOWN = 하락 (unknown이면 원래 부호 사용)
  if (v >= 0) {
    if (code === '5' || code === '2' || code.indexOf('RISE') >= 0 || code.indexOf('UP') >= 0) return Math.abs(v);
    if (code === '4' || code === '3' || code.indexOf('FALL') >= 0 || code.indexOf('DOWN') >= 0) return -Math.abs(v);
  }
  return v;
}
function deriveKrPreviousClose_(current, obj) {
  if (!(Number(current) > 0)) return null;
  var direct = Number(String(
    (obj && (obj.previousClosePrice || obj.prevClosePrice || obj.previousClose || obj.pcv)) || ''
  ).replace(/,/g,''));
  if (direct > 0) return direct;

  var change = signedKrChange_(obj);
  if (change != null) {
    var prev = Number(current) - Number(change);
    if (prev > 0) return prev;
  }

  var ratio = Number(String((obj && (obj.fluctuationsRatio || obj.changeRate || obj.cr)) || '').replace(/,/g,''));
  if (isFinite(ratio) && ratio > -100 && ratio !== 0) {
    var prevByRatio = Number(current) / (1 + ratio / 100);
    if (prevByRatio > 0) return prevByRatio;
  }
  return null;
}

function fetchKR_(codes, out) {
  if (!codes.length) return;

  codes.forEach(function(rawCode){
    var requestCode = String(rawCode == null ? '' : rawCode).trim();
    var code = normalizeKrCode_(requestCode);
    var responseKey = 'KR:' + requestCode;
    var ok = false;
    var errs = [];

    // A. 네이버 모바일 basic
    try {
      var r1 = UrlFetchApp.fetch(
        'https://m.stock.naver.com/api/stock/' + encodeURIComponent(code) + '/basic',
        {
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/124 Mobile Safari/537.36',
            'Accept':'application/json,text/plain,*/*',
            'Referer':'https://m.stock.naver.com/'
          }
        }
      );
      if (r1.getResponseCode() >= 200 && r1.getResponseCode() < 300) {
        var j1 = JSON.parse(r1.getContentText());
        var p1 = Number(String(
          j1.closePrice || j1.currentPrice || j1.tradePrice || ''
        ).replace(/,/g,''));
        if (p1 > 0) {
          out.prices[responseKey] = {
            price:p1,
            currency:'KRW',
            name:j1.stockName || '',
            source:'Naver Basic',
            previousClose:deriveKrPreviousClose_(p1,j1),
            marketStatus:j1.marketStatus || ''
          };
          ok = true;
        }
      } else {
        errs.push('NaverBasic HTTP ' + r1.getResponseCode());
      }
    } catch(e1) {
      errs.push('NaverBasic ' + String(e1 && e1.message ? e1.message : e1));
    }
    if (ok) return;

    // B. 네이버 실시간 domestic endpoint
    try {
      var r2 = UrlFetchApp.fetch(
        'https://polling.finance.naver.com/api/realtime/domestic/stock/' + encodeURIComponent(code),
        {
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
            'Accept':'application/json,text/plain,*/*',
            'Referer':'https://finance.naver.com/'
          }
        }
      );
      if (r2.getResponseCode() >= 200 && r2.getResponseCode() < 300) {
        var j2 = JSON.parse(r2.getContentText());
        var d2 = j2 && j2.datas && j2.datas[0];
        var p2 = d2 && Number(String(
          d2.closePrice || d2.currentPrice || d2.tradePrice || d2.nv || ''
        ).replace(/,/g,''));
        if (p2 > 0) {
          out.prices[responseKey] = {
            price:p2,
            currency:'KRW',
            name:d2.stockName || d2.name || '',
            source:'Naver Realtime',
            previousClose:deriveKrPreviousClose_(p2,d2),
            marketStatus:d2.marketStatus || d2.ms || ''
          };
          ok = true;
        }
      } else {
        errs.push('NaverRealtime HTTP ' + r2.getResponseCode());
      }
    } catch(e2) {
      errs.push('NaverRealtime ' + String(e2 && e2.message ? e2.message : e2));
    }
    if (ok) return;

    // C. 네이버 구형 polling endpoint
    try {
      var r3 = UrlFetchApp.fetch(
        'https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:' + encodeURIComponent(code),
        {
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
            'Accept':'application/json,text/plain,*/*',
            'Referer':'https://finance.naver.com/'
          }
        }
      );
      if (r3.getResponseCode() >= 200 && r3.getResponseCode() < 300) {
        var j3 = JSON.parse(r3.getContentText());
        var d3 = j3 && j3.result && j3.result.areas && j3.result.areas[0] &&
                 j3.result.areas[0].datas && j3.result.areas[0].datas[0];
        var p3 = d3 && Number(String(d3.nv || d3.closePrice || '').replace(/,/g,''));
        if (p3 > 0) {
          out.prices[responseKey] = {
            price:p3,
            currency:'KRW',
            source:'Naver Polling',
            previousClose:deriveKrPreviousClose_(p3,d3),
            marketStatus:d3.ms || ''
          };
          ok = true;
        }
      } else {
        errs.push('NaverPolling HTTP ' + r3.getResponseCode());
      }
    } catch(e3) {
      errs.push('NaverPolling ' + String(e3 && e3.message ? e3.message : e3));
    }
    if (ok) return;

    // D. 네이버 종목 메인 HTML 백업
    try {
      var r4 = UrlFetchApp.fetch(
        'https://finance.naver.com/item/main.naver?code=' + encodeURIComponent(code),
        {
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
            'Accept-Language':'ko-KR,ko;q=0.9,en-US;q=0.8',
            'Referer':'https://finance.naver.com/'
          }
        }
      );
      if (r4.getResponseCode() >= 200 && r4.getResponseCode() < 300) {
        var txt4 = r4.getContentText('UTF-8');
        var m4 = txt4.match(/no_today[\s\S]{0,1800}?<span class="blind">([\d,]+)<\/span>/i);
        if (!m4) m4 = txt4.match(/<dd>현재가\s*([\d,]+)\s*/i);
        var p4 = m4 && Number(String(m4[1] || '').replace(/,/g,''));
        if (p4 > 0) {
          out.prices[responseKey] = {
            price:p4,
            currency:'KRW',
            source:'Naver HTML'
          };
          ok = true;
        } else {
          errs.push('NaverHTML price 없음');
        }
      } else {
        errs.push('NaverHTML HTTP ' + r4.getResponseCode());
      }
    } catch(e4) {
      errs.push('NaverHTML ' + String(e4 && e4.message ? e4.message : e4));
    }
    if (ok) return;

    // E. 다음 금융 백업
    try {
      var symbol = 'A' + code;
      var r5 = UrlFetchApp.fetch(
        'https://finance.daum.net/api/quotes/' + encodeURIComponent(symbol),
        {
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
            'Accept':'application/json,text/plain,*/*',
            'Referer':'https://finance.daum.net/quotes/' + symbol
          }
        }
      );
      if (r5.getResponseCode() >= 200 && r5.getResponseCode() < 300) {
        var j5 = JSON.parse(r5.getContentText());
        var p5 = Number(
          j5.tradePrice || j5.currentPrice || j5.closePrice ||
          j5.regularMarketPrice || j5.price || 0
        );
        if (p5 > 0) {
          out.prices[responseKey] = {
            price:p5,
            currency:'KRW',
            name:j5.name || j5.symbolName || '',
            source:'Daum Finance',
            previousClose:deriveKrPreviousClose_(p5,j5),
            marketStatus:j5.marketStatus || ''
          };
          ok = true;
        } else {
          errs.push('Daum price 없음');
        }
      } else {
        errs.push('Daum HTTP ' + r5.getResponseCode());
      }
    } catch(e5) {
      errs.push('Daum ' + String(e5 && e5.message ? e5.message : e5));
    }

    // F. Google Finance 최종 백업
    // 예: https://www.google.com/finance/quote/000660:KRX
    try {
      var r6 = UrlFetchApp.fetch(
        'https://www.google.com/finance/quote/' + encodeURIComponent(code) + ':KRX?hl=ko&gl=KR',
        {
          method:'get',
          muteHttpExceptions:true,
          followRedirects:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
            'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language':'ko-KR,ko;q=0.9,en-US;q=0.8'
          }
        }
      );

      if (r6.getResponseCode() >= 200 && r6.getResponseCode() < 300) {
        var txt6 = r6.getContentText('UTF-8');
        var p6 = null;

        // Google Finance 페이지에서 자주 쓰이는 값들을 순서대로 시도
        var m61 = txt6.match(/data-last-price="([0-9.,]+)"/i);
        if (m61) p6 = Number(String(m61[1]).replace(/,/g,''));

        if (!(p6 > 0)) {
          var m62 = txt6.match(/class="YMlKec fxKbKc"[^>]*>\s*(?:₩|&#8361;)?\s*([0-9,]+(?:\.[0-9]+)?)/i);
          if (m62) p6 = Number(String(m62[1]).replace(/,/g,''));
        }

        if (!(p6 > 0)) {
          var m63 = txt6.match(/"price"\s*:\s*"?([0-9]+(?:\.[0-9]+)?)"?/i);
          if (m63) p6 = Number(m63[1]);
        }

        if (p6 > 0) {
          out.prices[responseKey] = {
            price:p6,
            currency:'KRW',
            source:'Google Finance'
          };
          ok = true;
        } else {
          errs.push('GoogleFinance price 없음');
        }
      } else {
        errs.push('GoogleFinance HTTP ' + r6.getResponseCode());
      }
    } catch(e6) {
      errs.push('GoogleFinance ' + String(e6 && e6.message ? e6.message : e6));
    }

    if (!ok) {
      out.errors[responseKey] =
        '국내시세 조회 실패 (' + errs.slice(-5).join(' / ') + ')';
    }
  });

  // 전일 종가는 앱의 자산변화 기록과 무관하게 시세 사이트에서 직접 가져온다.
  // 현재가 소스에 전일대비가 없었던 종목만 Naver 일봉 API를 fetchAll로 한 번에 보완.
  var prevTargets = [];
  codes.forEach(function(rawCode){
    var requestCode = String(rawCode == null ? '' : rawCode).trim();
    var key = 'KR:' + requestCode;
    var q = out.prices[key];
    if (q && Number(q.price) > 0 && !(Number(q.previousClose) > 0)) {
      prevTargets.push({key:key, code:normalizeKrCode_(requestCode)});
    }
  });

  if (prevTargets.length) {
    try {
      var prevReqs = prevTargets.map(function(x){
        return {
          url:'https://m.stock.naver.com/api/stock/' + encodeURIComponent(x.code) + '/price?pageSize=3&page=1',
          method:'get',
          muteHttpExceptions:true,
          headers:{
            'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/124 Mobile Safari/537.36',
            'Accept':'application/json,text/plain,*/*',
            'Referer':'https://m.stock.naver.com/'
          }
        };
      });
      var prevRes = UrlFetchApp.fetchAll(prevReqs);

      prevRes.forEach(function(r, i){
        if (r.getResponseCode() < 200 || r.getResponseCode() >= 300) return;
        try {
          var j = JSON.parse(r.getContentText());
          var rows = Array.isArray(j) ? j : (Array.isArray(j.result) ? j.result : (Array.isArray(j.prices) ? j.prices : []));
          var valid = rows.filter(function(x){
            return Number(String((x && (x.closePrice || x.tradePrice || x.currentPrice)) || '').replace(/,/g,'')) > 0;
          });
          if (valid.length >= 2) {
            var prevRow = valid[1];
            var prev = Number(String(prevRow.closePrice || prevRow.tradePrice || prevRow.currentPrice || '').replace(/,/g,''));
            if (prev > 0 && out.prices[prevTargets[i].key]) {
              out.prices[prevTargets[i].key].previousClose = prev;
              out.prices[prevTargets[i].key].previousCloseDate = prevRow.localTradedAt || prevRow.date || '';
            }
          }
        } catch(_e) {}
      });
    } catch(_batchErr) {}
  }

}
function fetchCrypto_(symbols, out) {
  if (symbols.indexOf('BTC') === -1) return;

  // 1순위: Upbit - 현재가와 전일 기준 종가를 함께 제공
  try {
    var r1 = UrlFetchApp.fetch('https://api.upbit.com/v1/ticker?markets=KRW-BTC', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/15.0)'}
    });
    if (r1.getResponseCode() >= 200 && r1.getResponseCode() < 300) {
      var j1 = JSON.parse(r1.getContentText());
      var d1 = j1 && j1[0];
      var price1 = d1 && Number(d1.trade_price);
      var prev1 = d1 && Number(d1.prev_closing_price);
      if (price1 > 0) {
        out.prices['CRYPTO:BTC'] = {
          price:price1,
          previousClose:prev1>0?prev1:null,
          currency:'KRW',
          source:'Upbit'
        };
        return;
      }
    }
  } catch(e1) {}

  // 2순위: CoinGecko (전일 종가가 없어 현재가만 백업)
  try {
    var r2 = UrlFetchApp.fetch('https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=krw', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/15.0)'}
    });
    if (r2.getResponseCode() >= 200 && r2.getResponseCode() < 300) {
      var j2 = JSON.parse(r2.getContentText());
      var price2 = j2.bitcoin && Number(j2.bitcoin.krw);
      if (price2 > 0) {
        out.prices['CRYPTO:BTC'] = {price:price2,currency:'KRW',source:'CoinGecko'};
        return;
      }
    }
  } catch(e2) {}

  out.errors['CRYPTO:BTC'] = 'BTC 현재가 조회 실패';
}

function fetchFx_(out) {
  // FX는 주식처럼 하나의 '정규장'이 있는 시장이 아니라 평일 거의 24시간 거래됩니다.
  // 앱에서는 Yahoo Finance의 USD/KRW(KRW=X) 1분 차트에서 가장 최근 유효값을 우선 사용합니다.
  function yahooFx_(host) {
    var url = 'https://' + host + '/v8/finance/chart/' + encodeURIComponent('KRW=X') +
              '?range=1d&interval=1m&includePrePost=true';
    var r = UrlFetchApp.fetch(url, {
      muteHttpExceptions:true,
      headers:{
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36',
        'Accept':'application/json,text/plain,*/*',
        'Referer':'https://finance.yahoo.com/'
      }
    });
    if (r.getResponseCode() < 200 || r.getResponseCode() >= 300) {
      throw new Error('HTTP ' + r.getResponseCode());
    }

    var j = JSON.parse(r.getContentText());
    var result = j && j.chart && j.chart.result && j.chart.result[0];
    if (!result) throw new Error('Yahoo FX 응답 없음');

    var meta = result.meta || {};
    var ts = result.timestamp || [];
    var quote = result.indicators && result.indicators.quote &&
                result.indicators.quote[0] || {};
    var closes = quote.close || [];

    var latest = null, latestTs = null;
    for (var i = closes.length - 1; i >= 0; i--) {
      var p = Number(closes[i]);
      if (p > 0) {
        latest = p;
        latestTs = Number(ts[i] || 0) || null;
        break;
      }
    }

    if (!(latest > 0)) {
      latest = Number(meta.regularMarketPrice) ||
               Number(meta.postMarketPrice) ||
               Number(meta.preMarketPrice);
      latestTs = Number(meta.regularMarketTime || 0) || null;
    }

    if (!(latest > 0)) throw new Error('유효한 USD/KRW 가격 없음');

    return {
      price: latest,
      timestamp: latestTs ? new Date(latestTs * 1000).toISOString() : new Date().toISOString()
    };
  }

  // 1순위: Yahoo Finance query1, 1분 시세
  try {
    var y1 = yahooFx_('query1.finance.yahoo.com');
    out.fx = y1.price;
    out.fxSource = 'Yahoo FX 1분';
    out.fxTimestamp = y1.timestamp;
    return;
  } catch(e0) {}

  // 2순위: Yahoo Finance query2 재시도
  try {
    var y2 = yahooFx_('query2.finance.yahoo.com');
    out.fx = y2.price;
    out.fxSource = 'Yahoo FX 1분';
    out.fxTimestamp = y2.timestamp;
    return;
  } catch(e1) {}

  // 3순위: 네이버 USD/KRW 시장지수
  try {
    var r0 = UrlFetchApp.fetch(
      'https://m.stock.naver.com/front-api/marketIndex/prices?category=exchange&reutersCode=FX_USDKRW&pageSize=1&page=1',
      {
        muteHttpExceptions:true,
        headers:{
          'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/120 Mobile Safari/537.36',
          'Referer':'https://m.stock.naver.com/'
        }
      }
    );
    if (r0.getResponseCode() >= 200 && r0.getResponseCode() < 300) {
      var j0 = JSON.parse(r0.getContentText());
      var arr0 = j0 && (j0.result || j0);
      var d0 = Array.isArray(arr0) ? arr0[0] : null;
      var fx0 = d0 && Number(String(d0.closePrice || '').replace(/,/g,''));
      if (fx0 > 0) {
        out.fx = fx0;
        out.fxSource = 'Naver FX 백업';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e2) {}

  // 4순위: Frankfurter - 일일 기준 백업
  try {
    var r1 = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=USD&to=KRW', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/15.0)'}
    });
    if (r1.getResponseCode() >= 200 && r1.getResponseCode() < 300) {
      var j1 = JSON.parse(r1.getContentText());
      var fx1 = j1.rates && Number(j1.rates.KRW);
      if (fx1 > 0) {
        out.fx = fx1;
        out.fxSource = 'Frankfurter 일일 백업';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e3) {}

  out.errors['FX:USDKRW'] = 'Yahoo/Naver/Frankfurter 환율 조회 실패';
}

