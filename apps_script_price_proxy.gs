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

function handleSyncAction_(action,p){
  var key=String(p.key||'').trim();
  if(!key||key.length<6)throw new Error('동기화 코드가 올바르지 않습니다.');
  var ss=getSyncSpreadsheet_(),ash=ss.getSheetByName(SYNC_ASSET_SHEET),msh=ss.getSheetByName(SYNC_META_SHEET);
  if(action==='syncInit')return{ok:true,updatedAt:new Date().toISOString(),assetCount:countAssetsForKey_(ash,key),spreadsheetUrl:ss.getUrl()};
  var lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    if(action==='syncGet'){
      var assets=readAssetsForKey_(ash,key),meta=readMetaForKey_(msh,key);
      return{ok:true,assets:assets,accounts:meta.accounts,updatedAt:meta.updatedAt||latestAssetUpdatedAt_(ash,key),spreadsheetUrl:ss.getUrl()};
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
      upsertMeta_(msh,key,acc);return{ok:true,updatedAt:new Date().toISOString()}
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
  return ss;
}
function ensureSheet_(ss,name,headers){
  var sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1)}
  return sh;
}
function countAssetsForKey_(sh,key){var last=sh.getLastRow();if(last<2)return 0;var v=sh.getRange(2,1,last-1,1).getValues(),n=0;v.forEach(function(r){if(String(r[0])===key)n++});return n}
function assetToRow_(key,a){return[key,Number(a.id||0),String(a.account||''),String(a.name||''),Number(a.qty||0),Number(a.avg||0),a.purchaseFx==null||a.purchaseFx===''?'':Number(a.purchaseFx),Number(a.cost||0),Number(a.price||0),String(a.currency||'KRW'),String(a.market||'MANUAL'),String(a.ticker||''),a.targetWeight==null||a.targetWeight===''?'':Number(a.targetWeight),String(a.memo||''),new Date().toISOString()]}
function rowToAsset_(r){return{id:Number(r[1]||0),account:String(r[2]||''),name:String(r[3]||''),qty:Number(r[4]||0),avg:Number(r[5]||0),purchaseFx:r[6]===''?null:Number(r[6]),cost:Number(r[7]||0),price:Number(r[8]||0),currency:String(r[9]||'KRW'),market:String(r[10]||'MANUAL'),ticker:String(r[11]||''),targetWeight:r[12]===''?'':Number(r[12]),memo:String(r[13]||'')}}
function findAssetRow_(sh,key,id){var last=sh.getLastRow();if(last<2)return-1;var v=sh.getRange(2,1,last-1,2).getValues();for(var i=0;i<v.length;i++)if(String(v[i][0])===key&&String(v[i][1])===String(id))return i+2;return-1}
function upsertAsset_(sh,key,a){if(!a||!a.id)throw new Error('자산 ID가 없습니다.');var row=assetToRow_(key,a),n=findAssetRow_(sh,key,a.id);if(n>0)sh.getRange(n,1,1,row.length).setValues([row]);else sh.appendRow(row)}
function deleteAsset_(sh,key,id){var n=findAssetRow_(sh,key,id);if(n>0)sh.deleteRow(n)}
function clearKeyAssets_(sh,key){var last=sh.getLastRow();if(last<2)return;var v=sh.getRange(2,1,last-1,1).getValues();for(var i=v.length-1;i>=0;i--)if(String(v[i][0])===key)sh.deleteRow(i+2)}
function readAssetsForKey_(sh,key){var last=sh.getLastRow();if(last<2)return[];var v=sh.getRange(2,1,last-1,15).getValues(),o=[];v.forEach(function(r){if(String(r[0])===key)o.push(rowToAsset_(r))});return o}
function latestAssetUpdatedAt_(sh,key){var last=sh.getLastRow();if(last<2)return'';var v=sh.getRange(2,1,last-1,15).getValues(),x='';v.forEach(function(r){if(String(r[0])===key&&String(r[14]||'')>x)x=String(r[14]||'')});return x}
function findMetaRow_(sh,key){var last=sh.getLastRow();if(last<2)return-1;var v=sh.getRange(2,1,last-1,1).getValues();for(var i=0;i<v.length;i++)if(String(v[i][0])===key)return i+2;return-1}
function upsertMeta_(sh,key,accounts){var row=[key,JSON.stringify(accounts||[]),new Date().toISOString()],n=findMetaRow_(sh,key);if(n>0)sh.getRange(n,1,1,3).setValues([row]);else sh.appendRow(row)}
function readMetaForKey_(sh,key){var n=findMetaRow_(sh,key);if(n<0)return{accounts:[],updatedAt:''};var r=sh.getRange(n,1,1,3).getValues()[0],a=[];try{a=JSON.parse(String(r[1]||'[]'))}catch(e){}if(!Array.isArray(a))a=[];return{accounts:a,updatedAt:String(r[2]||'')}}

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

    return {
      price: latestPrice,
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
}
function fetchKR_(codes, out) {
  if (!codes.length) return;

  codes.forEach(function(code){
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
          out.prices['KR:' + code] = {
            price:p1,
            currency:'KRW',
            name:j1.stockName || '',
            source:'Naver Basic',
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
          out.prices['KR:' + code] = {
            price:p2,
            currency:'KRW',
            name:d2.stockName || d2.name || '',
            source:'Naver Realtime',
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
          out.prices['KR:' + code] = {
            price:p3,
            currency:'KRW',
            source:'Naver Polling',
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
          out.prices['KR:' + code] = {
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
          out.prices['KR:' + code] = {
            price:p5,
            currency:'KRW',
            name:j5.name || j5.symbolName || '',
            source:'Daum Finance',
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

    if (!ok) {
      out.errors['KR:' + code] =
        '국내시세 조회 실패 (' + errs.slice(-4).join(' / ') + ')';
    }
  });
}
function fetchCrypto_(symbols, out) {
  if (symbols.indexOf('BTC') === -1) return;

  // 1순위: CoinGecko
  try {
    var r = UrlFetchApp.fetch('https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=krw', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/9.0)'}
    });
    if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) {
      var j = JSON.parse(r.getContentText());
      var price = j.bitcoin && Number(j.bitcoin.krw);
      if (price > 0) {
        out.prices['CRYPTO:BTC'] = {price:price,currency:'KRW',source:'CoinGecko'};
        return;
      }
    }
  } catch(e1) {}

  // 2순위: Upbit KRW-BTC
  try {
    var r2 = UrlFetchApp.fetch('https://api.upbit.com/v1/ticker?markets=KRW-BTC', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/9.0)'}
    });
    if (r2.getResponseCode() >= 200 && r2.getResponseCode() < 300) {
      var j2 = JSON.parse(r2.getContentText());
      var price2 = j2 && j2[0] && Number(j2[0].trade_price);
      if (price2 > 0) {
        out.prices['CRYPTO:BTC'] = {price:price2,currency:'KRW',source:'Upbit'};
        return;
      }
    }
  } catch(e2) {}

  out.errors['CRYPTO:BTC'] = 'CoinGecko/Upbit 모두 조회 실패';
}

function fetchFx_(out) {
  // 1순위: 네이버 USD/KRW 시장지수
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
        out.fxSource = 'Naver Finance';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e0) {}

  // 2순위: Frankfurter(일일 환율)
  try {
    var r1 = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=USD&to=KRW', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/9.0)'}
    });
    if (r1.getResponseCode() >= 200 && r1.getResponseCode() < 300) {
      var j1 = JSON.parse(r1.getContentText());
      var fx1 = j1.rates && Number(j1.rates.KRW);
      if (fx1 > 0) {
        out.fx = fx1;
        out.fxSource = 'Frankfurter';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e1) {}

  out.errors['FX:USDKRW'] = 'Naver/Frankfurter 모두 조회 실패';
}

