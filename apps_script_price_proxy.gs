/**
 * 자산관리 HTML용 시세 프록시 (Google Apps Script)
 * 배포: 배포 > 새 배포 > 웹 앱 > 실행 사용자: 나 > 액세스 권한: 모든 사용자
 * 배포 후 /exec URL을 자산관리 HTML의 설정에 입력하세요.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out = { ok: true, timestamp: new Date().toISOString(), prices: {}, fx: null, fxSource: null, fxTimestamp: null, errors: {} };

  try {
    var us = splitList_(p.us);
    var kr = splitList_(p.kr);
    var crypto = splitList_(p.crypto);

    fetchUS_(us, out);
    fetchKR_(kr, out);
    fetchCrypto_(crypto, out);
    if (p.fx === '1') fetchFx_(out);
  } catch (err) {
    out.ok = false;
    out.message = String(err && err.message ? err.message : err);
  }

  var body = JSON.stringify(out);
  var cb = String(p.callback || '');
  if (/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function splitList_(s) {
  if (!s) return [];
  var seen = {};
  return String(s).split(',').map(function(x){ return x.trim().toUpperCase(); })
    .filter(function(x){ if (!x || seen[x]) return false; seen[x] = true; return true; });
}

function fetchUS_(symbols, out) {
  if (!symbols.length) return;
  var reqs = symbols.map(function(sym){
    return {
      url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(sym) + '?range=1d&interval=1d',
      method: 'get', muteHttpExceptions: true,
      headers: {'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/3.0)'}
    };
  });
  var res = UrlFetchApp.fetchAll(reqs);
  res.forEach(function(r,i){
    var sym = symbols[i];
    try {
      if (r.getResponseCode() < 200 || r.getResponseCode() >= 300) throw new Error('HTTP ' + r.getResponseCode());
      var j = JSON.parse(r.getContentText());
      var meta = j.chart && j.chart.result && j.chart.result[0] && j.chart.result[0].meta;
      var price = meta && Number(meta.regularMarketPrice);
      if (!(price > 0)) throw new Error('price 없음');
      out.prices['US:' + sym] = {price: price, currency: meta.currency || 'USD', source:'Yahoo Finance'};
    } catch(err) { out.errors['US:' + sym] = String(err); }
  });
}

function fetchKR_(codes, out) {
  if (!codes.length) return;

  // 1차: 네이버 모바일 basic API (가장 가볍고 빠름)
  var reqs = codes.map(function(code){
    return {
      url: 'https://m.stock.naver.com/api/stock/' + encodeURIComponent(code) + '/basic',
      method:'get', muteHttpExceptions:true,
      headers:{
        'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/120 Mobile Safari/537.36',
        'Referer':'https://m.stock.naver.com/'
      }
    };
  });

  var res = UrlFetchApp.fetchAll(reqs);

  res.forEach(function(r,i){
    var code = codes[i];
    var ok = false;

    // 방법 A: /api/stock/{code}/basic
    try {
      if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) {
        var j = JSON.parse(r.getContentText());
        var price = Number(String(j.closePrice || '').replace(/,/g,''));
        if (price > 0) {
          out.prices['KR:' + code] = {
            price: price,
            currency:'KRW',
            name:j.stockName || '',
            source:'Naver Basic',
            marketStatus:j.marketStatus || ''
          };
          ok = true;
        }
      }
    } catch(e1) {}

    if (ok) return;

    // 방법 B: 신규/영문 혼합 ETF 코드 대응용 실시간 domestic endpoint
    try {
      var url2 = 'https://polling.finance.naver.com/api/realtime/domestic/stock/' + encodeURIComponent(code);
      var r2 = UrlFetchApp.fetch(url2, {
        method:'get',
        muteHttpExceptions:true,
        headers:{
          'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/120 Mobile Safari/537.36',
          'Referer':'https://finance.naver.com/'
        }
      });

      if (r2.getResponseCode() >= 200 && r2.getResponseCode() < 300) {
        var j2 = JSON.parse(r2.getContentText());
        var d2 = j2 && j2.datas && j2.datas[0];
        var price2 = d2 && Number(String(d2.closePrice || '').replace(/,/g,''));
        if (price2 > 0) {
          out.prices['KR:' + code] = {
            price:price2,
            currency:'KRW',
            name:d2.stockName || d2.name || '',
            source:'Naver Realtime',
            marketStatus:d2.marketStatus || ''
          };
          ok = true;
        }
      }
    } catch(e2) {}

    if (ok) return;

    // 방법 C: 구형 polling query endpoint
    try {
      var url3 = 'https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:' + encodeURIComponent(code);
      var r3 = UrlFetchApp.fetch(url3, {
        method:'get',
        muteHttpExceptions:true,
        headers:{
          'User-Agent':'Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/120 Mobile Safari/537.36',
          'Referer':'https://finance.naver.com/'
        }
      });

      if (r3.getResponseCode() >= 200 && r3.getResponseCode() < 300) {
        var j3 = JSON.parse(r3.getContentText());
        var d3 = j3 && j3.result && j3.result.areas && j3.result.areas[0] &&
                 j3.result.areas[0].datas && j3.result.areas[0].datas[0];
        var price3 = d3 && Number(d3.nv);
        if (price3 > 0) {
          out.prices['KR:' + code] = {
            price:price3,
            currency:'KRW',
            source:'Naver Polling',
            marketStatus:d3.ms || ''
          };
          ok = true;
        }
      }
    } catch(e3) {}

    if (!ok) {
      out.errors['KR:' + code] = '네이버 국내시세 3개 경로 모두 조회 실패';
    }
  });
}

function fetchCrypto_(symbols, out) {
  if (symbols.indexOf('BTC') === -1) return;

  // 1순위: CoinGecko
  try {
    var r = UrlFetchApp.fetch('https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=krw', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/3.0)'}
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
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/3.0)'}
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

  // 2순위: Yahoo Finance
  try {
    var r = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v8/finance/chart/KRW=X?range=1d&interval=1d', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/3.0)'}
    });
    if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) {
      var j = JSON.parse(r.getContentText());
      var meta = j.chart && j.chart.result && j.chart.result[0] && j.chart.result[0].meta;
      var fx = meta && Number(meta.regularMarketPrice);
      if (fx > 0) {
        out.fx = fx;
        out.fxSource = 'Yahoo Finance';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e1) {}

  // 3순위: Frankfurter(일일 환율)
  try {
    var r2 = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=USD&to=KRW', {
      muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/3.0)'}
    });
    if (r2.getResponseCode() >= 200 && r2.getResponseCode() < 300) {
      var j2 = JSON.parse(r2.getContentText());
      var fx2 = j2.rates && Number(j2.rates.KRW);
      if (fx2 > 0) {
        out.fx = fx2;
        out.fxSource = 'Frankfurter';
        out.fxTimestamp = new Date().toISOString();
        return;
      }
    }
  } catch(e2) {}

  out.errors['FX:USDKRW'] = 'Naver/Yahoo/Frankfurter 모두 조회 실패';
}

