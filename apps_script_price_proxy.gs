
/**
 * Twelve Data 설정
 * 1) https://twelvedata.com 에서 무료 API Key 발급
 * 2) 아래 TWELVE_DATA_API_KEY 값만 교체
 *
 * 무료 Basic: 미국 주식/ETF 정규장 실시간(또는 최신 이용 가능 가격)
 * 프리/애프터장: Twelve Data Pro 이상에서 TWELVE_DATA_USE_EXTENDED_HOURS = true
 */
var TWELVE_DATA_API_KEY = '여기에_TWELVE_DATA_API_KEY_입력';
var TWELVE_DATA_USE_EXTENDED_HOURS = false;

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
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/6.0)'}
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
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/6.0)'}
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
      headers:{'User-Agent':'Mozilla/5.0 (compatible; PortfolioPriceProxy/6.0)'}
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

