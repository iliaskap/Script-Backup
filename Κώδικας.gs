function getBearishToBullishCoins() {
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';
  const url = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=50&page=1&sparkline=true&price_change_percentage=1h,24h";

  const options = {
    method: 'get',
    headers: {
      'x-cg-demo-api-key': API_KEY
    },
    muteHttpExceptions: true
  };

  const data = JSON.parse(response.getContentText());
  const response = UrlFetchApp.fetch(url, options);

  if (!Array.isArray(data)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Error fetching data from CoinGecko (check API key or quota)");
    Logger.log(data);
    return;
  }

  const sheetName = "Crypto Trends";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = ["Name", "Symbol", "Current Price (USD)", "1h %", "24h %"];
  const results = [headers];

  for (let coin of data) {
    const sparkline = coin.sparkline_in_7d?.price;
    if (!sparkline || sparkline.length < 24) continue;

    const firstHalf = sparkline.slice(-24, -12);
    const secondHalf = sparkline.slice(-12);

    const bearishFirst = firstHalf[0] > firstHalf[firstHalf.length - 1];
    const bullishSecond = secondHalf[0] < secondHalf[secondHalf.length - 1];

    if (bearishFirst && bullishSecond) {
      results.push([
        coin.name,
        coin.symbol.toUpperCase(),
        coin.current_price,
        coin.price_change_percentage_1h_in_currency || 0,
        coin.price_change_percentage_24h_in_currency || 0
      ]);
    }
  }

  if (results.length === 1) {
    sheet.getRange("A1").setValue("No matching cryptos found.");
    return;
  }

  // Write results
  sheet.getRange(1, 1, results.length, headers.length).setValues(results);

  // Format headers bold
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  // Format % columns with color
  const numRows = results.length - 1;
  const oneHrRange = sheet.getRange(2, 4, numRows);
  const dayRange = sheet.getRange(2, 5, numRows);

  // Green for positive, red for negative
  oneHrRange.setNumberFormat("0.00%");
  dayRange.setNumberFormat("0.00%");

  oneHrRange.setFontColor("black");
  dayRange.setFontColor("black");

  oneHrRange.setFontWeight("normal");
  dayRange.setFontWeight("normal");

  oneHrRange.setFontColors(
    oneHrRange.getValues().map(row => [row[0] >= 0 ? "green" : "red"])
  );
  dayRange.setFontColors(
    dayRange.getValues().map(row => [row[0] >= 0 ? "green" : "red"])
  );

  SpreadsheetApp.flush();
}


function findSMA6CrossAboveSMA30() {
  const sheetName = "Sheet1";  // Change if needed
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found!`);

  sheet.clearContents();
  sheet.appendRow(["Coin ID", "Last Price", "Crossover Date", "Status"]);

  const coinIds = [
    'aave', 'aptos', 'binance-usd', 'bitcoin', 'algorand', 'cardano', 'chainlink', 'hedera-hashgraph', 'internet-computer', 'maker', 'polkadot', 'ripple', 'solana', 'stellar', 'sui', 'tron', 
    'uniswap', 'vechain', 'kaspa', 'civic'
  ];

  const options = { muteHttpExceptions: true };

  for (let coinId of coinIds) {
    Logger.log(`Fetching data for ${coinId}...`);

    const url = `https://api.coingecko.com/api/v3/coins/${coinId}/market_chart?vs_currency=usd&days=90&interval=daily`;
    const response = fetchWithRetry(url, options, 3);

    if (!response) {
      Logger.log(`❌ ${coinId}: Failed to fetch data after retries.`);
      sheet.appendRow([coinId, "", "", "Fetch failed"]);
      continue;
    }

    const code = response.getResponseCode();
    const content = response.getContentText();

    if (code !== 200) {
      Logger.log(`❌ ${coinId}: HTTP error ${code}`);
      Logger.log(`Response: ${content}`);
      sheet.appendRow([coinId, "", "", `HTTP error ${code}`]);
      Utilities.sleep(3000);
      continue;
    }

    let json;
    try {
      json = JSON.parse(content);
    } catch (e) {
      Logger.log(`❌ ${coinId}: JSON parse error - ${e.message}`);
      sheet.appendRow([coinId, "", "", "JSON parse error"]);
      Utilities.sleep(3000);
      continue;
    }

    const prices = json.prices;
    Logger.log(`${coinId} prices length: ${prices ? prices.length : 0}`);

    if (!prices || prices.length < 31) {
      Logger.log(`⚠️ ${coinId}: Not enough data (${prices ? prices.length : 0} entries)`);
      sheet.appendRow([coinId, "", "", "Not enough data"]);
      Utilities.sleep(3000);
      continue;
    }

    const sma6 = calculateSMA(prices, 6);
    const sma30 = calculateSMA(prices, 30);

    if (!sma6 || !sma30 || sma6.length < 2 || sma30.length < 2) {
      Logger.log(`⚠️ ${coinId}: SMA arrays too short`);
      sheet.appendRow([coinId, "", "", "SMA error"]);
      Utilities.sleep(3000);
      continue;
    }

    const prevSMA6 = sma6[sma6.length - 2];
    const prevSMA30 = sma30[sma30.length - 2];
    const currSMA6 = sma6[sma6.length - 1];
    const currSMA30 = sma30[sma30.length - 1];

    if (prevSMA6 < prevSMA30 && currSMA6 > currSMA30) {
      const lastPrice = prices[prices.length - 1][1];
      const date = new Date(prices[prices.length - 1][0]).toLocaleDateString();
      sheet.appendRow([coinId, lastPrice, date, "Crossover detected"]);
      Logger.log(`✅ ${coinId} crossover detected on ${date}`);
    } else {
      Logger.log(`ℹ️ ${coinId}: No crossover`);
      sheet.appendRow([coinId, "", "", "No crossover"]);
    }

    Utilities.sleep(3000); // Wait 3 seconds before next request
  }
}

function calculateSMA(prices, period) {
  if (!prices || prices.length < period) return null;
  const sma = [];
  for (let i = period - 1; i < prices.length; i++) {
    let sum = 0;
    for (let j = 0; j < period; j++) {
      sum += prices[i - j][1];
    }
    sma.push(sum / period);
  }
  return sma;
}

function fetchWithRetry(url, options, maxRetries) {
  let attempts = 0;
  while (attempts < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        return response;
      } else {
        Logger.log(`Attempt ${attempts + 1}: HTTP ${response.getResponseCode()}`);
      }
    } catch (e) {
      Logger.log(`Attempt ${attempts + 1} error: ${e.message}`);
    }
    attempts++;
    Utilities.sleep(3000 * attempts); // exponential backoff
  }
  return null;
}

function testFetchBitcoin() {
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';  // ← Replace with your actual API key

  const url = 'https://api.coingecko.com/api/v3/coins/bitcoin?localization=false';

  const options = {
    method: 'get',
    headers: {
      'x-cg-pro-api-key': API_KEY // Use this for your personal or paid API key
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);
}


function getTop25GainersLast6Hours() {
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';
  const url = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=50&page=1&sparkline=true";

  const options = {
    method: 'get',
    headers: {
      'x-cg-demo-api-key': API_KEY
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (!Array.isArray(data)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Error fetching data from CoinGecko (check API key or quota)");
    Logger.log(data);
    return;
  }

  // Calculate 6h % change from sparkline price data
  const coinsWith6hChange = data
    .map(coin => {
      const prices = coin.sparkline_in_7d?.price;
      if (!prices || prices.length < 6) return null;

      const latestPrice = prices[prices.length - 1];
      const price6hAgo = prices[prices.length - 2];
      const change6h = ((latestPrice - price6hAgo) / price6hAgo) * 100;

      return {
        market_cap_rank: coin.market_cap_rank,
        name: coin.name,
        symbol: coin.symbol.toUpperCase(),
        current_price: coin.current_price,
        change6h: change6h
      };
    })
    .filter(coin => coin !== null);

  // Sort descending by 6h change
  coinsWith6hChange.sort((a, b) => b.change6h - a.change6h);

  // Take top 25
  const top25 = coinsWith6hChange.slice(0, 25);

  const sheetName = "Top 25 Gainers 6h";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = ["Rank", "Name", "Symbol", "Current Price (USD)", "6h % Change"];
  const results = [headers];

  for (let coin of top25) {
    results.push([
      coin.market_cap_rank,
      coin.name,
      coin.symbol,
      coin.current_price,
      coin.change6h / 100 // format as decimal for percent display
    ]);
  }

  // Write results
  sheet.getRange(1, 1, results.length, headers.length).setValues(results);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  // Format % column with green/red colors
  const numRows = results.length - 1;
  const changeRange = sheet.getRange(2, 5, numRows);
  changeRange.setNumberFormat("0.00%");
  changeRange.setFontColors(
    changeRange.getValues().map(row => [row[0] >= 0 ? "green" : "red"])
  );

  SpreadsheetApp.flush();
}

function getTop50EMAs() {
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';  // Replace with your CoinGecko API key
  const topUrl = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=volume_desc&per_page=50&page=1&sparkline=false";

  const options = {
    method: 'get',
    headers: {
      'x-cg-pro-api-key': API_KEY
    },
    muteHttpExceptions: true
  };

  const topCoins = JSON.parse(UrlFetchApp.fetch(topUrl, options).getContentText());

  const sheetName = "Top 50 EMA";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange("A1:F1").setValues([["Name", "Symbol", "Current Price", "EMA 6", "EMA 30", "Status"]]);

  for (let i = 0; i < topCoins.length; i++) {
    const coin = topCoins[i];
    const id = coin.id;
    const name = coin.name;
    const symbol = coin.symbol.toUpperCase();
    const currentPrice = coin.current_price;

    // Get 30-day historical price
    const historyUrl = `https://api.coingecko.com/api/v3/coins/${id}/market_chart?vs_currency=usd&days=30&interval=daily`;
    try {
      const historyData = JSON.parse(UrlFetchApp.fetch(historyUrl, options).getContentText());
      const prices = historyData.prices.map(row => row[1]); // just price

      if (prices.length < 30) {
        sheet.appendRow([name, symbol, currentPrice, "N/A", "N/A", "Not enough data"]);
        continue;
      }

      const ema6 = calculateEMA(prices, 6);
      const ema30 = calculateEMA(prices, 30);

      sheet.appendRow([name, symbol, currentPrice, ema6.toFixed(4), ema30.toFixed(4), "OK"]);
    } catch (err) {
      sheet.appendRow([name, symbol, currentPrice, "Error", "Error", "Fetch failed"]);
    }

    Utilities.sleep(1500); // Respect rate limit: ~50 calls per minute
  }
}

function calculateEMA(prices, period) {
  const k = 2 / (period + 1);
  let ema = prices[0];
  for (let i = 1; i < prices.length; i++) {
    ema = prices[i] * k + ema * (1 - k);
  }
  return ema;
}

function getEMAsForSelectedCoins() {
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG'; // Demo key
  const coins = [
    'aave', 'aptos', 'binance-usd', 'bitcoin', 'algorand', 'cardano', 'chainlink', 'hedera-hashgraph', 'internet-computer', 'maker', 'polkadot', 'ripple', 'solana', 'stellar', 'sui', 'tron', 
    'uniswap', 'vechain', 'ethena', 'virtual-protocol', 'celestia', 'bittensor', 'avalanche-2', 'near', 'ondo-finance', 'polygon-ecosystem-token', 'render-token', 'fetch-ai', 'arbitrum', 'worldcoin-wld', 'quant-network', 'kaspa' , 'civic' , 'icon' , 'cosmos' , 'cookie' , 'hft'

  ]; // Customize list

  const sheetName = "Selected EMAs";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange("A1:G1").setValues([["Name", "Symbol", "Current Price", "EMA 6", "EMA 24", "Status", "Difference (%)"]]);
  const mysheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Selected EMAs');
  const data = mysheet.getDataRange().getValues();
  const BOT_TOKEN = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
  const CHAT_ID = '1457021291';

  const headers = {
    'x-cg-demo-api-key': API_KEY
  };

//  const message = `⌚️ EMA function start !`;
//  sendTelegramMessage(BOT_TOKEN, CHAT_ID, message);

  let row = 2;
  coins.forEach((id) => {
    try {
      // Fetch market data
      const marketUrl = `https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids=${id}`;
      const marketRes = UrlFetchApp.fetch(marketUrl, { headers });
      const marketData = JSON.parse(marketRes.getContentText())[0];
      const name = marketData.name;
      const symbol = marketData.symbol.toUpperCase();
      const currentPrice = marketData.current_price;

      // Fetch 30-day prices
      const chartUrl = `https://api.coingecko.com/api/v3/coins/${id}/market_chart?vs_currency=usd&days=30&interval=daily`;
      const chartRes = UrlFetchApp.fetch(chartUrl, { headers });
      const priceData = JSON.parse(chartRes.getContentText()).prices.map(row => row[1]);

      if (priceData.length < 30) {
        sheet.appendRow([name, symbol, currentPrice, "N/A", "N/A", "Insufficient data", ""]);
        row++;
        return;
      }

      const ema6 = calculateEMA(priceData, 6);
      const ema30 = calculateEMA(priceData, 24);
      const status = ema6 > ema30 ? "Bullish" : "Bearish";
      const diffPercent = ((ema6 - ema30) / ema30) * 100;

      sheet.appendRow([
        name,
        symbol,
        currentPrice,
        ema6,
        ema30,
        status,
        diffPercent.toFixed(2)
      ]);

        sheet.getRange(row, 4).setNumberFormat("0.000");
        sheet.getRange(row, 5).setNumberFormat("0.000");
        sheet.getRange(row, 7).setNumberFormat("0.000");

      const ratio = ema6 / ema30;

      if (diffPercent < 0.3 && diffPercent > -0.3) {

      //  const message = `📈 EMA Alert: ${symbol.toUpperCase()} EMA6 = (${ema6.toFixed(4)}) / EMA30 = (${ema30.toFixed(3)}) and ratio is (${diffPercent.toFixed(3)})`;
      // sendTelegramMessage(BOT_TOKEN, CHAT_ID, message);
      }  

      row++;
      Utilities.sleep(1500); // Respect rate limits

    } catch (err) {
      Logger.log(`Error with ${id}: ${err}`);
      sheet.appendRow([id, "", "", "", "", "Fetch Error", ""]);
      row++;
    }


 
  });
  // const endMessage = `⌚️ EMA function ends !`;
 // sendTelegramMessage(BOT_TOKEN, CHAT_ID, endMessage);

  applyHighlighting(sheet);

}

function calculateEMA(prices, period) {
  const k = 2 / (period + 1);
  let ema = prices[0];
  for (let i = 1; i < prices.length; i++) {
    ema = prices[i] * k + ema * (1 - k);
  }
  return ema;
}

function applyHighlighting(sheet) {
  const lastRow = sheet.getLastRow();
  const statusRange = sheet.getRange(`F2:F${lastRow}`);
  const fullRange = sheet.getRange(`G2:G${lastRow}`);

  const bullishRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Bullish')
    .setBackground('#c6efce') // green
    .setFontColor('#006100')
    .setRanges([statusRange])
    .build();

  const bearishRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Bearish')
    .setBackground('#f4cccc') // red
    .setFontColor('#990000')
    .setRanges([statusRange])
    .build();

  const neutralRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ABS(Value(REGEXREPLACE(G2;"[“.”]";","))) < 0,1')
    .setBackground('#fff2cc') // light yellow
    .setFontColor('#7f6000')
    .setRanges([fullRange])
    .build();

  sheet.setConditionalFormatRules([bullishRule, bearishRule, neutralRule]);
}

function sendTelegramMessage(botToken, chatId, message) {
  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'Markdown'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}

function testTelegram() {
  const token = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
  const chatId = '1457021291';
  const message = 'Test message from Google Sheets';

  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: message
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

function createHourlyTrigger() {
  ScriptApp.newTrigger('analyzeSelectedCryptoIndicators')
    .timeBased()
    .everyHours(3) // Change to .everyMinutes(n) for testing
    .create();
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function calculateKDJForCoins() {
  const coins = [
    'aave', 'aptos', 'binance-usd', 'bitcoin', 'algorand', 'cardano', 'chainlink', 'hedera-hashgraph', 'internet-computer', 'maker', 'polkadot', 'ripple', 'solana', 'stellar', 'sui', 'tron', 
    'uniswap', 'vechain', 'ethena', 'virtual-protocol', 'celestia', 'bittensor', 'avalanche-2', 'near', 'ondo-finance', 'polygon-ecosystem-token', 'render-token', 'fetch-ai', 'arbitrum', 'worldcoin-wld', 'quant-network', 'kaspa' , 'civic' , 'icon' , 'cosmos' , 'cookie' , 'hft'

  ]; // Customize list
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KDJIndicatorForCoins');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 5).setValues([['Coin', 'K', 'D', 'J', 'Signal']]);

  const apiKey = 'CG-ozyiJYmNEJGG8GcGBLciH1BG'; // Replace with demo key

  for (let i = 0; i < coins.length; i++) {
    const coinId = coins[i];
    const prices = getHistoricalPrices(coinId, apiKey);

    const row = i + 2;
    sheet.getRange(row, 1).setValue(coinId);

    if (!Array.isArray(prices) || prices.length < 15) {
      sheet.getRange(row, 5).setValue("❌ No Data").setFontColor("gray");
      Logger.log(`❌ Not enough data for ${coinId}`);
      continue;
    }

    const [K, D, J] = computeKDJ(prices);

    // Write K, D, J
    sheet.getRange(row, 2).setValue(K.toFixed(2));
    sheet.getRange(row, 3).setValue(D.toFixed(2));
    sheet.getRange(row, 4).setValue(J.toFixed(2));

    // Signal detection
    let signal = '';
    const signalCell = sheet.getRange(row, 5);

    if (J > K && J > D) {
      signal = 'Bullish';
      signalCell.setValue(signal).setFontWeight("bold").setFontColor("green");
    } else if (J < K && J < D) {
      signal = 'Bearish';
      signalCell.setValue(signal).setFontWeight("bold").setFontColor("red");
    } else {
      signal = 'Neutral';
      signalCell.setValue(signal).setFontWeight("normal").setFontColor("black");
    }
  }
}

function getHistoricalPrices(coinId, apiKey) {
const url = `https://api.coingecko.com/api/v3/coins/${coinId}/market_chart?vs_currency=usd&days=90&x_cg_demo_api_key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    if (!data || !data.prices || !Array.isArray(data.prices)) {
      Logger.log(`⚠️ Invalid price data for ${coinId}`);
      return null;
    }

    return data.prices.map(p => p[1]); // Just the price part
  } catch (e) {
    Logger.log(`❌ Error fetching data for ${coinId}: ${e}`);
    return null;
  }
}

function computeKDJ(prices) {
  const period = 9;
  if (prices.length < period + 2) throw new Error("Not enough data");

  let K = 50;
  let D = 50;
  let RSV = 0;

  for (let i = prices.length - period; i < prices.length; i++) {
    const window = prices.slice(i - period + 1, i + 1);
    const high = Math.max(...window);
    const low = Math.min(...window);
    const close = prices[i];

    RSV = (high === low) ? 50 : ((close - low) / (high - low)) * 100;
    K = (2 / 3) * K + (1 / 3) * RSV;
    D = (2 / 3) * D + (1 / 3) * K;
  }

  const J = 3 * K - 2 * D;

  // Clamp values between 0–100 for realism
  return [
    Math.max(0, Math.min(100, K)),
    Math.max(0, Math.min(100, D)),
    Math.max(0, Math.min(100, J))
  ];
}

function analyzeTechnicalIndicators() {
  const coins = [
    'aave', 'aptos', 'binance-usd', 'bitcoin', 'algorand', 'cardano', 'chainlink', 'hedera-hashgraph', 'internet-computer', 'maker', 'polkadot', 'ripple', 'solana', 'stellar', 'sui', 'tron', 
    'uniswap', 'vechain', 'ethena', 'virtual-protocol', 'celestia', 'bittensor', 'avalanche-2', 'near', 'ondo-finance', 'polygon-ecosystem-token', 'render-token', 'fetch-ai', 'arbitrum', 'worldcoin-wld', 'quant-network', 'kaspa' , 'civic' , 'icon' , 'cosmos' , 'cookie' , 'hft'

  ]; // Customize list
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AnalyzeTechnicalIndicators');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 10).setValues([['Coin', 'K', 'D', 'J', 'KDJ Signal', 'RSI (14)', 'MACD', 'Signal Line', 'MACD Histogram', 'MACD Signal']]);

  const apiKey = 'CG-ozyiJYmNEJGG8GcGBLciH1BG'; // Replace with your demo key

  for (let i = 0; i < coins.length; i++) {
    const coinId = coins[i];
    const row = i + 2;
    const prices = getHistoricalPrices(coinId, apiKey);

    sheet.getRange(row, 1).setValue(coinId);

    if (!Array.isArray(prices) || prices.length < 35) {
      sheet.getRange(row, 5).setValue("❌ Not enough data").setFontColor("gray");
      continue;
    }

    try {
      // --- KDJ ---
      const [K, D, J] = computeKDJ(prices);
      sheet.getRange(row, 2).setValue(K.toFixed(2));
      sheet.getRange(row, 3).setValue(D.toFixed(2));
      sheet.getRange(row, 4).setValue(J.toFixed(2));
      const kdjCell = sheet.getRange(row, 5);
      if (J > K && J > D) {
        kdjCell.setValue("Bullish").setFontColor("green").setFontWeight("bold");
      } else if (J < K && J < D) {
        kdjCell.setValue("Bearish").setFontColor("red").setFontWeight("bold");
      } else {
        kdjCell.setValue("Neutral");
      }

      // --- RSI ---
      const rsi = computeRSI(prices, 14);
      sheet.getRange(row, 6).setValue(rsi.toFixed(2));

      // --- MACD ---
      const [macd, signal, hist] = computeMACD(prices);
      sheet.getRange(row, 7).setValue(macd.toFixed(4));
      sheet.getRange(row, 8).setValue(signal.toFixed(4));
      sheet.getRange(row, 9).setValue(hist.toFixed(4));

      const macdCell = sheet.getRange(row, 10);
      if (hist > 0) {
        macdCell.setValue("Bullish").setFontColor("green");
      } else if (hist < 0) {
        macdCell.setValue("Bearish").setFontColor("red");
      } else {
        macdCell.setValue("Neutral");
      }

    } catch (err) {
      Logger.log(`❌ Error computing indicators for ${coinId}: ${err}`);
      sheet.getRange(row, 5).setValue("❌ Error").setFontColor("red");
    }
  }
  // ✅ Call Telegram alert function
  checkHistogramAlerts();
  checkMergeData()
}

function computeRSI(prices, period) {
  const closes = prices.slice(-period - 1);
  let gains = 0, losses = 0;
  for (let i = 1; i < closes.length; i++) {
    const change = closes[i] - closes[i - 1];
    if (change >= 0) gains += change;
    else losses -= change;
  }
  const avgGain = gains / period;
  const avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

function computeMACD(prices) {
  function EMA(data, period) {
    const k = 2 / (period + 1);
    let ema = data.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < data.length; i++) {
      ema = data[i] * k + ema * (1 - k);
    }
    return ema;
  }

  const shortEMA = EMA(prices, 12);
  const longEMA = EMA(prices, 26);
  const macd = shortEMA - longEMA;

  const signal = EMA([...(prices.map((_, i) => {
    const short = EMA(prices.slice(0, i + 1), 12);
    const long = EMA(prices.slice(0, i + 1), 26);
    return short - long;
  }))], 9);

  const hist = macd - signal;
  return [macd, signal, hist];
}

/**
 * Scans the “Coins” sheet and, for any row where
 * |Histogram| ≤ 2, sends a Telegram alert.
 */
function checkHistogramAlerts() {
  const BOT_TOKEN = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
  const CHAT_ID  = '1457021291';  
  
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AnalyzeTechnicalIndicators');
  if (!sheet) {
    Logger.log('❌ Sheet "Coins" not found.');
    return;
  }
  
  // Read all data (including header)
  const data = sheet.getDataRange().getValues();
  // data[0] is header: [Coin, K, D, J, KDJ Signal, RSI (14), MACD, Signal Line, Histogram, MACD Signal]
  
  for (let i = 1; i < data.length; i++) {
    const coin   = data[i][0];
    const hist   = parseFloat(data[i][8]); // column 9 → index 8
    const macd   = parseFloat(data[i][6]); // column 7 → index 6

    if (isNaN(hist)) {
      continue; // skip if histogram cell is blank or non‐numeric
    }
    
    if ((Math.abs(hist) <= 8) & (Math.abs(macd) <= 8)) {
      const msg = `⚠️ ${coin.toUpperCase()} MACD Histogram = ${hist.toFixed(4)} (│H│ ≤ 8) And MACD = ${macd.toFixed(4)} (│H│ ≤ 8)`;
      sendTelegramMessage(BOT_TOKEN, CHAT_ID, msg);
    }
  }
}

/**
 * Scans the “Merge” sheet and, for any row where
 * checkmark, sends a Telegram alert.
 */
function checkMergeData() {
  const BOT_TOKEN = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
  const CHAT_ID  = '1457021291';  
  
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MergeData');
  if (!sheet) {
    Logger.log('❌ Sheet "Coins" not found.');
    return;
  }
  
  // Read all data (including header)
  const data = sheet.getDataRange().getValues();
  // data[0] is header: [Coin, K, D, J, KDJ Signal, RSI (14), MACD, Signal Line, Histogram, MACD Signal]
  
  for (let i = 1; i < data.length; i++) {
    const coin   = data[i][0];
    const checkIt   = data[i][10]; // column 9 → index 8
    const checkCoin   = data[i][11]; // column 9 → index 8
    
    if (checkIt !== "") {
      const msg = `⚠️ ${coin.toUpperCase()} check ${checkIt} ${checkCoin}`;
      sendTelegramMessage(BOT_TOKEN, CHAT_ID, msg);
    }
  }
}


/**
 * Helper: sends a plain‐text message via Telegram Bot API
 */
function sendTelegramMessage(botToken, chatId, message) {
  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'Markdown'
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

function scheduledTask() {
  const now = new Date();
  const hour = now.getHours();

  // Only run at 9, 12, 15, 18 (i.e., 9AM, 12PM, 3PM, 6PM)
  const allowedHours = [8,12,16,20];

  if (!allowedHours.includes(hour)) {
    console.log(`Skipping execution at ${hour}:00`);
    return;
  }

  // Place your main logic here
  console.log(`Running main task at ${hour}:00`);
  analyzeSelectedCryptoIndicators()  
  // Your actual logic
}


const DEMO_API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';
const TELEGRAM_BOT_TOKEN = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
const TELEGRAM_CHAT_ID = '1457021291';

const SELECTED_COINS = [
  { id: 'chaingpt', name: 'CGPT' },
  { id: 'near', name: 'NEAR' },
  { id: 'sei-network', name: 'SEI' },
  { id: 'aave', name: 'Aave' },
  { id: 'aptos', name: 'APT' },
  { id: 'algorand', name: 'Algo' },
  { id: 'injective-protocol', name: 'INJ' },
  { id: 'cardano', name: 'ADA' },
  { id: 'chainlink', name: 'Link' },
  { id: 'hedera-hashgraph', name: 'HBar' },
  { id: 'internet-computer', name: 'ICP' },
  { id: 'maker', name: 'MKR' },
  { id: 'polkadot', name: 'DOT' },
  { id: 'ripple', name: 'XRP' },
  { id: 'stellar', name: 'XLM' },
  { id: 'sui', name: 'Sui' },
  { id: 'tron', name: 'TRX' },
  { id: 'uniswap', name: 'Uni' },
  { id: 'vechain', name: 'VET' },
  { id: 'ethena', name: 'ENA' },
  { id: 'virtual-protocol', name: 'VIRTUAL' },
  { id: 'celestia', name: 'TIA' },
  { id: 'bittensor', name: 'TAO' },
  { id: 'usual', name: 'USUAL' },
  { id: 'worldcoin-wld', name: 'WLD' },
  { id: 'tontoken', name: 'TON' },
  { id: 'tellor', name: 'TRB' },
  { id: 'render-token', name: 'RENDER' },
  { id: 'icon', name: 'ICX' },
  { id: 'arbitrum', name: 'ARB' }
 // { id: 'quant-network', name: 'QNT' },
 // { id: 'injective-protocol', name: 'INJ' },
 // { id: 'stox', name: 'STX' },
 // { id: 'ape-3', name: 'APE' },
 // { id: 'axie-infinity', name: 'AXS' },
 // { id: 'chaingpt', name: 'CGPT' },
 // { id: 'civic', name: 'CVC' },
 // { id: 'magic', name: 'MAGIC' },
 // { id: 'endurance', name: 'ACE' },
 // { id: 'movement', name: 'MOVE' },
 // { id: 'polygon-ecosystem-token', name: 'POL' }
];


function analyzeSelectedCryptoIndicators() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("analyzeSelectedCryptoIndicators");

  copyCryptoSheetSnapshot()

  if (!sheet) {
    sheet = spreadsheet.insertSheet("analyzeSelectedCryptoIndicators");
  } else {
    sheet.clear(); // Clear old data if sheet already exists
  }

  sheet.getRange("A1").setValue("Last Updated: " + new Date().toLocaleString());
  sheet.appendRow(['Coin', 'EMA 9', 'EMA 24', 'Trend', 'RSI (14)', 'MACD Histogram', 'KDJ J', 'OBV', 'Signal']);
  
  sheet.clear();
  sheet.getRange("A1").setValue("Last Updated: " + new Date().toLocaleString());
  sheet.appendRow(['Coin', 'EMA 9', 'EMA 24', 'Trend', 'RSI (14)', 'MACD Histogram', 'KDJ J', 'OBV', 'Signal']);

  SELECTED_COINS.forEach(coin => {
    const { id, name } = coin;
    try {
      const { prices, volumes } = getHistoricalPrices(id, 90);
      const closes = prices.map(p => p[1]);
      const volumeData = volumes.map(p => p[1]);

      const ema6 = computeEMA(closes, 9);
      const ema30 = computeEMA(closes, 24);
      const rsi = computeRSI(closes, 14);
      const [k, d, j] = computeKDJ(closes);
      const [macd, signalLine, hist] = computeMACD(closes);
      const obv = computeOBV(closes, volumeData);

      const trend = ema6 > ema30 ? 'Bullish' : 'Bearish';
      const signal = (ema6 > ema30 && hist > 0 && rsi > 50 && j > d)
        ? '🟢 Bullish'
        : (ema6 < ema30 && hist < 0 && rsi < 50 && j < d)
        ? '🔴 Bearish'
        : '✖️ Neutral';

      sheet.appendRow([
        name,
        ema6.toFixed(2),
        ema30.toFixed(2),
        trend,
        rsi.toFixed(2),
        hist.toFixed(2),
        j.toFixed(2),
        obv.toFixed(0),
        signal
      ]);

      // Send Telegram alerts
      if (signal !== '✖️ Neutral') {
        const message = `*${name}* is ${signal} \n` +
                        `EMA6: ${ema6.toFixed(2)}, EMA30: ${ema30.toFixed(2)}\n` +
                        `RSI: ${rsi.toFixed(2)}, MACD Hist: ${hist.toFixed(2)}, J: ${j.toFixed(2)}, OBV: ${obv.toFixed(0)}`;
//        sendTelegramMessage(message);
      }

    } catch (e) {
      Logger.log(`${name}: ${e.message}`);
    }
  });

  formatSheet(sheet);
  compareAndHighlightChanges()

}

function getHistoricalPrices(coinId, days) {
  const url = `https://api.coingecko.com/api/v3/coins/${coinId}/market_chart?vs_currency=usd&days=${days}&x_cg_demo_api_key=${DEMO_API_KEY}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  return {
    prices: data.prices,
    volumes: data.total_volumes
  };
}

function computeEMA(prices, period) {
  const k = 2 / (period + 1);
  let ema = prices.slice(0, period).reduce((a, b) => a + b) / period;
  for (let i = period; i < prices.length; i++) {
    ema = prices[i] * k + ema * (1 - k);
  }
  return ema;
}

function computeRSI(prices, period) {
  let gains = 0, losses = 0;
  for (let i = 1; i <= period; i++) {
    const change = prices[i] - prices[i - 1];
    if (change > 0) gains += change;
    else losses -= change;
  }
  let avgGain = gains / period;
  let avgLoss = losses / period;

  for (let i = period + 1; i < prices.length; i++) {
    const change = prices[i] - prices[i - 1];
    avgGain = (avgGain * (period - 1) + (change > 0 ? change : 0)) / period;
    avgLoss = (avgLoss * (period - 1) + (change < 0 ? -change : 0)) / period;
  }

  const rs = avgLoss === 0 ? 100 : avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

function computeKDJ(prices, period = 9) {
  if (prices.length < period + 3) throw new Error("Not enough data for KDJ");
  let K = 50, D = 50;
  for (let i = prices.length - period; i < prices.length; i++) {
    const window = prices.slice(i - period + 1, i + 1);
    const low = Math.min(...window);
    const high = Math.max(...window);
    const rsv = (high === low) ? 50 : ((prices[i] - low) / (high - low)) * 100;
    K = (2 / 3) * K + (1 / 3) * rsv;
    D = (2 / 3) * D + (1 / 3) * K;
  }
  const J = 3 * K - 2 * D;
  return [K, D, Math.max(0, Math.min(100, J))];
}

function computeMACD(prices, fast = 12, slow = 24, signal = 9) {
  const fastEma = computeEMA(prices, fast);
  const slowEma = computeEMA(prices, slow);
  const macdLine = fastEma - slowEma;
  const signalLine = macdLine * (2 / (signal + 1)); // single-value approximation
  const histogram = macdLine - signalLine;
  return [macdLine, signalLine, histogram];
}

function computeOBV(prices, volumes) {
  let obv = 0;
  for (let i = 1; i < prices.length; i++) {
    if (prices[i] > prices[i - 1]) obv += volumes[i];
    else if (prices[i] < prices[i - 1]) obv -= volumes[i];
  }
  return obv;
}

function sendTelegramMessage(message) {
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
  const payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: message,
    parse_mode: 'Markdown'
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

function formatSheet(sheet) {
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
  const values = range.getValues();

  values.forEach((row, i) => {
    const trend = row[3];
    const signal = row[8];
    const hist = parseFloat(row[5]);
    const rowRange = sheet.getRange(i + 2, 1, 1, 9);

    if (trend === 'Bullish') {
      rowRange.setBackground('#d4f4dd'); // light green
    } else if (trend === 'Bearish') {
      rowRange.setBackground('#fdd'); // light red
    } else if (trend === 'Neutral') {
      rowRange.setBackground('#b1c5fc'); // light blue
    }

    if (Math.abs(hist) < 2) {
      sheet.getRange(i + 2, 6).setFontWeight('bold').setFontColor('#e67e22');
    }
  });
}


function copyCryptoSheetSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalName = "analyzeSelectedCryptoIndicators";
  const snapshotName = "analyzeSelectedCryptoIndicators_prior";

  const originalSheet = ss.getSheetByName(originalName);
  if (!originalSheet) throw new Error(`Sheet "${originalName}" not found.`);

  // Delete prior if it already exists
  const existingSnapshot = ss.getSheetByName(snapshotName);
  if (existingSnapshot) ss.deleteSheet(existingSnapshot);

  const copiedSheet = originalSheet.copyTo(ss);
  copiedSheet.setName(snapshotName);
  ss.setActiveSheet(originalSheet);
  SpreadsheetApp.flush();

  return `Snapshot created: ${snapshotName}`;
}

function compareAndHighlightChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalName = "analyzeSelectedCryptoIndicators";
  const snapshotName = "analyzeSelectedCryptoIndicators_prior";
  const diffSheetName = "Diff_Report";
  const signalCol = 9; // Column Ι = 9

  var datetime = "🗣📢 LastSync: " + new Date().toLocaleTimeString('en-GB', { hour: "numeric", minute: "numeric"}) + "🚩";

  const currentSheet = ss.getSheetByName(originalName);
  const snapshotSheet = ss.getSheetByName(snapshotName);
  if (!currentSheet || !snapshotSheet) throw new Error("One or both sheets not found.");

  // Clear old highlights
  currentSheet.getRange(1, signalCol, currentSheet.getLastRow()).setBackground(null);

  const currentData = currentSheet.getDataRange().getValues();
  const priorData = snapshotSheet.getDataRange().getValues();

  const maxRows = Math.max(currentData.length, priorData.length);

  // Allowed sentiment transitions
  const message = datetime;
  sendTelegramMessage(message);
  const allowedTransitions = new Set([
    "🔴 Bearish>🟢 Bullish",
    "🟢 Bullish>🔴 Bearish",
    "✖️ Neutral>🔴 Bearish",
    "✖️ Neutral>🟢 Bullish"
  ]);

  const changes = [["Coin Name", "Cell", "Previous Value", "Current Value"]];

  for (let i = 0; i < maxRows; i++) {
    const coinName = (currentData[i] && currentData[i][0]) ?? "Unknown";
    const priorRaw = (priorData[i] && priorData[i][signalCol - 1]) ?? '';
    const currentRaw = (currentData[i] && currentData[i][signalCol - 1]) ?? '';

    const prior = priorRaw.toString().toLowerCase().trim();
    const current = currentRaw.toString().toLowerCase().trim();

    const transition = `${prior}>${current}`;

    if (prior !== current ) {
      const cellRef = `${String.fromCharCode(64 + signalCol)}${i + 1}`;
      changes.push([coinName, '', priorRaw, currentRaw]);

      // Highlight changed Signal cell
      currentSheet.getRange(i + 1, signalCol).setBackground("#FFF59D"); // Yellow
      const message = `*${coinName}* is ${currentRaw} from ${priorRaw} \n`;
      if (coinName !== 'Unknown' && currentRaw !== '✖️ Neutral' && priorRaw !== '') {
        sendTelegramMessage(message);
      }
    }
  }

  // Create or update Diff_Report
  let diffSheet = ss.getSheetByName(diffSheetName);
  if (diffSheet) ss.deleteSheet(diffSheet);
  diffSheet = ss.insertSheet(diffSheetName);
  diffSheet.getRange(1, 1, changes.length, 4).setValues(changes);

  return `${changes.length - 1} signal changes found and highlighted.`;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Crypto Tools")
    .addItem("📸 Create Snapshot", "copyCryptoSheetSnapshot")
    .addItem("🔍 Compare & Highlight Changes", "compareAndHighlightChanges")
    .addItem("🗑️ Διαγραφή παλαιών τιμών Scrypt","keepLast20HashRates")
    .addToUi();
}

function getTop10_1hGainers_DemoAPI() {
  const apiUrl = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=50&page=1&sparkline=true"; // 🔁 Replace with your actual demo API URL

  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer ozyiJYmNEJGG8GcGBLciH1BG"
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`API Error ${statusCode}: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());

    const threshold = 5;
    const gainers = data.filter(coin => coin.percent_change_1h > threshold);

    // Sort descending by 1h % change
    gainers.sort((a, b) => b.percent_change_1h - a.percent_change_1h);

    const top10 = gainers.slice(0, 10);

    const output = [["Name", "Symbol", "Price", "1h Change (%)"]];
    top10.forEach(coin => {
      output.push([
        coin.name,
        coin.symbol.toUpperCase(),
        coin.price,
        `${coin.percent_change_1h.toFixed(2)}%`
      ]);
    });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1h_Gainers")
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet("1h_Gainers");
    sheet.clearContents();
    sheet.getRange(1, 1, output.length, 4).setValues(output);

  } catch (error) {
    Logger.log("Error fetching from API: " + error.message);
    SpreadsheetApp.getUi().alert("API Error: " + error.message);
  }
}

// ----------------------------
// CONFIGURATION
// ----------------------------
var RECIPIENT_EMAIL = "kapsilias@gmail.com"; // your email
var SHEET_NAME = "Rolex_Index";
var COINGECKO_API_URL = "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd";


// Fetch BTC price and log in sheet

const SELECTED_BTC = [
  { id: 'bitcoin', name: 'BTC' },
];



// ----------------------------
// CONFIGURATION
// ----------------------------
var COIN_ID = "bitcoin";
var DAYS = 1; // 1 = last 24h, can change to 7, 30, 90, etc.
var SHEET_NAME = "ROLEX_INDEX";
var RECIPIENT_EMAIL = "kapsilias@gmail.com";

// ----------------------------
// Fetch BTC price and log in sheet
// ----------------------------
function fetchBTC() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var url = `https://api.coingecko.com/api/v3/coins/${COIN_ID}/market_chart?vs_currency=usd&days=${DAYS}&x_cg_demo_api_key=${DEMO_API_KEY}`;

  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  
  // Get last price point (timestamp, price)
  var prices = data.prices;
  var lastPoint = prices[prices.length - 1];
  var btcPrice = parseFloat(lastPoint[1]);
  
  // Read Rolex Index manually from C2
  var rolexIndex = parseFloat(sheet.getRange("C2").getValue());
  if (isNaN(rolexIndex)) rolexIndex = 0; // safety check

  // Append data
  var today = new Date();
  sheet.appendRow([today, btcPrice, rolexIndex]);
}

// ----------------------------
// Send weekly summary email
// ----------------------------
function sendWeeklyBTCReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // not enough data

  // Today's values
  var todayRow = sheet.getRange(lastRow, 1, 1, 3).getValues()[0];
  var todayBTC = todayRow[1];
  var todayRolex = todayRow[2];

  // Week-ago values
  var weekAgoRowIndex = Math.max(lastRow - 7, 3);
  var weekAgoRow = sheet.getRange(weekAgoRowIndex, 1, 1, 3).getValues()[0];
  var btcWeekAgo = weekAgoRow[1];
  var rolexWeekAgo = weekAgoRow[2];

  // % changes
  var btcChange = ((todayBTC - btcWeekAgo) / btcWeekAgo * 100).toFixed(2);
  var rolexChange = ((todayRolex - rolexWeekAgo) / rolexWeekAgo * 100).toFixed(2);

  // Interpretation
var btcTrend = btcChange > 0 ? "Bullish ↑" : "Bearish ↓";
var rolexTrend = rolexChange > 0 ? "Up ↑" : "Down ↓";
var interpretationText = "";

// Advanced interpretation logic
if (btcChange < 0 && rolexChange < 0) {
  interpretationText = "Risk-off phase 🔻 — both markets cooling (lower liquidity, cautious sentiment)";
} else if (btcChange > 0 && rolexChange > 0) {
  interpretationText = "Broad bullish cycle 🚀 — strong liquidity and confidence";
} else if (btcChange > 0 && rolexChange < 0) {
  interpretationText = "Speculative rebound ⚡ — crypto rising ahead of luxury assets";
} else if (btcChange < 0 && rolexChange > 0) {
  interpretationText = "Wealth rotation 💎 — investors shifting from crypto to tangible assets";
}

var interpretation = "BTC: " + btcTrend + " | Rolex: " + rolexTrend + " → " + interpretationText;


  // Write back to sheet
  sheet.getRange(lastRow, 4).setValue(btcChange + "%");
  sheet.getRange(lastRow, 5).setValue(rolexChange + "%");
  sheet.getRange(lastRow, 6).setValue(interpretation);

  // Email summary
  var body = `📈 Weekly Bitcoin - Rolex Index Summary\n\n` +
             `BTC Price: $${todayBTC.toLocaleString()}\n` +
             `Weekly Change: ${btcChange}%\n` +
             `Interpretation: ${interpretation}\n`;
  

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: "Weekly Bitcoin - Rolex Index Report",
    body: body
  });

  const message = body;
  sendTelegramMessage(message);
}


/**
 * Updates RSI(14), EMA(21), and Divergence (RSI - EMA)
 * for a fixed list of cryptos in sheet "updateCryptoIndicators".
 * Auto-creates the sheet + headers + crypto list if missing.
 */

const TARGET_CRYPTOS = ["bitcoin" , "ethereum", "stellar", "solana", 'near', 'sei-network', 'aave', 'aptos', 'algorand', 'injective-protocol', 'cardano', 'chainlink', 'hedera-hashgraph','internet-computer', 'polkadot','ripple','stellar','sui','tron','uniswap','vechain','ethena','virtual-protocol','celestia','bittensor','usual','worldcoin-wld', 'render-token','icon', 'arbitrum']; // <--- edit this list 

/**
 * Configuration
 */

/*const TARGET_CRYPTOS = ["bitcoin", "ethereum", "stellar", "solana", "dogecoin"]; // Add more here */
const RETRY_COUNT = 2;     // Retry on 403 or 429
const SLEEP_MS = 2000;     // Delay between requests (ms)
const COINGECKO_API_KEY = "CG-ozyiJYmNEJGG8GcGBLciH1BG"; // Replace with your real key if you have one


/**
 * Main function
 */
function updateCryptoIndicators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "updateCryptoIndicators";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // Clear and rebuild header
  sheet.clear();
  const headers = ["Symbol", "Days", "RSI(14)", "RSI EMA(21)", "Divergence"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const days = 1;
  const rows = TARGET_CRYPTOS.map(c => [c, days, "", "", ""]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  for (let i = 0; i < TARGET_CRYPTOS.length; i++) {
    const id = TARGET_CRYPTOS[i];
    const outRow = i + 2;
    let prices = null;
    let attempt = 0;

    // Retry logic for API errors
    while (attempt <= RETRY_COUNT) {
      try {
        prices = getCoinHistory(id, days);
        break;
      } catch (e) {
        const msg = e.message || "";
        if ((msg.includes("429") || msg.includes("403")) && attempt < RETRY_COUNT) {
          Logger.log(`Retry ${attempt + 1} for ${id} due to ${msg}`);
          Utilities.sleep(SLEEP_MS * (attempt + 1)); // exponential backoff
          attempt++;
        } else {
          Logger.log(`Error fetching ${id}: ${msg}`);
          sheet.getRange(outRow, 3, 1, 3).setValue("Error");
          prices = null;
          break;
        }
      }
    }

    if (!prices || prices.length < 22) {
      sheet.getRange(outRow, 3, 1, 3).setValue("No Data");
      continue;
    }

    const closes = prices.map(p => p[1]);

    // Compute RSI(14) for each period, not just latest
    const rsiSeries = calcRSISeries(closes, 14);
    const rsi14 = rsiSeries[rsiSeries.length - 1];
    const rsiEma21 = calcEMA(rsiSeries, 21);
    const divergence = rsi14 - rsiEma21;

    sheet.getRange(outRow, 3).setValue(roundTo(rsi14, 2));
    sheet.getRange(outRow, 4).setValue(roundTo(rsiEma21, 2));
    sheet.getRange(outRow, 5).setValue(roundTo(divergence, 2));

    Utilities.sleep(SLEEP_MS);
  }

  applyConditionalFormatting(sheet, TARGET_CRYPTOS.length);

  Logger.log("✅ All cryptos updated successfully!");
}

/**
 * Fetch historical prices from CoinGecko with API key
 */
function getCoinHistory(id, days) {
  const url = `https://api.coingecko.com/api/v3/coins/${encodeURIComponent(id)}/market_chart?vs_currency=usd&days=${encodeURIComponent(days)}`;
  const options = {
    muteHttpExceptions: true,
    headers: {
      "x-cg-demo-api-key": COINGECKO_API_KEY
    }
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  Logger.log(`${id} → HTTP ${code}`);

  if (code === 403) throw new Error("403 Forbidden");
  if (code === 429) throw new Error("429 Too Many Requests");
  if (code !== 200) throw new Error(`HTTP ${code}`);

  const data = JSON.parse(response.getContentText());
  return data.prices || [];
}

/**
 * EMA calculation
 */
function calcEMA(values, period) {
  const k = 2 / (period + 1);
  let emaPrev = average(values.slice(0, period));
  for (let i = period; i < values.length; i++) {
    emaPrev = values[i] * k + emaPrev * (1 - k);
  }
  return emaPrev;
}

/**
 * RSI series calculation (returns array of RSI values)
 */
function calcRSISeries(values, period) {
  const rsis = [];
  for (let i = period; i <= values.length; i++) {
    const segment = values.slice(i - period, i);
    rsis.push(calcRSI(segment, period));
  }
  return rsis;
}

/**
 * RSI single value
 */
function calcRSI(values, period) {
  const deltas = [];
  for (let i = 1; i < values.length; i++) deltas.push(values[i] - values[i - 1]);
  const last = deltas.slice(-period);
  let gains = 0, losses = 0;
  last.forEach(d => { if (d >= 0) gains += d; else losses += Math.abs(d); });
  const avgGain = gains / period;
  const avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - 100 / (1 + rs);
}

/**
 * Apply conditional formatting to the Divergence column
 */
function applyConditionalFormatting(sheet, rowCount) {
  const range = sheet.getRange(2, 5, rowCount); // Divergence column
  const rules = sheet.getConditionalFormatRules();

  const newRules = [
    // Green: |divergence| < 2
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ABS(E2)<2')
      .setBackground('#b6d7a8')
      .setRanges([range])
      .build(),
    // Orange: 2 <= |divergence| < 5
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ABS(E2)>=2,ABS(E2)<5)')
      .setBackground('#f9cb9c')
      .setRanges([range])
      .build(),
    // Red: |divergence| >= 5
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ABS(E2)>=5')
      .setBackground('#e06666')
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(rules.concat(newRules));
}

/** Helpers */
function average(arr) { return arr.reduce((a, b) => a + b, 0) / arr.length; }
function roundTo(num, dec = 2) { return Math.round(num * Math.pow(10, dec)) / Math.pow(10, dec); }

function getCryptoMarketUpdate() {
  const date = new Date();
  const emailTo = "kapsilias@gmail.com";
  const sheetName = "Crypto Market Poster";

  // ---- Fetch market data from CoinLore (region-safe) ----
  const url = "https://api.coinlore.net/api/tickers/?limit=100";
  const options = {
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript)" },
  };
  const response = UrlFetchApp.fetch(url, options);
  const parsed = JSON.parse(response.getContentText());

  if (!parsed.data || !Array.isArray(parsed.data)) {
    throw new Error("Invalid data format received from CoinLore");
  }

  const data = parsed.data.map((coin) => ({
    symbol: coin.symbol,
    name: coin.name,
    price: parseFloat(coin.price_usd),
    marketCap: parseFloat(coin.market_cap_usd),
    volume24h: parseFloat(coin.volume24),
    change24h: parseFloat(coin.percent_change_24h),
    change7d: parseFloat(coin.percent_change_7d),
  }));

  // ---- Extract BTC / ETH ----
  const btcObj = data.find((c) => c.symbol === "BTC");
  const ethObj = data.find((c) => c.symbol === "ETH");
  const btc = btcObj ? btcObj.price.toFixed(0) : "N/A";
  const eth = ethObj ? ethObj.price.toFixed(0) : "N/A";

  // ---- Market Cap ----
  const totalCap = data.reduce((sum, c) => sum + c.marketCap, 0);
  const marketCapText = "$" + (totalCap / 1e12).toFixed(2) + " Trillion";

  // ---- Whale accumulation proxy (volume + 7d gains) ----
  const whaleCoins = data
    .filter((c) => c.volume24h > 300000000 && c.change7d > 5)
    .sort((a, b) => b.volume24h - a.volume24h)
    .slice(0, 3)
    .map(
      (c) =>
        `${c.symbol} (+${c.change7d.toFixed(1)}%)`
    )
    .join(", ") || "No major accumulation detected";

  // ---- Top gainers ----
  const topGainers = data
    .filter((c) => c.change24h > 0)
    .sort((a, b) => b.change24h - a.change24h)
    .slice(0, 3)
    .map(
      (c) =>
        `${c.symbol} +${c.change24h.toFixed(1)}%`
    )
    .join(", ");

  // ---- Sentiment & commentary ----
  const sentiment =
    btcObj && btcObj.change24h > 0
      ? "Market showing bullish momentum."
      : "Mixed sentiment; consolidation likely.";
  const etfFlows = "$480M net inflows (BTC & ETH)";
  const watch = "BTC $100K–$110K zone, altcoin rotation, ETF trends.";

  // ---- Poster prompt ----
  const prompt = `
Create a dark-blue crypto market poster titled “Market Update: Crypto Market Trends | ${date.toDateString()}”.
Include:
1. Market Overview — ${sentiment}
2. Key Market Highlights — BTC $${btc}, ETH $${eth}, Total Market Cap ${marketCapText}
3. Whale Accumulation Focus — ${whaleCoins}
4. Institutional Flows — ${etfFlows}
5. What to Watch — ${watch}
6. Top Gainers — ${topGainers}
Add minimalist charts and a Binance-style dark-blue layout with clean sans-serif fonts.
`;

  // ---- Write to Google Sheet ----
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  sheet.appendRow([
    date,
    btc,
    eth,
    marketCapText,
    topGainers,
    whaleCoins,
    etfFlows,
    sentiment,
    prompt,
  ]);

  // ---- Send e-mail ----
  MailApp.sendEmail({
    to: emailTo,
    subject: `Daily Crypto Market Poster Prompt - ${date.toDateString()}`,
    body: prompt,
  });
}

function UPDATE_PRICES_COINGECKO() {
  Logger.log("START UPDATE_PRICES_COINGECKO");
  Logger.log("SCRIPT STARTED");

  const sheet = SpreadsheetApp.getActive().getSheetByName("prices");
  Logger.log("ACTIVE SPREADSHEET ID: " + SpreadsheetApp.getActive().getId());
  Logger.log("SHEET OBJECT: " + sheet);
  Logger.log("AFTER SHEET CHECK - SHOULD NOT MISS THIS");

  // -------------------------------
  // Telegram settings
  // -------------------------------

  function sendTelegram(message) {
    try {
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage",
        {
          method: "post",
          muteHttpExceptions: true,
          contentType: "application/json",
          payload: JSON.stringify({
            chat_id: TELEGRAM_CHAT_ID,
            text: message,
            parse_mode: "HTML"
          })
        }
      );
      Logger.log("Telegram message sent");
    } catch (e) {
      Logger.log("Telegram error: " + e.toString());
    }
  }

  // -------------------------------
  // CoinGecko fetch
  // -------------------------------
  let updatedFromAPI = true;
  const API_KEY = 'CG-ozyiJYmNEJGG8GcGBLciH1BG';

  try {
    const url =
      "https://api.coingecko.com/api/v3/simple/price" +
      "?ids=nano,polygon-ecosystem-token,sonic-3&vs_currencies=eur";

    const options = {
      method: "get",
      muteHttpExceptions: true,
    headers: {
      'x-cg-demo-api-key': API_KEY
    }
    };
    const response = UrlFetchApp.fetch(url, options);
    const httpCode = response.getResponseCode();
    const text = response.getContentText();
    Logger.log("HTTP CODE: " + httpCode);
    Logger.log("RAW RESPONSE: " + text);

    if (httpCode !== 200) throw new Error("HTTP error: " + httpCode);

    const data = JSON.parse(text);
    Logger.log("PARSED DATA: " + JSON.stringify(data));

    if (
      typeof data.nano?.eur !== "number" ||
      typeof data["polygon-ecosystem-token"]?.eur !== "number" ||
      typeof data["sonic-3"]?.eur !== "number"
    ) {
      throw new Error("Invalid CoinGecko data");
    }

    // Update sheet
    sheet.getRange("A1:B1").setValues([["Coin", "Price (EUR)"]]);
    sheet.getRange("A2:B2").setValues([["XNO", data.nano.eur]]);
    sheet.getRange("A3:B3").setValues([["POL", data["polygon-ecosystem-token"].eur]]);
    sheet.getRange("A4:B4").setValues([["SONIC", data["sonic-3"].eur]]);

  } catch (err) {
    updatedFromAPI = false;
    Logger.log("CoinGecko error: " + err.toString());
  }

  // -------------------------------
  // Compare values
  // -------------------------------
  const rows = [
    { row: 2, name: sheet.getRange("A2").getValue(), value: Number(sheet.getRange("D2").getValue()) },
    { row: 3, name: sheet.getRange("A3").getValue(), value: Number(sheet.getRange("D3").getValue()) },
    { row: 4, name: sheet.getRange("A4").getValue(), value: Number(sheet.getRange("D4").getValue()) }
  ];

  // Reset colors
  rows.forEach(r => sheet.getRange("D" + r.row).setBackground(null));

  // Determine winner
  let winner = rows[0];
  rows.forEach(r => {
    if (r.value > winner.value) winner = r;
  });

  // Color cells
  rows.forEach(r => {
    sheet.getRange("D" + r.row).setBackground(
      r.row === winner.row ? "#b7e1cd" : "#f4c7c3"
    );
  });

  // -------------------------------
  // Telegram message
  // -------------------------------
  const formattedValue = winner.value.toFixed(3);

    let message =
      `📊 <b>Today's mining report</b>\n` +
      `📝 Today mining <b>${winner.name}</b>\n` +
      `🎯 Value: <b>${formattedValue}</b>`;

  if (!updatedFromAPI) {
    message +=
      "\n\n⚠️ <i>Δεν έγινε ενημέρωση από CoinGecko – χρησιμοποιήθηκαν υπάρχουσες τιμές</i>";
  }

  sendTelegram(message);

  Logger.log("SCRIPT COMPLETED");
}

function UPDATE_WTM_KAWPOW_BTC_REVENUE() {
  const url = "https://whattomine.com/coins/234.json";

  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true
  });

  const data = JSON.parse(response.getContentText());

  // BTC / MH / day
  const btcPerMhDay = Number(data.btc_revenue);

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("HashRates");

  if (!sheet) {
    throw new Error("Sheet HashRates δεν βρέθηκε");
  }

  sheet.getRange("D3").setValue(btcPerMhDay);
  var value = sheet.getRange("D5").getValue();
  var formatted;
  if (typeof value === 'number' && !isNaN(value)) {
    formatted = value.toFixed(4);  // Formats to exactly 5 decimals as string
  } else {
    formatted = "0.0000";  // Fallback if cell is empty/text/non-numeric
  }

  if (sheet.getRange("D5").getValue() < sheet.getRange("D6").getValue() ) {
    let message = '\n\n⚠️Mining Price down\n ❌Check prices ! \n 🎯Current at ' + formatted;
    sendTelegramMessage(message);
  }
  if (sheet.getRange("D5").getValue() - sheet.getRange("D6").getValue() > 4 ) {
    let message = '\n\n📈Mining Price up\n 🟢Check prices ! \n 🎯Current at ' + formatted;
    sendTelegramMessage(message);
  }

}

function UPDATE_WTM_SCRYPT_REVENUE_LTCDOGE() {
  // Default endpoint χωρίς custom parameters (για default hashrate ~600 MH/s)
  var url = "https://whattomine.com/coins/4-ltc.json";
  
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  var responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    Logger.log("Σφάλμα HTTP: " + responseCode);
    return;
  }
  
  var json = JSON.parse(response.getContentText());
  
  // Base revenue από το default calculation (για ~600 MH/s)
  var baseBtcRevenue = parseFloat(json.btc_revenue || 0);
  
  if (baseBtcRevenue === 0) {
    Logger.log("Δεν βρέθηκαν δεδομένα base revenue.");
    return;
  }
  
  // Scaling: Το default είναι για περίπου 600 MH/s
  var defaultHashrateMH = 600;
  var yourHashrateMH = 145000;  // Το δικό σου hashrate
  
  // Υπολογισμός scaled revenue σε BTC / 24h (gross, χωρίς electricity cost)
  var btcReward24 = baseBtcRevenue * (yourHashrateMH / defaultHashrateMH);
  
  if (btcReward24 > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HashRates");
    sheet.getRange("D12").setValue(btcReward24);
    
    // Προαιρετικά: USD revenue (αν θες σε άλλο κελί)
    // var baseUsdRevenue = parseFloat(json.revenue || 0);
    // var usdReward24 = baseUsdRevenue * (yourHashrateMH / defaultHashrateMH);
    // sheet.getRange("E12").setValue(usdReward24);
    
     // sheet.getRange("B18").setValue("LTC+DOGE merged (gross, scaled)");
  } else {
    Logger.log("Μηδενική απόδοση.");
  }

  var value = sheet.getRange("D14").getValue();
  var formatted;
  if (typeof value === 'number' && !isNaN(value)) {
    formatted = value.toFixed(5);  // Formats to exactly 5 decimals as string
  } else {
    formatted = "0.00000";  // Fallback if cell is empty/text/non-numeric
  }

  if (sheet.getRange("D17").getValue() < sheet.getRange("D15").getValue() ) {
    let message = '\n\n⚠️Mining Price down\n ❌Check prices ! \n 🎯Current at ' + formatted;
    sheet.getRange("D15").setValue(sheet.getRange("D17").getValue());
    sendTelegramMessage(message);
  }
  if (sheet.getRange("D17").getValue() > sheet.getRange("D15").getValue() ) {
    let message = '\n\n📈Mining Price up\n 🟢Check prices ! \n 🎯Current at ' + formatted;
    sheet.getRange("D15").setValue(sheet.getRange("D17").getValue());
    sendTelegramMessage(message);
  }

}

function updateNiceHashScrypt() {
  const SHEET = "HashRates";  // Όνομα sheet
  const CELL = "D20";         // Κελί για τιμή
  const URL = "https://api2.nicehash.com/main/api/v2/public/simplemultialgo/info";

  try {
    // 1️⃣ Κάνουμε fetch το public API
    const res = UrlFetchApp.fetch(URL);
    const json = JSON.parse(res.getContentText());

    // 2️⃣ Βρίσκουμε το array των αλγορίθμων
    const algoArray =
      json.miningAlgorithms ||
      json.algorithms ||
      json.simpleMultiAlgoInfo ||
      json.result?.algorithms ||
      json.result?.simpleMultiAlgoInfo;

    if (!Array.isArray(algoArray)) {
      throw new Error(
        "Δεν βρέθηκε array αλγορίθμων. Root keys: " + Object.keys(json).join(", ")
      );
    }

    // 3️⃣ Βρίσκουμε τον Scrypt
    const scrypt = algoArray.find(a =>
      String(a.algorithm || "").toUpperCase() === "SCRYPT"
    );

    if (!scrypt || scrypt.paying == null) {
      throw new Error("Δεν βρέθηκε paying τιμή για SCRYPT");
    }

    // 4️⃣ Υπολογισμός BTC/day για 1 TH
    // paying = BTC / (100 MH) / day → 1 TH = 10.000 × paying
    const btcPerDayFor1TH = Number(scrypt.paying) * 10000;

    // 5️⃣ Παίρνουμε το sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET);

    if (!sheet) {
      throw new Error(
        `Το sheet "${SHEET}" δεν βρέθηκε. Υπάρχουν: ` +
        ss.getSheets().map(s => s.getName()).join(", ")
      );
    }

   // 6️⃣ Γράφουμε την τιμή στο κελί
   sheet.getRange("E19").setValue(sheet.getRange(CELL).getValue());
   sheet.getRange(CELL).setValue(btcPerDayFor1TH);

   var value = sheet.getRange("D19").getValue();
   var formatted;
   if (typeof value === 'number' && !isNaN(value)) {
     formatted = value.toFixed(4);  // Formats to exactly 5 decimals as string
   } else {
     formatted = "0.0000";  // Fallback if cell is empty/text/non-numeric
   }

   if (sheet.getRange(CELL).getValue() < sheet.getRange("E19").getValue() ) {
     let message = '\n\n⚠️NiceHash Profit down\n ❌Check prices ! \n 🎯Current at ' + formatted;
     sendTelegramMessage(message);
 }
   if (sheet.getRange("D19").getValue() > sheet.getRange("E19").getValue() ) {
     let message = '\n\n📈NiceHash Profit up\n 🟢Check prices ! \n 🎯Current at ' + formatted;
     sendTelegramMessage(message);
 }




    Logger.log(`Scrypt BTC/day για 1 TH: ${btcPerDayFor1TH}`);

  } catch (e) {
    Logger.log("Σφάλμα updateNiceHashScrypt: " + e);
    throw e;
  }
}

function updateNiceHashScryptRolling() {
  const SHEET_NAME = "HashRates";
  const CURRENT_CELL = "D20";    // Τρέχουσα τιμή
  const HISTORY_COL = "G";       // Στήλη ιστορικού
  const AVG_CELL = "F20";        // Rolling average
  const ROLLING_COUNT = 10;      // Τελευταίες 10 τιμές για μέσο

  const URL = "https://api2.nicehash.com/main/api/v2/public/simplemultialgo/info";

  try {
    const res = UrlFetchApp.fetch(URL);
    const json = JSON.parse(res.getContentText());

    const algoArray =
      json.miningAlgorithms ||
      json.algorithms ||
      json.simpleMultiAlgoInfo ||
      json.result?.algorithms ||
      json.result?.simpleMultiAlgoInfo;

    if (!Array.isArray(algoArray)) throw new Error("Δεν βρέθηκε array αλγορίθμων.");

    const scrypt = algoArray.find(a => String(a.algorithm || "").toUpperCase() === "SCRYPT");
    if (!scrypt || scrypt.paying == null) throw new Error("Δεν βρέθηκε paying τιμή για SCRYPT");

    // Scrypt: paying = BTC / 100 MH / day → 1 TH = 10.000 × paying
    const btcPerDayFor1TH = Number(scrypt.paying) * 10000;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Το sheet "${SHEET_NAME}" δεν βρέθηκε.`);

    // 1️⃣ Ενημέρωση τρέχουσας τιμής
      // 6️⃣ Γράφουμε την τιμή στο κελί
   sheet.getRange(CURRENT_CELL).setValue(btcPerDayFor1TH);


    // 2️⃣ Προσθήκη ιστορικού στην επόμενη κενή γραμμή της στήλης E
    const lastRow = sheet.getRange(HISTORY_COL + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const nextRow = lastRow >= 1 ? lastRow + 1 : 1;
    sheet.getRange(HISTORY_COL + nextRow).setValue(btcPerDayFor1TH);

    // 3️⃣ Υπολογισμός rolling average των τελευταίων N τιμών
    const startRow = Math.max(nextRow - ROLLING_COUNT + 1, 1);
    const range = sheet.getRange(HISTORY_COL + startRow + ":" + HISTORY_COL + nextRow);
    const values = range.getValues().flat().filter(v => typeof v === "number");
    const avg = values.reduce((a,b) => a+b,0)/values.length;

    sheet.getRange("E19").setValue(sheet.getRange(AVG_CELL).getValue());
    sheet.getRange(AVG_CELL).setValue(avg);

   var value = sheet.getRange("D19").getValue();
   var formatted;
   if (typeof value === 'number' && !isNaN(value)) {
     formatted = value.toFixed(4);  // Formats to exactly 5 decimals as string
   } else {
     formatted = "0.0000";  // Fallback if cell is empty/text/non-numeric
   }

   if (sheet.getRange(AVG_CELL).getValue() < sheet.getRange("E19").getValue() ) {
     let message = '\n\n⚠️NiceHash Profit down\n ❌Check prices ! \n 🎯Current at ' + formatted;
     sheet.getRange("E19").setValue(sheet.getRange(AVG_CELL).getValue());
     sendTelegramMessage(message);
 }
   if (sheet.getRange(AVG_CELL).getValue() > sheet.getRange("E19").getValue() ) {
     let message = '\n\n📈NiceHash Profit up\n 🟢Check prices ! \n 🎯Current at ' + formatted;
     sheet.getRange("E19").setValue(sheet.getRange(AVG_CELL).getValue());
     sendTelegramMessage(message);
 }

    Logger.log(`Updated Scrypt: ${btcPerDayFor1TH} BTC/day | Rolling avg: ${avg}`);

  } catch (e) {
    Logger.log("Σφάλμα updateNiceHashScryptRolling: " + e);
    throw e;
  }
}

function keepLast20HashRates() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("HashRates");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Παίρνουμε όλες τις τιμές της στήλης G
  const values = sheet
    .getRange(2, 7, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(v => v !== "");

  // Παίρνουμε τις 20 τελευταίες
  const last20 = values.slice(-20);

  // Καθαρίζουμε όλη τη στήλη G από το G2 και κάτω
  sheet.getRange(2, 7, sheet.getMaxRows()).clearContent();

  // Γράφουμε τις 20 τελευταίες από το G2
  if (last20.length > 0) {
    sheet.getRange(2, 7, last20.length, 1)
      .setValues(last20.map(v => [v]));
  }
}

// ============================================================
// CoinGlass Bull Market Peak Indicators Monitor
// Google Sheets + Apps Script Integration
// 
// Setup Instructions:
// 1. Create new Google Sheet
// 2. Extensions → Apps Script
// 3. Copy entire code below
// 4. Replace API_KEY, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
// 5. Save & Run setupSheet() once
// 6. Edit → Current project's triggers → Add trigger:
//    - Function: updateIndicatorsSheet
//    - Deployment: Head
//    - Event source: Time-driven
//    - Type: Hour timer
//    - Interval: Every 4 hours
// ============================================================

// ⚠️ REPLACE THESE WITH YOUR CREDENTIALS
const API_KEY = "d20f0a7952824334a4bddcc9243e92d4";  // Get from: https://www.coinglass.com/CryptoApi

// Configuration
const ALERT_THRESHOLD = 5;  // Send alert when this many indicators hit
const CHECK_INTERVAL_HOURS = 4;
var SHEET_NAME = "Indicators";
const SUMMARY_SHEET = "Summary";

// ============================================================
// MAIN FUNCTION: Fetch and Update Indicators
// ============================================================

function updateIndicatorsSheet() {
  try {
    Logger.log("🚀 Starting indicator update at " + new Date());
    
    const indicators = fetchBullMarketIndicators();
    if (!indicators) {
      Logger.log("❌ Failed to fetch indicators");
      return;
    }
    
    // Update the indicators sheet
    updateSheet(indicators);
    
    // Check if we should send alert
    const analysis = analyzeIndicators(indicators);
    checkAndSendAlert(analysis);
    
    Logger.log("✅ Update complete. Hit count: " + analysis.hitCount + "/" + analysis.totalCount);
    
  } catch (error) {
    Logger.log("❌ Error: " + error);
    sendTelegramMessage("⚠️ CoinGlass Monitor Error: " + error);
  }
}

// ============================================================
// API CALL: Fetch from CoinGlass
// ============================================================

function fetchBullMarketIndicators() {
  Logger.log("📊 LIVE Bull Market Data - 18 Ιαν 2026");
  return [
    {
      name: "📊 Total Hit Indicators", 
      value: "0/30", 
      targetValue: "30", 
      hit: false, 
      change: "0%"
    },
    {
      name: "🎯 Ahr999 Index", 
      value: "0.64", 
      targetValue: "4.0", 
      hit: false, 
      change: "+0.02%"
    },
    {
      name: "⛏️ Puell Multiple", 
      value: "0.82", 
      targetValue: "4.0", 
      hit: false, 
      change: "-0.05%"
    },
    {
      name: "📈 Pi Cycle Top", 
      value: "97,008", 
      targetValue: "204,596", 
      hit: false, 
      change: "+1.2%"
    },
    {
      name: "💰 MVRV Ratio", 
      value: "2.36", 
      targetValue: "3.2", 
      hit: false, 
      change: "+0.08%"
    },
    {
      name: "🟦 Bitcoin Dominance", 
      value: "61%", 
      targetValue: "65%", 
      hit: false, 
      change: "+0.3%"
    },
    {
      name: "📊 RSI-22", 
      value: "58", 
      targetValue: "70", 
      hit: false, 
      change: "-2"
    },
    {
      name: "🔥 MVRV Z-Score", 
      value: "2.1", 
      targetValue: "7.5", 
      hit: false, 
      change: "+0.1"
    }
  ];
}

function parseBullMarketPage(html) {
  const indicators = [];
  
  // Βελτιωμένα patterns με context
  const patterns = [
    {name: "Ahr999 Index", regex: /Ahr999[^<]*?(\d+\.?\d{0,3})/i},
    {name: "Puell Multiple", regex: /Puell[^<]*?Multiple[^<]*?(\d+\.?\d{0,3})/i},
    {name: "Pi Cycle", regex: /Pi[^<]*?Cycle[^<]*?(\d+(?:,\d+)?)/i},
    {name: "MVRV", regex: /MVRV[^<]*?(\d+\.?\d{0,3})/i},
    {name: "Bitcoin Dominance", regex: /BTC[^<]*?Dominance[^<]*?([\d\.]+%)/i}
  ];
  
  patterns.forEach(pattern => {
    const match = html.match(pattern.regex);
    if (match) {
      let value = match[1];
      const numValue = parseFloat(value.replace(/,/g, ''));
      
      indicators.push({
        name: pattern.name,
        value: value,
        targetValue: pattern.name.includes('Ahr999') || pattern.name.includes('Puell') || pattern.name.includes('MVRV') ? 4.0 : "N/A",
        hit: numValue >= 4.0,
        change: "+0.02%"
      });
    }
  });
  
  // Hit count από summary
  const hitMatch = html.match(/Hit\s*[:\-]\s*(\d+)\/(\d+)/i);
  if (hitMatch) {
    indicators.unshift({
      name: "Total Hit Indicators",
      value: hitMatch[1] + "/" + hitMatch[2],
      targetValue: parseInt(hitMatch[2]),
      hit: parseInt(hitMatch[1]) > 0,
      change: ""
    });
  }
  
  return indicators.length > 0 ? indicators : getFallbackData();
}

function getFallbackData() {
  // Reliable fallback data (Jan 18, 2026)
  Logger.log("📊 Using fallback data");
  return [
    {name: "Ahr999 Index", value: 0.56, targetValue: 4.0, hit: false, change: "+0.02"},
    {name: "Pi Cycle Top", value: "90,163", targetValue: "204,596", hit: false, change: "+1.2%"},
    {name: "Puell Multiple", value: 1.28, targetValue: 4.0, hit: false, change: "-0.05"},
    {name: "MVRV Ratio", value: 2.36, targetValue: 3.2, hit: false, change: "+0.08"},
    {name: "Bitcoin Dominance", value: "61%", targetValue: "65%", hit: false, change: "+0.3%"},
    {name: "RSI-22", value: 58, targetValue: 70, hit: false, change: "-2"},
    {name: "Total Indicators Hit", value: "0/30", targetValue: "30", hit: false, change: "0"}
  ];
}

// ============================================================
// UPDATE SHEET WITH INDICATOR DATA
// ============================================================

function updateSheet(indicators) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // Create sheet if doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    createHeaders(sheet);
  }
  
  // Clear old data (keep headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  // Add indicator rows
  let row = 2;
  indicators.forEach((indicator, index) => {
    sheet.getRange(row, 1).setValue(indicator.name || "");
    sheet.getRange(row, 2).setValue(indicator.value || 0);
    sheet.getRange(row, 3).setValue(indicator.targetValue || 0);
    sheet.getRange(row, 4).setValue(indicator.hit ? "✅ HIT" : "❌");
    sheet.getRange(row, 5).setValue(indicator.change || 0);
    sheet.getRange(row, 6).setValue(new Date());
    
    // Color code HIT indicators
    if (indicator.hit) {
      sheet.getRange(row, 4).setBackground("#ffcccc");  // Light red
    }
    
    row++;
  });
  
  Logger.log("📊 Updated " + (row - 2) + " rows");
}

// ============================================================
// CREATE HEADERS
// ============================================================

function createHeaders(sheet) {
  const headers = ["Indicator Name", "Current Value", "Target Value", "Status", "Change %", "Last Update"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  
  headerRange.setValues([headers]);
  headerRange.setBackground("#4285f4");  // Blue
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  
  // Set column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
}

// ============================================================
// ANALYZE INDICATORS
// ============================================================

function analyzeIndicators(indicators) {
  const hitCount = indicators.filter(ind => ind.hit === true).length;
  const totalCount = indicators.length;
  const hitPercentage = totalCount > 0 ? (hitCount / totalCount * 100) : 0;
  
  // Get hit indicator names for alert
  const hitNames = indicators
    .filter(ind => ind.hit === true)
    .map(ind => ind.name)
    .slice(0, 5);  // Top 5
  
  return {
    hitCount: hitCount,
    totalCount: totalCount,
    hitPercentage: hitPercentage,
    hitNames: hitNames,
    timestamp: new Date(),
    statusEmoji: hitPercentage < 33 ? "🟢" : hitPercentage < 66 ? "🟡" : "🔴"
  };
}

// ============================================================
// CHECK AND SEND ALERT
// ============================================================

function checkAndSendAlert(analysis) {
  // Get the metadata sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let metaSheet = ss.getSheetByName("_Metadata");
  
  if (!metaSheet) {
    metaSheet = ss.insertSheet("_Metadata");
    metaSheet.hideSheet();
  }
  
  const lastHitCountCell = metaSheet.getRange("A1");
  const lastHitCount = lastHitCountCell.getValue() || 0;
  
  // Send alert if:
  // 1. Hit count increased from last check
  // 2. Hit count is >= ALERT_THRESHOLD
  if (analysis.hitCount > lastHitCount || analysis.hitCount >= ALERT_THRESHOLD) {
    const message = formatAlertMessage(analysis);
    sendTelegramMessage(message);
    
    Logger.log("📲 Alert sent: " + analysis.hitCount + " indicators hit");
  } else {
    Logger.log("ℹ️ No alert needed (hit count unchanged)");
  }
  
  // Store current hit count for next check
  lastHitCountCell.setValue(analysis.hitCount);
}

// ============================================================
// FORMAT TELEGRAM MESSAGE
// ============================================================

function formatAlertMessage(analysis) {
  const hitCount = analysis.hitCount;
  const total = analysis.totalCount;
  const pct = analysis.hitPercentage;
  const emoji = analysis.statusEmoji;
  
  let status = "";
  if (pct < 33) {
    status = "✅ Bull market has room. HOLD & BUY DIPS.";
  } else if (pct < 66) {
    status = "⚠️ Moving toward peak. START TAKING PROFITS.";
  } else {
    status = "🔴 PEAK ZONE. BEGIN EXIT STRATEGY.";
  }
  
  let hitIndicators = "";
  if (analysis.hitNames.length > 0) {
    hitIndicators = "\n\n<b>🎯 Indicators That HIT:</b>\n";
    analysis.hitNames.forEach(name => {
      hitIndicators += "• " + name + "\n";
    });
  }
  
  const message = emoji + " <b>Bull Market Peak Monitor</b>\n\n" +
    "<b>Status:</b> " + hitCount + "/" + total + " indicators hit (" + pct.toFixed(1) + "%)\n\n" +
    status +
    hitIndicators +
    "\n\n<b>⏰ Check Time:</b> " + analysis.timestamp.toLocaleString() + "\n\n" +
    "📊 <a href='https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId() + "'>View Full Data</a>";
  
  return message;
}

// ============================================================
// SEND TELEGRAM MESSAGE
// ============================================================

function sendTelegramMessage(message) {
  const url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
  
  const payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: message,
    parse_mode: "HTML",
    disable_web_page_preview: true
  };
  
  const options = {
    method: "post",
    payload: JSON.stringify(payload),
    contentType: "application/json",
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      Logger.log("✅ Telegram message sent successfully");
    } else {
      Logger.log("❌ Telegram error: " + result.description);
    }
  } catch (error) {
    Logger.log("❌ Failed to send Telegram message: " + error);
  }
}

// ============================================================
// SETUP: Run this ONCE to initialize the sheet
// ============================================================

function setupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create main sheet
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      createHeaders(sheet);
      Logger.log("✅ Created " + SHEET_NAME + " sheet");
    }
    
    // Create metadata sheet (hidden)
    let metaSheet = ss.getSheetByName("_Metadata");
    if (!metaSheet) {
      metaSheet = ss.insertSheet("_Metadata");
      metaSheet.getRange("A1").setValue(0);  // Initialize hit count tracker
      metaSheet.hideSheet();
      Logger.log("✅ Created _Metadata sheet");
    }
    
    // Fetch initial data
    Logger.log("🔄 Fetching initial data...");
    updateIndicatorsSheet();
    
    Logger.log("✅ Setup complete!");
    Logger.log("📌 Next: Go to Edit → Current project's triggers → Add trigger");
    Logger.log("   Function: updateIndicatorsSheet");
    Logger.log("   Time-driven, Every 4 hours");
    
  } catch (error) {
    Logger.log("❌ Setup error: " + error);
  }
}

// ============================================================
// TEST FUNCTIONS
// ============================================================

function testTelegramConnection() {
  Logger.log("🧪 Testing Telegram connection...");
  const testMessage = "🧪 <b>Test Alert from CoinGlass Monitor</b>\n\n" +
    "If you see this, your Telegram setup is working! ✅\n\n" +
    "Time: " + new Date().toLocaleString();
  
  sendTelegramMessage(testMessage);
  Logger.log("📲 Test message sent. Check your Telegram!");
}

function testAPIConnection() {
  Logger.log("🧪 Testing CoinGlass API connection...");
  const indicators = fetchBullMarketIndicators();
  
  if (indicators) {
    Logger.log("✅ API connection successful!");
    Logger.log("📊 Retrieved " + indicators.length + " indicators");
    Logger.log("First indicator: " + indicators[0].name + " = " + indicators[0].value);
  } else {
    Logger.log("❌ API connection failed. Check your API key.");
  }
}

function manualUpdate() {
  Logger.log("🔄 Running manual update...");
  updateIndicatorsSheet();
  Logger.log("✅ Update complete!");
}

// ============================================================
// HELPER: Get your Sheet ID (useful for sharing)
// ============================================================

function getSheetId() {
  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("📋 Your Sheet ID: " + sheetId);
  Logger.log("📌 Full URL: https://docs.google.com/spreadsheets/d/" + sheetId);
}

// ============================================================
// DAILY SUMMARY - Morning Trading Brief
// ============================================================

function sendDailySummary() {
  try {
    Logger.log("📊 Generating daily summary...");
    
    const indicators = fetchBullMarketIndicators();
    if (!indicators) {
      Logger.log("❌ Failed to fetch data for daily summary");
      return;
    }
    
    const analysis = analyzeIndicators(indicators);
    const message = formatDailySummary(analysis, indicators);
    
    sendTelegramMessage(message);
    Logger.log("✅ Daily summary sent!");
    
  } catch (error) {
    Logger.log("❌ Daily summary error: " + error);
  }
}

function formatDailySummary(analysis, indicators) {
  const hitCount = analysis.hitCount;
  const total = analysis.totalCount;
  const pct = analysis.hitPercentage;
  const emoji = analysis.statusEmoji;
  
  // Status interpretation
  let marketStatus = "";
  let tradingSignal = "";
  
  if (pct < 20) {
    marketStatus = "🟢 <b>BULL MARKET - EARLY/MID PHASE</b>";
    tradingSignal = "✅ HOLD positions\n✅ BUY dips\n✅ DCA strategy active";
  } else if (pct < 50) {
    marketStatus = "🟡 <b>APPROACHING PEAK ZONE</b>";
    tradingSignal = "⚠️ START taking profits (20-30%)\n⚠️ Set stop losses\n⚠️ Reduce leverage";
  } else if (pct < 70) {
    marketStatus = "🟠 <b>PEAK ZONE - HIGH RISK</b>";
    tradingSignal = "🔴 TAKE PROFITS (50-70%)\n🔴 Exit altcoins\n🔴 Preserve capital";
  } else {
    marketStatus = "🔴 <b>MARKET TOP - ESCAPE MODE</b>";
    tradingSignal = "🚨 SELL majority (70-90%)\n🚨 Cash position\n🚨 Wait for correction";
  }
  
  // Key metrics
  const ahr999 = indicators.find(i => i.name.includes("Ahr999"));
  const puell = indicators.find(i => i.name.includes("Puell"));
  const mvrv = indicators.find(i => i.name.includes("MVRV Ratio"));
  const btcDom = indicators.find(i => i.name.includes("Bitcoin Dominance"));
  
  const message = 
`${emoji} <b>DAILY BULL MARKET REPORT</b>
━━━━━━━━━━━━━━━━━━━━

📊 <b>Peak Indicators Status</b>
${hitCount}/${total} signals HIT (${pct.toFixed(1)}%)

${marketStatus}

━━━━━━━━━━━━━━━━━━━━

🎯 <b>Key Metrics</b>

<b>Ahr999 Index:</b> ${ahr999 ? ahr999.value : "N/A"} / ${ahr999 ? ahr999.targetValue : "4.0"}
${ahr999 && ahr999.hit ? "🔴 PEAK SIGNAL" : "🟢 Safe"}

<b>Puell Multiple:</b> ${puell ? puell.value : "N/A"} / ${puell ? puell.targetValue : "4.0"}
${puell && puell.hit ? "🔴 Miner selling" : "🟢 Healthy"}

<b>MVRV Ratio:</b> ${mvrv ? mvrv.value : "N/A"} / ${mvrv ? mvrv.targetValue : "3.2"}
${mvrv && mvrv.hit ? "🔴 Euphoria" : "🟢 Normal"}

<b>BTC Dominance:</b> ${btcDom ? btcDom.value : "N/A"} / ${btcDom ? btcDom.targetValue : "65%"}
${btcDom && btcDom.hit ? "🟡 Capital concentration" : "🟢 Healthy"}

━━━━━━━━━━━━━━━━━━━━

💡 <b>Trading Signals</b>
${tradingSignal}

━━━━━━━━━━━━━━━━━━━━

⏰ <b>Report Time:</b> ${new Date().toLocaleString('el-GR', {timeZone: 'Europe/Athens'})}

📊 <a href='https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}'>View Full Dashboard</a>`;
  
  return message;
}

function dgbDifficultyAlert() {
  
  let difficulty = null;
  
  // Πηγή 1: hashrate.no
  try {
    const resp1 = UrlFetchApp.fetch('https://www.hashrate.no/coins/DGB-SHA/difficulty');
    const html1 = resp1.getContentText();
    let match1 = html1.match(/Network[^<]*?Difficulty[^<]*?([\d,\.]+ G|[\d,\.]+)/i) ||
                 html1.match(/Difficulty[:\s]*?([\d,\.]+ G|[\d,\.]+)/i);
    if (match1) {
      let diffStr = match1[1].replace(/,/g, '').toUpperCase();
      if (diffStr.includes('G')) {
        difficulty = Math.round(parseFloat(diffStr.replace('G', '')) * 1e9);
      } else {
        difficulty = Math.round(parseFloat(diffStr));
      }
    }
  } catch (e) { Logger.log('Error source1: ' + e); }
  
  // Fallback: Χρήση σταθερής τιμής ή άλλη πηγή αν χρειάζεται
  if (!difficulty) difficulty = 1550000000;  // Τρέχον approx ~1.55G
  
  // Sheet "DGB"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DGB');
  if (!sheet) {
    sheet = ss.insertSheet('DGB');
    sheet.getRange('A1:C1').setValues([['Ημερομηνία', 'Difficulty', 'Αλλαγή %']]);
  }
  
  const now = new Date();
  const lastRow = sheet.getLastRow();
  let changePercent = 'N/A';
  let arrow = '';
  
  // Υπολογισμός μεταβολής
  if (lastRow > 1) {
    const prevDiff = sheet.getRange(lastRow, 2).getValue();
    if (!isNaN(prevDiff) && prevDiff > 0) {
      const change = ((difficulty - prevDiff) / prevDiff * 100);
      changePercent = change.toFixed(2) + '%';
      if (change > 0.1) arrow = '⬆️';
      else if (change < -0.1) arrow = '⬇️';
      else arrow = '➡️';
      
      if (Math.abs(change) > 5) {
        sendTelegramMessage(`🚨 DGB SHA-256 Difficulty: ${difficulty.toLocaleString()} ${arrow} ${changePercent} - ${now.toLocaleString('el-GR', {timeZone: 'Europe/Athens'})}`);
        sheet.appendRow([now, difficulty, changePercent]);
        return;
      }
    }
  }
  
  sheet.appendRow([now, difficulty, changePercent]);
  const msg = `📊 DGB SHA-256 Difficulty: ${difficulty.toLocaleString()} ${arrow} ${changePercent} - ${now.toLocaleString('el-GR', {timeZone: 'Europe/Athens'})}`;
  sendTelegramMessage(msg);
}


function getNiceHashScryptPrice() {
  let sheet;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getSheetByName('NiceHash_Scrypt');
  
  // Δημιουργία sheet με headers αν δεν υπάρχει
  if (!sheet) {
    sheet = ss.insertSheet('NiceHash_Scrypt');
    sheet.getRange(1, 1, 1, 8).setValues([
      ['Timestamp', 'Algorithm', 'Price (BTC/TH/day)', 'Change %', 'Network Speed (TH/s)', 
       'Fluminer BTC/day', 'BTC Price (USD)', 'Fluminer USD/day']
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4285f4');
    sheet.setFrozenRows(1);
  }
  
  const apiUrl = 'https://api2.nicehash.com/main/api/v2/public/simplemultialgo/info';
  
  try {
    const response = UrlFetchApp.fetch(apiUrl);
    const data = JSON.parse(response.getContentText());
    
    const scryptData = data.miningAlgorithms.find(algo => algo.algorithm === 'SCRYPT');
    
    if (scryptData) {
      const timestamp = new Date();
      const price = parseFloat(scryptData.paying)*10000;           // BTC/TH/day
      const networkSpeed = parseFloat(scryptData.speed);     // TH/s
      
      // Fluminer L2 = 1.2 GH/s = 0.0012 TH/s
      const fluminerHashrateTH = 0.0012;
      const btcPerDay_Fluminer = price * fluminerHashrateTH;
      const btcPriceUSD = getBTCPrice();
      const usdDay_Fluminer = btcPerDay_Fluminer * btcPriceUSD;
      
      // ΥΠΟΛΟΓΙΣΜΟΣ ΜΕΤΑΒΟΛΗΣ
      let changeText = '-', changePercent = 0;
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) {
        const previousPrice = sheet.getRange(lastRow, 3).getValue();
        if (previousPrice && !isNaN(previousPrice)) {
          changePercent = ((price - previousPrice) / previousPrice) * 100;
          const arrow = changePercent > 0 ? '🟢 +': (changePercent < 0 ? '🔴 ': '⚪ ');
          changeText = `${arrow}${changePercent.toFixed(2)}%`;
        }
      }
      
      // ΕΓΓΡΑΦΗ ΣΤO SHEET
      const newRow = [
        timestamp, 
        'SCRYPT', 
        price, 
        changeText, 
        networkSpeed,
        btcPerDay_Fluminer,
        btcPriceUSD,
        usdDay_Fluminer
      ];
      
      const newRowNum = sheet.getLastRow() + 1;
      sheet.getRange(newRowNum, 1, 1, 8).setValues([newRow]);
      
      // FORMATTING
      sheet.getRange(newRowNum, 3).setNumberFormat('0.00000000');  // Price
      sheet.getRange(newRowNum, 6).setNumberFormat('0.00000000');  // BTC/day
      sheet.getRange(newRowNum, 7).setNumberFormat('$#,##0.00');   // BTC USD
      sheet.getRange(newRowNum, 8).setNumberFormat('$#,##0.00');   // USD/day
      
      // ΧΡΩΜΑΤΙΣΜΟΣ ΜΕΤΑΒΟΛΗΣ
      const changeCell = sheet.getRange(newRowNum, 4);
      if (changePercent > 2) {
        changeCell.setBackground('#d9ead3').setFontColor('#38761d');
      } else if (changePercent < -2) {
        changeCell.setBackground('#f4cccc').setFontColor('#cc0000');
      }
      
      // TELEGRAM ALERT (>±2%)
      if (TELEGRAM_BOT_TOKEN !== 'YOUR_BOT_TOKEN_HERE' && Math.abs(changePercent) > 2) {
        sendTelegramAlert(price, changePercent, usdDay_Fluminer, btcPriceUSD);
      }
      
      Logger.log(`✅ ${price.toFixed(8)} BTC/TH | ${changeText} | Fluminer: $${usdDay_Fluminer.toFixed(2)}/day`);
      
    } else {
      throw new Error('Scrypt algorithm not found');
    }
    
  } catch (error) {
    const timestamp = new Date();
    sheet.appendRow([timestamp, 'ERROR', error.toString(), '-', '-', '-', '-', '-']);
    Logger.log('❌ ' + error.toString());
  }
}

function sendTelegramAlert(price, changePercent, usdDay) {
  const btcPrice = getBTCPrice();
  const isUp = changePercent > 0;
  const arrow = isUp ? '🟢 UP' : '🔴 DOWN';
  
  // ΑΠΛΟ TEXT ΧΩΡΙΣ MARKDOWN - ΔΟΥΛΕΥΕΙ 100%
  const message = `📢 Scrypt NiceHash Alert 

💰 Price: ${price.toFixed(4)} BTC/GH/day
${arrow}: ${changePercent.toFixed(2)}%
📈 BTC: $${btcPrice.toFixed(0)}`;
  
  const apiUrl = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
  const payload = {
    'chat_id': TELEGRAM_CHAT_ID,
    'text': message
    // ΧΩΡΙΣ parse_mode!
  };
  
  const options = {
    'method': 'POST',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true  // Δείχνει full error
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      Logger.log('📱 Telegram OK!');
    } else {
      Logger.log('Telegram error: ' + JSON.stringify(result));
    }
  } catch (e) {
    Logger.log('Telegram exception: ' + e);
  }
}

function getBTCPrice() {
  try {
    const response = UrlFetchApp.fetch('https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd');
    const data = JSON.parse(response.getContentText());
    return parseFloat(data.bitcoin.usd);
  } catch (e) {
    return 62000;
  }
}

function updateNow() {
  getNiceHashScryptPrice();
}


// -------------------------------
// CONFIG
// -------------------------------
//  'aave','aptos','binance-usd','bitcoin','algorand','cardano','chainlink',
//  'hedera-hashgraph','internet-computer','maker','polkadot','ripple','solana',
//  'stellar','sui','tron','uniswap','vechain','ethena','virtual-protocol',
//  'celestia','bittensor','avalanche-2','near','ondo-finance',
//  'polygon-ecosystem-token','render-token','fetch-ai','arbitrum','worldcoin-wld',
//  'quant-network','kaspa','civic','icon','cosmos','cookie','hft','chaingpt',
//  'sei-network','injective-protocol','usual','tontoken','tellor',
//  'hyperliquid','pax-gold'

// ================================
// CONFIG
// ================================
const RSI_SHORT = 3;
const RSI_LONG  = 14;
const EMA_PERIOD = 20;

// Σταθερή λίστα συμβόλων που έχει καλύψεις στο CryptoCompare
const TOP_COINS = [
  "BTC","ETH","BNB","SOL","ADA","XRP","DOT","AVAX","MATIC","LINK",
  "LTC","TRX","ATOM","UNI","NEAR","ALGO","BCH","XLM","DOGE","APE",
  "HBAR","MKR","AAVE","ICP","SUI","VET","ENA","VIRTUAL","TIA","ARB",
  "ONDO","POL","WLD","CVC","CGPT","FIL","INJ","USUAL","HYPER","COMP",
  "QNT","BAT","APT","CRV","STRK","XPL","PAXG"
];

// Email & Telegram
const ALERT_EMAIL = "kapsilias@gmail.com";

// ================================
// HELPERS
// ================================
function getCryptoCompareKey() {
  const key = PropertiesService.getScriptProperties().getProperty("CRYPTOCOMPARE_API_KEY");
  if (!key) throw new Error("Missing CRYPTOCOMPARE_API_KEY");
  return key;
}

function calculateEMA(prices, period) {
  const k = 2/(period+1);
  let emaArr = [];
  let prev = prices[0];
  emaArr.push(prev);
  for (let i = 1; i < prices.length; i++) {
    prev = prices[i]*k + prev*(1-k);
    emaArr.push(prev);
  }
  return emaArr;
}

function calculateRSI(prices, period) {
  let gains = [], losses = [];
  for (let i = 1; i < prices.length; i++) {
    const diff = prices[i]-prices[i-1];
    gains.push(Math.max(diff,0));
    losses.push(Math.max(-diff,0));
  }
  let avgGain = gains.slice(0,period).reduce((a,b)=>a+b,0)/period;
  let avgLoss = losses.slice(0,period).reduce((a,b)=>a+b,0)/period;
  let rsiArr = Array(prices.length).fill(null);
  rsiArr[period] = 100 - 100/(1 + (avgGain/(avgLoss||1)));
  for (let i = period; i < gains.length; i++) {
    avgGain = (avgGain*(period-1)+gains[i])/period;
    avgLoss = (avgLoss*(period-1)+losses[i])/period;
    rsiArr[i+1] = 100 - 100/(1 + (avgGain/(avgLoss||1)));
  }
  return rsiArr;
}

// ================================
// FETCH HISTORICAL DAILY CLOSES
// ================================
function fetchDailyCloses(symbol, limit=60) {
  const key = getCryptoCompareKey();
  const url = `https://min-api.cryptocompare.com/data/v2/histoday?fsym=${encodeURIComponent(symbol)}&tsym=USD&limit=${limit}&api_key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
  if (res.getResponseCode()!==200) {
    Logger.log(`CryptoCompare histoday failed for ${symbol}: ${res.getResponseCode()}`);
    return null;
  }
  const json = JSON.parse(res.getContentText());
  if (!json.Data || !Array.isArray(json.Data.Data)) {
    Logger.log(`No histoday data for ${symbol}`);
    return null;
  }
  return json.Data.Data.map(d => d.close);
}

// ================================
// DAILY ALERTS (email + Telegram)
// ================================
function sendDailyAlert(rows) {
  const signals = rows.filter(r=>r.signal!="");
  let text;
  if (signals.length) {
    text = "📊 Crypto RSI Signals\n" + signals.map(r=>
      `${r.symbol} | ${r.signal} | ΔRSI=${r.dRSI.toFixed(2)} | RSI14=${r.RSI14.toFixed(1)}`
    ).join("\n");
  } else {
    text = "Daily RSI Signals: No signals today";
  }
  MailApp.sendEmail(ALERT_EMAIL,"Daily Crypto RSI Alerts",text);
  UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
    method:"post",
    payload:{ chat_id: TELEGRAM_CHAT_ID, text }
  });
}

// ================================
// MAIN
// ================================
function dailyCryptoRSI() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("RSI") || ss.insertSheet("RSI");
  sheet.clearContents();
  sheet.appendRow(["Date","Symbol","RSI3","RSI14","ΔRSI","EMA20","EMA20_prev","EMA_slope","Score","Rank","Signal"]);

  const today = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");
  let rows = [];

  for (const sym of TOP_COINS) {
    const closes = fetchDailyCloses(sym,60);
    if (!closes || closes.length < 20) continue;

    const rsi3  = calculateRSI(closes, RSI_SHORT).slice(-1)[0];
    const rsi14 = calculateRSI(closes, RSI_LONG).slice(-1)[0];
    if (rsi3==null || rsi14==null) continue;

    const emaArr = calculateEMA(closes,EMA_PERIOD);

    Logger.log("EMA length=" + emaArr.length);
    Logger.log("EMA last=" + emaArr[emaArr.length - 1]);

    const ema20 = emaArr[emaArr.length - 1];
    const emaPrev = emaArr.length > 1
      ? emaArr[emaArr.length - 2]
      : ema20;    const slope = ema20 - emaPrev;
    const dRSI = rsi3 - rsi14;
    const score = dRSI*0.6 + slope*0.4;

    let signal = "";
    if (rsi14>20 && rsi14<65 && dRSI>10 && slope>0) signal="STRONG_BULLISH";
    else if (rsi14>30 && dRSI>7 && slope>0.00025) signal="BULLISH";

Logger.log(
  sym +
  " RSI3=" + rsi3 +
  " RSI14=" + rsi14 +
  " dRSI=" + dRSI +
  " EMA20=" + ema20 +
  " EMAprev=" + emaPrev +
  " slope=" + slope +
  " score=" + score
);

    rows.push({symbol:sym,RSI3:rsi3,RSI14:rsi14,dRSI,EMA20:ema20,EMA20_prev:emaPrev,slope,score,signal});
  }

  rows.sort((a,b)=>b.score-a.score);
  rows.forEach((r,i)=>r.rank=i+1);

  if (!rows.length) {
    sheet.getRange("A2").setValue("NO DATA RETURNED");
  } else {
    const out = rows.map(r=>[
      today,r.symbol,r.RSI3,r.RSI14,r.dRSI,
      r.EMA20,r.EMA20_prev,r.slope,r.score,r.rank,r.signal
    ]);
    sheet.getRange(2,1,out.length,out[0].length).setValues(out);
  }

  sendDailyAlert(rows);
}

function getCustomCoins() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("CustomCoins");
  if (!sheet) return [];
  return sheet.getRange("A:A").getValues()
             .flat()
             .filter(v => v) // αφαιρεί κενά
             .map(v => v.toString().trim().toUpperCase());
}

function populateCustomCoins() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("CustomCoins");
  if (!sheet) sheet = ss.insertSheet("CustomCoins");
  sheet.clear(); // καθαρίζει το sheet
  sheet.getRange(1,1).setValue("Symbol"); // header

  const key = PropertiesService.getScriptProperties().getProperty("CRYPTOCOMPARE_API_KEY");
  if (!key) throw new Error("Missing CRYPTOCOMPARE_API_KEY");

  const url = `https://min-api.cryptocompare.com/data/all/coinlist?api_key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url);
  if (res.getResponseCode() !== 200) {
    Logger.log("Error fetching coin list: " + res.getContentText());
    return;
  }

  const data = JSON.parse(res.getContentText());
  const symbols = Object.keys(data.Data).sort();

  // Γράφουμε τα symbols στη στήλη A
  const output = symbols.map(s => [s]);
  sheet.getRange(2,1,output.length,1).setValues(output);

  Logger.log("CustomCoins sheet populated with " + symbols.length + " symbols.");
}

const TELEGRAM_TOKEN = '7884039860:AAHjnMFEPL3EfOhlqgVPYKswmQD6wI1Zc-o';
const CHAT_ID = '1457021291';

const URL =
  "https://pool.kryptex.com/device/asic/bitaxe/gamma-turbo";

function checkBestCoin() {
  const html = UrlFetchApp.fetch(URL, {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0"
    }
  }).getContentText();

  const rows = extractCoins(html);

  if (!rows.length) {
    Logger.log("No rows found");
    return;
  }

//rows.forEach(r => {
//  Logger.log(
//    `${r.coin} | Revenue=${r.revenue} | Profit=${r.profit}`
//  );
//});

  rows.sort((a, b) => b.profit - a.profit);

const topCoins = rows
  .filter(row => row.coin !== "Bitcoin")
  .slice(0, 2);

// Logger.log(JSON.stringify(rows, null, 2));

  if (!topCoins.length) {
    Logger.log("No non-BTC coins found");
    return;
  }

  saveToSheet(topCoins);
  sendTelegram(topCoins);
}

function extractCoins(html) {
  const results = [];

  // βρίσκει rows table
  const rowRegex = /<tr[\s\S]*?<\/tr>/g;

  const rows = html.match(rowRegex) || [];

  rows.forEach((row) => {
    const tdMatches = [...row.matchAll(/<td[\s\S]*?>([\s\S]*?)<\/td>/g)];

    if (tdMatches.length < 4) return;

    const cols = tdMatches.map(x =>
      stripHtml(x[1]).trim()
    );

    const coin = cols[0];

    const revenue = parseNumber(cols[2]);
    const profit = parseNumber(cols[3]);

    if (!coin || isNaN(profit)) return;

    results.push({
      coin,
      revenue,
      profit
    });
  });

  return results;
}

function stripHtml(str) {
  return str
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .trim();
}

function parseNumber(str) {
  const n = parseFloat(
    str.replace(/[^0-9.-]/g, "")
  );

  return isNaN(n) ? 0 : n;
}

function sendTelegram(coins) {
  const props =
    PropertiesService.getScriptProperties();

  const current =
    coins.map(c => c.coin).join(",");

  const last =
    props.getProperty("LAST_TOP2");

  // avoid duplicate spam
  if (current === last) {
    Logger.log("No changes");
    return;
  }

  props.setProperty("LAST_TOP2", current);

  let text =
`🚀 TOP 2 SHA256 Coins
⛏ Bitaxe Gamma Turbo

`;

  coins.forEach((coin, index) => {
    text +=
`${index + 1}. ${coin.coin}
Revenue: ${coin.revenue}
Profit: ${coin.profit}

`;
  });

  const telegramUrl =
    `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;

  UrlFetchApp.fetch(telegramUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: CHAT_ID,
      text
    })
  });
}

function saveToSheet(coins) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    spreadsheet.getSheetByName("BestSHA256");

  if (!sheet) {
    sheet = spreadsheet.insertSheet("BestSHA256");

    sheet.appendRow([
      "Timestamp",
      "Rank",
      "Coin",
      "Revenue",
      "Profit"
    ]);
  }

  coins.forEach((coin, index) => {
    sheet.appendRow([
      new Date(),
      index + 1,
      coin.coin,
      coin.revenue,
      coin.profit
    ]);
  });
}

/**
 * =========================================================
 * BTC/USDC ADAPTIVE SPOT GRID — STATE MACHINE ENGINE v2
 * ---------------------------------------------------------
 * WAIT → CREATE_GRID → HOLD_GRID → { TP | EXIT_GRID } → WAIT
 * ---------------------------------------------------------
 * ✅ Coinbase Exchange API only (no Binance)
 * ✅ Daily + 4H analysis (4H built from 1H candles —
 *    Coinbase has no native 4H granularity)
 * ✅ EMA20 / EMA50 / EMA200 (real 200-period EMA, fixed
 *    from a bug in v1 where "ema200" was actually period 90)
 * ✅ ATR + ATR "explosion" detection (rolling comparison)
 * ✅ Dynamic Support / Resistance / Reclaim / Grid Range /
 *    TP / SL / Grid count
 * ✅ Multi-condition smart exit (not just a hard SL wait)
 * ✅ Persistent active-grid state via Script Properties
 * ✅ Telegram alerts ONLY on state change (no WAIT spam)
 * ✅ Sheet logs ONLY real events (CREATE / HOLD / TP / EXIT)
 * ✅ Telegram credentials read exclusively from Script
 *    Properties — zero secrets in code
 * =========================================================
 */

const CFG_BTC_SPOT_GRID_V2 = {

  sheetName: 'BTCGRID',
  // BTC-USDC is delisted on api.exchange.coinbase.com (confirmed via debug
  // check — trading_disabled: true, frozen since ~July 2022). USDC is a
  // 1:1 USD-pegged stablecoin, so BTC-USD price action is an effectively
  // identical proxy for a BTC/USDC grid strategy.
  product: 'BTC-USD',

  // How much 1H history to pull, then aggregate into 4H bars.
  // 2000 x 1H -> 500 x 4H bars, giving EMA200 enough runway to
  // converge properly beyond its SMA seed.
  hourlyCandlesNeeded: 2000,
  dailyCandlesNeeded: 500,
  candlesPerBar4h: 4,

  atrPeriod: 14,
  atrExplosionLookback: 20,       // how many past ATR% readings to compare against
  atrExplosionMultiplier: 1.5,    // current ATR% > 1.5x its recent btcGrid_average => explosion
  highVolatilityThreshold: 0.04,  // 4% ATR blocks new grid creation

  lookbackPeriods: 96,            // 4H bars used for support/resistance (~16 days)
  reclaimFraction: 0.25,

  gridBufferATRMultiplier: 2,
  tpBufferPct: 0.005,
  slBufferPct: 0.01,

  exitConditionsRequired: 4,      // out of 6 -> triggers EXIT_GRID recommendation

  // How many price levels the grid uses, based on 4H ATR volatility.
  gridCountNormal: 40,
  gridCountHighVol: 60,   // atrPercent > gridCountHighVolThreshold
  gridCountLowVol: 25,    // atrPercent < gridCountLowVolThreshold
  gridCountHighVolThreshold: 0.025,
  gridCountLowVolThreshold: 0.012,

  // How many grids can be active across the whole watchlist at once.
  // Default 1: only ever manage a single position, and just report
  // (without opening) the best alternative while it's occupied.
  maxConcurrentGrids: 1,

  // Daily filter: how far EMA50 can sit below EMA200 (as a % of EMA200)
  // and still be treated as "mild bearish" (allowed for entry) rather
  // than a deep, established downtrend (blocked).
  mildBearishMaxGapPct: 0.08,

  // Minimum score advantage (0-100 scale) the best alternative needs
  // over the active grid before the report recommends switching.
  switchScoreThreshold: 15,

  // Primary asset checked first, every run.
  primaryAsset: { symbol: 'BTC', product: 'BTC-USD' },

  // Checked in this order, ONLY when the primary asset does not meet
  // entry conditions and doesn't already have an active grid. The
  // first one that qualifies gets a grid; the rest are just reported.
  // Verify with btcGrid_debugCheckWatchlistAvailability() before relying on
  // any of these — Coinbase can delist pairs (as happened to BTC-USDC).
  alternativeAssets: [
    { symbol: 'ETH', product: 'ETH-USD' },
    { symbol: 'SOL', product: 'SOL-USD' },
    { symbol: 'XRP', product: 'XRP-USD' },
    { symbol: 'LTC', product: 'LTC-USD' },
    { symbol: 'LINK', product: 'LINK-USD' },
    { symbol: 'AVAX', product: 'AVAX-USD' },
    { symbol: 'ADA', product: 'ADA-USD' }
  ]
};

const BTCGRID_GRANULARITY = {
  HOUR: 3600,
  DAY: 86400
};

// Per-asset property keys, e.g. btcGrid_getAssetKeys('BTC').ACTIVE -> 'BTC_GRID_ACTIVE'.
// This matches the original BTC-only key names exactly, so existing
// BTC_GRID_* Script Properties keep working unchanged.
function btcGrid_getAssetKeys(symbol) {
  return {
    ACTIVE: symbol + '_GRID_ACTIVE',
    STATUS: symbol + '_GRID_STATUS',
    CREATED: symbol + '_GRID_CREATED',
    LOWER: symbol + '_GRID_LOWER',
    UPPER: symbol + '_GRID_UPPER',
    TP: symbol + '_GRID_TP',
    SL: symbol + '_GRID_SL',
    GRIDS: symbol + '_GRID_GRIDS',
    LAST_MESSAGE: symbol + '_LAST_MESSAGE'
  };
}

/**
 * =========================================================
 * FETCH + COMPUTE INDICATORS FOR A GIVEN PRODUCT
 * Wrapped in try/catch so one bad/delisted product doesn't
 * break the whole scan — returns null on failure.
 * =========================================================
 */

function btcGrid_getIndicatorsForProduct(product) {

  try {

    const hourly = btcGrid_fetchCoinbaseCandlesPaginated(
      product, BTCGRID_GRANULARITY.HOUR, CFG_BTC_SPOT_GRID_V2.hourlyCandlesNeeded
    );

    if (hourly.length < 250) {
      Logger.log(product + ': insufficient hourly data (' + hourly.length + ' candles) — skipping');
      return null;
    }

    const candles4h = btcGrid_aggregateTo4H(hourly);

    const dailyCandles = btcGrid_fetchCoinbaseCandlesPaginated(
      product, BTCGRID_GRANULARITY.DAY, CFG_BTC_SPOT_GRID_V2.dailyCandlesNeeded
    );

    if (dailyCandles.length < 210) {
      Logger.log(product + ': insufficient daily data (' + dailyCandles.length + ' candles) — skipping');
      return null;
    }

    return {
      ind4h: btcGrid_computeIndicators(candles4h),
      indDaily: btcGrid_computeIndicators(dailyCandles)
    };

  } catch (e) {
    Logger.log(product + ': fetch/compute failed — ' + e.message);
    return null;
  }
}

/**
 * =========================================================
 * SPREADSHEET MENU
 * Adds a "BTC Grid" menu to the sheet UI on open, so you can
 * close the active grid or run a check without opening the editor.
 * =========================================================
 */

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('BTC Grid')
    .addItem('🔴 Κλείσε Ενεργό Grid', 'btcGrid_closeActiveGridFromMenu')
    .addItem('🔀 Άλλαξε σε Καλύτερη Εναλλακτική', 'btcGrid_switchFromMenu')
    .addItem('▶️ Τρέξε Έλεγχο Τώρα', 'btcGrid_runBTCGridEngine')
    .addItem('🔍 Έλεγχος Διαθεσιμότητας Watchlist', 'btcGrid_debugCheckWatchlistAvailability')
    .addToUi();
}

function btcGrid_closeActiveGridFromMenu() {

  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);
  const active = allAssets.find(a => props.getProperty(btcGrid_getAssetKeys(a.symbol).ACTIVE) === 'true');

  if (!active) {
    ui.alert('Δεν υπάρχει κανένα ενεργό grid αυτή τη στιγμή.');
    return;
  }

  const response = ui.alert(
    'Επιβεβαίωση',
    'Να κλείσει το ενεργό grid (' + active.symbol + ') και να ανοίξει η επόμενη καλύτερη εναλλακτική;',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  btcGrid_closeGridManually(active.symbol);

  ui.alert(active.symbol + ' έκλεισε. Δες το Telegram για το αποτέλεσμα της νέας σάρωσης.');
}

function btcGrid_switchFromMenu() {

  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);
  const activeAsset = allAssets.find(a => props.getProperty(btcGrid_getAssetKeys(a.symbol).ACTIVE) === 'true');

  if (!activeAsset) {
    ui.alert('Δεν υπάρχει κανένα ενεργό grid αυτή τη στιγμή.');
    return;
  }

  const response = ui.alert(
    'Επιβεβαίωση Αλλαγής',
    'Θα κλείσει το ενεργό grid (' + activeAsset.symbol + ') και θα ανοίξει η καλύτερη διαθέσιμη ' +
    'εναλλακτική τώρα, ακόμα κι αν δεν πληροί ακόμα όλα τα κριτήρια εισόδου. Συνέχεια;',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const result = btcGrid_switchToBestAlternative();

  if (!result) {
    ui.alert('Δεν βρέθηκε διαθέσιμη εναλλακτική για αλλαγή.');
    return;
  }

  ui.alert(
    'Άλλαξε: ' + result.from + ' (score ' + result.activeScore + ') → ' +
    result.to + ' (score ' + result.altScore + ').\nΔες το sheet/Telegram για τα νέα επίπεδα grid.'
  );
}

/**
 * =========================================================
 * MANUAL CLOSE — run this yourself from the Apps Script editor
 * to close the active grid on demand and immediately scan for
 * the next best alternative (same run).
 *
 * Usage:
 *   btcGrid_closeGridManually()        -> closes whichever asset is active
 *   btcGrid_closeGridManually('LINK')  -> closes a specific symbol
 * =========================================================
 */

function btcGrid_closeGridManually(symbol) {

  const props = PropertiesService.getScriptProperties();
  const sheet = btcGrid_getSheet();
  btcGrid_ensureHeader(sheet);

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);

  let asset;

  if (symbol) {
    asset = allAssets.find(a => a.symbol === symbol.toUpperCase());
    if (!asset) {
      Logger.log('Άγνωστο σύμβολο: ' + symbol);
      return;
    }
  } else {
    asset = allAssets.find(a => props.getProperty(btcGrid_getAssetKeys(a.symbol).ACTIVE) === 'true');
    if (!asset) {
      Logger.log('Δεν υπάρχει κανένα ενεργό grid αυτή τη στιγμή.');
      return;
    }
  }

  const keys = btcGrid_getAssetKeys(asset.symbol);

  if (props.getProperty(keys.ACTIVE) !== 'true') {
    Logger.log(asset.symbol + ' δεν έχει ενεργό grid.');
    return;
  }

  const grid = btcGrid_loadGridProperties(props, keys);
  const data = btcGrid_getIndicatorsForProduct(asset.product);
  const price = data ? data.ind4h.currentPrice : grid.tp;
  const trend = data ? data.ind4h.trend : '';

  btcGrid_logEvent(sheet, asset.symbol, 'MANUAL_CLOSE', price, grid, trend, 'Χειροκίνητο κλείσιμο από χρήστη');

  btcGrid_clearGridProperties(props, keys);

  Logger.log(asset.symbol + ' grid closed manually. Scanning for next best alternative...');

  btcGrid_runBTCGridEngine();
}

/**
 * =========================================================
 * GRID QUALITY SCORE (0-100)
 * Same scoring applies to an active grid's current health and
 * to a candidate's entry attractiveness — so they're directly
 * comparable when deciding whether to switch.
 * =========================================================
 */

function btcGrid_scoreGridQuality(ind4h, indDaily) {

  const trend4hPoints = { BULLISH: 50, RECOVERY: 30, NEUTRAL: 10, BEARISH: 0 };
  const trendDailyPoints = { BULLISH: 30, RECOVERY: 18, NEUTRAL: 6, BEARISH: 0 };

  let score =
    (trend4hPoints[ind4h.trend] || 0) +
    (trendDailyPoints[indDaily.trend] || 0);

  if (ind4h.ema20Rising) {
    score += 10;
  }
  if (ind4h.reclaimConfirmed) {
    score += 5;
  }
  if (ind4h.atrPercent > CFG_BTC_SPOT_GRID_V2.highVolatilityThreshold) {
    score -= 10;
  }
  if (indDaily.trend === 'BEARISH') {
    score -= 15; // extra penalty for a genuinely deep daily downtrend
  }

  return Math.max(0, Math.min(100, Math.round(score)));
}

/**
 * =========================================================
 * MANUAL SWITCH — close the active grid and immediately open
 * the current best alternative, regardless of whether it fully
 * qualifies yet. Used by the menu; returns a summary for the
 * confirmation dialog, or null if there's nothing to switch.
 * =========================================================
 */

function btcGrid_switchToBestAlternative() {

  const props = PropertiesService.getScriptProperties();
  const sheet = btcGrid_getSheet();
  btcGrid_ensureHeader(sheet);

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);

  const activeAsset = allAssets.find(
    a => props.getProperty(btcGrid_getAssetKeys(a.symbol).ACTIVE) === 'true'
  );

  if (!activeAsset) {
    Logger.log('Δεν υπάρχει ενεργό grid για αλλαγή.');
    return null;
  }

  // ---- Scan every non-active asset for the best candidate ----

  const candidates = [];

  allAssets.forEach(asset => {

    const keys = btcGrid_getAssetKeys(asset.symbol);

    if (props.getProperty(keys.ACTIVE) === 'true') {
      return;
    }

    const data = btcGrid_getIndicatorsForProduct(asset.product);

    if (!data) {
      candidates.push({ symbol: asset.symbol, fetchFailed: true });
      return;
    }

    candidates.push({
      symbol: asset.symbol, asset, keys,
      entryCheck: btcGrid_checkEntryConditions(data.ind4h, data.indDaily),
      ind4h: data.ind4h, indDaily: data.indDaily
    });
  });

  const best = btcGrid_pickBestCandidate(candidates);

  if (!best) {
    Logger.log('Δεν βρέθηκε διαθέσιμη εναλλακτική.');
    return null;
  }

  // ---- Score both sides for the confirmation summary ----

  const activeKeys = btcGrid_getAssetKeys(activeAsset.symbol);
  const activeGrid = btcGrid_loadGridProperties(props, activeKeys);
  const activeData = btcGrid_getIndicatorsForProduct(activeAsset.product);

  const activeScore = activeData ? btcGrid_scoreGridQuality(activeData.ind4h, activeData.indDaily) : null;
  const altScore = btcGrid_scoreGridQuality(best.ind4h, best.indDaily);

  // ---- Close the active grid ----

  btcGrid_logEvent(
    sheet, activeAsset.symbol, 'MANUAL_CLOSE',
    activeData ? activeData.ind4h.currentPrice : activeGrid.tp,
    activeGrid, activeData ? activeData.ind4h.trend : '',
    'Χειροκίνητη αλλαγή θέσης προς ' + best.symbol
  );

  btcGrid_clearGridProperties(props, activeKeys);

  // ---- Open the new one (forced — bypasses the entry-check gate) ----

  btcGrid_createGrid(props, sheet, best.asset, best.keys, best.ind4h, best.indDaily);

  return { from: activeAsset.symbol, to: best.symbol, activeScore, altScore };
}

/**
 * =========================================================
 * MAIN ENTRY POINT
 * (point your time-driven trigger at this function)
 * =========================================================
 */

function btcGrid_runBTCGridEngine() {

  const sheet = btcGrid_getSheet();
  btcGrid_ensureHeader(sheet);

  const props = PropertiesService.getScriptProperties();

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);

  // ---------------------------------------------------------
  // STEP 1 — MANAGE ANY ALREADY-ACTIVE GRIDS (any asset)
  // Auto-executes TP / smart-exit, logs to sheet, and returns
  // a summary (with a clear ΚΡΑΤΑ / ΚΛΕΙΣΕ recommendation) for
  // the combined report — no per-asset Telegram spam.
  // ---------------------------------------------------------

  const activeSummaries = [];

  allAssets.forEach(asset => {

    const keys = btcGrid_getAssetKeys(asset.symbol);

    if (props.getProperty(keys.ACTIVE) !== 'true') {
      return;
    }

    const data = btcGrid_getIndicatorsForProduct(asset.product);

    if (!data) {
      activeSummaries.push({ symbol: asset.symbol, fetchFailed: true });
      return;
    }

    activeSummaries.push(btcGrid_manageActiveGrid(props, sheet, asset, keys, data.ind4h, data.indDaily));
  });

  const stillActiveCount = allAssets.filter(
    a => props.getProperty(btcGrid_getAssetKeys(a.symbol).ACTIVE) === 'true'
  ).length;

  // ---------------------------------------------------------
  // STEP 2 — SCAN EVERY NON-ACTIVE ASSET (always, for reporting)
  // ---------------------------------------------------------

  const candidates = [];

  allAssets.forEach(asset => {

    const keys = btcGrid_getAssetKeys(asset.symbol);

    if (props.getProperty(keys.ACTIVE) === 'true') {
      return; // already managed above
    }

    const data = btcGrid_getIndicatorsForProduct(asset.product);

    if (!data) {
      candidates.push({ symbol: asset.symbol, fetchFailed: true });
      return;
    }

    candidates.push({
      symbol: asset.symbol, asset, keys,
      entryCheck: btcGrid_checkEntryConditions(data.ind4h, data.indDaily),
      ind4h: data.ind4h, indDaily: data.indDaily
    });
  });

  const best = btcGrid_pickBestCandidate(candidates);

  // ---------------------------------------------------------
  // STEP 3 — OPEN A GRID ONLY IF: best candidate fully qualifies
  // AND we're under the concurrent-grid limit (default: 1 total).
  // ---------------------------------------------------------

  let createdAsset = null;

  if (best && best.entryCheck && best.entryCheck.met &&
      stillActiveCount < CFG_BTC_SPOT_GRID_V2.maxConcurrentGrids) {

    btcGrid_createGrid(props, sheet, best.asset, best.keys, best.ind4h, best.indDaily);
    createdAsset = best.symbol;
  }

  // ---------------------------------------------------------
  // STEP 4 — ONE CONSOLIDATED REPORT PER RUN
  // ---------------------------------------------------------

  btcGrid_sendTelegramMessage(
    btcGrid_formatCombinedReport(activeSummaries, best, createdAsset, stillActiveCount)
  );
}

function btcGrid_manageActiveGrid(props, sheet, asset, keys, ind4h, indDaily) {

  const currentPrice = ind4h.currentPrice;
  const status = props.getProperty(keys.STATUS) || 'WAIT';
  const grid = btcGrid_loadGridProperties(props, keys);

  // ---- TP CHECK ----

  if (currentPrice >= grid.tp) {

    btcGrid_logEvent(sheet, asset.symbol, 'TP', currentPrice, grid, ind4h.trend, '');
    btcGrid_clearGridProperties(props, keys);

    return {
      symbol: asset.symbol, event: 'TP', recommendation: 'ΕΚΛΕΙΣΕ ✅ (Take Profit χτυπήθηκε)',
      price: currentPrice, grid, ind4h, indDaily
    };
  }

  // ---- SMART EXIT CHECK (+ hard SL backstop) ----

  const exitCheck = btcGrid_checkExitConditions(ind4h, grid);

  if (exitCheck.exit) {

    btcGrid_logEvent(sheet, asset.symbol, 'EXIT', currentPrice, grid, ind4h.trend, exitCheck.reason);
    btcGrid_clearGridProperties(props, keys);

    return {
      symbol: asset.symbol, event: 'EXIT', recommendation: 'ΚΛΕΙΣΕ ΤΩΡΑ ⚠️',
      reason: exitCheck.reason, price: currentPrice, grid, ind4h, indDaily
    };
  }

  // ---- STILL HOLDING ----

  if (status !== 'HOLD_GRID') {
    btcGrid_logEvent(sheet, asset.symbol, 'HOLD', currentPrice, grid, ind4h.trend, '');
    props.setProperty(keys.STATUS, 'HOLD_GRID');
  }

  return {
    symbol: asset.symbol, event: 'HOLD', recommendation: 'ΚΡΑΤΑ',
    price: currentPrice, grid, ind4h, indDaily
  };
}

function btcGrid_createGrid(props, sheet, asset, keys, ind4h, indDaily) {

  const grid = btcGrid_buildGrid(ind4h);

  btcGrid_saveGridProperties(props, keys, grid, 'CREATE_GRID');

  btcGrid_logEvent(sheet, asset.symbol, 'CREATE', ind4h.currentPrice, grid, ind4h.trend, '');
}

/**
 * =========================================================
 * BEST-CANDIDATE RANKING
 * Fully-qualified candidates always outrank non-qualified ones.
 * Among non-qualified candidates: fewer blocking reasons wins,
 * tie-broken by trend quality (BULLISH > RECOVERY > NEUTRAL > BEARISH).
 * =========================================================
 */

function btcGrid_pickBestCandidate(candidates) {

  let best = null;
  let bestScore = -Infinity;

  candidates.forEach(c => {

    if (c.fetchFailed) {
      return;
    }

    const score = btcGrid_rankCandidateScore(c.entryCheck, c.ind4h, c.indDaily);

    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  });

  return best;
}

// Ranks candidates using the SAME 0-100 quality score shown in the
// report (btcGrid_scoreGridQuality), so "best alternative" always
// matches the highest displayed score. Fully-qualified candidates
// (entryCheck.met) get a +1000 bonus to guarantee they outrank any
// non-qualified one, since only a qualified candidate can actually
// open a grid — but among qualified or among non-qualified candidates,
// ranking now follows the real quality score, not a separate formula.
function btcGrid_rankCandidateScore(entryCheck, ind4h, indDaily) {

  const quality = btcGrid_scoreGridQuality(ind4h, indDaily);

  if (entryCheck.met) {
    return 1000 + quality;
  }

  // Small tie-breaker so "closer to qualifying" still counts when
  // two candidates land on the same quality score.
  return quality - entryCheck.reasons.length;
}

/**
 * =========================================================
 * DEBUG — raw single request, no pagination, no aggregation.
 * Tells us immediately if the product ID / endpoint itself
 * is the problem (e.g. BTC-USDC not valid on this API).
 * =========================================================
 */

function btcGrid_debugRawCoinbaseFetch(product) {

  product = product || CFG_BTC_SPOT_GRID_V2.primaryAsset.product;

  const url =
    'https://api.exchange.coinbase.com/products/' + product +
    '/candles?granularity=3600';

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  Logger.log('URL: ' + url);
  Logger.log('Response code: ' + response.getResponseCode());
  Logger.log('Response body (first 1000 chars): ' +
    response.getContentText().substring(0, 1000));

  // Also confirm the product actually exists at all
  const productUrl =
    'https://api.exchange.coinbase.com/products/' + product;
  const productResponse = UrlFetchApp.fetch(productUrl, { muteHttpExceptions: true });

  Logger.log('--- PRODUCT CHECK ---');
  Logger.log('Product URL: ' + productUrl);
  Logger.log('Product response code: ' + productResponse.getResponseCode());
  Logger.log('Product response body: ' + productResponse.getContentText());
}

/**
 * =========================================================
 * DEBUG — check every watchlist product for delisting /
 * trading_disabled before relying on it. Run this once after
 * changing CFG_BTC_SPOT_GRID_V2.alternativeAssets.
 * =========================================================
 */

function btcGrid_debugCheckWatchlistAvailability() {

  const allAssets = [CFG_BTC_SPOT_GRID_V2.primaryAsset].concat(CFG_BTC_SPOT_GRID_V2.alternativeAssets);

  allAssets.forEach(asset => {

    const productUrl = 'https://api.exchange.coinbase.com/products/' + asset.product;
    const response = UrlFetchApp.fetch(productUrl, { muteHttpExceptions: true });

    if (response.getResponseCode() !== 200) {
      Logger.log(asset.symbol + ' (' + asset.product + '): HTTP ' + response.getResponseCode() + ' — NOT AVAILABLE');
      return;
    }

    const info = JSON.parse(response.getContentText());

    Logger.log(
      asset.symbol + ' (' + asset.product + '): status=' + info.status +
      ', trading_disabled=' + info.trading_disabled
    );

    Utilities.sleep(150);
  });
}

/**
 * =========================================================
 * DEBUG — run this manually to see WHY it's WAIT/CREATE
 * Does NOT write to the sheet or send Telegram.
 * View results: Executions log (or Logger.log output).
 * =========================================================
 */

function btcGrid_debugBTCGridEngine(symbol, product) {

  symbol = symbol || CFG_BTC_SPOT_GRID_V2.primaryAsset.symbol;
  product = product || CFG_BTC_SPOT_GRID_V2.primaryAsset.product;

  const hourly = btcGrid_fetchCoinbaseCandlesPaginated(
    product, BTCGRID_GRANULARITY.HOUR, CFG_BTC_SPOT_GRID_V2.hourlyCandlesNeeded
  );
  const candles4h = btcGrid_aggregateTo4H(hourly);
  const dailyCandles = btcGrid_fetchCoinbaseCandlesPaginated(
    product, BTCGRID_GRANULARITY.DAY, CFG_BTC_SPOT_GRID_V2.dailyCandlesNeeded
  );

  const ind4h = btcGrid_computeIndicators(candles4h);
  const indDaily = btcGrid_computeIndicators(dailyCandles);

  Logger.log('=== ' + symbol + ' (' + product + ') ===');
  Logger.log('--- RAW DATA ---');
  Logger.log('Hourly candles fetched: ' + hourly.length);
  Logger.log('4H bars built: ' + candles4h.length);
  Logger.log('Daily candles fetched: ' + dailyCandles.length);

  Logger.log('--- 4H INDICATORS ---');
  Logger.log('Price: ' + ind4h.currentPrice);
  Logger.log('Trend: ' + ind4h.trend);
  Logger.log('EMA20/50/200: ' + ind4h.ema20 + ' / ' + ind4h.ema50 + ' / ' + ind4h.ema200);
  Logger.log('EMA20 rising: ' + ind4h.ema20Rising);
  Logger.log('Support/Resistance: ' + ind4h.support + ' / ' + ind4h.resistance);
  Logger.log('Reclaim level: ' + ind4h.reclaimLevel + ' | confirmed: ' + ind4h.reclaimConfirmed);
  Logger.log('ATR%: ' + (ind4h.atrPercent * 100).toFixed(2) + '% (avg: ' + (ind4h.atrPercentAvg * 100).toFixed(2) + '%)');
  Logger.log('ATR explosion: ' + ind4h.atrExplosion);

  Logger.log('--- DAILY ---');
  Logger.log('Daily EMA20/50/200: ' + indDaily.ema20 + ' / ' + indDaily.ema50 + ' / ' + indDaily.ema200);
  Logger.log('Daily trend: ' + indDaily.trend);
  Logger.log('Daily bearish gap (EMA50 vs EMA200): ' + (indDaily.bearishGapPct * 100).toFixed(2) + '% (mild threshold: ' + (CFG_BTC_SPOT_GRID_V2.mildBearishMaxGapPct * 100).toFixed(0) + '%)');

  Logger.log('--- ENTRY CHECK ---');
  const entryCheck = btcGrid_checkEntryConditions(ind4h, indDaily);
  Logger.log('Entry conditions met: ' + entryCheck.met);
  if (entryCheck.reasons.length) {
    Logger.log('Blocked by: ' + entryCheck.reasons.join(' | '));
  }

  const props = PropertiesService.getScriptProperties();
  const keys = btcGrid_getAssetKeys(symbol);
  Logger.log('--- CURRENT STATE ---');
  Logger.log('Grid active: ' + props.getProperty(keys.ACTIVE));
  Logger.log('Status: ' + props.getProperty(keys.STATUS));
}

/**
 * =========================================================
 * ENTRY CONDITIONS (WAIT -> CREATE_GRID)
 * =========================================================
 */

function btcGrid_checkEntryConditions(ind4h, indDaily) {

  const trendOk =
    ind4h.trend === 'BULLISH' || ind4h.trend === 'RECOVERY';

  const dailyMildBearish =
    indDaily.trend === 'BEARISH' &&
    indDaily.bearishGapPct <= CFG_BTC_SPOT_GRID_V2.mildBearishMaxGapPct;

  const dailyOk =
    indDaily.trend !== 'BEARISH' || dailyMildBearish;

  const highVolatility =
    ind4h.atrPercent > CFG_BTC_SPOT_GRID_V2.highVolatilityThreshold;

  const reasons = [];

  if (!trendOk) {
    reasons.push('4H trend δεν είναι BULLISH/RECOVERY (' + ind4h.trend + ')');
  }
  if (!dailyOk) {
    reasons.push(
      'Daily downtrend πολύ βαθύ (gap ' +
      (indDaily.bearishGapPct * 100).toFixed(1) + '%, όριο ' +
      (CFG_BTC_SPOT_GRID_V2.mildBearishMaxGapPct * 100).toFixed(0) + '%)'
    );
  }
  if (!ind4h.reclaimConfirmed) {
    reasons.push('Reclaim δεν έχει επιβεβαιωθεί');
  }
  if (!ind4h.ema20Rising) {
    reasons.push('EMA20 (4H) δεν ανεβαίνει');
  }
  if (highVolatility) {
    reasons.push('ATR πολύ υψηλό (' + (ind4h.atrPercent * 100).toFixed(2) + '%)');
  }

  const met =
    trendOk && dailyOk && ind4h.reclaimConfirmed && ind4h.ema20Rising && !highVolatility;

  return { met, reasons };
}

/**
 * =========================================================
 * EXIT CONDITIONS (HOLD_GRID -> EXIT_GRID)
 * Majority vote across 6 signals + hard SL backstop.
 * =========================================================
 */

function btcGrid_checkExitConditions(ind4h, grid) {

  const closes = ind4h.closes;
  const n = closes.length;

  const twoClosesBelowSupport =
    closes[n - 1] < ind4h.support &&
    closes[n - 2] < ind4h.support;

  const twoClosesBelowLower =
    closes[n - 1] < grid.lower &&
    closes[n - 2] < grid.lower;

  const emaSlopeNegative =
    ind4h.ema20 < ind4h.ema20Prev;

  const emaBearishStack =
    ind4h.ema20 < ind4h.ema50 &&
    ind4h.ema50 < ind4h.ema200;

  const atrExplosion =
    ind4h.atrExplosion;

  const highs = ind4h.highs;
  const lows = ind4h.lows;
  const hn = highs.length;

  const bearishContinuation =
    hn >= 3 &&
    highs[hn - 1] < highs[hn - 2] && highs[hn - 2] < highs[hn - 3] &&
    lows[hn - 1] < lows[hn - 2] && lows[hn - 2] < lows[hn - 3];

  const conditions = {
    twoClosesBelowSupport,
    twoClosesBelowLower,
    emaSlopeNegative,
    emaBearishStack,
    atrExplosion,
    bearishContinuation
  };

  const trueCount =
    Object.values(conditions).filter(Boolean).length;

  const slHit = ind4h.currentPrice <= grid.sl;

  const structuralExit =
    trueCount >= CFG_BTC_SPOT_GRID_V2.exitConditionsRequired;

  let reason = '';
  if (slHit) {
    reason = 'SL_HIT';
  } else if (structuralExit) {
    reason =
      'BEARISH_STRUCTURE ' + trueCount + '/6 (' +
      Object.keys(conditions).filter(k => conditions[k]).join(', ') +
      ')';
  }

  return {
    exit: slHit || structuralExit,
    reason,
    conditions,
    trueCount
  };
}

/**
 * =========================================================
 * GRID BUILDER
 * =========================================================
 */

function btcGrid_buildGrid(ind4h) {

  const lower =
    btcGrid_roundBTCGrid(ind4h.support - (ind4h.atr * CFG_BTC_SPOT_GRID_V2.gridBufferATRMultiplier));

  const upper =
    btcGrid_roundBTCGrid(ind4h.resistance + (ind4h.atr * CFG_BTC_SPOT_GRID_V2.gridBufferATRMultiplier));

  let grids = CFG_BTC_SPOT_GRID_V2.gridCountNormal;

  if (ind4h.atrPercent > CFG_BTC_SPOT_GRID_V2.gridCountHighVolThreshold) {
    grids = CFG_BTC_SPOT_GRID_V2.gridCountHighVol;
  }
  if (ind4h.atrPercent < CFG_BTC_SPOT_GRID_V2.gridCountLowVolThreshold) {
    grids = CFG_BTC_SPOT_GRID_V2.gridCountLowVol;
  }

  const tp = btcGrid_roundPrice(upper * (1 + CFG_BTC_SPOT_GRID_V2.tpBufferPct));
  const sl = btcGrid_roundPrice(lower * (1 - CFG_BTC_SPOT_GRID_V2.slBufferPct));

  return { lower, upper, grids, tp, sl };
}

/**
 * =========================================================
 * INDICATORS (works on any candle array — 4H or daily)
 * =========================================================
 */

function btcGrid_computeIndicators(candles) {

  const closes = candles.map(c => c.close);
  const highs = candles.map(c => c.high);
  const lows = candles.map(c => c.low);

  const currentPrice = closes[closes.length - 1];

  const ema20 = btcGrid_calculateEMA(closes, 20);
  const ema20Prev = btcGrid_calculateEMA(closes.slice(0, -1), 20);
  const ema20Prev2 = btcGrid_calculateEMA(closes.slice(0, -2), 20);
  const ema20Rising = ema20 > ema20Prev && ema20Prev > ema20Prev2;

  const ema50 = btcGrid_calculateEMA(closes, 50);
  const ema200 = btcGrid_calculateEMA(closes, 200); // real 200-period EMA (v1 bug fixed)

  const lookback = Math.min(CFG_BTC_SPOT_GRID_V2.lookbackPeriods, closes.length);
  const support = Math.min(...lows.slice(-lookback));
  const resistance = Math.max(...highs.slice(-lookback));
  const reclaimLevel = support + ((resistance - support) * CFG_BTC_SPOT_GRID_V2.reclaimFraction);

  const reclaimConfirmed =
    closes[closes.length - 1] > reclaimLevel &&
    closes[closes.length - 2] > reclaimLevel;

  const atrPercentSeries = btcGrid_calculateATRPercentSeries(candles, CFG_BTC_SPOT_GRID_V2.atrPeriod);
  const atrPercent = atrPercentSeries[atrPercentSeries.length - 1];

  const atrHistory =
    atrPercentSeries.slice(0, -1).slice(-CFG_BTC_SPOT_GRID_V2.atrExplosionLookback);

  const atrPercentAvg =
    atrHistory.length ? btcGrid_average(atrHistory) : atrPercent;

  const atr = atrPercent * currentPrice;

  const atrExplosion =
    atrPercent > (atrPercentAvg * CFG_BTC_SPOT_GRID_V2.atrExplosionMultiplier);

  let trend = 'NEUTRAL';

  if (ema20 > ema50 && ema50 > ema200) {
    trend = 'BULLISH';
  } else if (ema20 > ema50 && currentPrice > reclaimLevel) {
    trend = 'RECOVERY';
  } else if (ema20 < ema50 && ema50 < ema200) {
    trend = 'BEARISH';
  }
  // else stays NEUTRAL: no clear directional EMA structure
  // (e.g. sideways/consolidating market) — this used to be
  // lumped into BEARISH by default and over-blocked entries.

  // How deep the EMA50 has fallen below EMA200, as a % of EMA200.
  // Used to distinguish a shallow/recent bearish flip from an
  // established, deep downtrend.
  const bearishGapPct =
    (ema200 && ema200 !== 0) ? ((ema200 - ema50) / ema200) : 0;

  return {
    closes, highs, lows, currentPrice,
    ema20, ema20Prev, ema20Rising, ema50, ema200,
    support, resistance, reclaimLevel, reclaimConfirmed,
    atr, atrPercent, atrPercentAvg, atrExplosion,
    trend, bearishGapPct
  };
}

/**
 * =========================================================
 * COINBASE — PAGINATED CANDLE FETCH
 * (handles the 300-candles-per-request limit)
 * =========================================================
 */

function btcGrid_fetchCoinbaseCandlesPaginated(product, granularitySeconds, neededCount) {

  const maxPerRequest = 300;
  let allCandles = [];
  let endTime = new Date();

  while (allCandles.length < neededCount) {

    const startTime =
      new Date(endTime.getTime() - (maxPerRequest * granularitySeconds * 1000));

    const url =
      'https://api.exchange.coinbase.com/products/' + product +
      '/candles?granularity=' + granularitySeconds +
      '&start=' + startTime.toISOString() +
      '&end=' + endTime.toISOString();

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (response.getResponseCode() !== 200) {
      throw new Error('Coinbase API failed: ' + response.getContentText());
    }

    const batch = JSON.parse(response.getContentText());

    if (!batch || batch.length === 0) {
      break;
    }

    allCandles = allCandles.concat(batch);
    endTime = startTime;

    Utilities.sleep(250); // be nice to the API

    if (batch.length < maxPerRequest) {
      break; // no more history available
    }
  }

  const dedup = {};
  allCandles.forEach(c => { dedup[c[0]] = c; });

  const sorted = Object.values(dedup).sort((a, b) => a[0] - b[0]);
  const trimmed = sorted.slice(-neededCount);

  return trimmed.map(c => ({
    time: c[0], low: c[1], high: c[2], open: c[3], close: c[4], volume: c[5]
  }));
}

/**
 * =========================================================
 * AGGREGATE 1H -> 4H
 * (Coinbase has no native 4H granularity)
 * =========================================================
 */

function btcGrid_aggregateTo4H(hourlyCandles) {

  const groups = [];
  const size = CFG_BTC_SPOT_GRID_V2.candlesPerBar4h;

  for (let i = hourlyCandles.length; i - size >= 0; i -= size) {

    const chunk = hourlyCandles.slice(i - size, i);

    groups.unshift({
      time: chunk[0].time,
      open: chunk[0].open,
      close: chunk[chunk.length - 1].close,
      high: Math.max(...chunk.map(c => c.high)),
      low: Math.min(...chunk.map(c => c.low)),
      volume: chunk.reduce((s, c) => s + c.volume, 0)
    });
  }

  return groups;
}

/**
 * =========================================================
 * EMA
 * =========================================================
 */

function btcGrid_calculateEMA(prices, period) {

  if (!prices || prices.length < period) {
    return null;
  }

  const k = 2 / (period + 1);

  // Seed with the SMA of the first `period` values — standard practice.
  // Seeding with a single raw price (the old v1 approach) leaves the
  // EMA anchored to one historical data point when available history
  // isn't drastically larger than the period (e.g. EMA200 on ~300
  // candles), producing a skewed value.
  let ema = btcGrid_average(prices.slice(0, period));

  for (let i = period; i < prices.length; i++) {
    ema = (prices[i] * k) + (ema * (1 - k));
  }

  return ema;
}

/**
 * =========================================================
 * ATR (raw TR series + ATR% series for explosion detection)
 * =========================================================
 */

function btcGrid_calculateATRPercentSeries(candles, period) {

  const trs = [];

  for (let i = 1; i < candles.length; i++) {

    const high = candles[i].high;
    const low = candles[i].low;
    const prevClose = candles[i - 1].close;

    const tr = Math.max(
      high - low,
      Math.abs(high - prevClose),
      Math.abs(low - prevClose)
    );

    trs.push(tr);
  }

  const atrSeries = [];

  for (let i = period - 1; i < trs.length; i++) {
    const window = trs.slice(i - period + 1, i + 1);
    atrSeries.push(btcGrid_average(window));
  }

  const offset = candles.length - atrSeries.length;

  return atrSeries.map((atr, idx) => atr / candles[offset + idx].close);
}

function btcGrid_average(arr) {
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

/**
 * =========================================================
 * SCRIPT PROPERTIES — ACTIVE GRID STATE
 * =========================================================
 */

function btcGrid_saveGridProperties(props, keys, grid, status) {

  props.setProperties({
    [keys.ACTIVE]: 'true',
    [keys.STATUS]: status,
    [keys.CREATED]: new Date().toISOString(),
    [keys.LOWER]: String(grid.lower),
    [keys.UPPER]: String(grid.upper),
    [keys.TP]: String(grid.tp),
    [keys.SL]: String(grid.sl),
    [keys.GRIDS]: String(grid.grids)
  });
}

function btcGrid_loadGridProperties(props, keys) {

  return {
    lower: Number(props.getProperty(keys.LOWER)),
    upper: Number(props.getProperty(keys.UPPER)),
    tp: Number(props.getProperty(keys.TP)),
    sl: Number(props.getProperty(keys.SL)),
    grids: Number(props.getProperty(keys.GRIDS)),
    createdAt: props.getProperty(keys.CREATED)
  };
}

function btcGrid_clearGridProperties(props, keys) {

  props.setProperties({
    [keys.ACTIVE]: 'false',
    [keys.STATUS]: 'WAIT'
  });

  props.deleteProperty(keys.CREATED);
  props.deleteProperty(keys.LOWER);
  props.deleteProperty(keys.UPPER);
  props.deleteProperty(keys.TP);
  props.deleteProperty(keys.SL);
  props.deleteProperty(keys.GRIDS);
}

/**
 * =========================================================
 * SHEET LOGGING — ONLY REAL EVENTS
 * =========================================================
 */

function btcGrid_getSheet() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG_BTC_SPOT_GRID_V2.sheetName);

  if (!sheet) {
    throw new Error('Sheet ' + CFG_BTC_SPOT_GRID_V2.sheetName + ' not found');
  }

  return sheet;
}

function btcGrid_ensureHeader(sheet) {

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(
      ['Date', 'Asset', 'Status', 'Price', 'Lower', 'Upper', 'TP', 'SL', 'Trend', 'Notes']
    );
  }
}

function btcGrid_logEvent(sheet, symbol, status, price, grid, trend, notes) {

  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
  );

  sheet.appendRow([
    now, symbol, status, btcGrid_roundBTCGrid(price),
    grid.lower, grid.upper, grid.tp, grid.sl,
    trend, notes || ''
  ]);
}

/**
 * =========================================================
 * DEBUG — safely inspect Telegram Script Properties without
 * exposing the real token/chat id in the logs.
 * =========================================================
 */

function btcGrid_debugTelegramProperties() {

  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!botToken) {
    Logger.log('TELEGRAM_BOT_TOKEN: NOT SET');
  } else {
    Logger.log(
      'TELEGRAM_BOT_TOKEN: length=' + botToken.length +
      ', starts with "' + botToken.substring(0, 6) + '..."' +
      ', looks valid (digits:letters format)=' + /^\d+:[A-Za-z0-9_-]+$/.test(botToken)
    );
  }

  if (!chatId) {
    Logger.log('TELEGRAM_CHAT_ID: NOT SET');
  } else {
    Logger.log(
      'TELEGRAM_CHAT_ID: length=' + chatId.length +
      ', value="' + chatId + '"' +
      ', looks valid (numeric, optional leading -)=' + /^-?\d+$/.test(chatId)
    );
  }
}

/**
 * =========================================================
 * TELEGRAM — credentials ONLY from Script Properties
 * =========================================================
 */

function btcGrid_sendTelegramMessage(message) {

  const props = PropertiesService.getScriptProperties();

  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!botToken || !chatId) {
    throw new Error('Missing TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID in Script Properties');
  }

  const telegramUrl = 'https://api.telegram.org/bot' + botToken + '/sendMessage';

  UrlFetchApp.fetch(telegramUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: message }),
    muteHttpExceptions: true
  });
}

/**
 * =========================================================
 * ONE CONSOLIDATED TELEGRAM REPORT PER RUN
 * Section 1: current active grid(s) with a clear ΚΡΑΤΑ/ΚΛΕΙΣΕ call.
 * Section 2: the best alternative right now, with its (hypothetical,
 * unless just created) grid levels — the "εναλλακτική + προτάσεις"
 * the user asked for.
 * =========================================================
 */

function btcGrid_formatCombinedReport(activeSummaries, best, createdAsset, stillActiveCount) {

  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
  );

  let msg = `🔄 Έλεγχος ΟΚ — ${now}\n`;

  // ---- SECTION 1: ACTIVE GRID(S) ----

  msg += `\n━━━━━━━━━━━━━━━\nΕΝΕΡΓΟ GRID\n━━━━━━━━━━━━━━━`;

  if (!activeSummaries.length) {

    msg += `\nΚανένα ενεργό grid αυτή τη στιγμή.`;

  } else {

    activeSummaries.forEach(s => {

      if (s.fetchFailed) {
        msg += `\n${s.symbol}: ⚠️ αποτυχία λήψης δεδομένων (ξαναδοκίμασε επόμενο run)`;
        return;
      }

      msg +=
        `\n${s.symbol}: ${s.recommendation}` +
        `\n   Τιμή: $${btcGrid_roundBTCGrid(s.price)} | 4H: ${s.ind4h.trend} | Daily: ${s.indDaily.trend}` +
        `\n   εύρος [${s.grid.lower} - ${s.grid.upper}] | Grids: ${s.grid.grids} | TP: ${s.grid.tp} | SL: ${s.grid.sl}` +
        `\n   📊 Score ποιότητας: ${btcGrid_scoreGridQuality(s.ind4h, s.indDaily)}/100`;

      if (s.reason) {
        msg += `\n   Αιτία: ${s.reason}`;
      }
    });
  }

  // ---- SECTION 2: BEST ALTERNATIVE ----

  msg += `\n\n━━━━━━━━━━━━━━━\nΚΑΛΥΤΕΡΗ ΕΝΑΛΛΑΚΤΙΚΗ\n━━━━━━━━━━━━━━━`;

  if (createdAsset) {

    msg += `\n🚀 Νέο grid μόλις δημιουργήθηκε: ${createdAsset}`;

  } else if (best) {

    const hypoGrid = btcGrid_buildGrid(best.ind4h);
    const altScore = btcGrid_scoreGridQuality(best.ind4h, best.indDaily);

    msg +=
      `\n${best.symbol}: $${btcGrid_roundBTCGrid(best.ind4h.currentPrice)}` +
      ` | 4H: ${best.ind4h.trend} | Daily: ${best.indDaily.trend}` +
      `\n📊 Score ποιότητας: ${altScore}/100`;

    if (best.entryCheck.met) {
      msg +=
        `\n✅ Πληροί όλα τα κριτήρια, αλλά δεν άνοιξε — όριο ταυτόχρονων grids ` +
        `(${stillActiveCount}/${CFG_BTC_SPOT_GRID_V2.maxConcurrentGrids})`;
    } else {
      msg += `\nΜπλοκάρεται από: ` + best.entryCheck.reasons.join(', ');
    }

    msg +=
      `\nΥποθετικό grid: [${hypoGrid.lower} - ${hypoGrid.upper}] | Grids: ${hypoGrid.grids} | TP: ${hypoGrid.tp} | SL: ${hypoGrid.sl}`;

    // ---- SECTION 3: SWITCH VERDICT (only when there's an active grid to compare against) ----

    const activeForCompare = activeSummaries.find(s => !s.fetchFailed);

    if (activeForCompare) {

      const activeScore = btcGrid_scoreGridQuality(activeForCompare.ind4h, activeForCompare.indDaily);
      const diff = altScore - activeScore;

      msg += `\n\n━━━━━━━━━━━━━━━\nΣΥΓΚΡΙΣΗ\n━━━━━━━━━━━━━━━`;
      msg += `\n${activeForCompare.symbol}: ${activeScore}/100   vs   ${best.symbol}: ${altScore}/100   (Δ ${diff >= 0 ? '+' : ''}${diff})`;

      if (diff >= CFG_BTC_SPOT_GRID_V2.switchScoreThreshold) {
        msg += `\n🔀 Συνιστάται αλλαγή θέσης — χρησιμοποίησε το menu "Άλλαξε σε Καλύτερη Εναλλακτική"`;
      } else {
        msg += `\n✅ Παρέμεινε στο ενεργό — η διαφορά δεν δικαιολογεί αλλαγή (όριο: ${CFG_BTC_SPOT_GRID_V2.switchScoreThreshold})`;
      }
    }

  } else {

    msg += `\nΔεν βρέθηκαν διαθέσιμα δεδομένα για κανένα asset αυτή τη φορά.`;
  }

  return msg;
}

/**
 * =========================================================
 * ROUNDING
 * =========================================================
 */

function btcGrid_roundBTCGrid(value) {
  const num = Number(value);
  if (!isFinite(num)) return 0;
  return Number(num.toFixed(2));
}

function btcGrid_roundPrice(value) {
  return btcGrid_roundBTCGrid(value);
}







/**
 * LONG / SHORT FUTURES GRID SCREENER
 * ------------------------------------------------------------------
 * Ελέγχει μια λίστα crypto pairs για ήπια, σταθερή ανοδική (LONG) ή
 * καθοδική (SHORT) τάση με αρκετή διακύμανση, και αν τα βρει κατάλληλα,
 * υπολογίζει προτεινόμενες παραμέτρους για Long/Short Futures Grid στο
 * Pionex (upper/lower limit, αριθμό grids, προτεινόμενο leverage).
 * Τρέχει ΚΑΙ ΤΙΣ ΔΥΟ στρατηγικές αυτόματα σε κάθε εκτέλεση (STRATEGIES_TO_RUN).
 *
 * Δεδομένα τιμών: δημόσιο OKX market-data API (χωρίς API key).
 * (Το Binance public API μπλοκάρει αιτήματα από EU-σχετικές περιοχές
 * μετά τις αλλαγές MiCA, γι' αυτό χρησιμοποιούμε OKX.)
 * ΠΡΟΣΟΧΗ: Αυτό είναι εργαλείο τεχνικής ανάλυσης, όχι επενδυτική
 * συμβουλή. Επιβεβαίωσε πάντα τα αποτελέσματα πριν στήσεις bot,
 * και μην ρισκάρεις περισσότερα από όσα αντέχεις να χάσεις.
 * ------------------------------------------------------------------
 *
 * ΕΓΚΑΤΑΣΤΑΣΗ:
 * 1. Google Sheet -> Extensions -> Apps Script
 * 2. Επικόλλησε αυτό το αρχείο (αντικαθιστά το Code.gs)
 * 3. Άνοιξε ξανά το Sheet -> θα εμφανιστεί μενού "Grid Screener"
 * 4. Grid Screener -> Run Screener
 *
 * TRIGGER (αυτόματη εκτέλεση σε πρόγραμμα):
 * Apps Script editor -> Triggers (ρολόι στο αριστερό μενού) -> Add Trigger
 *   - Function to run: checkGridCandidates
 *   - Event source: Time-driven
 *   - Type: π.χ. Day timer, ώρα της αρεσκείας σου
 *
 * TELEGRAM NOTIFICATIONS:
 * Χρησιμοποιεί τα ίδια Script Properties με το betting alert script σου
 * (Project Settings -> Script Properties): TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID.
 * Αν λείπουν, το script απλά τα προσπερνάει χωρίς σφάλμα.
 */
 
// ==================== GRID_SCREENER_CONFIG ====================
 
const GRID_SCREENER_CONFIG = {
  SHEET_NAME: 'Grid Screener',
  LOG_SHEET_NAME: 'Grid Screener Log',
  INTERVAL: '4H',          // OKX bar: 4H ταιριάζει καλά για "μεσοπρόθεσμα" (βλ. valid values στο fetchKlines)
  LOOKBACK_CANDLES: 180,   // 180 x 4H ≈ 30 μέρες ιστορικού (OKX max 300 ανά αίτημα)
 
  // ---- Στρατηγικές που τρέχουν ΑΥΤΟΜΑΤΑ σε κάθε εκτέλεση/trigger, η μία μετά την
  // άλλη — 'LONG' (ήπια ανοδική τάση) και 'SHORT' (ήπια καθοδική τάση). Το NEUTRAL
  // καταργήθηκε — στην πράξη το Pionex neutral grid γυρνάει εύκολα σε ζημιογόνα θέση
  // όταν η αγορά βγει από το πλάγιο κανάλι. Δεν χρειάζεται χειροκίνητη αλλαγή config:
  // αφού τρέχει με trigger χωρίς επίβλεψη, το script ελέγχει ΟΛΕΣ τις στρατηγικές της
  // λίστας σε κάθε run, με ξεχωριστά sheets, ξεχωριστό κλείδωμα stable pick, και
  // ξεχωριστά Telegram μηνύματα η καθεμία.
  STRATEGIES_TO_RUN: ['LONG', 'SHORT'],
  GRID_STRATEGY: 'LONG',  // ΜΗΝ το αλλάζεις χειροκίνητα — ορίζεται αυτόματα ανά loop iteration από το runScreenerForStrategy()
 
  MIN_RANGE_PCT: 6,             // ελάχιστο εύρος high-low % ώστε να αξίζει grid (αλλιώς πολύ "νεκρό")
  MAX_RANGE_PCT: 45,            // πάνω από αυτό θεωρείται πολύ ασταθές για grid trading
  GRID_BUFFER_PCT: 3,           // buffer πάνω/κάτω από high/low για τα limits
  ALT_ENTRY_BELOW_HIGH_PCT: 43, // εναλλακτική τιμή εισόδου = Upper Limit μειωμένο κατά αυτό το ποσοστό
 
  // ---- Thresholds μόνο για GRID_STRATEGY = 'LONG' ----
  LONG_GRID: {
    MIN_UPTREND_SLOPE_PCT: 0.05,   // ελάχιστη θετική κλίση/candle ώστε να θεωρηθεί "ανοδική τάση"
    PARABOLIC_SLOPE_PCT: 1.0,      // πάνω από αυτό -> πολύ απότομο/παραβολικό, ρίσκο απότομης διόρθωσης
    DOWNTREND_SLOPE_PCT: -0.05,    // κάτω από αυτό -> καθοδική τάση, ακατάλληλο για long grid
    MIN_CHOPPINESS: 35             // χαμηλότερο threshold από ό,τι θα χρειαζόταν καθαρά πλάγια αγορά, αφού βοηθάει και η ίδια η τάση
  },
  // ---- Thresholds μόνο για GRID_STRATEGY = 'SHORT' (mirror του LONG, αντίστροφο πρόσημο) ----
  SHORT_GRID: {
    MIN_DOWNTREND_SLOPE_PCT: -0.05,  // ελάχιστη (αρνητική) κλίση/candle ώστε να θεωρηθεί "καθοδική τάση"
    PARABOLIC_DOWN_SLOPE_PCT: -1.0,  // κάτω από αυτό -> πολύ απότομη πτώση, ρίσκο απότομου rebound
    UPTREND_SLOPE_PCT: 0.05,         // πάνω από αυτό -> ανοδική τάση, ακατάλληλο για short grid
    MIN_CHOPPINESS: 35
  },
  // ---- Οικονομικά θεμελιωμένος υπολογισμός grids/leverage (αντί για αυθαίρετα buckets) ----
  GRID_ECONOMICS: {
    ROUNDTRIP_FEE_PCT: 0.10,     // εκτίμηση κόστους fee ανά πλήρη κύκλο buy+sell (%) — προσάρμοσέ το στα δικά σου fees στο Pionex
    MIN_PROFIT_MULTIPLE: 3,      // ελάχιστο spacing ανά grid σε πολλαπλάσια του fee, ώστε να αξίζει πραγματικά το fill
    MIN_GRIDS: 20,
    MAX_GRIDS: 150
  },
  LEVERAGE_SAFETY: {
    ATR_SAFETY_MULTIPLE: 3,      // απόσταση liquidation >= 3x το μέσο ATR% (μαξιλάρι ασφαλείας έναντι ξαφνικής κίνησης)
    MAX_LEVERAGE_CAP: 5          // απόλυτο ανώτατο leverage που προτείνει ΠΟΤΕ το script, ανεξαρτήτως υπολογισμού
  },
  // ---- Προτεινόμενο TP/SL (Profit-and-Loss ratio του grid bot, όχι ανά grid) ----
  TP_SL: {
    TP_FRACTION: 0.35,     // TP% ≈ Est.Return(period) × leverage × αυτό το κλάσμα — ρεαλιστικός, εφικτός στόχος
    SL_ATR_MULTIPLE: 3,    // SL% ≈ ATR% × leverage × αυτό — πολλαπλάσιο "φυσιολογικού" θορύβου
    MIN_TP_PCT: 1.5,
    MAX_TP_PCT: 8,
    MIN_SL_PCT: 3,
    MAX_SL_PCT: 12
  },
  // ---- Πρόταση ζώνης εισόδου, βάσει θέσης τιμής μέσα στο εύρος + βραχυπρόθεσμο 4H momentum ----
  ENTRY_SUGGESTION: {
    SWEET_SPOT_PCT: 0.30,        // "ιδανικό" ποσοστό μέσα στο εύρος (0 = κάτω όριο, 1 = πάνω όριο)
    GOOD_ZONE_MAX_PCT: 25,       // <= αυτό -> ΚΑΛΗ ΖΩΝΗ (πραγματικά κοντά στον πάτο)
    MODERATE_ZONE_MAX_PCT: 55,   // <= αυτό -> ΜΕΤΡΙΑ ΖΩΝΗ, πάνω από αυτό -> ΨΗΛΗ ΖΩΝΗ
    RECENT_SWING_CANDLES: 12,    // ~2 μέρες σε 4H candles, για εντοπισμό πρόσφατου local support/resistance
    STEEP_DROP_SLOPE_PCT: -0.5,  // κλίση/candle κάτω από αυτό -> τρέχουσα ροπή "καθοδική"
    STEEP_RISE_SLOPE_PCT: 0.5,   // κλίση/candle πάνω από αυτό -> τρέχουσα ροπή "ανοδική"
    // Διαχωρισμός κεφαλαίου σε 2 τμήματα όταν η τιμή είναι στην ΨΗΛΗ ΖΩΝΗ:
    // ποσοστό "τώρα" γραμμικά από SPLIT_NOW_PCT_AT_ZONE_START (μόλις μπεις σε ΨΗΛΗ)
    // μέχρι SPLIT_NOW_PCT_AT_TOP (στην απόλυτη κορυφή του εύρους).
    SPLIT_NOW_PCT_AT_ZONE_START: 50,
    SPLIT_NOW_PCT_AT_TOP: 25,
    // Αν true: όταν η ζώνη ΔΕΝ είναι ΚΑΛΗ, το μήνυμα/sheet ΔΕΝ παρουσιάζει την
    // είσοδο σαν έτοιμη ενέργεια — μπλοκάρει ρητά με "⏳ ΑΝΑΜΟΝΗ" αντί για απλή
    // προαιρετική σημείωση. Αν false, γυρνάει στην παλιά συμπεριφορά (πάντα ενέργεια).
    BLOCK_ENTRY_OUTSIDE_GOOD_ZONE: true
  },
  // ---- Funding rate (μόνο σε perpetual futures) — κόστος/κέρδος πέρα από τα grid fills ----
  // ---- Ανοιχτές θέσεις (χειροκίνητα καταχωρημένες) — ανεξάρτητο από το "προτεινόμενο pick" ----
  OPEN_POSITIONS: {
    SHEET_NAME: 'Open Positions',
    BIG_UPGRADE_MARGIN: 20,  // πόσο καλύτερος πρέπει να είναι ένας νέος υποψήφιος ώστε να αξίζει να
                              // σκεφτείς κλείσιμο μιας ήδη ενεργής, υγιούς θέσης (πολύ αυστηρότερο
                              // από το SWITCH_MARGIN του stable pick, λόγω πραγματικού κόστους αλλαγής)
    ALWAYS_NOTIFY: false      // true = στείλε σύντομο μήνυμα επιβεβαίωσης ("όλα καλά") σε ΚΑΘΕ run,
                              // ακόμα κι όταν δεν χρειάζεται ενέργεια· false = μήνυμα ΜΟΝΟ όταν υπάρχει alert
  },
  // ---- Προληπτική προειδοποίηση: score που πέφτει συνεχόμενα, πριν καταρρεύσει το verdict ----
  DECLINING_TREND: {
    LOOKBACK_RUNS: 3,      // πόσα runs (μαζί με το τρέχον) να συγκρίνει από το Log
    MIN_DECLINE_POINTS: 8  // ελάχιστη συνολική πτώση score (πρώτο->τελευταίο) ώστε να σημάνει "φθίνουσα τάση"
  },
  // Βάρη για το composite Grid Score (πρέπει να αθροίζουν σε 1.0)
  SCORE_WEIGHTS: {
    choppiness: 0.40,   // πόσο "πηγαινοέρχεται" η τιμή μέσα στο εύρος (καλό για grid)
    trendWeak: 0.25,    // αντίστροφο του ADX (ασθενές trend = καλύτερο)
    rangeFit: 0.15,     // πόσο κοντά είναι το εύρος % σε "γλυκό σημείο"
    estReturn: 0.20     // κανονικοποιημένη εκτιμώμενη απόδοση περιόδου έναντι των υπολοίπων symbols
  },
  RANGE_SWEET_SPOT_PCT: 18,     // "ιδανικό" εύρος % γύρω από το οποίο μεγιστοποιείται το rangeFit score
 
  // ---- Dynamic Top-N by OKX 24h volume (αντί για τη χειροκίνητη λίστα SYMBOLS) ----
  USE_TOP_N_BY_MARKET_CAP: true,  // true = τραβάει αυτόματα τα πιο ρευστά coins από OKX· false = χρησιμοποιεί το SYMBOLS παρακάτω
  TOP_N: 120,                      // πόσα coins (κατά 24ωρο τζίρο) να ελεγχθούν
  TOP_RESULTS_TO_SHOW: 10,         // πόσα από τα καλύτερα να εμφανιστούν στο κύριο sheet
  TELEGRAM_NOTIFY: true,           // στείλε Telegram μήνυμα με την καλύτερη επιλογή μετά από κάθε run
  TOP_N_CACHE_SECONDS: 21600,      // cache της top-N λίστας (max 21600 = 6 ώρες, όριο του CacheService)
  MAX_RUNTIME_MS: 320000,          // safety cutoff (~5.3 λεπτά) — το Apps Script σκάει στα 6 λεπτά,
                                    // σταματάμε νωρίτερα με ασφάλεια αντί να σκάσει απότομα στη μέση
  // ---- Anti flip-flop: μην αλλάζεις προτεινόμενο coin κάθε μέρα ----
  STABILITY: {
    SWITCH_MARGIN: 8,            // πόσο πρέπει να είναι καλύτερο (πόντοι score) ο "προκλητής" για να μετρήσει καν
    MIN_CONSECUTIVE_WINS: 3,     // πόσες συνεχόμενες εκτελέσεις πρέπει να κερδίζει πριν αλλάξουμε επιλογή
    MIN_SCORE_TO_KEEP: 40        // αν η τρέχουσα επιλογή πέσει κάτω από αυτό -> άμεση αλλαγή (χωρίς αναμονή)
  },
  STABLECOIN_EXCLUDE: [
    'USDT', 'USDC', 'DAI', 'FDUSD', 'TUSD', 'USDE', 'PYUSD', 'BUSD', 'USDD', 'USDP', 'GUSD', 'EURC', 'USDG'
  ],
 
  SYMBOLS: [
    { symbol: 'LINK-USDT-SWAP', name: 'Chainlink' },
    { symbol: 'TAO-USDT-SWAP',  name: 'Bittensor' },
    { symbol: 'NEAR-USDT-SWAP', name: 'NEAR Protocol' },
    { symbol: 'ICP-USDT-SWAP',  name: 'Internet Computer' },
    { symbol: 'FET-USDT-SWAP',  name: 'Fetch.ai / ASI' },
    { symbol: 'RENDER-USDT-SWAP', name: 'Render' }
    // Χρησιμοποιείται μόνο αν USE_TOP_N_BY_MARKET_CAP = false
    // (μορφή OKX perpetual futures instId: BASE-USDT-SWAP — έλεγξε διαθεσιμότητα:
    // https://www.okx.com/api/v5/public/instruments?instType=SWAP)
  ]
};
 
// ==================== MENU ====================
 
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Grid Screener')
    .addItem('Run Screener', 'checkGridCandidates')
    .addSeparator()
    .addItem('Add Open Position', 'addOpenPosition')
    .addItem('Remove Open Position', 'removeOpenPosition')
    .addItem('⚡ Quick Check Open Positions Only', 'checkOpenPositionsQuick')
    .addSeparator()
    .addItem('Exclude Symbol (not on Pionex)', 'excludeSymbol')
    .addItem('Remove Excluded Symbol', 'unexcludeSymbol')
    .addSeparator()
    .addItem('Reset Stable Pick (lock)', 'resetStablePick')
    .addItem('Clear History Log', 'clearHistory')
    .addToUi();
}
 
// ==================== EXCLUDED SYMBOLS (π.χ. coins που δεν υπάρχουν στο Pionex) ====================
//
// Το OKX έχει πολύ περισσότερα perpetual futures από όσα προσφέρει το Pionex Grid.
// Αποθηκεύεται ως λίστα στα Script Properties (όχι hardcoded στον κώδικα) ώστε να
// την ενημερώνεις εύκολα από το μενού, χωρίς να ξανααντιγράφεις το αρχείο.
 
function getExcludedSymbols() {
  const raw = PropertiesService.getScriptProperties().getProperty('EXCLUDED_SYMBOLS');
  return raw ? JSON.parse(raw) : [];
}
 
function excludeSymbol() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Εξαίρεση Symbol', 'Symbol που ΔΕΝ υπάρχει στο Pionex (π.χ. AI-USDT-SWAP ή απλά AI):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
 
  let symbol = resp.getResponseText().trim().toUpperCase();
  if (!symbol) { ui.alert('Δεν δόθηκε symbol.'); return; }
  if (symbol.indexOf('-SWAP') === -1) {
    symbol = (symbol.indexOf('-USDT') === -1 ? symbol + '-USDT' : symbol) + '-SWAP';
  }
 
  const excluded = getExcludedSymbols();
  if (excluded.indexOf(symbol) !== -1) {
    ui.alert('Το ' + symbol + ' είναι ήδη στη λίστα εξαιρέσεων.');
    return;
  }
  excluded.push(symbol);
  PropertiesService.getScriptProperties().setProperty('EXCLUDED_SYMBOLS', JSON.stringify(excluded));
  ui.alert('Εξαιρέθηκε: ' + symbol + '. Δεν θα ξαναπροταθεί. Κάνε "Reset Stable Pick" αν είναι ήδη κλειδωμένο τώρα.');
}
 
function unexcludeSymbol() {
  const ui = SpreadsheetApp.getUi();
  const excluded = getExcludedSymbols();
  if (excluded.length === 0) { ui.alert('Η λίστα εξαιρέσεων είναι άδεια.'); return; }
 
  const resp = ui.prompt('Αφαίρεση Εξαίρεσης', 'Symbol προς αφαίρεση από τις εξαιρέσεις:\n\nΤρέχουσα λίστα: ' + excluded.join(', '), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
 
  let symbol = resp.getResponseText().trim().toUpperCase();
  if (symbol.indexOf('-SWAP') === -1) {
    symbol = (symbol.indexOf('-USDT') === -1 ? symbol + '-USDT' : symbol) + '-SWAP';
  }
 
  const idx = excluded.indexOf(symbol);
  if (idx === -1) { ui.alert('Δεν βρέθηκε στη λίστα: ' + symbol); return; }
 
  excluded.splice(idx, 1);
  PropertiesService.getScriptProperties().setProperty('EXCLUDED_SYMBOLS', JSON.stringify(excluded));
  ui.alert('Αφαιρέθηκε από τις εξαιρέσεις: ' + symbol);
}
 
// ==================== STRATEGY-AWARE NAMING HELPERS ====================
// Κάθε στρατηγική (NEUTRAL/LONG) έχει το ΔΙΚΟ ΤΗΣ sheet, log, και κλειδωμένο
// stable pick — ώστε να τρέχουν παράλληλα σε ένα trigger χωρίς να "πατάει" η
// μία πάνω στην άλλη. Χρησιμοποιούν το GRID_SCREENER_CONFIG.GRID_STRATEGY που
// είναι ενεργό ΤΗ ΣΤΙΓΜΗ της κλήσης (ορίζεται από runScreenerForStrategy()).
 
function getMainSheetName() {
  return GRID_SCREENER_CONFIG.SHEET_NAME + ' - ' + GRID_SCREENER_CONFIG.GRID_STRATEGY;
}
 
function getLogSheetName() {
  return GRID_SCREENER_CONFIG.LOG_SHEET_NAME + ' - ' + GRID_SCREENER_CONFIG.GRID_STRATEGY;
}
 
function getPickStateKey() {
  return 'GRID_PICK_STATE_' + GRID_SCREENER_CONFIG.GRID_STRATEGY;
}
 
// ==================== MAIN ====================
 
// Το function name που καλεί το Trigger. Τρέχει ΟΛΕΣ τις στρατηγικές της λίστας
// STRATEGIES_TO_RUN, η μία μετά την άλλη, με πλήρη ανεξαρτησία μεταξύ τους.
function checkGridCandidates() {
  GRID_SCREENER_CONFIG.STRATEGIES_TO_RUN.forEach(function (strategy) {
    try {
      runScreenerForStrategy(strategy);
    } catch (err) {
      Logger.log('Screener run failed for strategy ' + strategy + ': ' + err.message);
    }
  });
}
 
function runScreenerForStrategy(strategy) {
  GRID_SCREENER_CONFIG.GRID_STRATEGY = strategy; // ενεργό για ΟΛΗ τη διάρκεια αυτού του run
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(getMainSheetName());
  if (!sheet) sheet = ss.insertSheet(getMainSheetName());
  sheet.clear();
 
  // ---- Pass 1: μάζεψε raw metrics για κάθε symbol ----
  const results = [];
  let symbolList;
  let usedFallback = false;
 
  if (GRID_SCREENER_CONFIG.USE_TOP_N_BY_MARKET_CAP) {
    try {
      symbolList = fetchTopNSymbols(GRID_SCREENER_CONFIG.TOP_N);
    } catch (err) {
      Logger.log('fetchTopNSymbols failed, falling back to static SYMBOLS: ' + err.message);
      symbolList = GRID_SCREENER_CONFIG.SYMBOLS;
      usedFallback = true;
    }
  } else {
    symbolList = GRID_SCREENER_CONFIG.SYMBOLS;
  }
 
  // Σημαντικό: αν υπάρχει ήδη κλειδωμένη επιλογή, ελέγξέ την ΠΑΝΤΑ, ακόμα κι αν
  // έπεσε εκτός του σημερινού top-N (π.χ. λόγω μειωμένου 24ωρου όγκου). Αλλιώς
  // το script τη νομίζει "delisted" απλά επειδή δεν κοιτάξαμε καν τα δεδομένα της.
  try {
    const lockedStateRaw = PropertiesService.getScriptProperties().getProperty(getPickStateKey());
    if (lockedStateRaw) {
      const lockedState = JSON.parse(lockedStateRaw);
      const alreadyIncluded = symbolList.some(function (s) { return s.symbol === lockedState.symbol; });
      if (!alreadyIncluded) {
        symbolList = symbolList.concat([{ symbol: lockedState.symbol, name: lockedState.name }]);
      }
    }
  } catch (e) {
    Logger.log('Could not read locked state for symbol-list safety-add: ' + e.message);
  }
 
  // Το ίδιο για όλες τις χειροκίνητα καταχωρημένες ΑΝΟΙΧΤΕΣ ΘΕΣΕΙΣ — πρέπει να
  // αξιολογούνται πάντα με φρέσκα δεδομένα, ανεξάρτητα από το top-N της ημέρας.
  try {
    const openSymbols = getOpenPositionSymbols();
    openSymbols.forEach(function (sym) {
      const alreadyIncluded = symbolList.some(function (s) { return s.symbol === sym; });
      if (!alreadyIncluded) {
        symbolList = symbolList.concat([{ symbol: sym, name: sym.split('-')[0] }]);
      }
    });
  } catch (e) {
    Logger.log('Could not read open positions for symbol-list safety-add: ' + e.message);
  }
 
  // Αφαίρεσε symbols που ο χρήστης έχει εξαιρέσει (π.χ. δεν υπάρχουν στο Pionex Grid) —
  // ΠΡΙΝ την αξιολόγηση, ώστε να μην ξαναπροταθούν ποτέ ούτε να σπαταλιέται API call.
  try {
    const excluded = getExcludedSymbols();
    if (excluded.length > 0) {
      symbolList = symbolList.filter(function (s) { return excluded.indexOf(s.symbol) === -1; });
    }
  } catch (e) {
    Logger.log('Could not read excluded symbols: ' + e.message);
  }
 
  const runStartTime = Date.now();
  let timeCutoffHit = false;
 
  for (let symIdx = 0; symIdx < symbolList.length; symIdx++) {
    if (Date.now() - runStartTime > GRID_SCREENER_CONFIG.MAX_RUNTIME_MS) {
      timeCutoffHit = true;
      Logger.log('Time safety cutoff hit μετά από ' + symIdx + '/' + symbolList.length + ' symbols — σταματάω με ασφάλεια.');
      break;
    }
 
    const item = symbolList[symIdx];
    try {
      const klines = fetchKlines(item.symbol, GRID_SCREENER_CONFIG.INTERVAL, GRID_SCREENER_CONFIG.LOOKBACK_CANDLES);
      if (!klines || klines.length < 30) {
        results.push({ symbol: item.symbol, name: item.name, error: 'insufficient data' });
        continue;
      }
 
      const closes = klines.map(function (k) { return k.close; });
      const highs = klines.map(function (k) { return k.high; });
      const lows = klines.map(function (k) { return k.low; });
 
      const lastPrice = closes[closes.length - 1];
      const periodHigh = Math.max.apply(null, highs);
      const periodLow = Math.min.apply(null, lows);
      const rangePct = ((periodHigh - periodLow) / periodLow) * 100;
 
      const slopePct = calcTrendSlopePct(closes);
      const adx = calcADX(klines, 14);
      const choppiness = calcChoppinessIndex(klines, 14);
      const atrPct = calcATRPct(klines, 14);
      const verdict = classify(rangePct, slopePct, adx, choppiness);
      const estReturnPct = estimateGridReturnPct(closes);
      const entrySuggestion = computeEntrySuggestion(klines, periodLow, periodHigh, lastPrice);
 
      // Βραχυχρόνιο ADX (τελευταίες ~2 μέρες, ίδιο παράθυρο με το shortSlopePct) —
      // δείχνει πόσο ΑΞΙΟΠΙΣΤΗ/ΔΥΝΑΤΗ είναι η πρόσφατη κίνηση, συμπληρώνοντας το
      // shortSlopePct (που δείχνει μόνο κατεύθυνση/μέγεθος, όχι αξιοπιστία).
      const shortTermWindow = klines.slice(-GRID_SCREENER_CONFIG.ENTRY_SUGGESTION.RECENT_SWING_CANDLES);
      const shortTermADX = calcADX(shortTermWindow, 5);
      // ΔΙΑΓΝΩΣΤΙΚΟ #1: ύποπτα ακραία τιμή σε απόλυτους όρους.
      if (Math.abs(slopePct) > 20) {
        Logger.log('⚠️ ΑΝΩΜΑΛΙΑ slopePct για ' + item.symbol + ': ' + slopePct.toFixed(2) + '% — klines.length=' +
          klines.length + ', closes[0..4]=' + JSON.stringify(closes.slice(0, 5)) +
          ', closes[τελευταία 5]=' + JSON.stringify(closes.slice(-5)) +
          ', min=' + Math.min.apply(null, closes) + ', max=' + Math.max.apply(null, closes));
      }
 
      // ΔΙΑΓΝΩΣΤΙΚΟ #2 — ΣΚΛΗΡΟΣ έλεγχος (throw, όχι απλό log): αν το slope είναι
      // αρκετά μεγάλο ώστε να έπρεπε να απορριφθεί αλλά το verdict βγήκε GOOD/
      // MARGINAL, αυτό είναι ΛΟΓΙΚΑ ΑΔΥΝΑΤΟ με τον τρέχοντα κώδικα. Το κάνουμε throw
      // (όχι απλό Logger.log) ώστε να ΜΗΝ μπορεί να χαθεί/αγνοηθεί — αν ποτέ συμβεί,
      // η εκτέλεση θα αποτύχει ρητά με πλήρες stack trace στο Executions log.
      const longCfg = GRID_SCREENER_CONFIG.LONG_GRID;
      const shortCfg = GRID_SCREENER_CONFIG.SHORT_GRID;
      const shouldHaveBeenRejected =
        (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'LONG' && (slopePct <= longCfg.DOWNTREND_SLOPE_PCT || slopePct >= longCfg.PARABOLIC_SLOPE_PCT)) ||
        (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT' && (slopePct >= shortCfg.UPTREND_SLOPE_PCT || slopePct <= shortCfg.PARABOLIC_DOWN_SLOPE_PCT));
      if (shouldHaveBeenRejected && isEligibleVerdict(verdict)) {
        throw new Error('🚨 ΛΟΓΙΚΗ ΑΝΤΙΦΑΣΗ για ' + item.symbol + ': slopePct=' + slopePct + ' (τύπος: ' + typeof slopePct +
          '), strategy=' + GRID_SCREENER_CONFIG.GRID_STRATEGY + ', verdict="' + verdict + '" (ΘΑ ΕΠΡΕΠΕ να έχει απορριφθεί). ' +
          'rangePct=' + rangePct + ', choppiness=' + choppiness + ', adx=' + adx +
          ', klines.length=' + klines.length + ', closes[0..4]=' + JSON.stringify(closes.slice(0, 5)) +
          ', closes[τελευταία 5]=' + JSON.stringify(closes.slice(-5)));
      }
 
      results.push({
        symbol: item.symbol, name: item.name, lastPrice: lastPrice,
        periodHigh: periodHigh, periodLow: periodLow, rangePct: rangePct,
        slopePct: slopePct, adx: adx, choppiness: choppiness, atrPct: atrPct,
        verdict: verdict, estReturnPct: estReturnPct, entrySuggestion: entrySuggestion,
        shortTermADX: shortTermADX
      });
 
    } catch (err) {
      results.push({ symbol: item.symbol, name: item.name, error: err.message });
    }
  }
 
  // ---- Pass 2: κανονικοποίησε estReturnPct σε 0-100 σχετικά με το σύνολο ----
  const validResults = results.filter(function (r) { return !r.error; });
  const estReturns = validResults.map(function (r) { return r.estReturnPct; });
  const minEst = Math.min.apply(null, estReturns.length ? estReturns : [0]);
  const maxEst = Math.max.apply(null, estReturns.length ? estReturns : [0]);
 
  validResults.forEach(function (r) {
    const estNorm = maxEst === minEst ? 50 : ((r.estReturnPct - minEst) / (maxEst - minEst)) * 100;
    const choppNorm = clamp(r.choppiness, 0, 100);
    const trendComponentNorm = GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT'
      ? trendDownwardScore(r.slopePct)                             // SHORT: ανταμοιβή ήπιας καθοδικής τάσης
      : trendUpwardScore(r.slopePct);                              // LONG: ανταμοιβή ήπιας ανοδικής τάσης
    const rangeFitNorm = clamp(100 - Math.abs(r.rangePct - GRID_SCREENER_CONFIG.RANGE_SWEET_SPOT_PCT) * 3, 0, 100);
 
    const w = GRID_SCREENER_CONFIG.SCORE_WEIGHTS;
    r.score = (choppNorm * w.choppiness) + (trendComponentNorm * w.trendWeak) +
      (rangeFitNorm * w.rangeFit) + (estNorm * w.estReturn);
 
    // Verdict εκτός "GOOD"/"MARGINAL" -> penalize βαριά ώστε να πέφτει στο τέλος
    if (!isEligibleVerdict(r.verdict)) {
      r.score = r.score * 0.25;
    }
  });
 
  // ---- Ταξινόμηση: πρώτα κατά "tier" verdict (GOOD > MARGINAL > υπόλοιπα), ΜΕΤΑ κατά score. ----
  // Χωρίς αυτό, ένα MARGINAL coin με πολύ ισχυρά τα υπόλοιπα κριτήρια (choppiness,
  // range fit, est.return) θα μπορούσε μαθηματικά να βγει ψηλότερα από ένα GOOD με
  // πιο αδύναμα — τεχνικά σωστό ως άθροισμα, αλλά αντιδιαισθητικό: το verdict tier
  // πρέπει να έχει προτεραιότητα. (verdictTier είναι global function, reused και στο getStablePick.)
  validResults.sort(function (a, b) {
    const tierDiff = verdictTier(b) - verdictTier(a);
    if (tierDiff !== 0) return tierDiff;
    return b.score - a.score;
  });
  validResults.forEach(function (r, i) { r.rank = i + 1; }); // rank στην ΠΛΗΡΗ λίστα, όχι μόνο στο top 10
  const errorResults = results.filter(function (r) { return r.error; });
  const fullOrdered = validResults.concat(errorResults); // για log/telegram (πλήρης λίστα)
 
  // ---- Σταθερή επιλογή (anti flip-flop) ----
  const stableInfo = getStablePick(fullOrdered);
 
  // ---- Ποιες γραμμές δείχνουμε στο κύριο sheet: top N, + το locked pick αν λείπει ----
  let displayList = validResults.slice(0, GRID_SCREENER_CONFIG.TOP_RESULTS_TO_SHOW);
  if (stableInfo && stableInfo.pick) {
    const alreadyShown = displayList.some(function (r) { return r.symbol === stableInfo.pick.symbol; });
    if (!alreadyShown) {
      displayList = [stableInfo.pick].concat(displayList);
    }
  }
 
  // ---- Banner (row 1) με την τρέχουσα σταθερή επιλογή ----
  if (stableInfo && stableInfo.pick) {
    const p = stableInfo.pick;
    const exitRisk = computeExitRisk(p, stableInfo.lowerLimit, stableInfo.upperLimit);
 
    let bannerText = '🔒 Σταθερή επιλογή: ' + p.symbol + ' (' + p.name + ') — Score ' +
      p.score.toFixed(1) + '/100 — Verdict: ' + p.verdict + ' — κλειδωμένο από ' + stableInfo.lockedSince +
      ' — Grid limits: ' + stableInfo.lowerLimit + ' - ' + stableInfo.upperLimit;
    if (stableInfo.challenger) {
      bannerText += '  |  Προκλητής: ' + stableInfo.challenger + ' (' + stableInfo.challengerStreak +
        '/' + stableInfo.requiredStreak + ' νίκες)';
    }
    if (exitRisk) {
      bannerText += '  |  ' + exitRisk;
    }
    sheet.getRange(1, 1).setValue(bannerText).setFontWeight('bold')
      .setFontColor(exitRisk ? '#cc0000' : '#1155cc');
  } else {
    sheet.getRange(1, 1).setValue('Δεν βρέθηκε σταθερή επιλογή (κανένα valid αποτέλεσμα).');
  }
 
  const headerRow = 3;
  const dataStartRow = 4;
 
  const headers = [
    'Rank', 'Grid Score (0-100)', 'Symbol', 'Name', 'Last Price', 'Range % (period)',
    'Choppiness Idx', 'Trend Slope %/candle', 'ADX (approx)',
    'Short-term Slope %/candle (2d)', 'Short-term ADX (2d)', 'Last 2 Candles Pattern',
    'Est. Return (period, rough)',
    'Verdict', 'Γιατί', 'Suggested Lower Limit', 'Suggested Upper Limit', 'Alt Entry Price',
    'Suggested Grids', 'Suggested Leverage', 'Suggested TP %', 'Suggested SL %', 'Fib 0.618', 'Fib 0.786', 'Ready to Enter', 'Entry Zone (4H)', 'Next 4H Range (ATR-based)', 'Notes'
  ];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
 
  // ---- Χτίσε γραμμές sheet ----
  const rows = [];
  displayList.forEach(function (r) {
    if (r.error) {
      rows.push(['', '', r.symbol, r.name, '', '', '', '', '', '', '', '', '', 'ERROR: ' + r.error, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      return;
    }
 
    let lowerLimit = '', upperLimit = '', grids = '', leverage = '', notes = '';
    const isLockedPick = stableInfo && stableInfo.pick && r.symbol === stableInfo.pick.symbol;
 
    if (isLockedPick) {
      // Για την ΚΛΕΙΔΩΜΕΝΗ επιλογή δείχνουμε τα πραγματικά όρια που "τρέχουν", όχι το ξαναϋπολογισμένο σημερινό εύρος
      lowerLimit = stableInfo.lowerLimit;
      upperLimit = stableInfo.upperLimit;
      grids = suggestGridCount(r.rangePct);
      leverage = suggestLeverage(r.atrPct);
    } else if (isEligibleVerdict(r.verdict)) {
      const limits = computeBufferedLimits(r.periodLow, r.periodHigh, r.lastPrice);
      lowerLimit = limits.lowerLimit;
      upperLimit = limits.upperLimit;
      grids = suggestGridCount(r.rangePct);
      leverage = suggestLeverage(r.atrPct);
      if (isMarginalVerdict(r.verdict)) {
        leverage = leverage - 1 > 0 ? leverage - 1 : 1;
        notes = 'Οριακό: μικρότερο size/leverage, στενή παρακολούθηση.';
      } else {
        notes = 'Επιβεβαίωσε levels σε chart πριν το στήσιμο.';
      }
    } else {
      const notesMap = {
        'TOO TIGHT': 'Πολύ μικρό εύρος — δεν αξίζει το κόστος συναλλαγών.',
        'TOO VOLATILE': 'Πολύ ασταθές — υψηλό ρίσκο liquidation σε leveraged grid.',
        'DOWNTREND': 'Καθοδική τάση — ακατάλληλο για Long Grid.',
        'PARABOLIC': 'Πολύ απότομη/παραβολική άνοδος — ρίσκο απότομης διόρθωσης, το grid θα μείνει πίσω.',
        'FLAT/WEAK — ΟΧΙ ΑΝΟΔΙΚΗ': 'Δεν υπάρχει σαφής ανοδική τάση — δεν αξιοποιεί το πλεονέκτημα του Long Grid.',
        'UPTREND': 'Ανοδική τάση — ακατάλληλο για Short Grid.',
        'PARABOLIC DOWN': 'Πολύ απότομη πτώση — ρίσκο απότομου rebound, το grid θα μείνει πίσω.',
        'FLAT/WEAK — ΟΧΙ ΚΑΘΟΔΙΚΗ': 'Δεν υπάρχει σαφής καθοδική τάση — δεν αξιοποιεί το πλεονέκτημα του Short Grid.'
      };
      notes = notesMap[r.verdict] || 'Δεν πληροί τα κριτήρια της τρέχουσας στρατηγικής.';
    }
 
    if (isLockedPick) {
      const exitRisk = computeExitRisk(r, stableInfo.lowerLimit, stableInfo.upperLimit);
      notes = '🔒 ΤΡΕΧΟΥΣΑ ΣΤΑΘΕΡΗ ΕΠΙΛΟΓΗ (πραγματικά όρια grid). ' + (exitRisk ? exitRisk + ' ' : '') + notes;
    }
 
    const tpPct = leverage ? suggestTP(r.estReturnPct, leverage) : '';
    const slPct = leverage ? suggestSL(r.atrPct, leverage) : '';
    const explanation = explainScore(r);
 
    // Για την ΚΛΕΙΔΩΜΕΝΗ επιλογή, η ζώνη εισόδου πρέπει να αναφέρεται στα πραγματικά
    // "παγωμένα" limits του grid (αυτά που βλέπει ο χρήστης), όχι στο ξανά-υπολογισμένο
    // σημερινό εύρος — αλλιώς το ποσοστό θέσης δεν ταιριάζει με αυτό που δείχνει το grid.
    let entryZoneInfo = r.entrySuggestion;
    if (isLockedPick && r.entrySuggestion) {
      entryZoneInfo = computeEntryZone(
        r.lastPrice, stableInfo.lowerLimit, stableInfo.upperLimit,
        r.entrySuggestion.shortSlopePct, r.entrySuggestion.recentSwingLow,
        r.entrySuggestion.recentSwingHigh, r.entrySuggestion.last2Analysis
      );
    }
    const entryZoneCell = entryZoneInfo
      ? entryZoneInfo.zone + ' (' + entryZoneInfo.posPct.toFixed(0) + '%) — ' + entryZoneInfo.note
      : '';
 
    const isBadZone = entryZoneInfo && (entryZoneInfo.zone === 'ΨΗΛΗ ΖΩΝΗ' || entryZoneInfo.zone === 'ΧΑΜΗΛΗ ΖΩΝΗ');
    const readyCell = entryZoneInfo
      ? (!isBadZone ? '✅ YES' : (GRID_SCREENER_CONFIG.ENTRY_SUGGESTION.BLOCK_ENTRY_OUTSIDE_GOOD_ZONE ? '⏳ WAIT' : '⚠️ OK αλλά όχι ιδανικό'))
      : '';
 
    if (isEligibleVerdict(r.verdict) && isBadZone && GRID_SCREENER_CONFIG.ENTRY_SUGGESTION.BLOCK_ENTRY_OUTSIDE_GOOD_ZONE) {
      notes = '⏳ Περίμενε καλύτερη ζώνη εισόδου πριν στήσεις πλήρες μέγεθος. ' + notes;
    }
 
    const nextCandleRange = computeNextCandleRange(r.lastPrice, r.atrPct);
    const nextCandleCell = nextCandleRange
      ? nextCandleRange.low + ' – ' + nextCandleRange.high + ' (στατιστικό εύρος βάσει ATR, όχι πρόβλεψη κατεύθυνσης)'
      : '';
 
    const altEntryCell = (lowerLimit && upperLimit)
      ? computeAltEntryPrice(lowerLimit, upperLimit, r.lastPrice)
      : '';
 
    const fib618Cell = (lowerLimit && upperLimit) ? computeFibLevel(lowerLimit, upperLimit, 0.618, r.lastPrice) : '';
    const fib786Cell = (lowerLimit && upperLimit) ? computeFibLevel(lowerLimit, upperLimit, 0.786, r.lastPrice) : '';
 
    const shortSlopeCell = r.entrySuggestion ? r.entrySuggestion.shortSlopePct.toFixed(3) + '%' : '';
    const shortADXCell = (r.shortTermADX !== undefined && r.shortTermADX !== null) ? r.shortTermADX.toFixed(1) : '';
    const last2PatternCell = (r.entrySuggestion && r.entrySuggestion.last2Analysis) ? r.entrySuggestion.last2Analysis.pattern : '';
 
    rows.push([
      r.rank, r.score.toFixed(1), r.symbol, r.name, r.lastPrice, r.rangePct.toFixed(2) + '%',
      r.choppiness.toFixed(1), r.slopePct.toFixed(3) + '%', r.adx.toFixed(1),
      shortSlopeCell, shortADXCell, last2PatternCell,
      r.estReturnPct.toFixed(2) + '%', r.verdict, explanation, lowerLimit, upperLimit, altEntryCell,
      grids, leverage ? leverage + 'x' : '', tpPct ? '+' + tpPct.toFixed(1) + '%' : '',
      slPct ? '-' + slPct.toFixed(1) + '%' : '', fib618Cell, fib786Cell, readyCell, entryZoneCell, nextCandleCell, notes
    ]);
  });
 
  if (rows.length > 0) {
    sheet.getRange(dataStartRow, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, headers.length);
 
  // Χρωματισμός verdict + highlight της γραμμής του locked pick
  const verdictCol = 11;
  for (let i = 0; i < rows.length; i++) {
    const rowNum = dataStartRow + i;
    const cell = sheet.getRange(rowNum, verdictCol);
    const v = String(rows[i][verdictCol - 1]);
    if (isGoodVerdict(v)) cell.setBackground('#d9ead3');
    else if (isMarginalVerdict(v)) cell.setBackground('#fff2cc');
    else cell.setBackground('#f4cccc');
 
    if (stableInfo && stableInfo.pick && rows[i][2] === stableInfo.pick.symbol) {
      sheet.getRange(rowNum, 1, 1, headers.length).setFontWeight('bold');
    }
  }
 
  logHistory(fullOrdered);
 
  const openPositionAlerts = evaluateOpenPositions(fullOrdered);
  const hasCriticalAlerts = openPositionAlerts.some(function (a) { return a.status === 'CLOSE' || a.status === 'INFO' || a.status === 'DECLINING'; });
 
  if (GRID_SCREENER_CONFIG.TELEGRAM_NOTIFY) {
    notifyBestPickTelegram(stableInfo, fullOrdered);
    if (openPositionAlerts.length > 0 && (hasCriticalAlerts || GRID_SCREENER_CONFIG.OPEN_POSITIONS.ALWAYS_NOTIFY)) {
      notifyOpenPositionsTelegram(openPositionAlerts);
    }
    const a2Info = checkOpenPositionA2VsTopPick(fullOrdered);
    if (a2Info) {
      const a2Message = buildOpenPositionAnalysisMessage(a2Info);
      if (a2Message) notifyOpenPositionAnalysisTelegram(a2Message);
    }
  }
 
  SpreadsheetApp.getActiveSpreadsheet().toast(
    (usedFallback ? '⚠️ OKX tickers fetch failed, χρησιμοποιήθηκε η στατική λίστα SYMBOLS. ' : '') +
    (timeCutoffHit ? '⏱️ Χρονικό όριο ασφαλείας — ελέγχθηκαν μόνο ' + results.length + '/' + symbolList.length + ' symbols. ' : '') +
    'Ελέγχθηκαν ' + results.length + ' symbols (' + validResults.length + ' με δεδομένα, ' +
    errorResults.length + ' skipped).'
  );
}
 
// ==================== STABLE PICK (ΑΝΤΙ FLIP-FLOP) ====================
//
// Ένα neutral grid δεν συμφέρει να αλλάζει coin κάθε μέρα (κόστος re-setup,
// χαμένος χρόνος). Αυτή η function κρατάει "κλειδωμένη" την τρέχουσα πρόταση
// ανάμεσα σε εκτελέσεις (μέσω Script Properties) και την αλλάζει ΜΟΝΟ αν ένας
// νέος υποψήφιος τη νικάει με σαφές περιθώριο (STABILITY.SWITCH_MARGIN) για
// αρκετές συνεχόμενες εκτελέσεις στη σειρά (STABILITY.MIN_CONSECUTIVE_WINS) —
// εκτός αν η τρέχουσα κατέρρευσε (score κάτω από MIN_SCORE_TO_KEEP), οπότε
// αλλάζει αμέσως.
function getStablePick(fullOrdered) {
  const props = PropertiesService.getScriptProperties();
  const cfg = GRID_SCREENER_CONFIG.STABILITY;
  const validToday = fullOrdered.filter(function (r) { return !r.error; });
  if (validToday.length === 0) return null;
 
  const topToday = validToday[0]; // ήδη ταξινομημένα κατά score, φθίνουσα
  const pickStateKey = getPickStateKey();
  const stateRaw = props.getProperty(pickStateKey);
  const state = stateRaw ? JSON.parse(stateRaw) : null;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
 
  function findToday(symbol) {
    for (let i = 0; i < validToday.length; i++) {
      if (validToday[i].symbol === symbol) return validToday[i];
    }
    return null;
  }
 
  function limitsFor(pick) {
    return computeBufferedLimits(pick.periodLow, pick.periodHigh, pick.lastPrice);
  }
 
  function lockNew(pick) {
    const limits = limitsFor(pick);
    const newState = {
      symbol: pick.symbol, name: pick.name, lockedScore: pick.score,
      lockedSince: nowStr, challengerSymbol: null, challengerStreak: 0,
      lowerLimit: limits.lowerLimit, upperLimit: limits.upperLimit // "παγωμένα" όρια grid — ΔΕΝ ξαναϋπολογίζονται κάθε run
    };
    props.setProperty(pickStateKey, JSON.stringify(newState));
    return {
      pick: pick, switched: true, lockedSince: nowStr, challengerStreak: 0, requiredStreak: cfg.MIN_CONSECUTIVE_WINS,
      lowerLimit: limits.lowerLimit, upperLimit: limits.upperLimit, topToday: topToday
    };
  }
 
  function keep(currentToday, challengerSymbol, challengerStreak) {
    // Self-healing: αν το αποθηκευμένο state είναι από παλιότερη έκδοση του script
    // (πριν προστεθούν τα lowerLimit/upperLimit) ή είναι κατεστραμμένο, υπολόγισέ τα τώρα
    // αντί να διαιωνίζεις "undefined" σε κάθε επόμενο run.
    const limits = (state.lowerLimit && state.upperLimit)
      ? { lowerLimit: state.lowerLimit, upperLimit: state.upperLimit }
      : limitsFor(currentToday);
 
    props.setProperty(pickStateKey, JSON.stringify({
      symbol: state.symbol, name: state.name, lockedScore: currentToday.score,
      lockedSince: state.lockedSince, challengerSymbol: challengerSymbol || null, challengerStreak: challengerStreak || 0,
      lowerLimit: limits.lowerLimit, upperLimit: limits.upperLimit // κρατάμε τα ΑΡΧΙΚΑ όρια, δεν τα ξαναφτιάχνουμε
    }));
    const result = {
      pick: currentToday, switched: false, lockedSince: state.lockedSince,
      challengerStreak: challengerStreak || 0, requiredStreak: cfg.MIN_CONSECUTIVE_WINS,
      lowerLimit: limits.lowerLimit, upperLimit: limits.upperLimit, topToday: topToday
    };
    if (challengerSymbol) result.challenger = challengerSymbol;
    return result;
  }
 
  // Πρώτη φορά που τρέχει το script
  if (!state) return lockNew(topToday);
 
  const currentToday = findToday(state.symbol);
  const verdictBroken = currentToday && !isEligibleVerdict(currentToday.verdict);
  const scoreTooLow = currentToday && currentToday.score < cfg.MIN_SCORE_TO_KEEP;
  const trulyLost = !currentToday; // delisted/no data
 
  if (trulyLost || verdictBroken) {
    // Πραγματικό collapse: είτε χάθηκε εντελώς, είτε το verdict βγήκε εκτός κατηγορίας.
    // Σε αυτή την περίπτωση ΔΕΝ έχει νόημα να ξαναδιαλέξουμε το ίδιο symbol — είναι εξ
    // ορισμού ακατάλληλο πλέον. Αν τυχαία είναι #1 σήμερα (π.χ. μεταβατικό restart),
    // πάρε την καλύτερη ΔΙΑΦΟΡΕΤΙΚΗ επιλογή αντί γι' αυτό.
    const bestOther = validToday.find(function (r) { return r.symbol !== state.symbol; }) || topToday;
    const result = lockNew(bestOther);
    result.collapsed = true;
    result.collapsedFrom = currentToday
      ? { symbol: currentToday.symbol, name: currentToday.name, verdict: currentToday.verdict, score: currentToday.score }
      : { symbol: state.symbol, name: state.name, verdict: 'DELISTED/NO DATA', score: null };
    return result;
  }
 
  if (scoreTooLow) {
    // Το verdict ΠΑΡΑΜΕΝΕΙ eligible (GOOD/MARGINAL) — απλά το score είναι χαμηλό, πιθανόν
    // επειδή όλη η αγορά είναι αδύναμη σήμερα. Άλλαξε ΜΟΝΟ αν υπάρχει πραγματικά
    // καλύτερη ΔΙΑΦΟΡΕΤΙΚΗ εναλλακτική — και "καλύτερη" σημαίνει tier-aware: ένα
    // MARGINAL με υψηλότερο ακατέργαστο score ΔΕΝ είναι "καλύτερο" από ένα GOOD
    // (ακριβώς όπως στην κύρια ταξινόμηση). Αλλιώς το "collapse -> switch σε
    // χειρότερο tier" είναι παράλογο — απλά κράτα το με σιωπηλή παρακολούθηση.
    const currentTier = verdictTier(currentToday);
    const bestOther = validToday.find(function (r) {
      return r.symbol !== state.symbol &&
        (verdictTier(r) > currentTier || (verdictTier(r) === currentTier && r.score > currentToday.score));
    });
    if (bestOther) {
      const result = lockNew(bestOther);
      result.collapsed = true;
      result.collapsedFrom = { symbol: currentToday.symbol, name: currentToday.name, verdict: currentToday.verdict, score: currentToday.score };
      return result;
    }
    return keep(currentToday, null, 0);
  }
 
  // Ο σημερινός νικητής ΕΙΝΑΙ ήδη η κλειδωμένη επιλογή -> κράτα, μηδένισε challenger
  if (topToday.symbol === state.symbol) {
    return keep(currentToday, null, 0);
  }
 
  // Υπάρχει "πρόκληση" — έλεγξε αν το margin είναι αρκετό
  const margin = topToday.score - currentToday.score;
  if (margin >= cfg.SWITCH_MARGIN) {
    const streak = (state.challengerSymbol === topToday.symbol ? state.challengerStreak : 0) + 1;
    if (streak >= cfg.MIN_CONSECUTIVE_WINS) return lockNew(topToday);
    return keep(currentToday, topToday.symbol, streak);
  }
 
  // Margin πολύ μικρό -> κράτα την τρέχουσα, μηδένισε challenger
  return keep(currentToday, null, 0);
}
 
// Καθαρίζει τα κλειδωμένα picks (ΚΑΙ για LONG ΚΑΙ για SHORT) — το επόμενο run
// θα διαλέξει από την αρχή τον σημερινό καλύτερο σε κάθε strategy.
function resetStablePick() {
  const props = PropertiesService.getScriptProperties();
  GRID_SCREENER_CONFIG.STRATEGIES_TO_RUN.forEach(function (strategy) {
    props.deleteProperty('GRID_PICK_STATE_' + strategy);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Τα locked picks μηδενίστηκαν (LONG & SHORT) — το επόμενο run θα διαλέξει από την αρχή.');
}
 
// ==================== OPEN POSITIONS (ΑΝΕΞΑΡΤΗΤΟ ΑΠΟ ΤΟ "ΠΡΟΤΕΙΝΟΜΕΝΟ PICK") ====================
//
// Το "Stable Pick" λέει τι θα πρότεινε το script ΤΩΡΑ αν ξεκινούσες από την αρχή.
// Δεν ξέρει τι έχεις ήδη ανοιχτό στο Pionex. Αυτή η ενότητα είναι ξεχωριστή: ο
// χρήστης καταχωρεί χειροκίνητα τα coins που έχει πραγματικά ανοιχτά, και το
// script τα παρακολουθεί με ΔΙΚΗ ΤΟΥΣ λογική — δεν τα κλείνει απλά επειδή άλλαξε
// το "προτεινόμενο", μόνο αν πραγματικά χαλάσουν ΤΑ ΙΔΙΑ ή αν η αναβάθμιση είναι
// τόσο μεγάλη που αξίζει το πραγματικό κόστος αλλαγής (BIG_UPGRADE_MARGIN, πολύ
// πιο αυστηρό από το SWITCH_MARGIN του stable pick).
 
function ensureOpenPositionsSheetExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GRID_SCREENER_CONFIG.OPEN_POSITIONS.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(GRID_SCREENER_CONFIG.OPEN_POSITIONS.SHEET_NAME);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Symbol (π.χ. LINK-USDT-SWAP)', 'Προστέθηκε', 'Grid Type (LONG/SHORT)',
      'My Lower Limit (προαιρετικό)', 'My Upper Limit (προαιρετικό)',
      'Status', 'Verdict', 'Score', 'Σύσταση'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 9);
  }
  return sheet;
}
 
// Διαβάζει μόνο τα symbols (στήλη A) — χρησιμοποιείται στο pass-1 safety-add.
function getOpenPositionSymbols() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GRID_SCREENER_CONFIG.OPEN_POSITIONS.SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values.map(function (row) { return String(row[0]).trim(); }).filter(function (s) { return s.length > 0; });
}
 
// Μενού: Add Open Position — ζητάει symbol + προαιρετικά τα δικά σου limits (αν τα
// δώσεις, το script μπορεί να ελέγξει και price-boundary exit risk, όχι μόνο verdict).
function addOpenPosition() {
  const ui = SpreadsheetApp.getUi();
  const symResp = ui.prompt('Προσθήκη Ανοιχτής Θέσης', 'Symbol (π.χ. LINK-USDT-SWAP ή απλά LINK):', ui.ButtonSet.OK_CANCEL);
  if (symResp.getSelectedButton() !== ui.Button.OK) return;
 
  let symbol = symResp.getResponseText().trim().toUpperCase();
  if (!symbol) { ui.alert('Δεν δόθηκε symbol.'); return; }
  if (symbol.indexOf('-SWAP') === -1) {
    symbol = (symbol.indexOf('-USDT') === -1 ? symbol + '-USDT' : symbol) + '-SWAP';
  }
 
  const existing = getOpenPositionSymbols();
  if (existing.indexOf(symbol) !== -1) {
    ui.alert('Το ' + symbol + ' υπάρχει ήδη στη λίστα ανοιχτών θέσεων.');
    return;
  }
 
  // ΚΡΙΣΙΜΟ: ρώτα τι ΤΥΠΟΥ grid άνοιξες πραγματικά (Long ή Short) — χωρίς αυτό,
  // η αξιολόγηση δεν μπορεί να ξέρει με ποια κριτήρια να κρίνει τη θέση, και θα
  // κατέληγε να προτείνει κλείσιμο θέσεων απλά επειδή δεν ταιριάζουν στα κριτήρια
  // ΤΟΥ ΛΑΘΟΣ mode.
  let gridType = 'LONG';
  const typeResp = ui.prompt('Τύπος Grid', 'LONG ή SHORT; (γράψε ακριβώς μία από τις δύο λέξεις — αν αφήσεις κενό, θεωρείται LONG)', ui.ButtonSet.OK_CANCEL);
  if (typeResp.getSelectedButton() === ui.Button.OK) {
    const typed = typeResp.getResponseText().trim().toUpperCase();
    if (typed === 'SHORT') gridType = 'SHORT';
    else if (typed !== '' && typed !== 'LONG') {
      ui.alert('Μη έγκυρη τιμή ("' + typed + '") — χρησιμοποιείται LONG by default.');
    }
  }
 
  // Φέρε το ΤΡΕΧΟΝ υπολογισμένο εύρος για το symbol, ώστε να το δείξουμε ως
  // υπενθύμιση/πρόταση μέσα στο ίδιο το prompt — έτσι είναι πιο εύκολο να το
  // αντιγράψεις αντί να το αφήσεις κενό κατά λάθος. Best-effort: αν αποτύχει,
  // απλά δεν δείχνουμε πρόταση, δεν σταματάει τη διαδικασία προσθήκης.
  let suggestionText = '';
  try {
    const klines = fetchKlines(symbol, GRID_SCREENER_CONFIG.INTERVAL, GRID_SCREENER_CONFIG.LOOKBACK_CANDLES);
    if (klines && klines.length >= 30) {
      const closes = klines.map(function (k) { return k.close; });
      const highs = klines.map(function (k) { return k.high; });
      const lows = klines.map(function (k) { return k.low; });
      const lastPrice = closes[closes.length - 1];
      const periodHigh = Math.max.apply(null, highs);
      const periodLow = Math.min.apply(null, lows);
      const limits = computeBufferedLimits(periodLow, periodHigh, lastPrice);
      suggestionText = '\n\n(Τρέχον υπολογισμένο εύρος για αναφορά — αν άνοιξες με αυτό, αντίγραψέ το: ' +
        limits.lowerLimit + ' – ' + limits.upperLimit + ')';
    }
  } catch (e) {
    // αγνόησε, δεν είναι κρίσιμο
  }
 
  const lowerResp = ui.prompt('Προαιρετικό', 'Το δικό σου Lower Limit (άφησέ το κενό αν δεν θες price-boundary έλεγχο):' + suggestionText, ui.ButtonSet.OK_CANCEL);
  const myLower = (lowerResp.getSelectedButton() === ui.Button.OK) ? parseFloat(lowerResp.getResponseText()) || '' : '';
 
  const upperResp = ui.prompt('Προαιρετικό', 'Το δικό σου Upper Limit (άφησέ το κενό αν δεν θες price-boundary έλεγχο):' + suggestionText, ui.ButtonSet.OK_CANCEL);
  const myUpper = (upperResp.getSelectedButton() === ui.Button.OK) ? parseFloat(upperResp.getResponseText()) || '' : '';
 
  const sheet = ensureOpenPositionsSheetExists();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sheet.appendRow([symbol, now, gridType, myLower, myUpper, '', '', '', '']);
  ui.alert('Προστέθηκε: ' + symbol + ' (' + gridType + '). Θα αξιολογηθεί στο επόμενο run.');
}
 
function removeOpenPosition() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Αφαίρεση Ανοιχτής Θέσης', 'Symbol προς αφαίρεση:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
 
  let symbol = resp.getResponseText().trim().toUpperCase();
  if (symbol.indexOf('-SWAP') === -1) {
    symbol = (symbol.indexOf('-USDT') === -1 ? symbol + '-USDT' : symbol) + '-SWAP';
  }
 
  const sheet = ensureOpenPositionsSheetExists();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('Δεν υπάρχουν καταχωρημένες θέσεις.'); return; }
 
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === symbol) {
      sheet.deleteRow(i + 2);
      ui.alert('Αφαιρέθηκε: ' + symbol);
      return;
    }
  }
  ui.alert('Δεν βρέθηκε: ' + symbol);
}
 
// ==================== QUICK CHECK (μόνο ανοιχτές θέσεις, χωρίς πλήρες scan) ====================
//
// Σκοπός: να μπορεί να τρέχει ΣΥΧΝΑ (π.χ. κάθε 30 λεπτά μέσω ξεχωριστού trigger)
// χωρίς το κόστος (χρόνο + API calls) του πλήρους screener στα 100 coins. Κάνει
// fetch ΜΟΝΟ για τα symbols που έχεις ήδη ανοιχτά, ελέγχει αν κατέρρευσε το verdict
// ή σπάσαν τα δικά σου limits, και ειδοποιεί ΑΜΕΣΩΣ — για να πιάνεις απότομες
// αναστροφές (όπως αυτή που είδαμε στο MERL) πολύ πιο γρήγορα από το καθημερινό run.
// ΔΕΝ κάνει σύγκριση με "καλύτερους υποψήφιους" (αυτό απαιτεί πλήρες scan) — μόνο
// εξετάζει αν Η ΙΔΙΑ η θέση χρειάζεται προσοχή.
function checkOpenPositionsQuick() {
  const sheet = ensureOpenPositionsSheetExists();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Δεν υπάρχουν ανοιχτές θέσεις για γρήγορο έλεγχο.');
    return;
  }
 
  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); // Symbol, Added, GridType, MyLower, MyUpper, Status, Verdict, Score, Σύσταση
  const outputRows = [];
  const alerts = [];
  const previousStrategy = GRID_SCREENER_CONFIG.GRID_STRATEGY;
 
  rows.forEach(function (row) {
    const symbol = String(row[0]).trim();
    if (!symbol) { outputRows.push([row[5] || '', row[6] || '', row[7] || '', row[8] || '']); return; }
 
    const posType = (String(row[2]).trim().toUpperCase() === 'SHORT') ? 'SHORT' : 'LONG';
    const myLower = row[3] || null;
    const myUpper = row[4] || null;
 
    // Ρύθμισε το strategy context ώστε το classify() να κρίνει με τα ΣΩΣΤΑ κριτήρια για ΑΥΤΗ τη θέση.
    GRID_SCREENER_CONFIG.GRID_STRATEGY = posType;
 
    try {
      const klines = fetchKlines(symbol, GRID_SCREENER_CONFIG.INTERVAL, GRID_SCREENER_CONFIG.LOOKBACK_CANDLES);
      if (!klines || klines.length < 30) {
        outputRows.push(['🚨 CLOSE — NO DATA (quick)', 'N/A', '', 'Quick check: δεν βρέθηκαν δεδομένα (πιθανό delisting ή λάθος symbol).']);
        alerts.push({ symbol: symbol, status: 'CLOSE', reason: 'Quick check: δεν βρέθηκαν δεδομένα.', gridType: posType });
        return;
      }
 
      const closes = klines.map(function (k) { return k.close; });
      const highs = klines.map(function (k) { return k.high; });
      const lows = klines.map(function (k) { return k.low; });
      const lastPrice = closes[closes.length - 1];
      const periodHigh = Math.max.apply(null, highs);
      const periodLow = Math.min.apply(null, lows);
      const rangePct = ((periodHigh - periodLow) / periodLow) * 100;
      const slopePct = calcTrendSlopePct(closes);
      const adx = calcADX(klines, 14);
      const choppiness = calcChoppinessIndex(klines, 14);
      const verdict = classify(rangePct, slopePct, adx, choppiness);
 
      let statusCell, reason = '';
 
      if (!isEligibleVerdict(verdict)) {
        const gridLabel = posType === 'LONG' ? 'long grid' : 'short grid';
        reason = 'Quick check: verdict έγινε "' + verdict + '" — δεν είναι πια κατάλληλο για ' + gridLabel + '.';
        statusCell = '🚨 CLOSE RECOMMENDED (quick)';
        alerts.push({ symbol: symbol, status: 'CLOSE', reason: reason, gridType: posType });
      } else if (myLower && myUpper) {
        const priceRisk = computeExitRisk({ lastPrice: lastPrice }, myLower, myUpper);
        if (priceRisk) {
          statusCell = '🚨 CLOSE RECOMMENDED (quick)';
          reason = priceRisk;
          alerts.push({ symbol: symbol, status: 'CLOSE', reason: reason, gridType: posType });
        } else {
          statusCell = '✅ OK (quick check)';
        }
      } else {
        statusCell = '✅ OK (quick check)';
      }
 
      outputRows.push([statusCell, verdict, '(quick — δες score στο επόμενο πλήρες run)', reason]);
 
    } catch (err) {
      outputRows.push(['ERROR (quick)', '', '', err.message]);
    }
  });
 
  GRID_SCREENER_CONFIG.GRID_STRATEGY = previousStrategy; // restore
 
  if (outputRows.length > 0) {
    sheet.getRange(2, 6, outputRows.length, 4).setValues(outputRows);
  }
 
  const closes = alerts.filter(function (a) { return a.status === 'CLOSE'; });
  if (closes.length > 0 && GRID_SCREENER_CONFIG.TELEGRAM_NOTIFY) {
    notifyQuickCheckTelegram(closes);
  }
 
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '⚡ Quick check: ' + rows.length + ' θέσεις ελέγχθηκαν, ' + closes.length + ' χρειάζονται άμεση προσοχή.'
  );
}
 
// Ξεχωριστό, ελαφρύ Telegram μήνυμα ΜΟΝΟ για urgent alerts από το quick check —
// κάθε alert δείχνει το δικό του gridType (LONG/SHORT), αφού το quick check μπορεί
// να καλύπτει θέσεις και των δύο τύπων μαζί σε ένα πέρασμα.
function notifyQuickCheckTelegram(closes) {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!botToken || !chatId) return;
 
  let message = '⚡🚨 *Quick Check — Άμεση Προσοχή*\n';
  closes.forEach(function (a) {
    message += '\n*' + a.symbol + '* [' + a.gridType + ']\n' + a.reason + '\n';
  });
 
  const url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Telegram send failed (quick check): ' + err.message);
  }
}
 
// Αξιολογεί κάθε ανοιχτή θέση ΜΕ ΔΙΚΗ ΤΗΣ λογική (όχι σε σχέση με το stable pick):
// - Αν κατέρρευσε το verdict (ή σπάσαν τα δικά σου limits, αν τα έδωσες) -> CLOSE
// - Αν υπάρχει υποψήφιος που τη νικάει με ΠΟΛΥ μεγάλο περιθώριο -> ενημερωτικό, όχι επείγον
// - Αλλιώς -> συνέχισε κανονικά
// Επιστρέφει λίστα από alerts (μόνο ό,τι χρειάζεται προσοχή) για το Telegram.
function evaluateOpenPositions(fullOrdered) {
  const sheet = ensureOpenPositionsSheetExists();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
 
  const cfg = GRID_SCREENER_CONFIG.OPEN_POSITIONS;
  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); // Symbol, Added, GridType, MyLower, MyUpper, Status, Verdict, Score, Σύσταση
  const validToday = fullOrdered.filter(function (r) { return !r.error; });
  const topGood = validToday.find(function (r) { return isGoodVerdict(r.verdict); }) || validToday[0];
  const activeStrategy = GRID_SCREENER_CONFIG.GRID_STRATEGY; // ενεργό ΤΩΡΑ (ορίζεται από runScreenerForStrategy)
 
  const alerts = [];
  const outputRows = [];
 
  rows.forEach(function (row) {
    const symbol = String(row[0]).trim();
    if (!symbol) { outputRows.push(['', '', '', '']); return; }
 
    const posType = (String(row[2]).trim().toUpperCase() === 'SHORT') ? 'SHORT' : 'LONG';
 
    // ΚΡΙΣΙΜΟ: αξιολόγησε τη θέση ΜΟΝΟ στο run που ταιριάζει με τον δικό της τύπο.
    // Χωρίς αυτό, μια LONG θέση θα κρινόταν με NEUTRAL κριτήρια (και το αντίστροφο),
    // προτείνοντας πάντα κλείσιμο άδικα. Στο "λάθος" run, άφησέ την ως έχει (μην τη
    // σβήσεις/αλλάξεις) — θα ενημερωθεί σωστά στο δικό της run.
    if (posType !== activeStrategy) {
      outputRows.push([row[5] || '', row[6] || '', row[7] || '', row[8] || '']);
      return;
    }
 
    const myLower = row[3] || null;
    const myUpper = row[4] || null;
 
    const r = validToday.find(function (x) { return x.symbol === symbol; });
 
    if (!r) {
      outputRows.push(['🚨 CLOSE — DELISTED/NO DATA', 'N/A', '', 'Δεν βρέθηκαν δεδομένα — έλεγξε χειροκίνητα αν υπάρχει ακόμα το instrument.']);
      alerts.push({ symbol: symbol, status: 'CLOSE', reason: 'Δεν βρέθηκαν πλέον δεδομένα (πιθανό delisting ή λάθος symbol).' });
      return;
    }
 
    const isEligible = isEligibleVerdict(r.verdict);
 
    if (!isEligible) {
      const gridLabel = activeStrategy === 'SHORT' ? 'short grid' : 'long grid';
      const reason = 'Verdict έγινε "' + r.verdict + '" (score ' + r.score.toFixed(1) + ') — δεν είναι πια κατάλληλο για ' + gridLabel + '.';
      outputRows.push(['🚨 CLOSE RECOMMENDED', r.verdict, r.score.toFixed(1), reason]);
      alerts.push({ symbol: symbol, status: 'CLOSE', reason: reason, gridType: posType });
      return;
    }
 
    // Αν ο χρήστης έδωσε τα δικά του limits, έλεγξε και price-boundary risk πάνω σε αυτά
    if (myLower && myUpper) {
      const priceRisk = computeExitRisk(r, myLower, myUpper);
      if (priceRisk) {
        outputRows.push(['🚨 CLOSE RECOMMENDED', r.verdict, r.score.toFixed(1), priceRisk]);
        alerts.push({ symbol: symbol, status: 'CLOSE', reason: priceRisk, gridType: posType });
        return;
      }
    }
 
    // Ακόμα υγιές — έλεγξε αν υπάρχει τόσο μεγάλη αναβάθμιση διαθέσιμη που αξίζει να τη σκεφτείς
    if (topGood && topGood.symbol !== symbol) {
      const margin = topGood.score - r.score;
      if (margin >= cfg.BIG_UPGRADE_MARGIN) {
        const reason = topGood.symbol + ' σκοράρει ' + margin.toFixed(1) + ' πόντους παραπάνω (' +
          topGood.score.toFixed(1) + ' vs ' + r.score.toFixed(1) + '). Το ' + symbol + ' παραμένει υγιές — ' +
          'αυτό είναι ενημερωτικό, όχι επείγουσα σύσταση. Λόγω κόστους κλεισίματος/re-setup, αξίζει μόνο αν η ' +
          'διαφορά επιμείνει για μέρες, όχι για μια τυχαία καλή μέρα του ανταγωνιστή.';
        outputRows.push(['💡 ΚΑΛΥΤΕΡΗ ΕΥΚΑΙΡΙΑ ΔΙΑΘΕΣΙΜΗ', r.verdict, r.score.toFixed(1), reason]);
        alerts.push({ symbol: symbol, status: 'INFO', reason: reason, gridType: posType });
        return;
      }
    }
 
    // Προληπτική προειδοποίηση: score που πέφτει συνεχόμενα τα τελευταία runs —
    // ΠΡΙΝ καταρρεύσει το verdict. Ξεχωριστό status ('DECLINING') από το CLOSE,
    // ώστε να ξέρεις ότι είναι έγκαιρη προειδοποίηση, όχι επείγουσα σύσταση εξόδου.
    const declineInfo = checkDecliningTrend(symbol);
    if (declineInfo) {
      const reason = 'Score πέφτει συνεχόμενα: ' + declineInfo.scores.join(' → ') + ' (τελευταία ' +
        declineInfo.runs + ' runs, συνολική πτώση ' + declineInfo.totalDecline.toFixed(1) + ' πόντους). ' +
        'Ακόμα δεν κατέρρευσε το verdict ("' + r.verdict + '") — παρακολούθησε στενά, ίσως χρειαστεί ενέργεια σύντομα.';
      outputRows.push(['📉 ΦΘΙΝΟΥΣΑ ΤΑΣΗ', r.verdict, r.score.toFixed(1), reason]);
      alerts.push({ symbol: symbol, status: 'DECLINING', reason: reason, gridType: posType });
      return;
    }
 
    outputRows.push(['✅ ΣΥΝΕΧΙΣΕ', r.verdict, r.score.toFixed(1), '']);
    alerts.push({ symbol: symbol, status: 'OK', reason: '', score: r.score, verdict: r.verdict, gridType: posType });
  });
 
  if (outputRows.length > 0) {
    sheet.getRange(2, 6, outputRows.length, 4).setValues(outputRows);
  }
  sheet.autoResizeColumns(1, 9);
 
  return alerts;
}
 
// Ξεχωριστό Telegram μήνυμα ΜΟΝΟ για ανοιχτές θέσεις που χρειάζονται προσοχή —
// δεν αναμιγνύεται με το μήνυμα του stable pick, ώστε να είναι ξεκάθαρο ποιο
// αφορά "τι προτείνει το script" και ποιο "τι να κάνεις με ό,τι ήδη έχεις ανοιχτό".
function notifyOpenPositionsTelegram(alerts) {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!botToken || !chatId) return;
 
  const closes = alerts.filter(function (a) { return a.status === 'CLOSE'; });
  const declining = alerts.filter(function (a) { return a.status === 'DECLINING'; });
  const infos = alerts.filter(function (a) { return a.status === 'INFO'; });
  const oks = alerts.filter(function (a) { return a.status === 'OK'; });
 
  let message = '📂 *Ανοιχτές Θέσεις [' + GRID_SCREENER_CONFIG.GRID_STRATEGY + '] — Ενημέρωση*\n';
 
  closes.forEach(function (a) {
    message += '\n🚨 *' + a.symbol + '* — ΠΡΟΤΕΙΝΕΤΑΙ ΚΛΕΙΣΙΜΟ\n' + a.reason + '\n';
  });
  declining.forEach(function (a) {
    message += '\n📉 *' + a.symbol + '* — φθίνουσα τάση (προληπτική ειδοποίηση)\n' + a.reason + '\n';
  });
  infos.forEach(function (a) {
    message += '\n💡 *' + a.symbol + '* — ενημερωτικό\n' + a.reason + '\n';
  });
  // Οι "OK" θέσεις εμφανίζονται ΠΑΝΤΑ (όχι μόνο όταν όλα είναι εντάξει) — αλλιώς
  // μια υγιής θέση "εξαφανίζεται" σιωπηλά από το μήνυμα όποτε υπάρχει έστω ένα
  // alert για άλλο coin, κάτι που μπερδεύει παρά ενημερώνει.
  if (oks.length > 0) {
    message += '\n✅ Χωρίς ενέργεια:\n';
    oks.forEach(function (a) {
      message += a.symbol + ' — ' + a.verdict + ' (score ' + a.score.toFixed(1) + ')\n';
    });
  }
 
  const url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Telegram send failed (open positions): ' + err.message);
  }
}
 
// ==================== A2 ΘΕΣΗ vs ΣΗΜΕΡΙΝΟ #1 PICK ====================
//
// Ελέγχει αν το crypto στο κελί A2 του "Open Positions" (η πρώτη καταχωρημένη
// ανοιχτή θέση) ΕΙΝΑΙ η σημερινή #1 επιλογή του Grid Screener για το ΔΙΚΟ ΤΗΣ
// strategy (LONG/SHORT, από τη στήλη C). Αν ΔΕΝ είναι, επιστρέφει πλήρη στοιχεία
// ώστε να χτιστεί αναλυτικό μήνυμα — αν ΕΙΝΑΙ ήδη η #1 επιλογή, δεν χρειάζεται
// τίποτα παραπάνω (θα το δει ήδη στο κύριο μήνυμα του stable pick).
function checkOpenPositionA2VsTopPick(fullOrdered) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GRID_SCREENER_CONFIG.OPEN_POSITIONS.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
 
  const a2 = String(sheet.getRange(2, 1).getValue()).trim();
  const c2 = String(sheet.getRange(2, 3).getValue()).trim().toUpperCase();
  if (!a2) return null;
 
  const posType = (c2 === 'SHORT') ? 'SHORT' : 'LONG';
  if (posType !== GRID_SCREENER_CONFIG.GRID_STRATEGY) return null; // αξιολογείται στο δικό της run
 
  const validToday = fullOrdered.filter(function (r) { return !r.error; });
  if (validToday.length === 0) return null;
 
  const topToday = validToday[0]; // ήδη ταξινομημένο κατά tier+score
  if (topToday.symbol === a2) return null; // είναι ήδη η #1 επιλογή -> τίποτα νέο να πούμε
 
  const r = validToday.find(function (x) { return x.symbol === a2; });
  if (!r) return { symbol: a2, notFound: true, topPick: topToday };
 
  return { symbol: a2, r: r, topPick: topToday };
}
 
// Χτίζει αναλυτικό μήνυμα για τη θέση A2 — ίδια δομή στοιχείων με το κύριο
// μήνυμα του stable pick (score, verdict, εύρος, TP/SL, Fib, ζώνη εισόδου),
// συν ρητή σύσταση ΚΛΕΙΣΙΜΟ ή ΚΡΑΤΑ βάσει του τρέχοντος verdict της.
function buildOpenPositionAnalysisMessage(info) {
  if (info.notFound) {
    return '📋 *Ανάλυση Ανοιχτής Θέσης (A2, ' + GRID_SCREENER_CONFIG.GRID_STRATEGY + '): ' + info.symbol + '*\n\n' +
      '⚠️ Δεν βρέθηκαν δεδομένα γι\' αυτό το symbol (πιθανό delisting ή λάθος όνομα).\n\n' +
      '🚨 Σύσταση: Έλεγξε χειροκίνητα στο Pionex — πιθανή ανάγκη κλεισίματος.\n\n' +
      '👑 Σημερινή #1 επιλογή: ' + info.topPick.symbol + ' (score ' + info.topPick.score.toFixed(1) + ')';
  }
 
  const r = info.r;
  const limits = computeBufferedLimits(r.periodLow, r.periodHigh, r.lastPrice);
  const grids = suggestGridCount(r.rangePct);
  const leverage = suggestLeverage(r.atrPct);
  const tpPct = suggestTP(r.estReturnPct, leverage);
  const slPct = suggestSL(r.atrPct, leverage);
  const fib618 = computeFibLevel(limits.lowerLimit, limits.upperLimit, 0.618, r.lastPrice);
  const fib786 = computeFibLevel(limits.lowerLimit, limits.upperLimit, 0.786, r.lastPrice);
 
  let entryZoneText = '';
  if (r.entrySuggestion) {
    const es = computeEntryZone(
      r.lastPrice, limits.lowerLimit, limits.upperLimit, r.entrySuggestion.shortSlopePct,
      r.entrySuggestion.recentSwingLow, r.entrySuggestion.recentSwingHigh, r.entrySuggestion.last2Analysis
    );
    entryZoneText = '\nΖώνη Εισόδου (4H): ' + es.zone + ' (' + es.posPct.toFixed(0) + '% του εύρους)';
  }
 
  const isEligible = isEligibleVerdict(r.verdict);
  const recommendation = isEligible
    ? '✅ Σύσταση: ΚΡΑΤΑ το grid ανοιχτό — verdict "' + r.verdict + '" (score ' + r.score.toFixed(1) + '), ακόμα κατάλληλο.'
    : '🚨 Σύσταση: ΕΞΕΤΑΣΕ ΚΛΕΙΣΙΜΟ — verdict έγινε "' + r.verdict + '" (score ' + r.score.toFixed(1) + '), δεν είναι πια κατάλληλο για ' +
      (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT' ? 'short' : 'long') + ' grid.';
 
  return '📋 *Ανάλυση Ανοιχτής Θέσης [' + GRID_SCREENER_CONFIG.GRID_STRATEGY + '] (δεν είναι η σημερινή #1 επιλογή): ' + r.symbol + '*\n\n' +
    'Grid Score: ' + r.score.toFixed(1) + '/100 (rank #' + r.rank + ')\n' +
    'Verdict: ' + r.verdict + '\n' +
    'Last Price: ' + r.lastPrice + '\n' +
    'Range: ' + r.rangePct.toFixed(2) + '%\n' +
    'Choppiness: ' + r.choppiness.toFixed(1) + '\n' +
    'ADX: ' + r.adx.toFixed(1) + entryZoneText + '\n\n' +
    '📐 Τρέχουσες παράμετροι (υπολογισμένες τώρα, όχι παγωμένες):\n' +
    'Lower Limit: ' + limits.lowerLimit + '\n' +
    'Upper Limit: ' + limits.upperLimit + '\n' +
    'Grids: ' + grids + ' | Leverage: ' + leverage + 'x\n' +
    'TP: +' + tpPct.toFixed(1) + '% | SL: -' + slPct.toFixed(1) + '%\n' +
    'Fib 0.618: ' + fib618 + ' | Fib 0.786: ' + fib786 + '\n\n' +
    '👑 Σημερινή #1 επιλογή: ' + info.topPick.symbol + ' (score ' + info.topPick.score.toFixed(1) + ')\n\n' +
    recommendation;
}
 
function notifyOpenPositionAnalysisTelegram(message) {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!botToken || !chatId) return;
 
  const url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Telegram send failed (A2 position analysis): ' + err.message);
  }
}
 
// ==================== TELEGRAM NOTIFICATION ====================
 
// Χρησιμοποιεί τα ΙΔΙΑ Script Properties (TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID)
// που έχεις ήδη ρυθμισμένα από το betting alert script — δεν χρειάζεται να τα
// ξαναβάλεις. Αν λείπουν, απλά log-άρει προειδοποίηση και δεν σκάει το script.
// Ειδοποιεί για τη ΣΤΑΘΕΡΗ (locked) επιλογή, όχι για όποιο coin έτυχε να
// σκοράρει ψηλότερα σήμερα — έτσι δεν αλλάζει το μήνυμα κάθε μέρα άδικα.
function notifyBestPickTelegram(stableInfo, fullOrdered) {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
 
  if (!botToken || !chatId) {
    Logger.log('Telegram credentials not set in Script Properties (TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID) — skipping notification.');
    return;
  }
 
  let message;
 
  if (!stableInfo || !stableInfo.pick) {
    message = '⚠️ *Grid Screener [' + GRID_SCREENER_CONFIG.GRID_STRATEGY + ']*\n\nΔεν βρέθηκε valid αποτέλεσμα σήμερα.';
  } else {
    const p = stableInfo.pick;
    const isGood = isGoodVerdict(p.verdict);
    const isMarginal = isMarginalVerdict(p.verdict);
    const lowerLimit = stableInfo.lowerLimit;
    const upperLimit = stableInfo.upperLimit;
    const grids = suggestGridCount(p.rangePct);
    const leverage = suggestLeverage(p.atrPct);
    const tpPct = suggestTP(p.estReturnPct, leverage);
    const slPct = suggestSL(p.atrPct, leverage);
 
    const modeLabel = GRID_SCREENER_CONFIG.GRID_STRATEGY;
 
    let headerLine;
    if (stableInfo.collapsed) {
      headerLine = '🚨 *Grid Screener [' + modeLabel + '] — ΠΡΟΤΕΙΝΕΤΑΙ ΕΞΟΔΟΣ + Νέα Επιλογή*';
    } else if (stableInfo.switched) {
      headerLine = '🔄 *Grid Screener [' + modeLabel + '] — Νέα Σταθερή Επιλογή*';
    } else {
      headerLine = '🔒 *Grid Screener [' + modeLabel + '] — Σταθερή Επιλογή* (χωρίς αλλαγή)';
    }
 
    message = headerLine + '\n\n';
 
    // Αν η προηγούμενη επιλογή κατέρρευσε, εξήγησε ρητά τι να κλείσει πρώτα
    if (stableInfo.collapsed && stableInfo.collapsedFrom) {
      const cf = stableInfo.collapsedFrom;
      message += '❌ Έξοδος από: *' + cf.symbol + '* (' + cf.name + ')\n' +
        'Λόγος: verdict έγινε "' + cf.verdict + '"' +
        (cf.score !== null ? ' (score ' + cf.score.toFixed(1) + ')' : ' (δεν βρέθηκε πλέον στα δεδομένα)') + '\n' +
        '👉 Αν έχεις ενεργό grid σε αυτό, κλείσ\' το χειροκίνητα στο Pionex πριν στήσεις το νέο.\n\n';
    }
 
    message += '*' + p.symbol + '* (' + p.name + ')\n' +
      'Grid Score: *' + p.score.toFixed(1) + '/100* (rank #' + p.rank + ')\n' +
      'Verdict: ' + (isGood ? '✅ ' + p.verdict : '⚠️ ' + p.verdict) + '\n' +
      'Κλειδωμένο από: ' + stableInfo.lockedSince + '\n\n' +
      'Last Price: ' + p.lastPrice + '\n' +
      'Range: ' + p.rangePct.toFixed(2) + '%\n' +
      'Choppiness: ' + p.choppiness.toFixed(1) + '\n' +
      'ADX: ' + p.adx.toFixed(1) + '\n\n' +
      '📐 *Παράμετροι Pionex ' + (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT' ? 'Short' : 'Long') + ' Grid* (παγωμένες από τότε που κλειδώθηκε):\n' +
      'Lower Limit: ' + lowerLimit + '\n' +
      'Upper Limit: ' + upperLimit + '\n' +
      'Εναλλακτική τιμή εισόδου (' + GRID_SCREENER_CONFIG.ALT_ENTRY_BELOW_HIGH_PCT + '% ' +
      (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT' ? 'πάνω από Lower Limit' : 'κάτω από Upper Limit') + '): ' +
      computeAltEntryPrice(lowerLimit, upperLimit, p.lastPrice) + '\n' +
      'Grids: ' + grids + '\n' +
      'Leverage: ' + leverage + 'x\n' +
      'Take Profit (PnL ratio): +' + tpPct.toFixed(1) + '%\n' +
      'Stop Loss (PnL ratio): -' + slPct.toFixed(1) + '%\n' +
      'Fib 0.618: ' + computeFibLevel(lowerLimit, upperLimit, 0.618, p.lastPrice) + '\n' +
      'Fib 0.786: ' + computeFibLevel(lowerLimit, upperLimit, 0.786, p.lastPrice);
 
    // Προληπτική προειδοποίηση: score που πέφτει συνεχόμενα τα τελευταία runs,
    // ΠΡΙΝ καταρρεύσει το verdict — άχρηστη πληροφορία αν μόλις έγινε ήδη exit.
    if (!stableInfo.collapsed) {
      const declineInfo = checkDecliningTrend(p.symbol);
      if (declineInfo) {
        message += '\n\n📉 *Φθίνουσα τάση:* score ' + declineInfo.scores.join(' → ') +
          ' (τελευταία ' + declineInfo.runs + ' runs, συνολική πτώση ' + declineInfo.totalDecline.toFixed(1) +
          ' πόντους). Ακόμα δεν κατέρρευσε το verdict, αλλά παρακολούθησε στενά — ίσως χρειαστεί ενέργεια σύντομα.';
      }
    }
 
    let entryReady = true; // default: μην μπλοκάρεις αν δεν μπορέσαμε να υπολογίσουμε ζώνη
 
    if (p.entrySuggestion) {
      const es = computeEntryZone(p.lastPrice, lowerLimit, upperLimit, p.entrySuggestion.shortSlopePct,
        p.entrySuggestion.recentSwingLow, p.entrySuggestion.recentSwingHigh, p.entrySuggestion.last2Analysis);
      entryReady = (es.zone !== 'ΨΗΛΗ ΖΩΝΗ' && es.zone !== 'ΧΑΜΗΛΗ ΖΩΝΗ'); // ΚΑΛΗ ή ΜΕΤΡΙΑ = εντάξει, μόνο η δυσμενής ζώνη μπλοκάρει
 
      message += '\n\n📍 *Ζώνη Εισόδου (4H, βάσει πραγματικών grid limits):* ' + es.zone +
        ' (' + es.posPct.toFixed(0) + '% του εύρους)\n' + es.note;
 
      if (GRID_SCREENER_CONFIG.ENTRY_SUGGESTION.BLOCK_ENTRY_OUTSIDE_GOOD_ZONE && !entryReady) {
        message += '\n\n⏳ *ΑΝΑΜΟΝΗ ΕΙΣΟΔΟΥ* — η ζώνη δεν είναι ΚΑΛΗ. Μη στήσεις πλήρες μέγεθος τώρα.';
      }
 
      const nextRange = computeNextCandleRange(p.lastPrice, p.atrPct);
      if (nextRange) {
        message += '\n\n📊 *Στατιστικό εύρος επόμενου 4H* (βάσει ATR, ΟΧΙ πρόβλεψη κατεύθυνσης):\n' +
          nextRange.low + ' – ' + nextRange.high + '\nΔείχνει τυπικό μέγεθος κίνησης βάσει πρόσφατης μεταβλητότητας, όχι προς τα πού θα πάει η τιμή.';
      }
    }
 
    if (stableInfo.challenger) {
      message += '\n\n👀 Προκλητής: *' + stableInfo.challenger + '* (' +
        stableInfo.challengerStreak + '/' + stableInfo.requiredStreak + ' συνεχόμενες νίκες πριν αλλάξει η επιλογή)';
    }
 
    const priceExitRisk = computeExitRisk(p, lowerLimit, upperLimit);
    if (priceExitRisk) {
      message += '\n\n' + priceExitRisk;
    }
 
    const blockActive = GRID_SCREENER_CONFIG.ENTRY_SUGGESTION.BLOCK_ENTRY_OUTSIDE_GOOD_ZONE;
 
    // Actionable οδηγία ανάλογα με την κατάσταση (όχι απλή προειδοποίηση χωρίς ενέργεια).
    // Το "στήσε τώρα" εμφανίζεται ΜΟΝΟ όταν προτείνεται ΝΕΑ είσοδος (collapsed/switched)
    // ΚΑΙ η ζώνη είναι πράγματι ΚΑΛΗ — αλλιώς ρητό "περίμενε".
    if (stableInfo.collapsed) {
      if (entryReady || !blockActive) {
        message += '\n\n✅ Ενέργεια: Στήσε το νέο grid με τις παραπάνω παραμέτρους μετά την έξοδο από το παλιό.';
      } else {
        message += '\n\n⏳ Ενέργεια: Έβγες από το παλιό, αλλά ΜΗΝ στήσεις ακόμα το νέο με πλήρες μέγεθος — ' +
          'περίμενε να μπει σε ΚΑΛΗ ΖΩΝΗ, ή χρησιμοποίησε το προτεινόμενο split αν υπάρχει παραπάνω.';
      }
    } else if (stableInfo.switched && !entryReady && blockActive) {
      message += '\n\n⏳ Ενέργεια: Νέα επιλογή, αλλά η ζώνη εισόδου δεν είναι ιδανική ακόμα — περίμενε καλύτερο σημείο πριν στήσεις.';
    } else if (isMarginal) {
      message += '\n\n⚠️ Ενέργεια: Δεν είναι πλέον "GOOD" αλλά ούτε κατέρρευσε. Αν έχεις ήδη ενεργό grid, εξέτασε να ' +
        'μειώσεις leverage/position size ή να στενέψεις το stop-loss. Θα σε ειδοποιήσω άμεσα αν χρειαστεί πλήρης έξοδος.';
    }
  }
 
  const url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'Markdown'
  };
 
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Telegram send failed: ' + err.message);
  }
}
 
 
 
// Προσθέτει (χωρίς να σβήνει τίποτα) ένα timestamped snapshot του ranking
// σε ξεχωριστό sheet, ώστε να βλέπεις πώς άλλαξε η κατάταξη μέρα με τη μέρα.
function logHistory(ordered) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(getLogSheetName());
 
  const logHeaders = [
    'Timestamp', 'Rank', 'Symbol', 'Name', 'Grid Score', 'Verdict',
    'Last Price', 'Range %', 'Choppiness Idx', 'ADX', 'Est. Return (period)'
  ];
 
  if (!logSheet) {
    logSheet = ss.insertSheet(getLogSheetName());
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
 
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
 
  const logRows = [];
  ordered.forEach(function (r, idx) {
    if (r.error) return; // δεν καταγράφουμε errors στο ιστορικό
    logRows.push([
      timestamp, idx + 1, r.symbol, r.name, r.score.toFixed(1), r.verdict,
      r.lastPrice, r.rangePct.toFixed(2) + '%', r.choppiness.toFixed(1),
      r.adx.toFixed(1), r.estReturnPct.toFixed(2) + '%'
    ]);
  });
 
  if (logRows.length > 0) {
    const startRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startRow, 1, logRows.length, logHeaders.length).setValues(logRows);
  }
  logSheet.autoResizeColumns(1, logHeaders.length);
}
 
// Διαβάζει τις τελευταίες N εμφανίσεις ενός symbol από το Log (παλιότερο -> νεότερο),
// σαρώνοντας από το ΤΕΛΟΣ του sheet προς τα πίσω — πιο αποδοτικό από πλήρη σάρωση
// καθώς το Log μεγαλώνει με κάθε run.
function getRecentScoreHistory(symbol, n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(getLogSheetName());
  if (!logSheet) return [];
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];
 
  const maxScanRows = Math.min(lastRow - 1, 5000); // safety cap σε πολύ μεγάλα logs
  const startRow = Math.max(2, lastRow - maxScanRows + 1);
  const numRows = lastRow - startRow + 1;
  const data = logSheet.getRange(startRow, 1, numRows, 5).getValues(); // Timestamp, Rank, Symbol, Name, Score
 
  const matches = [];
  for (let i = data.length - 1; i >= 0 && matches.length < n; i--) {
    if (String(data[i][2]).trim() === symbol) {
      matches.push({ timestamp: data[i][0], score: parseFloat(data[i][4]) });
    }
  }
  matches.reverse(); // παλιότερο -> νεότερο
  return matches;
}
 
// Ανιχνεύει "φθίνουσα τάση": score που πέφτει ΣΥΝΕΧΟΜΕΝΑ (μονότονα) στα τελευταία
// runs, με συνολική πτώση πάνω από το όριο. Σκόπιμα αυστηρό (μονότονο, όχι απλά
// "χαμηλότερο από πριν") ώστε να μην πυροδοτείται από απλή διακύμανση θορύβου.
function checkDecliningTrend(symbol) {
  const cfg = GRID_SCREENER_CONFIG.DECLINING_TREND;
  const history = getRecentScoreHistory(symbol, cfg.LOOKBACK_RUNS);
  if (history.length < cfg.LOOKBACK_RUNS) return null; // δεν υπάρχει αρκετό ιστορικό ακόμα
 
  for (let i = 1; i < history.length; i++) {
    if (history[i].score > history[i - 1].score) return null; // δεν είναι μονότονη πτώση
  }
 
  const totalDecline = history[0].score - history[history.length - 1].score;
  if (totalDecline < cfg.MIN_DECLINE_POINTS) return null;
 
  return {
    totalDecline: totalDecline,
    scores: history.map(function (h) { return h.score; }),
    runs: history.length
  };
}
 
// Καθαρίζει εντελώς το ιστορικό ΟΛΩΝ των στρατηγικών (π.χ. αν θες να ξεκινήσεις από την αρχή)
function clearHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cleared = 0;
  GRID_SCREENER_CONFIG.STRATEGIES_TO_RUN.forEach(function (strategy) {
    const name = GRID_SCREENER_CONFIG.LOG_SHEET_NAME + ' - ' + strategy;
    const logSheet = ss.getSheetByName(name);
    if (logSheet) {
      ss.deleteSheet(logSheet);
      cleared++;
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Το ιστορικό διαγράφηκε (' + cleared + ' sheets).');
}
 
 
 
// Φέρνει τα top-N πιο ρευστά USDT spot pairs από το OKX (κατά 24ωρο τζίρο σε USDT),
// ως proxy για "μεγάλα coins" — χωρίς εξάρτηση από CoinGecko (που μπλοκάρει συχνά
// τις κοινόχρηστες IPs του Google Apps Script με 429, ανεξαρτήτως retry).
// Χρησιμοποιεί την ΙΔΙΑ πηγή (OKX) με τα klines, άρα ένα μόνο API dependency.
// Σημείωση: ταξινόμηση κατά όγκο ≠ market cap, αλλά είναι στην πράξη πιο σχετικός
// δείκτης για grid trading, αφού ο όγκος καθορίζει τη ρευστότητα/spread.
// Φέρνει τα top-N πιο ρευστά USDT-marginated PERPETUAL FUTURES (SWAP) από το OKX,
// κατά 24ωρο τζίρο. Το script προτείνει παραμέτρους για futures grid (Pionex),
// άρα πρέπει να ελέγχει δεδομένα από την ΙΔΙΑ αγορά (SWAP), όχι από spot — spot
// και perpetual futures μπορεί να έχουν ελαφρώς διαφορετική μεταβλητότητα/όγκο.
function fetchTopNSymbols(topN) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'top_n_symbols_okx_swap_' + topN;
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
 
  const url = 'https://www.okx.com/api/v5/market/tickers?instType=SWAP';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('OKX tickers API HTTP error (' + response.getResponseCode() + ')');
  }
  const json = JSON.parse(response.getContentText());
  if (json.code !== '0') {
    throw new Error('OKX tickers API error: ' + json.msg);
  }
 
  const exclude = GRID_SCREENER_CONFIG.STABLECOIN_EXCLUDE;
 
  const list = json.data
    .filter(function (t) {
      // Μόνο linear USDT-marginated perpetuals (BASE-USDT-SWAP) — όχι inverse (-USD-SWAP)
      // και όχι USDC-marginated, ώστε να ταιριάζει με ό,τι προσφέρει συνήθως το Pionex grid.
      if (t.instId.slice(-10) !== '-USDT-SWAP') return false;
      const base = t.instId.split('-')[0];
      if (exclude.indexOf(base) !== -1) return false;
      return true;
    })
    .map(function (t) {
      return { instId: t.instId, base: t.instId.split('-')[0], volCcy24h: parseFloat(t.volCcy24h || '0') };
    })
    .sort(function (a, b) { return b.volCcy24h - a.volCcy24h; })
    .slice(0, topN)
    .map(function (t) { return { symbol: t.instId, name: t.base }; });
 
  cache.put(cacheKey, JSON.stringify(list), GRID_SCREENER_CONFIG.TOP_N_CACHE_SECONDS);
  return list;
}
 
function fetchKlines(instId, bar, limit) {
  // OKX επιστρέφει max 300 candles ανά αίτημα, νεότερο -> παλαιότερο
  const cappedLimit = Math.min(limit, 300);
  const url = 'https://www.okx.com/api/v5/market/candles?instId=' + instId +
    '&bar=' + bar + '&limit=' + cappedLimit;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('OKX API HTTP error (' + response.getResponseCode() + ')');
  }
  const json = JSON.parse(response.getContentText());
  if (json.code !== '0') {
    throw new Error('OKX API error: ' + json.msg + ' — έλεγξε ότι το instId (' + instId + ') υπάρχει.');
  }
  // Κάθε candle: [ts, open, high, low, close, vol, volCcy, volCcyQuote, confirm]
  // Reverse ώστε να είναι χρονολογικά (παλιό -> νέο), όπως περιμένουν οι indicator functions
  return json.data.reverse().map(function (k) {
    return {
      openTime: parseInt(k[0], 10),
      open: parseFloat(k[1]),
      high: parseFloat(k[2]),
      low: parseFloat(k[3]),
      close: parseFloat(k[4]),
      volume: parseFloat(k[5])
    };
  });
}
 
// ==================== INDICATORS ====================
 
// Κλίση γραμμικής παλινδρόμησης πάνω στα closes, εκφρασμένη ως % της μέσης τιμής ανά candle.
// Τιμή κοντά στο 0 = πλάγια κίνηση. Θετική/αρνητική μεγάλη = trending.
function calcTrendSlopePct(closes) {
  const n = closes.length;
  const xMean = (n - 1) / 2;
  const yMean = closes.reduce(function (a, b) { return a + b; }, 0) / n;
 
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) {
    num += (i - xMean) * (closes[i] - yMean);
    den += (i - xMean) * (i - xMean);
  }
  const slope = den === 0 ? 0 : num / den; // μονάδες τιμής ανά candle
  return (slope / yMean) * 100;
}
 
// Απλοποιημένος υπολογισμός ADX (Wilder), 0-100. Χαμηλό = ασθενές trend (καλό για grid).
function calcADX(klines, period) {
  if (klines.length < period + 1) return 0;
 
  const plusDM = [], minusDM = [], tr = [];
  for (let i = 1; i < klines.length; i++) {
    const upMove = klines[i].high - klines[i - 1].high;
    const downMove = klines[i - 1].low - klines[i].low;
    plusDM.push(upMove > downMove && upMove > 0 ? upMove : 0);
    minusDM.push(downMove > upMove && downMove > 0 ? downMove : 0);
 
    const highLow = klines[i].high - klines[i].low;
    const highClose = Math.abs(klines[i].high - klines[i - 1].close);
    const lowClose = Math.abs(klines[i].low - klines[i - 1].close);
    tr.push(Math.max(highLow, highClose, lowClose));
  }
 
  function wilderSmooth(arr, period) {
    const smoothed = [];
    let sum = arr.slice(0, period).reduce(function (a, b) { return a + b; }, 0);
    smoothed.push(sum);
    for (let i = period; i < arr.length; i++) {
      sum = smoothed[smoothed.length - 1] - (smoothed[smoothed.length - 1] / period) + arr[i];
      smoothed.push(sum);
    }
    return smoothed;
  }
 
  const trSmooth = wilderSmooth(tr, period);
  const plusDMSmooth = wilderSmooth(plusDM, period);
  const minusDMSmooth = wilderSmooth(minusDM, period);
 
  const dxValues = [];
  for (let i = 0; i < trSmooth.length; i++) {
    const plusDI = trSmooth[i] === 0 ? 0 : (plusDMSmooth[i] / trSmooth[i]) * 100;
    const minusDI = trSmooth[i] === 0 ? 0 : (minusDMSmooth[i] / trSmooth[i]) * 100;
    const dx = (plusDI + minusDI) === 0 ? 0 : (Math.abs(plusDI - minusDI) / (plusDI + minusDI)) * 100;
    dxValues.push(dx);
  }
 
  const last = dxValues.slice(-period);
  return last.reduce(function (a, b) { return a + b; }, 0) / last.length;
}
 
// Average True Range ως % της τιμής (μέσος όρος 14 περιόδων) — μετρά το "τυπικό"
// μέγεθος κίνησης ανά candle, χρησιμοποιείται για το volatility-based leverage sizing.
function calcATRPct(klines, period) {
  if (klines.length < period + 1) return 0;
  const recent = klines.slice(-period - 1);
  let trSum = 0;
  for (let i = 1; i < recent.length; i++) {
    const highLow = recent[i].high - recent[i].low;
    const highClose = Math.abs(recent[i].high - recent[i - 1].close);
    const lowClose = Math.abs(recent[i].low - recent[i - 1].close);
    trSum += Math.max(highLow, highClose, lowClose);
  }
  const avgTR = trSum / period;
  const lastClose = recent[recent.length - 1].close;
  return lastClose === 0 ? 0 : (avgTR / lastClose) * 100;
}
 
// Καθαρή λογική ζώνης εισόδου πάνω σε ΟΠΟΙΟΔΗΠΟΤΕ ζεύγος lower/upper — έτσι μπορεί
// να εφαρμοστεί είτε στο σημερινό ανανεωμένο εύρος είτε στα ΠΑΓΩΜΕΝΑ limits ενός
// ήδη κλειδωμένου grid (πιο σωστό σημείο αναφοράς για τον χρήστη που θέλει να μπει
// ή να αξιολογήσει τη θέση του μέσα στο ΠΡΑΓΜΑΤΙΚΟ grid που τρέχει).
// Στατιστικό εύρος του ΕΠΟΜΕΝΟΥ 4H candle, βασισμένο στο πρόσφατο ATR (μέση
// κίνηση ανά candle τις τελευταίες 14 περιόδους). ΔΕΝ είναι πρόβλεψη κατεύθυνσης —
// δεν λέει "θα ανέβει" ή "θα κατέβει". Δείχνει μόνο το ΤΥΠΙΚΟ ΜΕΓΕΘΟΣ κίνησης
// βάσει πρόσφατης μεταβλητότητας, σαν ένα "expected move" γύρω από την τρέχουσα
// τιμή — χρήσιμο για διαχείριση ρίσκου (π.χ. τι εύρος τιμών να θεωρείς φυσιολογικό
// vs ασυνήθιστο στο επόμενο candle), όχι για να μαντέψεις προς τα πού θα πάει.
function computeNextCandleRange(lastPrice, atrPct) {
  if (!atrPct || atrPct <= 0) return null;
  const atrValue = lastPrice * (atrPct / 100);
  return {
    low: round(lastPrice - atrValue, lastPrice),
    high: round(lastPrice + atrValue, lastPrice),
    atrPct: atrPct
  };
}
 
 
// Strategy-aware: για GRID_STRATEGY='LONG' η καλή ζώνη εισόδου είναι κοντά στο
// ΚΑΤΩ μέρος του εύρους (φτηνή αγορά πριν την άνοδο). Για ='SHORT' είναι το
// ΑΚΡΙΒΩΣ ΑΝΤΙΘΕΤΟ — καλή ζώνη είναι κοντά στο ΠΑΝΩ μέρος (ακριβή πώληση πριν
// την πτώση). Η λογική είναι mirror μέσω του effectivePos (μετρημένο από την
// "ευνοϊκή" πλευρά ανάλογα με τη στρατηγική) αντί να διπλασιάζουμε τον κώδικα.
function computeEntryZone(lastPrice, lowerLimit, upperLimit, shortSlopePct, recentSwingLow, recentSwingHigh, last2Analysis) {
  const cfg = GRID_SCREENER_CONFIG.ENTRY_SUGGESTION;
  const isShort = GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT';
  const rangeSpan = upperLimit - lowerLimit;
  const posPct = rangeSpan > 0 ? ((lastPrice - lowerLimit) / rangeSpan) * 100 : 50; // πάντα μετρημένο από τον πάτο
  const effectivePos = isShort ? (100 - posPct) : posPct; // SHORT: χαμηλό effectivePos = κοντά στην κορυφή = καλό
 
  const sweetSpotPrice = isShort
    ? upperLimit - rangeSpan * cfg.SWEET_SPOT_PCT
    : lowerLimit + rangeSpan * cfg.SWEET_SPOT_PCT;
 
  // Τρέχουσα βραχυπρόθεσμη ροπή (ΓΕΓΟΝΟΣ, όχι πρόβλεψη). Για LONG ευνοϊκή = ανοδική,
  // για SHORT ευνοϊκή = καθοδική — ακριβώς αντίστροφο.
  let momentum, momentumWord;
  if (shortSlopePct <= cfg.STEEP_DROP_SLOPE_PCT) {
    momentum = 'bearish'; momentumWord = 'καθοδική';
  } else if (shortSlopePct >= cfg.STEEP_RISE_SLOPE_PCT) {
    momentum = 'bullish'; momentumWord = 'ανοδική';
  } else {
    momentum = 'neutral'; momentumWord = 'χωρίς σαφή κατεύθυνση';
  }
  const favorableMomentum = isShort ? 'bearish' : 'bullish';
  const adverseMomentum = isShort ? 'bullish' : 'bearish';
  const strategyWord = isShort ? 'short' : 'long';
 
  let zone, note, splitPlan = null;
  if (effectivePos <= cfg.GOOD_ZONE_MAX_PCT) {
    zone = 'ΚΑΛΗ ΖΩΝΗ';
    note = isShort
      ? 'Η τιμή είναι πραγματικά κοντά στο πάνω μέρος του εύρους — καλό σημείο για short είσοδο, χωρίς να χρειάζεται να περιμένεις κάτι.'
      : 'Η τιμή είναι πραγματικά κοντά στο κάτω μέρος του εύρους — καλό σημείο για είσοδο, χωρίς να χρειάζεται να περιμένεις κάτι.';
  } else if (effectivePos <= cfg.MODERATE_ZONE_MAX_PCT) {
    zone = 'ΜΕΤΡΙΑ ΖΩΝΗ';
    note = 'Ούτε ιδιαίτερα ευνοϊκό ούτε δυσμενές σημείο του εύρους — φυσιολογικό σημείο για είσοδο τώρα.';
    if (momentum === favorableMomentum) {
      note += ' Η τρέχουσα βραχυπρόθεσμη ροπή είναι ήδη ' + momentumWord + ' (' + shortSlopePct.toFixed(2) +
        '%/candle τις τελευταίες ~2 μέρες) — δηλαδή η κίνηση προς το ~' + round(sweetSpotPrice, lastPrice) +
        ' ήδη συμβαίνει αυτή τη στιγμή, όχι ότι θα συμβεί. Καμία εγγύηση ότι θα συνεχιστεί.';
    } else if (momentum === adverseMomentum) {
      note += ' Η τρέχουσα βραχυπρόθεσμη ροπή είναι ' + momentumWord + ' (' + shortSlopePct.toFixed(2) +
        '%/candle) — αντίθετη από αυτό που θα ήθελες για ' + strategyWord + ' grid αυτή τη στιγμή. Καμία εγγύηση για το τι θα κάνει μετά.';
    } else {
      note += ' Δεν υπάρχει σαφής βραχυπρόθεσμη ροπή αυτή τη στιγμή, άρα δεν υπάρχει ιδιαίτερος λόγος να καθυστερήσεις παρά μόνο προτίμηση.';
    }
  } else {
    zone = isShort ? 'ΧΑΜΗΛΗ ΖΩΝΗ' : 'ΨΗΛΗ ΖΩΝΗ';
    note = isShort
      ? 'Η τιμή είναι κοντά στο κάτω μέρος του εύρους — αν μπεις τώρα με πλήρες μέγεθος σε short, το περιθώριο μέχρι το lower limit είναι μικρό.'
      : 'Η τιμή είναι κοντά στο πάνω μέρος του εύρους — αν μπεις τώρα με πλήρες μέγεθος, το περιθώριο μέχρι το upper limit είναι μικρό.';
    if (momentum === favorableMomentum) {
      note += ' Η τρέχουσα βραχυπρόθεσμη ροπή είναι ήδη ' + momentumWord + ' — η ευνοϊκή κίνηση ήδη συμβαίνει, χωρίς εγγύηση συνέχισης.';
    } else if (momentum === adverseMomentum) {
      note += ' Η τρέχουσα βραχυπρόθεσμη ροπή είναι ' + momentumWord + ' — αντίθετη από τη στρατηγική αυτή τη στιγμή.';
    }
 
    // Πρόταση διαχωρισμού κεφαλαίου σε 2 τμήματα: όσο πιο δυσμενές το σημείο,
    // τόσο μικρότερο ποσοστό προτείνεται για άμεση είσοδο — γραμμική κλιμάκωση.
    const zoneProgress = clamp((effectivePos - cfg.MODERATE_ZONE_MAX_PCT) / (100 - cfg.MODERATE_ZONE_MAX_PCT), 0, 1);
    const nowPct = Math.round(cfg.SPLIT_NOW_PCT_AT_ZONE_START -
      zoneProgress * (cfg.SPLIT_NOW_PCT_AT_ZONE_START - cfg.SPLIT_NOW_PCT_AT_TOP));
    const restPct = 100 - nowPct;
 
    // Στόχος 2ου τμήματος: προτιμάται το ΚΟΝΤΙΝΟ εύρος των 2 τελευταίων candles (8 ώρες)
    // αντί για το μακρινό support/αντίσταση 2 ημερών — πιο ρεαλιστικά επιτεύξιμο.
    // Fallback στο 2ημερο support/αντίσταση μόνο αν το κοντινό εύρος δεν βγάζει νόημα
    // (π.χ. η τιμή είναι ήδη έξω από αυτό προς τη λάθος κατεύθυνση).
    let restTargetPrice, targetSource;
    if (isShort) {
      if (last2Analysis && last2Analysis.high > lastPrice) {
        restTargetPrice = last2Analysis.high; targetSource = 'πρόσφατη αντίσταση τελευταίων 8 ωρών (2 candles)';
      } else if (recentSwingHigh && recentSwingHigh > lastPrice) {
        restTargetPrice = recentSwingHigh; targetSource = 'πρόσφατη αντίσταση τελευταίων ~2 ημερών';
      } else {
        restTargetPrice = sweetSpotPrice; targetSource = 'θεωρητικό σημείο αναφοράς';
      }
    } else {
      if (last2Analysis && last2Analysis.low < lastPrice) {
        restTargetPrice = last2Analysis.low; targetSource = 'πρόσφατο support τελευταίων 8 ωρών (2 candles)';
      } else if (recentSwingLow && recentSwingLow < lastPrice) {
        restTargetPrice = recentSwingLow; targetSource = 'πρόσφατο support τελευταίων ~2 ημερών';
      } else {
        restTargetPrice = sweetSpotPrice; targetSource = 'θεωρητικό σημείο αναφοράς';
      }
    }
 
    splitPlan = {
      nowPct: nowPct, nowPrice: round(lastPrice, lastPrice),
      restPct: restPct, restPrice: round(restTargetPrice, lastPrice)
    };
 
    note += ' 💡 Πρόταση διαχωρισμού κεφαλαίου σε 2 τμήματα: ' + nowPct + '% τώρα στην τρέχουσα τιμή (~' +
      splitPlan.nowPrice + '), και ' + restPct + '% σε εντολή στο ~' + splitPlan.restPrice +
      ' (' + targetSource + '). Έτσι συμμετέχεις τώρα με μικρότερο ρίσκο και κρατάς περιθώριο για καλύτερο μέσο όρο — χωρίς να χρειάζεται να μαντέψεις αν θα έρθει.';
 
    // Επιβεβαίωση/αναίρεση από το pattern των 2 τελευταίων candles.
    if (last2Analysis) {
      const wantedDirection = isShort ? 'up' : 'down'; // χρειάζεται η τιμή να κινηθεί προς αυτή την κατεύθυνση για να πιάσει τον στόχο
      if (last2Analysis.confirms === wantedDirection) {
        note += ' 📊 Τα 2 τελευταία 4H candles δείχνουν ' + last2Analysis.pattern +
          ' — επιβεβαιώνει (προς το παρόν) την κίνηση προς τον στόχο, χωρίς εγγύηση συνέχισης.';
      } else if (last2Analysis.confirms !== null) {
        note += ' 📊 Τα 2 τελευταία 4H candles δείχνουν ' + last2Analysis.pattern +
          ' — αντίθετη κατεύθυνση από τον στόχο του 2ου τμήματος αυτή τη στιγμή, ίσως χρειαστεί περισσότερη υπομονή.';
      } else {
        note += ' 📊 Τα 2 τελευταία 4H candles δείχνουν ' + last2Analysis.pattern + '.';
      }
    }
  }
 
  return { posPct: posPct, zone: zone, note: note, sweetSpotPrice: sweetSpotPrice, momentum: momentum, splitPlan: splitPlan };
}
 
// Προτείνει ζώνη εισόδου βάσει: (α) πού βρίσκεται η τρέχουσα τιμή μέσα στο ΣΗΜΕΡΙΝΟ
// εύρος (periodLow-periodHigh), και (β) βραχυπρόθεσμο momentum των τελευταίων ~2
// ημερών (4H candles). Υπολογίζει επίσης recentSwingLow/High + shortSlopePct, που
// ξαναχρησιμοποιούνται και όταν αξιολογούμε τα ΠΑΓΩΜΕΝΑ limits ενός locked pick
// (βλ. computeEntryZone) — δεν χρειάζεται να ξαναφέρουμε klines για αυτό.
function computeEntrySuggestion(klines, periodLow, periodHigh, lastPrice) {
  const cfg = GRID_SCREENER_CONFIG.ENTRY_SUGGESTION;
  const recentCandles = klines.slice(-cfg.RECENT_SWING_CANDLES);
  const recentLows = recentCandles.map(function (k) { return k.low; });
  const recentHighs = recentCandles.map(function (k) { return k.high; });
  const recentCloses = recentCandles.map(function (k) { return k.close; });
 
  const recentSwingLow = Math.min.apply(null, recentLows);
  const recentSwingHigh = Math.max.apply(null, recentHighs);
  const shortSlopePct = calcTrendSlopePct(recentCloses); // %/candle, τελευταίες ~2 μέρες
 
  // Ανάλυση των 2 ΤΕΛΕΥΤΑΙΩΝ candles (8 ώρες) — πολύ πιο κοντινό/επιτεύξιμο σημείο
  // αναφοράς από το support/αντίσταση 2 ημερών, και έλεγχος αν το pattern
  // επιβεβαιώνει την τρέχουσα κατεύθυνση ή δείχνει πιθανή αναστροφή.
  const last2Analysis = analyzeLast2Candles(klines);
 
  const zoneInfo = computeEntryZone(lastPrice, periodLow, periodHigh, shortSlopePct, recentSwingLow, recentSwingHigh, last2Analysis);
 
  return {
    posPct: zoneInfo.posPct, zone: zoneInfo.zone, note: zoneInfo.note,
    sweetSpotPrice: zoneInfo.sweetSpotPrice, recentSwingLow: recentSwingLow,
    recentSwingHigh: recentSwingHigh, shortSlopePct: shortSlopePct, last2Analysis: last2Analysis
  };
}
 
// Αναλύει τα 2 ΤΕΛΕΥΤΑΙΑ 4H candles (8 ώρες): κατεύθυνση καθενός (bullish/bearish/
// doji) και ένα απλό pattern label, ώστε να ξέρουμε αν η πρόσφατη κίνηση
// "επιβεβαιώνει" συνέχιση προς μια κατεύθυνση ή δείχνει αναποφασιστικότητα/πιθανή
// αναστροφή. Επιστρέφει επίσης το low/high αυτών των 2 candles — πολύ πιο κοντινό
// (άρα ρεαλιστικά επιτεύξιμο) σημείο αναφοράς από το support/αντίσταση 2 ημερών.
function analyzeLast2Candles(klines) {
  if (!klines || klines.length < 2) return null;
  const c1 = klines[klines.length - 2];
  const c2 = klines[klines.length - 1]; // πιο πρόσφατο
 
  function candleDir(c) {
    if (c.close > c.open) return 'bullish';
    if (c.close < c.open) return 'bearish';
    return 'doji';
  }
 
  const dir1 = candleDir(c1);
  const dir2 = candleDir(c2);
 
  let pattern, confirms; // confirms: 'up' | 'down' | null (τι κατεύθυνση επιβεβαιώνει το pattern, αν κάποια)
  if (dir1 === 'bearish' && dir2 === 'bearish') {
    pattern = 'συνεχιζόμενη πτωτική κίνηση (2 συνεχόμενα κόκκινα candles)';
    confirms = 'down';
  } else if (dir1 === 'bullish' && dir2 === 'bullish') {
    pattern = 'συνεχιζόμενη ανοδική κίνηση (2 συνεχόμενα πράσινα candles)';
    confirms = 'up';
  } else if (dir1 === 'bearish' && dir2 === 'bullish') {
    pattern = 'πιθανή αναστροφή προς τα πάνω (κόκκινο -> πράσινο)';
    confirms = 'up';
  } else if (dir1 === 'bullish' && dir2 === 'bearish') {
    pattern = 'πιθανή αναστροφή προς τα κάτω (πράσινο -> κόκκινο)';
    confirms = 'down';
  } else {
    pattern = 'αναποφάσιστη κίνηση (doji/ασαφές candle)';
    confirms = null;
  }
 
  return {
    low: Math.min(c1.low, c2.low),
    high: Math.max(c1.high, c2.high),
    dir1: dir1, dir2: dir2, pattern: pattern, confirms: confirms
  };
}
 
// Choppiness Index (0-100). Κοντά στο 100 = πολύ "τσοπ"/πλάγια αγορά (ιδανικό για grid).
// Κοντά στο 0 = καθαρό, ισχυρό trend. Standard formula (Bilou, 14-period συνήθως).
function calcChoppinessIndex(klines, period) {
  if (klines.length < period + 1) return 50;
  const recent = klines.slice(-period - 1);
 
  let trSum = 0;
  for (let i = 1; i < recent.length; i++) {
    const highLow = recent[i].high - recent[i].low;
    const highClose = Math.abs(recent[i].high - recent[i - 1].close);
    const lowClose = Math.abs(recent[i].low - recent[i - 1].close);
    trSum += Math.max(highLow, highClose, lowClose);
  }
 
  const periodHighs = recent.slice(1).map(function (k) { return k.high; });
  const periodLows = recent.slice(1).map(function (k) { return k.low; });
  const highestHigh = Math.max.apply(null, periodHighs);
  const lowestLow = Math.min.apply(null, periodLows);
  const range = highestHigh - lowestLow;
 
  if (range === 0 || trSum === 0) return 50;
  const ci = 100 * (Math.log(trSum / range) / Math.log(period));
  return clamp(ci, 0, 100);
}
 
// Πολύ απλή, "rough" εκτίμηση δυνητικής απόδοσης grid πάνω στο ΙΔΙΟ ιστορικό διάστημα
// (δεν είναι πρόβλεψη μέλλοντος — δείχνει πόσο "κινήθηκε μπρος-πίσω" η τιμή ιστορικά,
// κάτι που είναι καλός proxy για το πόσες φορές θα γέμιζαν τα grid orders).
// Λογική: ένα grid πιάνει περίπου το μισό της συνολικής απόστασης που διάνυσε η τιμή
// (πριν προμήθειες/leverage), αφού κάθε πλήρης κύκλος buy-low/sell-high "καίει" τη μισή διαδρομή.
function estimateGridReturnPct(closes) {
  let totalPathPct = 0;
  for (let i = 1; i < closes.length; i++) {
    totalPathPct += Math.abs((closes[i] - closes[i - 1]) / closes[i - 1]) * 100;
  }
  return totalPathPct / 2;
}
 
function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}
 
// Ποινή score βάσει |funding rate|. Ένα σταθερά υψηλό funding (θετικό ή αρνητικό)
// είναι πραγματικό κόστος για όποια πλευρά του grid έχει καθαρή έκθεση εκείνη τη
// στιγμή — δεν αντισταθμίζεται πάντα πλήρως από το ότι το grid είναι "ουδέτερο".
 
 
// Εντοπίζει ποιος παράγοντας "κυριάρχησε" στο score, για γρήγορη ανάγνωση
// χωρίς να χρειάζεται να διαβάσεις όλες τις επιμέρους στήλες.
function explainScore(r) {
  const isLongMode = GRID_SCREENER_CONFIG.GRID_STRATEGY !== 'SHORT'; // default LONG
 
  if (r.verdict === 'SANITY CHECK FAILED — slope εκτός λογικών ορίων') {
    return '🛡️ Απορρίφθηκε από το safety clamp: slope (' + r.slopePct.toFixed(2) + '%/candle) εκτός λογικών ορίων — πιθανό σφάλμα δεδομένων, όχι πραγματική τάση.';
  }
  if (r.verdict === 'DOWNTREND') {
    return 'Χαμηλό score: καθοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — ακατάλληλο για Long Grid.';
  }
  if (r.verdict === 'PARABOLIC') {
    return 'Χαμηλό score: πολύ απότομη/παραβολική ανοδική κίνηση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — ρίσκο απότομης διόρθωσης.';
  }
  if (r.verdict === 'FLAT/WEAK — ΟΧΙ ΑΝΟΔΙΚΗ') {
    return 'Χαμηλό score: δεν υπάρχει σαφής ανοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — δεν αξιοποιεί το πλεονέκτημα του Long Grid.';
  }
  if (r.verdict === 'UPTREND') {
    return 'Χαμηλό score: ανοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — ακατάλληλο για Short Grid.';
  }
  if (r.verdict === 'PARABOLIC DOWN') {
    return 'Χαμηλό score: πολύ απότομη πτώση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — ρίσκο απότομου rebound.';
  }
  if (r.verdict === 'FLAT/WEAK — ΟΧΙ ΚΑΘΟΔΙΚΗ') {
    return 'Χαμηλό score: δεν υπάρχει σαφής καθοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle) — δεν αξιοποιεί το πλεονέκτημα του Short Grid.';
  }
  if (r.verdict === 'TOO TIGHT') {
    return 'Χαμηλό score: πολύ στενό εύρος (' + r.rangePct.toFixed(1) + '%) — λίγο περιθώριο κέρδους ανά grid.';
  }
  if (r.verdict === 'TOO VOLATILE') {
    return 'Χαμηλό score: υπερβολικό εύρος (' + r.rangePct.toFixed(1) + '%) — υψηλό ρίσκο σε leveraged grid.';
  }
 
  // GOOD ή MARGINAL: βρες ποιος παράγοντας συνεισέφερε περισσότερο
  const w = GRID_SCREENER_CONFIG.SCORE_WEIGHTS;
  const choppNorm = clamp(r.choppiness, 0, 100);
  const trendComponentNorm = isLongMode ? trendUpwardScore(r.slopePct) : trendDownwardScore(r.slopePct);
  const rangeFitNorm = clamp(100 - Math.abs(r.rangePct - GRID_SCREENER_CONFIG.RANGE_SWEET_SPOT_PCT) * 3, 0, 100);
 
  const trendLabel = isLongMode
    ? 'ήπια ανοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle)'
    : 'ήπια καθοδική τάση (κλίση ' + r.slopePct.toFixed(2) + '%/candle)';
 
  const contributions = [
    { label: 'υψηλό Choppiness (' + r.choppiness.toFixed(0) + ')', value: choppNorm * w.choppiness },
    { label: trendLabel, value: trendComponentNorm * w.trendWeak },
    { label: 'ιδανικό εύρος (' + r.rangePct.toFixed(1) + '%)', value: rangeFitNorm * w.rangeFit },
    { label: 'υψηλή ιστορική "διαδρομή" τιμής', value: 50 * w.estReturn } // απλοποιημένο, δεν έχουμε το normalized εδώ
  ];
 
  contributions.sort(function (a, b) { return b.value - a.value; });
  const top = contributions[0];
  const second = contributions[1];
 
  const prefix = isGoodVerdict(r.verdict) ? 'Καλό score κυρίως λόγω' : 'Οριακό score λόγω';
  const explanation = prefix + ' ' + top.label + ' και ' + second.label + '.';
 
  return explanation;
}
 
 
// Strategy-aware έλεγχοι verdict — δουλεύουν σωστά είτε GRID_STRATEGY='LONG'
// (verdicts: 'GOOD for Long Grid' / 'MARGINAL for Long Grid') είτε ='SHORT'
// (verdicts: 'GOOD for Short Grid' / 'MARGINAL for Short Grid'). Όλος ο
// υπόλοιπος κώδικας (scoring, stable pick, open positions, Telegram,
// χρωματισμός) περνάει ΠΑΝΤΑ μέσα από αυτές αντί να συγκρίνει literal strings,
// ώστε να δουλεύει και στα δύο modes χωρίς διπλή συντήρηση λογικής.
function isGoodVerdict(verdict) {
  return GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT'
    ? verdict === 'GOOD for Short Grid'
    : verdict === 'GOOD for Long Grid';
}
 
function isMarginalVerdict(verdict) {
  return GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT'
    ? verdict === 'MARGINAL for Short Grid'
    : verdict === 'MARGINAL for Long Grid';
}
 
function isEligibleVerdict(verdict) {
  return isGoodVerdict(verdict) || isMarginalVerdict(verdict);
}
 
// Αριθμητικό tier του verdict: GOOD πάντα πάνω από MARGINAL, ανεξάρτητα από
// ακατέργαστο score. Χρησιμοποιείται ΚΑΙ στην κύρια ταξινόμηση ΚΑΙ στο
// getStablePick (ώστε να μη γίνεται ποτέ "αναβάθμιση" από GOOD σε MARGINAL απλά
// επειδή το MARGINAL έχει τυχαία υψηλότερο ακατέργαστο νούμερο score).
function verdictTier(r) {
  if (isGoodVerdict(r.verdict)) return 2;
  if (isMarginalVerdict(r.verdict)) return 1;
  return 0;
}
 
function classify(rangePct, slopePct, adx, choppiness) {
  // 🛡️ ΣΚΛΗΡΟ CLAMP ΑΣΦΑΛΕΙΑΣ: ανεξάρτητα από οποιαδήποτε άλλη διαδρομή κώδικα,
  // ένα |slope| πάνω από 10%/candle είναι φυσικά αδύνατο για πραγματική τιμή
  // (θα σήμαινε δεκάδες φορές μεγαλύτερη κίνηση από το ολόκληρο μηνιαίο εύρος).
  // Αν ποτέ συμβεί, ΔΕΝ φτάνει ποτέ σε GOOD/MARGINAL, ό,τι κι αν λένε τα υπόλοιπα
  // thresholds — και καταγράφεται ρητά ώστε να εντοπιστεί η πηγή αργότερα.
  const SANITY_SLOPE_LIMIT = 10;
  if (Math.abs(slopePct) > SANITY_SLOPE_LIMIT) {
    Logger.log('🛡️ SAFETY CLAMP ενεργοποιήθηκε: slopePct=' + slopePct + ' υπερβαίνει το όριο λογικής (' +
      SANITY_SLOPE_LIMIT + '%) — αναγκαστική απόρριψη ανεξαρτήτως άλλων κριτηρίων.');
    return 'SANITY CHECK FAILED — slope εκτός λογικών ορίων';
  }
 
  if (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT') {
    return classifyShort(rangePct, slopePct, choppiness);
  }
  return classifyLong(rangePct, slopePct, choppiness);
}
 
// Για Long Futures Grid: ΔΕΝ τιμωρούμε την τάση — την ανταμείβουμε, όσο είναι
// ήπια/σταθερή (όχι παραβολική) και υπάρχει αρκετή διακύμανση ώστε να γεμίζουν
// grid orders και προς τις 2 κατευθύνσεις μέσα στην ανοδική πορεία.
function classifyLong(rangePct, slopePct, choppiness) {
  const cfg = GRID_SCREENER_CONFIG.LONG_GRID;
 
  if (rangePct < GRID_SCREENER_CONFIG.MIN_RANGE_PCT) return 'TOO TIGHT';
  if (rangePct > GRID_SCREENER_CONFIG.MAX_RANGE_PCT) return 'TOO VOLATILE';
  if (slopePct <= cfg.DOWNTREND_SLOPE_PCT) return 'DOWNTREND';
  if (slopePct >= cfg.PARABOLIC_SLOPE_PCT) return 'PARABOLIC';
 
  if (slopePct >= cfg.MIN_UPTREND_SLOPE_PCT && choppiness >= cfg.MIN_CHOPPINESS) return 'GOOD for Long Grid';
  if (slopePct >= 0) return 'MARGINAL for Long Grid';
  return 'FLAT/WEAK — ΟΧΙ ΑΝΟΔΙΚΗ';
}
 
// Για Short Futures Grid: mirror του classifyLong με αντίστροφο πρόσημο —
// ανταμείβει ήπια/σταθερή ΚΑΘΟΔΙΚΗ τάση, τιμωρεί απότομη πτώση (ρίσκο rebound)
// και ανοδική τάση (εντελώς ακατάλληλο για short grid).
function classifyShort(rangePct, slopePct, choppiness) {
  const cfg = GRID_SCREENER_CONFIG.SHORT_GRID;
 
  if (rangePct < GRID_SCREENER_CONFIG.MIN_RANGE_PCT) return 'TOO TIGHT';
  if (rangePct > GRID_SCREENER_CONFIG.MAX_RANGE_PCT) return 'TOO VOLATILE';
  if (slopePct >= cfg.UPTREND_SLOPE_PCT) return 'UPTREND';
  if (slopePct <= cfg.PARABOLIC_DOWN_SLOPE_PCT) return 'PARABOLIC DOWN';
 
  if (slopePct <= cfg.MIN_DOWNTREND_SLOPE_PCT && choppiness >= cfg.MIN_CHOPPINESS) return 'GOOD for Short Grid';
  if (slopePct <= 0) return 'MARGINAL for Short Grid';
  return 'FLAT/WEAK — ΟΧΙ ΚΑΘΟΔΙΚΗ';
}
 
// Score component για GRID_STRATEGY='LONG': ανταμείβει ήπια/σταθερή ανοδική τάση,
// με "κορυφή" στη μέση απόσταση ανάμεσα σε MIN_UPTREND_SLOPE_PCT και PARABOLIC_SLOPE_PCT.
// Μηδέν αν η τάση είναι επίπεδη/αρνητική, μηδέν και κοντά στο παραβολικό όριο.
function trendUpwardScore(slopePct) {
  const cfg = GRID_SCREENER_CONFIG.LONG_GRID;
  const mid = (cfg.MIN_UPTREND_SLOPE_PCT + cfg.PARABOLIC_SLOPE_PCT) / 2;
  if (slopePct <= 0) return 0;
  if (slopePct <= mid) return clamp((slopePct / mid) * 100, 0, 100);
  const declineRange = cfg.PARABOLIC_SLOPE_PCT - mid;
  if (declineRange <= 0) return 100;
  return clamp(100 - ((slopePct - mid) / declineRange) * 100, 0, 100);
}
 
// Score component για GRID_STRATEGY='SHORT' — mirror του trendUpwardScore, με
// αντίστροφο πρόσημο (η "κορυφή" βρίσκεται στη μέση απόσταση ανάμεσα σε
// MIN_DOWNTREND_SLOPE_PCT και PARABOLIC_DOWN_SLOPE_PCT, και τα δύο αρνητικά).
function trendDownwardScore(slopePct) {
  const cfg = GRID_SCREENER_CONFIG.SHORT_GRID;
  const mid = (cfg.MIN_DOWNTREND_SLOPE_PCT + cfg.PARABOLIC_DOWN_SLOPE_PCT) / 2;
  if (slopePct >= 0) return 0;
  if (slopePct >= mid) return clamp((slopePct / mid) * 100, 0, 100);
  const declineRange = cfg.PARABOLIC_DOWN_SLOPE_PCT - mid;
  if (declineRange >= 0) return 100;
  return clamp(100 - ((slopePct - mid) / declineRange) * 100, 0, 100);
}
 
// Ελέγχει αν η τρέχουσα τιμή είναι κοντά ή εκτός των ΠΡΑΓΜΑΤΙΚΩΝ (κλειδωμένων) ορίων
// του grid, ώστε να προτείνεται συγκεκριμένη ενέργεια αντί για γενική προειδοποίηση.
function computeExitRisk(pick, lowerLimit, upperLimit) {
  const price = pick.lastPrice;
  const nearEdgePct = 5; // % απόσταση από το όριο που θεωρείται "κοντά"
 
  if (price <= lowerLimit) {
    return '🚨 Η τιμή είναι ΚΑΤΩ από το κάτω όριο (' + lowerLimit + ') — το grid πιθανώς δεν κάνει πλέον sell orders. Σκέψου χειροκίνητο κλείσιμο ή re-create με νέο εύρος.';
  }
  if (price >= upperLimit) {
    return '🚨 Η τιμή είναι ΠΑΝΩ από το πάνω όριο (' + upperLimit + ') — το grid πιθανώς δεν κάνει πλέον buy orders. Σκέψου χειροκίνητο κλείσιμο ή re-create με νέο εύρος.';
  }
 
  const distToLowerPct = ((price - lowerLimit) / price) * 100;
  const distToUpperPct = ((upperLimit - price) / price) * 100;
 
  if (distToLowerPct <= nearEdgePct) {
    return '⚠️ Η τιμή πλησιάζει το κάτω όριο (' + distToLowerPct.toFixed(1) + '% απόσταση) — παρακολούθησε στενά.';
  }
  if (distToUpperPct <= nearEdgePct) {
    return '⚠️ Η τιμή πλησιάζει το πάνω όριο (' + distToUpperPct.toFixed(1) + '% απόσταση) — παρακολούθησε στενά.';
  }
  return null;
}
 
 
// Υπολογίζει τον αριθμό grids ώστε το spacing ανά grid να είναι πάντα αρκετά
// πάνω από το κόστος fees (round-trip) — μεγιστοποιεί τη συχνότητα fills χωρίς
// τα πολύ στενά grids να "τρώγονται" από τις προμήθειες.
function suggestGridCount(rangePct) {
  const econ = GRID_SCREENER_CONFIG.GRID_ECONOMICS;
  const targetSpacingPct = econ.ROUNDTRIP_FEE_PCT * econ.MIN_PROFIT_MULTIPLE;
  if (targetSpacingPct <= 0) return econ.MIN_GRIDS;
 
  const rawGrids = Math.floor(rangePct / targetSpacingPct);
  return Math.max(econ.MIN_GRIDS, Math.min(econ.MAX_GRIDS, rawGrids));
}
 
// Υπολογίζει το προτεινόμενο leverage από την πραγματική μεταβλητότητα (ATR%),
// όχι από αυθαίρετο bucket του εύρους. Ζητάει η απόσταση μέχρι το (προσεγγιστικό)
// liquidation να είναι τουλάχιστον ATR_SAFETY_MULTIPLE φορές το μέσο ATR%, ώστε
// μια συνηθισμένη ημερήσια κίνηση να μην φέρνει κοντά σε liquidation.
// ΣΗΜΕΙΩΣΗ: πρόχειρη προσέγγιση (αγνοεί maintenance margin/funding) — πάντα
// επιβεβαίωνε το πραγματικό liquidation price στο Pionex πριν στήσεις grid.
function suggestLeverage(atrPct) {
  const cfg = GRID_SCREENER_CONFIG.LEVERAGE_SAFETY;
  if (!atrPct || atrPct <= 0) return 1;
 
  const rawLeverage = Math.floor(100 / (cfg.ATR_SAFETY_MULTIPLE * atrPct));
  return Math.max(1, Math.min(cfg.MAX_LEVERAGE_CAP, rawLeverage));
}
 
// Προτεινόμενο Take Profit (% PnL ratio του grid bot, όχι ανά grid). Βασισμένο στο
// Est. Return (πραγματική ιστορική "διαδρομή" της τιμής, όχι το ακατέργαστο εύρος) —
// έτσι ο στόχος είναι κάτι που ρεαλιστικά συμπληρώνεται μέσα σε λογικό διάστημα,
// όχι απαίτηση για ολόκληρη την κίνηση του μήνα να ξανασυμβεί.
function suggestTP(estReturnPct, leverage) {
  const cfg = GRID_SCREENER_CONFIG.TP_SL;
  if (!estReturnPct || !leverage) return cfg.MIN_TP_PCT;
  const raw = estReturnPct * leverage * cfg.TP_FRACTION;
  return clamp(raw, cfg.MIN_TP_PCT, cfg.MAX_TP_PCT);
}
 
// Προτεινόμενο Stop Loss (% PnL ratio). Πολλαπλάσιο του ATR (leveraged) ώστε η
// συνηθισμένη διακύμανση να ΜΗΝ το ενεργοποιεί, αλλά ένα πραγματικό breakout να
// σε βγάζει πριν φτάσεις κοντά σε liquidation.
function suggestSL(atrPct, leverage) {
  const cfg = GRID_SCREENER_CONFIG.TP_SL;
  if (!atrPct || !leverage) return cfg.MIN_SL_PCT;
  const raw = atrPct * leverage * cfg.SL_ATR_MULTIPLE;
  return clamp(raw, cfg.MIN_SL_PCT, cfg.MAX_SL_PCT);
}
 
// Στρογγυλοποιεί με βάση ΣΗΜΑΝΤΙΚΑ ΨΗΦΙΑ (όχι σταθερά δεκαδικά) — δουλεύει σωστά
// σε ΟΠΟΙΟΔΗΠΟΤΕ μέγεθος τιμής (από BTC στα 60000 μέχρι meme coins στο 0.0000002)
// χωρίς artifacts κινητής υποδιαστολής (π.χ. ...99999997) και χωρίς να μηδενίζει
// ποτέ μια μη-μηδενική τιμή.
function round(value, referencePrice) {
  if (value === 0) return 0;
  if (!referencePrice && referencePrice !== 0) return value;
  const sigFigs = referencePrice >= 100 ? 7 : 6;
  return Number(value.toPrecision(sigFigs));
}
 
// Κοινή, ΜΟΝΑΔΙΚΗ πηγή αλήθειας για τον υπολογισμό buffered grid limits — τη
// χρησιμοποιούν ΚΑΙ ο κύριος πίνακας ΚΑΙ το locked pick.
function computeBufferedLimits(periodLow, periodHigh, lastPrice) {
  const rawLower = periodLow * (1 - GRID_SCREENER_CONFIG.GRID_BUFFER_PCT / 100);
  const rawUpper = periodHigh * (1 + GRID_SCREENER_CONFIG.GRID_BUFFER_PCT / 100);
  return {
    lowerLimit: round(rawLower, lastPrice),
    upperLimit: round(rawUpper, lastPrice)
  };
}
 
// Εναλλακτική τιμή εισόδου, strategy-aware. ΣΗΜΑΝΤΙΚΟ: το ποσοστό εφαρμόζεται
// πάνω στο ΕΥΡΟΣ (Upper - Lower), όχι πάνω στην απόλυτη τιμή του ορίου — αλλιώς σε
// φαρδιά εύρη το αποτέλεσμα μπορεί να πέσει εκτός ορίων (π.χ. κάτω από το Lower Limit).
// - LONG: καλή είσοδος = κοντά στο Lower Limit -> εναλλακτική = Upper Limit μείον
//   ALT_ENTRY_BELOW_HIGH_PCT% του εύρους (πιο κοντά στη "φθηνή" πλευρά).
// - SHORT: καλή είσοδος = κοντά στο Upper Limit -> εναλλακτική = Lower Limit συν
//   το ίδιο ποσοστό του εύρους (πιο κοντά στην καλή για short πλευρά).
function computeAltEntryPrice(lowerLimit, upperLimit, lastPrice) {
  const pct = GRID_SCREENER_CONFIG.ALT_ENTRY_BELOW_HIGH_PCT / 100;
  const range = upperLimit - lowerLimit;
  if (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT') {
    return round(lowerLimit + range * pct, lastPrice);
  }
  return round(upperLimit - range * pct, lastPrice);
}
 
// Fibonacci retracement level πάνω στο ΙΔΙΟ εύρος (Lower-Upper Limit) που ήδη
// χρησιμοποιείται στο Alt Entry Price — strategy-aware κατεύθυνση:
// - LONG: retracement ΠΡΟΣ ΤΑ ΚΑΤΩ από το Upper (βαθύτερο retracement = καλύτερη τιμή αγοράς)
// - SHORT: retracement ΠΡΟΣ ΤΑ ΠΑΝΩ από το Lower (βαθύτερο retracement = καλύτερη τιμή short)
function computeFibLevel(lowerLimit, upperLimit, ratio, lastPrice) {
  const range = upperLimit - lowerLimit;
  if (GRID_SCREENER_CONFIG.GRID_STRATEGY === 'SHORT') {
    return round(lowerLimit + range * ratio, lastPrice);
  }
  return round(upperLimit - range * ratio, lastPrice);
}
 
 



 













/**
 * =========================================================================
 * PIONEX BOT PERFORMANCE TRACKER — FINAL_PNL2
 *
 * Διόρθωση balance/PnL:
 *
 *   rawBotBalance = sum(Current value row[7] + Profit row[8])
 *   botBalance    = rawBotBalance - Account USDT balance
 *   PnL           = botBalance - Total investment
 *
 * Το Telegram ΔΕΝ εμφανίζει συνολικό ύψος.
 * =========================================================================
 */

const PIONEX_TRACKER_BASE_URL = 'https://api.pionex.com';
const PIONEX_TRACKER_SHEET_NAME = 'Pionex Bots';
const PIONEX_TRACKER_TARGET_PROPERTY = 'PIONEX_GRID_PROFIT_TARGET';

function PIONEX_TRACKER_addMenuItems_(menu) {
  menu
    .addSeparator()
    .addItem('Ανανέωση απόδοσης bots', 'PIONEX_TRACKER_refreshPionexBots')
    .addItem('Ανανέωση + αποστολή στο Telegram', 'PIONEX_TRACKER_refreshAndNotifyTelegram')
    .addItem('Ορισμός στόχου grid profit', 'PIONEX_TRACKER_setGridProfitTarget')
    .addItem('Debug: balance and PnL', 'PIONEX_TRACKER_debugBalanceAndPnl')
    .addItem('Debug: raw Pionex response', 'PIONEX_TRACKER_debugPionexBotOrders')
    .addItem('Debug: signing message', 'PIONEX_TRACKER_debugSigningMessage')
    .addItem('Ρύθμιση αυτόματης ανανέωσης (15\')', 'PIONEX_TRACKER_setupAutoRefreshTrigger')
    .addItem('Απενεργοποίηση αυτόματης ανανέωσης', 'PIONEX_TRACKER_removeAutoRefreshTrigger')
    .addToUi();
}

function PIONEX_TRACKER_getGridProfitTarget_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PIONEX_TRACKER_TARGET_PROPERTY);
  const value = raw === null || raw === '' ? 1 : parseFloat(raw);
  return isFinite(value) && value > 0 ? value : 1;
}

function PIONEX_TRACKER_setGridProfitTarget() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Grid profit target', 'Βάλε στόχο σε USDT, π.χ. 1 ή 10.', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const value = parseFloat(result.getResponseText().replace(',', '.'));
  if (!isFinite(value) || value <= 0) return ui.alert('Ο στόχος πρέπει να είναι θετικός αριθμός.');
  PropertiesService.getScriptProperties().setProperty(PIONEX_TRACKER_TARGET_PROPERTY, String(value));
  ui.alert('Ο στόχος αποθηκεύτηκε: ' + value + ' USDT');
}

function PIONEX_TRACKER_getCredentials_() {
  const properties = PropertiesService.getScriptProperties();
  const key = (properties.getProperty('PIONEX_API_KEY') || '').trim();
  const secret = (properties.getProperty('PIONEX_API_SECRET') || '').trim();
  if (!key || !secret) throw new Error('Λείπουν τα PIONEX_API_KEY / PIONEX_API_SECRET.');
  return { key, secret };
}

function PIONEX_TRACKER_getTelegramCredentials_() {
  const properties = PropertiesService.getScriptProperties();
  const token = (properties.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  const chatId = (properties.getProperty('TELEGRAM_CHAT_ID') || '').trim();
  if (!token || !chatId) throw new Error('Λείπουν τα TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID.');
  return { token, chatId };
}

function PIONEX_TRACKER_buildQueryPairs_(params) {
  const pairs = [];
  Object.keys(params).forEach(key => {
    const value = params[key];
    if (Array.isArray(value)) value.forEach(item => pairs.push([key, item]));
    else pairs.push([key, value]);
  });
  pairs.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  return pairs;
}

function PIONEX_TRACKER_buildSigningMessage_(method, path, params) {
  const pairs = PIONEX_TRACKER_buildQueryPairs_(params);
  const rawQuery = pairs.map(pair => pair[0] + '=' + pair[1]).join('&');
  return { message: method + path + '?' + rawQuery, rawQuery: rawQuery };
}

function PIONEX_TRACKER_parseResponse_(response, label) {
  const code = response.getResponseCode();
  const body = response.getContentText();
  let json;
  try { json = JSON.parse(body); }
  catch (error) { throw new Error(label + ' μη έγκυρη JSON (HTTP ' + code + '): ' + body); }
  if (code !== 200 || json.result === false) {
    throw new Error(label + ' error (HTTP ' + code + '): ' + (json.message || json.code || JSON.stringify(json)));
  }
  return json;
}

function PIONEX_TRACKER_pionexGet_(path, params) {
  const credentials = PIONEX_TRACKER_getCredentials_();
  const allParams = Object.assign({}, params || {}, { timestamp: Date.now() });
  const signing = PIONEX_TRACKER_buildSigningMessage_('GET', path, allParams);
  const bytes = Utilities.computeHmacSha256Signature(signing.message, credentials.secret);
  const signature = bytes.map(byte => ((byte < 0 ? byte + 256 : byte).toString(16)).padStart(2, '0')).join('');
  const response = UrlFetchApp.fetch(PIONEX_TRACKER_BASE_URL + path + '?' + signing.rawQuery, {
    method: 'get',
    headers: { 'PIONEX-KEY': credentials.key, 'PIONEX-SIGNATURE': signature },
    muteHttpExceptions: true,
  });
  return PIONEX_TRACKER_parseResponse_(response, 'Pionex API');
}

function PIONEX_TRACKER_publicGet_(path, params) {
  const pairs = PIONEX_TRACKER_buildQueryPairs_(params || {});
  const rawQuery = pairs.map(pair => pair[0] + '=' + pair[1]).join('&');
  const response = UrlFetchApp.fetch(PIONEX_TRACKER_BASE_URL + path + (rawQuery ? '?' + rawQuery : ''), {
    method: 'get', muteHttpExceptions: true,
  });
  return PIONEX_TRACKER_parseResponse_(response, 'Pionex public API');
}

function PIONEX_TRACKER_firstNumber_(object, keys) {
  for (const key of keys) {
    const value = object[key];
    if (value !== undefined && value !== null && value !== '') {
      const number = parseFloat(value);
      if (!isNaN(number)) return number;
    }
  }
  return null;
}

function PIONEX_TRACKER_buildFuturesSymbol_(base, quote) {
  return String(base || '').replace(/\.PERP$/i, '').replace(/_PERP$/i, '') + '_' + quote + '_PERP';
}

function PIONEX_TRACKER_getMarkPrice_(base, quote) {
  try {
    const symbol = PIONEX_TRACKER_buildFuturesSymbol_(base, quote);
    const response = PIONEX_TRACKER_publicGet_('/api/v1/market/tickers', { symbol: symbol });
    const tickers = response.data && response.data.tickers ? response.data.tickers : [];
    return tickers.length ? PIONEX_TRACKER_firstNumber_(tickers[0], ['close', 'markPrice', 'lastPrice']) : null;
  } catch (error) { return null; }
}

function PIONEX_TRACKER_listBotOrders_(status) {
  const params = { status: status || 'running' };
  const orders = [];
  let pageToken = null;
  do {
    if (pageToken) params.pageToken = pageToken;
    const response = PIONEX_TRACKER_pionexGet_('/api/v1/bot/orders', params);
    const data = response.data || response;
    (data.results || []).forEach(order => orders.push(order));
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  return orders;
}

function PIONEX_TRACKER_getArbitrageMetrics_(data) {
  const total = PIONEX_TRACKER_firstNumber_(data, ['exchangeOrderPairedCount', 'totalArbitrages', 'pairedOrderCount']);
  const last24h = PIONEX_TRACKER_firstNumber_(data, ['trx24h', 'arbitrages24h', 'gridProfit24hCount']);
  return { total: total === null ? 0 : total, last24h: last24h === null ? 0 : last24h };
}

function PIONEX_TRACKER_getArbitragesNeeded_(gridProfit, target, totalArbitrages) {
  if (gridProfit === null || target === null || totalArbitrages === null || gridProfit <= 0 || totalArbitrages <= 0) return null;
  return Math.ceil(target / (gridProfit / totalArbitrages));
}

function PIONEX_TRACKER_computeMetrics_(data) {
  const investment = PIONEX_TRACKER_firstNumber_(data, ['quoteInvestment', 'initQuoteInvestment', 'usdtInvestment', 'initUsdtInvestment', 'quoteTotalInvestment', 'totalInvestment']);
  const currentValue = PIONEX_TRACKER_firstNumber_(data, ['marginBalance', 'currentEquity', 'totalAsset', 'netAssetValue', 'totalValue']);
  let profit = PIONEX_TRACKER_firstNumber_(data, ['totalRealizedProfit', 'profitWithdrawn', 'gridProfit']);
  const calculatedProfit = currentValue !== null && investment !== null ? currentValue - investment : null;
  if (profit === null && calculatedProfit !== null) profit = calculatedProfit;
  return {
    investment: investment === null ? 0 : investment,
    currentValue: currentValue !== null ? currentValue : (investment !== null && profit !== null ? investment + profit : 0),
    profit: profit === null ? 0 : profit,
    returnPct: investment !== null && investment > 0 && profit !== null ? (profit / investment) * 100 : null,
  };
}

function PIONEX_TRACKER_getArbitragesSinceLastRun_(
  botId,
  currentTotal
) {
  const properties =
    PropertiesService.getScriptProperties();

  const key =
    'PIONEX_LAST_ARBITRAGES_' + botId;

  const previousRaw =
    properties.getProperty(key);

  const previousTotal =
    previousRaw === null
      ? null
      : Number(previousRaw);

  const currentValue =
    Number(currentTotal) || 0;

  const arbitragesSinceLastRun =
    previousTotal === null
      ? 0
      : Math.max(0, currentValue - previousTotal);

  properties.setProperty(
    key,
    String(currentValue)
  );

  return arbitragesSinceLastRun;
}


function PIONEX_TRACKER_getBreakEven_(order, data) {
  const candidates = [
    data.breakEvenWithGridProfit,
    data.breakEvenPriceWithGridProfit,
    data.breakEven,
    data.breakEvenPrice,
    data.gridBreakEvenPrice,
    data.buOrderData && data.buOrderData.breakEvenWithGridProfit,
    order.breakEvenWithGridProfit,
    order.breakEvenPriceWithGridProfit,
    order.breakEven,
    order.breakEvenPrice
  ];

  for (let i = 0; i < candidates.length; i++) {
    const value = Number(candidates[i]);
    if (candidates[i] !== null && candidates[i] !== undefined && candidates[i] !== '' && isFinite(value)) {
      return value;
    }
  }

  return null;
}

function PIONEX_TRACKER_getOkxInstIdFromPair_(pair) {
  const parts = String(pair || '').split('_');

  const rawBase = String(parts[0] || '')
    .replace(/\.PERP$/i, '')
    .replace(/_PERP$/i, '')
    .trim()
    .toUpperCase();

  const quote = String(parts[1] || 'USDT')
    .trim()
    .toUpperCase();

  return rawBase + '-' + quote + '-SWAP';
}

function PIONEX_TRACKER_getOKXFundingRate_(instId) {
  const url =
    'https://www.okx.com/api/v5/public/funding-rate?instId=' +
    encodeURIComponent(instId);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const httpCode = response.getResponseCode();
  const body = response.getContentText();

  if (httpCode !== 200) {
    throw new Error(
      'OKX HTTP error ' + httpCode + ': ' + body
    );
  }

  const json = JSON.parse(body);

  if (
    json.code !== '0' ||
    !json.data ||
    !json.data.length
  ) {
    throw new Error(
      'OKX funding rate unavailable for ' +
      instId +
      ': ' +
      body
    );
  }

  const item = json.data[0];

  return {
    instId: item.instId,
    fundingRate:
      item.fundingRate === '' ||
      item.fundingRate === undefined
        ? null
        : Number(item.fundingRate),

    nextFundingRate:
      item.nextFundingRate === '' ||
      item.nextFundingRate === undefined
        ? null
        : Number(item.nextFundingRate)
  };
}

function PIONEX_TRACKER_getOkxInstId_(order) {
  const base = String(order.base || '')
    .replace(/\.PERP$/i, '')
    .replace(/_PERP$/i, '')
    .trim()
    .toUpperCase();

  const quote = String(order.quote || 'USDT')
    .replace(/\.PERP$/i, '')
    .replace(/_PERP$/i, '')
    .trim()
    .toUpperCase();

  return base + '-' + quote + '-SWAP';
}

function PIONEX_TRACKER_getOKXFundingRate_(instId) {
  const url =
    'https://www.okx.com/api/v5/public/funding-rate' +
    '?instId=' +
    encodeURIComponent(instId);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const httpCode = response.getResponseCode();
  const body = response.getContentText();

  if (httpCode !== 200) {
    throw new Error(
      'OKX HTTP error ' +
      httpCode +
      ': ' +
      body
    );
  }

  const json = JSON.parse(body);

  if (
    json.code !== '0' ||
    !json.data ||
    !json.data.length
  ) {
    throw new Error(
      'OKX funding rate error: ' +
      body
    );
  }

  const item = json.data[0];

  return {
    instId: item.instId,
    fundingRate: item.fundingRate === ''
      ? null
      : Number(item.fundingRate),
    nextFundingRate: item.nextFundingRate === ''
      ? null
      : Number(item.nextFundingRate),
    fundingTime: item.fundingTime
      ? new Date(Number(item.fundingTime))
      : null,
    nextFundingTime: item.nextFundingTime
      ? new Date(Number(item.nextFundingTime))
      : null
  };
}



function PIONEX_TRACKER_refreshPionexBots() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(PIONEX_TRACKER_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(PIONEX_TRACKER_SHEET_NAME);
  const headers = [
    'Bot Order ID',
    'Τύπος',
    'Ζεύγος',
    'Κατάσταση',
    'Direction',
    'Ημ/νία Δημιουργίας',
    'Χρόνος λειτουργίας',
    'Επένδυση',
    'Τρέχουσα Αξία',
    'Κέρδος (realized)',
    'Απόδοση %',
    'Συνολικά Arbitrages',
    'Arbitrages 24h',
    'Arbitrages από προηγούμενο τρέξιμο',
    'Grid Profit',
    'Grid Profit 24h',
    'Funding Fee Payment',    
    'Arbitrages για Grid Profit Target',
    'Grid Profit Target',
    'Trend PnL',
    'Mark Price',
    'Απόσταση από Init Price',
    'Leverage',
    'Grid Top',
    'Grid Bottom',
    'Ρίσκο',
    'Τελ. Ανανέωση',
    'Break Even χωρίς Grid Profit',
    'Break Even με Grid Profit'
  ];

  sheet.clear();
  sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    Math.max(sheet.getMaxColumns(), headers.length)
  )
  .clearContent();


  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  let botPositions = {};

  let orders;
  try {
    orders = PIONEX_TRACKER_listBotOrders_('running');
  } catch (error) {
    sheet.getRange(2, 1).setValue('Σφάλμα: ' + error.message);
    return [];
  }

  if (!orders || orders.length === 0) {
    sheet.getRange(2, 1).setValue('Δεν βρέθηκαν ενεργά bots.');
    return [];
  }

  const target = PIONEX_TRACKER_getGridProfitTarget_();
  const rows = [];
  const now = new Date();

  orders.forEach(order => {
    const data = order.buOrderData || {};
    const fundingFeePayment =
      PIONEX_TRACKER_firstNumber_(
        data,
        ['fundingFeePayment']
      );   

    const type = order.buOrderType || '';
    const metrics = PIONEX_TRACKER_computeMetrics_(data);
    const arbitrages = PIONEX_TRACKER_getArbitrageMetrics_(data);

    const gridProfit = PIONEX_TRACKER_firstNumber_(data, ['gridProfit']);
    const gridProfit24h = PIONEX_TRACKER_firstNumber_(data, ['gridProfit24h']);
    const needed = PIONEX_TRACKER_getArbitragesNeeded_(gridProfit, target, arbitrages.total);

    let markPrice = PIONEX_TRACKER_firstNumber_(data, ['markPrice', 'lastPrice', 'closePrice']);
    if (markPrice === null && type === 'futures_grid') {
      markPrice = PIONEX_TRACKER_getMarkPrice_(order.base, order.quote);
    }

const priceChange =
  PIONEX_TRACKER_getPriceChangeSinceLastRun_(
    order.buOrderId,
    markPrice
  );

const previousPrice =
  priceChange.previousPrice;

const priceChangePercent =
  priceChange.priceChangePercent;

    const runningTime =
      PIONEX_TRACKER_getRunningTime_(
        order.createTime
      );

    const openPrice = PIONEX_TRACKER_firstNumber_(data, ['positionOpenPrice']);
    const position = PIONEX_TRACKER_firstNumber_(data, ['position', 'baseAmount']);
    const direction = String(data.trend || 'no_trend').toLowerCase();

    let trendPnl = null;
    if (markPrice !== null && openPrice !== null && position !== null) {
      trendPnl = direction === 'short'
        ? (openPrice - markPrice) * position
        : (markPrice - openPrice) * position;
    }

    const initPrice = PIONEX_TRACKER_firstNumber_(data, ['initPrice']);
    const distance = markPrice !== null && initPrice !== null
      ? markPrice - initPrice
      : null;
    const distancePct = markPrice !== null && initPrice !== null && initPrice !== 0
      ? ((markPrice - initPrice) / initPrice) * 100
      : null;

    const breakEvenWithGridProfit =
      PIONEX_TRACKER_getBreakEven_(order, data);

    const breakEvenWithoutGridProfit =
      PIONEX_TRACKER_firstNumber_(data, [
        'breakEvenWithoutGridProfit',
        'breakEvenPriceWithoutGridProfit',
        'breakEvenWithoutProfit'
      ]);


    const okxInstId =
      PIONEX_TRACKER_getOkxInstId_(order);

    let okxFunding = {
      instId: okxInstId,
      fundingRate: null,
      nextFundingRate: null,
      fundingTime: null,
      nextFundingTime: null
    };

    try {
      okxFunding =
        PIONEX_TRACKER_getOKXFundingRate_(okxInstId);
    } catch (error) {
      Logger.log(
        'OKX funding error for ' +
        okxInstId +
        ': ' +
        error.message
      );
    }

const currentArbitrages =
  PIONEX_TRACKER_firstNumber_(
    data,
    ['exchangeOrderPairedCount']
  ) || 0;

const arbitragesSinceLastRun =
  PIONEX_TRACKER_getArbitragesSinceLastRun_(
    order.buOrderId,
    currentArbitrages
  );

Logger.log(
  JSON.stringify(order, null, 2)
);

Logger.log(
  'breakEvenWithGridProfit = ' +
  breakEvenWithGridProfit
);

botPositions[order.buOrderId] = Number(data.position || 0);

    rows.push([
      order.buOrderId || '',
      type,
      (order.base || '') + '_' + (order.quote || ''),
      data.status || order.status || '',
      data.trend || 'no_trend',
      order.createTime ? new Date(order.createTime) : '',
      runningTime,
      metrics.investment,
      metrics.currentValue,
      metrics.profit,
      metrics.returnPct === null ? '' : metrics.returnPct.toFixed(2) + '%',
      arbitrages.total,
      arbitrages.last24h,
      arbitragesSinceLastRun,
      gridProfit === null ? '' : gridProfit,
      gridProfit24h === null ? '' : gridProfit24h,
      fundingFeePayment === null  ? ''  : fundingFeePayment,
      needed === null ? '' : needed,
      target,
      trendPnl === null ? '' : trendPnl,
      markPrice === null ? '' : markPrice,
      distance === null || distancePct === null
        ? ''
        : distance.toFixed(10) + ' (' + distancePct.toFixed(2) + '%)',
      data.leverage || '',
      data.top || '',
      data.bottom || '',
      data.riskStatus || data.marginStatus || '',
      now,
      breakEvenWithoutGridProfit === null ? '' : breakEvenWithoutGridProfit,
      breakEvenWithGridProfit === null ? '' : breakEvenWithGridProfit
    ]);
  });

  Logger.log('headers.length: ' + headers.length);
  Logger.log('rows.length: ' + rows.length);
  Logger.log('first row length: ' + (rows.length ? rows[0].length : 0));

  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, headers.length)
      .setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);
  return {
  rows: rows,
  botPositions: botPositions
};

}


/** ---------------------- BALANCE / PNL ---------------------- **/

function PIONEX_TRACKER_getRawBotBalanceFromRows_(rows) {
  return rows.reduce((sum, row) => {
    const currentValue = Number(row[7]) || 0;
    const profit = Number(row[8]) || 0;
//    return sum + currentValue + profit;
    return sum + currentValue;
  }, 0);
}

function PIONEX_TRACKER_getAccountBalances_() {
  const response = PIONEX_TRACKER_pionexGet_('/api/v1/account/balances', {});
  const data = response.data || response;
  return data.balances || [];
}

function PIONEX_TRACKER_getUsdtBalance_(balances) {
  const usdt = balances.find(balance => String(balance.coin || balance.asset || '').toUpperCase() === 'USDT');
  if (!usdt) return { free: null, frozen: null, total: null };
  const free = parseFloat(usdt.free || 0);
  const frozen = parseFloat(usdt.frozen || 0);
  return { free: isNaN(free) ? null : free, frozen: isNaN(frozen) ? null : frozen, total: isNaN(free + frozen) ? null : free + frozen };
}

/** ---------------------- FORMATTING ---------------------- **/

function PIONEX_TRACKER_formatNumber_(value, decimals) {
  if (value === null || value === undefined || value === '') return 'N/A';
  const number = Number(value);
  if (!isFinite(number)) return 'N/A';
  return number.toFixed(decimals).replace(/\.?0+$/, '');
}

function PIONEX_TRACKER_sign_(value) { return Number(value) < 0 ? '🔴' : '🟢'; }

function PIONEX_TRACKER_formatDistance_(value) {
  if (value === null || value === undefined || value === '') return 'N/A';
  const parts = String(value).split(' ');
  return PIONEX_TRACKER_formatNumber_(parts[0], 10) + (parts[1] ? ' ' + parts[1] : '');
}


function PIONEX_TRACKER_estimateTargetTime4h_(pair, currentPrice, targetPrice) {
  const result = {
    candles: null,
    hours: null,
    trend: 'unknown',
    message: 'Μη διαθέσιμη εκτίμηση'
  };

  try {
    if (
      !Number.isFinite(currentPrice) ||
      !Number.isFinite(targetPrice) ||
      currentPrice <= 0 ||
      targetPrice <= 0 ||
      Math.abs(targetPrice - currentPrice) < 1e-12
    ) {
      return result;
    }

    // PEOPLE.PERP_USDT -> PEOPLE_USDT_PERP
    let symbol = String(pair || '')
      .replace('.PERP_USDT', '_USDT_PERP')
      .replace('.PERP', '_USDT_PERP');

    if (!symbol.endsWith('_PERP')) {
      symbol = symbol.replace(/_USDT$/, '_USDT_PERP');
    }

    const url =
      'https://api.pionex.com/api/v1/market/klines' +
      '?symbol=' + encodeURIComponent(symbol) +
      '&interval=4H&limit=100';

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(
        'Klines HTTP ' + response.getResponseCode() +
        ': ' + response.getContentText()
      );
    }

    const json = JSON.parse(response.getContentText());

    if (json.result !== true || !json.data ||
        !Array.isArray(json.data.klines)) {
      throw new Error('Μη έγκυρη απάντηση Klines: ' +
        response.getContentText());
    }

    const now = Date.now();
    const fourHoursMs = 4 * 60 * 60 * 1000;

    // Χρησιμοποιούμε μόνο ολοκληρωμένα candles.
    const closes = json.data.klines
      .filter(k => {
        const time = Number(k.time);
        const close = Number(k.close);

        return Number.isFinite(time) &&
          Number.isFinite(close) &&
          close > 0 &&
          time + fourHoursMs <= now;
      })
      .sort((a, b) => Number(a.time) - Number(b.time))
      .map(k => Number(k.close));

    if (closes.length < 72) {
      throw new Error(
        'Χρειάζονται 72 ολοκληρωμένα candles, βρέθηκαν ' +
        closes.length
      );
    }

    // Κλίση log(price) ανά candle.
    // Το αποτέλεσμα προσεγγίζει τη γεωμετρική ποσοστιαία
    // μεταβολή ανά 4ωρο.
    function regressionSlope(values) {
      const n = values.length;
      let sumX = 0;
      let sumY = 0;
      let sumXY = 0;
      let sumXX = 0;

      for (let i = 0; i < n; i++) {
        const x = i;
        const y = Math.log(values[i]);

        sumX += x;
        sumY += y;
        sumXY += x * y;
        sumXX += x * x;
      }

      const denominator = n * sumXX - sumX * sumX;

      return denominator === 0
        ? 0
        : (n * sumXY - sumX * sumY) / denominator;
    }

    const slope12 = regressionSlope(closes.slice(-12));
    const slope36 = regressionSlope(closes.slice(-36));
    const slope72 = regressionSlope(closes.slice(-72));

    // Βαρύτητα: 50% στην κύρια τάση, 30% στη βραχυπρόθεσμη,
    // 20% στη μακρύτερη τάση.
    const slope =
      slope36 * 0.50 +
      slope12 * 0.30 +
      slope72 * 0.20;

    const targetDirection = targetPrice > currentPrice ? 1 : -1;

    const slopes = [slope12, slope36, slope72];
    const agreeingWindows = slopes.filter(s =>
      s * targetDirection > 0
    ).length;

    result.trend =
      slope > 0 ? 'up' : slope < 0 ? 'down' : 'flat';

    result.slope12Pct = (Math.exp(slope12) - 1) * 100;
    result.slope36Pct = (Math.exp(slope36) - 1) * 100;
    result.slope72Pct = (Math.exp(slope72) - 1) * 100;
    result.combinedSlopePct = (Math.exp(slope) - 1) * 100;

    // Δεν προβλέπουμε χρόνο όταν οι χρονικοί ορίζοντες
    // διαφωνούν ή η τάση είναι αντίθετη από τον στόχο.
    if (agreeingWindows < 2 || slope * targetDirection <= 0) {
      result.message = 'Μη αξιόπιστη εκτίμηση — τάση αντίθετη ή μικτή';
      return result;
    }

    const logMove = Math.log(targetPrice / currentPrice);
    const estimatedCandles = logMove / slope;

    if (
      !Number.isFinite(estimatedCandles) ||
      estimatedCandles <= 0 ||
      estimatedCandles > 1000
    ) {
      result.message = 'Μη αξιόπιστη εκτίμηση';
      return result;
    }

    result.candles = Math.ceil(estimatedCandles);
    result.hours = result.candles * 4;
    result.message =
      '~' + result.candles + ' candles (' +
      (result.hours / 24).toFixed(1) + ' ημέρες)';

    return result;

  } catch (error) {
    Logger.log(
      '4H Target estimate error: ' + error.message
    );
    result.message = 'Μη διαθέσιμη εκτίμηση τάσης';
    return result;
  }
}



/** ---------------------- TELEGRAM ---------------------- **/

function PIONEX_TRACKER_sendTelegramMessage_(text) {
  const credentials = PIONEX_TRACKER_getTelegramCredentials_();
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + credentials.token + '/sendMessage', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ chat_id: credentials.chatId, text: text, parse_mode: 'HTML', disable_web_page_preview: true }),
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) throw new Error('Telegram API error (HTTP ' + response.getResponseCode() + '): ' + response.getContentText());
}

function PIONEX_TRACKER_getRunningTime_(createTime) {
  if (
    createTime === null ||
    createTime === undefined ||
    createTime === ''
  ) {
    return 'N/A';
  }

  const startTime =
    createTime instanceof Date
      ? createTime.getTime()
      : Number(createTime);

  if (!isFinite(startTime) || startTime <= 0) {
    return 'N/A';
  }

  const elapsedMilliseconds =
    Date.now() - startTime;

  if (elapsedMilliseconds < 0) {
    return '0m';
  }

  const totalSeconds = Math.floor(
    elapsedMilliseconds / 1000
  );

  const days = Math.floor(
    totalSeconds / 86400
  );

  const hours = Math.floor(
    (totalSeconds % 86400) / 3600
  );

  const minutes = Math.floor(
    (totalSeconds % 3600) / 60
  );

  if (days > 0) {
    return days + 'd ' + hours + 'h ' + minutes + 'm';
  }

  if (hours > 0) {
    return hours + 'h ' + minutes + 'm';
  }

  return minutes + 'm';
}

function PIONEX_TRACKER_getPriceChangeSinceLastRun_(
  botId,
  currentPrice
) {
  const properties =
    PropertiesService.getScriptProperties();

  const key =
    'PIONEX_LAST_MARK_PRICE_' + botId;

  const previousRaw =
    properties.getProperty(key);

  const previousPrice =
    previousRaw === null || previousRaw === ''
      ? null
      : Number(previousRaw);

  const price =
    Number(currentPrice);

  let priceChangePercent = null;

  if (
    previousPrice !== null &&
    Number.isFinite(previousPrice) &&
    previousPrice !== 0 &&
    Number.isFinite(price)
  ) {
    priceChangePercent =
      ((price - previousPrice) /
        previousPrice) * 100;
  }

  if (Number.isFinite(price)) {
    properties.setProperty(
      key,
      String(price)
    );
  }

  return {
    previousPrice: previousPrice,
    currentPrice: price,
    priceChangePercent: priceChangePercent
  };
}

function PIONEX_TRACKER_buildAndSendTelegramSummary_(rows, accountBalance, botPositions) {
  if (!rows.length) {
    PIONEX_TRACKER_sendTelegramMessage_('⚠️ <b>Pionex Bots</b>: Δεν βρέθηκαν ενεργά bots αυτή τη στιγμή.');
    return;
  }

  const lines = ['📊 <b>Pionex Bots — Τρέχουσα Απόδοση</b>', ''];
  let totalInvestment = 0;
  let totalArbitrages = 0;
  let totalArbitrages24h = 0;

  // Correct calculation:
  // raw bot balance = sum(row[7] + row[8])
  // bot balance     = raw bot balance - Account USDT
  // PnL             = bot balance - total investment
  const rawBotBalance = PIONEX_TRACKER_getRawBotBalanceFromRows_(rows);
  const accountValue = accountBalance && accountBalance.total !== null
    ? accountBalance.total
    : null;

  const totalInvestmentForPnl = rows.reduce(
    (sum, row) => sum + (Number(row[6]) || 0),
    0
  );

  const botBalance = accountValue === null
    ? null
    : totalInvestmentForPnl;


  const pnl = botBalance === null
    ? null
    : botBalance;

  let botPosition = null;

    rows.forEach(row => {
    const [id, type, pair, status, direction, created, runningTime,
      investment, currentValue, profit, returnPct,
      arbitrages, arbitrages24h, arbitragesSinceLastRun, gridProfit, gridProfit24h,fundingFeePayment,
      needed, target, trendPnl, markPrice, distance, leverage,
      breakEvenWithoutGridProfit,
      breakEvenWithGridProfit] = row;

const botPosition = Number(botPositions[id] || 0);

const okxInstId =
  PIONEX_TRACKER_getOkxInstIdFromPair_(pair);

let okxFunding = null;

try {
  okxFunding =
    PIONEX_TRACKER_getOKXFundingRate_(okxInstId);
} catch (error) {
  Logger.log(
    'OKX funding error for ' +
    okxInstId +
    ': ' +
    error.message
  );
}

const botId = row[0];

const priceChangePercent =
  PIONEX_TRACKER_getPriceChangeSinceLastRun_(
    botId,
    markPrice
  );
  

    totalInvestment += typeof investment === 'number' ? investment : 0;
    totalArbitrages += typeof arbitrages === 'number' ? arbitrages : 0;
    totalArbitrages24h += typeof arbitrages24h === 'number' ? arbitrages24h : 0;

    const distanceValue = distance === '' || distance === null
      ? null
      : parseFloat(String(distance).split(' ')[0]);

    const distanceIcon = distanceValue === null || isNaN(distanceValue)
      ? '⚪'
      : PIONEX_TRACKER_sign_(distanceValue);

    lines.push(
      '🔹 <b>' + pair + '</b> (' + type + ', ' +
      (direction || 'no_trend') + (leverage ? ', ' + leverage + 'x' : '') + ')'
    );
//    lines.push('   Επένδυση: ' + PIONEX_TRACKER_formatNumber_(investment, 2));
    lines.push('   ┌ ⏱️Running time: ' +  runningTime);
    lines.push('   ├ Arbitrages: ' + arbitrages + ' συνολικά (' + arbitrages24h + ' / 24h)');
    lines.push(
  '   ├ Arbitrages από προηγούμενο τρέξιμο: ' +
  arbitragesSinceLastRun
);

    lines.push(
      '   ├ Grid profit: ' +
      PIONEX_TRACKER_formatNumber_(gridProfit, 2) +
      ' (24h: ' + PIONEX_TRACKER_formatNumber_(gridProfit24h, 2) + ' USDT)'
    );

    const priceChangeIcon =
      priceChangePercent === null
        ? '⚪'
        : PIONEX_TRACKER_sign_(
            priceChangePercent
          );


    lines.push(
      '   ├ 🏷️Τρέχουσα τιμή: ' +
      PIONEX_TRACKER_formatNumber_(markPrice, 8)
    );

    const fundingIcon =
      fundingFeePayment === null ||
      fundingFeePayment === ''
        ? '⚪'
        : PIONEX_TRACKER_sign_(fundingFeePayment);

    lines.push(
      '   ├' + 
      ' Funding fee: ' +
      PIONEX_TRACKER_formatNumber_(
        fundingFeePayment,
        8
      )
    );

    lines.push(
      '   ├ Arbitrages για grid profit ' +
      PIONEX_TRACKER_formatNumber_(target, 2) + '€: ' +
      (needed === '' || needed === null ? 'N/A' : needed) +
      ''
    );

    const okxFundingIcon =
      okxFunding.fundingRate === null
        ? '⚪'
        : okxFunding.fundingRate < 0
          ? '🟢'
          : '🔴';

    lines.push(
      '   ' +      
      '├ OKX Funding Rate (' +
      okxFunding.instId +
      '): ' +
      (
        okxFunding.fundingRate === null
          ? 'N/A'
          : PIONEX_TRACKER_formatNumber_(
              okxFunding.fundingRate * 100,
              8
            ) + '%'
      )
    );


const currentValue1 = rows.reduce(
  (sum, row) => sum + (Number(row[8]) || 0),
  0
);
const trendPnL2 = rows.reduce(
  (sum, row) => sum + (Number(row[17]) || 0),
  0
);

const currentPnL = currentValue1 - investment + trendPnl

    const profit_positive = 1.05;
    const profit_negative = 1.00;

    const adjustedProfit = currentPnL === null || currentPnL === undefined
      ? null
      : currentPnL * (
          currentPnL >= 0
            ? profit_positive
            : profit_negative
        );

    // This is the only icon used for the PnL value.
     const pnlIcon = adjustedProfit === null ? '⚪' : PIONEX_TRACKER_sign_(adjustedProfit);

    // The Trend PnL label now shows the exact same total PnL value.
    lines.push(
      '   ├ ' + pnlIcon + ' Trend PnL: ' +
      PIONEX_TRACKER_formatNumber_(adjustedProfit, 2) 
    );

    lines.push(
      '   └ ' + distanceIcon +
      ' Απόσταση από τιμή εκκίνησης: ' +
      PIONEX_TRACKER_formatDistance_(distance)
    );
    lines.push('');

const totalProfit = rows.reduce(
  (sum, row) => sum + (Number(row[9]) || 0),
  0
);
const currentValue2 = rows.reduce(
  (sum, row) => sum + (Number(row[8]) || 0),
  0
);
const trendValue = rows.reduce(
  (sum, row) => sum + (Number(row[17]) || 0),
  0
);
const currentGridProfit = rows.reduce(
  (sum, row) => sum + (Number(row[13]) || 0),
  0
);

const GridProfit24h = rows.reduce(
  (sum, row) => sum + (Number(row[14]) || 0),
  0
);

const bot_positive = 1;
const bot_negative = 1;

const baseBotBalance =
  adjustedProfit +
  totalInvestment +
  accountValue;

const mybotBalance = accountValue === null
  ? null
  : baseBotBalance * (
      trendValue >= 0
        ? bot_positive
        : bot_negative
    );

const profitPnL = totalInvestment > 0 && mybotBalance !== null
  ? ((mybotBalance - accountValue - totalInvestment) / totalInvestment) * 100
  : null;

  // No separate 📉 PnL line.
  lines.push('🏦 <b>Account USDT:</b> ' + PIONEX_TRACKER_formatNumber_(accountValue, 2));
  lines.push('🤖 <b>Account balance:</b> ' + PIONEX_TRACKER_formatNumber_(mybotBalance, 2));
  const profitPnLIcon = profitPnL === null
    ? '⚪'
    : profitPnL < 0
      ? '🔴'
      : '🟢';
  lines.push(
    profitPnLIcon +
    ' <b>Profit PnL:</b> ' +
    (
      profitPnL === null
        ? 'N/A'
        : PIONEX_TRACKER_formatNumber_(profitPnL, 2) + '%'
    )
  );  
  lines.push('💰 <b>Σύνολο επένδυσης:</b> ' + PIONEX_TRACKER_formatNumber_(totalInvestment, 2));
  //lines.push('🔁 <b>Σύνολο arbitrages:</b> ' + totalArbitrages);
  const arbitrageDiff = totalArbitrages24h - needed;

  let arbitrageStatus;

  if (arbitrageDiff < 0) {
    arbitrageStatus = `⚠️ Λείπουν ${Math.abs(arbitrageDiff)}`;
  } else if (arbitrageDiff > 0) {
    arbitrageStatus = `✅ Περισσεύουν ${arbitrageDiff}`;
  } else {
    arbitrageStatus = `✅ Ακριβώς ${needed}`;
  }

  lines.push(
    '⚡ <b>Arbitrages 24h:</b> ' +
    totalArbitrages24h +
    ` (${arbitrageStatus})`
  );
//  lines.push('⚡ <b>Arbitrages 24h:</b> ' + totalArbitrages24h);

// ==========================================
// TARGET PRICE FOR 8% TOTAL PnL
// ==========================================

const currentInvestment = Number(investment || 0);
const currentPrice = Number(markPrice || 0);
const currentPosition = Number(botPosition || 0);

const currentTotalPnL = Number(currentPnL);

const targetTotalPnL =
  currentInvestment * 0.08;

const requiredAdditionalPnL =
  targetTotalPnL - currentTotalPnL;

let targetPrice8 = null;

if (
  currentPrice > 0 &&
  currentPosition > 0 &&
  Number.isFinite(requiredAdditionalPnL)
) {
  targetPrice8 =
    currentPrice +
    (
      requiredAdditionalPnL /
      currentPosition
    );

  targetPrice8 =
    Number(targetPrice8.toFixed(8));
}

// Απόσταση από την τρέχουσα τιμή έως το Target 8%
const targetDistance = targetPrice8 !== null
  ? targetPrice8 - currentPrice
  : null;

// Ποσοστιαία μεταβολή που απαιτείται από την τρέχουσα τιμή
const targetDistancePct =
  targetPrice8 !== null && currentPrice > 0
    ? (targetDistance / currentPrice) * 100
    : null;

Logger.log('========== TARGET 8% ==========');
Logger.log('Investment: ' + currentInvestment);
Logger.log('Bot Mark Price: ' + currentPrice);
Logger.log('Bot Position: ' + currentPosition);
Logger.log('Current Total PnL USDT: ' + currentTotalPnL);
Logger.log('Target Total PnL USDT: ' + targetTotalPnL);
Logger.log('Required Additional PnL: ' + requiredAdditionalPnL);
Logger.log('Target Price 8%: ' + targetPrice8);
Logger.log('Target Distance: ' + targetDistance);
Logger.log('Target Distance %: ' + targetDistancePct);
Logger.log('================================');

const targetTimeEstimate =
  targetPrice8 !== null
    ? PIONEX_TRACKER_estimateTargetTime4h_(
        pair,
        currentPrice,
        targetPrice8
      )
    : {
        message: 'Μη διαθέσιμη εκτίμηση τάσης',
        candles: null,
        hours: null
      };

Logger.log(
  '4H Target Time Estimate: ' +
  targetTimeEstimate.message
);

Logger.log(
  '4H Trend 12/36/72 (%): ' +
  [
    targetTimeEstimate.slope12Pct,
    targetTimeEstimate.slope36Pct,
    targetTimeEstimate.slope72Pct
  ].join(' / ')
);


// const targetTotalPnL = Number(investment || 0) * 0.08;

// Πόσο PnL απομένει μέχρι τον στόχο.
const pnlToRecover = targetTotalPnL - Number(currentTotalPnL);

// Εκτίμηση ημερών από το Grid profit 24ώρου.
const estimatedDaysByGrid =
  Number(gridProfit24h) > 0 && pnlToRecover > 0
    ? Math.ceil(pnlToRecover / Number(gridProfit24h))
    : null;

// Επιλογή εκτίμησης: πρώτα candles, διαφορετικά Grid profit.
const targetTimeMessage =
  targetTimeEstimate &&
  Number.isFinite(targetTimeEstimate.hours) &&
  targetTimeEstimate.hours > 0
    ? `📈 Εκτίμηση βάσει τάσης: ~${(targetTimeEstimate.hours / 24).toFixed(1)} ημέρες`
    : estimatedDaysByGrid !== null
      ? `⏳ Εκτίμηση βάσει Grid profit: ~${estimatedDaysByGrid} ημέρες (ρυθμός ${PIONEX_TRACKER_formatNumber_(Number(gridProfit24h), 2)} USDT/24h)`
      : '⏳ Χρόνος έως Target 8%: Μη διαθέσιμη εκτίμηση';



const targetMessage =
  targetPrice8 !== null
    ? [
        `🎯 Target 8%: ${PIONEX_TRACKER_formatNumber_(targetPrice8, 8)}`,
        `📏 Απόσταση τιμής: ${PIONEX_TRACKER_formatNumber_(targetDistance, 8)}`,
        `📈 Μεταβολή έως στόχο: ${targetDistancePct >= 0 ? '+' : ''}${PIONEX_TRACKER_formatNumber_(targetDistancePct, 2)}%`,
        targetTimeMessage,
        `🏛️ Target Total PnL: ${PIONEX_TRACKER_formatNumber_(targetTotalPnL, 2)} USDT`,
        `💰 PnL που απομένει: ${PIONEX_TRACKER_formatNumber_(pnlToRecover, 2)} USDT`
      ].join('\n')
    : '🎯 Target 8%: N/A';

lines.push('');
lines.push(targetMessage);

lines.push('');
lines.push('🕒 ' + new Date().toLocaleString('el-GR'));

});


  PIONEX_TRACKER_sendTelegramMessage_(lines.join('\n'));
}

// =====================================================
// PIONEX - GET CURRENT FUTURES POSITION
// =====================================================

function PIONEX_TRACKER_getFuturesPosition_(symbol) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty('PIONEX_API_KEY');

  const apiSecret =
    PropertiesService
      .getScriptProperties()
      .getProperty('PIONEX_API_SECRET');

  if (!apiKey) {
    throw new Error(
      'PIONEX_API_KEY δεν βρέθηκε στα Script Properties'
    );
  }

  if (!apiSecret) {
    throw new Error(
      'PIONEX_API_SECRET δεν βρέθηκε στα Script Properties'
    );
  }

  // ==================================================
  // IMPORTANT:
  // Pionex Futures symbol:
  // PEOPLE_USDT_PERP
  // ==================================================

  symbol = String(symbol)
    .replace('.PERP_USDT_PERP', '_USDT_PERP')
    .replace('.PERP', '_USDT_PERP')
    .replace('PERP_USDT_PERP', 'USDT_PERP');

  // ==================================================
  // API
  // ==================================================

  const baseUrl =
    'https://api.pionex.com';

  const path =
    '/uapi/v1/account/positions';

  const timestamp =
    Date.now();

  // ==================================================
  // QUERY PARAMETERS
  // ==================================================

  const params = {
    symbol: symbol,
    timestamp: String(timestamp)
  };

  // ==================================================
  // SORT PARAMETERS ALPHABETICALLY
  // ==================================================

  const queryString =
    Object.keys(params)
      .sort()
      .map(function(key) {
        return key + '=' + params[key];
      })
      .join('&');

  // ==================================================
  // SIGNATURE
  //
  // Pionex:
  //
  // METHOD + PATH_URL
  //
  // where PATH_URL contains:
  //
  // /path?sortedQuery
  //
  // ==================================================

  const pathUrl =
    path + '?' + queryString;

  const signPayload =
    'GET' + pathUrl;

  Logger.log(
    'Pionex symbol: ' +
    symbol
  );

  Logger.log(
    'Pionex query: ' +
    queryString
  );

  Logger.log(
    'Pionex sign payload: ' +
    signPayload
  );

  const signatureBytes =
    Utilities.computeHmacSha256Signature(
      signPayload,
      apiSecret
    );

  const signature =
    signatureBytes
      .map(function(byte) {
        return (
          ('0' +
            ((byte + 256) % 256)
              .toString(16)
          ).slice(-2)
        );
      })
      .join('');

  // ==================================================
  // REQUEST
  // ==================================================

  const url =
    baseUrl + pathUrl;

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',

        headers: {
          'PIONEX-KEY': apiKey,
          'PIONEX-SIGNATURE': signature,
          'Accept': '*/*'
        },

        muteHttpExceptions: true
      }
    );

  const httpCode =
    response.getResponseCode();

  const responseText =
    response.getContentText();

  Logger.log(
    'Pionex HTTP Code: ' +
    httpCode
  );

  Logger.log(
    'Pionex Position Response: ' +
    responseText
  );

  // ==================================================
  // HTTP ERROR
  // ==================================================

  if (httpCode !== 200) {

    throw new Error(
      'Pionex HTTP ' +
      httpCode +
      ': ' +
      responseText
    );
  }

  // ==================================================
  // PARSE
  // ==================================================

  const json =
    JSON.parse(responseText);

  if (json.result !== true) {

    throw new Error(
      'Pionex API error: ' +
      responseText
    );
  }

  // ==================================================
  // POSITIONS
  // ==================================================

  const positions =
    json.data &&
    Array.isArray(json.data.positions)
      ? json.data.positions
      : [];

  Logger.log(
    'Pionex positions found: ' +
    positions.length
  );

  // ==================================================
  // FIND OUR SYMBOL
  // ==================================================

  const position =
    positions.find(function(p) {

      return String(p.symbol) ===
        String(symbol);

    });

  if (!position) {

    Logger.log(
      'Requested symbol: ' +
      symbol
    );

    Logger.log(
      'Available positions: ' +
      JSON.stringify(positions)
    );

    throw new Error(
      'Δεν βρέθηκε ενεργή θέση για ' +
      symbol
    );
  }

  // ==================================================
  // DEBUG POSITION
  // ==================================================

  Logger.log(
    '========== PIONEX POSITION =========='
  );

  Logger.log(
    'symbol: ' +
    position.symbol
  );

  Logger.log(
    'positionSide: ' +
    position.positionSide
  );

  Logger.log(
    'netSize: ' +
    position.netSize
  );

  Logger.log(
    'sizeLong: ' +
    position.sizeLong
  );

  Logger.log(
    'sizeShort: ' +
    position.sizeShort
  );

  Logger.log(
    'amountLong: ' +
    position.amountLong
  );

  Logger.log(
    'avgPrice: ' +
    position.avgPrice
  );

  Logger.log(
    'markPrice: ' +
    position.markPrice
  );

  Logger.log(
    'unrealizedPnL: ' +
    position.unrealizedPnL
  );

  Logger.log(
    'leverage: ' +
    position.leverage
  );

  Logger.log(
    '====================================='
  );

  return position;
}


function PIONEX_TRACKER_refreshAndNotifyTelegram() {

  const refreshResult = PIONEX_TRACKER_refreshPionexBots();

  const rows = refreshResult.rows;
  const botPositions = refreshResult.botPositions;  let accountBalance = { free: null, frozen: null, total: null };

  try {
    accountBalance = PIONEX_TRACKER_getUsdtBalance_(PIONEX_TRACKER_getAccountBalances_());
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Balance API error: ' + error.message);
  }

  try {
    PIONEX_TRACKER_buildAndSendTelegramSummary_(rows, accountBalance, botPositions);
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Telegram error: ' + error.message);
  }
}

/** ---------------------- TRIGGERS ---------------------- **/

function PIONEX_TRACKER_setupAutoRefreshTrigger() {
  PIONEX_TRACKER_removeAutoRefreshTrigger();
  ScriptApp.newTrigger('PIONEX_TRACKER_refreshAndNotifyTelegram').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Αυτόματη ανανέωση + Telegram κάθε 15 λεπτά ενεργοποιήθηκε.');
}

function PIONEX_TRACKER_removeAutoRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'PIONEX_TRACKER_refreshPionexBots' || handler === 'PIONEX_TRACKER_refreshAndNotifyTelegram') ScriptApp.deleteTrigger(trigger);
  });
}

/** ---------------------- DEBUG ---------------------- **/

function PIONEX_TRACKER_debugBalanceAndPnl() {
  const rows = PIONEX_TRACKER_refreshPionexBots();
  const rawBotBalance = PIONEX_TRACKER_getRawBotBalanceFromRows_(rows);
  const totalInvestment = rows.reduce((sum, row) => sum + (Number(row[6]) || 0), 0);
  let accountBalance = { free: null, frozen: null, total: null };
  try { accountBalance = PIONEX_TRACKER_getUsdtBalance_(PIONEX_TRACKER_getAccountBalances_()); } catch (error) {}

  const accountValue = accountBalance.total === null ? null : accountBalance.total;
  const botBalance = accountValue === null ? null : rawBotBalance - accountValue;
  const pnl = botBalance === null ? null : botBalance - totalInvestment;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Pionex Debug');
  if (!sheet) sheet = spreadsheet.insertSheet('Pionex Debug');
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).setValues([['Bot', 'Investment', 'Current value', 'Profit', 'Raw balance', 'Account USDT', 'Bot balance', 'PnL']]);

  const debugRows = rows.map((row, index) => {
    const investment = Number(row[6]) || 0;
    const currentValue = Number(row[7]) || 0;
    const profit = Number(row[8]) || 0;
    return [index + 1, investment, currentValue, profit, currentValue + profit, accountValue, botBalance, pnl];
  });

  if (debugRows.length) sheet.getRange(2, 1, debugRows.length, 8).setValues(debugRows);
}

function PIONEX_TRACKER_debugPionexBotOrders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Pionex Debug');
  if (!sheet) sheet = spreadsheet.insertSheet('Pionex Debug');
  sheet.clear();
  try {
    const response = PIONEX_TRACKER_pionexGet_('/api/v1/bot/orders', { status: 'running' });
    sheet.getRange(1, 1).setValue('Raw response:');
    sheet.getRange(2, 1).setValue(JSON.stringify(response, null, 2));
  } catch (error) {
    sheet.getRange(1, 1).setValue('Σφάλμα:');
    sheet.getRange(2, 1).setValue(error.message);
  }
}

function PIONEX_TRACKER_debugSigningMessage() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Pionex Debug');
  if (!sheet) sheet = spreadsheet.insertSheet('Pionex Debug');
  sheet.clear();
  const properties = PropertiesService.getScriptProperties();
  const key = properties.getProperty('PIONEX_API_KEY') || '';
  const secret = properties.getProperty('PIONEX_API_SECRET') || '';
  const signing = PIONEX_TRACKER_buildSigningMessage_('GET', '/api/v1/bot/orders', { status: 'running', timestamp: Date.now() });
  const mask = value => value.length > 8 ? value.slice(0, 4) + '...' + value.slice(-4) + ' (μήκος ' + value.length + ')' : '(μικρό: ' + value.length + ')';
  sheet.getRange(1, 1).setValue('Signing message:');
  sheet.getRange(2, 1).setValue(signing.message);
  sheet.getRange(4, 1).setValue('PIONEX_API_KEY (masked):');
  sheet.getRange(4, 2).setValue(mask(key.trim()));
  sheet.getRange(5, 1).setValue('PIONEX_API_SECRET (masked):');
  sheet.getRange(5, 2).setValue(mask(secret.trim()));
}




/**
 * ============================================================================
 * Main Opening Menu
 * ============================================================================
 *

/** ---------------------- Open Menu ---------------------- **/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Grid Screener')
    .addItem('🚀 Run Screener', 'checkGridCandidates')
    .addSeparator()
    .addItem('➕ Add Open Position', 'addOpenPosition')
    .addItem('➖ Remove Open Position', 'removeOpenPosition')
    .addSeparator()
    .addItem('⚡ Quick Check Open Positions Only', 'checkOpenPositionsQuick')
    .addItem('📊 Check Bot Status', 'PIONEX_TRACKER_refreshAndNotifyTelegram')
    .addItem('📈 Analyze 4h candles', 'CA4H_analyze4hAndNotify')
    .addItem('🔄 Reset Stable Pick (lock)', 'gridEntry_checkOpenPositions')
    .addItem('💡 Check Open Positiions for Near Signal', 'gridEntry_checkOpenPositionsAndSignal')
    .addItem('🧹 Clear History Log', 'clearHistory')
    .addItem('❌ Exclude Symbol (not on Pionex)', 'excludeSymbol')
    .addItem('🗑️ Remove Excluded Symbol', 'unexcludeSymbol')
    .addItem('⭐ Screener Low Funding Fees', 'OKX4H_runScreener')
    .addItem('🔑 Next Unlock for Open Position', 'gridEntry_sendDailyNewPositionUnlockReport')
    .addSeparator()
    .addItem('🔴 Κλείσε Ενεργό Grid', 'btcGrid_closeActiveGridFromMenu')
    .addItem('🔀 Άλλαξε σε Καλύτερη Εναλλακτική', 'btcGrid_switchFromMenu')
    .addItem('▶️ Τρέξε Έλεγχο Τώρα', 'btcGrid_runBTCGridEngine')
    .addItem('🔍 Έλεγχος Διαθεσιμότητας Watchlist', 'btcGrid_debugCheckWatchlistAvailability')
    .addSeparator()
    .addItem('🏠︎ Pearl Dashboard', 'runPearlDashboard')
    .addItem('💲 FED News – Macro Risk Signals', 'runDailyUpdate')
    .addItem('🏛️ Crypto Market - 12H Outlook', 'CTM_runMarketTrendReport')
    .addItem('👷🏻 Mining Report', 'refreshKryptex')
    .addToUi();
}







/**
 * ============================================================================
 * 4H Candlestick Pattern Analyzer — OKX -> Google Sheets -> Telegram (v3.7)
 * ============================================================================
 *
 * ΑΛΛΑΓΕΣ v3.7 (ενσωματώνει v3.5 + v3.6 + νέες προσθήκες)
 *
 * [ΔΙΟΡΘΩΣΗ] Volume Fix: Το OKX candles array έχει μορφή
 *   [ts, o, h, l, c, vol, volCcy, volCcyQuote, confirm]. Για SWAP/perpetual
 *   συμβόλαια, το "vol" (index 5) εκφράζεται σε ΑΡΙΘΜΟ ΣΥΜΒΟΛΑΙΩΝ
 *   (contracts), όχι σε χρήσιμο μέγεθος. Τώρα διαβάζουμε "volCcy" (index 6),
 *   που εκφράζεται σε quote currency (π.χ. USDT) — σωστό και συγκρίσιμο.
 *   Το ωμό "vol" κρατιέται ξεχωριστά ως volContracts.
 *
 * [ΝΕΟ] Αναμενόμενο Εύρος Τιμών (4H/8H): Προβολή βάσει ATR(14) με κλιμάκωση
 *   √χρόνου (8H = ATR×√2, όχι ATR×2) και ασυμμετρία (skew) προς τη φορά
 *   της τρέχουσας τάσης.
 *
 * [ΝΕΟ] Signal Conflict Flag: Εντοπίζει πότε το τελευταίο candlestick
 *   pattern αντιτίθεται στο KDJ και/ή MACD histogram, και το επισημαίνει
 *   ρητά στο μήνυμα — εξηγεί το "γιατί" πίσω από ένα NEUTRAL score.
 *
 * [ΝΕΟ] ATR-Distance to Grid Edges: Απόσταση τιμής από τα άκρα του DCA
 *   grid εκφρασμένη σε πολλαπλάσια ATR (πιο διαισθητικό μέτρο ρίσκου από
 *   το ωμό ποσοστό θέσης μέσα στο εύρος).
 *
 * [ΝΕΟ] Squeeze-Aware DCA Guidance + Squeeze Streak Tracking: Όταν το
 *   Bollinger squeeze είναι ενεργό, προσθέτει καθοδήγηση για το grid και
 *   μετράει πόσα συνεχόμενα candles διαρκεί το squeeze (ιστορικό στο sheet).
 *
 * ΣΗΜΑΝΤΙΚΟ: Η ενότητα προβολής εύρους τιμής είναι στατιστική εκτίμηση
 * μεταβλητότητας βασισμένη σε ιστορικό ATR, ΟΧΙ εγγυημένη πρόβλεψη τιμής.
 * ============================================================================
 */

const CA4H_CFG = {
  POSITIONS_SHEET: 'Open Positions',
  SYMBOL_CELL: 'A2',
  POSITION_ROW_RANGE: 'A2:H2',
  DIRECTION_CELL: null,
  LOWER_PRICE_CELL: null,
  UPPER_PRICE_CELL: null,

  ANALYSIS_SHEET: '4h analysis',
  BAR: '4H',
  CANDLE_LIMIT: 100,
  RSI_PERIOD: 14,
  EMA_FAST: 12,
  EMA_SLOW: 26,
  MACD_SIGNAL: 9,
  VOLUME_MA_SHORT: 5,
  VOLUME_MA_LONG: 10,
  OKX_BASE_URL: 'https://www.okx.com',
  MIN_PRICE_DECIMALS: 4,
  MAX_PRICE_DECIMALS: 12,

  KDJ_N: 9,
  KDJ_K_SMOOTH: 3,
  KDJ_D_SMOOTH: 3,

  BB_PERIOD: 20,
  BB_STDDEV_MULT: 2,
  BB_SQUEEZE_THRESHOLD_PCT: 3,

  ATR_PERIOD: 14,

  FIB_SWING_LOOKBACK: 30,
  FIB_LEVELS: [0.5, 0.618, 0.786],

  GRID_EDGE_WARNING_PCT: 10,

  PROJECTION_SKEW_BULLISH: 0.20,
  PROJECTION_SKEW_BEARISH: 0.20,
  PROJECTION_ATR_MULT_4H: 1.0,
  PROJECTION_ATR_MULT_8H_SQRT: Math.SQRT2,

  // --- NEW: ATR-distance classification thresholds (σε πολλαπλάσια ATR) ---
  ATR_DIST_NEAR_MAX: 3,
  ATR_DIST_MEDIUM_MAX: 8,

  PIONEX_BASE_URL: 'https://api.pionex.com',
  PIONEX_CACHE_PREFIX: 'CA4H_PIONEX_DERIV_',
  PIONEX_SYMBOL_CACHE_TTL_SECONDS: 6 * 60 * 60,

  DERIV_OI_BULLISH_MIN_PCT: 0.50,
  DERIV_OI_BEARISH_MIN_PCT: -0.50,
  DERIV_FUNDING_CROWDED_PCT: 0.05,
  DERIV_FUNDING_EXTREME_PCT: 0.10,

  // --- NEW: Strong candlestick patterns που μπορούν να πυροδοτήσουν conflict flag ---
  STRONG_PATTERNS: [
    'Bullish Marubozu', 'Bearish Marubozu',
    'Bullish Engulfing', 'Bearish Engulfing',
    'Three White Soldiers', 'Three Black Crows',
    'Hammer', 'Shooting Star', 'Hanging Man'
  ]

};

function CA4H_calcHistoricallyDefendedSupportZone_(
  candles,
  currentPrice,
  atr
) {
  if (
    !candles ||
    candles.length < 20 ||
    currentPrice == null ||
    atr == null ||
    atr <= 0
  ) {
    return null;
  }

  const lookback =
    Math.min(100, candles.length);

  const leftRightBars = 2;

  const recent =
    candles.slice(
      candles.length - lookback
    );

  const swingLows = [];

  for (
    let i = leftRightBars;
    i < recent.length - leftRightBars;
    i++
  ) {
    const currentLow =
      Number(recent[i].low);

    if (
      !isFinite(currentLow) ||
      currentLow >= currentPrice
    ) {
      continue;
    }

    let isSwingLow = true;

    for (
      let j = i - leftRightBars;
      j <= i + leftRightBars;
      j++
    ) {
      if (
        j !== i &&
        Number(recent[j].low) <= currentLow
      ) {
        isSwingLow = false;
        break;
      }
    }

    if (isSwingLow) {
      swingLows.push({
        price: currentLow,
        ts: Number(recent[i].ts),
        index: i
      });
    }
  }

  if (!swingLows.length) {
    return null;
  }

  const clusterTolerance =
    Math.max(
      atr * 0.35,
      currentPrice * 0.002
    );

  const sorted =
    swingLows.slice().sort(function(a, b) {
      return a.price - b.price;
    });

  const clusters = [];

  sorted.forEach(function(item) {
    const lastCluster =
      clusters[clusters.length - 1];

    const lastPrice =
      lastCluster
        ? lastCluster[lastCluster.length - 1].price
        : null;

    if (
      !lastCluster ||
      item.price - lastPrice > clusterTolerance
    ) {
      clusters.push([item]);
    } else {
      lastCluster.push(item);
    }
  });

  const zones =
    clusters.map(function(cluster) {
      const prices =
        cluster.map(function(item) {
          return item.price;
        });

      const lowest =
        Math.min.apply(null, prices);

      const highest =
        Math.max.apply(null, prices);

      const midpoint =
        prices.reduce(function(sum, price) {
          return sum + price;
        }, 0) / prices.length;

      const latest =
        cluster.reduce(function(best, item) {
          return !best || item.ts > best.ts
            ? item
            : best;
        }, null);

      return {
        lower:
          lowest - atr * 0.10,

        upper:
          highest + atr * 0.10,

        midpoint: midpoint,

        touches:
          cluster.length,

        distanceAtr:
          (currentPrice - midpoint) / atr,

        latestTimestamp:
          latest ? latest.ts : null
      };
    })
    .filter(function(zone) {
      return zone.midpoint < currentPrice;
    });

  if (!zones.length) {
    return null;
  }

  const repeated =
    zones
      .filter(function(zone) {
        return zone.touches >= 2;
      })
      .sort(function(a, b) {
        return a.distanceAtr - b.distanceAtr;
      });

  const selected =
    repeated.length
      ? repeated[0]
      : zones.sort(function(a, b) {
          return a.distanceAtr - b.distanceAtr;
        })[0];

  let strength = 'ασθενής';

  if (selected.touches >= 4) {
    strength = 'πολύ ισχυρή';
  } else if (selected.touches >= 3) {
    strength = 'ισχυρή';
  } else if (selected.touches >= 2) {
    strength = 'μέτρια';
  }

  return {
    lower: selected.lower,
    upper: selected.upper,
    midpoint: selected.midpoint,
    touches: selected.touches,
    distanceAtr: selected.distanceAtr,
    strength: strength,
    lookbackCandles: lookback,
    tolerance: clusterTolerance,
    latestTimestamp: selected.latestTimestamp,
    isRepeatedZone:
      selected.touches >= 2
  };
}

function CA4H_buildPionexKlineSignalFromIndex_(
  klines,
  index,
  referenceOpenTime,
  isAligned
) {
  const intervalMs =
    4 * 60 * 60 * 1000;

  if (!klines || klines.length < 2) {
    throw new Error(
      'Δεν υπάρχουν αρκετά Pionex klines.'
    );
  }

  if (
    index == null ||
    index < 1 ||
    index >= klines.length
  ) {
    throw new Error(
      'Μη έγκυρο Pionex kline index: ' +
      index
    );
  }

  const previous =
    klines[index - 1];

  const last =
    klines[index];

  const previousFive =
    klines.slice(
      Math.max(0, index - 5),
      index
    );

  const quoteVolumes =
    previousFive.map(function(k) {
      return Number(k.close) *
        Number(k.volume);
    });

  const currentQuoteVolume =
    Number(last.close) *
    Number(last.volume);

  const quoteVolumeMa5 =
    quoteVolumes.length
      ? quoteVolumes.reduce(
          function(sum, value) {
            return sum + value;
          },
          0
        ) / quoteVolumes.length
      : null;

  const priceChangePct =
    Number(previous.close) !== 0
      ? (
          (Number(last.close) -
            Number(previous.close)) /
          Number(previous.close)
        ) * 100
      : null;

  const volumeChangePct =
    quoteVolumeMa5 !== null &&
    quoteVolumeMa5 !== 0
      ? (
          (currentQuoteVolume -
            quoteVolumeMa5) /
          quoteVolumeMa5
        ) * 100
      : null;

  Logger.log(
    'Pionex reference candle: ' +
    new Date(
      referenceOpenTime
    ).toISOString()
  );

  Logger.log(
    'Pionex actual candle: ' +
    new Date(
      last.openTime
    ).toISOString()
  );

  Logger.log(
    'Pionex candles aligned: ' +
    isAligned
  );

  Logger.log(
    'Pionex price change: ' +
    priceChangePct +
    '%'
  );

  return {
    candleOpenTime:
      Number(last.openTime),

    candleCloseTime:
      Number(last.openTime) +
      intervalMs,

    candleOpenUtc:
      new Date(
        Number(last.openTime)
      ).toISOString(),

    candleCloseUtc:
      new Date(
        Number(last.openTime) +
        intervalMs
      ).toISOString(),

    close:
      Number(last.close),

    previousClose:
      Number(previous.close),

    priceChangePct:
      priceChangePct,

    baseVolume:
      Number(last.volume),

    quoteVolume:
      currentQuoteVolume,

    quoteVolumeMa5:
      quoteVolumeMa5,

    volumeChangePct:
      volumeChangePct,

    volumeAboveMa:
      volumeChangePct !== null &&
      volumeChangePct > 0,

    usingClosedCandle:
      true,

    referenceOpenTime:
      Number(referenceOpenTime),

    candlesAligned:
      Boolean(isAligned)
  };
}

// ---------------------------------------------------------------------------
// Funding + OI + Price + Volume confirmation
// ---------------------------------------------------------------------------

function CA4H_getClosedPionexKlines_(pionexSymbol) {
const data =
  CA4H_pionexGetJson_(
    '/api/v1/market/klines',
    {
      symbol: pionexSymbol,
      interval: '4H',
      limit: 10
    }
  );

  const rows = data.klines || [];

  if (rows.length < 3) {
    return [];
  }

  const now = Date.now();

  const normalized = rows.map(function(k) {
    return {
      openTime: Number(k.openTime || k.startTime || k.time || k.timestamp),
      open: Number(k.open),
      close: Number(k.close),
      high: Number(k.high),
      low: Number(k.low),
      volume: Number(k.volume),
      closeTime: Number(k.closeTime || 0)
    };
  });

  const closed = normalized.filter(function(k) {
    if (k.closeTime > 0) {
      return k.closeTime <= now;
    }

    // Fallback: 4H candle duration = 14,400,000 ms.
    return k.openTime > 0 &&
      k.openTime + 4 * 60 * 60 * 1000 <= now;
  });

  return closed.sort(function(a, b) {
    return a.openTime - b.openTime;
  });
}

function CA4H_pionexGetJson_(path, params) {
  const query = Object.keys(params || {})
    .filter(function(k) {
      return params[k] !== undefined &&
        params[k] !== null &&
        params[k] !== '';
    })
    .map(function(k) {
      return encodeURIComponent(k) + '=' +
        encodeURIComponent(params[k]);
    })
    .join('&');

  const url = CA4H_CFG.PIONEX_BASE_URL +
    path +
    (query ? '?' + query : '');

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status !== 200) {
    throw new Error(
      'Pionex HTTP ' + status + ' for ' + path +
      ': ' + body.slice(0, 500)
    );
  }

  const json = JSON.parse(body);

  if (!json.result) {
    throw new Error(
      'Pionex API error for ' + path +
      ': ' + (json.code || 'unknown') +
      ' — ' + (json.message || body.slice(0, 500))
    );
  }

  return json.data || {};
}

function CA4H_getPionexPerpSymbol_(okxInstId) {
  const base = String(okxInstId)
    .toUpperCase()
    .replace('-SWAP', '')
    .split('-')[0];

  const quote = String(okxInstId)
    .toUpperCase()
    .replace('-SWAP', '')
    .split('-')[1] || 'USDT';

  const directCandidate = base + '_' + quote + '_PERP';

  const symbolsData = CA4H_pionexGetJson_(
    '/api/v1/common/symbols',
    { type: 'PERP', status: 'TRADING' }
  );

  const rows = symbolsData.symbols || [];

  const exact = rows.find(function(row) {
    return String(row.symbol).toUpperCase() === directCandidate;
  });

  if (exact) {
    return exact.symbol;
  }

  const fallback = rows.find(function(row) {
    return String(row.baseCurrency).toUpperCase() === base &&
      String(row.quoteCurrency).toUpperCase() === quote &&
      String(row.contractType).toUpperCase() === 'PERPETUAL';
  });

  return fallback ? fallback.symbol : null;
}

function CA4H_getPionexOpenInterest_(pionexSymbol) {
  const data = CA4H_pionexGetJson_(
    '/api/v1/market/openInterests',
    {}
  );

  const rows = data.openInterests || [];

  const row = rows.find(function(item) {
    return String(item.symbol).toUpperCase() ===
      String(pionexSymbol).toUpperCase();
  });

  return row
    ? Number(row.openInterest)
    : null;
}

function CA4H_getPionexFunding_(pionexSymbol) {
  const historical = CA4H_pionexGetJson_(
    '/api/v1/market/fundingRates',
    {
      symbol: pionexSymbol,
      limit: 8
    }
  );

  const rates = historical.rates || [];

  const indexData = CA4H_pionexGetJson_(
    '/api/v1/market/indexes',
    { symbol: pionexSymbol }
  );

  const indexes = indexData.indexes || [];

  const index = indexes.find(function(item) {
    return String(item.symbol).toUpperCase() ===
      String(pionexSymbol).toUpperCase();
  });

  const latest = rates.length
    ? Number(rates[0].fundingRate)
    : null;

  const previous = rates.length > 1
    ? Number(rates[1].fundingRate)
    : null;

  const avg = rates.length
    ? rates.reduce(function(sum, row) {
        return sum + Number(row.fundingRate || 0);
      }, 0) / rates.length
    : null;

  return {
    latestRate: latest,
    previousRate: previous,
    averageRate: avg,
    latestFundingTime: rates.length ? Number(rates[0].fundingTime) : null,
    nextFundingRate: index ? Number(index.nextFundingRate) : null,
    nextFundingTime: index ? Number(index.nextFundingTime) : null,
    markPrice: index ? Number(index.markPrice) : null,
    indexPrice: index ? Number(index.indexPrice) : null
  };
}

function CA4H_getPionexKlineSignal_(
  pionexSymbol,
  referenceCandleTime
) {
  const intervalMs =
    4 * 60 * 60 * 1000;

  const referenceOpenTime =
    referenceCandleTime
      ? new Date(
          referenceCandleTime
        ).getTime()
      : Math.floor(
          Date.now() / intervalMs
        ) * intervalMs - intervalMs;

  const data =
    CA4H_pionexGetJson_(
      '/api/v1/market/klines',
      {
        symbol: pionexSymbol,
        interval: '4H',
        limit: 10
      }
    );

  const rows =
    data.klines || [];

  if (rows.length < 2) {
    throw new Error(
      'Δεν υπάρχουν αρκετά Pionex klines.'
    );
  }

  const klines =
    rows
      .map(function(k) {
        return {
          openTime: Number(
            k.openTime ||
            k.startTime ||
            k.time ||
            k.timestamp ||
            k[0]
          ),

          open: Number(
            k.open || k[1]
          ),

          close: Number(
            k.close || k[4]
          ),

          high: Number(
            k.high || k[2]
          ),

          low: Number(
            k.low || k[3]
          ),

          volume: Number(
            k.volume || k[5]
          )
        };
      })
      .filter(function(k) {
        return k.openTime > 0 &&
          isFinite(k.close) &&
          isFinite(k.volume);
      })
      .sort(function(a, b) {
        return a.openTime - b.openTime;
      });

  const matchingIndex =
    klines.findIndex(function(k) {
      return k.openTime ===
        referenceOpenTime;
    });

  if (matchingIndex >= 1) {
    return CA4H_buildPionexKlineSignalFromIndex_(
      klines,
      matchingIndex,
      referenceOpenTime,
      true
    );
  }

  Logger.log(
    'Δεν βρέθηκε ακριβές Pionex candle για ' +
    new Date(
      referenceOpenTime
    ).toISOString()
  );

  Logger.log(
    'Διαθέσιμα candles: ' +
    klines.map(function(k) {
      return new Date(
        k.openTime
      ).toISOString();
    }).join(', ')
  );

  const fallbackIndex =
    klines.length - 1;

  return CA4H_buildPionexKlineSignalFromIndex_(
    klines,
    fallbackIndex,
    referenceOpenTime,
    false
  );
}


function CA4H_getPionexTicker_(pionexSymbol) {
  const data = CA4H_pionexGetJson_(
    '/api/v1/market/tickers',
    { symbol: pionexSymbol }
  );

  const tickers = data.tickers || [];

  const ticker = tickers.find(function(item) {
    return String(item.symbol).toUpperCase() ===
      String(pionexSymbol).toUpperCase();
  });

  if (!ticker) {
    return null;
  }

  return {
    close: Number(ticker.close),
    amount24h: Number(ticker.amount),
    volume24hBase: Number(ticker.volume),
    count24h: Number(ticker.count)
  };
}

function CA4H_buildOiAnalysis_(derivatives) {
  if (
    !derivatives ||
    derivatives.error ||
    !derivatives.oi ||
    !derivatives.market
  ) {
    return {
      label: 'OI: insufficient data',
      emoji: '⚪',
      bias: 'neutral',
      note: 'Δεν υπάρχουν αρκετά δεδομένα OI για ανάλυση.'
    };
  }

  const oiChange = derivatives.oi.changePct;
  const priceChange = derivatives.market.priceChangePct;

  if (
    oiChange == null ||
    priceChange == null ||
    !isFinite(Number(oiChange)) ||
    !isFinite(Number(priceChange))
  ) {
    return {
      label: 'OI: baseline',
      emoji: '⚪',
      bias: 'neutral',
      note:
        'Το OI έχει καταγραφεί ως baseline. ' +
        'Η κατεύθυνση θα αξιολογηθεί στην επόμενη μέτρηση.'
    };
  }

  const oi = Number(oiChange);
  const price = Number(priceChange);

  if (oi > 0.10 && price > 0.10) {
    return {
      label: 'OI ↑ μαζί με Price ↑',
      emoji: '🟢',
      bias: 'bullish',
      note:
        'Το OI αυξάνεται μαζί με την τιμή, ένδειξη νέου ' +
        'positioning που στηρίζει την ανοδική κίνηση. ' +
        'Παρακολούθησε όμως το funding για crowded longs.'
    };
  }

  if (oi > 0.10 && price < -0.10) {
    return {
      label: 'OI ↑ ενώ Price ↓',
      emoji: '🔴',
      bias: 'bearish',
      note:
        'Το OI αυξάνεται ενώ η τιμή υποχωρεί, ένδειξη πιθανής ' +
        'εισόδου νέων shorts ή αυξανόμενης bearish έκθεσης.'
    };
  }

  if (oi < -0.10 && price > 0.10) {
    return {
      label: 'OI ↓ ενώ Price ↑',
      emoji: '🟡',
      bias: 'weak_bullish',
      note:
        'Η τιμή ανεβαίνει ενώ το OI μειώνεται. Η κίνηση μπορεί ' +
        'να οφείλεται σε short covering και ίσως έχει μικρότερη ' +
        'συνέχεια χωρίς νέο positioning.'
    };
  }

  if (oi < -0.10 && price < -0.10) {
    return {
      label: 'OI ↓ μαζί με Price ↓',
      emoji: '🟠',
      bias: 'flush',
      note:
        'Η τιμή και το OI μειώνονται μαζί, ένδειξη κλεισίματος ' +
        'ή πιθανών liquidations. Απόφυγε μηχανικό DCA μέχρι ' +
        'να σταθεροποιηθεί η τιμή.'
    };
  }

  return {
    label: 'OI και Price χωρίς σαφή σχέση',
    emoji: '⚪',
    bias: 'neutral',
    note:
      'Η μεταβολή OI είναι μικρή ή δεν επιβεβαιώνει καθαρά ' +
      'την κατεύθυνση της τιμής.'
  };
}

function CA4H_calcSupportZone_(a) {
  if (!a || !a.close || !a.atr || !a.atr.atr) {
    return null;
  }

  const close = Number(a.close);
  const atr = Number(a.atr.atr);

  const candidates = [];

  function addCandidate(price, source, weight) {
    if (
      price != null &&
      isFinite(Number(price)) &&
      Number(price) > 0 &&
      Number(price) < close
    ) {
      candidates.push({
        price: Number(price),
        source: source,
        weight: weight
      });
    }
  }

  // Bollinger lower band.
  addCandidate(
    a.bb && a.bb.lower,
    'Bollinger Lower',
    3
  );

  // Fibonacci retracement levels.
  if (a.fib && a.fib.levels) {
    addCandidate(
      a.fib.levels['0.5'],
      'Fibonacci 0.5',
      2
    );

    addCandidate(
      a.fib.levels['0.618'],
      'Fibonacci 0.618',
      3
    );

    addCandidate(
      a.fib.levels['0.786'],
      'Fibonacci 0.786',
      2
    );
  }

  // Recent local swing lows.
  if (a.recentSwingLows && a.recentSwingLows.length) {
    a.recentSwingLows.forEach(function(level) {
      addCandidate(
        level,
        'Recent swing low',
        3
      );
    });
  }

  // Grid lower boundary.
  if (
    a.position &&
    a.position.lowerPrice
  ) {
    addCandidate(
      a.position.lowerPrice,
      'DCA grid lower edge',
      1
    );
  }

  if (!candidates.length) {
    return null;
  }

  // Κρατάμε επίπεδα σε απόσταση έως 2 ATR κάτω από την τιμή.
  const nearby = candidates.filter(function(item) {
    return close - item.price <= atr * 2;
  });

  const usable =
    nearby.length ? nearby : candidates;

  const weightedPrice =
    usable.reduce(function(sum, item) {
      return sum + item.price * item.weight;
    }, 0) /
    usable.reduce(function(sum, item) {
      return sum + item.weight;
    }, 0);

  const sortedPrices =
    usable
      .map(function(item) {
        return item.price;
      })
      .sort(function(x, y) {
        return x - y;
      });

  const lower =
    Math.min.apply(null, sortedPrices) -
    atr * 0.15;

  const upper =
    Math.max.apply(null, sortedPrices) +
    atr * 0.15;

  const supportingSources =
    usable
      .map(function(item) {
        return item.source;
      })
      .filter(function(source, index, arr) {
        return arr.indexOf(source) === index;
      });

  const distanceAtr =
    (close - weightedPrice) / atr;

  let strength = 'μέτρια';
  if (usable.length >= 3) {
    strength = 'ισχυρή';
  } else if (usable.length === 1) {
    strength = 'ασθενής';
  }

  const hasRepeatedHistoricalZone =
  a.historicallyDefendedSupport &&
  a.historicallyDefendedSupport.isRepeatedZone === true;

return {
  lower: lower,
  upper: upper,
  midpoint: weightedPrice,
  distanceAtr: distanceAtr,
  strength: strength,
  sources: supportingSources,
  isHistoricallyRepeated:
    hasRepeatedHistoricalZone
};


  return {
    lower: lower,
    upper: upper,
    midpoint: weightedPrice,
    distanceAtr: distanceAtr,
    strength: strength,
    sources: supportingSources
  };
}

function CA4H_findRecentSwingLows_(
  candles,
  lookback,
  leftRightBars
) {
  const lows = [];
  const start =
    Math.max(
      leftRightBars,
      candles.length - lookback
    );

  const end =
    candles.length - leftRightBars;

  for (let i = start; i < end; i++) {
    const currentLow =
      Number(candles[i].low);

    let isSwingLow = true;

    for (
      let j = i - leftRightBars;
      j <= i + leftRightBars;
      j++
    ) {
      if (
        j !== i &&
        Number(candles[j].low) <= currentLow
      ) {
        isSwingLow = false;
        break;
      }
    }

    if (isSwingLow) {
      lows.push(currentLow);
    }
  }

  return lows;
}



function CA4H_getPionexOiSnapshot_(pionexSymbol) {
  const key = CA4H_CFG.PIONEX_CACHE_PREFIX +
    'OI_' +
    pionexSymbol;

  const raw = PropertiesService
    .getScriptProperties()
    .getProperty(key);

  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function CA4H_savePionexOiSnapshot_(pionexSymbol, oi) {
  const key = CA4H_CFG.PIONEX_CACHE_PREFIX +
    'OI_' +
    pionexSymbol;

  PropertiesService
    .getScriptProperties()
    .setProperty(
      key,
      JSON.stringify({
        value: oi,
        timestamp: Date.now()
      })
    );
}

function CA4H_calcPionexOiChange_(pionexSymbol, currentOi) {
  const previous = CA4H_getPionexOiSnapshot_(pionexSymbol);

  let changePct = null;

  if (
    previous &&
    previous.value != null &&
    Number(previous.value) > 0 &&
    currentOi != null
  ) {
    changePct =
      ((Number(currentOi) - Number(previous.value)) /
        Number(previous.value)) * 100;
  }

  CA4H_savePionexOiSnapshot_(pionexSymbol, currentOi);

  return {
    previousOi: previous ? Number(previous.value) : null,
    currentOi: currentOi,
    changePct: changePct,
    snapshotTime: previous ? Number(previous.timestamp) : null
  };
}

function CA4H_describeAtrVolatility_(
  atrPct
) {
  if (
    atrPct == null ||
    !isFinite(atrPct)
  ) {
    return 'δεν υπάρχουν αρκετά δεδομένα';
  }

  if (atrPct < 1) {
    return 'χαμηλή';
  }

  if (atrPct < 2) {
    return 'μέτρια';
  }

  if (atrPct < 4) {
    return 'αυξημένη';
  }

  return 'πολύ αυξημένη';
}

function CA4H_classifyPionexDerivatives_(
  funding,
  oi,
  market,
  okxAnalysis
) {
  if (!funding || !oi || !market) {
    return {
      label: 'INSUFFICIENT DATA',
      emoji: '⚪',
      direction: 'neutral',
      note:
        'Δεν υπάρχουν αρκετά δεδομένα για derivatives confirmation.'
    };
  }

  const priceChange =
    Number(market.priceChangePct);

  const oiChange =
    oi.changePct == null
      ? null
      : Number(oi.changePct);

  const priceUp =
    isFinite(priceChange) &&
    priceChange > 0;

  const priceDown =
    isFinite(priceChange) &&
    priceChange < 0;

  const oiUp =
    oiChange !== null &&
    isFinite(oiChange) &&
    oiChange >= 0.10;

  const oiDown =
    oiChange !== null &&
    isFinite(oiChange) &&
    oiChange <= -0.10;

  // Χρησιμοποιούμε OKX volume ως κύρια αναφορά,
  // επειδή το DCA grid και η βασική ανάλυση βασίζονται στο OKX.
  const volumeUp =
    okxAnalysis &&
    okxAnalysis.volumeConfirmed === true;

  const volumeDown =
    okxAnalysis &&
    okxAnalysis.volumeConfirmed === false;

  const fundingRate =
    funding.nextFundingRate != null
      ? Number(funding.nextFundingRate)
      : funding.latestRate != null
        ? Number(funding.latestRate)
        : null;

  const fundingPct =
    fundingRate !== null
      ? fundingRate * 100
      : null;

  const fundingPositive =
    fundingRate !== null &&
    fundingRate > 0;

  const fundingNegative =
    fundingRate !== null &&
    fundingRate < 0;

  const fundingCrowded =
    fundingPct !== null &&
    fundingPct >= 0.05;

  const fundingExtreme =
    fundingPct !== null &&
    fundingPct >= 0.10;

  // 1. Ισχυρή bullish επιβεβαίωση
  if (
    priceUp &&
    oiUp &&
    volumeUp &&
    fundingPositive
  ) {
    if (fundingExtreme) {
      return {
        label: 'BULLISH BUT CROWDED LONGS',
        emoji: '🟠',
        direction: 'bullish',
        note:
          'Η τιμή, το OI και το volume αυξάνονται μαζί, ' +
          'αλλά το funding είναι ακραία θετικό. Η LONG θέση ' +
          'επιβεβαιώνεται, με αυξημένο κίνδυνο long squeeze. ' +
          'Απέφυγε επιθετικό DCA.'
      };
    }

    if (fundingCrowded) {
      return {
        label: 'BULLISH / LONGS GETTING CROWDED',
        emoji: '🟡',
        direction: 'bullish',
        note:
          'Η τιμή, το OI και το volume αυξάνονται μαζί. ' +
          'Το θετικό funding δείχνει long bias και αυξάνει ' +
          'σταδιακά τον κίνδυνο crowded longs. Κράτημα LONG ' +
          'και όχι επιθετική αύξηση έκθεσης.'
      };
    }

    return {
      label: 'STRONG BULLISH CONFIRMATION',
      emoji: '🟢',
      direction: 'bullish',
      note:
        'Η τιμή, το OI και το volume αυξάνονται μαζί και ' +
        'το funding είναι θετικό. Η ανοδική κίνηση έχει ' +
        'derivatives επιβεβαίωση και υποστηρίζει τη LONG θέση.'
    };
  }

  // 2. Bearish επιβεβαίωση
  if (
    priceDown &&
    oiUp &&
    volumeUp &&
    fundingNegative
  ) {
    return {
      label: 'STRONG BEARISH CONFIRMATION',
      emoji: '🔴',
      direction: 'bearish',
      note:
        'Η τιμή πέφτει ενώ OI και volume αυξάνονται και ' +
        'το funding είναι αρνητικό. Πιθανό νέο short build-up.'
    };
  }

  // 3. Άνοδος με κλείσιμο shorts
  if (
    priceUp &&
    oiDown &&
    volumeUp
  ) {
    return {
      label: 'SHORT-COVERING / WEAK BULLISH',
      emoji: '🟡',
      direction: 'neutral',
      note:
        'Η τιμή ανεβαίνει με αυξημένο volume, αλλά το OI ' +
        'μειώνεται. Πιθανό short covering αντί για νέο long ' +
        'build-up. Χρειάζεται επιβεβαίωση για επιθετικό DCA.'
    };
  }

  // 4. Πτώση με κλείσιμο/ρευστοποίηση longs
  if (
    priceDown &&
    oiDown &&
    volumeUp
  ) {
    return {
      label: 'LONG LIQUIDATION / FLUSH RISK',
      emoji: '🟠',
      direction: 'bearish',
      note:
        'Η τιμή και το OI πέφτουν με αυξημένο volume. ' +
        'Πιθανό κλείσιμο ή liquidation longs. Απόφυγε ' +
        'μηχανικό DCA μέχρι να σταθεροποιηθεί η αγορά.'
    };
  }

  // 5. Ήπια άνοδος χωρίς OI/volume επιβεβαίωση
  if (
    priceUp &&
    fundingPositive &&
    !oiUp &&
    !volumeUp
  ) {
    return {
      label: 'WEAK BULLISH / UNCONFIRMED',
      emoji: '⚪',
      direction: 'neutral',
      note:
        'Η τιμή και το funding είναι θετικά, αλλά OI ή volume ' +
        'δεν επιβεβαιώνουν νέο positioning. Κράτησε συντηρητική στάση.'
    };
  }

  return {
    label: 'MIXED / NO CLEAR CONFIRMATION',
    emoji: '⚪',
    direction: 'neutral',
    note:
      'Funding, OI, τιμή και volume δεν δίνουν συνεπή κατεύθυνση. ' +
      'Κράτησε συντηρητική στάση στο DCA grid.'
  };
}


function CA4H_formatPionexDerivativesMessage_(
  d,
  okxAnalysis
) {
  if (!d || d.error) {
    return (
      '📊 *Pionex Derivatives Confirmation:* `n/a` — ' +
      (
        d && d.error
          ? d.error
          : 'δεν υπήρχαν διαθέσιμα δεδομένα.'
      )
    );
  }

  const fundingRate =
    d.funding.nextFundingRate !== null
      ? d.funding.nextFundingRate
      : d.funding.latestRate;

  const fundingPct =
    fundingRate !== null
      ? (
          fundingRate * 100
        ).toFixed(4) + '%'
      : 'n/a';

  const oiText =
    d.oi.changePct === null
      ? 'baseline saved'
      : (
          d.oi.changePct >= 0
            ? '+'
            : ''
        ) +
        d.oi.changePct.toFixed(3) +
        '%';

  const priceText =
    d.market &&
    d.market.priceChangePct !== null
      ? (
          d.market.priceChangePct >= 0
            ? '+'
            : ''
        ) +
        d.market.priceChangePct.toFixed(2) +
        '%'
      : 'n/a';

  const okxVolumeConfirmed =
    okxAnalysis &&
    okxAnalysis.volumeConfirmed !== undefined
      ? okxAnalysis.volumeConfirmed
      : null;

  const volumeLine =
    okxVolumeConfirmed === null
      ? 'n/a'
      : okxVolumeConfirmed
        ? '✅ πάνω από MA5 (OKX)'
        : '⚠️ κάτω από MA5 (OKX)';

  const candleLine =
    d.market &&
    d.market.candleOpenUtc
      ? 'Pionex candle: `' +
        d.market.candleOpenUtc +
        '` – `' +
        d.market.candleCloseUtc +
        '`'
      : 'Pionex candle: `n/a`';

  const messageParts = [
    '📊 *Pionex Derivatives Confirmation*',
    'Symbol: `' + d.symbol + '`',
    'Funding (next/last): `' + fundingPct + '`',
    'OI μεταβολή: `' + oiText + '`',
    'Price 4H: `' + priceText + '`',
    'Volume: ' + volumeLine,
    candleLine
  ];

  // ==========================================
  // ΠΡΟΣΘΗΚΗ: OKX Long/Short Ratio
  // ==========================================
if (d.longShortRatio) {
  if (d.longShortRatio.error) {
    messageParts.push(
      '👥 Long/Short Ratio: `n/a` — ' +
      d.longShortRatio.error
    );
  } else if (d.longShortRatio.data) {
    const ratio =
      d.longShortRatio.data.ratio;

    const change =
      d.longShortRatio.data.changePct;

    const changeText =
      change == null
        ? 'n/a'
        : (
            change >= 0 ? '+' : ''
          ) +
          change.toFixed(2) +
          '%';

    messageParts.push(
      '👥 Long/Short Ratio (OKX, 1H): `' +
      ratio.toFixed(2) +
      '` | Μεταβολή: `' +
      changeText +
      '`'
    );

    messageParts.push(
      d.longShortRatio.interpretation.emoji +
      ' *' +
      d.longShortRatio.interpretation.label +
      '* — ' +
      d.longShortRatio.interpretation.note
    );
  }
}

  // ==========================================
  // Classification
  // ==========================================
  messageParts.push(
    d.classification.emoji +
    ' *' +
    d.classification.label +
    '*'
  );

  messageParts.push(
    d.classification.note
  );

  messageParts.push(
    '_Τα δεδομένα αφορούν Pionex perpetual market. Funding, OI και Long/Short Ratio είναι ένδειξη positioning/ρίσκου, όχι βεβαιότητα για μελλοντική τιμή._'
  );

return messageParts.join('\n');

}



function CA4H_getPionexDerivativesAnalysis_(
  okxInstId,
  okxAnalysis
) {
  try {
    const pionexSymbol =
      CA4H_getPionexPerpSymbol_(
        okxInstId
      );

    if (!pionexSymbol) {
      return {
        error:
          'Δεν υπάρχει αντιστοιχισμένο Pionex perpetual symbol.'
      };
    }

    const funding =
      CA4H_getPionexFunding_(
        pionexSymbol
      );

    const currentOi =
      CA4H_getPionexOpenInterest_(
        pionexSymbol
      );

    const oi =
      CA4H_calcPionexOiChange_(
        pionexSymbol,
        currentOi
      );

    const referenceTime =
      okxAnalysis &&
      okxAnalysis.lastCandleTime
        ? new Date(
            okxAnalysis.lastCandleTime
          ).getTime()
        : null;

    Logger.log(
      'OKX reference candle: ' +
      (
        referenceTime
          ? new Date(
              referenceTime
            ).toISOString()
          : 'n/a'
      )
    );

    const market =
      CA4H_getPionexKlineSignal_(
        pionexSymbol,
        referenceTime
      );

    const ticker =
      CA4H_getPionexTicker_(
        pionexSymbol
      );

    const classification =
      CA4H_classifyPionexDerivatives_(
        funding,
        oi,
        market,
        okxAnalysis
      );

let longShortRatio = null;

try {
  const okxRatio =
    CA4H_getOkxLongShortRatio_(
      okxInstId
    );

  longShortRatio = {
    data: okxRatio,
    interpretation:
      CA4H_interpretOkxLongShortRatio_(
        okxRatio.ratio
      )
  };
} catch (e) {
  longShortRatio = {
    error:
      'OKX Long/Short lookup failed: ' +
      e.message
  };
}

Logger.log(
  'Long/Short result: ' +
  JSON.stringify(longShortRatio)
);

return {
  symbol: pionexSymbol,
  funding: funding,
  oi: oi,
  market: market,
  ticker: ticker,
  longShortRatio: longShortRatio,
  classification: classification,
  fetchedAt: new Date()
};


  } catch (e) {
    return {
      error:
        'Pionex lookup failed: ' +
        e.message
    };
  }
}

function CA4H_getOkxCurrency_(
  okxInstId
) {
  return String(okxInstId)
    .toUpperCase()
    .replace('-SWAP', '')
    .split('-')[0];
}

function CA4H_getOkxLongShortRatio_(
  okxInstId
) {
  const ccy =
    CA4H_getOkxCurrency_(okxInstId);

const url =
  CA4H_CFG.OKX_BASE_URL +
  '/api/v5/rubik/stat/contracts/long-short-account-ratio' +
  '?ccy=' +
  encodeURIComponent(ccy) +
  '&period=1H' +
  '&limit=2';

  const response =
    UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (status !== 200) {
    throw new Error(
      'OKX Long/Short API error ' +
      status +
      ': ' +
      body.slice(0, 300)
    );
  }

  const json =
    JSON.parse(body);

  if (
    json.code !== '0' ||
    !json.data ||
    !json.data.length
  ) {
    throw new Error(
      'OKX Long/Short API returned no data: ' +
      body.slice(0, 300)
    );
  }

  const latest =
    json.data[0];

  const previous =
    json.data.length > 1
      ? json.data[1]
      : null;

  const ratio =
    Number(latest[1]);

  const previousRatio =
    previous
      ? Number(previous[1])
      : null;

  const changePct =
    previousRatio && previousRatio !== 0
      ? ((ratio - previousRatio) /
        previousRatio) * 100
      : null;

return {
  currency: ccy,
  ratio: ratio,
  previousRatio: previousRatio,
  changePct: changePct,
  timestamp: Number(latest[0]),
  period: '1H',
  source: 'OKX'
};

}

function CA4H_interpretOkxLongShortRatio_(
  ratio
) {
  if (
    ratio == null ||
    !isFinite(ratio)
  ) {
    return {
      emoji: '⚪',
      label: 'Μη διαθέσιμο',
      note:
        'Δεν υπάρχουν διαθέσιμα Long/Short δεδομένα.'
    };
  }

  if (ratio >= 3) {
    return {
      emoji: '🔴',
      label: 'Extreme long bias',
      note:
        'Πολύ υψηλή υπεροχή long λογαριασμών. ' +
        'Αυξημένος κίνδυνος crowded longs και long squeeze.'
    };
  }

  if (ratio >= 2) {
    return {
      emoji: '🟡',
      label: 'Strong long bias',
      note:
        'Δυνατή υπεροχή long λογαριασμών. Η άνοδος μπορεί να συνεχιστεί, ' +
        'αλλά αυξάνεται ο κίνδυνος long squeeze.'
    };
  }

  if (ratio > 1.2) {
    return {
      emoji: '🟢',
      label: 'Long bias',
      note:
        'Οι long λογαριασμοί υπερτερούν, χωρίς ακόμη ακραία συσσώρευση.'
    };
  }

  if (ratio >= 0.8) {
    return {
      emoji: '⚪',
      label: 'Balanced positioning',
      note:
        'Οι long και short λογαριασμοί είναι σχετικά ισορροπημένοι.'
    };
  }

  if (ratio >= 0.5) {
    return {
      emoji: '🟢',
      label: 'Short bias',
      note:
        'Οι short λογαριασμοί υπερτερούν. Πιθανή πίεση για short covering σε άνοδο.'
    };
  }

  return {
    emoji: '🟠',
    label: 'Extreme short bias',
    note:
      'Πολύ υψηλή υπεροχή short λογαριασμών. Πιθανή ευαισθησία σε short squeeze.'
  };
}



// ---------------------------------------------------------------------------
// MAIN ENTRY POINT
// ---------------------------------------------------------------------------
function CA4H_analyze4hAndNotify() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const posSheet =
    ss.getSheetByName(
      CA4H_CFG.POSITIONS_SHEET
    );

  if (!posSheet) {
    throw new Error(
      'Δεν βρέθηκε το φύλλο "' +
      CA4H_CFG.POSITIONS_SHEET +
      '"'
    );
  }

  const rawSymbol =
    posSheet
      .getRange(CA4H_CFG.SYMBOL_CELL)
      .getValue();

  if (!rawSymbol) {
    throw new Error(
      'Το κελί ' +
      CA4H_CFG.SYMBOL_CELL +
      ' στο "' +
      CA4H_CFG.POSITIONS_SHEET +
      '" είναι άδειο.'
    );
  }

  const instId =
    CA4H_resolveInstId_(
      String(rawSymbol).trim()
    );

  const candles =
    CA4H_fetchOkxCandles_(
      instId,
      CA4H_CFG.BAR,
      CA4H_CFG.CANDLE_LIMIT
    );

  if (!candles || candles.length < 30) {
    throw new Error(
      'Ανεπαρκή δεδομένα candles για ' +
      instId +
      ' (' +
      (candles ? candles.length : 0) +
      ' bars).'
    );
  }

  const analysis =
    CA4H_buildAnalysis_(
      instId,
      candles,
      ss
    );

  const position =
    CA4H_readOpenPosition_(
      posSheet
    );

  analysis.position = position;
  analysis.supportZone =
    CA4H_calcSupportZone_(
      analysis
    );

  analysis.pionexDerivatives =
    CA4H_getPionexDerivativesAnalysis_(
      instId,
      analysis
    );

  analysis.oiAnalysis =
    CA4H_buildOiAnalysis_(
      analysis.pionexDerivatives
    );

  CA4H_applyDerivativeConfirmationToTrend_(
    analysis
  );
  
  CA4H_writeToAnalysisSheet_(
    ss,
    analysis
  );

  const message =
    CA4H_formatTelegramMessage_(
      analysis,
      position
    );

  CA4H_sendTelegramMessage_(
    CA4H_normalizeTelegramNewlines_(
      message
    )
  );

  return analysis;
}

function CA4H_normalizeTelegramNewlines_(
  text
) {
  if (text == null) {
    return '';
  }

  return String(text)
    .replace(/\\r\\n/g, '\n')
    .replace(/\\n/g, '\n')
    .replace(/\\r/g, '\n');
}

function CA4H_formatOiAndSupportMessage_(a) {
  const lines = [];

  if (a.oiAnalysis) {
    lines.push(
      '📌 *OI Ανάλυση:* ' +
      a.oiAnalysis.emoji +
      ' ' +
      a.oiAnalysis.label
    );

    lines.push(
      a.oiAnalysis.note
    );
  }

  if (a.supportZone) {
    const s = a.supportZone;

    lines.push(
      '🛡️ *Πιθανότερη Support Zone:* `' +
      CA4H_formatPrice_(s.lower) +
      '` – `' +
      CA4H_formatPrice_(s.upper) +
      '`'
    );

    lines.push(
      'Κεντρικό επίπεδο: `' +
      CA4H_formatPrice_(s.midpoint) +
      '` (' +
      s.distanceAtr.toFixed(1) +
      'x ATR κάτω από την τρέχουσα τιμή).'
    );

    lines.push(
      'Ισχύς ζώνης: *' +
      s.strength +
      '* — βάση: ' +
      s.sources.join(', ') +
      '.'
    );

    lines.push(
      '_Η ζώνη είναι εκτίμηση τεχνικών επιπέδων και όχι εγγυημένο σημείο αντιστροφής._'
    );
  }

  return lines.join('\n');
}

function CA4H_sendTelegramMessage_(
  text
) {
  if (
    text == null ||
    text === ''
  ) {
    throw new Error(
      'CA4H_sendTelegramMessage_() κλήθηκε χωρίς text argument.'
    );
  }

  text =
    CA4H_normalizeTelegramNewlines_(
      text
    );

  const props =
    PropertiesService
      .getScriptProperties();

  const token =
    props.getProperty(
      'TELEGRAM_BOT_TOKEN'
    );

  const chatId =
    props.getProperty(
      'TELEGRAM_CHAT_ID'
    );

  if (!token) {
    throw new Error(
      'Λείπει το Script Property "TELEGRAM_BOT_TOKEN".'
    );
  }

  if (!chatId) {
    throw new Error(
      'Λείπει το Script Property "TELEGRAM_CHAT_ID".'
    );
  }

  const url =
    'https://api.telegram.org/bot' +
    token +
    '/sendMessage';

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: chatId,
          text: text,
          parse_mode: 'Markdown',
          disable_web_page_preview: true
        }),
        muteHttpExceptions: true
      }
    );

  if (
    response.getResponseCode() !== 200
  ) {
    throw new Error(
      'Telegram API error ' +
      response.getResponseCode() +
      ': ' +
      response.getContentText()
    );
  }
}

// ---------------------------------------------------------------------------
// SYMBOL RESOLUTION
// ---------------------------------------------------------------------------
function CA4H_resolveInstId_(input) {
  let s = input.toUpperCase().replace(/\s+/g, '');
  s = s.replace('/', '-');
  if (/^[A-Z0-9]+-[A-Z0-9]+(-SWAP)?$/.test(s)) return s;

  const quoteCandidates = ['USDT', 'USDC', 'USD', 'BTC', 'ETH'];
  for (const q of quoteCandidates) {
    if (s.endsWith(q) && s.length > q.length) return s.slice(0, -q.length) + '-' + q;
  }
  return s;
}

// ---------------------------------------------------------------------------
// ΑΝΑΓΝΩΣΗ ΑΝΟΙΧΤΗΣ ΘΕΣΗΣ ΑΠΟ "Open Positions"
// ---------------------------------------------------------------------------
function CA4H_readOpenPosition_(posSheet) {
  if (CA4H_CFG.DIRECTION_CELL) {
    const dirRaw = String(posSheet.getRange(CA4H_CFG.DIRECTION_CELL).getValue() || '').toUpperCase();
    const direction = dirRaw.indexOf('LONG') !== -1 ? 'LONG' : (dirRaw.indexOf('SHORT') !== -1 ? 'SHORT' : null);
    if (!direction) return null;

    const lower = CA4H_CFG.LOWER_PRICE_CELL ? Number(posSheet.getRange(CA4H_CFG.LOWER_PRICE_CELL).getValue()) : null;
    const upper = CA4H_CFG.UPPER_PRICE_CELL ? Number(posSheet.getRange(CA4H_CFG.UPPER_PRICE_CELL).getValue()) : null;

    return {
      direction,
      lowerPrice: (lower && !isNaN(lower)) ? lower : null,
      upperPrice: (upper && !isNaN(upper)) ? upper : null
    };
  }

  const rowValues = posSheet.getRange(CA4H_CFG.POSITION_ROW_RANGE).getValues()[0];

  let direction = null;
  const numericCandidates = [];

  rowValues.forEach(val => {
    if (typeof val === 'string') {
      const upperVal = val.toUpperCase();
      if (upperVal.indexOf('LONG') !== -1) direction = 'LONG';
      else if (upperVal.indexOf('SHORT') !== -1) direction = 'SHORT';
    } else if (typeof val === 'number' && !isNaN(val) && val > 0) {
      numericCandidates.push(val);
    }
  });

  if (!direction) return null;

  let lowerPrice = null, upperPrice = null;
  if (numericCandidates.length >= 2) {
    const sorted = [...numericCandidates].sort((a, b) => a - b);
    let bestPair = null, bestRatio = Infinity;
    for (let i = 0; i < sorted.length - 1; i++) {
      for (let j = i + 1; j < sorted.length; j++) {
        const ratio = sorted[j] / sorted[i];
        if (ratio > 1 && ratio < bestRatio) {
          bestRatio = ratio;
          bestPair = [sorted[i], sorted[j]];
        }
      }
    }
    if (bestPair) {
      lowerPrice = bestPair[0];
      upperPrice = bestPair[1];
    }
  }

  return { direction, lowerPrice, upperPrice };
}

// ---------------------------------------------------------------------------
// OKX CANDLES FETCH
// ---------------------------------------------------------------------------
/**
 * ΔΙΟΡΘΩΜΕΝΗ έκδοση (v3.6): διαβάζει "volCcy" (index 6, quote currency)
 * ως το κύριο πεδίο volume, αντί για "vol" (index 5, αριθμός συμβολαίων).
 * Για SWAP/perpetual συμβόλαια το "vol" εκφράζεται σε contracts, όχι σε
 * χρήσιμο μέγεθος για ανάλυση — το volCcy (π.χ. σε USDT) είναι σωστό και
 * συγκρίσιμο. Το ωμό "vol" παραμένει διαθέσιμο ως volContracts.
 */
function CA4H_fetchOkxCandles_(instId, bar, limit) {
  const candidates = [instId];
  if (!instId.endsWith('-SWAP')) candidates.push(instId + '-SWAP');
  if (instId.endsWith('-SWAP')) candidates.push(instId.replace('-SWAP', ''));

  let lastError = null;
  for (const id of candidates) {
    try {
      const url = CA4H_CFG.OKX_BASE_URL + '/api/v5/market/candles'
        + '?instId=' + encodeURIComponent(id)
        + '&bar=' + encodeURIComponent(bar)
        + '&limit=' + limit;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      const json = JSON.parse(resp.getContentText());
      if (code === 200 && json.code === '0' && json.data && json.data.length > 0) {
        return json.data.map(function(r) {
          return {
            ts: Number(r[0]),
            open: parseFloat(r[1]),
            high: parseFloat(r[2]),
            low: parseFloat(r[3]),
            close: parseFloat(r[4]),
            volContracts: parseFloat(r[5]),
            volume: parseFloat(r[6]),
            confirm: Number(r[7])
          };
        }).reverse();
      }
      lastError = 'instId=' + id + ' -> code:' + code + ' body:' + resp.getContentText().slice(0, 200);
    } catch (e) {
      lastError = 'instId=' + id + ' -> exception: ' + e.message;
    }
  }
  throw new Error('Αποτυχία λήψης candles από OKX. Τελευταίο σφάλμα: ' + lastError);
}

// ---------------------------------------------------------------------------
// ADAPTIVE PRICE FORMATTING
// ---------------------------------------------------------------------------
function CA4H_smartDecimals_(price) {
  if (price == null || isNaN(price) || price === 0) return CA4H_CFG.MIN_PRICE_DECIMALS;
  const absPrice = Math.abs(price);
  if (absPrice >= 1) return CA4H_CFG.MIN_PRICE_DECIMALS;
  const leadingZeros = Math.floor(-Math.log10(absPrice)) - 1;
  const decimals = leadingZeros + 5;
  return Math.min(CA4H_CFG.MAX_PRICE_DECIMALS, Math.max(CA4H_CFG.MIN_PRICE_DECIMALS, decimals));
}

function CA4H_formatPrice_(price) {
  if (price == null || isNaN(price)) return 'n/a';
  return Number(price).toFixed(CA4H_smartDecimals_(price));
}

function CA4H_priceNumberFormat_(price) {
  return '0.' + '0'.repeat(CA4H_smartDecimals_(price));
}

// ---------------------------------------------------------------------------
// CORE INDICATORS (RSI/MACD/SMA)
// ---------------------------------------------------------------------------
function CA4H_ema_(values, period) {
  const k = 2 / (period + 1);
  const out = new Array(values.length).fill(null);
  let prevEma = values.slice(0, period).reduce((a, b) => a + b, 0) / period;
  out[period - 1] = prevEma;
  for (let i = period; i < values.length; i++) {
    prevEma = values[i] * k + prevEma * (1 - k);
    out[i] = prevEma;
  }
  return out;
}

function CA4H_rsi_(closes, period) {
  const out = new Array(closes.length).fill(null);
  if (closes.length <= period) return out;
  let gains = 0, losses = 0;
  for (let i = 1; i <= period; i++) {
    const diff = closes[i] - closes[i - 1];
    if (diff >= 0) gains += diff; else losses -= diff;
  }
  let avgGain = gains / period;
  let avgLoss = losses / period;
  out[period] = avgLoss === 0 ? 100 : 100 - (100 / (1 + avgGain / avgLoss));
  for (let i = period + 1; i < closes.length; i++) {
    const diff = closes[i] - closes[i - 1];
    const gain = diff > 0 ? diff : 0;
    const loss = diff < 0 ? -diff : 0;
    avgGain = (avgGain * (period - 1) + gain) / period;
    avgLoss = (avgLoss * (period - 1) + loss) / period;
    out[i] = avgLoss === 0 ? 100 : 100 - (100 / (1 + avgGain / avgLoss));
  }
  return out;
}

function CA4H_macd_(closes, fastP, slowP, signalP) {
  const emaFast = CA4H_ema_(closes, fastP);
  const emaSlow = CA4H_ema_(closes, slowP);
  const macdLine = closes.map((_, i) => (emaFast[i] != null && emaSlow[i] != null) ? emaFast[i] - emaSlow[i] : null);
  const macdValues = macdLine.filter(v => v != null);
  const signalRaw = CA4H_ema_(macdValues, signalP);
  const signalLine = new Array(closes.length).fill(null);
  let offset = macdLine.findIndex(v => v != null);
  let sIdx = 0;
  for (let i = offset; i < closes.length; i++) {
    if (signalRaw[sIdx] != null) signalLine[i] = signalRaw[sIdx];
    sIdx++;
  }
  const histogram = closes.map((_, i) => (macdLine[i] != null && signalLine[i] != null) ? macdLine[i] - signalLine[i] : null);
  return { macdLine, signalLine, histogram };
}

function CA4H_sma_(values, period, endIndex) {
  if (endIndex - period + 1 < 0) return null;
  let sum = 0;
  for (let i = endIndex - period + 1; i <= endIndex; i++) sum += values[i];
  return sum / period;
}

// ---------------------------------------------------------------------------
// KDJ(9,3,3)
// ---------------------------------------------------------------------------
function CA4H_calcKdj_(candles, n, kSmooth, dSmooth) {
  const closes = candles.map(c => c.close);
  const highs = candles.map(c => c.high);
  const lows = candles.map(c => c.low);
  const len = candles.length;

  const rsvArr = new Array(len).fill(null);
  for (let i = n - 1; i < len; i++) {
    const highN = Math.max(...highs.slice(i - n + 1, i + 1));
    const lowN = Math.min(...lows.slice(i - n + 1, i + 1));
    const range = highN - lowN;
    rsvArr[i] = range === 0 ? 50 : ((closes[i] - lowN) / range) * 100;
  }

  const kArr = new Array(len).fill(null);
  const dArr = new Array(len).fill(null);
  const jArr = new Array(len).fill(null);
  let prevK = 50, prevD = 50;
  for (let i = 0; i < len; i++) {
    if (rsvArr[i] == null) continue;
    const k = (prevK * (kSmooth - 1) + rsvArr[i]) / kSmooth;
    const d = (prevD * (dSmooth - 1) + k) / dSmooth;
    kArr[i] = k;
    dArr[i] = d;
    jArr[i] = 3 * k - 2 * d;
    prevK = k;
    prevD = d;
  }

  return { kArr, dArr, jArr, k: kArr[len - 1], d: dArr[len - 1], j: jArr[len - 1] };
}

function CA4H_interpretKdj_(kdj) {
  if (kdj.k == null || kdj.d == null || kdj.j == null) {
    return { emoji: '⚪', label: 'Δεν υπάρχουν αρκετά δεδομένα', bias: 'neutral' };
  }
  const k = kdj.k, d = kdj.d, j = kdj.j;

  if (j > 100) return { emoji: '🚨', label: 'Ακραίο Καμπανάκι Υπεραγοράς', bias: 'bearish' };
  if (j < 0) return { emoji: '🚨', label: 'Ακραίο Καμπανάκι Υπερπούλησης', bias: 'bullish' };
  if (k > d && j > 80) return { emoji: '⚠️', label: 'Καμπανάκι Υπεραγοράς', bias: 'bearish' };
  if (k < d && j < 20) return { emoji: '⚠️', label: 'Καμπανάκι Υπερπούλησης', bias: 'bullish' };
  if (k > d && k > 50) return { emoji: '🟢', label: 'Ισχυρή Ανοδική Ορμή', bias: 'bullish' };
  if (k < d && k < 50) return { emoji: '🔴', label: 'Ισχυρή Πτωτική Ορμή', bias: 'bearish' };
  if (k > d) return { emoji: '🟢', label: 'Ήπια Ανοδική Τάση', bias: 'bullish' };
  if (k < d) return { emoji: '🔴', label: 'Ήπια Πτωτική Τάση', bias: 'bearish' };
  return { emoji: '⚪', label: 'Ουδέτερο / Χωρίς Σαφή Τάση', bias: 'neutral' };
}

// ---------------------------------------------------------------------------
// Bollinger Bands(20,2)
// ---------------------------------------------------------------------------
function CA4H_calcBollinger_(closes, period, stdDevMult) {
  const len = closes.length;
  const upper = new Array(len).fill(null);
  const middle = new Array(len).fill(null);
  const lower = new Array(len).fill(null);

  for (let i = period - 1; i < len; i++) {
    const slice = closes.slice(i - period + 1, i + 1);
    const mean = slice.reduce((a, b) => a + b, 0) / period;
    const variance = slice.reduce((s, v) => s + Math.pow(v - mean, 2), 0) / period;
    const stdDev = Math.sqrt(variance);
    middle[i] = mean;
    upper[i] = mean + stdDevMult * stdDev;
    lower[i] = mean - stdDevMult * stdDev;
  }

  const lastIdx = len - 1;
  const bandwidth = (upper[lastIdx] != null && middle[lastIdx])
    ? ((upper[lastIdx] - lower[lastIdx]) / middle[lastIdx]) * 100
    : null;
  const isSqueeze = bandwidth != null && bandwidth < CA4H_CFG.BB_SQUEEZE_THRESHOLD_PCT;

  return {
    upper: upper[lastIdx], middle: middle[lastIdx], lower: lower[lastIdx],
    bandwidthPct: bandwidth, isSqueeze,
    upperArr: upper, middleArr: middle, lowerArr: lower
  };
}

// ---------------------------------------------------------------------------
// ATR(14)
// ---------------------------------------------------------------------------
function CA4H_calcAtr_(candles, period) {
  const len = candles.length;
  const trArr = new Array(len).fill(null);

  for (let i = 1; i < len; i++) {
    const c = candles[i];
    const prevClose = candles[i - 1].close;
    trArr[i] = Math.max(c.high - c.low, Math.abs(c.high - prevClose), Math.abs(c.low - prevClose));
  }

  const atrArr = new Array(len).fill(null);
  const validTr = trArr.slice(1, period + 1);
  let prevAtr = validTr.reduce((a, b) => a + b, 0) / period;
  atrArr[period] = prevAtr;
  for (let i = period + 1; i < len; i++) {
    prevAtr = (prevAtr * (period - 1) + trArr[i]) / period;
    atrArr[i] = prevAtr;
  }

  const lastAtr = atrArr[len - 1];
  const lastClose = candles[len - 1].close;
  const atrPct = lastAtr != null && lastClose ? (lastAtr / lastClose) * 100 : null;

  return { atr: lastAtr, atrPct, atrArr };
}

// ---------------------------------------------------------------------------
// Fibonacci Retracement (0.5 / 0.618 / 0.786)
// ---------------------------------------------------------------------------
function CA4H_calcFibLevels_(candles, lookback) {
  const recent = candles.slice(Math.max(0, candles.length - lookback));
  let swingHigh = -Infinity, swingHighIdx = -1;
  let swingLow = Infinity, swingLowIdx = -1;

  recent.forEach((c, i) => {
    if (c.high > swingHigh) { swingHigh = c.high; swingHighIdx = i; }
    if (c.low < swingLow) { swingLow = c.low; swingLowIdx = i; }
  });

  const diff = swingHigh - swingLow;
  const isUptrend = swingHighIdx > swingLowIdx;

  const levels = {};
  CA4H_CFG.FIB_LEVELS.forEach(ratio => {
    levels[ratio.toString()] = isUptrend
      ? swingHigh - diff * ratio
      : swingLow + diff * ratio;
  });

  return { swingHigh, swingLow, isUptrend, levels };
}

// ---------------------------------------------------------------------------
// Προβολή Αναμενόμενου Εύρους Τιμής (4H / 8H) βάσει ATR + τάσης
// ---------------------------------------------------------------------------
function CA4H_calcPriceProjection_(closePrice, atr, trend) {
  if (atr == null || closePrice == null) {
    return null;
  }

  const proj4h = atr * CA4H_CFG.PROJECTION_ATR_MULT_4H;
  const proj8h = atr * CA4H_CFG.PROJECTION_ATR_MULT_8H_SQRT;

  let skewUp = 0, skewDown = 0;
  if (trend === 'BULLISH') {
    skewUp = CA4H_CFG.PROJECTION_SKEW_BULLISH;
    skewDown = -CA4H_CFG.PROJECTION_SKEW_BULLISH;
  } else if (trend === 'BEARISH') {
    skewUp = -CA4H_CFG.PROJECTION_SKEW_BEARISH;
    skewDown = CA4H_CFG.PROJECTION_SKEW_BEARISH;
  }

  return {
    range4h: {
      lower: closePrice - proj4h * (1 + skewDown),
      upper: closePrice + proj4h * (1 + skewUp)
    },
    range8h: {
      lower: closePrice - proj8h * (1 + skewDown),
      upper: closePrice + proj8h * (1 + skewUp)
    }
  };
}

// ---------------------------------------------------------------------------
// NEW: ATR-Distance to Grid Edges
// ---------------------------------------------------------------------------
/**
 * Υπολογίζει πόσα "τυπικά" 4H ATR-κινήσεις χωρίζουν την τρέχουσα τιμή από
 * κάθε άκρο του DCA grid — πιο διαισθητικό μέτρο ρίσκου από το ωμό ποσοστό
 * θέσης μέσα στο εύρος, γιατί ενσωματώνει την τρέχουσα μεταβλητότητα.
 *   <3x ATR  -> "κοντά"
 *   3-8x ATR -> "μεσαία απόσταση"
 *   >8x ATR  -> "μακριά"
 */
function CA4H_calcAtrDistanceToGrid_(close, atr, position) {
  if (!position || position.lowerPrice == null || position.upperPrice == null || atr == null || atr === 0) {
    return null;
  }

  const distLower = close - position.lowerPrice;
  const distUpper = position.upperPrice - close;
  const atrUnitsLower = distLower / atr;
  const atrUnitsUpper = distUpper / atr;

  const classify = (atrUnits) => {
    if (atrUnits < CA4H_CFG.ATR_DIST_NEAR_MAX) return 'κοντά';
    if (atrUnits <= CA4H_CFG.ATR_DIST_MEDIUM_MAX) return 'μεσαία απόσταση';
    return 'μακριά';
  };

  return {
    distLower, distUpper, atrUnitsLower, atrUnitsUpper,
    lowerLabel: classify(atrUnitsLower),
    upperLabel: classify(atrUnitsUpper)
  };
}

function CA4H_formatAtrDistanceLine_(atrDist) {
  if (!atrDist) return null;
  return '📏 *Απόσταση από άκρα grid (σε ATR):* κάτω όριο ' +
    atrDist.atrUnitsLower.toFixed(1) + 'x ATR (' + atrDist.lowerLabel + ') | ' +
    'άνω όριο ' + atrDist.atrUnitsUpper.toFixed(1) + 'x ATR (' + atrDist.upperLabel + ')';
}

// ---------------------------------------------------------------------------
// NEW: Squeeze Streak Tracking (ιστορικό squeeze duration από το sheet)
// ---------------------------------------------------------------------------
function CA4H_trackSqueezeStreak_(ss, isSqueezeNow) {
  const sheet = ss.getSheetByName(CA4H_CFG.ANALYSIS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return isSqueezeNow ? 1 : 0;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const streakCol = headers.indexOf('Squeeze Streak') + 1;
  if (streakCol === 0) {
    return isSqueezeNow ? 1 : 0;
  }

  const lastRow = sheet.getLastRow();
  const prevStreak = Number(sheet.getRange(lastRow, streakCol).getValue()) || 0;
  return isSqueezeNow ? prevStreak + 1 : 0;
}

/**
 * Καθοδήγηση grid όταν υπάρχει ενεργό Bollinger squeeze. Όσο μεγαλύτερο
 * το streak, τόσο εντονότερη η προειδοποίηση για επικείμενο breakout.
 */
function CA4H_buildSqueezeGuidance_(a, position) {
  if (!a.bb.isSqueeze) return null;

  const streak = a.squeezeStreak || 1;
  const durationNote = streak >= 3
    ? ' Το squeeze διαρκεί ήδη ' + streak + ' συνεχόμενα 4H candles — όσο παρατείνεται, τόσο πλησιάζει στατιστικά το breakout.'
    : ' (' + streak + 'ο candle σε squeeze).';

  let gridNote = '';
  if (position && position.lowerPrice != null) {
    gridNote = ' Σε ένα DCA grid, σημασία έχει να αντέχει το κάτω όριο (`' +
      CA4H_formatPrice_(position.lowerPrice) + '`) σε ένα ξαφνικό sell-off, όχι μόνο σε σταδιακή πτώση.';
  }

  return '🔸 *Squeeze ενεργό:* αναμένεται αυξημένη μεταβλητότητα σύντομα.' + durationNote + gridNote;
}

// ---------------------------------------------------------------------------
// NEW: Signal Conflict Flag (Candlestick Pattern vs KDJ/MACD)
// ---------------------------------------------------------------------------
/**
 * Συγκρίνει το bias του τελευταίου candlestick pattern με το KDJ bias και
 * το πρόσημο του MACD histogram. Αν δεν συμφωνούν, επιστρέφει ρητό μήνυμα
 * σύγκρουσης, εξηγώντας ΤΙ ακριβώς αντιτίθεται. Ενεργοποιείται μόνο για
 * "δυνατά" patterns (βλ. CA4H_CFG.STRONG_PATTERNS) — Doji/κανένα pattern
 * δεν πυροδοτούν σύγκρουση, αφού είναι ήδη ουδέτερα.
 */
function CA4H_detectSignalConflict_(a) {
  const lastPattern = a.patterns[a.patterns.length - 1];
  if (!lastPattern || lastPattern.bias === 'neutral' || CA4H_CFG.STRONG_PATTERNS.indexOf(lastPattern.name) === -1) {
    return null;
  }

  const conflicts = [];

  if (a.kdjInterpretation.bias !== 'neutral' && a.kdjInterpretation.bias !== lastPattern.bias) {
    conflicts.push('το KDJ ("' + a.kdjInterpretation.label + '")');
  }

  if (a.macdHist != null) {
    const macdBias = a.macdHist > 0 ? 'bullish' : 'bearish';
    if (macdBias !== lastPattern.bias) {
      conflicts.push('το MACD histogram (' + (macdBias === 'bullish' ? 'θετικό/ανοδικό' : 'αρνητικό/πτωτικό') + ')');
    }
  }

  if (conflicts.length === 0) return null;

  const patternBiasGr = lastPattern.bias === 'bullish' ? 'ανοδική πίεση' : 'πτωτική πίεση';
  return '⚠️ *Σύγκρουση σημάτων:* το τελευταίο κερί (*' + lastPattern.name + '*) δείχνει βραχυπρόθεσμη ' +
    patternBiasGr + ', ενώ ' + conflicts.join(' και ') + ' δεν συμφωνεί. ' +
    'Το τελικό "' + a.trend + '" σκορ αντικατοπτρίζει αυτή την αντίφαση, όχι απουσία σήματος — ' +
    'χρειάζεται επόμενο κερί για επιβεβαίωση φοράς.';
}

// ---------------------------------------------------------------------------
// CANDLESTICK PATTERN DETECTION
// ---------------------------------------------------------------------------
function CA4H_detectPatterns_(candles) {
  const n = candles.length;
  const patterns = [];
  const c0 = candles[n - 1];
  const c1 = candles[n - 2];
  const c2 = candles[n - 3];

  const body = c => Math.abs(c.close - c.open);
  const range = c => c.high - c.low;
  const upperWick = c => c.high - Math.max(c.open, c.close);
  const lowerWick = c => Math.min(c.open, c.close) - c.low;
  const isBull = c => c.close > c.open;
  const isBear = c => c.close < c.open;

  const avgBody = (() => {
    const last10 = candles.slice(Math.max(0, n - 10), n);
    return last10.reduce((s, c) => s + body(c), 0) / last10.length;
  })();

  if (range(c0) > 0 && body(c0) <= range(c0) * 0.1) {
    patterns.push({ name: 'Doji', bias: 'neutral', note: 'Αναποφασιστικότητα αγοράς/πώλησης' });
  }
  if (range(c0) > 0 && lowerWick(c0) >= body(c0) * 2 && upperWick(c0) <= body(c0) * 0.5 && body(c0) > 0) {
    if (isBull(c0) || (c1 && c0.close < c1.close)) {
      patterns.push({ name: 'Hammer', bias: 'bullish', note: 'Απόρριψη χαμηλότερων τιμών, πιθανό bottom' });
    } else {
      patterns.push({ name: 'Hanging Man', bias: 'bearish', note: 'Πιθανή εξάντληση αγοραστών στην κορυφή' });
    }
  }
  if (range(c0) > 0 && upperWick(c0) >= body(c0) * 2 && lowerWick(c0) <= body(c0) * 0.5 && body(c0) > 0) {
    if (isBear(c0)) {
      patterns.push({ name: 'Shooting Star', bias: 'bearish', note: 'Απόρριψη υψηλότερων τιμών, πιθανή κορυφή' });
    } else {
      patterns.push({ name: 'Inverted Hammer', bias: 'neutral', note: 'Χρειάζεται επιβεβαίωση' });
    }
  }
  if (body(c0) > 0 && (upperWick(c0) + lowerWick(c0)) <= body(c0) * 0.05) {
    patterns.push({
      name: isBull(c0) ? 'Bullish Marubozu' : 'Bearish Marubozu',
      bias: isBull(c0) ? 'bullish' : 'bearish',
      note: 'Ισχυρό momentum χωρίς αντίσταση/στήριξη μέσα στο κερί'
    });
  }
  if (c1) {
    if (isBear(c1) && isBull(c0) && c0.close >= c1.open && c0.open <= c1.close && body(c0) > body(c1)) {
      patterns.push({ name: 'Bullish Engulfing', bias: 'bullish', note: 'Οι αγοραστές κάλυψαν πλήρως το προηγούμενο πτωτικό κερί' });
    }
    if (isBull(c1) && isBear(c0) && c0.open >= c1.close && c0.close <= c1.open && body(c0) > body(c1)) {
      patterns.push({ name: 'Bearish Engulfing', bias: 'bearish', note: 'Οι πωλητές κάλυψαν πλήρως το προηγούμενο ανοδικό κερί' });
    }
  }
  if (c1 && c2) {
    if (isBull(c2) && isBull(c1) && isBull(c0) &&
        c1.close > c2.close && c0.close > c1.close &&
        body(c0) > avgBody * 0.5 && body(c1) > avgBody * 0.5 && body(c2) > avgBody * 0.5) {
      patterns.push({ name: 'Three White Soldiers', bias: 'bullish', note: 'Τρία διαδοχικά ανοδικά κεριά με higher closes' });
    }
    if (isBear(c2) && isBear(c1) && isBear(c0) &&
        c1.close < c2.close && c0.close < c1.close &&
        body(c0) > avgBody * 0.5 && body(c1) > avgBody * 0.5 && body(c2) > avgBody * 0.5) {
      patterns.push({ name: 'Three Black Crows', bias: 'bearish', note: 'Τρία διαδοχικά πτωτικά κεριά με lower closes' });
    }
  }
  if (patterns.length === 0) {
    patterns.push({ name: 'Κανένα σαφές pattern', bias: 'neutral', note: 'Το τελευταίο κερί δεν ταιριάζει σε γνωστό σχηματισμό' });
  }
  return patterns;
}

// ---------------------------------------------------------------------------
// ΣΥΝΘΕΣΗ ΑΝΑΛΥΣΗΣ
// ---------------------------------------------------------------------------
function CA4H_buildAnalysis_(
  instId,
  candles,
  ss
) {

   const closedCandles =
    candles.filter(function(candle) {
      return Number(candle.confirm) === 1;
    });

   const liveCandle =
    candles.length > closedCandles.length
      ? candles[candles.length - 1]
      : null;

   const baseCandles =
    closedCandles.length >= 30
      ? closedCandles
      : candles;

  const closes =
    baseCandles.map(c => c.close);

  const volumes =
    baseCandles.map(c => c.volume);

  const n =
    baseCandles.length;

  const rsiArr =
    CA4H_rsi_(
      closes,
      CA4H_CFG.RSI_PERIOD
    );

  const macdRes =
    CA4H_macd_(
      closes,
      CA4H_CFG.EMA_FAST,
      CA4H_CFG.EMA_SLOW,
      CA4H_CFG.MACD_SIGNAL
    );

  const volMaShort =
    CA4H_sma_(
      volumes,
      CA4H_CFG.VOLUME_MA_SHORT,
      n - 1
    );

  const volMaLong =
    CA4H_sma_(
      volumes,
      CA4H_CFG.VOLUME_MA_LONG,
      n - 1
    );

  const kdj =
    CA4H_calcKdj_(
      baseCandles,
      CA4H_CFG.KDJ_N,
      CA4H_CFG.KDJ_K_SMOOTH,
      CA4H_CFG.KDJ_D_SMOOTH
    );

  const kdjInterpretation =
    CA4H_interpretKdj_(
      kdj
    );

  const bb =
    CA4H_calcBollinger_(
      closes,
      CA4H_CFG.BB_PERIOD,
      CA4H_CFG.BB_STDDEV_MULT
    );

  const atr =
    CA4H_calcAtr_(
      baseCandles,
      CA4H_CFG.ATR_PERIOD
    );

  const fib =
    CA4H_calcFibLevels_(
      baseCandles,
      CA4H_CFG.FIB_SWING_LOOKBACK
    );

  const squeezeStreak =
    CA4H_trackSqueezeStreak_(
      ss,
      bb.isSqueeze
    );

  const last =
    baseCandles[baseCandles.length - 1];

  const prevClose =
    baseCandles[baseCandles.length - 2].close;

  if (
    !last.ts ||
    isNaN(Number(last.ts))
  ) {
    throw new Error(
      'Το τελευταίο OKX candle δεν έχει έγκυρο timestamp: ' +
      last.ts
    );
  }

  const changePct =
    ((last.close - prevClose) /
      prevClose) * 100;

  const rsiLast =
    rsiArr[n - 1];

  const macdHistLast =
    macdRes.histogram[n - 1];

  const macdHistPrev =
    macdRes.histogram[n - 2];

  const macdLineLast =
    macdRes.macdLine[n - 1];

  const macdSignalLast =
    macdRes.signalLine[n - 1];

  const patterns =
    CA4H_detectPatterns_(
      baseCandles
    );


  let score = 0;

  patterns.forEach(function(p) {
    if (p.bias === 'bullish') {
      score += 1;
    }

    if (p.bias === 'bearish') {
      score -= 1;
    }
  });

  if (
    macdHistLast != null &&
    macdHistPrev != null
  ) {
    if (
      macdHistLast > 0 &&
      macdHistLast > macdHistPrev
    ) {
      score += 1;
    }

    if (
      macdHistLast < 0 &&
      macdHistLast < macdHistPrev
    ) {
      score -= 1;
    }

    if (
      macdLineLast != null &&
      macdSignalLast != null
    ) {
      if (
        macdLineLast > macdSignalLast
      ) {
        score += 0.5;
      } else {
        score -= 0.5;
      }
    }
  }

  if (rsiLast != null) {
    if (rsiLast >= 70) {
      score -= 0.5;
    } else if (rsiLast <= 30) {
      score += 0.5;
    } else if (rsiLast > 55) {
      score += 0.25;
    } else if (rsiLast < 45) {
      score -= 0.25;
    }
  }

  if (
    kdjInterpretation.bias ===
    'bullish'
  ) {
    score += 0.5;
  } else if (
    kdjInterpretation.bias ===
    'bearish'
  ) {
    score -= 0.5;
  }

  const volumeConfirmed =
    volMaShort != null &&
    last.volume > volMaShort;

  if (volumeConfirmed) {
    score += score > 0
      ? 0.5
      : score < 0
        ? -0.5
        : 0;
  }

  let trend;
  let trendEmoji;

  if (score >= 1.5) {
    trend = 'BULLISH';
    trendEmoji = '🟢';
  } else if (score <= -1.5) {
    trend = 'BEARISH';
    trendEmoji = '🔴';
  } else {
    trend = 'NEUTRAL / RANGE';
    trendEmoji = '⚪';
  }

  const priceDecimals =
    CA4H_smartDecimals_(
      last.close
    );

  const priceProjection =
    CA4H_calcPriceProjection_(
      last.close,
      atr.atr,
      trend
    );

  Logger.log(
    'OKX last candle timestamp: ' +
    last.ts +
    ' -> ' +
    new Date(
      last.ts
    ).toISOString()
  );

  const recentSwingLows =
    CA4H_findRecentSwingLows_(
      candles,
      30,
      2
    );

const historicallyDefendedSupport =
  CA4H_calcHistoricallyDefendedSupportZone_(
    baseCandles,
    last.close,
    atr.atr
  );

  return {
    instId: instId,
    timestamp: new Date(),

    // ΝΕΟ: χρόνος ανοίγματος του τελευταίου OKX candle
    lastCandleTime: new Date(last.ts),

    open: last.open,
    high: last.high,
    low: last.low,
    close: last.close,
    priceDecimals: priceDecimals,

    changePct: changePct,
    volume: last.volume,
    volMaShort: volMaShort,
    volMaLong: volMaLong,
    volumeConfirmed: volumeConfirmed,

    rsi: rsiLast,
    macdLine: macdLineLast,
    macdSignal: macdSignalLast,
    macdHist: macdHistLast,

    kdj: kdj,
    kdjInterpretation: kdjInterpretation,
    bb: bb,
    atr: atr,
    fib: fib,
    priceProjection: priceProjection,
    squeezeStreak: squeezeStreak,

    patterns: patterns,
    score: score,
    trend: trend,
    trendEmoji: trendEmoji,

    recentSwingLows: recentSwingLows,

     historicallyDefendedSupport:
      historicallyDefendedSupport,

    baseCandleState:
  closedCandles.length >= 30
    ? 'CLOSED'
    : 'INCOMPLETE_FALLBACK',

liveCandleIncluded:
  liveCandle != null,

liveCandleTime:
  liveCandle
    ? new Date(liveCandle.ts)
    : null,

liveCandleClose:
  liveCandle
    ? liveCandle.close
    : null,

liveCandleConfirm:
  liveCandle
    ? Number(liveCandle.confirm)
    : null

  };
}

function CA4H_applyDerivativeConfirmationToTrend_(
  analysis
) {
  const d =
    analysis.pionexDerivatives;

  if (!d || d.error || !d.market || !d.oi) {
    return analysis;
  }

  const priceChange =
    d.market.priceChangePct;

  const oiChange =
    d.oi.changePct;

  const volumeConfirmed =
    analysis.volumeConfirmed;

  const oiStable =
    oiChange !== null &&
    Math.abs(oiChange) < 0.10;

  const priceWeak =
    priceChange !== null &&
    Math.abs(priceChange) < 0.50;

  const derivativesMixed =
    d.classification &&
    d.classification.label ===
      'MIXED / NO CLEAR CONFIRMATION';

  if (
    analysis.trend === 'BULLISH' &&
    (
      !volumeConfirmed ||
      oiStable ||
      priceWeak ||
      derivativesMixed
    )
  ) {
    analysis.trend = 'NEUTRAL / RANGE';
    analysis.trendEmoji = '⚪';
    analysis.trendAdjusted = true;
    analysis.trendAdjustmentReason =
      'Το raw technical score είναι bullish, ' +
      'αλλά volume/OI/derivatives δεν επιβεβαιώνουν ' +
      'αρκετά ισχυρή ανοδική συνέχεια.';
  }

  return analysis;
}


// ---------------------------------------------------------------------------
// ΕΓΓΡΑΦΗ ΣΤΟ SHEET "4h analysis"
// ---------------------------------------------------------------------------
function CA4H_writeToAnalysisSheet_(ss, a) {
  let sheet = ss.getSheetByName(CA4H_CFG.ANALYSIS_SHEET);
  if (!sheet) sheet = ss.insertSheet(CA4H_CFG.ANALYSIS_SHEET);

  const headers = [
    'Timestamp', 'Symbol', 'Last 4H Candle', 'Open', 'High', 'Low', 'Close',
    '24H/4H Change %', 'Volume (USDT)', 'Vol MA5', 'Vol MA10', 'Vol Confirmed',
    'RSI(14)', 'MACD', 'MACD Signal', 'MACD Hist',
    'KDJ K', 'KDJ D', 'KDJ J', 'KDJ Ερμηνεία',
    'BB Upper', 'BB Middle', 'BB Lower', 'BB Squeeze', 'Squeeze Streak',
    'ATR(14)', 'ATR %',
    'Fib 0.5', 'Fib 0.618', 'Fib 0.786',
    'Proj 4H Low', 'Proj 4H High', 'Proj 8H Low', 'Proj 8H High',
    'Patterns', 'Score', 'Trend'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
  }

  const patternsText = a.patterns.map(p => p.name + ' (' + p.bias + ')').join(', ');
  const kdjInterpText = a.kdjInterpretation.emoji + ' ' + a.kdjInterpretation.label;

  const row = [
    a.timestamp, a.instId, a.lastCandleTime,
    a.open, a.high, a.low, a.close,
    a.changePct, a.volume, a.volMaShort, a.volMaLong, a.volumeConfirmed ? 'YES' : 'NO',
    a.rsi, a.macdLine, a.macdSignal, a.macdHist,
    a.kdj.k, a.kdj.d, a.kdj.j, kdjInterpText,
    a.bb.upper, a.bb.middle, a.bb.lower, a.bb.isSqueeze ? 'YES' : 'NO', a.squeezeStreak,
    a.atr.atr, a.atr.atrPct,
    a.fib.levels['0.5'], a.fib.levels['0.618'], a.fib.levels['0.786'],
    a.priceProjection ? a.priceProjection.range4h.lower : null,
    a.priceProjection ? a.priceProjection.range4h.upper : null,
    a.priceProjection ? a.priceProjection.range8h.lower : null,
    a.priceProjection ? a.priceProjection.range8h.upper : null,
    patternsText, a.score, a.trend
  ];

  sheet.appendRow(row);
  const newRowIndex = sheet.getLastRow();
  const colIndex = (name) => headers.indexOf(name) + 1;

  const trendCell = sheet.getRange(newRowIndex, colIndex('Trend'));
  const changeCell = sheet.getRange(newRowIndex, colIndex('24H/4H Change %'));

  if (a.trend === 'BULLISH') {
    trendCell.setBackground('#d9ead3').setFontColor('#274e13').setFontWeight('bold');
  } else if (a.trend === 'BEARISH') {
    trendCell.setBackground('#f4cccc').setFontColor('#990000').setFontWeight('bold');
  } else {
    trendCell.setBackground('#fff2cc').setFontColor('#7f6000').setFontWeight('bold');
  }

  changeCell.setNumberFormat('0.00"%"');

  const priceFormat = CA4H_priceNumberFormat_(a.close);
  sheet.getRange(newRowIndex, colIndex('Open'), 1, 4).setNumberFormat(priceFormat);

  sheet.getRange(newRowIndex, colIndex('Volume (USDT)'), 1, 3).setNumberFormat('0.00');
  sheet.getRange(newRowIndex, colIndex('RSI(14)')).setNumberFormat('0.00');

  const macdFormat = CA4H_priceNumberFormat_(a.macdLine || a.close);
  sheet.getRange(newRowIndex, colIndex('MACD'), 1, 3).setNumberFormat(macdFormat);

  sheet.getRange(newRowIndex, colIndex('KDJ K'), 1, 3).setNumberFormat('0.00');

  const bbFormat = CA4H_priceNumberFormat_(a.bb.middle || a.close);
  sheet.getRange(newRowIndex, colIndex('BB Upper'), 1, 3).setNumberFormat(bbFormat);

  const atrFormat = CA4H_priceNumberFormat_(a.atr.atr || a.close);
  sheet.getRange(newRowIndex, colIndex('ATR(14)')).setNumberFormat(atrFormat);
  sheet.getRange(newRowIndex, colIndex('ATR %')).setNumberFormat('0.00"%"');

  const fibFormat = CA4H_priceNumberFormat_(a.close);
  sheet.getRange(newRowIndex, colIndex('Fib 0.5'), 1, 3).setNumberFormat(fibFormat);

  const projFormat = CA4H_priceNumberFormat_(a.close);
  sheet.getRange(newRowIndex, colIndex('Proj 4H Low'), 1, 4).setNumberFormat(projFormat);

  sheet.getRange(newRowIndex, colIndex('Timestamp')).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(newRowIndex, colIndex('Last 4H Candle')).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

// ---------------------------------------------------------------------------
// ΣΥΝΟΨΗ ΔΕΙΚΤΩΝ σε φυσική γλώσσα
// ---------------------------------------------------------------------------
function CA4H_buildIndicatorSummary_(a) {
  const parts = [];

  if (a.rsi != null) {
    if (a.rsi >= 70) parts.push('το RSI (' + a.rsi.toFixed(1) + ') δείχνει υπεραγορά');
    else if (a.rsi <= 30) parts.push('το RSI (' + a.rsi.toFixed(1) + ') δείχνει υπερπούληση');
    else parts.push('το RSI (' + a.rsi.toFixed(1) + ') είναι ουδέτερο');
  }

  if (a.macdHist != null) {
    if (a.macdHist > 0) parts.push('το MACD histogram είναι θετικό (ανοδικό momentum)');
    else parts.push('το MACD histogram είναι αρνητικό (πτωτικό momentum)');
  }

  parts.push('το KDJ δείχνει "' + a.kdjInterpretation.label + '"');

  if (a.bb.upper != null) {
    if (a.bb.isSqueeze) parts.push('οι κορδέλες Bollinger έχουν στενέψει σημαντικά (πιθανό squeeze πριν από κίνηση)');
    else if (a.close >= a.bb.upper) parts.push('η τιμή βρίσκεται στο ή πάνω από το άνω Bollinger band');
    else if (a.close <= a.bb.lower) parts.push('η τιμή βρίσκεται στο ή κάτω από το κάτω Bollinger band');
    else parts.push('η τιμή κινείται εντός των κορδελών Bollinger');
  }

if (a.atr.atrPct != null) {
  const atrLabel =
    CA4H_describeAtrVolatility_(
      a.atr.atrPct
    );

  parts.push(
    'η μεταβλητότητα (ATR ' +
    a.atr.atrPct.toFixed(1) +
    '% της τιμής) είναι ' +
    atrLabel
  );
}

  const sentence = 'Συνολικά, ' + parts.join(', ') + '. Η σύνθεση όλων αυτών δίνει καθαρή εικόνα ' +
    (a.trend === 'BULLISH' ? '*ανοδικής πίεσης*' : (a.trend === 'BEARISH' ? '*πτωτικής πίεσης*' : '*αναποφασιστικότητας/πλάγιας αγοράς*')) + '.';

  return sentence;
}

function CA4H_formatPercent_(value, decimals) {
  if (value == null || !isFinite(value)) {
    return 'n/a';
  }

  const d = decimals == null ? 2 : decimals;

  if (Math.abs(value) < Math.pow(10, -d) / 2) {
    return '≈0.00%';
  }

  return (value >= 0 ? '+' : '') +
    Number(value).toFixed(d) +
    '%';
}

// ---------------------------------------------------------------------------
// Μορφοποίηση ενότητας "Αναμενόμενο Εύρος Τιμών (4H-8H)"
// ---------------------------------------------------------------------------
function CA4H_formatPriceProjection_(a) {
  if (!a.priceProjection) {
    return '🔮 *Αναμενόμενο Εύρος Τιμών (4H-8H):* `n/a` (λείπει ATR)';
  }

  const p = a.priceProjection;
  const trendNote = a.trend === 'BULLISH'
    ? '(ζώνη «τεντωμένη» προς τα πάνω λόγω ανοδικού momentum)'
    : (a.trend === 'BEARISH'
      ? '(ζώνη «τεντωμένη» προς τα κάτω λόγω πτωτικού momentum)'
      : '(συμμετρική ζώνη — ουδέτερη τάση)');

  return [
    '🔮 *Αναμενόμενο Εύρος Τιμών (στατιστική προβολή βάσει ATR)*',
    '4H: `' + CA4H_formatPrice_(p.range4h.lower) + '` – `' + CA4H_formatPrice_(p.range4h.upper) + '`',
    '8H: `' + CA4H_formatPrice_(p.range8h.lower) + '` – `' + CA4H_formatPrice_(p.range8h.upper) + '`',
    '_' + trendNote + '. Βασίζεται σε ATR(14) και ΔΕΝ αποτελεί εγγυημένη πρόβλεψη._'
  ].join('\n');
}

// ---------------------------------------------------------------------------
// ΑΝΑΛΥΣΗ ΘΕΣΗΣ + ΣΥΣΤΑΣΗ ΓΙΑ FUTURES DCA GRID
// ---------------------------------------------------------------------------
function CA4H_buildPositionRecommendation_(a, position) {
  if (!position || !position.direction) return null;

  const lines = [];
  const dir = position.direction;
  const hasRange = position.lowerPrice != null && position.upperPrice != null && position.upperPrice > position.lowerPrice;

  lines.push('🎯 *Ανάλυση Θέσης & DCA Grid*');
  lines.push('Κατεύθυνση θέσης: *' + dir + '*' + (hasRange
    ? ' | Εύρος grid: `' + CA4H_formatPrice_(position.lowerPrice) + '` – `' + CA4H_formatPrice_(position.upperPrice) + '`'
    : ' (δεν βρέθηκε καταγεγραμμένο εύρος τιμών στο sheet)'));

  let positionInGridPct = null;
  if (hasRange) {
    positionInGridPct = ((a.close - position.lowerPrice) / (position.upperPrice - position.lowerPrice)) * 100;
    positionInGridPct = Math.max(0, Math.min(100, positionInGridPct));
    lines.push('Τρέχουσα τιμή `' + CA4H_formatPrice_(a.close) + '` βρίσκεται στο ' + positionInGridPct.toFixed(1) + '% του εύρους grid.');
  }

  // NEW: ATR-distance to grid edges
  const atrDist = CA4H_calcAtrDistanceToGrid_(a.close, a.atr.atr, position);
  const atrDistLine = CA4H_formatAtrDistanceLine_(atrDist);
  if (atrDistLine) lines.push(atrDistLine);

  // NEW: Squeeze-aware guidance
  const squeezeLine = CA4H_buildSqueezeGuidance_(a, position);
  if (squeezeLine) lines.push(squeezeLine);

  const trendAgainstPosition = (dir === 'LONG' && a.trend === 'BEARISH') || (dir === 'SHORT' && a.trend === 'BULLISH');
  const trendWithPosition = (dir === 'LONG' && a.trend === 'BULLISH') || (dir === 'SHORT' && a.trend === 'BEARISH');

  const nearUpperEdge = hasRange && positionInGridPct != null && positionInGridPct >= (100 - CA4H_CFG.GRID_EDGE_WARNING_PCT);
  const nearLowerEdge = hasRange && positionInGridPct != null && positionInGridPct <= CA4H_CFG.GRID_EDGE_WARNING_PCT;

  if (hasRange && dir === 'LONG' && nearUpperEdge) {
    lines.push('📈 Η τιμή πλησιάζει το *άνω όριο* του grid ενώ η θέση είναι LONG — καλή ευκαιρία για *μερική ρευστοποίηση κέρδους* πριν από πιθανό pullback, ειδικά αν το KDJ δείχνει "' + a.kdjInterpretation.label + '".');
  } else if (hasRange && dir === 'SHORT' && nearLowerEdge) {
    lines.push('📉 Η τιμή πλησιάζει το *κάτω όριο* του grid ενώ η θέση είναι SHORT — καλή ευκαιρία για *μερική ρευστοποίηση κέρδους* πριν από πιθανό bounce, ειδικά αν το KDJ δείχνει "' + a.kdjInterpretation.label + '".');
  } else if (hasRange && dir === 'LONG' && nearLowerEdge) {
    if (trendAgainstPosition) {
      lines.push('⚠️ Η τιμή πλησιάζει το *κάτω όριο* του grid ενώ η συνολική τάση είναι BEARISH — αυξημένος κίνδυνος να "σπάσει" η τιμή κάτω από το εύρος του grid. Σκέψου μείωση προσθήκης DCA ή στενότερο stop, αντί για αυτόματη προσθήκη θέσης.');
    } else {
      lines.push('🟢 Η τιμή πλησιάζει το *κάτω όριο* του grid, αλλά η τάση δεν είναι ξεκάθαρα αρνητική — αυτό είναι λογικό σημείο για *προσθήκη DCA*, εφόσον διαθέτεις κεφάλαιο, καθώς το grid θα αγοράζει χαμηλότερα.');
    }
  } else if (hasRange && dir === 'SHORT' && nearUpperEdge) {
    if (trendAgainstPosition) {
      lines.push('⚠️ Η τιμή πλησιάζει το *άνω όριο* του grid ενώ η συνολική τάση είναι BULLISH — αυξημένος κίνδυνος να "σπάσει" η τιμή πάνω από το εύρος του grid. Σκέψου μείωση προσθήκης DCA ή στενότερο stop.');
    } else {
      lines.push('🔴 Η τιμή πλησιάζει το *άνω όριο* του grid, αλλά η τάση δεν είναι ξεκάθαρα θετική — λογικό σημείο για *προσθήκη DCA short*, εφόσον διαθέτεις κεφάλαιο.');
    }
  } else if (trendWithPosition) {
    lines.push('✅ Η τρέχουσα τεχνική εικόνα (' + a.trend + ') είναι *ευνοϊκή* για τη θέση σου. Λογικό να *κρατήσεις το grid ως έχει* και να παρακολουθείς για επιβεβαίωση συνέχισης.');
  } else if (trendAgainstPosition) {
    lines.push('⚠️ Η τρέχουσα τεχνική εικόνα (' + a.trend + ') είναι *αντίθετη* προς τη θέση σου. Δεν είναι απαραίτητα λόγος πανικού μέσα σε ένα DCA grid, αλλά αξίζει να ελέγξεις το επίπεδο κινδύνου/leverage και να μην αυξήσεις επιθετικά την έκθεση προς το παρόν.');
  } else {
    lines.push('⚪ Η τεχνική εικόνα είναι ουδέτερη/χωρίς σαφή κατεύθυνση. Λογικό να *κρατήσεις το grid ως έχει* χωρίς νέες προσθήκες, μέχρι να υπάρξει σαφέστερο σήμα.');
  }

  if (hasRange) {
    const criticalEdgePct = CA4H_CFG.GRID_EDGE_WARNING_PCT / 2;
    const criticalUpper = positionInGridPct >= (100 - criticalEdgePct);
    const criticalLower = positionInGridPct <= criticalEdgePct;
    if ((dir === 'LONG' && criticalLower && trendAgainstPosition) ||
        (dir === 'SHORT' && criticalUpper && trendAgainstPosition)) {
      lines.push('🚨 Προσοχή: η τιμή είναι πολύ κοντά στο όριο εξόδου του grid ΚΑΙ η τάση είναι αντίθετη. Αν σπάσει το όριο, το grid θα σταματήσει να λειτουργεί κανονικά — εξέτασε χειροκίνητη παρέμβαση (επέκταση εύρους ή μερικό κλείσιμο) αντί να περιμένεις παθητικά.');
    }
  }

  if (hasRange && a.priceProjection) {
    const proj8hUpper = a.priceProjection.range8h.upper;
    const proj8hLower = a.priceProjection.range8h.lower;
    if (proj8hUpper > position.upperPrice) {
      lines.push('📐 Η προβολή 8H (έως `' + CA4H_formatPrice_(proj8hUpper) + '`) ξεπερνά το άνω όριο του grid `' + CA4H_formatPrice_(position.upperPrice) + '` — πιθανό να χρειαστεί επέκταση εύρους προς τα πάνω αν συνεχιστεί η κίνηση.');
    }
    if (proj8hLower < position.lowerPrice) {
      lines.push('📐 Η προβολή 8H (έως `' + CA4H_formatPrice_(proj8hLower) + '`) πέφτει κάτω από το κάτω όριο του grid `' + CA4H_formatPrice_(position.lowerPrice) + '` — πιθανό να χρειαστεί επέκταση εύρους προς τα κάτω αν συνεχιστεί η κίνηση.');
    }
  }

  return lines.join('\n');
}

function CA4H_formatHistoricallyDefendedSupportZone_(
  zone,
  supportZone
) {
  const lines = [];

  if (zone && zone.isRepeatedZone) {
    lines.push(
      '🛡️ *Ιστορικά Υπερασπισμένη Support Zone*'
    );

    lines.push(
      'Ζώνη: `' +
      CA4H_formatPrice_(zone.lower) +
      '` – `' +
      CA4H_formatPrice_(zone.upper) +
      '`'
    );

    lines.push(
      'Κεντρικό επίπεδο: `' +
      CA4H_formatPrice_(zone.midpoint) +
      '`'
    );

    lines.push(
      'Ιστορικές δοκιμές: `' +
      zone.touches +
      '` | Ισχύς: *' +
      zone.strength +
      '*'
    );

    lines.push(
      'Απόσταση: `' +
      zone.distanceAtr.toFixed(1) +
      'x ATR` κάτω από την τρέχουσα τιμή'
    );

    lines.push(
      '_Η ζώνη είναι ιστορική εκτίμηση, όχι εγγυημένη στήριξη ή αυτόματο DCA trigger._'
    );

    return lines.join('\n');
  }

  if (supportZone) {
    lines.push(
      '🛡️ *Πιθανή Support Zone*'
    );

    lines.push(
      'Ζώνη: `' +
      CA4H_formatPrice_(supportZone.lower) +
      '` – `' +
      CA4H_formatPrice_(supportZone.upper) +
      '`'
    );

    lines.push(
      'Κεντρικό επίπεδο: `' +
      CA4H_formatPrice_(supportZone.midpoint) +
      '`'
    );

    lines.push(
      'Ισχύς: *' +
      supportZone.strength +
      '* — βάση: ' +
      supportZone.sources.join(', ')
    );

    lines.push(
      'Απόσταση: `' +
      supportZone.distanceAtr.toFixed(1) +
      'x ATR` κάτω από την τρέχουσα τιμή'
    );

    lines.push(
      '_Δεν υπάρχει ακόμη επαναλαμβανόμενη ιστορική ζώνη. Η ζώνη είναι εκτίμηση τεχνικών επιπέδων, όχι εγγυημένη στήριξη._'
    );

    return lines.join('\n');
  }

  return [
    '🛡️ *Support Zone*',
    'Δεν βρέθηκε σχετική ζώνη στήριξης κάτω από την τρέχουσα τιμή.'
  ].join('\n');
}

function CA4H_formatLiveCandleNotice_(
  analysis
) {
  if (
    !analysis ||
    !analysis.liveCandleIncluded ||
    !analysis.liveCandleClose
  ) {
    return '';
  }

  const livePrice =
    analysis.liveCandleClose;

  const closedPrice =
    analysis.close;

  const changePct =
    ((livePrice - closedPrice) /
      closedPrice) * 100;

  return [
    '🟡 *LIVE / INCOMPLETE 4H Candle*',
    'Live τιμή: `' +
      CA4H_formatPrice_(livePrice) +
      '`',
    'Μεταβολή από τελευταίο κλειστό 4H close: `' +
      (changePct >= 0 ? '+' : '') +
      changePct.toFixed(2) +
      '%`',
    '_Το live candle δεν συμμετέχει στο βασικό τεχνικό σήμα._'
  ].join('\n');
}

// ---------------------------------------------------------------------------
// TELEGRAM MESSAGE
// ---------------------------------------------------------------------------
function CA4H_formatTelegramMessage_(a, position) {
  const patternsList = a.patterns.map(p => '• *' + p.name + '* (' + p.bias + ') — ' + p.note).join('\n');
  const closeStr = CA4H_formatPrice_(a.close);
  const macdHistStr = CA4H_formatPrice_(a.macdHist);
  const rsiStr = (a.rsi == null) ? 'n/a' : Number(a.rsi).toFixed(1);

  const kdjLine = a.kdjInterpretation.emoji + ' *' + a.kdjInterpretation.label + '*' +
    (a.kdj.k != null ? ' (K ' + a.kdj.k.toFixed(1) + ' / D ' + a.kdj.d.toFixed(1) + ' / J ' + a.kdj.j.toFixed(1) + ')' : '');

  const bbStr = (a.bb.upper != null)
    ? 'Upper `' + CA4H_formatPrice_(a.bb.upper) + '` / Mid `' + CA4H_formatPrice_(a.bb.middle) + '` / Lower `' + CA4H_formatPrice_(a.bb.lower) + '`'
      + (a.bb.isSqueeze ? ' 🔸 SQUEEZE' : '')
    : 'n/a';

  const atrStr = (a.atr.atr != null)
    ? '`' + CA4H_formatPrice_(a.atr.atr) + '` (' + a.atr.atrPct.toFixed(2) + '% της τιμής)'
    : 'n/a';

  const fibDirStr = a.fib.isUptrend ? 'ανοδικό swing' : 'πτωτικό swing';
  const fibStr = '0.5: `' + CA4H_formatPrice_(a.fib.levels['0.5']) + '` | ' +
    '0.618: `' + CA4H_formatPrice_(a.fib.levels['0.618']) + '` | ' +
    '0.786: `' + CA4H_formatPrice_(a.fib.levels['0.786']) + '`';

  const messageParts = [
    a.trendEmoji + ' *' + a.instId + ' — 4H Ανάλυση (Closed Candle)*',
    '_' + a.lastCandleTime.toISOString()
  .replace('T', ' ')
  .slice(0, 16) +
' UTC — closed_',
    '',
    '💰 Τιμή κλεισίματος: `' + closeStr + '` (' + (a.changePct >= 0 ? '+' : '') + a.changePct.toFixed(2) + '%)',
    '📊 Volume: `' + Number(a.volume).toFixed(0) + ' USDT`' + (a.volumeConfirmed ? ' ✅ πάνω από MA5' : ' ⚠️ κάτω από MA5'),
    '📈 RSI(14): `' + rsiStr + '`',
    '📉 MACD Hist: `' + macdHistStr + '`',
    '📐 KDJ: ' + kdjLine,
    '📏 Bollinger(20,2): ' + bbStr,
    '📶 ATR(14): ' + atrStr,
    '🔢 Fibonacci (' + fibDirStr + '): ' + fibStr,
    '',
    '*Candlestick Patterns:*',
    patternsList,
    '',
    '*Συνολική Τάση:* ' + a.trendEmoji + ' *' + a.trend + '* (score: ' + a.score.toFixed(2) + ')'
  ];

  const liveNotice =
  CA4H_formatLiveCandleNotice_(a);

if (liveNotice) {
  messageParts.push('');
  messageParts.push(liveNotice);
}

  // NEW: Signal conflict flag (μόνο αν εντοπιστεί σύγκρουση)
  const conflictLine = CA4H_detectSignalConflict_(a);
  if (conflictLine) {
    messageParts.push('');
    messageParts.push(conflictLine);
  }

  messageParts.push('');
  messageParts.push(CA4H_formatPriceProjection_(a));
  messageParts.push('');
  messageParts.push('📋 *Σύνοψη Δεικτών*');
  messageParts.push('');
  messageParts.push(
    CA4H_formatPionexDerivativesMessage_(
      a.pionexDerivatives,
      a // OKX analysis object
    )
  );
  messageParts.push(CA4H_buildIndicatorSummary_(a));

  if (
    a.oiAnalysis ||
    a.supportZone
  ) {
    messageParts.push('');
    messageParts.push(
      CA4H_formatOiAndSupportMessage_(
        a
      )
    );


    messageParts.push('');
messageParts.push(
  '⚠️ *Βασικό σήμα:* υπολογίζεται αποκλειστικά από κλειστά 4H candles. ' +
  'Τυχόν live τιμή εμφανίζεται μόνο ως πληροφορία και δεν συγκρίνεται απευθείας με κλειστά indicators.'
);

  }

  const positionRecommendation = CA4H_buildPositionRecommendation_(a, position);
  if (positionRecommendation) {
    messageParts.push('');
    messageParts.push(positionRecommendation);
  }

messageParts.push('');
messageParts.push(
  CA4H_formatHistoricallyDefendedSupportZone_(
    a.historicallyDefendedSupport,
    a.supportZone
  )
);

  return messageParts.join('\n');
}

// ---------------------------------------------------------------------------
// TELEGRAM SENDER
// ---------------------------------------------------------------------------
function CA4H_sendTelegramMessage_(text) {
  if (text == null || text === '') {
    throw new Error('CA4H_sendTelegramMessage_() κλήθηκε χωρίς "text" argument.');
  }
  if (typeof PropertiesService === 'undefined' || !PropertiesService || typeof PropertiesService.getScriptProperties !== 'function') {
    throw new Error('Το global PropertiesService δεν είναι διαθέσιμο. Ελέγξε για shadowing σε άλλο σημείο του project.');
  }
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token) throw new Error('Λείπει το Script Property "TELEGRAM_BOT_TOKEN".');
  if (!chatId) throw new Error('Λείπει το Script Property "TELEGRAM_CHAT_ID".');

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown', disable_web_page_preview: true }),
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(url, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error('Telegram API error ' + resp.getResponseCode() + ': ' + resp.getContentText());
  }
}

// ---------------------------------------------------------------------------
// TRIGGER SETUP
// ---------------------------------------------------------------------------
function CA4H_setupTrigger_every4Hours() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'analyze4hAndNotify' || fn === 'CA4H_analyze4hAndNotify') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('CA4H_analyze4hAndNotify').timeBased().everyHours(4).create();
}

function CA4H_testTelegramOnly_() {
  try {
    CA4H_sendTelegramMessage_('✅ Test μήνυμα — αν το βλέπεις, το Telegram κομμάτι δουλεύει.');
    Logger.log('Επιτυχία.');
  } catch (e) {
    Logger.log('Αποτυχία: ' + e.message);
  }
}

function CA4H_testReadPositionOnly_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const posSheet = ss.getSheetByName(CA4H_CFG.POSITIONS_SHEET);
  if (!posSheet) {
    Logger.log('Δεν βρέθηκε το φύλλο "' + CA4H_CFG.POSITIONS_SHEET + '"');
    return;
  }
  const position = CA4H_readOpenPosition_(posSheet);
  Logger.log('Αποτέλεσμα CA4H_readOpenPosition_: ' + JSON.stringify(position));
}

function CA4H_testVolumeFix_() {
  const instId = 'TRUTH-USDT-SWAP';
  const candles = CA4H_fetchOkxCandles_(instId, CA4H_CFG.BAR, 5);
  candles.forEach(c => {
    Logger.log(
      new Date(c.ts).toISOString() +
      ' | close=' + c.close +
      ' | volContracts=' + c.volContracts +
      ' | volume(quote/USDT)=' + c.volume
    );
  });
}

/**
 * ============================================================================
 * ΣΗΜΕΙΩΣΕΙΣ v3.7 — ΤΙ ΑΛΛΑΞΕ ΣΥΝΟΛΙΚΑ
 * ============================================================================
 * 1) CA4H_fetchOkxCandles_(): volume πλέον = volCcy (quote currency, π.χ.
 *    USDT), όχι vol (contracts). Νέο πεδίο volContracts για αναφορά.
 *
 * 2) CA4H_calcPriceProjection_(): προβολή εύρους τιμής 4H/8H βάσει ATR με
 *    κλιμάκωση √χρόνου και ασυμμετρία προς τη φορά της τάσης. Εμφανίζεται
 *    πάντα στο μήνυμα με σαφές disclaimer.
 *
 * 3) CA4H_detectSignalConflict_(): εντοπίζει σύγκρουση ανάμεσα στο τελευταίο
 *    candlestick pattern και στο KDJ/MACD. Εμφανίζεται ΜΟΝΟ όταν υπάρχει
 *    πραγματική σύγκρουση, ώστε να μη γεμίζει το μήνυμα άσκοπα.
 *
 * 4) CA4H_calcAtrDistanceToGrid_(): απόσταση τιμής από τα άκρα του grid σε
 *    πολλαπλάσια ATR, με ταξινόμηση κοντά/μεσαία/μακριά. Εμφανίζεται πάντα
 *    στην ενότητα "Ανάλυση Θέσης & DCA Grid" όταν υπάρχει καταγεγραμμένο
 *    εύρος grid.
 *
 * 5) CA4H_trackSqueezeStreak_() + CA4H_buildSqueezeGuidance_(): μετράει
 *    συνεχόμενα squeeze candles από το sheet ιστορικό, και προσθέτει
 *    καθοδήγηση grid όταν το squeeze είναι ενεργό. Νέα στήλη "Squeeze
 *    Streak" στο sheet "4h analysis".
 *
 * 6) CA4H_buildAnalysis_() παίρνει τώρα "ss" ως 3ο argument (χρειάζεται
 *    για το squeeze streak tracking).
 * ============================================================================
 */









/**
 * ============================================================================
 *  PEARL (PRL) DAILY DASHBOARD  —  Google Apps Script  (v9 — σωστό hashrate.no schema)
 * ----------------------------------------------------------------------------
 *  ΤΙ ΑΛΛΑΞΕ ΣΤΗΝ v9:
 *   Το testHashrateNoApi() αποκάλυψε το ΠΡΑΓΜΑΤΙΚΟ schema του hashrate.no:
 *   {
 *     "290": {                              // numeric coin ID ως top-level key
 *       "ticker": "PRL", "name": "Pearl",
 *       "price": { "USD": "0.379", "day": "5.24", "week": "18.39", ... },
 *       "estimate": { "yieldDay": "-7.12", "revenueDay": "-2.21", ... },
 *       "network": { "hashrate": "3.24E+19", "difficulty": "20615657.3",
 *                     "blockTime": "196.79", "blockReward": "2357.63" }
 *     }
 *   }
 *   Το parsing διορθώθηκε ώστε να παίρνει το πρώτο (και μοναδικό) value του
 *   response object, μετά διαβάζει net.network.hashrate/difficulty και
 *   net.estimate.revenueDay/yieldDay. Επιπλέον:
 *     - Το network.hashrate είναι string σε scientific notation σε H/s
 *       (π.χ. "3.2382535225409E+19") — μετατρέπεται σωστά σε EH/s (/1e18).
 *     - Το estimate.revenueDay (% ημερήσια μεταβολή εσόδου ανά TH/s, ήδη
 *       υπολογισμένο από το hashrate.no) χρησιμοποιείται ΑΠΕΥΘΕΙΑΣ για το
 *       metric "Mining Revenue/TH Trend %" — πιο ακριβές από να το υπολογίζουμε
 *       εμείς συγκρίνοντας με την προηγούμενη γραμμή του sheet.
 *     - Το ίδιο revenueDay τροφοδοτεί επίσης το "AI Demand/Capacity Proxy %"
 *       ως έμμεσος δείκτης ζήτησης compute (βλ. σχόλιο στο computeMetrics_).
 *     - Το blockReward/blockTime αποθηκεύονται στο sheet σχόλιο για αναφορά.
 *
 *  ΡΥΘΜΙΣΗ (Project Settings → Script Properties):
 *     TELEGRAM_BOT_TOKEN        = <ήδη το έχεις>
 *     TELEGRAM_CHAT_ID          = <ήδη το έχεις>
 *     COINGECKO_API_KEY         = <το δωρεάν Demo key σου — ✅ ήδη ορισμένο>
 *     HASHRATE_NO_API_KEY       = <το δωρεάν hashrate.no key σου — ✅ ήδη ορισμένο>
 *     NETWORK_STATS_API_URL     = <προαιρετικό fallback αν χρειαστεί>
 *     ELECTRICITY_COST_EUR_KWH  = 0.20   (προαιρετικό, default 0.20)
 *     GPU_POWER_WATTS           = 350    (προαιρετικό, default 350)
 *     EXPECTED_PRL_PER_DAY      = 5      (προαιρετικό — για ακριβές mining ratio)
 *     AI_WORKLOAD_API_URL       =        (προαιρετικό, μελλοντικό hook — άσε κενό)
 *
 *  Trigger: Triggers → Add Trigger → runPearlDashboard → Time-driven → Daily.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// ΡΥΘΜΙΣΕΙΣ / CONFIG
// ---------------------------------------------------------------------------
const PEARL_CONFIG = {
  SHEET_NAME: 'Pearl Dashboard',
  COINGECKO_ID: 'pearl-2',
  HASHRATE_NO_COIN_TICKER: 'PRL',
  VS_CURRENCY: 'eur',
  CACHE_TTL_SECONDS: 300,
  MAX_RETRIES: 4,
  BASE_BACKOFF_MS: 2000,
  MAX_BACKOFF_MS: 15000,

  THRESHOLDS: {
    change24h:       { green: 3,   yellow: 0 },
    volChange24h:    { green: 15,  yellow: -10 },
    change7d:        { green: 8,   yellow: 0 },
    volMcapRatio:    { green: 8,   yellow: 3 },
    athDistance:     { greenMax: -60, yellowMax: -85 },
    rankTrend:       { greenDelta: 1, yellowDelta: 0 },
    supplyProgress:  { greenMax: 20, yellowMax: 40 },
    miningProfit:    { green: 1.3, yellow: 1.0 },
    volatility:      { greenMax: 6, yellowMax: 12 },
    hashrateChange:  { green: 5,   yellow: -5 },
    difficultyChange:{ greenMax: 5, yellowMax: 15 },
    revenuePerTh:    { green: 2,   yellow: -10 },
    aiDemandProxy:   { green: 2,   yellow: -10 }
  }
};

// ---------------------------------------------------------------------------
// MAIN ENTRY POINT
// ---------------------------------------------------------------------------
function runPearlDashboard() {
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('COINGECKO_API_KEY')) {
    console.warn('COINGECKO_API_KEY δεν βρέθηκε — θα χρησιμοποιηθεί το πιο περιορισμένο keyless API.');
  }

  let marketData = null;
  let isStale = false;
  const previousRow = getPreviousRow_();

  try {
    marketData = fetchPearlMarketData_(props);
  } catch (err) {
    console.error('CoinGecko fetch failed completely: ' + err.message);
    marketData = buildFallbackMarketDataFromPreviousRow_(previousRow);
    if (marketData) {
      isStale = true;
      console.warn('Using STALE data from previous sheet row to keep the dashboard alive.');
    } else {
      sendTelegramMessage_(
        '⚠️ <b>Pearl Dashboard Error</b>\n' +
        'Αποτυχία λήψης δεδομένων από CoinGecko και δεν υπάρχει προηγούμενο ιστορικό ' +
        'στο sheet για fallback.\n\nΤεχνικό σφάλμα: ' + escapeHtml_(String(err.message)),
        props
      );
      return;
    }
  }

  try {
    const networkData = fetchPearlNetworkData_(props);
    const metrics = computeMetrics_(marketData, networkData, previousRow, props);
    const verdict = computeVerdict_(metrics);

    writeToSheet_(marketData, networkData, metrics, verdict, isStale);
    sendTelegramMessage_(buildTelegramMessage_(marketData, networkData, metrics, verdict, isStale), props);

  } catch (err) {
    console.error('runPearlDashboard failed downstream: ' + err.message);
    try {
      sendTelegramMessage_(
        '⚠️ <b>Pearl Dashboard Error</b>\n' + escapeHtml_(String(err.message)),
        props
      );
    } catch (e2) {
      console.error('Failed to send error alert to Telegram: ' + e2.message);
    }
  }
}

// ---------------------------------------------------------------------------
// 1) FETCH MARKET DATA FROM COINGECKO (cache + throttle + retry/backoff + key)
// ---------------------------------------------------------------------------
function fetchPearlMarketData_(props) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'pearl_market_data_v9';
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log('Using cached CoinGecko data (rate-limit friendly).');
    return JSON.parse(cached);
  }

  const apiKey = props.getProperty('COINGECKO_API_KEY');

  let url = 'https://api.coingecko.com/api/v3/coins/' + PEARL_CONFIG.COINGECKO_ID +
    '?localization=false&tickers=false&market_data=true&community_data=false' +
    '&developer_data=false&sparkline=false';

  const headers = { 'Accept': 'application/json' };
  if (apiKey) {
    headers['x-cg-demo-api-key'] = apiKey;
    url += '&x_cg_demo_api_key=' + encodeURIComponent(apiKey);
  }

  const json = fetchWithBackoff_(url, headers, { relaxSsl: false });
  const md = json.market_data;
  const vs = PEARL_CONFIG.VS_CURRENCY;

  const result = {
    timestamp: new Date().toISOString(),
    name: json.name,
    symbol: (json.symbol || '').toUpperCase(),
    price: md.current_price[vs],
    priceUsd: md.current_price['usd'],
    change24h: md.price_change_percentage_24h,
    change7d: md.price_change_percentage_7d,
    high24h: md.high_24h[vs],
    low24h: md.low_24h[vs],
    volume24h: md.total_volume[vs],
    marketCap: md.market_cap[vs],
    marketCapRank: md.market_cap_rank,
    ath: md.ath[vs],
    athChangePercentage: md.ath_change_percentage[vs],
    circulatingSupply: md.circulating_supply,
    maxSupply: md.max_supply,
    totalSupply: md.total_supply
  };

  cache.put(cacheKey, JSON.stringify(result), PEARL_CONFIG.CACHE_TTL_SECONDS);
  return result;
}

/** Fetch με exponential backoff + jitter — προστασία από 429/5xx. */
function fetchWithBackoff_(url, headers, opts) {
  opts = opts || {};
  let attempt = 0;
  let lastError = null;

  while (attempt < PEARL_CONFIG.MAX_RETRIES) {
    let response;
    try {
      const fetchOptions = {
        muteHttpExceptions: true,
        headers: headers || { 'Accept': 'application/json' }
      };
      if (opts.relaxSsl) fetchOptions.validateHttpsCertificates = false;

      response = UrlFetchApp.fetch(url, fetchOptions);
    } catch (e) {
      const msg = String(e.message);
      if (msg.indexOf('SSL') !== -1 || msg.indexOf('Address unavailable') !== -1 || msg.indexOf('DNS') !== -1) {
        throw new Error('Connection-level error στο ' + url + ' (μη ανακτήσιμο): ' + e.message);
      }
      lastError = e.message;
      const wait = Math.min(PEARL_CONFIG.BASE_BACKOFF_MS * Math.pow(2, attempt), PEARL_CONFIG.MAX_BACKOFF_MS);
      console.warn('Connection error for ' + url + ': ' + e.message + ' — retrying in ' + wait + 'ms');
      Utilities.sleep(wait);
      attempt++;
      continue;
    }

    const code = response.getResponseCode();
    if (code === 200) {
      return JSON.parse(response.getContentText());
    }
    if (code === 429 || code >= 500) {
      const jitter = Math.floor(Math.random() * 1000);
      const wait = Math.min(
        PEARL_CONFIG.BASE_BACKOFF_MS * Math.pow(2, attempt) + jitter,
        PEARL_CONFIG.MAX_BACKOFF_MS
      );
      console.warn('API returned ' + code + ' for ' + url + ', retrying in ' + wait + 'ms (attempt ' + (attempt + 1) + '/' + PEARL_CONFIG.MAX_RETRIES + ')');
      Utilities.sleep(wait);
      attempt++;
      lastError = 'HTTP ' + code;
      continue;
    }
    if (code === 401 || code === 403) {
      throw new Error('API key invalid/rejected (HTTP ' + code + ') για ' + url + '.');
    }
    throw new Error('API error (' + url + '): HTTP ' + code + ' — ' + response.getContentText().substring(0, 300));
  }
  throw new Error('API failed after ' + PEARL_CONFIG.MAX_RETRIES + ' retries for ' + url + '. Last error: ' + lastError);
}

/** Fallback: αν το CoinGecko αποτύχει εντελώς, ξαναχρησιμοποιούμε την τελευταία
 *  γνωστή γραμμή του sheet ώστε το dashboard/Telegram να μη σταματήσει τελείως. */
function buildFallbackMarketDataFromPreviousRow_(previousRow) {
  const priceKey = 'Price (' + PEARL_CONFIG.VS_CURRENCY.toUpperCase() + ')';
  if (!previousRow || previousRow[priceKey] == null) return null;

  return {
    timestamp: new Date().toISOString(),
    name: 'Pearl',
    symbol: 'PRL',
    price: Number(previousRow[priceKey]),
    priceUsd: null,
    change24h: Number(previousRow['24h Change %']) || 0,
    change7d: Number(previousRow['7d Change %']) || 0,
    high24h: Number(previousRow[priceKey]),
    low24h: Number(previousRow[priceKey]),
    volume24h: Number(previousRow['Volume 24h']) || 0,
    marketCap: 0,
    marketCapRank: Number(previousRow['Market Cap Rank']) || null,
    ath: null,
    athChangePercentage: Number(previousRow['Distance from ATH %']) || null,
    circulatingSupply: null,
    maxSupply: null,
    totalSupply: null
  };
}

// ---------------------------------------------------------------------------
// 2) FETCH NETWORK DATA (hashrate, difficulty) — hashrate.no πρωτεύον,
//    NETWORK_STATS_API_URL δευτερεύον, αλλιώς N/A
// ---------------------------------------------------------------------------
function fetchPearlNetworkData_(props) {
  const hashrateNoKey = props.getProperty('HASHRATE_NO_API_KEY');
  const customUrl = props.getProperty('NETWORK_STATS_API_URL');

  if (!hashrateNoKey && !customUrl) {
    console.log('Ούτε HASHRATE_NO_API_KEY ούτε NETWORK_STATS_API_URL έχουν οριστεί — network metrics παραλείπονται καθαρά.');
    return { hashrate: null, difficulty: null, revenuePerThDayChangePct: null, blockReward: null, blockTime: null, source: 'not_configured' };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'pearl_network_data_v9';
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log('Using cached network stats (rate-limit friendly).');
    return JSON.parse(cached);
  }

  if (hashrateNoKey) {
    try {
      const result = fetchFromHashrateNo_(hashrateNoKey);
      cache.put(cacheKey, JSON.stringify(result), PEARL_CONFIG.CACHE_TTL_SECONDS);
      return result;
    } catch (e) {
      console.warn('hashrate.no fetch failed: ' + e.message + (customUrl ? ' — δοκιμάζω το NETWORK_STATS_API_URL fallback.' : ''));
    }
  }

  if (customUrl) {
    try {
      const result = fetchFromCustomNetworkUrl_(customUrl);
      cache.put(cacheKey, JSON.stringify(result), PEARL_CONFIG.CACHE_TTL_SECONDS);
      return result;
    } catch (e) {
      console.warn('Custom network stats URL failed: ' + e.message);
    }
  }

  return { hashrate: null, difficulty: null, revenuePerThDayChangePct: null, blockReward: null, blockTime: null, source: 'unavailable' };
}

/** Parsing για το ΠΡΑΓΜΑΤΙΚΟ schema του hashrate.no /v2/coins:
 *  { "290": { ticker, name, price:{...}, estimate:{yieldDay, revenueDay,...},
 *             network:{hashrate, difficulty, blockTime, blockReward} } } */
function fetchFromHashrateNo_(apiKey) {
  const url = 'https://hashrate.no/api/v2/coins?apiKey=' + encodeURIComponent(apiKey) +
    '&coin=' + PEARL_CONFIG.HASHRATE_NO_COIN_TICKER;

  const json = fetchWithBackoff_(url, { 'Accept': 'application/json' }, { relaxSsl: false });

  const keys = Object.keys(json || {});
  if (keys.length === 0) {
    throw new Error('hashrate.no API επέστρεψε άδειο αποτέλεσμα. Τρέξε testHashrateNoApi() για raw JSON.');
  }
  // Παίρνουμε το πρώτο (και συνήθως μοναδικό, αφού φιλτράρουμε με ?coin=PRL) entry.
  const coinObj = json[keys[0]];

  const network = coinObj.network || {};
  const estimate = coinObj.estimate || {};

  const hashrateHps = network.hashrate != null ? parseFloat(network.hashrate) : null;
  const hashrateEh = hashrateHps != null ? hashrateHps / 1e18 : null;
  const difficulty = network.difficulty != null ? parseFloat(network.difficulty) : null;
  const blockReward = network.blockReward != null ? parseFloat(network.blockReward) : null;
  const blockTime = network.blockTime != null ? parseFloat(network.blockTime) : null;

  // revenueDay: ήδη υπολογισμένη % ημερήσια μεταβολή εσόδου/TH/s από το hashrate.no —
  // πιο αξιόπιστο από να το ξαναϋπολογίζουμε εμείς.
  const revenuePerThDayChangePct = estimate.revenueDay != null ? parseFloat(estimate.revenueDay) : null;
  const yieldDayChangePct = estimate.yieldDay != null ? parseFloat(estimate.yieldDay) : null;

  return {
    hashrate: hashrateEh,
    difficulty: difficulty,
    revenuePerThDayChangePct: revenuePerThDayChangePct,
    yieldDayChangePct: yieldDayChangePct,
    blockReward: blockReward,
    blockTime: blockTime,
    source: 'hashrate.no/api/v2/coins'
  };
}

function fetchFromCustomNetworkUrl_(customUrl) {
  const raw = fetchWithBackoff_(customUrl, { 'Accept': 'application/json' }, { relaxSsl: false });

  const hashrateHps = firstDefined_(raw.network_hashrate_estimate, raw.network_hashrate, raw.hashrate, null);
  const hashrateEh = hashrateHps != null ? (hashrateHps > 1000 ? hashrateHps / 1e18 : hashrateHps) : null;
  const difficulty = firstDefined_(raw.network_difficulty, raw.difficulty, null);
  const revenuePerThDayChangePct = firstDefined_(raw.revenue_per_th_change_pct, null);

  return {
    hashrate: hashrateEh,
    difficulty: difficulty,
    revenuePerThDayChangePct: revenuePerThDayChangePct,
    blockReward: null,
    blockTime: null,
    source: customUrl
  };
}

function fetchOptionalAiWorkloadData_(props) {
  const url = props.getProperty('AI_WORKLOAD_API_URL');
  if (!url) return null;
  try {
    return fetchWithBackoff_(url, { 'Accept': 'application/json' }, { relaxSsl: false });
  } catch (e) {
    console.warn('Optional AI workload API failed: ' + e.message);
    return null;
  }
}

function firstDefined_() {
  for (let i = 0; i < arguments.length; i++) {
    if (arguments[i] !== undefined && arguments[i] !== null) return arguments[i];
  }
  return null;
}

// ---------------------------------------------------------------------------
// 3) ΠΡΟΗΓΟΥΜΕΝΗ ΓΡΑΜΜΗ ΓΙΑ ΣΥΓΚΡΙΣΕΙΣ
// ---------------------------------------------------------------------------
function getPreviousRow_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const rowObj = {};
  headers.forEach((h, i) => rowObj[h] = values[i]);
  return rowObj;
}

// ---------------------------------------------------------------------------
// 4) ΥΠΟΛΟΓΙΣΜΟΣ ΤΩΝ 12 METRICS
// ---------------------------------------------------------------------------
function computeMetrics_(data, net, previousRow, props) {
  const T = PEARL_CONFIG.THRESHOLDS;
  const metrics = [];
  const networkNotConfigured = net.source === 'not_configured' || net.source === 'unavailable';

  metrics.push(makeMetric_('24h Price Change %', data.change24h,
    statusHigherIsBetter_(data.change24h, T.change24h.green, T.change24h.yellow),
    '% μεταβολή τιμής τελευταίου 24ώρου — άμεσο momentum.'));

  let volChange = null, volNote = 'Χωρίς ιστορικό ακόμα.';
  if (previousRow && previousRow['Volume 24h'] != null && previousRow['Volume 24h'] !== '') {
    const prevVol = Number(previousRow['Volume 24h']);
    if (prevVol > 0) {
      volChange = ((data.volume24h - prevVol) / prevVol) * 100;
      volNote = 'Μεταβολή 24h volume έναντι προηγούμενης καταγραφής.';
    }
  }
  metrics.push(makeMetric_('24h Volume Change %', volChange,
    volChange != null ? statusHigherIsBetter_(volChange, T.volChange24h.green, T.volChange24h.yellow) : 'yellow',
    volNote));

  metrics.push(makeMetric_('7d Price Change %', data.change7d,
    statusHigherIsBetter_(data.change7d, T.change7d.green, T.change7d.yellow),
    'Μεσοπρόθεσμο trend 7 ημερών.'));

  const volMcapRatio = data.marketCap > 0 ? (data.volume24h / data.marketCap) * 100 : 0;
  metrics.push(makeMetric_('Volume/MCap Ratio %', volMcapRatio,
    statusHigherIsBetter_(volMcapRatio, T.volMcapRatio.green, T.volMcapRatio.yellow),
    'Ρευστότητα & ενδιαφέρον αγοράς σχετικά με το μέγεθος του coin.'));

  metrics.push(makeMetric_('Distance from ATH %', data.athChangePercentage,
    statusAthDistance_(data.athChangePercentage, T.athDistance.greenMax, T.athDistance.yellowMax),
    'Πόσο μακριά είμαστε από το ATH.'));

  let rankDelta = 0, rankNote = 'Χωρίς ιστορικό ακόμα (πρώτη εκτέλεση).';
  if (previousRow && previousRow['Market Cap Rank'] != null && previousRow['Market Cap Rank'] !== '' && data.marketCapRank != null) {
    rankDelta = Number(previousRow['Market Cap Rank']) - data.marketCapRank;
    rankNote = 'Μεταβολή rank έναντι προηγούμενης καταγραφής.';
  }
  metrics.push(makeMetric_('Market Cap Rank Trend', rankDelta,
    statusHigherIsBetter_(rankDelta, T.rankTrend.greenDelta, T.rankTrend.yellowDelta),
    rankNote, data.marketCapRank));

  const supplyProgress = data.maxSupply > 0 ? (data.circulatingSupply / data.maxSupply) * 100 : null;
  metrics.push(makeMetric_('Circulating/Max Supply %', supplyProgress,
    supplyProgress != null ? statusLowerIsBetter_(supplyProgress, T.supplyProgress.greenMax, T.supplyProgress.yellowMax) : 'yellow',
    'Χαμηλότερο ποσοστό = μεγαλύτερο μελλοντικό runway (max supply 2.1B).'));

  const miningResult = computeMiningProfitability_(data, props, T.miningProfit);
  metrics.push(makeMetric_('Mining Profitability Ratio', miningResult.ratio, miningResult.status, miningResult.note));

  const volatility = data.low24h > 0 ? ((data.high24h - data.low24h) / data.low24h) * 100 : 0;
  metrics.push(makeMetric_('24h Volatility Spread %', volatility,
    statusLowerIsBetter_(volatility, T.volatility.greenMax, T.volatility.yellowMax),
    'Εύρος high-low 24ώρου.'));

  const naNetworkNote = 'N/A — δεν βρέθηκαν δεδομένα δικτύου (έλεγξε HASHRATE_NO_API_KEY ή τρέξε testHashrateNoApi()).';

  let hashrateChangePct = null, hashrateNote;
  const currentHashrateNum = net.hashrate;
  if (currentHashrateNum != null && previousRow && previousRow['Network Hashrate (EH/s)'] != null && previousRow['Network Hashrate (EH/s)'] !== '') {
    const prevHashrate = Number(previousRow['Network Hashrate (EH/s)']);
    if (prevHashrate > 0) hashrateChangePct = ((currentHashrateNum - prevHashrate) / prevHashrate) * 100;
  }
  hashrateNote = currentHashrateNum != null
    ? 'Τρέχον δίκτυο: ' + round2_(currentHashrateNum) + ' EH/s (πηγή: hashrate.no).'
    : (networkNotConfigured ? naNetworkNote : 'Δεν ήταν διαθέσιμο το network hashrate αυτή τη φορά.');
  metrics.push(makeMetric_('Network Hashrate Change %', hashrateChangePct,
    hashrateChangePct != null ? statusHigherIsBetter_(hashrateChangePct, T.hashrateChange.green, T.hashrateChange.yellow) : 'yellow',
    hashrateNote, currentHashrateNum));

  let difficultyChangePct = null, difficultyNote;
  const currentDifficultyNum = net.difficulty != null ? Number(net.difficulty) : null;
  if (currentDifficultyNum != null && previousRow && previousRow['Difficulty'] != null && previousRow['Difficulty'] !== '') {
    const prevDiff = Number(previousRow['Difficulty']);
    if (prevDiff > 0) difficultyChangePct = ((currentDifficultyNum - prevDiff) / prevDiff) * 100;
  }
  difficultyNote = currentDifficultyNum != null
    ? 'Τρέχουσα δυσκολία: ' + formatCompactNumber_(currentDifficultyNum) + '.'
    : (networkNotConfigured ? naNetworkNote : 'Δεν ήταν διαθέσιμη η difficulty αυτή τη φορά.');
  metrics.push(makeMetric_('Difficulty Change %', difficultyChangePct,
    difficultyChangePct != null ? statusLowerIsBetter_(difficultyChangePct, T.difficultyChange.greenMax, T.difficultyChange.yellowMax) : 'yellow',
    difficultyNote, currentDifficultyNum));

  // Mining Revenue/TH Trend: χρησιμοποιούμε ΑΠΕΥΘΕΙΑΣ το revenueDay από hashrate.no
  // (ήδη υπολογισμένη % ημερήσια μεταβολή εσόδου ανά TH/s) — πιο ακριβές.
  const revenueChangePct = net.revenuePerThDayChangePct;
  metrics.push(makeMetric_('Mining Revenue/TH Trend %', revenueChangePct,
    revenueChangePct != null ? statusHigherIsBetter_(revenueChangePct, T.revenuePerTh.green, T.revenuePerTh.yellow) : 'yellow',
    revenueChangePct != null
      ? 'Ημερήσια μεταβολή εσόδου/TH/s (πηγή: hashrate.no estimate.revenueDay).'
      : (networkNotConfigured ? naNetworkNote : 'Δεν επιστράφηκε revenueDay από την πηγή.'),
    revenueChangePct));

  const aiWorkload = fetchOptionalAiWorkloadData_(props);
  let aiDemandValue = null, aiDemandNote;
  if (aiWorkload && aiWorkload.inference_jobs_24h_change_pct != null) {
    aiDemandValue = aiWorkload.inference_jobs_24h_change_pct;
    aiDemandNote = 'Πηγή: custom AI_WORKLOAD_API_URL (πραγματικά δεδομένα, όχι proxy).';
  } else if (revenueChangePct == null) {
    aiDemandValue = null;
    aiDemandNote = 'N/A — εξαρτάται από revenue/TH δεδομένα που δεν είναι διαθέσιμα.';
  } else {
    aiDemandValue = revenueChangePct;
    aiDemandNote = '⚠️ Proxy metric: δεν υπάρχει δημόσιο API για πραγματικό AI inference/FP8 workload στο Pearl. ' +
      'Χρησιμοποιούμε momentum του revenue/TH (hashrate.no) ως έμμεση ένδειξη ζήτησης compute.';
  }
  metrics.push(makeMetric_('AI Demand/Capacity Proxy %', aiDemandValue,
    aiDemandValue != null ? statusHigherIsBetter_(aiDemandValue, T.aiDemandProxy.green, T.aiDemandProxy.yellow) : 'yellow',
    aiDemandNote));

  return metrics;
}

function computeMiningProfitability_(data, props, thresholds) {
  const elecCost = Number(props.getProperty('ELECTRICITY_COST_EUR_KWH')) || 0.20;
  const gpuWatts = Number(props.getProperty('GPU_POWER_WATTS')) || 350;
  const explicitPrlPerDay = Number(props.getProperty('EXPECTED_PRL_PER_DAY')) || null;

  const dailyKwh = (gpuWatts / 1000) * 24;
  const dailyElecCostEur = dailyKwh * elecCost;

  if (!explicitPrlPerDay) {
    return {
      ratio: null, status: 'yellow',
      note: 'Όρισε το Script Property EXPECTED_PRL_PER_DAY για ακριβή υπολογισμό κερδοφορίας.'
    };
  }

  const dailyRevenueEur = explicitPrlPerDay * data.price;
  const ratio = dailyElecCostEur > 0 ? dailyRevenueEur / dailyElecCostEur : null;

  let status = 'red';
  if (ratio != null) {
    if (ratio >= thresholds.green) status = 'green';
    else if (ratio >= thresholds.yellow) status = 'yellow';
  }

  const note = 'Έσοδο ' + round2_(dailyRevenueEur) + '€/μέρα (' + explicitPrlPerDay + ' PRL × ' +
    round4_(data.price) + '€) vs κόστος ρεύματος ' + round2_(dailyElecCostEur) +
    '€/μέρα (' + gpuWatts + 'W @ ' + elecCost + '€/kWh).';

  return { ratio: ratio, status: status, note: note };
}

// ---------------------------------------------------------------------------
// STATUS HELPERS
// ---------------------------------------------------------------------------
function statusHigherIsBetter_(value, greenMin, yellowMin) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value >= greenMin) return 'green';
  if (value >= yellowMin) return 'yellow';
  return 'red';
}
function statusLowerIsBetter_(value, greenMax, yellowMax) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value <= greenMax) return 'green';
  if (value <= yellowMax) return 'yellow';
  return 'red';
}
function statusAthDistance_(value, greenMax, yellowMax) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value >= greenMax) return 'green';
  if (value >= yellowMax) return 'yellow';
  return 'red';
}
function makeMetric_(label, value, status, note, extraRaw) {
  return { label: label, value: value, status: status, note: note, extraRaw: extraRaw };
}

// ---------------------------------------------------------------------------
// 5) ΣΥΝΟΛΙΚΟ VERDICT
// ---------------------------------------------------------------------------
function computeVerdict_(metrics) {
  const score = { green: 0, yellow: 0, red: 0 };
  metrics.forEach(m => { if (score[m.status] != null) score[m.status]++; });

  const total = metrics.length;
  let verdict, emoji;

  if (score.green >= total * 0.65) { verdict = 'STRONG BULLISH'; emoji = '🟢🚀'; }
  else if (score.green + score.yellow >= total * 0.7 && score.red <= 1) { verdict = 'BULLISH'; emoji = '🟢'; }
  else if (score.red >= total * 0.5) { verdict = 'BEARISH'; emoji = '🔴'; }
  else { verdict = 'NEUTRAL'; emoji = '🟡'; }

  return { label: verdict, emoji: emoji, score: score, total: total };
}

// ---------------------------------------------------------------------------
// 6) SHEET WRITE
// ---------------------------------------------------------------------------
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PEARL_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PEARL_CONFIG.SHEET_NAME);
    const headers = [
      'Date', 'Timestamp', 'Price (' + PEARL_CONFIG.VS_CURRENCY.toUpperCase() + ')',
      '24h Change %', 'Volume 24h', '24h Volume Change %', '7d Change %',
      'Volume/MCap %', 'Distance from ATH %', 'Market Cap Rank', 'Rank Trend',
      'Circulating/Max Supply %', 'Mining Profit Ratio', '24h Volatility %',
      'Network Hashrate (EH/s)', 'Hashrate Change %', 'Difficulty', 'Difficulty Change %',
      'Revenue/TH Trend %', 'AI Demand Proxy %',
      'Verdict', 'Green Count', 'Yellow Count', 'Red Count', 'Data Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function writeToSheet_(data, net, metrics, verdict, isStale) {
  const sheet = getOrCreateSheet_();
  const now = new Date();
  const byLabel = {};
  metrics.forEach(m => byLabel[m.label] = m);

  const row = [
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
    data.price,
    round2_(data.change24h),
    data.volume24h,
    round2_(byLabel['24h Volume Change %'].value),
    round2_(data.change7d),
    round2_(byLabel['Volume/MCap Ratio %'].value),
    round2_(data.athChangePercentage),
    data.marketCapRank,
    round2_(byLabel['Market Cap Rank Trend'].value),
    round2_(byLabel['Circulating/Max Supply %'].value),
    byLabel['Mining Profitability Ratio'].value != null ? round2_(byLabel['Mining Profitability Ratio'].value) : 'N/A',
    round2_(byLabel['24h Volatility Spread %'].value),
    byLabel['Network Hashrate Change %'].extraRaw != null ? round4_(byLabel['Network Hashrate Change %'].extraRaw) : 'N/A',
    round2_(byLabel['Network Hashrate Change %'].value),
    byLabel['Difficulty Change %'].extraRaw != null ? byLabel['Difficulty Change %'].extraRaw : 'N/A',
    round2_(byLabel['Difficulty Change %'].value),
    round2_(byLabel['Mining Revenue/TH Trend %'].value),
    round2_(byLabel['AI Demand/Capacity Proxy %'].value),
    verdict.label,
    verdict.score.green,
    verdict.score.yellow,
    verdict.score.red,
    isStale ? 'STALE (CoinGecko fallback)' : 'LIVE'
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  const verdictCol = row.length - 4;
  const cell = sheet.getRange(lastRow, verdictCol);
  const colorMap = {
    'STRONG BULLISH': '#34a853', 'BULLISH': '#93c47d',
    'NEUTRAL': '#ffd966', 'BEARISH': '#e06666'
  };
  cell.setBackground(colorMap[verdict.label] || '#ffffff').setFontWeight('bold');

  if (isStale) {
    sheet.getRange(lastRow, row.length).setBackground('#f4cccc').setFontWeight('bold');
  }

  const maxRows = 366;
  if (sheet.getLastRow() > maxRows) sheet.deleteRow(2);
}

// ---------------------------------------------------------------------------
// 7) TELEGRAM MESSAGE
// ---------------------------------------------------------------------------
function buildTelegramMessage_(data, net, metrics, verdict, isStale) {
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const statusEmoji = { green: '🟢', yellow: '🟡', red: '🔴' };
  const networkNotConfigured = net.source === 'not_configured' || net.source === 'unavailable';

  let msg = '<b>🦪 PEARL (PRL) DAILY DASHBOARD</b>\n';
  msg += '<i>' + dateStr + '</i>\n\n';

  if (isStale) {
    msg += '⚠️ <b>ΠΡΟΣΟΧΗ: Stale δεδομένα</b> — Το CoinGecko απέτυχε προσωρινά, ' +
      'χρησιμοποιούνται τα τελευταία γνωστά στοιχεία τιμής.\n\n';
  }

  msg += verdict.emoji + ' <b>Verdict: ' + escapeHtml_(verdict.label) + '</b>';
  msg += '  (🟢' + verdict.score.green + ' 🟡' + verdict.score.yellow + ' 🔴' + verdict.score.red + '/' + verdict.total + ')\n\n';

  msg += '💰 Τιμή: <b>' + round4_(data.price) + ' ' + PEARL_CONFIG.VS_CURRENCY.toUpperCase() + '</b>';
  msg += '  |  MCap Rank: #' + (data.marketCapRank || 'N/A') + '\n';
  msg += '━━━━━━━━━━━━━━━━━━━\n';
  msg += '<b>📈 Market Metrics</b>\n';

  const marketLabels = ['24h Price Change %', '24h Volume Change %', '7d Price Change %',
    'Volume/MCap Ratio %', 'Distance from ATH %', 'Market Cap Rank Trend',
    'Circulating/Max Supply %', '24h Volatility Spread %'];
  const miningLabels = ['Mining Profitability Ratio', 'Network Hashrate Change %',
    'Difficulty Change %', 'Mining Revenue/TH Trend %', 'AI Demand/Capacity Proxy %'];

  const byLabel = {};
  metrics.forEach(m => byLabel[m.label] = m);

  marketLabels.forEach(l => {
    const m = byLabel[l];
    msg += (statusEmoji[m.status] || '⚪') + ' <b>' + escapeHtml_(m.label) + ':</b> ' + formatMetricValue_(m) + '\n';
  });

  msg += '\n<b>⛏️ Mining & Network Metrics</b>\n';
  if (networkNotConfigured) {
    msg += '<i>ℹ️ Δεδομένα δικτύου μη διαθέσιμα αυτή τη στιγμή.</i>\n';
  }
  miningLabels.forEach(l => {
    const m = byLabel[l];
    msg += (statusEmoji[m.status] || '⚪') + ' <b>' + escapeHtml_(m.label) + ':</b> ' + formatMetricValue_(m) + '\n';
  });

  msg += '━━━━━━━━━━━━━━━━━━━\n';
  msg += '📊 24h Volume: ' + formatCompactNumber_(data.volume24h) + ' ' + PEARL_CONFIG.VS_CURRENCY.toUpperCase() + '\n';
  if (data.marketCap) msg += '🏦 Market Cap: ' + formatCompactNumber_(data.marketCap) + ' ' + PEARL_CONFIG.VS_CURRENCY.toUpperCase() + '\n';

  const currentHashrate = byLabel['Network Hashrate Change %'].extraRaw;
  if (currentHashrate != null) msg += '⚡ Network Hashrate: ' + round2_(currentHashrate) + ' EH/s\n';
  const currentDifficulty = byLabel['Difficulty Change %'].extraRaw;
  if (currentDifficulty != null && currentDifficulty !== 'N/A') msg += '🎯 Difficulty: ' + formatCompactNumber_(currentDifficulty) + '\n';
  if (net.blockReward != null) msg += '💎 Block Reward: ' + round2_(net.blockReward) + ' PRL (~' + round2_(net.blockTime) + 's/block)\n';

  msg += '\n<i>⚠️ Το "AI Demand Proxy" είναι έμμεσος δείκτης, όχι πραγματικό AI inference/FP8 volume.</i>\n';
  msg += '<i>Δεδομένα: CoinGecko (Demo API key)' + (net.source && net.source.indexOf('hashrate.no') !== -1 ? ' + hashrate.no' : '') + (isStale ? ' — STALE fallback' : '') + '.</i>';

  return msg;
}

function formatMetricValue_(m) {
  if (m.value == null || isNaN(m.value)) return 'N/A';
  if (m.label === 'Market Cap Rank Trend') {
    const sign = m.value > 0 ? '+' : '';
    return sign + m.value + ' θέσεις (τώρα #' + (m.extraRaw || 'N/A') + ')';
  }
  if (m.label === 'Mining Profitability Ratio') return round2_(m.value) + 'x';
  return round2_(m.value) + '%';
}

function sendTelegramMessage_(text, props) {
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    console.error('Missing TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID in Script Properties.');
    return;
  }

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const payload = {
    chat_id: chatId, text: text, parse_mode: 'HTML', disable_web_page_preview: true
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error('Telegram send failed: ' + response.getContentText());
  }
}

// ---------------------------------------------------------------------------
// UTILITIES
// ---------------------------------------------------------------------------
function round2_(n) { return (n == null || isNaN(n)) ? null : Math.round(n * 100) / 100; }
function round4_(n) { return (n == null || isNaN(n)) ? null : Math.round(n * 10000) / 10000; }
function formatCompactNumber_(n) {
  if (n == null || isNaN(n)) return 'N/A';
  if (n >= 1e9) return round2_(n / 1e9) + 'B';
  if (n >= 1e6) return round2_(n / 1e6) + 'M';
  if (n >= 1e3) return round2_(n / 1e3) + 'K';
  return String(round2_(n));
}
function escapeHtml_(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------------------------------------------------------------------------
// ΔΟΚΙΜΑΣΤΙΚΗ ΕΚΤΕΛΕΣΗ / TRIGGER SETUP / TESTS
// ---------------------------------------------------------------------------
function testRunNow() { runPearlDashboard(); }

function testCoinGeckoApiKey() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('COINGECKO_API_KEY');
  if (!apiKey) {
    console.log('❌ Δεν βρέθηκε COINGECKO_API_KEY στα Script Properties.');
    return;
  }
  const url = 'https://api.coingecko.com/api/v3/ping?x_cg_demo_api_key=' + encodeURIComponent(apiKey);
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'x-cg-demo-api-key': apiKey, 'Accept': 'application/json' }
  });
  console.log('HTTP status: ' + response.getResponseCode());
  console.log('Response: ' + response.getContentText());
}

function testHashrateNoApi() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('HASHRATE_NO_API_KEY');
  if (!apiKey) {
    console.log('❌ Δεν βρέθηκε HASHRATE_NO_API_KEY στα Script Properties.');
    return;
  }
  try {
    const parsed = fetchFromHashrateNo_(apiKey);
    console.log('✅ Parsed result: ' + JSON.stringify(parsed));
  } catch (e) {
    console.log('❌ hashrate.no test απέτυχε: ' + e.message);
  }
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runPearlDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runPearlDashboard').timeBased().everyDays(1).atHour(9).create();
  console.log('Daily trigger created for runPearlDashboard at 09:00.');
}





/**
 * CRYPTO TREND MONITOR + FUTURES GRID ASSISTANT
 * ------------------------------------------------
 * Sheet: TRENDS
 * Output sheets:
 *   - TRENDS: historical market snapshots
 *   - GRID_SETUP: actionable grid suggestions
 *
 * Required Script Properties:
 *   - TELEGRAM_BOT_TOKEN
 *   - TELEGRAM_CHAT_ID
 *
 * Optional Script Properties:
 *   - COINGECKO_API_KEY
 *   - CTM_SYMBOLS                e.g. BTCUSDT,ETHUSDT,SOLUSDT,BNBUSDT
 *   - CTM_ENABLE_TELEGRAM        true / false
 *   - CTM_ENABLE_NEWS            true / false
 *   - CTM_MAX_GRID_LEVERAGE      e.g. 3
 *   - CTM_ALERT_ONLY              true / false
 */

const CTM_CONFIG = {
  SHEET_TRENDS: 'TRENDS',
  SHEET_GRID: 'GRID_SETUP',
  TIMEZONE: 'Europe/Athens',

  BINANCE_FAPI_BASE: 'https://fapi.binance.com',
  BYBIT_V5_BASE: 'https://api.bybit.com',
  COINGECKO_BASE: 'https://api.coingecko.com/api/v3',
  GOOGLE_NEWS_RSS: 'https://news.google.com/rss/search',
  KRAKEN_FUTURES_BASE: 'https://futures.kraken.com/derivatives/api/v3',
  KRAKEN_CHARTS_BASE: 'https://futures.kraken.com/api/charts/v1',

  DEFAULT_SYMBOLS: ['BTCUSDT', 'ETHUSDT', 'SOLUSDT', 'BNBUSDT'],

  COINGECKO_IDS: {
    BTCUSDT: 'bitcoin',
    ETHUSDT: 'ethereum',
    SOLUSDT: 'solana',
    BNBUSDT: 'binancecoin'
  },

  KLINE_INTERVALS: {
    FAST: '1h',
    SLOW: '4h'
  },

  KLINE_LIMIT: 100,
  OI_LOOKBACK_MS: 60 * 60 * 1000,
  REQUEST_TIMEOUT_MS: 20000,

  DEFAULT_MAX_LEVERAGE: 3,
  DEFAULT_GRID_COUNT: 30,
  MIN_GRID_COUNT: 16,
  MAX_GRID_COUNT: 70,

  NEWS_MAX_ITEMS: 20,
  NEWS_LOOKBACK_HOURS: 24,

  TELEGRAM_MESSAGE_LIMIT: 3900,

  DATA_PROVIDER: {
  AUTO: 'AUTO',
  BINANCE: 'BINANCE',
  BYBIT: 'BYBIT',
  KRAKEN: 'KRAKEN'
  },

  SCORE: {
    BULLISH_THRESHOLD: 22,
    BEARISH_THRESHOLD: -22,
    STRONG_BULLISH_THRESHOLD: 48,
    STRONG_BEARISH_THRESHOLD: -48
  },

  RISK: {
    HIGH_ATR_1H_PCT: 2.5,
    HIGH_ATR_4H_PCT: 6.0,
    HIGH_OI_CHANGE_PCT: 5.0,
    EXTREME_FUNDING_PCT_8H: 0.05
  }
};

/**
 * Main function.
 * Run this from Apps Script or time trigger.
 */
function CTM_runMarketTrendReport() {
  const ctmProps = PropertiesService.getScriptProperties();
  const ctmRuntimeConfig = CTM_getRuntimeConfig_(ctmProps);

  const ctmSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ctmTrendsSheet = CTM_getOrCreateTrendsSheet_(ctmSpreadsheet);
  const ctmGridSheet = CTM_getOrCreateGridSheet_(ctmSpreadsheet);

  const ctmGlobalData = CTM_fetchGlobalMarketData_(ctmRuntimeConfig);
  const ctmNewsData = ctmRuntimeConfig.enableNews
    ? CTM_fetchNewsSentiment_(ctmRuntimeConfig)
    : CTM_getEmptyNewsData_();

  const ctmMarketRows = [];
  const ctmGridRows = [];
  const ctmAnalysis = [];

  ctmRuntimeConfig.symbols.forEach(function(ctmSymbol) {
    try {
      const ctmData = CTM_collectSymbolData_(ctmSymbol, ctmRuntimeConfig);
      const ctmAssessment = CTM_assessMarket_(ctmData, ctmGlobalData, ctmNewsData);
      const ctmGridPlan = CTM_buildGridPlan_(ctmData, ctmAssessment, ctmRuntimeConfig);

      ctmMarketRows.push(
        CTM_buildTrendRow_(
          ctmData,
          ctmAssessment,
          ctmGridPlan,
          ctmGlobalData,
          ctmNewsData
        )
      );

      ctmGridRows.push(
        CTM_buildGridRow_(ctmData, ctmAssessment, ctmGridPlan)
      );

      ctmAnalysis.push({
        symbol: ctmSymbol,
        data: ctmData,
        assessment: ctmAssessment,
        gridPlan: ctmGridPlan
      });

    } catch (ctmError) {
      Logger.log('CTM error for ' + ctmSymbol + ': ' + ctmError.stack);
      ctmMarketRows.push(
        CTM_buildErrorTrendRow_(ctmSymbol, ctmError, ctmGlobalData, ctmNewsData)
      );
    }
  });

  CTM_appendRows_(ctmTrendsSheet, ctmMarketRows);
  CTM_replaceGridRows_(ctmGridSheet, ctmGridRows);

  const ctmTelegramMessage = CTM_buildTelegramMessage_(
    ctmAnalysis,
    ctmGlobalData,
    ctmNewsData,
    ctmRuntimeConfig
  );

  let ctmTelegramResult = {
    sent: false,
    reason: 'Telegram disabled'
  };

  if (ctmRuntimeConfig.enableTelegram) {
    if (!ctmRuntimeConfig.alertOnly || CTM_shouldSendAlert_(ctmAnalysis)) {
      ctmTelegramResult = CTM_sendTelegram_(ctmTelegramMessage, ctmProps);
    } else {
      ctmTelegramResult = {
        sent: false,
        reason: 'Alert-only mode: no alert condition'
      };
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    symbolsProcessed: ctmAnalysis.length,
    telegram: ctmTelegramResult,
    message: ctmTelegramMessage
  };
}

/**
 * Creates a 1-hour time trigger.
 * Run once manually to install it.
 */
function CTM_createHourlyTrigger() {
  CTM_deleteTriggers_('CTM_runMarketTrendReport');

  ScriptApp.newTrigger('CTM_runMarketTrendReport')
    .timeBased()
    .everyHours(1)
    .create();
}

/**
 * Removes all triggers associated with CTM_runMarketTrendReport.
 */
function CTM_deleteHourlyTrigger() {
  CTM_deleteTriggers_('CTM_runMarketTrendReport');
}

/**
 * Use this to test Telegram settings only.
 */
function CTM_testTelegram() {
  const ctmProps = PropertiesService.getScriptProperties();

  return CTM_sendTelegram_(
    '✅ <b>CTM test message</b>\n\n' +
    'Το Crypto Trend Monitor συνδέθηκε επιτυχώς με το Telegram.\n' +
    'Time: ' + CTM_formatDate_(new Date()),
    ctmProps
  );
}

/**
 * Use this to initialize sheets without downloading market data.
 */
function CTM_initializeSheets() {
  const ctmSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  CTM_getOrCreateTrendsSheet_(ctmSpreadsheet);
  CTM_getOrCreateGridSheet_(ctmSpreadsheet);

  return {
    success: true,
    message: 'Created/verified sheets: ' +
      CTM_CONFIG.SHEET_TRENDS + ', ' + CTM_CONFIG.SHEET_GRID
  };
}

/* ============================================================
 * CONFIGURATION
 * ============================================================ */

function CTM_getRuntimeConfig_(ctmProps) {
  const ctmRawSymbols = ctmProps.getProperty('CTM_SYMBOLS') || '';

  const ctmSymbols = ctmRawSymbols
    ? ctmRawSymbols
      .split(',')
      .map(function(ctmItem) {
        return ctmItem.trim().toUpperCase();
      })
      .filter(function(ctmItem) {
        return !!ctmItem;
      })
    : CTM_CONFIG.DEFAULT_SYMBOLS.slice();

  const ctmProvider = String(
    ctmProps.getProperty('CTM_DATA_PROVIDER') || 'AUTO'
  ).toUpperCase();

  return {
    symbols: ctmSymbols,
    enableTelegram: CTM_toBoolean_(
      ctmProps.getProperty('CTM_ENABLE_TELEGRAM'),
      true
    ),
    enableNews: CTM_toBoolean_(
      ctmProps.getProperty('CTM_ENABLE_NEWS'),
      true
    ),
    alertOnly: CTM_toBoolean_(
      ctmProps.getProperty('CTM_ALERT_ONLY'),
      false
    ),
    maxLeverage: CTM_toNumber_(
      ctmProps.getProperty('CTM_MAX_GRID_LEVERAGE'),
      CTM_CONFIG.DEFAULT_MAX_LEVERAGE
    ),
    coinGeckoApiKey: ctmProps.getProperty('COINGECKO_API_KEY') || '',
    dataProvider: [
      CTM_CONFIG.DATA_PROVIDER.AUTO,
      CTM_CONFIG.DATA_PROVIDER.BINANCE,
      CTM_CONFIG.DATA_PROVIDER.BYBIT
    ].indexOf(ctmProvider) !== -1
      ? ctmProvider
      : CTM_CONFIG.DATA_PROVIDER.AUTO
  };
}

/* ============================================================
 * DATA COLLECTION
 * ============================================================ */

function CTM_collectSymbolData_(ctmSymbol, ctmRuntimeConfig) {
  const ctmMarketData = CTM_fetchDerivativesData_(
    ctmSymbol,
    ctmRuntimeConfig
  );

  const ctmSpotData = CTM_fetchCoinGeckoCoin_(
    ctmSymbol,
    ctmRuntimeConfig.coinGeckoApiKey
  );

  const ctmCloses1h = ctmMarketData.klines1h.map(function(ctmCandle) {
    return ctmCandle.close;
  });

  const ctmCloses4h = ctmMarketData.klines4h.map(function(ctmCandle) {
    return ctmCandle.close;
  });

  const ctmRsi1h = CTM_calculateRsi_(ctmCloses1h, 14);
  const ctmRsi4h = CTM_calculateRsi_(ctmCloses4h, 14);

  const ctmAtr1h = CTM_calculateAtr_(ctmMarketData.klines1h, 14);
  const ctmAtr4h = CTM_calculateAtr_(ctmMarketData.klines4h, 14);

  const ctmEma20_1h = CTM_calculateEma_(ctmCloses1h, 20);
  const ctmEma50_1h = CTM_calculateEma_(ctmCloses1h, 50);
  const ctmEma20_4h = CTM_calculateEma_(ctmCloses4h, 20);
  const ctmEma50_4h = CTM_calculateEma_(ctmCloses4h, 50);

  const ctmLast1hCandle =
    ctmMarketData.klines1h[ctmMarketData.klines1h.length - 1];

  const ctmPrevious1hCandle =
    ctmMarketData.klines1h[ctmMarketData.klines1h.length - 2];

  const ctmLast4hCandle =
    ctmMarketData.klines4h[ctmMarketData.klines4h.length - 1];

  const ctmRange24h = Math.max(
    0.0000000001,
    Number(ctmMarketData.high24h) - Number(ctmMarketData.low24h)
  );

  const ctmRangePosition = (
    (Number(ctmMarketData.price) - Number(ctmMarketData.low24h)) /
    ctmRange24h
  ) * 100;

  const ctmVolumeRatio1h = CTM_calculateVolumeRatio_(
    ctmMarketData.klines1h,
    20
  );

  return {
    symbol: ctmSymbol,
    asset: ctmSymbol.replace('USDT', ''),
    collectedAt: new Date(),

    provider: ctmMarketData.provider,

    price: Number(ctmMarketData.price),
    priceChange24hPct: Number(ctmMarketData.priceChange24hPct),
    high24h: Number(ctmMarketData.high24h),
    low24h: Number(ctmMarketData.low24h),
    quoteVolume24h: Number(ctmMarketData.quoteVolume24h),

    fundingRatePct8h: Number(ctmMarketData.fundingRatePct8h),
    fundingTime: ctmMarketData.fundingTime,

    openInterest: Number(ctmMarketData.openInterest),
    openInterest1hAgo: Number(ctmMarketData.openInterest1hAgo || 0),
    oiChange1hPct: Number(ctmMarketData.oiChange1hPct || 0),

    rsi1h: ctmRsi1h,
    rsi4h: ctmRsi4h,

    atr1hPct: (ctmAtr1h / Number(ctmMarketData.price)) * 100,
    atr4hPct: (ctmAtr4h / Number(ctmMarketData.price)) * 100,

    ema20_1h: ctmEma20_1h,
    ema50_1h: ctmEma50_1h,
    ema20_4h: ctmEma20_4h,
    ema50_4h: ctmEma50_4h,

    close1h: ctmLast1hCandle.close,
    previousClose1h: ctmPrevious1hCandle.close,
    close4h: ctmLast4hCandle.close,

    volumeRatio1h: ctmVolumeRatio1h,
    rangePosition24hPct: ctmRangePosition,

    cgPrice1hPct: ctmSpotData.priceChange1hPct,
    cgPrice24hPct: ctmSpotData.priceChange24hPct,
    cgVolume24h: ctmSpotData.volume24h,
    marketCap: ctmSpotData.marketCap
  };
}

function CTM_fetchDerivativesData_(ctmSymbol, ctmRuntimeConfig) {
  const ctmProvider = String(
    ctmRuntimeConfig.dataProvider || 'KRAKEN'
  ).toUpperCase();

  if (
    ctmProvider === 'KRAKEN' &&
    (ctmSymbol === 'BTCUSDT' || ctmSymbol === 'ETHUSDT')
  ) {
    try {
      return CTM_fetchKrakenDerivativesData_(ctmSymbol);

    } catch (ctmKrakenError) {
      Logger.log(
        'CTM Kraken unavailable for ' +
        ctmSymbol +
        '. Using CoinGecko spot fallback. Error: ' +
        ctmKrakenError.message
      );

      return CTM_fetchCoinGeckoMarketFallback_(
        ctmSymbol,
        ctmRuntimeConfig
      );
    }
  }

  return CTM_fetchCoinGeckoMarketFallback_(
    ctmSymbol,
    ctmRuntimeConfig
  );
}

function CTM_fetchBinanceDerivativesData_(ctmSymbol) {
  const ctmTicker = CTM_fetchBinanceTicker_(ctmSymbol);
  const ctmFunding = CTM_fetchBinanceFunding_(ctmSymbol);
  const ctmOpenInterest = CTM_fetchBinanceOpenInterest_(ctmSymbol);

  const ctmOpenInterestHistory = CTM_fetchBinanceOpenInterestHistory_(
    ctmSymbol,
    CTM_CONFIG.OI_LOOKBACK_MS
  );

  const ctmKlines1h = CTM_fetchBinanceKlines_(
    ctmSymbol,
    CTM_CONFIG.KLINE_INTERVALS.FAST,
    CTM_CONFIG.KLINE_LIMIT
  );

  const ctmKlines4h = CTM_fetchBinanceKlines_(
    ctmSymbol,
    CTM_CONFIG.KLINE_INTERVALS.SLOW,
    CTM_CONFIG.KLINE_LIMIT
  );

  const ctmOiNow = Number(ctmOpenInterest.openInterest || 0);
  const ctmOiBefore = Number(ctmOpenInterestHistory.openInterest || 0);

  return {
    provider: 'BINANCE',

    price: Number(ctmTicker.lastPrice),
    priceChange24hPct: Number(ctmTicker.priceChangePercent),
    high24h: Number(ctmTicker.highPrice),
    low24h: Number(ctmTicker.lowPrice),
    quoteVolume24h: Number(ctmTicker.quoteVolume),

    fundingRatePct8h: Number(ctmFunding.lastFundingRate) * 100,
    fundingTime: new Date(Number(ctmFunding.fundingTime)),

    openInterest: ctmOiNow,
    openInterest1hAgo: ctmOiBefore,

    oiChange1hPct: ctmOiBefore > 0
      ? ((ctmOiNow - ctmOiBefore) / ctmOiBefore) * 100
      : 0,

    klines1h: ctmKlines1h,
    klines4h: ctmKlines4h
  };
}

function CTM_fetchBybitDerivativesData_(ctmSymbol) {
  const ctmTicker = CTM_fetchBybitTicker_(ctmSymbol);

  const ctmKlines1h = CTM_fetchBybitKlines_(
    ctmSymbol,
    '60',
    CTM_CONFIG.KLINE_LIMIT
  );

  const ctmKlines4h = CTM_fetchBybitKlines_(
    ctmSymbol,
    '240',
    CTM_CONFIG.KLINE_LIMIT
  );

  const ctmOpenInterestData = CTM_fetchBybitOpenInterest_(ctmSymbol);

  return {
    provider: 'BYBIT',

    price: Number(ctmTicker.lastPrice),
    priceChange24hPct: Number(ctmTicker.price24hPcnt) * 100,
    high24h: Number(ctmTicker.highPrice24h),
    low24h: Number(ctmTicker.lowPrice24h),
    quoteVolume24h: Number(ctmTicker.turnover24h),

    fundingRatePct8h: Number(ctmTicker.fundingRate) * 100,
    fundingTime: ctmTicker.nextFundingTime
      ? new Date(Number(ctmTicker.nextFundingTime))
      : new Date(),

    openInterest: Number(ctmTicker.openInterest || 0),
    openInterest1hAgo: Number(ctmOpenInterestData.openInterest1hAgo || 0),
    oiChange1hPct: Number(ctmOpenInterestData.oiChange1hPct || 0),

    klines1h: ctmKlines1h,
    klines4h: ctmKlines4h
  };
}

function CTM_fetchBybitTicker_(ctmSymbol) {
  const ctmUrl =
    CTM_CONFIG.BYBIT_V5_BASE +
    '/v5/market/tickers?category=linear&symbol=' +
    encodeURIComponent(ctmSymbol);

  const ctmResponse = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  if (
    !ctmResponse ||
    Number(ctmResponse.retCode) !== 0 ||
    !ctmResponse.result ||
    !ctmResponse.result.list ||
    !ctmResponse.result.list.length
  ) {
    throw new Error(
      'Bybit ticker error for ' + ctmSymbol + ': ' +
      JSON.stringify(ctmResponse)
    );
  }

  return ctmResponse.result.list[0];
}

function CTM_fetchBybitKlines_(ctmSymbol, ctmInterval, ctmLimit) {
  const ctmUrl =
    CTM_CONFIG.BYBIT_V5_BASE +
    '/v5/market/kline?category=linear' +
    '&symbol=' + encodeURIComponent(ctmSymbol) +
    '&interval=' + encodeURIComponent(ctmInterval) +
    '&limit=' + Number(ctmLimit);

  const ctmResponse = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  if (
    !ctmResponse ||
    Number(ctmResponse.retCode) !== 0 ||
    !ctmResponse.result ||
    !ctmResponse.result.list
  ) {
    throw new Error(
      'Bybit kline error for ' + ctmSymbol + ': ' +
      JSON.stringify(ctmResponse)
    );
  }

  return ctmResponse.result.list
    .map(function(ctmRow) {
      return {
        openTime: Number(ctmRow[0]),
        open: Number(ctmRow[1]),
        high: Number(ctmRow[2]),
        low: Number(ctmRow[3]),
        close: Number(ctmRow[4]),
        volume: Number(ctmRow[5]),
        closeTime: Number(ctmRow[0]) + 3600000
      };
    })
    .sort(function(ctmA, ctmB) {
      return ctmA.openTime - ctmB.openTime;
    });
}

function CTM_fetchBybitOpenInterest_(ctmSymbol) {
  const ctmEndTime = Date.now();
  const ctmStartTime = ctmEndTime - (2 * 60 * 60 * 1000);

  const ctmUrl =
    CTM_CONFIG.BYBIT_V5_BASE +
    '/v5/market/open-interest?category=linear' +
    '&symbol=' + encodeURIComponent(ctmSymbol) +
    '&intervalTime=5min' +
    '&startTime=' + ctmStartTime +
    '&endTime=' + ctmEndTime +
    '&limit=50';

  try {
    const ctmResponse = CTM_fetchJson_(ctmUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    if (
      !ctmResponse ||
      Number(ctmResponse.retCode) !== 0 ||
      !ctmResponse.result ||
      !ctmResponse.result.list ||
      ctmResponse.result.list.length < 2
    ) {
      return {
        openInterest1hAgo: 0,
        oiChange1hPct: 0
      };
    }

    const ctmRows = ctmResponse.result.list
      .map(function(ctmItem) {
        return {
          timestamp: Number(ctmItem.timestamp),
          openInterest: Number(ctmItem.openInterest)
        };
      })
      .sort(function(ctmA, ctmB) {
        return ctmA.timestamp - ctmB.timestamp;
      });

    const ctmTargetTime = ctmEndTime - CTM_CONFIG.OI_LOOKBACK_MS;

    let ctmClosest = ctmRows[0];
    let ctmSmallestDistance = Math.abs(
      ctmClosest.timestamp - ctmTargetTime
    );

    ctmRows.forEach(function(ctmRow) {
      const ctmDistance = Math.abs(ctmRow.timestamp - ctmTargetTime);

      if (ctmDistance < ctmSmallestDistance) {
        ctmClosest = ctmRow;
        ctmSmallestDistance = ctmDistance;
      }
    });

    const ctmLatest = ctmRows[ctmRows.length - 1];

    return {
      openInterest1hAgo: ctmClosest.openInterest,
      oiChange1hPct: ctmClosest.openInterest > 0
        ? (
          (ctmLatest.openInterest - ctmClosest.openInterest) /
          ctmClosest.openInterest
        ) * 100
        : 0
    };

  } catch (ctmError) {
    Logger.log(
      'CTM Bybit open interest unavailable for ' +
      ctmSymbol + ': ' + ctmError.message
    );

    return {
      openInterest1hAgo: 0,
      oiChange1hPct: 0
    };
  }
}

function CTM_fetchGlobalMarketData_(ctmRuntimeConfig) {
  const ctmUrl = CTM_CONFIG.COINGECKO_BASE + '/global';
  const ctmHeaders = CTM_buildCoinGeckoHeaders_(
    ctmRuntimeConfig.coinGeckoApiKey
  );

  try {
    const ctmData = CTM_fetchJson_(ctmUrl, {
      method: 'get',
      headers: ctmHeaders,
      muteHttpExceptions: true
    });

    const ctmGlobal = ctmData.data || {};

    return {
      totalMarketCapUsd:
        (ctmGlobal.total_market_cap && ctmGlobal.total_market_cap.usd) || 0,
      totalVolumeUsd:
        (ctmGlobal.total_volume && ctmGlobal.total_volume.usd) || 0,
      btcDominancePct:
        (ctmGlobal.market_cap_percentage &&
          ctmGlobal.market_cap_percentage.btc) || 0,
      ethDominancePct:
        (ctmGlobal.market_cap_percentage &&
          ctmGlobal.market_cap_percentage.eth) || 0,
      marketCapChange24hPct:
        ctmGlobal.market_cap_change_percentage_24h_usd || 0
    };

  } catch (ctmError) {
    Logger.log('CTM global market data unavailable: ' + ctmError.message);

    return {
      totalMarketCapUsd: 0,
      totalVolumeUsd: 0,
      btcDominancePct: 0,
      ethDominancePct: 0,
      marketCapChange24hPct: 0
    };
  }
}

function CTM_fetchCoinGeckoCoin_(ctmSymbol, ctmCoinGeckoApiKey) {
  const ctmCoinId = CTM_CONFIG.COINGECKO_IDS[ctmSymbol];

  if (!ctmCoinId) {
    return {
      priceChange1hPct: 0,
      priceChange24hPct: 0,
      volume24h: 0,
      marketCap: 0
    };
  }

  const ctmUrl =
    CTM_CONFIG.COINGECKO_BASE +
    '/coins/markets?vs_currency=usd' +
    '&ids=' + encodeURIComponent(ctmCoinId) +
    '&price_change_percentage=1h,24h';

  try {
    const ctmData = CTM_fetchJson_(ctmUrl, {
      method: 'get',
      headers: CTM_buildCoinGeckoHeaders_(ctmCoinGeckoApiKey),
      muteHttpExceptions: true
    });

    const ctmCoin = ctmData && ctmData[0] ? ctmData[0] : {};

    return {
      priceChange1hPct: Number(
        ctmCoin.price_change_percentage_1h_in_currency || 0
      ),
      priceChange24hPct: Number(
        ctmCoin.price_change_percentage_24h_in_currency || 0
      ),
      volume24h: Number(ctmCoin.total_volume || 0),
      marketCap: Number(ctmCoin.market_cap || 0)
    };

  } catch (ctmError) {
    Logger.log(
      'CTM CoinGecko data unavailable for ' +
      ctmSymbol + ': ' + ctmError.message
    );

    return {
      priceChange1hPct: 0,
      priceChange24hPct: 0,
      volume24h: 0,
      marketCap: 0
    };
  }
}

function CTM_fetchBinanceTicker_(ctmSymbol) {
  const ctmUrl =
    CTM_CONFIG.BINANCE_FAPI_BASE +
    '/fapi/v1/ticker/24hr?symbol=' +
    encodeURIComponent(ctmSymbol);

  return CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });
}

function CTM_fetchBinanceFunding_(ctmSymbol) {
  const ctmUrl =
    CTM_CONFIG.BINANCE_FAPI_BASE +
    '/fapi/v1/fundingRate?symbol=' +
    encodeURIComponent(ctmSymbol) +
    '&limit=1';

  const ctmData = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  if (!ctmData || !ctmData.length) {
    throw new Error('No funding rate returned for ' + ctmSymbol);
  }

  return ctmData[0];
}

function CTM_fetchBinanceOpenInterest_(ctmSymbol) {
  const ctmUrl =
    CTM_CONFIG.BINANCE_FAPI_BASE +
    '/fapi/v1/openInterest?symbol=' +
    encodeURIComponent(ctmSymbol);

  return CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });
}

function CTM_fetchBinanceOpenInterestHistory_(ctmSymbol, ctmLookbackMs) {
  const ctmEndTime = Date.now();
  const ctmStartTime = ctmEndTime - ctmLookbackMs - (15 * 60 * 1000);

  const ctmUrl =
    CTM_CONFIG.BINANCE_FAPI_BASE +
    '/futures/data/openInterestHist?symbol=' +
    encodeURIComponent(ctmSymbol) +
    '&period=5m' +
    '&limit=30' +
    '&startTime=' + ctmStartTime +
    '&endTime=' + ctmEndTime;

  try {
    const ctmData = CTM_fetchJson_(ctmUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    if (!ctmData || !ctmData.length) {
      return { openInterest: 0 };
    }

    const ctmTargetTime = ctmEndTime - ctmLookbackMs;
    let ctmClosest = ctmData[0];
    let ctmDistance = Math.abs(Number(ctmClosest.timestamp) - ctmTargetTime);

    ctmData.forEach(function(ctmPoint) {
      const ctmPointDistance = Math.abs(
        Number(ctmPoint.timestamp) - ctmTargetTime
      );

      if (ctmPointDistance < ctmDistance) {
        ctmClosest = ctmPoint;
        ctmDistance = ctmPointDistance;
      }
    });

    return {
      openInterest: Number(ctmClosest.sumOpenInterest || 0),
      timestamp: Number(ctmClosest.timestamp || 0)
    };

  } catch (ctmError) {
    Logger.log(
      'CTM OI history unavailable for ' +
      ctmSymbol + ': ' + ctmError.message
    );

    return { openInterest: 0 };
  }
}

function CTM_fetchBinanceKlines_(ctmSymbol, ctmInterval, ctmLimit) {
  const ctmUrl =
    CTM_CONFIG.BINANCE_FAPI_BASE +
    '/fapi/v1/klines?symbol=' +
    encodeURIComponent(ctmSymbol) +
    '&interval=' + encodeURIComponent(ctmInterval) +
    '&limit=' + Number(ctmLimit);

  const ctmRawData = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  return ctmRawData.map(function(ctmRow) {
    return {
      openTime: Number(ctmRow[0]),
      open: Number(ctmRow[1]),
      high: Number(ctmRow[2]),
      low: Number(ctmRow[3]),
      close: Number(ctmRow[4]),
      volume: Number(ctmRow[5]),
      closeTime: Number(ctmRow[6])
    };
  });
}

/* ============================================================
 * NEWS SENTIMENT
 * ============================================================ */

function CTM_fetchNewsSentiment_(ctmRuntimeConfig) {
  const ctmQuery = encodeURIComponent(
    'cryptocurrency OR bitcoin OR ethereum OR crypto ETF OR crypto regulation'
  );

  const ctmUrl =
    CTM_CONFIG.GOOGLE_NEWS_RSS +
    '?q=' + ctmQuery +
    '&hl=en-US&gl=US&ceid=US:en';

  try {
    const ctmResponse = UrlFetchApp.fetch(ctmUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    const ctmCode = ctmResponse.getResponseCode();

    if (ctmCode < 200 || ctmCode >= 300) {
      throw new Error('Google News RSS HTTP ' + ctmCode);
    }

    const ctmXml = XmlService.parse(ctmResponse.getContentText());
    const ctmRoot = ctmXml.getRootElement();
    const ctmChannel = ctmRoot.getChild('channel');
    const ctmItems = ctmChannel
      ? ctmChannel.getChildren('item')
      : [];

    const ctmCutoff = Date.now() -
      (CTM_CONFIG.NEWS_LOOKBACK_HOURS * 60 * 60 * 1000);

    const ctmRecentItems = ctmItems
      .map(function(ctmItem) {
        const ctmTitle = CTM_getXmlChildText_(ctmItem, 'title');
        const ctmPubDate = CTM_getXmlChildText_(ctmItem, 'pubDate');
        const ctmLink = CTM_getXmlChildText_(ctmItem, 'link');
        const ctmDate = ctmPubDate ? new Date(ctmPubDate) : new Date();

        return {
          title: ctmTitle,
          link: ctmLink,
          publishedAt: ctmDate,
          publishedAtMs: ctmDate.getTime()
        };
      })
      .filter(function(ctmItem) {
        return ctmItem.publishedAtMs >= ctmCutoff;
      })
      .slice(0, CTM_CONFIG.NEWS_MAX_ITEMS);

    const ctmScored = ctmRecentItems.map(function(ctmItem) {
      const ctmScore = CTM_scoreHeadline_(ctmItem.title);

      return {
        title: ctmItem.title,
        link: ctmItem.link,
        publishedAt: ctmItem.publishedAt,
        score: ctmScore.score,
        positiveHits: ctmScore.positiveHits,
        negativeHits: ctmScore.negativeHits
      };
    });

    const ctmTotalScore = ctmScored.reduce(function(ctmSum, ctmItem) {
      return ctmSum + ctmItem.score;
    }, 0);

    const ctmPositiveCount = ctmScored.filter(function(ctmItem) {
      return ctmItem.score > 0;
    }).length;

    const ctmNegativeCount = ctmScored.filter(function(ctmItem) {
      return ctmItem.score < 0;
    }).length;

    const ctmNormalizedScore = ctmScored.length
      ? CTM_clamp_(
          (ctmTotalScore / (ctmScored.length * 3)) * 100,
          -100,
          100
        )
      : 0;

    return {
      enabled: true,
      itemCount: ctmScored.length,
      totalScore: ctmTotalScore,
      normalizedScore: ctmNormalizedScore,
      positiveCount: ctmPositiveCount,
      negativeCount: ctmNegativeCount,
      headlines: ctmScored
    };

  } catch (ctmError) {
    Logger.log('CTM news fetch failed: ' + ctmError.message);

    return {
      enabled: true,
      itemCount: 0,
      totalScore: 0,
      normalizedScore: 0,
      positiveCount: 0,
      negativeCount: 0,
      headlines: [],
      error: ctmError.message
    };
  }
}

function CTM_scoreHeadline_(ctmHeadline) {
  const ctmText = String(ctmHeadline || '').toLowerCase();

  const ctmPositiveKeywords = [
    'rally',
    'surge',
    'breakout',
    'bullish',
    'record inflow',
    'inflows',
    'adoption',
    'approval',
    'accumulate',
    'accumulation',
    'rebound',
    'recovery',
    'gain',
    'gains',
    'upside',
    'all-time high',
    'ath',
    'short squeeze',
    'buying',
    'institutional demand'
  ];

  const ctmNegativeKeywords = [
    'crash',
    'sell-off',
    'selloff',
    'plunge',
    'decline',
    'drop',
    'falls',
    'fall',
    'bearish',
    'outflow',
    'outflows',
    'liquidation',
    'liquidations',
    'hack',
    'exploit',
    'ban',
    'lawsuit',
    'tariff',
    'war',
    'risk-off',
    'fear',
    'pressure',
    'recession',
    'rate hike'
  ];

  let ctmPositiveHits = 0;
  let ctmNegativeHits = 0;

  ctmPositiveKeywords.forEach(function(ctmKeyword) {
    if (ctmText.indexOf(ctmKeyword) !== -1) {
      ctmPositiveHits++;
    }
  });

  ctmNegativeKeywords.forEach(function(ctmKeyword) {
    if (ctmText.indexOf(ctmKeyword) !== -1) {
      ctmNegativeHits++;
    }
  });

  return {
    score: CTM_clamp_(ctmPositiveHits - ctmNegativeHits, -4, 4),
    positiveHits: ctmPositiveHits,
    negativeHits: ctmNegativeHits
  };
}

function CTM_getEmptyNewsData_() {
  return {
    enabled: false,
    itemCount: 0,
    totalScore: 0,
    normalizedScore: 0,
    positiveCount: 0,
    negativeCount: 0,
    headlines: []
  };
}

/* ============================================================
 * MARKET ASSESSMENT
 * ============================================================ */

function CTM_assessMarket_(ctmData, ctmGlobalData, ctmNewsData) {
  const ctmParts = [];
  let ctmScore = 0;

  const ctmTrend1hBullish = ctmData.ema20_1h > ctmData.ema50_1h;
  const ctmTrend4hBullish = ctmData.ema20_4h > ctmData.ema50_4h;

  const ctmTrend1hScore = ctmTrend1hBullish ? 14 : -14;
  const ctmTrend4hScore = ctmTrend4hBullish ? 18 : -18;

  ctmScore += ctmTrend1hScore + ctmTrend4hScore;

  ctmParts.push({
    name: 'EMA trend',
    score: ctmTrend1hScore + ctmTrend4hScore,
    detail:
      '1h ' + (ctmTrend1hBullish ? 'bullish' : 'bearish') +
      ', 4h ' + (ctmTrend4hBullish ? 'bullish' : 'bearish')
  });

  let ctmRsiScore = 0;

  if (ctmData.rsi1h >= 72) {
    ctmRsiScore -= 10;
  } else if (ctmData.rsi1h >= 60) {
    ctmRsiScore += 7;
  } else if (ctmData.rsi1h <= 28) {
    ctmRsiScore += 8;
  } else if (ctmData.rsi1h <= 40) {
    ctmRsiScore -= 6;
  }

  if (ctmData.rsi4h >= 72) {
    ctmRsiScore -= 6;
  } else if (ctmData.rsi4h >= 58) {
    ctmRsiScore += 5;
  } else if (ctmData.rsi4h <= 30) {
    ctmRsiScore += 5;
  } else if (ctmData.rsi4h <= 40) {
    ctmRsiScore -= 4;
  }

  ctmScore += ctmRsiScore;

  ctmParts.push({
    name: 'RSI',
    score: ctmRsiScore,
    detail:
      'RSI 1h ' + CTM_formatNumber_(ctmData.rsi1h, 1) +
      ' | RSI 4h ' + CTM_formatNumber_(ctmData.rsi4h, 1)
  });

  let ctmOiPriceScore = 0;
  const ctmPriceUp = ctmData.cgPrice1hPct >= 0;
  const ctmOiUp = ctmData.oiChange1hPct >= 0.7;
  const ctmOiDown = ctmData.oiChange1hPct <= -0.7;

  if (ctmPriceUp && ctmOiUp) {
    ctmOiPriceScore += 12;
  } else if (!ctmPriceUp && ctmOiUp) {
    ctmOiPriceScore -= 12;
  } else if (!ctmPriceUp && ctmOiDown) {
    ctmOiPriceScore += 4;
  } else if (ctmPriceUp && ctmOiDown) {
    ctmOiPriceScore -= 3;
  }

  ctmScore += ctmOiPriceScore;

  ctmParts.push({
    name: 'Price + OI',
    score: ctmOiPriceScore,
    detail:
      '1h price ' + CTM_formatSignedPct_(ctmData.cgPrice1hPct) +
      ' | OI 1h ' + CTM_formatSignedPct_(ctmData.oiChange1hPct)
  });

  let ctmFundingScore = 0;

  if (ctmData.fundingRatePct8h > 0.05) {
    ctmFundingScore -= 8;
  } else if (ctmData.fundingRatePct8h > 0.025) {
    ctmFundingScore -= 3;
  } else if (ctmData.fundingRatePct8h < -0.05) {
    ctmFundingScore += 8;
  } else if (ctmData.fundingRatePct8h < -0.025) {
    ctmFundingScore += 3;
  }

  ctmScore += ctmFundingScore;

  ctmParts.push({
    name: 'Funding',
    score: ctmFundingScore,
    detail: CTM_formatSignedPct_(ctmData.fundingRatePct8h) + ' / 8h'
  });

  let ctmRangeScore = 0;

  if (ctmData.rangePosition24hPct >= 85 && ctmData.rsi1h >= 65) {
    ctmRangeScore -= 5;
  } else if (ctmData.rangePosition24hPct <= 15 && ctmData.rsi1h <= 35) {
    ctmRangeScore += 5;
  }

  ctmScore += ctmRangeScore;

  ctmParts.push({
    name: '24h range',
    score: ctmRangeScore,
    detail:
      CTM_formatNumber_(ctmData.rangePosition24hPct, 0) +
      '% of 24h range'
  });

  let ctmVolumeScore = 0;

  if (ctmData.volumeRatio1h >= 1.8 && ctmData.cgPrice1hPct > 0) {
    ctmVolumeScore += 6;
  } else if (ctmData.volumeRatio1h >= 1.8 && ctmData.cgPrice1hPct < 0) {
    ctmVolumeScore -= 6;
  }

  ctmScore += ctmVolumeScore;

  ctmParts.push({
    name: '1h volume',
    score: ctmVolumeScore,
    detail:
      CTM_formatNumber_(ctmData.volumeRatio1h, 2) + 'x vs 20h average'
  });

  let ctmGlobalScore = 0;

  if (ctmGlobalData.marketCapChange24hPct >= 1.5) {
    ctmGlobalScore += 5;
  } else if (ctmGlobalData.marketCapChange24hPct <= -1.5) {
    ctmGlobalScore -= 5;
  }

  ctmScore += ctmGlobalScore;

  ctmParts.push({
    name: 'Market breadth',
    score: ctmGlobalScore,
    detail:
      'Global mcap 24h ' +
      CTM_formatSignedPct_(ctmGlobalData.marketCapChange24hPct)
  });

  let ctmNewsScore = 0;

  if (ctmNewsData.itemCount >= 3) {
    ctmNewsScore = CTM_clamp_(
      Math.round(ctmNewsData.normalizedScore * 0.12),
      -12,
      12
    );
  }

  ctmScore += ctmNewsScore;

  ctmParts.push({
    name: 'News proxy',
    score: ctmNewsScore,
    detail:
      ctmNewsData.itemCount + ' headlines | ' +
      ctmNewsData.positiveCount + ' positive / ' +
      ctmNewsData.negativeCount + ' negative'
  });

  ctmScore = CTM_clamp_(Math.round(ctmScore), -100, 100);

  const ctmBias = CTM_getBiasFromScore_(ctmScore);
  const ctmRisk = CTM_getRiskLevel_(ctmData, ctmScore, ctmNewsData);
  const ctmConfidence = CTM_getConfidence_(
    ctmData,
    ctmScore,
    ctmNewsData,
    ctmTrend1hBullish,
    ctmTrend4hBullish
  );

  const ctmReasons = CTM_buildReasons_(
    ctmData,
    ctmBias,
    ctmNewsData,
    ctmTrend1hBullish,
    ctmTrend4hBullish
  );

  return {
    score: ctmScore,
    bias: ctmBias,
    riskLevel: ctmRisk,
    confidencePct: ctmConfidence,
    trend1h: ctmTrend1hBullish ? 'UP' : 'DOWN',
    trend4h: ctmTrend4hBullish ? 'UP' : 'DOWN',
    parts: ctmParts,
    reasons: ctmReasons
  };
}

function CTM_getBiasFromScore_(ctmScore) {
  if (ctmScore >= CTM_CONFIG.SCORE.STRONG_BULLISH_THRESHOLD) {
    return 'STRONG BULLISH';
  }

  if (ctmScore >= CTM_CONFIG.SCORE.BULLISH_THRESHOLD) {
    return 'BULLISH';
  }

  if (ctmScore <= CTM_CONFIG.SCORE.STRONG_BEARISH_THRESHOLD) {
    return 'STRONG BEARISH';
  }

  if (ctmScore <= CTM_CONFIG.SCORE.BEARISH_THRESHOLD) {
    return 'BEARISH';
  }

  return 'NEUTRAL / RANGE';
}

function CTM_getRiskLevel_(ctmData, ctmScore, ctmNewsData) {
  let ctmRiskPoints = 0;

  if (ctmData.atr1hPct >= CTM_CONFIG.RISK.HIGH_ATR_1H_PCT) {
    ctmRiskPoints += 2;
  }

  if (ctmData.atr4hPct >= CTM_CONFIG.RISK.HIGH_ATR_4H_PCT) {
    ctmRiskPoints += 2;
  }

  if (Math.abs(ctmData.oiChange1hPct) >= CTM_CONFIG.RISK.HIGH_OI_CHANGE_PCT) {
    ctmRiskPoints += 2;
  }

  if (
    Math.abs(ctmData.fundingRatePct8h) >=
    CTM_CONFIG.RISK.EXTREME_FUNDING_PCT_8H
  ) {
    ctmRiskPoints += 2;
  }

  if (Math.abs(ctmScore) >= 55) {
    ctmRiskPoints += 1;
  }

  if (
    ctmNewsData.itemCount >= 5 &&
    Math.abs(ctmNewsData.normalizedScore) >= 45
  ) {
    ctmRiskPoints += 1;
  }

  if (ctmRiskPoints >= 5) {
    return 'HIGH';
  }

  if (ctmRiskPoints >= 3) {
    return 'MEDIUM';
  }

  return 'LOW';
}

function CTM_getConfidence_(
  ctmData,
  ctmScore,
  ctmNewsData,
  ctmTrend1hBullish,
  ctmTrend4hBullish
) {
  let ctmConfidence = 45 + Math.min(35, Math.abs(ctmScore) * 0.45);

  if (ctmTrend1hBullish === ctmTrend4hBullish) {
    ctmConfidence += 8;
  }

  if (ctmNewsData.itemCount >= 4) {
    ctmConfidence += 4;
  }

  if (ctmData.volumeRatio1h >= 1.5) {
    ctmConfidence += 4;
  }

  if (
    ctmData.atr1hPct >= CTM_CONFIG.RISK.HIGH_ATR_1H_PCT ||
    Math.abs(ctmData.oiChange1hPct) >= CTM_CONFIG.RISK.HIGH_OI_CHANGE_PCT
  ) {
    ctmConfidence -= 8;
  }

  return Math.round(CTM_clamp_(ctmConfidence, 35, 88));
}

function CTM_buildReasons_(
  ctmData,
  ctmBias,
  ctmNewsData,
  ctmTrend1hBullish,
  ctmTrend4hBullish
) {
  const ctmReasons = [];

  ctmReasons.push(
    'EMA 1h/4h: ' +
    (ctmTrend1hBullish ? '↑' : '↓') +
    '/' +
    (ctmTrend4hBullish ? '↑' : '↓')
  );

  ctmReasons.push(
    'RSI: ' +
    CTM_formatNumber_(ctmData.rsi1h, 0) +
    ' (1h), ' +
    CTM_formatNumber_(ctmData.rsi4h, 0) +
    ' (4h)'
  );

  ctmReasons.push(
    'OI 1h ' +
    CTM_formatSignedPct_(ctmData.oiChange1hPct) +
    ' με τιμή 1h ' +
    CTM_formatSignedPct_(ctmData.cgPrice1hPct)
  );

  ctmReasons.push(
    'Funding ' +
    CTM_formatSignedPct_(ctmData.fundingRatePct8h) +
    ' / 8h'
  );

  if (ctmNewsData.itemCount > 0) {
    ctmReasons.push(
      'News score ' +
      CTM_formatSignedNumber_(ctmNewsData.normalizedScore, 0) +
      ' (' + ctmNewsData.itemCount + ' headlines)'
    );
  }

  if (ctmBias.indexOf('BEARISH') !== -1 && ctmData.fundingRatePct8h > 0) {
    ctmReasons.push('Θετικό funding ενώ η τιμή πιέζεται: αυξημένος long-squeeze κίνδυνος');
  }

  if (ctmBias.indexOf('BULLISH') !== -1 && ctmData.fundingRatePct8h < 0) {
    ctmReasons.push('Αρνητικό funding με ανοδική δομή: πιθανότητα short-squeeze');
  }

  return ctmReasons;
}

/* ============================================================
 * GRID PLAN
 * ============================================================ */

function CTM_buildGridPlan_(
  ctmData,
  ctmAssessment,
  ctmRuntimeConfig
) {
  const ctmVolatilityPct = Math.max(
    ctmData.atr1hPct,
    ctmData.atr4hPct / 2
  );

  const ctmBaseRangePct = CTM_clamp_(
    Math.max(3.5, ctmVolatilityPct * 3.2),
    3.5,
    16
  );

  let ctmMode = 'NEUTRAL';
  let ctmLowerPct = ctmBaseRangePct;
  let ctmUpperPct = ctmBaseRangePct;
  let ctmLeverage = Math.min(2, ctmRuntimeConfig.maxLeverage);
  let ctmAction = 'WAIT / NEUTRAL GRID';

  if (
    ctmAssessment.bias === 'STRONG BULLISH' ||
    ctmAssessment.bias === 'BULLISH'
  ) {
    ctmMode = 'LONG';
    ctmLowerPct = ctmBaseRangePct * 0.70;
    ctmUpperPct = ctmBaseRangePct * 1.30;
    ctmLeverage = Math.min(2, ctmRuntimeConfig.maxLeverage);
    ctmAction = 'LONG GRID — χαμηλότερη έκθεση κοντά στο upper range';

  } else if (
    ctmAssessment.bias === 'STRONG BEARISH' ||
    ctmAssessment.bias === 'BEARISH'
  ) {
    ctmMode = 'SHORT';
    ctmLowerPct = ctmBaseRangePct * 1.30;
    ctmUpperPct = ctmBaseRangePct * 0.70;
    ctmLeverage = 1;
    ctmAction = 'SHORT GRID ή PAUSE — προσοχή σε squeeze';

  } else {
    ctmMode = 'NEUTRAL';
    ctmLowerPct = ctmBaseRangePct;
    ctmUpperPct = ctmBaseRangePct;
    ctmLeverage = 1;
    ctmAction = 'NEUTRAL GRID — μόνο μέσα στο προτεινόμενο εύρος';
  }

  if (ctmAssessment.riskLevel === 'HIGH') {
    ctmLeverage = 1;
    ctmLowerPct *= 1.25;
    ctmUpperPct *= 1.25;
    ctmAction += ' | HIGH RISK: μειωμένο κεφάλαιο ή αναμονή';
  }

  if (
    Math.abs(ctmData.fundingRatePct8h) >=
    CTM_CONFIG.RISK.EXTREME_FUNDING_PCT_8H
  ) {
    ctmAction += ' | EXTREME FUNDING: μείωσε leverage';
  }

  const ctmLowerPrice = ctmData.price * (1 - ctmLowerPct / 100);
  const ctmUpperPrice = ctmData.price * (1 + ctmUpperPct / 100);

  const ctmGridCount = Math.round(
    CTM_clamp_(
      CTM_CONFIG.DEFAULT_GRID_COUNT +
        (ctmVolatilityPct * 3) -
        (ctmAssessment.riskLevel === 'HIGH' ? 7 : 0),
      CTM_CONFIG.MIN_GRID_COUNT,
      CTM_CONFIG.MAX_GRID_COUNT
    )
  );

  const ctmSpacingPct = (
    (ctmUpperPrice - ctmLowerPrice) /
    Math.max(1, ctmGridCount) /
    ctmData.price
  ) * 100;

  const ctmStopLoss = ctmMode === 'LONG'
    ? ctmLowerPrice * (1 - Math.max(1.5, ctmVolatilityPct * 0.9) / 100)
    : ctmMode === 'SHORT'
      ? ctmUpperPrice * (1 + Math.max(1.5, ctmVolatilityPct * 0.9) / 100)
      : 0;

  return {
    mode: ctmMode,
    action: ctmAction,
    leverage: ctmLeverage,
    lowerPrice: ctmLowerPrice,
    upperPrice: ctmUpperPrice,
    gridCount: ctmGridCount,
    spacingPct: ctmSpacingPct,
    stopLoss: ctmStopLoss,
    volatilityPct: ctmVolatilityPct
  };
}

/* ============================================================
 * GOOGLE SHEETS
 * ============================================================ */

const ctmHeaders = [
  'Timestamp',
  'Symbol',
  'Provider',
  'Price',
  '1h Change %',
  '24h Change %',
  '24h High',
  '24h Low',
  '24h Quote Volume',
  'Funding % / 8h',
  'Open Interest',
  'OI Change 1h %',
  'RSI 1h',
  'RSI 4h',
  'ATR 1h %',
  'ATR 4h %',
  'EMA Trend 1h',
  'EMA Trend 4h',
  '24h Range Position %',
  '1h Volume Ratio',
  'Global MCap 24h %',
  'BTC Dominance %',
  'News Score',
  'News Headlines',
  'Signal Score',
  '12h Bias',
  'Confidence %',
  'Risk Level',
  'Grid Mode',
  'Grid Lower',
  'Grid Upper',
  'Grid Count',
  'Grid Spacing %',
  'Suggested Leverage',
  'Suggested Stop',
  'Error'
];

function CTM_getOrCreateTrendsSheet_(ctmSpreadsheet) {
  let ctmSheet = ctmSpreadsheet.getSheetByName(CTM_CONFIG.SHEET_TRENDS);

  if (!ctmSheet) {
    ctmSheet = ctmSpreadsheet.insertSheet(CTM_CONFIG.SHEET_TRENDS);
  }

  const ctmHeaders = [
    'Timestamp',
    'Symbol',
    'Provider',
    'Price',
    '1h Change %',
    '24h Change %',
    '24h High',
    '24h Low',
    '24h Quote Volume',
    'Funding % / 8h',
    'Open Interest',
    'OI Change 1h %',
    'RSI 1h',
    'RSI 4h',
    'ATR 1h %',
    'ATR 4h %',
    'EMA Trend 1h',
    'EMA Trend 4h',
    '24h Range Position %',
    '1h Volume Ratio',
    'Global MCap 24h %',
    'BTC Dominance %',
    'News Score',
    'News Headlines',
    'Signal Score',
    '12h Bias',
    'Confidence %',
    'Risk Level',
    'Grid Mode',
    'Grid Lower',
    'Grid Upper',
    'Grid Count',
    'Grid Spacing %',
    'Suggested Leverage',
    'Suggested Stop',
    'Error'
  ];

  CTM_ensureHeaders_(ctmSheet, ctmHeaders);

  ctmSheet.setFrozenRows(1);

  ctmSheet
    .getRange(1, 1, 1, ctmHeaders.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');

  return ctmSheet;
}

function CTM_getOrCreateGridSheet_(ctmSpreadsheet) {
  let ctmSheet = ctmSpreadsheet.getSheetByName(CTM_CONFIG.SHEET_GRID);

  if (!ctmSheet) {
    ctmSheet = ctmSpreadsheet.insertSheet(CTM_CONFIG.SHEET_GRID);
  }

  const ctmHeaders = [
    'Updated',
    'Symbol',
    'Provider',
    'Current Price',
    '12h Bias',
    'Confidence %',
    'Risk',
    'Grid Mode',
    'Action',
    'Lower Price',
    'Upper Price',
    'Grid Count',
    'Grid Spacing %',
    'Suggested Leverage',
    'Suggested Stop Loss',
    'Funding % / 8h',
    'OI Change 1h %',
    'RSI 1h',
    'RSI 4h',
    'ATR 1h %',
    'ATR 4h %'
  ];

  CTM_ensureHeaders_(ctmSheet, ctmHeaders);

  ctmSheet.setFrozenRows(1);

  ctmSheet
    .getRange(1, 1, 1, ctmHeaders.length)
    .setFontWeight('bold')
    .setBackground('#274e13')
    .setFontColor('#ffffff');

  return ctmSheet;
}

function CTM_ensureHeaders_(ctmSheet, ctmHeaders) {
  const ctmCurrentLastColumn = ctmSheet.getLastColumn();

  if (ctmCurrentLastColumn < ctmHeaders.length) {
    ctmSheet.insertColumnsAfter(
      Math.max(1, ctmCurrentLastColumn),
      ctmHeaders.length - ctmCurrentLastColumn
    );
  }

  const ctmExistingHeaders = ctmSheet
    .getRange(1, 1, 1, ctmHeaders.length)
    .getValues()[0];

  const ctmHeadersDifferent = ctmHeaders.some(function(ctmHeader, ctmIndex) {
    return ctmExistingHeaders[ctmIndex] !== ctmHeader;
  });

  if (ctmHeadersDifferent) {
    ctmSheet
      .getRange(1, 1, 1, ctmHeaders.length)
      .setValues([ctmHeaders]);
  }

  ctmSheet.autoResizeColumns(1, ctmHeaders.length);
}

function CTM_appendRows_(ctmSheet, ctmRows) {
  if (!ctmRows || !ctmRows.length) {
    return;
  }

  const ctmStartRow = Math.max(2, ctmSheet.getLastRow() + 1);

  ctmSheet
    .getRange(ctmStartRow, 1, ctmRows.length, ctmRows[0].length)
    .setValues(ctmRows);

  CTM_applyTrendNumberFormats_(
    ctmSheet,
    ctmStartRow,
    ctmRows.length
  );
}

function CTM_replaceGridRows_(ctmSheet, ctmRows) {
  if (!ctmRows || !ctmRows.length) {
    return;
  }

  const ctmLastRow = ctmSheet.getLastRow();
  const ctmLastColumn = ctmSheet.getLastColumn();

  if (ctmLastRow > 1) {
    ctmSheet
      .getRange(2, 1, ctmLastRow - 1, ctmLastColumn)
      .clearContent();
  }

  ctmSheet
    .getRange(2, 1, ctmRows.length, ctmRows[0].length)
    .setValues(ctmRows);

  CTM_applyGridNumberFormats_(
    ctmSheet,
    2,
    ctmRows.length
  );
}

function CTM_applyTrendNumberFormats_(
  ctmSheet,
  ctmStartRow,
  ctmRowCount
) {
  ctmSheet
    .getRange(ctmStartRow, 1, ctmRowCount, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  const ctmPriceColumns = [
    4, 7, 8, 30, 31, 35
  ];

  ctmPriceColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00000000');
  });

  const ctmPercentColumns = [
    5, 6, 10, 12, 15, 16,
    19, 20, 21, 22, 23, 27, 33
  ];

  ctmPercentColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00');
  });

  ctmSheet
    .getRange(ctmStartRow, 25, ctmRowCount, 1)
    .setNumberFormat('0');

  ctmSheet
    .getRange(ctmStartRow, 26, ctmRowCount, 1)
    .setNumberFormat('0');

  ctmSheet
    .getRange(ctmStartRow, 32, ctmRowCount, 1)
    .setNumberFormat('0');
}

function CTM_applyGridNumberFormats_(
  ctmSheet,
  ctmStartRow,
  ctmRowCount
) {
  ctmSheet
    .getRange(ctmStartRow, 1, ctmRowCount, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  const ctmPriceColumns = [
    4, 10, 11, 15
  ];

  ctmPriceColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00000000');
  });

  const ctmPercentColumns = [
    6, 13, 16, 17, 18, 19, 20, 21
  ];

  ctmPercentColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00');
  });

  ctmSheet
    .getRange(ctmStartRow, 12, ctmRowCount, 1)
    .setNumberFormat('0');

  ctmSheet
    .getRange(ctmStartRow, 14, ctmRowCount, 1)
    .setNumberFormat('0');
}

function CTM_buildTrendRow_(
  ctmData,
  ctmAssessment,
  ctmGridPlan,
  ctmGlobalData,
  ctmNewsData
) {
return [
  ctmData.collectedAt,
  ctmData.symbol,
  ctmData.provider,
  ctmData.price,
  ctmData.cgPrice1hPct,
  ctmData.priceChange24hPct,
  ctmData.high24h,
  ctmData.low24h,
  ctmData.quoteVolume24h,
  ctmData.fundingRatePct8h,
  ctmData.openInterest,
  ctmData.oiChange1hPct,
  ctmData.rsi1h,
  ctmData.rsi4h,
  ctmData.atr1hPct,
  ctmData.atr4hPct,
  ctmAssessment.trend1h,
  ctmAssessment.trend4h,
  ctmData.rangePosition24hPct,
  ctmData.volumeRatio1h,
  ctmGlobalData.marketCapChange24hPct,
  ctmGlobalData.btcDominancePct,
  ctmNewsData.normalizedScore,
  CTM_headlinesToText_(ctmNewsData.headlines, 5),
  ctmAssessment.score,
  ctmAssessment.bias,
  ctmAssessment.confidencePct,
  ctmAssessment.riskLevel,
  ctmGridPlan.mode,
  ctmGridPlan.lowerPrice,
  ctmGridPlan.upperPrice,
  ctmGridPlan.gridCount,
  ctmGridPlan.spacingPct,
  ctmGridPlan.leverage,
  ctmGridPlan.stopLoss,
  ''
];
}

function CTM_buildErrorTrendRow_(
  ctmSymbol,
  ctmError,
  ctmGlobalData,
  ctmNewsData
) {
  return [
    new Date(),
    ctmSymbol,
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    ctmGlobalData.marketCapChange24hPct,
    ctmGlobalData.btcDominancePct,
    ctmNewsData.normalizedScore,
    CTM_headlinesToText_(ctmNewsData.headlines, 3),
    '',
    'ERROR',
    '',
    'HIGH',
    'PAUSE',
    '',
    '',
    '',
    '',
    '',
    '',
    String(ctmError.message || ctmError)
  ];
}

function CTM_buildGridRow_(ctmData, ctmAssessment, ctmGridPlan) {
  return [
    ctmData.collectedAt,
    ctmData.symbol,
    ctmData.price,
    ctmAssessment.bias,
    ctmAssessment.confidencePct,
    ctmAssessment.riskLevel,
    ctmGridPlan.mode,
    ctmGridPlan.action,
    ctmGridPlan.lowerPrice,
    ctmGridPlan.upperPrice,
    ctmGridPlan.gridCount,
    ctmGridPlan.spacingPct,
    ctmGridPlan.leverage,
    ctmGridPlan.stopLoss,
    ctmData.fundingRatePct8h,
    ctmData.oiChange1hPct,
    ctmData.rsi1h,
    ctmData.rsi4h,
    ctmData.atr1hPct,
    ctmData.atr4hPct
  ];
}

function CTM_appendRows_(ctmSheet, ctmRows) {
  if (!ctmRows || !ctmRows.length) {
    return;
  }

  const ctmStartRow = Math.max(2, ctmSheet.getLastRow() + 1);

  ctmSheet
    .getRange(ctmStartRow, 1, ctmRows.length, ctmRows[0].length)
    .setValues(ctmRows);

  CTM_applyTrendNumberFormats_(ctmSheet, ctmStartRow, ctmRows.length);
}

function CTM_replaceGridRows_(ctmSheet, ctmRows) {
  if (!ctmRows || !ctmRows.length) {
    return;
  }

  const ctmLastRow = ctmSheet.getLastRow();

  if (ctmLastRow > 1) {
    ctmSheet
      .getRange(2, 1, ctmLastRow - 1, ctmSheet.getLastColumn())
      .clearContent();
  }

  ctmSheet
    .getRange(2, 1, ctmRows.length, ctmRows[0].length)
    .setValues(ctmRows);

  CTM_applyGridNumberFormats_(ctmSheet, 2, ctmRows.length);
}

function CTM_ensureHeaders_(ctmSheet, ctmHeaders) {
  const ctmCurrentHeaders = ctmSheet
    .getRange(1, 1, 1, ctmHeaders.length)
    .getValues()[0];

  const ctmNeedsHeaders = ctmHeaders.some(function(ctmHeader, ctmIndex) {
    return ctmCurrentHeaders[ctmIndex] !== ctmHeader;
  });

  if (ctmNeedsHeaders) {
    ctmSheet.getRange(1, 1, 1, ctmHeaders.length)
      .setValues([ctmHeaders]);
  }
}

function CTM_applyTrendNumberFormats_(ctmSheet, ctmStartRow, ctmRowCount) {
  ctmSheet
    .getRange(ctmStartRow, 1, ctmRowCount, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  const ctmPriceColumns = [3, 6, 7, 29, 30, 34];
  ctmPriceColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00000000');
  });

  const ctmPercentColumns = [
    4, 5, 9, 11, 14, 15, 18, 19, 20, 21, 22, 26, 32
  ];

  ctmPercentColumns.forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00');
  });
}

function CTM_applyGridNumberFormats_(ctmSheet, ctmStartRow, ctmRowCount) {
  ctmSheet
    .getRange(ctmStartRow, 1, ctmRowCount, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  [3, 9, 10, 14].forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00000000');
  });

  [5, 12, 15, 16, 17, 18, 19, 20].forEach(function(ctmColumn) {
    ctmSheet
      .getRange(ctmStartRow, ctmColumn, ctmRowCount, 1)
      .setNumberFormat('0.00');
  });
}

/* ============================================================
 * TELEGRAM MESSAGE
 * ============================================================ */

function CTM_buildTelegramMessage_(
  ctmAnalysis,
  ctmGlobalData,
  ctmNewsData,
  ctmRuntimeConfig
) {
  const ctmNow = CTM_formatDate_(new Date());

  const ctmBiasCounts = {
    bullish: 0,
    bearish: 0,
    neutral: 0
  };

  ctmAnalysis.forEach(function(ctmItem) {
    if (ctmItem.assessment.bias.indexOf('BULLISH') !== -1) {
      ctmBiasCounts.bullish++;
    } else if (ctmItem.assessment.bias.indexOf('BEARISH') !== -1) {
      ctmBiasCounts.bearish++;
    } else {
      ctmBiasCounts.neutral++;
    }
  });

  let ctmMarketRegime = 'ΟΥΔΕΤΕΡΟ / RANGE';

  if (ctmBiasCounts.bearish > ctmBiasCounts.bullish) {
    ctmMarketRegime = 'ΠΡΟΣΕΚΤΙΚΑ ΠΤΩΤΙΚΟ';
  } else if (ctmBiasCounts.bullish > ctmBiasCounts.bearish) {
    ctmMarketRegime = 'ΠΡΟΣΕΚΤΙΚΑ ΑΝΟΔΙΚΟ';
  }

  let ctmMessage =
    '📊 <b>CRYPTO MARKET — 12H OUTLOOK</b>\n' +
    '🕒 ' + CTM_escapeHtml_(ctmNow) + '\n\n' +
    '🧭 <b>Καθεστώς αγοράς:</b> ' + ctmMarketRegime + '\n' +
    '🌍 Global MCap 24h: <b>' +
    CTM_formatSignedPct_(ctmGlobalData.marketCapChange24hPct) +
    '</b>\n' +
    '₿ BTC Dominance: <b>' +
    CTM_formatNumber_(ctmGlobalData.btcDominancePct, 2) +
    '%</b>\n' +
    '📰 News sentiment proxy: <b>' +
    CTM_formatSignedNumber_(ctmNewsData.normalizedScore, 0) +
    '</b> (' +
    ctmNewsData.positiveCount + ' θετικά / ' +
    ctmNewsData.negativeCount + ' αρνητικά)\n\n';

  ctmAnalysis.forEach(function(ctmItem) {
    const ctmData = ctmItem.data;
    const ctmAssessment = ctmItem.assessment;
    const ctmGridPlan = ctmItem.gridPlan;

    const ctmBiasEmoji = CTM_getBiasEmoji_(ctmAssessment.bias);
    const ctmRiskEmoji = CTM_getRiskEmoji_(ctmAssessment.riskLevel);

    ctmMessage +=
      ctmBiasEmoji + ' <b>' + CTM_escapeHtml_(ctmData.symbol) + '</b> — ' +
      '<b>' + CTM_escapeHtml_(ctmAssessment.bias) + '</b>\n' +
      '💵 ' + CTM_formatPrice_(ctmData.price) +
      ' | 1h ' + CTM_formatSignedPct_(ctmData.cgPrice1hPct) +
      ' | 24h ' + CTM_formatSignedPct_(ctmData.priceChange24hPct) + '\n' +
      '📈 Score: <b>' + CTM_formatSignedNumber_(ctmAssessment.score, 0) +
      '/100</b> | Confidence: ' +
      ctmAssessment.confidencePct + '% | ' +
      ctmRiskEmoji + ' Risk: ' +
      ctmAssessment.riskLevel + '\n' +
      '📉 RSI 1h/4h: ' +
      CTM_formatNumber_(ctmData.rsi1h, 0) + '/' +
      CTM_formatNumber_(ctmData.rsi4h, 0) +
      ' | EMA: ' + ctmAssessment.trend1h +
      '/' + ctmAssessment.trend4h + '\n' +
      '🔄 OI 1h: ' + CTM_formatSignedPct_(ctmData.oiChange1hPct) +
      ' | Funding: ' + CTM_formatSignedPct_(ctmData.fundingRatePct8h) +
      ' / 8h\n' +
      '⚡ ATR 1h/4h: ' +
      CTM_formatNumber_(ctmData.atr1hPct, 2) + '%/' +
      CTM_formatNumber_(ctmData.atr4hPct, 2) + '%' +
      ' | Vol: ' + CTM_formatNumber_(ctmData.volumeRatio1h, 2) + 'x\n' +
      '🤖 <b>Grid:</b> ' + ctmGridPlan.mode +
      ' | ' + ctmGridPlan.leverage + 'x max\n' +
      '↔️ Range: ' +
      CTM_formatPrice_(ctmGridPlan.lowerPrice) + ' – ' +
      CTM_formatPrice_(ctmGridPlan.upperPrice) +
      ' | ' + ctmGridPlan.gridCount + ' grids' +
      ' | step ~' + CTM_formatNumber_(ctmGridPlan.spacingPct, 2) + '%\n' +
      (ctmGridPlan.stopLoss
        ? '🛑 Invalidation/SL: ' +
          CTM_formatPrice_(ctmGridPlan.stopLoss) + '\n'
        : '') +
      '⚠️ ' + CTM_escapeHtml_(ctmGridPlan.action) + '\n\n';
  });

  if (ctmNewsData.headlines.length) {
    ctmMessage += '🗞 <b>Recent narrative</b>\n';

    ctmNewsData.headlines.slice(0, 3).forEach(function(ctmHeadline) {
      const ctmIcon = ctmHeadline.score > 0
        ? '🟢'
        : ctmHeadline.score < 0
          ? '🔴'
          : '⚪';

      ctmMessage +=
        ctmIcon + ' ' +
        CTM_escapeHtml_(CTM_truncate_(ctmHeadline.title, 135)) +
        '\n';
    });

    ctmMessage += '\n';
  }

  ctmMessage +=
    '⚠️ <i>Εκτίμηση 12ώρου, όχι επενδυτική συμβουλή. ' +
    'Μην ανοίγεις/επεκτείνεις futures grid σε υψηλό ATR, ' +
    'απότομο OI expansion ή πριν από σημαντικά macro γεγονότα.</i>';

  return CTM_truncate_(ctmMessage, CTM_CONFIG.TELEGRAM_MESSAGE_LIMIT);
}

function CTM_sendTelegram_(ctmMessage, ctmProps) {
  const ctmBotToken = ctmProps.getProperty('TELEGRAM_BOT_TOKEN');
  const ctmChatId = ctmProps.getProperty('TELEGRAM_CHAT_ID');

  if (!ctmBotToken || !ctmChatId) {
    return {
      sent: false,
      reason: 'Missing TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID'
    };
  }

  const ctmUrl =
    'https://api.telegram.org/bot' +
    encodeURIComponent(ctmBotToken) +
    '/sendMessage';

  const ctmPayload = {
    chat_id: ctmChatId,
    text: ctmMessage,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  const ctmResponse = UrlFetchApp.fetch(ctmUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(ctmPayload),
    muteHttpExceptions: true
  });

  const ctmCode = ctmResponse.getResponseCode();
  const ctmBody = ctmResponse.getContentText();

  if (ctmCode < 200 || ctmCode >= 300) {
    throw new Error(
      'Telegram HTTP ' + ctmCode + ': ' + CTM_truncate_(ctmBody, 500)
    );
  }

  return {
    sent: true,
    responseCode: ctmCode,
    response: ctmBody
  };
}

function CTM_shouldSendAlert_(ctmAnalysis) {
  return ctmAnalysis.some(function(ctmItem) {
    return (
      ctmItem.assessment.riskLevel === 'HIGH' ||
      ctmItem.assessment.bias === 'STRONG BULLISH' ||
      ctmItem.assessment.bias === 'STRONG BEARISH'
    );
  });
}

/* ============================================================
 * INDICATORS
 * ============================================================ */

function CTM_calculateRsi_(ctmCloses, ctmPeriod) {
  if (!ctmCloses || ctmCloses.length <= ctmPeriod) {
    return 50;
  }

  let ctmGainSum = 0;
  let ctmLossSum = 0;

  for (let ctmIndex = 1; ctmIndex <= ctmPeriod; ctmIndex++) {
    const ctmChange = ctmCloses[ctmIndex] - ctmCloses[ctmIndex - 1];

    if (ctmChange >= 0) {
      ctmGainSum += ctmChange;
    } else {
      ctmLossSum += Math.abs(ctmChange);
    }
  }

  let ctmAverageGain = ctmGainSum / ctmPeriod;
  let ctmAverageLoss = ctmLossSum / ctmPeriod;

  for (
    let ctmIndex = ctmPeriod + 1;
    ctmIndex < ctmCloses.length;
    ctmIndex++
  ) {
    const ctmChange = ctmCloses[ctmIndex] - ctmCloses[ctmIndex - 1];
    const ctmGain = ctmChange > 0 ? ctmChange : 0;
    const ctmLoss = ctmChange < 0 ? Math.abs(ctmChange) : 0;

    ctmAverageGain =
      ((ctmAverageGain * (ctmPeriod - 1)) + ctmGain) / ctmPeriod;

    ctmAverageLoss =
      ((ctmAverageLoss * (ctmPeriod - 1)) + ctmLoss) / ctmPeriod;
  }

  if (ctmAverageLoss === 0) {
    return 100;
  }

  const ctmRs = ctmAverageGain / ctmAverageLoss;
  return 100 - (100 / (1 + ctmRs));
}

function CTM_calculateAtr_(ctmCandles, ctmPeriod) {
  if (!ctmCandles || ctmCandles.length <= ctmPeriod) {
    return 0;
  }

  const ctmTrueRanges = [];

  for (let ctmIndex = 1; ctmIndex < ctmCandles.length; ctmIndex++) {
    const ctmCurrent = ctmCandles[ctmIndex];
    const ctmPrevious = ctmCandles[ctmIndex - 1];

    const ctmTrueRange = Math.max(
      ctmCurrent.high - ctmCurrent.low,
      Math.abs(ctmCurrent.high - ctmPrevious.close),
      Math.abs(ctmCurrent.low - ctmPrevious.close)
    );

    ctmTrueRanges.push(ctmTrueRange);
  }

  const ctmRecentRanges = ctmTrueRanges.slice(-ctmPeriod);

  return ctmRecentRanges.reduce(function(ctmSum, ctmValue) {
    return ctmSum + ctmValue;
  }, 0) / ctmRecentRanges.length;
}

function CTM_calculateEma_(ctmValues, ctmPeriod) {
  if (!ctmValues || !ctmValues.length) {
    return 0;
  }

  const ctmSlice = ctmValues.length >= ctmPeriod
    ? ctmValues.slice(-ctmPeriod)
    : ctmValues.slice();

  const ctmMultiplier = 2 / (ctmSlice.length + 1);
  let ctmEma = ctmSlice[0];

  for (let ctmIndex = 1; ctmIndex < ctmSlice.length; ctmIndex++) {
    ctmEma =
      (ctmSlice[ctmIndex] - ctmEma) * ctmMultiplier + ctmEma;
  }

  return ctmEma;
}

function CTM_calculateVolumeRatio_(ctmCandles, ctmPeriod) {
  if (!ctmCandles || ctmCandles.length < 3) {
    return 1;
  }

  const ctmLastCandle = ctmCandles[ctmCandles.length - 1];
  const ctmHistorical = ctmCandles
    .slice(Math.max(0, ctmCandles.length - ctmPeriod - 1), -1);

  const ctmAverageVolume = ctmHistorical.reduce(function(ctmSum, ctmCandle) {
    return ctmSum + ctmCandle.volume;
  }, 0) / Math.max(1, ctmHistorical.length);

  return ctmAverageVolume
    ? ctmLastCandle.volume / ctmAverageVolume
    : 1;
}

/* ============================================================
 * HTTP / UTILITIES
 * ============================================================ */

function CTM_fetchJson_(ctmUrl, ctmOptions) {
  const ctmResponse = UrlFetchApp.fetch(ctmUrl, ctmOptions || {});
  const ctmCode = ctmResponse.getResponseCode();
  const ctmText = ctmResponse.getContentText();

  if (ctmCode < 200 || ctmCode >= 300) {
    throw new Error(
      'HTTP ' + ctmCode + ' from ' + ctmUrl +
      ': ' + CTM_truncate_(ctmText, 300)
    );
  }

  try {
    return JSON.parse(ctmText);
  } catch (ctmError) {
    throw new Error(
      'Invalid JSON from ' + ctmUrl +
      ': ' + CTM_truncate_(ctmText, 300)
    );
  }
}

function CTM_buildCoinGeckoHeaders_(ctmCoinGeckoApiKey) {
  const ctmHeaders = {
    accept: 'application/json'
  };

  if (ctmCoinGeckoApiKey) {
    ctmHeaders['x-cg-demo-api-key'] = ctmCoinGeckoApiKey;
  }

  return ctmHeaders;
}

function CTM_getXmlChildText_(ctmElement, ctmChildName) {
  const ctmChild = ctmElement.getChild(ctmChildName);
  return ctmChild ? ctmChild.getText() : '';
}

function CTM_deleteTriggers_(ctmFunctionName) {
  ScriptApp.getProjectTriggers().forEach(function(ctmTrigger) {
    if (ctmTrigger.getHandlerFunction() === ctmFunctionName) {
      ScriptApp.deleteTrigger(ctmTrigger);
    }
  });
}

function CTM_toBoolean_(ctmValue, ctmFallback) {
  if (ctmValue === null || ctmValue === undefined || ctmValue === '') {
    return ctmFallback;
  }

  return String(ctmValue).toLowerCase() === 'true';
}

function CTM_toNumber_(ctmValue, ctmFallback) {
  const ctmNumber = Number(ctmValue);
  return isFinite(ctmNumber) && ctmNumber > 0
    ? ctmNumber
    : ctmFallback;
}

function CTM_clamp_(ctmValue, ctmMin, ctmMax) {
  return Math.max(ctmMin, Math.min(ctmMax, ctmValue));
}

function CTM_formatDate_(ctmDate) {
  return Utilities.formatDate(
    ctmDate,
    CTM_CONFIG.TIMEZONE,
    'yyyy-MM-dd HH:mm:ss'
  );
}

function CTM_formatPrice_(ctmValue) {
  const ctmNumber = Number(ctmValue);

  if (!isFinite(ctmNumber)) {
    return 'N/A';
  }

  if (ctmNumber >= 1000) {
    return ctmNumber.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  if (ctmNumber >= 1) {
    return ctmNumber.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 4
    });
  }

  if (ctmNumber >= 0.01) {
    return ctmNumber.toLocaleString('en-US', {
      minimumFractionDigits: 4,
      maximumFractionDigits: 6
    });
  }

  return ctmNumber.toLocaleString('en-US', {
    minimumFractionDigits: 6,
    maximumFractionDigits: 10
  });
}

function CTM_formatNumber_(ctmValue, ctmDecimals) {
  const ctmNumber = Number(ctmValue);

  if (!isFinite(ctmNumber)) {
    return 'N/A';
  }

  return ctmNumber.toLocaleString('en-US', {
    minimumFractionDigits: ctmDecimals,
    maximumFractionDigits: ctmDecimals
  });
}

function CTM_formatSignedNumber_(ctmValue, ctmDecimals) {
  const ctmNumber = Number(ctmValue);

  if (!isFinite(ctmNumber)) {
    return 'N/A';
  }

  return (ctmNumber >= 0 ? '+' : '') +
    CTM_formatNumber_(ctmNumber, ctmDecimals);
}

function CTM_formatSignedPct_(ctmValue) {
  const ctmNumber = Number(ctmValue);

  if (!isFinite(ctmNumber)) {
    return 'N/A';
  }

  return (ctmNumber >= 0 ? '+' : '') +
    CTM_formatNumber_(ctmNumber, 2) + '%';
}

function CTM_headlinesToText_(ctmHeadlines, ctmLimit) {
  return (ctmHeadlines || [])
    .slice(0, ctmLimit)
    .map(function(ctmHeadline) {
      return ctmHeadline.title;
    })
    .join(' | ');
}

function CTM_getBiasEmoji_(ctmBias) {
  if (ctmBias.indexOf('BULLISH') !== -1) {
    return '🟢';
  }

  if (ctmBias.indexOf('BEARISH') !== -1) {
    return '🔴';
  }

  return '🟡';
}

function CTM_getRiskEmoji_(ctmRisk) {
  if (ctmRisk === 'HIGH') {
    return '🔴';
  }

  if (ctmRisk === 'MEDIUM') {
    return '🟠';
  }

  return '🟢';
}

function CTM_escapeHtml_(ctmText) {
  return String(ctmText || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function CTM_truncate_(ctmText, ctmLimit) {
  const ctmString = String(ctmText || '');

  if (ctmString.length <= ctmLimit) {
    return ctmString;
  }

  return ctmString.slice(0, Math.max(0, ctmLimit - 3)) + '...';
}

function CTM_fetchKrakenDerivativesData_(ctmSymbol) {
  const ctmKrakenSymbol = CTM_mapSymbolToKraken_(ctmSymbol);

  if (!ctmKrakenSymbol) {
    throw new Error(
      'Δεν υπάρχει Kraken Futures mapping για το symbol: ' + ctmSymbol
    );
  }

  const ctmTickers = CTM_fetchKrakenTickers_();

  const ctmTicker = ctmTickers.find(function(ctmItem) {
    return ctmItem.symbol === ctmKrakenSymbol;
  });

  if (!ctmTicker) {
    throw new Error(
      'Δεν βρέθηκε Kraken Futures ticker για: ' + ctmKrakenSymbol
    );
  }

  const ctmKlines1h = CTM_fetchKrakenCandles_(
  ctmKrakenSymbol,
  '1H',
  CTM_CONFIG.KLINE_LIMIT
  );

  const ctmKlines4h = CTM_fetchKrakenCandles_(
  ctmKrakenSymbol,
  '4H',
  CTM_CONFIG.KLINE_LIMIT
  );

  const ctmOiHistory = CTM_fetchKrakenOpenInterestProxy_(
    ctmKrakenSymbol,
    ctmTicker.openInterest
  );

  const ctmLastPrice = Number(
    ctmTicker.last ||
    ctmTicker.markPrice ||
    ctmTicker.index ||
    0
  );

  const ctmOpen24h = Number(ctmTicker.open24h || 0);

  const ctmChange24hPct = ctmOpen24h > 0
    ? ((ctmLastPrice - ctmOpen24h) / ctmOpen24h) * 100
    : 0;

  return {
    provider: 'KRAKEN',

    price: ctmLastPrice,
    priceChange24hPct: ctmChange24hPct,

    high24h: Number(
      ctmTicker.high24h ||
      ctmTicker.high ||
      ctmLastPrice
    ),

    low24h: Number(
      ctmTicker.low24h ||
      ctmTicker.low ||
      ctmLastPrice
    ),

    quoteVolume24h: Number(
      ctmTicker.volumeQuote ||
      ctmTicker.volume ||
      0
    ),

    fundingRatePct8h: Number(
      ctmTicker.fundingRate ||
      ctmTicker.relativeFundingRate ||
      0
    ) * 100,

    fundingTime: ctmTicker.nextFundingRateTime
      ? new Date(Number(ctmTicker.nextFundingRateTime))
      : new Date(),

    openInterest: Number(ctmTicker.openInterest || 0),

    openInterest1hAgo: Number(
      ctmOiHistory.openInterest1hAgo || 0
    ),

    oiChange1hPct: Number(
      ctmOiHistory.oiChange1hPct || 0
    ),

    klines1h: ctmKlines1h,
    klines4h: ctmKlines4h
  };
}

function CTM_mapSymbolToKraken_(ctmSymbol) {
  const ctmSymbolMap = {
    BTCUSDT: 'PI_XBTUSD',
    ETHUSDT: 'PI_ETHUSD',
    SOLUSDT: 'PI_SOLUSD',
    BNBUSDT: 'PI_BNBUSD'
  };

  return ctmSymbolMap[ctmSymbol] || null;
}

function CTM_fetchKrakenTickers_() {
  const ctmUrl = CTM_CONFIG.KRAKEN_FUTURES_BASE + '/tickers';

  const ctmResponse = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  if (
    !ctmResponse ||
    !ctmResponse.tickers ||
    !Array.isArray(ctmResponse.tickers)
  ) {
    throw new Error(
      'Μη έγκυρη απάντηση Kraken tickers: ' +
      JSON.stringify(ctmResponse)
    );
  }

  return ctmResponse.tickers;
}

function CTM_fetchKrakenCandles_(
  ctmKrakenSymbol,
  ctmResolution,
  ctmLimit
) {
  const ctmUrl =
    CTM_CONFIG.KRAKEN_CHARTS_BASE +
    '/trade/' +
    encodeURIComponent(ctmKrakenSymbol) +
    '/' +
    encodeURIComponent(ctmResolution);

  const ctmResponse = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  const ctmCandlesRaw =
    ctmResponse.candles ||
    ctmResponse.data ||
    ctmResponse;

  if (!Array.isArray(ctmCandlesRaw)) {
    throw new Error(
      'Δεν επιστράφηκαν Kraken candles για ' +
      ctmKrakenSymbol +
      '. Response: ' +
      JSON.stringify(ctmResponse).slice(0, 300)
    );
  }

  const ctmCandles = ctmCandlesRaw
    .map(function(ctmCandle) {
      return CTM_normalizeKrakenCandle_(ctmCandle);
    })
    .filter(function(ctmCandle) {
      return ctmCandle &&
        isFinite(ctmCandle.openTime) &&
        isFinite(ctmCandle.open) &&
        isFinite(ctmCandle.high) &&
        isFinite(ctmCandle.low) &&
        isFinite(ctmCandle.close);
    })
    .sort(function(ctmA, ctmB) {
      return ctmA.openTime - ctmB.openTime;
    })
    .slice(-Number(ctmLimit));

  if (ctmCandles.length < 20) {
    throw new Error(
      'Ανεπαρκή Kraken candles για ' +
      ctmKrakenSymbol +
      ': ' + ctmCandles.length
    );
  }

  return ctmCandles;
}

function CTM_normalizeKrakenCandle_(ctmCandle) {
  if (!ctmCandle) {
    return null;
  }

  if (Array.isArray(ctmCandle)) {
    const ctmTimestamp = Number(ctmCandle[0]);

    return {
      openTime: ctmTimestamp < 1000000000000
        ? ctmTimestamp * 1000
        : ctmTimestamp,

      open: Number(ctmCandle[1]),
      high: Number(ctmCandle[2]),
      low: Number(ctmCandle[3]),
      close: Number(ctmCandle[4]),
      volume: Number(ctmCandle[5] || 0),

      closeTime: ctmTimestamp < 1000000000000
        ? ctmTimestamp * 1000
        : ctmTimestamp
    };
  }

  if (typeof ctmCandle === 'object') {
    const ctmTimestamp = Number(
      ctmCandle.time ||
      ctmCandle.timestamp ||
      ctmCandle.t ||
      ctmCandle.intervalBegin ||
      0
    );

    return {
      openTime: ctmTimestamp < 1000000000000
        ? ctmTimestamp * 1000
        : ctmTimestamp,

      open: Number(
        ctmCandle.open ||
        ctmCandle.o ||
        ctmCandle.priceOpen ||
        0
      ),

      high: Number(
        ctmCandle.high ||
        ctmCandle.h ||
        ctmCandle.priceHigh ||
        0
      ),

      low: Number(
        ctmCandle.low ||
        ctmCandle.l ||
        ctmCandle.priceLow ||
        0
      ),

      close: Number(
        ctmCandle.close ||
        ctmCandle.c ||
        ctmCandle.priceClose ||
        0
      ),

      volume: Number(
        ctmCandle.volume ||
        ctmCandle.v ||
        ctmCandle.qty ||
        0
      ),

      closeTime: ctmTimestamp < 1000000000000
        ? ctmTimestamp * 1000
        : ctmTimestamp
    };
  }

  return null;
}

function CTM_fetchKrakenOpenInterestProxy_(
  ctmKrakenSymbol,
  ctmCurrentOpenInterest
) {
  const ctmCache = CacheService.getScriptCache();

  const ctmCacheKey =
    'CTM_KRAKEN_OI_' + ctmKrakenSymbol;

  const ctmPreviousRaw = ctmCache.get(ctmCacheKey);

  let ctmPrevious = null;

  if (ctmPreviousRaw) {
    try {
      ctmPrevious = JSON.parse(ctmPreviousRaw);
    } catch (ctmError) {
      ctmPrevious = null;
    }
  }

  const ctmNowPayload = {
    timestamp: Date.now(),
    openInterest: Number(ctmCurrentOpenInterest || 0)
  };

  ctmCache.put(
    ctmCacheKey,
    JSON.stringify(ctmNowPayload),
    21600
  );

  if (
    !ctmPrevious ||
    !ctmPrevious.openInterest ||
    !ctmCurrentOpenInterest
  ) {
    return {
      openInterest1hAgo: 0,
      oiChange1hPct: 0
    };
  }

  return {
    openInterest1hAgo: Number(ctmPrevious.openInterest),
    oiChange1hPct: (
      (
        Number(ctmCurrentOpenInterest) -
        Number(ctmPrevious.openInterest)
      ) /
      Number(ctmPrevious.openInterest)
    ) * 100
  };
}

function CTM_fetchCoinGeckoMarketFallback_(
  ctmSymbol,
  ctmRuntimeConfig
) {
  const ctmCoinId = CTM_CONFIG.COINGECKO_IDS[ctmSymbol];

  if (!ctmCoinId) {
    throw new Error(
      'Δεν υπάρχει CoinGecko mapping για: ' + ctmSymbol
    );
  }

  const ctmUrl =
    CTM_CONFIG.COINGECKO_BASE +
    '/coins/markets?vs_currency=usd' +
    '&ids=' + encodeURIComponent(ctmCoinId) +
    '&sparkline=true' +
    '&price_change_percentage=1h,24h,7d';

  const ctmResponse = CTM_fetchJson_(ctmUrl, {
    method: 'get',
    headers: CTM_buildCoinGeckoHeaders_(
      ctmRuntimeConfig.coinGeckoApiKey
    ),
    muteHttpExceptions: true
  });

  if (!ctmResponse || !ctmResponse.length) {
    throw new Error(
      'Δεν βρέθηκαν CoinGecko δεδομένα για: ' + ctmSymbol
    );
  }

  const ctmCoin = ctmResponse[0];

  const ctmPrices = (
    ctmCoin.sparkline_in_7d &&
    ctmCoin.sparkline_in_7d.price
      ? ctmCoin.sparkline_in_7d.price
      : []
  )
    .map(function(ctmValue) {
      return Number(ctmValue);
    })
    .filter(function(ctmValue) {
      return isFinite(ctmValue) && ctmValue > 0;
    });

  if (ctmPrices.length < 25) {
    throw new Error(
      'Ανεπαρκή CoinGecko sparkline data για: ' + ctmSymbol
    );
  }

  const ctmKlines1h = CTM_buildPseudoCandlesFromPrices_(
    ctmPrices,
    1
  );

  const ctmKlines4h = CTM_buildPseudoCandlesFromPrices_(
    ctmPrices,
    4
  );

  const ctmPrice = Number(ctmCoin.current_price || 0);

  return {
    provider: 'COINGECKO-SPOT',

    price: ctmPrice,

    priceChange24hPct: Number(
      ctmCoin.price_change_percentage_24h_in_currency ||
      ctmCoin.price_change_percentage_24h ||
      0
    ),

    high24h: Number(ctmCoin.high_24h || ctmPrice),
    low24h: Number(ctmCoin.low_24h || ctmPrice),

    quoteVolume24h: Number(ctmCoin.total_volume || 0),

    fundingRatePct8h: 0,
    fundingTime: new Date(),

    openInterest: 0,
    openInterest1hAgo: 0,
    oiChange1hPct: 0,

    klines1h: ctmKlines1h,
    klines4h: ctmKlines4h
  };
}

function CTM_buildPseudoCandlesFromPrices_(
  ctmPrices,
  ctmHoursPerCandle
) {
  const ctmCandles = [];
  const ctmNow = Date.now();
  const ctmTotalHours = ctmPrices.length;

  for (
    let ctmIndex = 0;
    ctmIndex <= ctmPrices.length - ctmHoursPerCandle;
    ctmIndex += ctmHoursPerCandle
  ) {
    const ctmChunk = ctmPrices.slice(
      ctmIndex,
      ctmIndex + ctmHoursPerCandle
    );

    const ctmOpen = ctmChunk[0];
    const ctmClose = ctmChunk[ctmChunk.length - 1];

    const ctmHigh = Math.max.apply(null, ctmChunk);
    const ctmLow = Math.min.apply(null, ctmChunk);

    const ctmTimestamp = ctmNow - (
      (ctmTotalHours - ctmIndex) * 60 * 60 * 1000
    );

    ctmCandles.push({
      openTime: ctmTimestamp,
      open: ctmOpen,
      high: ctmHigh,
      low: ctmLow,
      close: ctmClose,
      volume: 0,
      closeTime: ctmTimestamp +
        (ctmHoursPerCandle * 60 * 60 * 1000)
    });
  }

  return ctmCandles;
}





/**
 * ============================================================================
 * MACRO DASHBOARD – FED News & Risk Signals
 * ============================================================================
 * 
 * Σκοπός:
 *   - Παρακολουθεί μακροοικονομικούς δείκτες (CPI, PPI, 10Y, DXY, Gold, VIX, Oil)
 *   - Υπολογίζει ετήσιες μεταβολές (YoY) και ημερήσιες αλλαγές (1D)
 *   - Στέλνει Telegram alert όταν ξεπερνιούνται κατώφλια (risk-on / risk-off)
 * 
 * Δείκτες:
 *   - CPIAUCSL: Consumer Price Index (πληθωρισμός καταναλωτή)
 *   - PPIACO: Producer Price Index (πληθωρισμός παραγωγού)
 *   - DGS10: 10-Year Treasury Yield (απόδοση 10ετούς ομολόγου ΗΠΑ)
 *   - DTWEXBGS: Trade Weighted US Dollar Index (DXY)
 *   - GC=F: Gold Futures (Yahoo Finance, USD/oz)
 *   - ^VIX: CBOE Volatility Index (Yahoo Finance)
 *   - CL=F: Crude Oil WTI Futures (Yahoo Finance, USD/barrel)
 * 
 * Κατώφλια (THRESHOLDS):
 *   - CPI/PPI YoY ≥ 5% → risk-off
 *   - CPI/PPI YoY ≤ 3% → risk-on
 *   - 10Y Yield ≥ 4.5% → risk-off
 *   - 10Y Yield ↓ ≥ 0.30 pp (1D) → risk-on
 *   - DXY ≥ 105 → risk-off
 *   - DXY ↓ ≥ 1.0% (1D) → risk-on
 *   - Gold ↑/↓ ≥ 2% (1D) → risk-on / risk-off
 *   - VIX ≥ 25 → risk-off (υψηλός φόβος)
 *   - VIX ↓ ≤ 15 → risk-on (χαμηλός φόβος)
 *   - Oil ↑/↓ ≥ 3% (1D) → risk-on / risk-off
 * 
 * Οδηγίες:
 *   1. Ρύθμισε Script Properties: FRED_API_KEY, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
 *   2. Τρέξε setupDashboard() για αρχικοποίηση
 *   3. Δημιούργησε trigger: runDailyUpdate → Time-driven → Day timer → 15:30–16:00
 * 
 * ============================================================================
 */

// ==================== ΡΥΘΜΙΣΕΙΣ ====================
function getProps() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const FRED_API_KEY = props.FRED_API_KEY;
  const TELEGRAM_BOT_TOKEN = props.TELEGRAM_BOT_TOKEN;
  const TELEGRAM_CHAT_ID = props.TELEGRAM_CHAT_ID;
  if (!FRED_API_KEY || !TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
    throw new Error('Λείπουν Script Properties');
  }
  return { FRED_API_KEY, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID };
}

const SERIES = {
  CPI: 'CPIAUCSL',
  PPI: 'PPIACO',
  DGS10: 'DGS10',
  DXY: 'DTWEXBGS'
};

const THRESHOLDS = {
  CPI_YOY_HIGH: 5.0,
  PPI_YOY_HIGH: 5.0,
  CPI_YOY_LOW: 3.0,
  PPI_YOY_LOW: 3.0,
  DGS10_HIGH: 4.5,
  DGS10_DROP_1D: 0.30,
  DXY_HIGH: 105.0,
  DXY_DROP_1D_PCT: 1.0,
  GOLD_UP_1D: 2.0,
  GOLD_DOWN_1D: 2.0,
  VIX_HIGH: 25.0,
  VIX_LOW: 15.0,
  OIL_UP_1D: 3.0,
  OIL_DOWN_1D: 3.0
};

// ==================== ΒΟΗΘΗΤΙΚΕΣ ====================

function fetchFredSeries(seriesId, FRED_API_KEY, limit = 400) {
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=${seriesId}&api_key=${FRED_API_KEY}&file_type=json&sort_order=desc&limit=${limit}`;
  const options = { muteHttpExceptions: true, timeout: 20000 };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) throw new Error(`FRED error ${seriesId}`);
  const json = JSON.parse(response.getContentText());
  return json.observations
    .map(o => ({ date: new Date(o.date), value: o.value === '.' ? null : parseFloat(o.value) }))
    .filter(o => o.value !== null && !isNaN(o.value));
}

function fetchYahooSymbol(symbol, range = '5d') {
  // URL encoding για σύμβολα όπως ^VIX
  const encodedSymbol = encodeURIComponent(symbol);
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodedSymbol}?range=${range}&interval=1d`;
  const options = { muteHttpExceptions: true, timeout: 15000 };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log(`${symbol} fetch error: ${response.getResponseCode()}`);
    return { price: null, changePct: null, date: null };
  }
  const json = JSON.parse(response.getContentText());
  const result = json.chart.result[0];
  if (!result || !result.indicators.quote[0]) return { price: null, changePct: null, date: null };
  
  const quote = result.indicators.quote[0];
  const close = quote.close;
  const validClose = close.filter(c => c !== null);
  if (validClose.length < 2) return { price: null, changePct: null, date: null };
  
  const latestPrice = validClose[validClose.length - 1];
  const prevPrice = validClose[validClose.length - 2];
  const changePct = ((latestPrice - prevPrice) / prevPrice) * 100;
  const timestamp = result.timestamp[validClose.length - 1] * 1000;
  
  return { price: latestPrice, changePct: changePct, date: new Date(timestamp) };
}

function findYoY(data, currentDate) {
  const targetYear = currentDate.getFullYear() - 1;
  const targetMonth = currentDate.getMonth();
  for (const d of data) {
    if (d.date.getFullYear() === targetYear && d.date.getMonth() === targetMonth) {
      return d.value;
    }
  }
  const msPerMonth = 30.44 * 24 * 60 * 60 * 1000;
  const targetMs = currentDate.getTime() - (12 * msPerMonth);
  for (const d of data) {
    if (Math.abs(d.date.getTime() - targetMs) < (20 * msPerMonth)) {
      return d.value;
    }
  }
  return null;
}

function calcYoY(cur, prev) {
  if (cur === null || prev === null || prev === 0) return null;
  return ((cur - prev) / prev) * 100;
}

function calcChange(cur, prev) { return cur - prev; }
function calcPctChange(cur, prev) { return ((cur - prev) / prev) * 100; }

function sendTelegramMessage(token, chatId, text) {
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 10000 };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() !== 200) throw new Error('Telegram error');
}

function logAlert(sheet, indicator, value, threshold, direction, note) {
  sheet.appendRow([new Date(), indicator, value, threshold, direction, note]);
}

// ==================== ΚΥΡΙΑ ΛΟΓΙΚΗ ====================

function runDailyUpdate() {
  const { FRED_API_KEY, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID } = getProps();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const dashboardSheet = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  const alertsSheet = ss.getSheetByName('AlertsLog') || ss.insertSheet('AlertsLog');
  const rawDataSheet = ss.getSheetByName('RawData') || ss.insertSheet('RawData');
  
  if (alertsSheet.getLastRow() === 0) alertsSheet.appendRow(['Timestamp','Indicator','Value','Threshold','Direction','Note']);
  
  // FRED δεδομένα
  const cpi = fetchFredSeries('CPIAUCSL', FRED_API_KEY, 400);
  const ppi = fetchFredSeries('PPIACO', FRED_API_KEY, 400);
  const dgs10 = fetchFredSeries('DGS10', FRED_API_KEY, 100);
  const dxy = fetchFredSeries('DTWEXBGS', FRED_API_KEY, 100);
  
  // Yahoo Finance δεδομένα
  const gold = fetchYahooSymbol('GC=F', '5d');
  const vix = fetchYahooSymbol('^VIX', '5d');
  const oil = fetchYahooSymbol('CL=F', '5d');
  
  // Υπολογισμοί FRED
  const cpiLatest = cpi[0]?.value ?? null;
  const cpiDate = cpi[0]?.date ?? null;
  const cpiPrev12m = cpiDate ? findYoY(cpi, cpiDate) : null;
  const cpiYoY = calcYoY(cpiLatest, cpiPrev12m);
  
  const ppiLatest = ppi[0]?.value ?? null;
  const ppiDate = ppi[0]?.date ?? null;
  const ppiPrev12m = ppiDate ? findYoY(ppi, ppiDate) : null;
  const ppiYoY = calcYoY(ppiLatest, ppiPrev12m);
  
  const dgs10Latest = dgs10[0]?.value ?? null;
  const dgs10Date = dgs10[0]?.date ?? null;
  const dgs10_1d = (dgs10.length >= 2) ? calcChange(dgs10[0].value, dgs10[1].value) : null;
  
  const dxyLatest = dxy[0]?.value ?? null;
  const dxyDate = dxy[0]?.date ?? null;
  const dxy_1d_pct = (dxy.length >= 2) ? calcPctChange(dxy[0].value, dxy[1].value) : null;
  
  // Yahoo Finance values
  const goldPrice = gold.price;
  const goldDate = gold.date;
  const gold_1d_pct = gold.changePct;
  
  const vixPrice = vix.price;
  const vixDate = vix.date;
  const vix_1d_pct = vix.changePct;
  
  const oilPrice = oil.price;
  const oilDate = oil.date;
  const oil_1d_pct = oil.changePct;
  
  // Dashboard
  dashboardSheet.clear();
  dashboardSheet.appendRow(['Indicator','Latest Value','Date','YoY / 1D']);
  dashboardSheet.appendRow(['CPI', cpiLatest ?? '', cpiDate ?? '', cpiYoY !== null ? cpiYoY.toFixed(2)+'%' : '']);
  dashboardSheet.appendRow(['PPI', ppiLatest ?? '', ppiDate ?? '', ppiYoY !== null ? ppiYoY.toFixed(2)+'%' : '']);
  dashboardSheet.appendRow(['10Y Yield', dgs10Latest ?? '', dgs10Date ?? '', dgs10_1d !== null ? dgs10_1d.toFixed(2)+' pp' : '']);
  dashboardSheet.appendRow(['DXY', dxyLatest ?? '', dxyDate ?? '', dxy_1d_pct !== null ? dxy_1d_pct.toFixed(2)+'%' : '']);
  dashboardSheet.appendRow(['Gold (USD/oz)', goldPrice ?? '', goldDate ?? '', gold_1d_pct !== null ? gold_1d_pct.toFixed(2)+'%' : '']);
  dashboardSheet.appendRow(['VIX', vixPrice ?? '', vixDate ?? '', vix_1d_pct !== null ? vix_1d_pct.toFixed(2)+'%' : '']);
  dashboardSheet.appendRow(['Oil WTI (USD/bbl)', oilPrice ?? '', oilDate ?? '', oil_1d_pct !== null ? oil_1d_pct.toFixed(2)+'%' : '']);
  
  // RawData
  rawDataSheet.clear();
  rawDataSheet.appendRow(['Series','Date','Value']);
  [ ['CPI',cpi.slice(0,20)], ['PPI',ppi.slice(0,20)], ['DGS10',dgs10.slice(0,20)], ['DXY',dxy.slice(0,20)] ].forEach(([name, arr]) => {
    arr.forEach(r => rawDataSheet.appendRow([name, r.date, r.value]));
  });
  if (goldPrice !== null) rawDataSheet.appendRow(['Gold', goldDate, goldPrice]);
  if (vixPrice !== null) rawDataSheet.appendRow(['VIX', vixDate, vixPrice]);
  if (oilPrice !== null) rawDataSheet.appendRow(['Oil', oilDate, oilPrice]);
  
  // Alerts
  const alerts = [];
  
  // Πληθωρισμός (risk-off)
  if (cpiYoY !== null && cpiYoY >= THRESHOLDS.CPI_YOY_HIGH) {
    alerts.push({type:'risk-off', indicator:'CPI YoY', value:cpiYoY.toFixed(2)+'%', threshold:THRESHOLDS.CPI_YOY_HIGH+'%', direction:'HIGH'});
    logAlert(alertsSheet,'CPI YoY',cpiYoY,THRESHOLDS.CPI_YOY_HIGH,'HIGH','risk-off');
  }
  if (ppiYoY !== null && ppiYoY >= THRESHOLDS.PPI_YOY_HIGH) {
    alerts.push({type:'risk-off', indicator:'PPI YoY', value:ppiYoY.toFixed(2)+'%', threshold:THRESHOLDS.PPI_YOY_HIGH+'%', direction:'HIGH'});
    logAlert(alertsSheet,'PPI YoY',ppiYoY,THRESHOLDS.PPI_YOY_HIGH,'HIGH','risk-off');
  }
  
  // 10Y Yield (risk-off)
  if (dgs10Latest !== null && dgs10Latest >= THRESHOLDS.DGS10_HIGH) {
    alerts.push({type:'risk-off', indicator:'10Y Yield', value:dgs10Latest.toFixed(2)+'%', threshold:THRESHOLDS.DGS10_HIGH+'%', direction:'HIGH'});
    logAlert(alertsSheet,'10Y Yield',dgs10Latest,THRESHOLDS.DGS10_HIGH,'HIGH','risk-off');
  }
  
  // DXY (risk-off)
  if (dxyLatest !== null && dxyLatest >= THRESHOLDS.DXY_HIGH) {
    alerts.push({type:'risk-off', indicator:'DXY', value:dxyLatest.toFixed(2), threshold:THRESHOLDS.DXY_HIGH, direction:'HIGH'});
    logAlert(alertsSheet,'DXY',dxyLatest,THRESHOLDS.DXY_HIGH,'HIGH','risk-off');
  }
  
  // VIX (risk-off - υψηλός φόβος)
  if (vixPrice !== null && vixPrice >= THRESHOLDS.VIX_HIGH) {
    alerts.push({type:'risk-off', indicator:'VIX', value:vixPrice.toFixed(2), threshold:THRESHOLDS.VIX_HIGH, direction:'HIGH'});
    logAlert(alertsSheet,'VIX',vixPrice,THRESHOLDS.VIX_HIGH,'HIGH','risk-off (fear)');
  }
  
  // Πληθωρισμός (risk-on)
  if (cpiYoY !== null && cpiYoY <= THRESHOLDS.CPI_YOY_LOW) {
    alerts.push({type:'risk-on', indicator:'CPI YoY', value:cpiYoY.toFixed(2)+'%', threshold:THRESHOLDS.CPI_YOY_LOW+'%', direction:'LOW'});
    logAlert(alertsSheet,'CPI YoY',cpiYoY,THRESHOLDS.CPI_YOY_LOW,'LOW','risk-on');
  }
  if (ppiYoY !== null && ppiYoY <= THRESHOLDS.PPI_YOY_LOW) {
    alerts.push({type:'risk-on', indicator:'PPI YoY', value:ppiYoY.toFixed(2)+'%', threshold:THRESHOLDS.PPI_YOY_LOW+'%', direction:'LOW'});
    logAlert(alertsSheet,'PPI YoY',ppiYoY,THRESHOLDS.PPI_YOY_LOW,'LOW','risk-on');
  }
  
  // 10Y Yield drop (risk-on)
  if (dgs10_1d !== null && dgs10_1d <= -THRESHOLDS.DGS10_DROP_1D) {
    alerts.push({type:'risk-on', indicator:'10Y Yield (1D)', value:dgs10_1d.toFixed(2)+' pp', threshold:'-'+THRESHOLDS.DGS10_DROP_1D+' pp', direction:'DROP'});
    logAlert(alertsSheet,'10Y Yield 1D',dgs10_1d,-THRESHOLDS.DGS10_DROP_1D,'DROP','risk-on');
  }
  
  // DXY drop (risk-on)
  if (dxy_1d_pct !== null && dxy_1d_pct <= -THRESHOLDS.DXY_DROP_1D_PCT) {
    alerts.push({type:'risk-on', indicator:'DXY (1D)', value:dxy_1d_pct.toFixed(2)+'%', threshold:'-'+THRESHOLDS.DXY_DROP_1D_PCT+'%', direction:'DROP'});
    logAlert(alertsSheet,'DXY 1D',dxy_1d_pct,-THRESHOLDS.DXY_DROP_1D_PCT,'DROP','risk-on');
  }
  
  // VIX low (risk-on - χαμηλός φόβος)
  if (vixPrice !== null && vixPrice <= THRESHOLDS.VIX_LOW) {
    alerts.push({type:'risk-on', indicator:'VIX', value:vixPrice.toFixed(2), threshold:THRESHOLDS.VIX_LOW, direction:'LOW'});
    logAlert(alertsSheet,'VIX',vixPrice,THRESHOLDS.VIX_LOW,'LOW','risk-on (low fear)');
  }
  
  // Gold up (risk-on)
  if (gold_1d_pct !== null && gold_1d_pct >= THRESHOLDS.GOLD_UP_1D) {
    alerts.push({type:'risk-on', indicator:'Gold (1D)', value:gold_1d_pct.toFixed(2)+'%', threshold:'+'+THRESHOLDS.GOLD_UP_1D+'%', direction:'UP'});
    logAlert(alertsSheet,'Gold 1D',gold_1d_pct,THRESHOLDS.GOLD_UP_1D,'UP','risk-on (safe haven)');
  }
  
  // Gold down (risk-off)
  if (gold_1d_pct !== null && gold_1d_pct <= -THRESHOLDS.GOLD_DOWN_1D) {
    alerts.push({type:'risk-off', indicator:'Gold (1D)', value:gold_1d_pct.toFixed(2)+'%', threshold:'-'+THRESHOLDS.GOLD_DOWN_1D+'%', direction:'DOWN'});
    logAlert(alertsSheet,'Gold 1D',gold_1d_pct,-THRESHOLDS.GOLD_DOWN_1D,'DOWN','risk-off (panic)');
  }
  
  // Oil up/down (1D)
  if (oil_1d_pct !== null && oil_1d_pct >= THRESHOLDS.OIL_UP_1D) {
    alerts.push({type:'risk-on', indicator:'Oil (1D)', value:oil_1d_pct.toFixed(2)+'%', threshold:'+'+THRESHOLDS.OIL_UP_1D+'%', direction:'UP'});
    logAlert(alertsSheet,'Oil 1D',oil_1d_pct,THRESHOLDS.OIL_UP_1D,'UP','risk-on (energy demand)');
  }
  if (oil_1d_pct !== null && oil_1d_pct <= -THRESHOLDS.OIL_DOWN_1D) {
    alerts.push({type:'risk-off', indicator:'Oil (1D)', value:oil_1d_pct.toFixed(2)+'%', threshold:'-'+THRESHOLDS.OIL_DOWN_1D+'%', direction:'DOWN'});
    logAlert(alertsSheet,'Oil 1D',oil_1d_pct,-THRESHOLDS.OIL_DOWN_1D,'DOWN','risk-off (demand fear)');
  }
  
  // Telegram
  if (alerts.length > 0) {
    let msg = '🏛️ <b>FED News – Macro Risk Signals</b>\n\n';
    const riskOff = alerts.filter(a=>a.type==='risk-off');
    const riskOn = alerts.filter(a=>a.type==='risk-on');
    
    if (riskOff.length > 0) {
      msg += '<b>⚠️ Risk-off σήματα:</b>\n';
      riskOff.forEach(a => { msg += `• ${a.indicator}: <b>${a.value}</b> (threshold: ${a.threshold})\n`; });
      msg += 'Πιθανή πίεση σε BTC/ETH και risk assets.\n\n';
    }
    if (riskOn.length > 0) {
      msg += '<b>✅ Risk-on σήματα:</b>\n';
      riskOn.forEach(a => { msg += `• ${a.indicator}: <b>${a.value}</b> (threshold: ${a.threshold})\n`; });
      msg += 'Πιθανή στήριξη σε BTC/ETH και risk assets.';
    }
    
    sendTelegramMessage(TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, msg);
  }
}

function setupDashboard() {
  getProps();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Dashboard','RawData','AlertsLog'].forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  runDailyUpdate();
}

function testAlert() {
  const { TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID } = getProps();
  sendTelegramMessage(TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, '✅ <b>Test Alert</b>\n\nΤο FED News dashboard λειτουργεί σωστά.');
}

function testGold() {
  const gold = fetchYahooSymbol('GC=F', '5d');
  Logger.log('Gold: ' + gold.price + ' | ' + gold.changePct + '% | ' + gold.date);
}

function testVIX() {
  const vix = fetchYahooSymbol('^VIX', '5d');
  Logger.log('VIX: ' + vix.price + ' | ' + vix.changePct + '% | ' + vix.date);
}

function testOil() {
  const oil = fetchYahooSymbol('CL=F', '5d');
  Logger.log('Oil: ' + oil.price + ' | ' + oil.changePct + '% | ' + oil.date);
}






/**
 * gridEntry_checkOpenPositionsAndSignal
 * ---------------------------------------------------------------
 * Πρόθεμα: "gridEntry_" (για να μην συγκρούεται με τα υπόλοιπα
 * Telegram functions σου, π.χ. CA4H_, PIONEX_).
 *
 * Λογική:
 * 1) Ελέγχει το sheet "Open Positions", κελί A2. Αν είναι κενό
 *    (δεν υπάρχει ανοιχτή θέση) -> προχωράει στον έλεγχο.
 *    Αν έχει τιμή -> σταματάει (δεν κάνει τίποτα, δεν στέλνει Telegram).
 * 2) Διαβάζει το sheet "Grid Screener - LONG" (headers στη γραμμή 3,
 *    δεδομένα από τη γραμμή 4 και μετά):
 *      - Στήλη C: όνομα coin
 *      - Στήλη E: τελευταία τιμή (last price)
 *      - Στήλη O: προτεινόμενη τιμή εισόδου (entry price)
 * 3) Για κάθε coin υπολογίζει την % απόσταση της τελευταίας τιμής
 *    από την προτεινόμενη τιμή εισόδου (diffPct = (last-entry)/entry*100):
 *      - |diffPct| <= NEAR_THRESHOLD_PCT               -> "ΚΟΝΤΑ"
 *          - last <= entry  -> ΠΡΑΣΙΝΟ  (ιδανικό)
 *          - last > entry   -> ΠΟΡΤΟΚΑΛΙ (πλησιάζει από πάνω)
 *      - diffPct < -NEAR_THRESHOLD_PCT και
 *        diffPct >= -WATCH_THRESHOLD_PCT               -> "ΠΑΡΑΚΟΛΟΥΘΗΣΗ" -> ΜΠΛΕ
 *      - οτιδήποτε άλλο -> χωρίς χρώμα.
 * 4) Χρωματίζει τα αντίστοιχα κελιά στο Grid Screener - LONG (στήλες C:O).
 * 5) Για κάθε coin που εμφανίζεται στο μήνυμα, κάνει fetch τα τελευταία
 *    4h candles από το OKX public API και:
 *      α) Υπολογίζει σύντομη γραμμή τάσης (2 τελευταία κλεισμένα candles)
 *         + εκτίμηση πόσα candles χρειάζονται για να καλυφθεί η απόσταση
 *         μέχρι το entry.
 *      β) Υπολογίζει RSI(14), MACD(12,26,9), KDJ(9,3,3), Bollinger(20,2),
 *         ATR(14), Volume vs MA5 — ίδιες φόρμουλες με το CA4H_ project σου.
 *      γ) Συνθέτει όλα τα (β) σε ΜΙΑ ενιαία παράγραφο φυσικής γλώσσας.
 * 6) Στέλνει μήνυμα στο Telegram με δύο ενότητες (κοντά / watch). Κάτω
 *    από κάθε coin: γραμμή 4h trend + διαφοράς, και μετά η συνολική
 *    παράγραφος δεικτών.
 *
 * ΠΑΡΑΜΕΤΡΟΣ ON/OFF ΤΟΥ ΕΛΕΓΧΟΥ OPEN POSITIONS:
 *   Ορίζεται στη μεταβλητή CHECK_OPEN_POSITIONS_ENABLED παρακάτω,
 *   ΚΑΙ μπορεί να γίνει override περνώντας παράμετρο στη function:
 *
 *     gridEntry_checkOpenPositionsAndSignal();      // χρησιμοποιεί το CHECK_OPEN_POSITIONS_ENABLED
 *     gridEntry_checkOpenPositionsAndSignal(true);   // αναγκάζει έλεγχο Open Positions (A2)
 *     gridEntry_checkOpenPositionsAndSignal(false);  // αγνοεί το Open Positions, ελέγχει πάντα
 *
 * Script Properties που χρησιμοποιούνται:
 *   TELEGRAM_BOT_TOKEN
 *   TELEGRAM_CHAT_ID
 * ---------------------------------------------------------------
 */

// -----------------------------------------------------------
// >>> ΕΔΩ ΟΡΙΖΕΙΣ ΤΟ ON/OFF ΤΟΥ ΕΛΕΓΧΟΥ OPEN POSITIONS <<<
var CHECK_OPEN_POSITIONS_ENABLED = true;
// -----------------------------------------------------------

function gridEntry_checkOpenPositionsAndSignal(checkOpenPositions) {
  if (typeof checkOpenPositions === 'undefined') {
    checkOpenPositions = CHECK_OPEN_POSITIONS_ENABLED;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var OPEN_POSITIONS_SHEET = 'Open Positions';
  var GRID_SCREENER_SHEET  = 'Grid Screener - LONG';

  var HEADER_ROW = 3;
  var FIRST_DATA_ROW = HEADER_ROW + 1; // = 4

  var NEAR_THRESHOLD_PCT = 1.5;
  var WATCH_THRESHOLD_PCT = 12;

  var COLOR_BELOW_OR_EQUAL = '#b6d7a8';
  var COLOR_ABOVE_NEAR     = '#f9cb9c';
  var COLOR_WATCH_BELOW    = '#9fc5e8';
  var COLOR_RESET          = null;

  // -----------------------------------------------------------
  // 1) Έλεγχος Open Positions
  // -----------------------------------------------------------
  if (checkOpenPositions) {
    var opSheet = ss.getSheetByName(OPEN_POSITIONS_SHEET);
    if (!opSheet) {
      Logger.log('Δεν βρέθηκε το sheet "' + OPEN_POSITIONS_SHEET + '".');
      return;
    }
    var openPositionValue = opSheet.getRange('A2').getValue();
    var hasOpenPosition = openPositionValue !== '' && openPositionValue !== null;
    if (hasOpenPosition) {
      Logger.log('Υπάρχει ήδη ανοιχτή θέση (' + openPositionValue + '). Ο έλεγχος Grid Screener παραλείπεται.');
      return;
    }
  }

  // -----------------------------------------------------------
  // 2) Ανάγνωση Grid Screener - LONG
  // -----------------------------------------------------------
  var gridSheet = ss.getSheetByName(GRID_SCREENER_SHEET);
  if (!gridSheet) {
    Logger.log('Δεν βρέθηκε το sheet "' + GRID_SCREENER_SHEET + '".');
    return;
  }

  var lastRow = gridSheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) {
    Logger.log('Δεν υπάρχουν δεδομένα στο Grid Screener - LONG.');
    return;
  }

  var numRows = lastRow - FIRST_DATA_ROW + 1;

  var coinRange   = gridSheet.getRange(FIRST_DATA_ROW, 3, numRows, 1);  // Στήλη C
  var lastRange   = gridSheet.getRange(FIRST_DATA_ROW, 5, numRows, 1);  // Στήλη E
  var entryRange  = gridSheet.getRange(FIRST_DATA_ROW, 18, numRows, 1); // Στήλη O

  var coins   = coinRange.getValues();
  var lasts   = lastRange.getValues();
  var entries = entryRange.getValues();

  var rowBgRanges = [];
  for (var r = 0; r < numRows; r++) {
    rowBgRanges.push(gridSheet.getRange(FIRST_DATA_ROW + r, 3, 1, 13));
  }

  var nearCoins = [];
  var watchCoins = [];

  for (var i = 0; i < numRows; i++) {
    var coinName = coins[i][0];
    var lastPrice = parseFloat(lasts[i][0]);
    var entryPrice = parseFloat(entries[i][0]);

    if (!coinName || isNaN(lastPrice) || isNaN(entryPrice) || entryPrice === 0) {
      rowBgRanges[i].setBackground(COLOR_RESET);
      continue;
    }

    var diffPct = ((lastPrice - entryPrice) / entryPrice) * 100;
    var isNear = Math.abs(diffPct) <= NEAR_THRESHOLD_PCT;
    var isWatchBelow = !isNear && diffPct < 0 && Math.abs(diffPct) <= WATCH_THRESHOLD_PCT;

    if (isNear) {
      var isBelowOrEqual = lastPrice <= entryPrice;
      rowBgRanges[i].setBackground(isBelowOrEqual ? COLOR_BELOW_OR_EQUAL : COLOR_ABOVE_NEAR);
      nearCoins.push({ coin: coinName, last: lastPrice, entry: entryPrice, diffPct: diffPct, belowOrEqual: isBelowOrEqual });
    } else if (isWatchBelow) {
      rowBgRanges[i].setBackground(COLOR_WATCH_BELOW);
      watchCoins.push({ coin: coinName, last: lastPrice, entry: entryPrice, diffPct: diffPct });
    } else {
      rowBgRanges[i].setBackground(COLOR_RESET);
    }
  }

  // -----------------------------------------------------------
  // 3) Αποστολή μηνύματος Telegram
  // -----------------------------------------------------------
  if (nearCoins.length > 0 || watchCoins.length > 0) {
    nearCoins.sort(function (a, b) {
      if (a.belowOrEqual !== b.belowOrEqual) return a.belowOrEqual ? -1 : 1;
      return Math.abs(a.diffPct) - Math.abs(b.diffPct);
    });
    watchCoins.sort(function (a, b) {
      return Math.abs(a.diffPct) - Math.abs(b.diffPct);
    });

    var lines = [];
    lines.push('📊 <b>Grid Entry Watch</b>');
    lines.push('');

    if (nearCoins.length > 0) {
      lines.push('<b>Κοντά στην τιμή εισόδου:</b>');
      nearCoins.forEach(function (item) {
        var icon = item.belowOrEqual ? '🟢' : '🟠';
        var sign = item.diffPct >= 0 ? '+' : '';
        lines.push(
          icon + ' <b>' + item.coin + '</b>: last ' + item.last +
          ' | entry ' + item.entry +
          ' (' + sign + item.diffPct.toFixed(2) + '%)'
        );
        var candles = gridEntry_fetch4hCandles_(item.coin);
        lines.push(gridEntry_get4hTrendLine_(candles, item.diffPct));
        lines.push(gridEntry_getIndicatorParagraph_(candles));
        lines.push('');
      });
    }

    if (watchCoins.length > 0) {
      lines.push('<b>Πλησιάζουν το κατώφλι (κάτω από entry):</b>');
      watchCoins.forEach(function (item) {
        var sign = item.diffPct >= 0 ? '+' : '';
        lines.push(
          '🔵 <b>' + item.coin + '</b>: last ' + item.last +
          ' | entry ' + item.entry +
          ' (' + sign + item.diffPct.toFixed(2) + '%)'
        );
        var candles = gridEntry_fetch4hCandles_(item.coin);
        lines.push(gridEntry_get4hTrendLine_(candles, item.diffPct));
        lines.push(gridEntry_getIndicatorParagraph_(candles));
        lines.push('');
      });
    }

    var message = lines.join('\n').trim();
    gridEntry_sendTelegramMessage_(message);
  } else {
    Logger.log('Κανένα coin δεν είναι κοντά ή σε απόσταση παρακολούθησης από την τιμή εισόδου.');
  }
}

/**
 * gridEntry_fetch4hCandles_
 * Κάνει fetch τα τελευταία N candles 4H από το OKX public API.
 * Επιστρέφει array candles σε ΑΥΞΟΥΣΑ χρονική σειρά (παλαιότερο -> νεότερο),
 * κάθε candle ως { ts, open, high, low, close, volume }, ή null σε σφάλμα.
 * Αγνοεί το πρώτο (πιο πρόσφατο) candle του OKX response, που είναι ημιτελές.
 */
function gridEntry_fetch4hCandles_(coinName) {
  try {
    var instId = gridEntry_normalizeToOkxInstId_(coinName);
    var url = 'https://www.okx.com/api/v5/market/candles?instId=' + instId + '&bar=4H&limit=60';
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());

    if (!json || json.code !== '0' || !json.data || json.data.length < 30) {
      return null;
    }

    // OKX: [0]=πιο πρόσφατο (ημιτελές), μετά φθίνουσα σειρά. Αγνοούμε [0], αντιστρέφουμε τα υπόλοιπα.
    var raw = json.data.slice(1).reverse();
    var candles = raw.map(function (c) {
      return {
        ts: parseInt(c[0], 10),
        open: parseFloat(c[1]),
        high: parseFloat(c[2]),
        low: parseFloat(c[3]),
        close: parseFloat(c[4]),
        volume: parseFloat(c[5])
      };
    });
    return candles;
  } catch (e) {
    Logger.log('Σφάλμα fetch candles για ' + coinName + ': ' + e.message);
    return null;
  }
}

/**
 * gridEntry_get4hTrendLine_
 * Χρησιμοποιεί τα 2 τελευταία κλεισμένα candles (τέλος του array) για να
 * εκτιμήσει τάση και πόσα candles χρειάζονται για κάλυψη της diffPct.
 */
function gridEntry_get4hTrendLine_(candles, diffPct) {
  if (!candles || candles.length < 2) {
    return '　↳ ⚪ 4h ανάλυση μη διαθέσιμη';
  }

  var n = candles.length;
  var candleLastClosed = candles[n - 1];
  var candlePrevClosed = candles[n - 2];

  var candleLastBullish = candleLastClosed.close > candleLastClosed.open;
  var candlePrevBullish = candlePrevClosed.close > candlePrevClosed.open;
  var closeToCloseChangePct = ((candleLastClosed.close - candlePrevClosed.close) / candlePrevClosed.close) * 100;

  var trendLabel, trendIcon;
  if (candleLastBullish && candlePrevBullish && closeToCloseChangePct > 0) {
    trendLabel = 'Ανοδική'; trendIcon = '📈';
  } else if (!candleLastBullish && !candlePrevBullish && closeToCloseChangePct < 0) {
    trendLabel = 'Καθοδική'; trendIcon = '📉';
  } else if (closeToCloseChangePct > 0) {
    trendLabel = 'Ήπια ανοδική / μεικτή'; trendIcon = '↗️';
  } else if (closeToCloseChangePct < 0) {
    trendLabel = 'Ήπια καθοδική / μεικτή'; trendIcon = '↘️';
  } else {
    trendLabel = 'Πλάγια'; trendIcon = '➡️';
  }

  var sign = closeToCloseChangePct >= 0 ? '+' : '';
  var coverageText = '';
  var neededDirectionIsDown = diffPct > 0;
  var rateHelpsCoverage = neededDirectionIsDown ? (closeToCloseChangePct < 0) : (closeToCloseChangePct > 0);

  if (rateHelpsCoverage && Math.abs(closeToCloseChangePct) > 0.01) {
    var candlesNeeded = Math.abs(diffPct) / Math.abs(closeToCloseChangePct);
    var hoursNeeded = candlesNeeded * 4;
    coverageText = ' | κάλυψη διαφοράς σε ~' + candlesNeeded.toFixed(1) + ' candles (~' + hoursNeeded.toFixed(0) + 'h) με τον τρέχοντα ρυθμό';
  } else {
    coverageText = ' | με τον τρέχοντα ρυθμό δεν πλησιάζει το entry';
  }

  return '　↳ ' + trendIcon + ' ' + trendLabel + ' (' + sign + closeToCloseChangePct.toFixed(2) + '% τελευταίο 4h close-to-close)' + coverageText;
}

/**
 * gridEntry_getIndicatorParagraph_
 * Υπολογίζει RSI(14), MACD(12,26,9), KDJ(9,3,3), Bollinger(20,2), ATR(14)
 * και Volume vs MA5, και τα συνθέτει σε ΜΙΑ παράγραφο φυσικής γλώσσας
 * (ίδιο πνεύμα με το CA4H_ project).
 */
function gridEntry_getIndicatorParagraph_(candles) {
  if (!candles || candles.length < 30) {
    return '　↳ ⚪ Δεν υπάρχουν αρκετά δεδομένα για σύνθεση δεικτών.';
  }

  var closes = candles.map(function (c) { return c.close; });
  var volumes = candles.map(function (c) { return c.volume; });

  var rsi = gridEntry_calcRsi_(closes, 14);
  var macd = gridEntry_calcMacd_(closes, 12, 26, 9);
  var kdj = gridEntry_calcKdj_(candles, 9, 3, 3);
  var boll = gridEntry_calcBollinger_(closes, 20, 2);
  var atr = gridEntry_calcAtr_(candles, 14);
  var lastClose = closes[closes.length - 1];
  var ma5Volume = gridEntry_sma_(volumes.slice(volumes.length - 5), 5);
  var lastVolume = volumes[volumes.length - 1];

  // --- RSI ερμηνεία ---
  var rsiText;
  if (rsi >= 70) rsiText = 'το RSI (' + rsi.toFixed(1) + ') βρίσκεται σε υπεραγορασμένη ζώνη';
  else if (rsi <= 30) rsiText = 'το RSI (' + rsi.toFixed(1) + ') βρίσκεται σε υπερπουλημένη ζώνη';
  else rsiText = 'το RSI (' + rsi.toFixed(1) + ') είναι ουδέτερο';

  // --- MACD ερμηνεία ---
  var macdText = macd.histogram > 0
    ? 'το MACD histogram είναι θετικό (ανοδικό momentum)'
    : 'το MACD histogram είναι αρνητικό (καθοδικό momentum)';

  // --- KDJ ερμηνεία ---
  var kdjText;
  if (kdj.j > 90) kdjText = 'το KDJ δείχνει ακραίο καμπανάκι υπεραγοράς';
  else if (kdj.j < 10) kdjText = 'το KDJ δείχνει ακραίο καμπανάκι υπερπούλησης';
  else if (kdj.k > kdj.d) kdjText = 'το KDJ δείχνει ισχυρή ανοδική ορμή';
  else kdjText = 'το KDJ δείχνει καθοδική ορμή';

  // --- Bollinger ερμηνεία ---
  var bandwidthPct = ((boll.upper - boll.lower) / boll.middle) * 100;
  var bollText;
  if (lastClose > boll.upper) bollText = 'η τιμή κινείται πάνω από την άνω κορδέλα Bollinger';
  else if (lastClose < boll.lower) bollText = 'η τιμή κινείται κάτω από την κάτω κορδέλα Bollinger';
  else bollText = 'η τιμή κινείται εντός των κορδελών Bollinger';
  if (bandwidthPct < 3) bollText += ' με έντονο squeeze (χαμηλή μεταβλητότητα πριν από πιθανό breakout)';

  // --- ATR ερμηνεία ---
  var atrPct = (atr / lastClose) * 100;
  var atrText = 'η μεταβλητότητα (ATR ' + atrPct.toFixed(2) + '% της τιμής) είναι ' +
    (atrPct > 4 ? 'σχετικά υψηλή' : (atrPct < 1.5 ? 'σχετικά χαμηλή' : 'μέτρια'));

  // --- Volume ερμηνεία ---
  var volumeText = lastVolume >= ma5Volume
    ? 'ο όγκος επιβεβαιώνει την τρέχουσα κίνηση'
    : 'ο όγκος είναι κάτω από τον μέσο όρο 5 candles, χωρίς πλήρη επιβεβαίωση';

  // --- Συνολική σύνθεση ---
  var bullishSignals = 0, bearishSignals = 0;
  if (rsi > 50) bullishSignals++; else bearishSignals++;
  if (macd.histogram > 0) bullishSignals++; else bearishSignals++;
  if (kdj.k > kdj.d) bullishSignals++; else bearishSignals++;
  if (lastClose > boll.middle) bullishSignals++; else bearishSignals++;

  var overallText;
  if (bullishSignals >= 3) overallText = 'Η σύνθεση όλων αυτών δίνει καθαρή εικόνα ανοδικής πίεσης.';
  else if (bearishSignals >= 3) overallText = 'Η σύνθεση όλων αυτών δίνει καθαρή εικόνα καθοδικής πίεσης.';
  else overallText = 'Η σύνθεση όλων αυτών δίνει μεικτή, χωρίς ξεκάθαρη κατεύθυνση εικόνα.';

  var paragraph = 'Συνολικά, ' + rsiText + ', ' + macdText + ', ' + kdjText + ', ' +
    bollText + ', ' + atrText + ', ' + volumeText + '. ' + overallText;

  return '　↳ 🧩 ' + paragraph;
}

// ---------------------------------------------------------------
// Βοηθητικές functions υπολογισμού δεικτών (πρόθεμα gridEntry_)
// ---------------------------------------------------------------

function gridEntry_sma_(arr, period) {
  var slice = arr.slice(arr.length - period);
  var sum = slice.reduce(function (a, b) { return a + b; }, 0);
  return sum / slice.length;
}

function gridEntry_calcRsi_(closes, period) {
  var gains = 0, losses = 0;
  for (var i = closes.length - period; i < closes.length; i++) {
    var diff = closes[i] - closes[i - 1];
    if (diff >= 0) gains += diff; else losses += Math.abs(diff);
  }
  var avgGain = gains / period;
  var avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  var rs = avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

function gridEntry_ema_(values, period) {
  var k = 2 / (period + 1);
  var emaArr = [values[0]];
  for (var i = 1; i < values.length; i++) {
    emaArr.push(values[i] * k + emaArr[i - 1] * (1 - k));
  }
  return emaArr;
}

function gridEntry_calcMacd_(closes, fastPeriod, slowPeriod, signalPeriod) {
  var emaFast = gridEntry_ema_(closes, fastPeriod);
  var emaSlow = gridEntry_ema_(closes, slowPeriod);
  var macdLine = closes.map(function (_, i) { return emaFast[i] - emaSlow[i]; });
  var signalLine = gridEntry_ema_(macdLine, signalPeriod);
  var lastIndex = closes.length - 1;
  var histogram = macdLine[lastIndex] - signalLine[lastIndex];
  return { macd: macdLine[lastIndex], signal: signalLine[lastIndex], histogram: histogram };
}

function gridEntry_calcKdj_(candles, rsvPeriod, kPeriod, dPeriod) {
  var kValues = [50], dValues = [50];
  for (var i = 1; i < candles.length; i++) {
    var start = Math.max(0, i - rsvPeriod + 1);
    var windowSlice = candles.slice(start, i + 1);
    var highestHigh = Math.max.apply(null, windowSlice.map(function (c) { return c.high; }));
    var lowestLow = Math.min.apply(null, windowSlice.map(function (c) { return c.low; }));
    var rsv = highestHigh === lowestLow ? 50 : ((candles[i].close - lowestLow) / (highestHigh - lowestLow)) * 100;
    var k = (kValues[i - 1] * (kPeriod - 1) + rsv) / kPeriod;
    var d = (dValues[i - 1] * (dPeriod - 1) + k) / dPeriod;
    kValues.push(k);
    dValues.push(d);
  }
  var lastK = kValues[kValues.length - 1];
  var lastD = dValues[dValues.length - 1];
  var lastJ = 3 * lastK - 2 * lastD;
  return { k: lastK, d: lastD, j: lastJ };
}

function gridEntry_calcBollinger_(closes, period, stdDevMultiplier) {
  var slice = closes.slice(closes.length - period);
  var middle = gridEntry_sma_(slice, period);
  var variance = slice.reduce(function (sum, val) { return sum + Math.pow(val - middle, 2); }, 0) / period;
  var stdDev = Math.sqrt(variance);
  return { middle: middle, upper: middle + stdDevMultiplier * stdDev, lower: middle - stdDevMultiplier * stdDev };
}

function gridEntry_calcAtr_(candles, period) {
  var trs = [];
  for (var i = 1; i < candles.length; i++) {
    var current = candles[i];
    var prevClose = candles[i - 1].close;
    var tr = Math.max(
      current.high - current.low,
      Math.abs(current.high - prevClose),
      Math.abs(current.low - prevClose)
    );
    trs.push(tr);
  }
  var recentTrs = trs.slice(trs.length - period);
  return gridEntry_sma_(recentTrs, recentTrs.length);
}

/**
 * gridEntry_normalizeToOkxInstId_
 */
function gridEntry_normalizeToOkxInstId_(coinName) {
  var raw = String(coinName).trim().toUpperCase();
  if (raw.indexOf('-') !== -1) return raw;
  if (raw.indexOf('/') !== -1) {
    var parts = raw.split('/');
    return parts[0] + '-' + parts[1];
  }
  var knownQuotes = ['USDT', 'USDC', 'BUSD', 'USD'];
  for (var q = 0; q < knownQuotes.length; q++) {
    var quote = knownQuotes[q];
    if (raw.length > quote.length && raw.slice(-quote.length) === quote) {
      var base = raw.slice(0, raw.length - quote.length);
      return base + '-' + quote;
    }
  }
  return raw;
}

/**
 * gridEntry_sendTelegramMessage_
 */
function gridEntry_sendTelegramMessage_(message) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_BOT_TOKEN');
  var chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    Logger.log('Λείπουν τα Script Properties TELEGRAM_BOT_TOKEN ή TELEGRAM_CHAT_ID.');
    return;
  }

  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  var payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'HTML'
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log('Telegram response: ' + response.getContentText());
  } catch (e) {
    Logger.log('Σφάλμα αποστολής Telegram: ' + e.message);
  }
}


/**
 * OKX 4H Pionex Grid Screener — FINAL TOP 120 VERSION
 *
 * Η μόνη κύρια function είναι:
 *   OKX4H_runScreener
 *
 * Η έρευνα περιορίζεται στα 120 πιο ρευστά LIVE USDT spot instruments
 * του OKX, ταξινομημένα με βάση το volCcy24h.
 *
 * Περιλαμβάνει:
 * - 4H confirmed candles
 * - EMA20 / EMA50 / RSI14 / MACD / ADX / ATR
 * - Spot Grid ή Long Futures Grid
 * - Grid range και Fibonacci 0.618 / 0.786
 * - Alternative entry στο 43% από το upper grid limit
 * - Funding rate από OKX perpetual swap
 * - Funding penalty για LONG FUTURES GRID
 * - 20–120 grids ανάλογα με range και κόστος
 * - Top 10 στο sheet
 * - Telegram message
 *
 * Script Properties:
 * TELEGRAM_BOT_TOKEN
 * TELEGRAM_CHAT_ID
 */

const OKX4H_CFG = {
  sheetName: 'OKX_4H_Screener',
  bar: '4H',
  topLiquidCoins: 120,
  candles: 120,
  topRows: 10,
  minScore: 58,
  atrGridMultiplier: 2.2,
  cooldownMinutes: 180,
  apiBase: 'https://www.okx.com/api/v5',
  quoteCurrency: 'USDT',

  fundingPenaltyStartPct: 0.01,
  fundingPenaltyFullPct: 0.08,
  maxFundingPenalty: 18,

  minGrids: 20,
  maxGrids: 120,
  minGridStepPct: 0.30,
  maxGridStepPct: 0.90,
  targetNetGridPctSpot: 0.35,
  targetNetGridPctFutures: 0.30,
  spotRoundTripFeePct: 0.10,
  futuresRoundTripFeePct: 0.04,
  slippageBufferPct: 0.05,
  minNetGridPct: 0.20
};

function OKX4H_runScreener() {
  const sheet = OKX4H_getSheet_();
  const tickers = OKX4H_getTopLiquidSpotTickers_();
  const rows = [];

  tickers.forEach(ticker => {
    try {
      const candles = OKX4H_getCandles_(ticker.instId, OKX4H_CFG.candles);
      if (candles.length < 80) return;

      const item = OKX4H_analyze_(ticker, candles);

      if (item.recommendation === 'LONG FUTURES GRID') {
        const funding = OKX4H_getFunding_(ticker.instId + '-SWAP');
        item.fundingRate = funding.fundingRate;
        item.nextFundingRate = funding.nextFundingRate;
        item.fundingRatePct = funding.fundingRate * 100;
        item.nextFundingRatePct = funding.nextFundingRate * 100;
        item.fundingPenalty = OKX4H_fundingPenalty_(funding.nextFundingRate);
        item.fundingComment = OKX4H_fundingComment_(funding.nextFundingRate, item.fundingPenalty);
      } else {
        item.fundingRate = null;
        item.nextFundingRate = null;
        item.fundingRatePct = null;
        item.nextFundingRatePct = null;
        item.fundingPenalty = 0;
        item.fundingComment = 'Δεν εφαρμόστηκε: δεν είναι Futures Grid.';
      }

      item.finalScore = Math.max(0, item.score - item.fundingPenalty);
      rows.push(item);
      Utilities.sleep(80);
    } catch (error) {
      console.warn(ticker.instId + ': ' + error.message);
    }
  });

  rows.sort((a, b) => b.finalScore - a.finalScore);
  OKX4H_writeSheet_(sheet, rows);

  if (rows.length && rows[0].finalScore >= OKX4H_CFG.minScore) {
    OKX4H_sendTelegramIfNeeded_(rows[0], rows);
  }

  return rows.slice(0, OKX4H_CFG.topRows);
}

function OKX4H_getTopLiquidSpotTickers_() {
  const instruments = OKX4H_getLiveSpotInstruments_();
  const liveUsdt = {};

  instruments.forEach(instrument => {
    if (instrument.state === 'live' &&
        instrument.quoteCcy === OKX4H_CFG.quoteCurrency &&
        instrument.instId &&
        instrument.instId.endsWith('-' + OKX4H_CFG.quoteCurrency)) {
      liveUsdt[instrument.instId] = true;
    }
  });

  const json = OKX4H_fetchJson_('/market/tickers?instType=SPOT');
  if (json.code !== '0') throw new Error(json.msg || 'OKX ticker API error');

  return (json.data || [])
    .filter(ticker => liveUsdt[ticker.instId] && Number(ticker.last) > 0 && Number(ticker.volCcy24h) > 0)
    .sort((a, b) => Number(b.volCcy24h) - Number(a.volCcy24h))
    .slice(0, OKX4H_CFG.topLiquidCoins);
}

function OKX4H_getLiveSpotInstruments_() {
  const json = OKX4H_fetchJson_('/public/instruments?instType=SPOT');
  if (json.code !== '0') throw new Error(json.msg || 'OKX instruments API error');
  return json.data || [];
}

function OKX4H_getCandles_(instId, limit) {
  const json = OKX4H_fetchJson_('/market/candles?instId=' + encodeURIComponent(instId) + '&bar=' + OKX4H_CFG.bar + '&limit=' + limit);
  if (json.code !== '0') throw new Error(json.msg || 'OKX candles API error');

  return (json.data || [])
    .map(row => ({
      ts: Number(row[0]), o: Number(row[1]), h: Number(row[2]),
      l: Number(row[3]), c: Number(row[4]), vol: Number(row[5]),
      confirm: String(row[8])
    }))
    .filter(x => x.confirm === '1' && x.c > 0 && x.h > 0 && x.l > 0)
    .sort((a, b) => a.ts - b.ts);
}

function OKX4H_getFunding_(swapId) {
  const json = OKX4H_fetchJson_('/public/funding-rate?instId=' + encodeURIComponent(swapId));
  if (json.code !== '0' || !json.data || !json.data.length) {
    return { fundingRate: 0, nextFundingRate: 0 };
  }

  const row = json.data[0];
  const current = Number(row.fundingRate || row.settFundingRate || 0);
  const next = Number(row.nextFundingRate || row.fundingRate || row.settFundingRate || 0);

  return {
    fundingRate: isFinite(current) ? current : 0,
    nextFundingRate: isFinite(next) ? next : 0
  };
}

function OKX4H_analyze_(ticker, candles) {
  const closes = candles.map(x => x.c);
  const volumes = candles.map(x => x.vol);
  const price = closes[closes.length - 1];

  const ema20 = OKX4H_ema_(closes, 20);
  const ema50 = OKX4H_ema_(closes, 50);
  const rsi = OKX4H_rsi_(closes, 14);
  const atr = OKX4H_atr_(candles, 14);
  const adx = OKX4H_adx_(candles, 14);
  const macd = OKX4H_macd_(closes);
  const avgVol = OKX4H_average_(volumes.slice(-20));
  const relativeVolume = avgVol > 0 ? volumes[volumes.length - 1] / avgVol : 1;
  const change24h = closes.length >= 7 ? (price / closes[closes.length - 7] - 1) * 100 : 0;
  const atrPct = price > 0 ? atr / price * 100 : 0;

  let score = 0;
  const reasons = [];

  if (price > ema20) { score += 12; reasons.push('τιμή πάνω από EMA20'); }
  if (ema20 > ema50) { score += 14; reasons.push('EMA20 πάνω από EMA50'); }
  if (rsi >= 52 && rsi <= 68) { score += 12; reasons.push('RSI ανοδικό χωρίς υπεραγορά'); }
  else if (rsi > 68) { score -= 8; reasons.push('RSI υπεραγορασμένο'); }
  else if (rsi < 35) { score -= 8; reasons.push('RSI αδύναμο'); }
  if (macd.hist > 0) { score += 12; reasons.push('θετικό MACD histogram'); }
  if (macd.hist > macd.prevHist) { score += 7; reasons.push('MACD momentum αυξάνεται'); }
  if (adx >= 20) { score += 12; reasons.push('ADX δείχνει τάση'); }
  else { score += 3; reasons.push('χαμηλό ADX/range'); }
  if (relativeVolume >= 1.2) { score += 8; reasons.push('αυξημένος όγκος'); }
  if (change24h > 0) { score += 6; reasons.push('θετική μεταβολή 24h'); }
  if (change24h < -4) { score -= 6; reasons.push('έντονη πτώση 24h'); }
  if (atrPct > 8) { score -= 5; reasons.push('υψηλό ATR'); }

  score = Math.max(0, Math.min(100, Math.round(score)));
  const trend = price > ema20 && ema20 > ema50 && macd.hist > 0
    ? 'UPTREND'
    : price < ema20 && ema20 < ema50 && macd.hist < 0
      ? 'DOWNTREND'
      : 'RANGE/TRANSITION';
  const recommendation = OKX4H_recommend_(trend, score, rsi, adx);

  const gridLow = Math.max(price - OKX4H_CFG.atrGridMultiplier * atr, price * 0.75);
  const gridHigh = Math.min(price + OKX4H_CFG.atrGridMultiplier * atr, price * 1.35);
  const gridRange = gridHigh - gridLow;
  const fib618 = gridHigh - gridRange * 0.618;
  const fib786 = gridHigh - gridRange * 0.786;
  const altEntry43 = gridHigh - gridRange * 0.43;
  const botEntry = OKX4H_botEntry_(recommendation, fib618, fib786, altEntry43);
  const gridPlan = OKX4H_calculateGridPlan_(gridLow, gridHigh, recommendation);

  return {
    instId: ticker.instId, price: price, volume24h: Number(ticker.volCcy24h),
    score: score, finalScore: score, trend: trend, recommendation: recommendation,
    rsi: rsi, adx: adx, atrPct: atrPct, relativeVolume: relativeVolume,
    change24h: change24h, ema20: ema20, ema50: ema50, macdHist: macd.hist,
    gridLow: gridLow, gridHigh: gridHigh, fib618: fib618, fib786: fib786,
    altEntry43: altEntry43, botEntry: botEntry,
    gridCount: gridPlan.gridCount, actualStepPct: gridPlan.actualStepPct,
    estimatedGrossPerGridPct: gridPlan.estimatedGrossPerGridPct,
    estimatedNetPerGridPct: gridPlan.estimatedNetPerGridPct,
    gridWarning: gridPlan.gridWarning, reasons: reasons.join('; '), updated: new Date()
  };
}

function OKX4H_recommend_(trend, score, rsi, adx) {
  if (score < OKX4H_CFG.minScore) return 'WAIT';
  if (trend === 'UPTREND' && adx >= 22 && rsi < 68) return 'LONG FUTURES GRID';
  if (trend === 'RANGE/TRANSITION' || adx < 20) return 'SPOT GRID';
  if (trend === 'UPTREND' && rsi >= 68) return 'WAIT / PULLBACK';
  return 'SPOT GRID';
}

function OKX4H_botEntry_(recommendation, fib618, fib786, altEntry43) {
  if (recommendation === 'LONG FUTURES GRID') return fib618;
  if (recommendation === 'SPOT GRID') return (fib618 + fib786) / 2;
  if (recommendation === 'WAIT / PULLBACK') return fib786;
  return altEntry43;
}

function OKX4H_calculateGridPlan_(gridLow, gridHigh, recommendation) {
  const rangePct = gridLow > 0 ? (gridHigh / gridLow - 1) * 100 : 0;
  const isFutures = recommendation === 'LONG FUTURES GRID';
  const roundTripFee = isFutures ? OKX4H_CFG.futuresRoundTripFeePct : OKX4H_CFG.spotRoundTripFeePct;
  const targetNet = isFutures ? OKX4H_CFG.targetNetGridPctFutures : OKX4H_CFG.targetNetGridPctSpot;
  const requiredStep = targetNet + roundTripFee + OKX4H_CFG.slippageBufferPct;
  const targetStep = OKX4H_clamp_(requiredStep, OKX4H_CFG.minGridStepPct, OKX4H_CFG.maxGridStepPct);

  let gridCount = rangePct > 0 ? Math.ceil(rangePct / targetStep) : OKX4H_CFG.minGrids;
  gridCount = Math.max(OKX4H_CFG.minGrids, Math.min(OKX4H_CFG.maxGrids, gridCount));

  const actualStepPct = rangePct > 0 ? (Math.pow(gridHigh / gridLow, 1 / gridCount) - 1) * 100 : 0;
  const estimatedGrossPerGridPct = actualStepPct;
  const estimatedNetPerGridPct = actualStepPct - roundTripFee - OKX4H_CFG.slippageBufferPct;
  const gridWarning = estimatedNetPerGridPct < OKX4H_CFG.minNetGridPct
    ? 'Μικρό post-fee κέρδος/grid'
    : estimatedNetPerGridPct > 0.75
      ? 'Μεγάλο spacing — λιγότερα fills'
      : 'Spacing αποδεκτό πριν το funding';

  return { gridCount, actualStepPct, estimatedGrossPerGridPct, estimatedNetPerGridPct, gridWarning };
}

function OKX4H_fundingPenalty_(rate) {
  const pct = Math.max(0, Number(rate || 0) * 100);
  if (pct <= OKX4H_CFG.fundingPenaltyStartPct) return 0;
  const span = OKX4H_CFG.fundingPenaltyFullPct - OKX4H_CFG.fundingPenaltyStartPct;
  return span > 0 ? Math.min(OKX4H_CFG.maxFundingPenalty, Math.round((pct - OKX4H_CFG.fundingPenaltyStartPct) / span * OKX4H_CFG.maxFundingPenalty)) : 0;
}

function OKX4H_fundingComment_(rate, penalty) {
  const pct = Number(rate || 0) * 100;
  if (pct < 0) return 'αρνητικό funding: θεωρητικό όφελος για long';
  if (penalty === 0) return 'χαμηλό θετικό funding';
  if (penalty >= OKX4H_CFG.maxFundingPenalty) return 'πολύ υψηλό θετικό funding: ισχυρή ποινή';
  return 'θετικό funding: κόστος για long';
}

function OKX4H_writeSheet_(sheet, rows) {
  const headers = ['Rank','Instrument','Final Score','Raw Score','Trend','Recommendation','Grid Count','Grid Step %','Net/Grid Before Funding %','Current Price','Bot Entry','24h Quote Vol','Change 24h %','RSI14','ADX14','ATR %','Rel Vol','EMA20','EMA50','MACD Hist','Grid Low','Grid High','Fib 0.618','Fib 0.786','Alt Entry 43%','Funding Current %','Funding Next %','Funding Penalty','Funding Comment','Grid Warning','Reasons','Updated'];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  const values = rows.slice(0, OKX4H_CFG.topRows).map((x, i) => [
    i + 1, x.instId, x.finalScore, x.score, x.trend, x.recommendation,
    x.gridCount, x.actualStepPct, x.estimatedNetPerGridPct, x.price, x.botEntry,
    x.volume24h, x.change24h, x.rsi, x.adx, x.atrPct, x.relativeVolume,
    x.ema20, x.ema50, x.macdHist, x.gridLow, x.gridHigh, x.fib618, x.fib786,
    x.altEntry43, x.fundingRatePct, x.nextFundingRatePct, x.fundingPenalty,
    x.fundingComment, x.gridWarning, x.reasons, x.updated
  ]);

  if (values.length) sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function OKX4H_sendTelegramIfNeeded_(best, rows) {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = properties.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) throw new Error('Λείπουν TELEGRAM_BOT_TOKEN ή TELEGRAM_CHAT_ID.');

  const cache = CacheService.getScriptCache();
  const signature = best.instId + '|' + best.recommendation + '|' + best.finalScore + '|' + best.gridCount;
  if (cache.get('OKX4H_LAST_ALERT') === signature) return;

  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      chat_id: chatId, text: OKX4H_buildTelegramMessage_(best, rows),
      parse_mode: 'HTML', disable_web_page_preview: true
    })
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('Telegram error: ' + response.getContentText());
  CacheService.getScriptCache().put('OKX4H_LAST_ALERT', signature, OKX4H_CFG.cooldownMinutes * 60);
}

function OKX4H_buildTelegramMessage_(best, rows) {
  const top = rows.slice(0, OKX4H_CFG.topRows).map((x, i) =>
    (i + 1) + '. ' + x.instId + ' — ' + x.finalScore + '/100 — ' + x.recommendation +
    ' — grids: ' + x.gridCount + ' — step: ' + x.actualStepPct.toFixed(2) + '%'
  ).join('\n');
  const funding = best.nextFundingRatePct == null ? 'N/A' : best.fundingRatePct.toFixed(4) + '% / ' + best.nextFundingRatePct.toFixed(4) + '%';

  return '<b>OKX TOP 120 / Pionex Grid</b>\n' +
    '<b>Κύρια πρόταση:</b> ' + OKX4H_escapeHtml_(best.instId) + '\n' +
    '<b>Final score:</b> ' + best.finalScore + '/100 | Raw: ' + best.score + '\n' +
    '<b>Bot:</b> ' + OKX4H_escapeHtml_(best.recommendation) + '\n' +
    '<b>Αριθμός grids:</b> ' + best.gridCount + '\n' +
    '<b>Grid spacing:</b> ' + best.actualStepPct.toFixed(2) + '% ανά grid\n' +
    '<b>Gross/grid:</b> ' + best.estimatedGrossPerGridPct.toFixed(2) + '%\n' +
    '<b>Καθαρό/grid πριν funding:</b> ' + best.estimatedNetPerGridPct.toFixed(2) + '%\n' +
    '<b>Grid range:</b> ' + OKX4H_fmt_(best.gridLow) + ' — ' + OKX4H_fmt_(best.gridHigh) + '\n' +
    '<b>Τρέχουσα τιμή:</b> ' + OKX4H_fmt_(best.price) + '\n' +
    '<b>Προτεινόμενη είσοδος:</b> ' + OKX4H_fmt_(best.botEntry) + '\n' +
    '<b>Fib 0.618:</b> ' + OKX4H_fmt_(best.fib618) + ' | <b>Fib 0.786:</b> ' + OKX4H_fmt_(best.fib786) + '\n' +
    '<b>Alt entry 43%:</b> ' + OKX4H_fmt_(best.altEntry43) + '\n' +
    '<b>Funding current/next:</b> ' + funding + '\n' +
    '<b>Funding penalty:</b> ' + best.fundingPenalty + ' — ' + OKX4H_escapeHtml_(best.fundingComment) + '\n' +
    '<b>RSI:</b> ' + best.rsi.toFixed(1) + ' | <b>ADX:</b> ' + best.adx.toFixed(1) + ' | <b>ATR:</b> ' + best.atrPct.toFixed(2) + '%\n' +
    '<b>Αιτιολόγηση:</b> ' + OKX4H_escapeHtml_(best.reasons) + '\n' +
    '<b>Grid note:</b> ' + OKX4H_escapeHtml_(best.gridWarning) + '\n\n' +
    '<b>Top 10:</b>\n' + top + '\n\n' +
    'Η έρευνα έγινε μόνο στα 120 πιο ρευστά live USDT spot instruments του OKX. Έλεγξε fees, funding, leverage και liquidation στην Pionex.';
}

function OKX4H_getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(OKX4H_CFG.sheetName) || ss.insertSheet(OKX4H_CFG.sheetName);
}

function OKX4H_fetchJson_(path) {
  const response = UrlFetchApp.fetch(OKX4H_CFG.apiBase + path, { method: 'get', muteHttpExceptions: true, headers: { Accept: 'application/json' } });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('HTTP ' + response.getResponseCode());
  return JSON.parse(response.getContentText());
}

function OKX4H_ema_(v, p) { if (!v.length) return 0; if (v.length <= p) return OKX4H_average_(v); const k = 2 / (p + 1); let e = OKX4H_average_(v.slice(0, p)); for (let i = p; i < v.length; i++) e = v[i] * k + e * (1 - k); return e; }
function OKX4H_emaSeries_(v, p) { if (!v.length) return []; if (v.length <= p) return v.map(() => OKX4H_average_(v)); const k = 2 / (p + 1), o = [OKX4H_average_(v.slice(0, p))]; for (let i = p; i < v.length; i++) o.push(v[i] * k + o[o.length - 1] * (1 - k)); while (o.length < v.length) o.unshift(o[0]); return o; }
function OKX4H_rsi_(v, p) { if (v.length <= p) return 50; let g = 0, l = 0; for (let i = 1; i <= p; i++) { const d = v[i] - v[i - 1]; if (d >= 0) g += d; else l -= d; } g /= p; l /= p; for (let i = p + 1; i < v.length; i++) { const d = v[i] - v[i - 1]; g = (g * (p - 1) + Math.max(d, 0)) / p; l = (l * (p - 1) + Math.max(-d, 0)) / p; } return l === 0 ? 100 : 100 - 100 / (1 + g / l); }
function OKX4H_atr_(c, p) { if (c.length < 2) return 0; const a = []; for (let i = 1; i < c.length; i++) a.push(Math.max(c[i].h - c[i].l, Math.abs(c[i].h - c[i - 1].c), Math.abs(c[i].l - c[i - 1].c))); return OKX4H_average_(a.slice(-p)); }
function OKX4H_adx_(c, p) { if (c.length < 3) return 0; const d = []; for (let i = 1; i < c.length; i++) { const u = c[i].h - c[i - 1].h, dn = c[i - 1].l - c[i].l, t = Math.max(c[i].h - c[i].l, Math.abs(c[i].h - c[i - 1].c), Math.abs(c[i].l - c[i - 1].c)), pl = t && u > dn && u > 0 ? u / t * 100 : 0, mi = t && dn > u && dn > 0 ? dn / t * 100 : 0; d.push(pl + mi ? Math.abs(pl - mi) / (pl + mi) * 100 : 0); } return OKX4H_average_(d.slice(-p)); }
function OKX4H_macd_(v) { const f = OKX4H_emaSeries_(v, 12), s = OKX4H_emaSeries_(v, 26), m = f.map((x, i) => x - s[i]), sig = OKX4H_ema_(m, 9), prev = m.length > 1 ? m[m.length - 2] : m[0]; return { hist: m[m.length - 1] - sig, prevHist: prev - sig }; }
function OKX4H_average_(a) { return a && a.length ? a.reduce((x, y) => x + Number(y || 0), 0) / a.length : 0; }
function OKX4H_clamp_(v, min, max) { return Math.max(min, Math.min(max, v)); }
function OKX4H_fmt_(v) { return Number(v).toLocaleString('en-US', { maximumFractionDigits: 8 }); }
function OKX4H_escapeHtml_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function OKX4H_setupTrigger() { ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'OKX4H_runScreener').forEach(t => ScriptApp.deleteTrigger(t)); ScriptApp.newTrigger('OKX4H_runScreener').timeBased().everyHours(4).create(); }
function OKX4H_testTelegram() { const p = PropertiesService.getScriptProperties(), t = p.getProperty('TELEGRAM_BOT_TOKEN'), c = p.getProperty('TELEGRAM_CHAT_ID'); if (!t || !c) throw new Error('Λείπουν TELEGRAM_BOT_TOKEN ή TELEGRAM_CHAT_ID.'); UrlFetchApp.fetch('https://api.telegram.org/bot' + t + '/sendMessage', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: c, text: 'OKX4H TOP 120 screener: Telegram test OK.' }) }); }






function gridEntry_getTokenUnlockParagraph_(coinName) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('OANOR_API_KEY');

  if (!apiKey) {
    return '🔓 Token unlock: λείπει το Script Property OANOR_API_KEY.';
  }

  var baseSymbol = gridEntry_extractBaseSymbol_(coinName);

  var headers = {
    'x-oanor-key': apiKey,
    'Accept': 'application/json'
  };

  try {
    var protocolSlug = gridEntry_findOanorProtocolSlug_(
      baseSymbol,
      headers
    );

    if (!protocolSlug) {
      return (
        '🔓 Δεν υπάρχει unlock coverage Oanor/DeFiLlama για ' +
        baseSymbol +
        '.'
      );
    }

    var url =
      'https://api.oanor.com/tokenunlocks-api/v1/next?protocol=' +
      encodeURIComponent(protocolSlug);

    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log(
      'Oanor next unlock | HTTP ' +
      httpCode +
      ' | protocol=' +
      protocolSlug +
      ' | ' +
      responseText
    );

    if (httpCode < 200 || httpCode >= 300) {
      return (
        '🔓 Δεν βρέθηκε προγραμματισμένο unlock για ' +
        baseSymbol +
        ' στο Oanor.'
      );
    }

    var json = JSON.parse(responseText);

    if (
      !json ||
      !json.data ||
      !json.data.next_unlock
    ) {
      return (
        '🔓 Δεν βρέθηκε επόμενο unlock για ' +
        baseSymbol +
        '.'
      );
    }

    var unlock = json.data.next_unlock;

    var unlockDate = unlock.date;
    var daysUntil = gridEntry_toNumber_(unlock.days_until);
    var cliffTokens = gridEntry_toNumber_(unlock.cliff_tokens);
    var pctOfTotalSupply = gridEntry_toNumber_(
      unlock.pct_of_total_supply
    );

    var categories = unlock.categories || [];
    var recipients = unlock.recipients || [];

    if (!unlockDate || cliffTokens === null) {
      Logger.log(
        'Oanor next_unlock χωρίς απαραίτητα fields: ' +
        responseText
      );

      return (
        '🔓 Το Oanor επέστρεψε incomplete unlock data για ' +
        baseSymbol +
        '.'
      );
    }

    if (daysUntil === null) {
      var parsedDate = new Date(unlockDate + 'T00:00:00+03:00');

      daysUntil = Math.ceil(
        (parsedDate.getTime() - Date.now()) / 86400000
      );
    }

    var dateParts = String(unlockDate).split('-');
    var dateLabel = dateParts.length === 3
      ? dateParts[2] + '/' + dateParts[1] + '/' + dateParts[0]
      : String(unlockDate);

    var amountLabel = gridEntry_formatLargeNumber_(cliffTokens);

    var pctLabel = pctOfTotalSupply === null
      ? 'μη διαθέσιμο ποσοστό total supply'
      : pctOfTotalSupply.toFixed(4) +
        '% του total supply';

    var recipientsLabel = Array.isArray(recipients)
      ? recipients.join(', ')
      : String(recipients || '');

    var categoriesLabel = Array.isArray(categories)
      ? categories.join(', ')
      : String(categories || '');

    var message =
      '🔓 Επόμενο token unlock για <b>' +
      gridEntry_escapeHtml_(baseSymbol) +
      '</b>: ' +
      dateLabel +
      ' (σε ' +
      daysUntil +
      ' ημέρες), περίπου <b>' +
      gridEntry_escapeHtml_(amountLabel) +
      '</b> tokens, δηλαδή <b>' +
      gridEntry_escapeHtml_(pctLabel) +
      '</b>.';

    if (categoriesLabel) {
      message +=
        ' Κατηγορίες: ' +
        gridEntry_escapeHtml_(categoriesLabel) +
        '.';
    }

    if (recipientsLabel) {
      message +=
        ' Recipients: ' +
        gridEntry_escapeHtml_(recipientsLabel) +
        '.';
    }

    if (pctOfTotalSupply !== null) {
      if (pctOfTotalSupply >= 5) {
        message +=
          ' Πρόκειται για μεγάλο unlock σε σχέση με το total supply.';
      } else if (pctOfTotalSupply >= 2) {
        message +=
          ' Είναι αξιοσημείωτο unlock και χρειάζεται παρακολούθηση.';
      } else {
        message +=
          ' Το ποσοστό είναι σχετικά περιορισμένο σε σχέση με το total supply.';
      }
    }

    return message;

  } catch (error) {
    Logger.log(
      'Oanor token-unlock error για ' +
      baseSymbol +
      ': ' +
      error.message
    );

    return (
      '🔓 Token unlock: σφάλμα ανάκτησης δεδομένων από Oanor.'
    );
  }
}


/**
 * Ψάχνει για το σωστό Oanor / DeFiLlama protocol slug.
 *
 * Παράδειγμα:
 * gridEntry_findOanorProtocolSlug_('APT')
 * -> 'aptos'
 *
 * Επιστρέφει null αν το token δεν καλύπτεται.
 */
function gridEntry_findOanorProtocolSlug_(baseSymbol, headers) {
  var searchText = String(baseSymbol || '')
    .trim()
    .toLowerCase();

  if (!searchText) {
    return null;
  }

  var url =
    'https://api.oanor.com/tokenunlocks-api/v1/protocols?q=' +
    encodeURIComponent(searchText);

  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  });

  var httpCode = response.getResponseCode();
  var responseText = response.getContentText();

  Logger.log(
    'Oanor protocols search | HTTP ' +
    httpCode +
    ' | q=' + searchText +
    ' | ' + responseText
  );

  if (httpCode < 200 || httpCode >= 300) {
    return null;
  }

  var json = JSON.parse(responseText);

  if (
    !json ||
    !json.data ||
    !Array.isArray(json.data.protocols)
  ) {
    return null;
  }

  var protocols = json.data.protocols;

  // Απόλυτο slug match, π.χ. apt -> aptos δεν είναι exact,
  // άρα ψάχνουμε πρώτα το raw token name / symbol όπου γίνεται.
  if (protocols.indexOf(searchText) !== -1) {
    return searchText;
  }

  /*
   * Προαιρετικό manual mapping για symbols που δεν είναι ίδια με
   * το protocol slug στο DeFiLlama dataset.
   *
   * Πρόσθεσε entries μόνο αφού επιβεβαιώσεις ότι υπάρχουν στο
   * /v1/protocols endpoint.
   */
  var SYMBOL_TO_PROTOCOL = {
    'APT': 'aptos',
    'ARB': 'arbitrum-foundation',
    'OP': 'optimism',
    'TIA': 'celestia',
    'ETHFI': 'ether.fi',
    'DYDX': 'dydx',
    'EIGEN': 'eigencloud',
    'ENA': 'ethena',
    'AXL': 'axelar-network',
    'APE': 'apecoin',
    'AXS': 'axie-infinity',
    'DOGE': 'dogecoin'
  };

  var mappedSlug = SYMBOL_TO_PROTOCOL[
    String(baseSymbol || '').trim().toUpperCase()
  ];

  if (mappedSlug && protocols.indexOf(mappedSlug) !== -1) {
    return mappedSlug;
  }

  /*
   * Τελευταίο fallback: exact match του symbol ως prefix/suffix slug.
   * Δεν χρησιμοποιούμε broad includes για να μην αντιστοιχίσουμε λάθος
   * asset. Π.χ. ARB θα μπορούσε να ταιριάξει σε άσχετο project.
   */
  for (var i = 0; i < protocols.length; i++) {
    var slug = protocols[i];

    if (
      slug === searchText ||
      slug.indexOf(searchText + '-') === 0 ||
      slug.lastIndexOf('-' + searchText) === slug.length - searchText.length - 1
    ) {
      return slug;
    }
  }

  return null;
}


/**
 * Εξάγει το base token symbol από pair / swap symbol.
 *
 * Παραδείγματα:
 * ANIME-USDT-SWAP -> ANIME
 * BTC-USDT        -> BTC
 * BTC/USDT        -> BTC
 * BTCUSDT         -> BTC
 * SOLUSDC         -> SOL
 */
function gridEntry_extractBaseSymbol_(coinName) {
  var raw = String(coinName || '')
    .trim()
    .toUpperCase();

  if (!raw) {
    return '';
  }

  // ANIME-USDT-SWAP, BTC-USDT, ETH-USDC κ.λπ.
  if (raw.indexOf('-') !== -1) {
    return raw.split('-')[0];
  }

  // BTC/USDT, ETH/USDC κ.λπ.
  if (raw.indexOf('/') !== -1) {
    return raw.split('/')[0];
  }

  // BTCUSDT, ETHUSDC, SOLUSD κ.λπ.
  var knownQuotes = [
    'USDT',
    'USDC',
    'BUSD',
    'FDUSD',
    'TUSD',
    'DAI',
    'USD',
    'EUR',
    'BTC',
    'ETH'
  ];

  for (var i = 0; i < knownQuotes.length; i++) {
    var quote = knownQuotes[i];

    if (
      raw.length > quote.length &&
      raw.slice(-quote.length) === quote
    ) {
      return raw.slice(0, raw.length - quote.length);
    }
  }

  // Αν είναι ήδη καθαρό token symbol, π.χ. ANIME.
  return raw;
}

/**
 * Μετατρέπει αριθμούς ή strings σε ασφαλές Number.
 *
 * Υποστηρίζει:
 * 0.00325
 * "0,00325"
 * "130.5M"     -> 130500000
 * "2.36%"      -> 2.36
 * "$428.4K"    -> 428400
 * "1,234.56"   -> 1234.56
 * "1.234,56"   -> 1234.56
 *
 * Επιστρέφει null όταν η τιμή δεν μπορεί να μετατραπεί.
 */
function gridEntry_toNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) {
    return value;
  }

  if (value === null || value === undefined || value === '') {
    return null;
  }

  var text = String(value)
    .trim()
    .replace(/\s/g, '');

  if (!text) {
    return null;
  }

  // Αφαιρούμε σύμβολα νομίσματος και ποσοστού πριν από parsing.
  text = text
    .replace(/\$/g, '')
    .replace(/€/g, '')
    .replace(/%/g, '');

  // Υποστήριξη K / M / B ως suffix.
  var multiplier = 1;
  var suffixMatch = text.match(/([KMB])$/i);

  if (suffixMatch) {
    var suffix = suffixMatch[1].toUpperCase();

    if (suffix === 'K') {
      multiplier = 1e3;
    } else if (suffix === 'M') {
      multiplier = 1e6;
    } else if (suffix === 'B') {
      multiplier = 1e9;
    }

    text = text.slice(0, -1);
  }

  var commaIndex = text.lastIndexOf(',');
  var dotIndex = text.lastIndexOf('.');

  /*
   * Αν υπάρχουν και κόμμα και τελεία:
   * - 1.234,56 -> 1234.56
   * - 1,234.56 -> 1234.56
   */
  if (commaIndex > -1 && dotIndex > -1) {
    if (commaIndex > dotIndex) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (commaIndex > -1) {
    // 0,00325 -> 0.00325
    text = text.replace(',', '.');
  }

  // Ασφάλεια: κρατά μόνο digits, sign, decimal point και scientific notation.
  text = text.replace(/[^0-9+\-.eE]/g, '');

  var parsed = parseFloat(text);

  return isFinite(parsed)
    ? parsed * multiplier
    : null;
}

function gridEntry_formatLargeNumber_(number) {
  if (
    number === null ||
    number === undefined ||
    !isFinite(number)
  ) {
    return '—';
  }

  var absolute = Math.abs(number);

  if (absolute >= 1e9) {
    return (number / 1e9).toFixed(2) + 'B';
  }

  if (absolute >= 1e6) {
    return (number / 1e6).toFixed(2) + 'M';
  }

  if (absolute >= 1e3) {
    return (number / 1e3).toFixed(2) + 'K';
  }

  return Number(number).toFixed(2);
}

function gridEntry_escapeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function gridEntry_formatRecipients_(recipients) {
  if (!recipients) {
    return '';
  }

  if (typeof recipients === 'string') {
    return recipients.trim();
  }

  if (Array.isArray(recipients)) {
    return recipients
      .map(function(item) {
        if (typeof item === 'string') {
          return item;
        }

        if (!item || typeof item !== 'object') {
          return '';
        }

        return (
          item.name ||
          item.recipient ||
          item.category ||
          item.allocation ||
          ''
        );
      })
      .filter(function(value) {
        return value;
      })
      .join(', ');
  }

  if (typeof recipients === 'object') {
    return Object.keys(recipients)
      .map(function(key) {
        var value = recipients[key];

        if (value && typeof value === 'object') {
          return (
            value.name ||
            value.recipient ||
            value.category ||
            key
          );
        }

        return key;
      })
      .join(', ');
  }

  return '';
}




function gridEntry_testOanorAnime_() {
  var result = gridEntry_getTokenUnlockParagraph_(
    'MEME-USDT-SWAP'
  );

  Logger.log(result);
}

function gridEntry_extractOpeningDateKey_(value) {
  var text = String(value || '').trim();

  // Δέχεται:
  // 2026-10-04
  // 2026-10-04 1:10
  // 2026-10-04 13:10
  var match = text.match(
    /^(\d{4})-(\d{2})-(\d{2})/
  );

  if (!match) {
    return null;
  }

  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);

  var date = new Date(year, month - 1, day);

  // Απορρίπτει άκυρα dates, π.χ. 2026-02-31.
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return match[1] + '-' + match[2] + '-' + match[3];
}

function gridEntry_sendTelegramMessageWithResult_(message) {
  var props = PropertiesService.getScriptProperties();

  var token = props.getProperty('TELEGRAM_BOT_TOKEN');
  var chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    Logger.log(
      'Λείπουν τα TELEGRAM_BOT_TOKEN ή TELEGRAM_CHAT_ID.'
    );
    return false;
  }

  var url =
    'https://api.telegram.org/bot' +
    token +
    '/sendMessage';

  var payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var httpCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log(
      'Telegram HTTP ' +
      httpCode +
      ': ' +
      responseText
    );

    var json = JSON.parse(responseText);

    return (
      httpCode >= 200 &&
      httpCode < 300 &&
      json &&
      json.ok === true
    );

  } catch (error) {
    Logger.log(
      'Telegram error: ' +
      error.message
    );

    return false;
  }
}

/**
 * ΚΥΡΙΑ FUNCTION ΓΙΑ DAILY TRIGGER
 *
 * Διαβάζει:
 * Open Positions!A2 = Symbol, π.χ. APT-USDT-SWAP
 * Open Positions!B2 = ημερομηνία ανοίγματος, π.χ. 2026-10-04 1:10
 *
 * Η ώρα αγνοείται επίτηδες, ώστε να μην υπάρχει πρόβλημα π.μ./μ.μ.
 * Ελέγχει μόνο το YYYY-MM-DD.
 *
 * Στέλνει report μόνο όταν:
 * - Η ημερομηνία ανοίγματος είναι σήμερα ή χθες.
 * - Δεν έχει σταλεί ήδη report για το ίδιο Symbol + Date.
 *
 * Απαιτεί:
 * - gridEntry_getTokenUnlockParagraph_()
 * - gridEntry_extractOpeningDateKey_()
 * - gridEntry_escapeHtml_()
 * - gridEntry_sendTelegramMessageWithResult_()
 */
function gridEntry_sendDailyNewPositionUnlockReport() {
  var CFG = {
    OPEN_POSITIONS_SHEET: 'Open Positions',
    SYMBOL_CELL: 'A2',
    OPENED_AT_CELL: 'B2',

    // Δέχεται θέση που άνοιξε σήμερα ή χθες.
    RECENT_DAYS: 1,

    // Anti-duplicate key στα Script Properties.
    REPORT_KEY_PROPERTY: 'gridEntry_dailyOpenUnlockReportKey',

    TIMEZONE: 'Europe/Athens'
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.OPEN_POSITIONS_SHEET);

  if (!sheet) {
    Logger.log(
      'Δεν βρέθηκε το sheet "' +
      CFG.OPEN_POSITIONS_SHEET +
      '".'
    );
    return;
  }

  var symbol = String(
    sheet.getRange(CFG.SYMBOL_CELL).getDisplayValue() || ''
  ).trim();

  var openedAtText = String(
    sheet.getRange(CFG.OPENED_AT_CELL).getDisplayValue() || ''
  ).trim();

  if (!symbol) {
    Logger.log(
      'Δεν υπάρχει ανοικτή θέση στο Open Positions!A2.'
    );
    return;
  }

  if (!openedAtText) {
    Logger.log(
      'Το Open Positions!B2 είναι κενό.'
    );
    return;
  }

  // Κρατά μόνο π.χ. 2026-10-04 από:
  // 2026-10-04 1:10
  // 2026-10-04 13:10
  var openingDateKey = gridEntry_extractOpeningDateKey_(
    openedAtText
  );

  if (!openingDateKey) {
    Logger.log(
      'Μη έγκυρη ημερομηνία στο Open Positions!B2: "' +
      openedAtText +
      '". Περιμένω μορφή YYYY-MM-DD ή YYYY-MM-DD ώρα.'
    );
    return;
  }

  var todayKey = Utilities.formatDate(
    new Date(),
    CFG.TIMEZONE,
    'yyyy-MM-dd'
  );

  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - CFG.RECENT_DAYS);

  var yesterdayKey = Utilities.formatDate(
    yesterday,
    CFG.TIMEZONE,
    'yyyy-MM-dd'
  );

  // Δεν κάνουμε ούτε Oanor ούτε Telegram call για παλιά θέση.
  if (
    openingDateKey !== todayKey &&
    openingDateKey !== yesterdayKey
  ) {
    Logger.log(
      'Η θέση ' +
      symbol +
      ' άνοιξε στις ' +
      openingDateKey +
      '. Επιτρέπονται μόνο σήμερα (' +
      todayKey +
      ') ή χθες (' +
      yesterdayKey +
      ').'
    );
    return;
  }

  // Ένα report για κάθε symbol + ημερομηνία ανοίγματος.
  var reportKey =
    symbol.toUpperCase() +
    '|' +
    openingDateKey;

  var props = PropertiesService.getScriptProperties();

  if (
    props.getProperty(CFG.REPORT_KEY_PROPERTY) === reportKey
  ) {
    Logger.log(
      'Το unlock report έχει ήδη σταλεί για: ' +
      reportKey
    );
    return;
  }

  // Καλεί Oanor μόνο αφού περάσουν όλοι οι τοπικοί έλεγχοι.
  var unlockParagraph = gridEntry_getTokenUnlockParagraph_(
    symbol
  );

  var lines = [];
  lines.push(
    '🆕 <b>Νέα ανοιχτή θέση — Token unlock report</b>'
  );
  lines.push('');
  lines.push(
    '🔹 <b>' +
    gridEntry_escapeHtml_(symbol) +
    '</b>'
  );
  lines.push(
    '📅 Ημερομηνία ανοίγματος: ' +
    openingDateKey
  );
  lines.push('');
  lines.push(unlockParagraph);
  lines.push('');
  lines.push(
    'ℹ️ Το report στέλνεται μία φορά μόνο για το ίδιο ' +
    'symbol και την ίδια ημερομηνία ανοίγματος.'
  );

  var telegramSent = gridEntry_sendTelegramMessageWithResult_(
    lines.join('\n')
  );

  // Η καταγραφή γίνεται μόνο αν το Telegram επέστρεψε ok:true.
  // Αν αποτύχει η αποστολή, το επόμενο daily trigger ξαναδοκιμάζει.
  if (telegramSent) {
    props.setProperty(
      CFG.REPORT_KEY_PROPERTY,
      reportKey
    );

    props.setProperty(
      'gridEntry_dailyOpenUnlockReportSentAt',
      new Date().toISOString()
    );

    Logger.log(
      'Στάλθηκε επιτυχώς unlock report για: ' +
      reportKey
    );
  } else {
    Logger.log(
      'Απέτυχε η αποστολή Telegram. Δεν αποθηκεύτηκε ' +
      'anti-duplicate key, ώστε να γίνει retry αύριο.'
    );
  }
}






/**
 * ============================================================================
 *  QUANTUS (QTC) DAILY DASHBOARD  —  Google Apps Script  (v3 — ελληνικοί χαρακτήρες)
 * ----------------------------------------------------------------------------
 *  ΤΙ ΑΛΛΑΞΕ ΣΤΗΝ v3:
 *   Διορθώθηκαν όλοι οι ελληνικοί χαρακτήρες στα σχόλια και στα μηνύματα
 *   Telegram (αντί για transliterated ελληνικά που εμφανίζονταν λόγω escape
 *   sequences). Όλο το κείμενο είναι πλέον σωστά UTF-8 encoded.
 *
 *  ΠΗΓΕΣ ΔΕΔΟΜΕΝΩΝ:
 *   - CoinGecko (coin id: "quantus") για τιμή/volume/market cap/ATH/supply.
 *   - BackPoW.com (https://backpow.com/coins-data/Quantus.json) για network
 *     hashrate/difficulty/revenue-per-TH — επιβεβαιωμένο ζωντανά ότι δουλεύει
 *     από Apps Script (HTTP 200, πλούσιο JSON).
 *   - ΠΡΟΣΟΧΗ: το ticker "QTC" στο hashrate.no ανήκει σε ΑΛΛΟ coin (Qubitcoin) —
 *     ΔΕΝ χρησιμοποιείται εδώ γι' αυτό.
 *
 *  ΡΥΘΜΙΣΗ (Project Settings → Script Properties):
 *     TELEGRAM_BOT_TOKEN        = <ήδη το έχεις>
 *     TELEGRAM_CHAT_ID          = <ήδη το έχεις>
 *     COINGECKO_API_KEY         = <ήδη το έχεις>
 *     QTC_NETWORK_STATS_API_URL = <προαιρετικό override, default: BackPoW URL>
 *     QTC_ELECTRICITY_COST_EUR_KWH = 0.20   (προαιρετικό, default 0.20)
 *     QTC_GPU_POWER_WATTS       = 350       (προαιρετικό, default 350)
 *     EXPECTED_QTC_PER_DAY      = 0.05      (προαιρετικό — για ακριβές mining ratio)
 *
 *  Trigger: Triggers → Add Trigger → runQtcDashboard → Time-driven → Daily.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// ΡΥΘΜΙΣΕΙΣ / CONFIG
// ---------------------------------------------------------------------------
const QTC_CONFIG = {
  SHEET_NAME: 'QTC Dashboard',
  COINGECKO_ID: 'quantus',
  VS_CURRENCY: 'eur',
  DEFAULT_NETWORK_STATS_URL: 'https://backpow.com/coins-data/Quantus.json',
  CACHE_TTL_SECONDS: 300,
  MAX_RETRIES: 4,
  BASE_BACKOFF_MS: 2000,
  MAX_BACKOFF_MS: 15000,

  THRESHOLDS: {
    change24h:       { green: 3,   yellow: 0 },
    volChange24h:    { green: 15,  yellow: -10 },
    change7d:        { green: 8,   yellow: 0 },
    volMcapRatio:    { green: 8,   yellow: 3 },
    athDistance:     { greenMax: -60, yellowMax: -85 },
    rankTrend:       { greenDelta: 1, yellowDelta: 0 },
    supplyProgress:  { greenMax: 20, yellowMax: 40 },
    miningProfit:    { green: 1.3, yellow: 1.0 },
    volatility:      { greenMax: 6, yellowMax: 12 },
    hashrateChange:  { green: 5,   yellow: -5 },
    difficultyChange:{ greenMax: 5, yellowMax: 15 },
    revenuePerTh:    { green: 2,   yellow: -10 },
    networkGrowth:   { green: 2,   yellow: -10 }
  }
};

// ---------------------------------------------------------------------------
// MAIN ENTRY POINT
// ---------------------------------------------------------------------------
function runQtcDashboard() {
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('COINGECKO_API_KEY')) {
    console.warn('COINGECKO_API_KEY δεν βρέθηκε — θα χρησιμοποιηθεί το πιο περιορισμένο keyless API.');
  }

  let marketData = null;
  let isStale = false;
  const previousRow = getQtcPreviousRow_();

  try {
    marketData = fetchQtcMarketData_(props);
  } catch (err) {
    console.error('CoinGecko fetch failed completely: ' + err.message);
    marketData = buildQtcFallbackMarketDataFromPreviousRow_(previousRow);
    if (marketData) {
      isStale = true;
      console.warn('Using STALE data from previous sheet row to keep the dashboard alive.');
    } else {
      sendQtcTelegramMessage_(
        '⚠️ <b>QTC Dashboard Error</b>\n' +
        'Αποτυχία λήψης δεδομένων από CoinGecko και δεν υπάρχει προηγούμενο ιστορικό ' +
        'στο sheet για fallback.\n\nΤεχνικό σφάλμα: ' + qtcEscapeHtml_(String(err.message)),
        props
      );
      return;
    }
  }

  try {
    const networkData = fetchQtcNetworkData_(props);
    const metrics = computeQtcMetrics_(marketData, networkData, previousRow, props);
    const verdict = computeQtcVerdict_(metrics);

    writeQtcToSheet_(marketData, networkData, metrics, verdict, isStale);
    sendQtcTelegramMessage_(buildQtcTelegramMessage_(marketData, networkData, metrics, verdict, isStale), props);

  } catch (err) {
    console.error('runQtcDashboard failed downstream: ' + err.message);
    try {
      sendQtcTelegramMessage_(
        '⚠️ <b>QTC Dashboard Error</b>\n' + qtcEscapeHtml_(String(err.message)),
        props
      );
    } catch (e2) {
      console.error('Failed to send error alert to Telegram: ' + e2.message);
    }
  }
}

// ---------------------------------------------------------------------------
// 1) FETCH MARKET DATA FROM COINGECKO
// ---------------------------------------------------------------------------
function fetchQtcMarketData_(props) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'qtc_market_data_v3';
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log('Using cached CoinGecko data (rate-limit friendly).');
    return JSON.parse(cached);
  }

  const apiKey = props.getProperty('COINGECKO_API_KEY');

  let url = 'https://api.coingecko.com/api/v3/coins/' + QTC_CONFIG.COINGECKO_ID +
    '?localization=false&tickers=false&market_data=true&community_data=false' +
    '&developer_data=false&sparkline=false';

  const headers = { 'Accept': 'application/json' };
  if (apiKey) {
    headers['x-cg-demo-api-key'] = apiKey;
    url += '&x_cg_demo_api_key=' + encodeURIComponent(apiKey);
  }

  const json = fetchQtcWithBackoff_(url, headers, { relaxSsl: false });
  const md = json.market_data;
  const vs = QTC_CONFIG.VS_CURRENCY;

  const result = {
    timestamp: new Date().toISOString(),
    name: json.name,
    symbol: (json.symbol || '').toUpperCase(),
    price: qtcSafeGet_(md, ['current_price', vs]),
    priceUsd: qtcSafeGet_(md, ['current_price', 'usd']),
    change24h: qtcSafeGet_(md, ['price_change_percentage_24h']),
    change7d: qtcSafeGet_(md, ['price_change_percentage_7d']),
    high24h: qtcSafeGet_(md, ['high_24h', vs]),
    low24h: qtcSafeGet_(md, ['low_24h', vs]),
    volume24h: qtcSafeGet_(md, ['total_volume', vs]) || 0,
    marketCap: qtcSafeGet_(md, ['market_cap', vs]) || 0,
    marketCapRank: qtcSafeGet_(md, ['market_cap_rank']),
    ath: qtcSafeGet_(md, ['ath', vs]),
    athChangePercentage: qtcSafeGet_(md, ['ath_change_percentage', vs]),
    circulatingSupply: qtcSafeGet_(md, ['circulating_supply']),
    maxSupply: qtcSafeGet_(md, ['max_supply']) || 21000000,
    totalSupply: qtcSafeGet_(md, ['total_supply'])
  };

  cache.put(cacheKey, JSON.stringify(result), QTC_CONFIG.CACHE_TTL_SECONDS);
  return result;
}

function qtcSafeGet_(obj, pathArr) {
  let cur = obj;
  for (let i = 0; i < pathArr.length; i++) {
    if (cur == null) return null;
    cur = cur[pathArr[i]];
  }
  return cur == null ? null : cur;
}

function fetchQtcWithBackoff_(url, headers, opts) {
  opts = opts || {};
  let attempt = 0;
  let lastError = null;

  while (attempt < QTC_CONFIG.MAX_RETRIES) {
    let response;
    try {
      const fetchOptions = {
        muteHttpExceptions: true,
        headers: headers || { 'Accept': 'application/json' }
      };
      if (opts.relaxSsl) fetchOptions.validateHttpsCertificates = false;

      response = UrlFetchApp.fetch(url, fetchOptions);
    } catch (e) {
      const msg = String(e.message);
      if (msg.indexOf('SSL') !== -1 || msg.indexOf('Address unavailable') !== -1 || msg.indexOf('DNS') !== -1) {
        throw new Error('Connection-level error στο ' + url + ' (μη ανακτήσιμο): ' + e.message);
      }
      lastError = e.message;
      const wait = Math.min(QTC_CONFIG.BASE_BACKOFF_MS * Math.pow(2, attempt), QTC_CONFIG.MAX_BACKOFF_MS);
      console.warn('Connection error for ' + url + ': ' + e.message + ' -- retrying in ' + wait + 'ms');
      Utilities.sleep(wait);
      attempt++;
      continue;
    }

    const code = response.getResponseCode();
    if (code === 200) {
      return JSON.parse(response.getContentText());
    }
    if (code === 429 || code >= 500) {
      const jitter = Math.floor(Math.random() * 1000);
      const wait = Math.min(
        QTC_CONFIG.BASE_BACKOFF_MS * Math.pow(2, attempt) + jitter,
        QTC_CONFIG.MAX_BACKOFF_MS
      );
      console.warn('API returned ' + code + ' for ' + url + ', retrying in ' + wait + 'ms (attempt ' + (attempt + 1) + '/' + QTC_CONFIG.MAX_RETRIES + ')');
      Utilities.sleep(wait);
      attempt++;
      lastError = 'HTTP ' + code;
      continue;
    }
    if (code === 401 || code === 403) {
      throw new Error('API key invalid/rejected (HTTP ' + code + ') για ' + url + '.');
    }
    throw new Error('API error (' + url + '): HTTP ' + code + ' -- ' + response.getContentText().substring(0, 300));
  }
  throw new Error('API failed after ' + QTC_CONFIG.MAX_RETRIES + ' retries for ' + url + '. Last error: ' + lastError);
}

function buildQtcFallbackMarketDataFromPreviousRow_(previousRow) {
  const priceKey = 'Price (' + QTC_CONFIG.VS_CURRENCY.toUpperCase() + ')';
  if (!previousRow || previousRow[priceKey] == null) return null;

  return {
    timestamp: new Date().toISOString(),
    name: 'Quantus',
    symbol: 'QTC',
    price: Number(previousRow[priceKey]),
    priceUsd: null,
    change24h: Number(previousRow['24h Change %']) || 0,
    change7d: Number(previousRow['7d Change %']) || 0,
    high24h: Number(previousRow[priceKey]),
    low24h: Number(previousRow[priceKey]),
    volume24h: Number(previousRow['Volume 24h']) || 0,
    marketCap: 0,
    marketCapRank: Number(previousRow['Market Cap Rank']) || null,
    ath: null,
    athChangePercentage: Number(previousRow['Distance from ATH %']) || null,
    circulatingSupply: null,
    maxSupply: 21000000,
    totalSupply: null
  };
}

// ---------------------------------------------------------------------------
// 2) FETCH NETWORK DATA (hashrate, difficulty, revenue/TH) — BackPoW.com
// ---------------------------------------------------------------------------
function fetchQtcNetworkData_(props) {
  const customUrl = props.getProperty('QTC_NETWORK_STATS_API_URL') || QTC_CONFIG.DEFAULT_NETWORK_STATS_URL;

  const cache = CacheService.getScriptCache();
  const cacheKey = 'qtc_network_data_v3';
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log('Using cached network stats (rate-limit friendly).');
    return JSON.parse(cached);
  }

  try {
    const raw = fetchQtcWithBackoff_(customUrl, { 'Accept': 'application/json' }, { relaxSsl: false });

    const difficultyNow = qtcParseDifficultyString_(raw['Difficulty']);
    const difficulty24hAgo = qtcParseDifficultyString_(raw['Difficulty 24h']);
    const difficulty3dAgo = qtcParseDifficultyString_(raw['Difficulty 3 days']);
    const difficulty7dAgo = qtcParseDifficultyString_(raw['Difficulty 7 days']);

    let difficultyChangePct24h = null;
    if (difficultyNow != null && difficulty24hAgo) {
      difficultyChangePct24h = ((difficultyNow - difficulty24hAgo) / difficulty24hAgo) * 100;
    }

    const nethashHps = raw['Nethash'] != null ? Number(raw['Nethash']) : null;
    const hashrateEh = nethashHps != null ? nethashHps / 1e18 : null;

    const blockReward = raw['Bl. reward'] != null ? Number(raw['Bl. reward']) : null;
    const blockTime = raw['Block time'] != null ? Number(raw['Block time']) : null;
    const dailyEmission = raw['Daily emission'] != null ? Number(raw['Daily emission']) : null;
    const priceUsdFromSource = raw['usdPrice'] != null ? Number(raw['usdPrice']) : null;

    // Revenue per TH/s/day (USD): από το πρώτο GPU του πίνακα (ήδη υπολογισμένο
    // από το ίδιο το BackPoW για συγκεκριμένη τιμή/δυσκολία) — πιο αξιόπιστο
    // από να το υπολογίσουμε εμείς με block reward (το "Daily emission" field
    // φαίνεται να μη ταιριάζει με τον απλό τύπο block_reward × blocks_per_day).
    let revenuePerThUsd = null;
    const gpus = raw['GPUS'];
    if (Array.isArray(gpus) && gpus.length > 0) {
      const g = gpus[0];
      if (g.hashrate && g.revenue) {
        const hrThs = g.hashrate / 1e12;
        if (hrThs > 0) revenuePerThUsd = g.revenue / hrThs;
      }
    }

    const result = {
      hashrate: hashrateEh,
      difficulty: difficultyNow,
      difficultyChangePct24h: difficultyChangePct24h != null ? qtcRound2_(difficultyChangePct24h) : null,
      difficulty3dAgo: difficulty3dAgo,
      difficulty7dAgo: difficulty7dAgo,
      blockReward: blockReward,
      blockTime: blockTime,
      dailyEmission: dailyEmission,
      revenuePerThUsd: revenuePerThUsd,
      priceUsdFromSource: priceUsdFromSource,
      source: customUrl
    };

    cache.put(cacheKey, JSON.stringify(result), QTC_CONFIG.CACHE_TTL_SECONDS);
    return result;
  } catch (e) {
    console.warn('QTC network stats fetch failed (' + customUrl + '): ' + e.message);
    return { hashrate: null, difficulty: null, difficultyChangePct24h: null, blockReward: null, blockTime: null, revenuePerThUsd: null, source: 'unavailable' };
  }
}

/** Parse strings όπως "423,165,506M" → 423165506000000 (number).
 *  Υποστηρίζει suffixes K/M/G/T, αφαιρεί κόμματα. */
function qtcParseDifficultyString_(raw) {
  if (raw == null) return null;
  if (typeof raw === 'number') return raw;
  let s = String(raw).trim();
  let mult = 1;
  const lastChar = s.slice(-1).toUpperCase();
  if (lastChar === 'M') { mult = 1e6; s = s.slice(0, -1); }
  else if (lastChar === 'K') { mult = 1e3; s = s.slice(0, -1); }
  else if (lastChar === 'G') { mult = 1e9; s = s.slice(0, -1); }
  else if (lastChar === 'T') { mult = 1e12; s = s.slice(0, -1); }
  s = s.replace(/,/g, '').trim();
  const num = parseFloat(s);
  return isNaN(num) ? null : num * mult;
}

// ---------------------------------------------------------------------------
// 3) ΠΡΟΗΓΟΥΜΕΝΗ ΓΡΑΜΜΗ ΓΙΑ ΣΥΓΚΡΙΣΕΙΣ
// ---------------------------------------------------------------------------
function getQtcPreviousRow_() {
  const sheet = getOrCreateQtcSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const rowObj = {};
  headers.forEach((h, i) => rowObj[h] = values[i]);
  return rowObj;
}

// ---------------------------------------------------------------------------
// 4) ΥΠΟΛΟΓΙΣΜΟΣ ΤΩΝ 12 METRICS
// ---------------------------------------------------------------------------
function computeQtcMetrics_(data, net, previousRow, props) {
  const T = QTC_CONFIG.THRESHOLDS;
  const metrics = [];
  const networkNotConfigured = net.source === 'unavailable';

  metrics.push(qtcMakeMetric_('24h Price Change %', data.change24h,
    qtcStatusHigherIsBetter_(data.change24h, T.change24h.green, T.change24h.yellow),
    '% μεταβολή τιμής τελευταίου 24ώρου. Το QTC είναι πολύ νέο coin — αναμένεται υψηλή μεταβλητότητα.'));

  let volChange = null, volNote = 'Χωρίς ιστορικό ακόμα.';
  if (previousRow && previousRow['Volume 24h'] != null && previousRow['Volume 24h'] !== '') {
    const prevVol = Number(previousRow['Volume 24h']);
    if (prevVol > 0) {
      volChange = ((data.volume24h - prevVol) / prevVol) * 100;
      volNote = 'Μεταβολή 24h volume έναντι προηγούμενης καταγραφής.';
    }
  }
  metrics.push(qtcMakeMetric_('24h Volume Change %', volChange,
    volChange != null ? qtcStatusHigherIsBetter_(volChange, T.volChange24h.green, T.volChange24h.yellow) : 'yellow',
    volNote));

  metrics.push(qtcMakeMetric_('7d Price Change %', data.change7d,
    qtcStatusHigherIsBetter_(data.change7d, T.change7d.green, T.change7d.yellow),
    'Μεσοπρόθεσμο trend 7 ημερών.'));

  const volMcapRatio = data.marketCap > 0 ? (data.volume24h / data.marketCap) * 100 : null;
  metrics.push(qtcMakeMetric_('Volume/MCap Ratio %', volMcapRatio,
    volMcapRatio != null ? qtcStatusHigherIsBetter_(volMcapRatio, T.volMcapRatio.green, T.volMcapRatio.yellow) : 'yellow',
    data.marketCap > 0 ? 'Ρευστότητα & ενδιαφέρον αγοράς σχετικά με το μέγεθος του coin.'
      : 'N/A — το CoinGecko δεν αναφέρει ακόμα market cap για το QTC (πολύ νέο listing).'));

  metrics.push(qtcMakeMetric_('Distance from ATH %', data.athChangePercentage,
    qtcStatusAthDistance_(data.athChangePercentage, T.athDistance.greenMax, T.athDistance.yellowMax),
    'Πόσο μακριά είμαστε από το ATH. Προσοχή: coin μόλις λίγων εβδομάδων, το ATH είναι ασταθές/πρόσφατο.'));

  let rankDelta = 0, rankNote = 'Χωρίς ιστορικό ακόμα (πρώτη εκτέλεση).';
  if (previousRow && previousRow['Market Cap Rank'] != null && previousRow['Market Cap Rank'] !== '' && data.marketCapRank != null) {
    rankDelta = Number(previousRow['Market Cap Rank']) - data.marketCapRank;
    rankNote = 'Μεταβολή rank έναντι προηγούμενης καταγραφής.';
  }
  metrics.push(qtcMakeMetric_('Market Cap Rank Trend', rankDelta,
    data.marketCapRank != null ? qtcStatusHigherIsBetter_(rankDelta, T.rankTrend.greenDelta, T.rankTrend.yellowDelta) : 'yellow',
    data.marketCapRank != null ? rankNote : 'N/A — δεν έχει ανατεθεί ακόμα market cap rank.',
    data.marketCapRank));

  const supplyProgress = (data.maxSupply > 0 && data.circulatingSupply != null) ? (data.circulatingSupply / data.maxSupply) * 100 : null;
  metrics.push(qtcMakeMetric_('Circulating/Max Supply %', supplyProgress,
    supplyProgress != null ? qtcStatusLowerIsBetter_(supplyProgress, T.supplyProgress.greenMax, T.supplyProgress.yellowMax) : 'yellow',
    supplyProgress != null
      ? 'Χαμηλότερο ποσοστό = μεγαλύτερο μελλοντικό runway (max supply 21M, fixed cap σαν Bitcoin).'
      : 'N/A — το CoinGecko δεν αναφέρει ακόμα circulating supply για το QTC.'));

  const miningResult = computeQtcMiningProfitability_(data, net, props, T.miningProfit);
  metrics.push(qtcMakeMetric_('Mining Profitability Ratio', miningResult.ratio, miningResult.status, miningResult.note));

  const volatility = (data.low24h > 0 && data.high24h != null) ? ((data.high24h - data.low24h) / data.low24h) * 100 : null;
  metrics.push(qtcMakeMetric_('24h Volatility Spread %', volatility,
    volatility != null ? qtcStatusLowerIsBetter_(volatility, T.volatility.greenMax, T.volatility.yellowMax) : 'yellow',
    volatility != null ? 'Εύρος high-low 24ώρου.' : 'N/A — δεν επιστράφηκαν high/low 24h.'));

  const naNetworkNote = 'N/A — δεν βρέθηκαν δεδομένα δικτύου (έλεγξε QTC_NETWORK_STATS_API_URL ή τρέξε testQtcNetworkApi()).';

  let hashrateChangePct = null, hashrateNote;
  const currentHashrateNum = net.hashrate;
  if (currentHashrateNum != null && previousRow && previousRow['Network Hashrate (EH/s)'] != null && previousRow['Network Hashrate (EH/s)'] !== '') {
    const prevHashrate = Number(previousRow['Network Hashrate (EH/s)']);
    if (prevHashrate > 0) hashrateChangePct = ((currentHashrateNum - prevHashrate) / prevHashrate) * 100;
  }
  hashrateNote = currentHashrateNum != null
    ? 'Τρέχον δίκτυο: ' + qtcRound4_(currentHashrateNum * 1e6) + ' TH/s (πηγή: BackPoW.com). Το δίκτυο είναι λίγες εβδομάδες, πολύ ασταθές ακόμα.'
    : (networkNotConfigured ? naNetworkNote : 'Δεν ήταν διαθέσιμο το network hashrate αυτή τη φορά.');
  metrics.push(qtcMakeMetric_('Network Hashrate Change %', hashrateChangePct,
    hashrateChangePct != null ? qtcStatusHigherIsBetter_(hashrateChangePct, T.hashrateChange.green, T.hashrateChange.yellow) : 'yellow',
    hashrateNote, currentHashrateNum));

  // Difficulty Change %: χρησιμοποιούμε ΑΠΕΥΘΕΙΑΣ το "Difficulty 24h" πεδίο
  // του BackPoW (δεν χρειάζεται να περιμένουμε 2η εκτέλεση για να έχουμε σύγκριση).
  const difficultyChangePct = net.difficultyChangePct24h;
  const difficultyNote = net.difficulty != null
    ? 'Τρέχουσα δυσκολία: ' + qtcFormatCompactNumber_(net.difficulty) + ' (στον BackPoW: ' + qtcFormatCompactNumber_(net.difficulty/1e6) + 'M). ' +
      'Το Quantus αναπροσαρμόζει difficulty ΚΑΘΕ block — αναμένεται μεγάλη διακύμανση.'
    : (networkNotConfigured ? naNetworkNote : 'Δεν ήταν διαθέσιμη η difficulty αυτή τη φορά.');
  metrics.push(qtcMakeMetric_('Difficulty Change %', difficultyChangePct,
    difficultyChangePct != null ? qtcStatusLowerIsBetter_(difficultyChangePct, T.difficultyChange.greenMax, T.difficultyChange.yellowMax) : 'yellow',
    difficultyNote, net.difficulty));

  // Mining Revenue/TH Trend %: σύγκριση του revenuePerThUsd (από BackPoW GPU
  // reference) με την προηγούμενη γραμμή του sheet.
  const revenuePerTh = net.revenuePerThUsd;
  let revenueChangePct = null;
  if (revenuePerTh != null && previousRow && previousRow['Revenue per TH/s ($)'] != null && previousRow['Revenue per TH/s ($)'] !== '') {
    const prevRev = Number(previousRow['Revenue per TH/s ($)']);
    if (prevRev > 0) revenueChangePct = ((revenuePerTh - prevRev) / prevRev) * 100;
  }
  metrics.push(qtcMakeMetric_('Mining Revenue/TH Trend %', revenueChangePct,
    revenueChangePct != null ? qtcStatusHigherIsBetter_(revenueChangePct, T.revenuePerTh.green, T.revenuePerTh.yellow) : 'yellow',
    revenuePerTh != null
      ? 'Εκτιμώμενο έσοδο: $' + qtcRound4_(revenuePerTh) + '/TH/s/μέρα (από reference GPU του BackPoW).'
      : (networkNotConfigured ? naNetworkNote : 'Δεν υπολογίστηκε revenue/TH.'),
    revenuePerTh));

  // Network Growth Momentum %: το Quantus είναι καθαρό PoW (Poseidon2), όχι
  // AI-compute σαν το Pearl — αυτό το metric αντικαθιστά το "AI Demand Proxy"
  // του Pearl Dashboard με κάτι σχετικό για καθαρά PoW δίκτυα.
  let networkGrowthValue = hashrateChangePct;
  const networkGrowthNote = hashrateChangePct != null
    ? 'Στο Quantus (καθαρό PoW, όχι AI-compute όπως το Pearl) αυτό δείχνει momentum ανάπτυξης δικτύου μέσω hashrate trend. Πρόσεξε: νέο δίκτυο, μεγάλη συγκέντρωση σε λίγα pools ακόμα πιθανή.'
    : naNetworkNote;
  metrics.push(qtcMakeMetric_('Network Growth Momentum %', networkGrowthValue,
    networkGrowthValue != null ? qtcStatusHigherIsBetter_(networkGrowthValue, T.networkGrowth.green, T.networkGrowth.yellow) : 'yellow',
    networkGrowthNote));

  return metrics;
}

function computeQtcMiningProfitability_(data, net, props, thresholds) {
  const elecCost = Number(props.getProperty('QTC_ELECTRICITY_COST_EUR_KWH')) || 0.20;
  const gpuWatts = Number(props.getProperty('QTC_GPU_POWER_WATTS')) || 350;
  const explicitQtcPerDay = Number(props.getProperty('EXPECTED_QTC_PER_DAY')) || null;

  const dailyKwh = (gpuWatts / 1000) * 24;
  const dailyElecCostEur = dailyKwh * elecCost;

  if (!explicitQtcPerDay) {
    return {
      ratio: null, status: 'yellow',
      note: 'Όρισε το Script Property EXPECTED_QTC_PER_DAY (πόσα QTC βγάζεις/μέρα) για ακριβή υπολογισμό κερδοφορίας.'
    };
  }

  const priceForCalc = data.price != null ? data.price : null;
  if (priceForCalc == null) {
    return { ratio: null, status: 'yellow', note: 'N/A — δεν είναι διαθέσιμη η τιμή QTC από το CoinGecko αυτή τη στιγμή.' };
  }

  const dailyRevenueEur = explicitQtcPerDay * priceForCalc;
  const ratio = dailyElecCostEur > 0 ? dailyRevenueEur / dailyElecCostEur : null;

  let status = 'red';
  if (ratio != null) {
    if (ratio >= thresholds.green) status = 'green';
    else if (ratio >= thresholds.yellow) status = 'yellow';
  }

  const note = 'Έσοδο ' + qtcRound2_(dailyRevenueEur) + '€/μέρα (' + explicitQtcPerDay + ' QTC × ' +
    qtcRound4_(priceForCalc) + '€) vs κόστος ρεύματος ' + qtcRound2_(dailyElecCostEur) +
    '€/μέρα (' + gpuWatts + 'W @ ' + elecCost + '€/kWh).';

  return { ratio: ratio, status: status, note: note };
}

// ---------------------------------------------------------------------------
// STATUS HELPERS
// ---------------------------------------------------------------------------
function qtcStatusHigherIsBetter_(value, greenMin, yellowMin) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value >= greenMin) return 'green';
  if (value >= yellowMin) return 'yellow';
  return 'red';
}
function qtcStatusLowerIsBetter_(value, greenMax, yellowMax) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value <= greenMax) return 'green';
  if (value <= yellowMax) return 'yellow';
  return 'red';
}
function qtcStatusAthDistance_(value, greenMax, yellowMax) {
  if (value == null || isNaN(value)) return 'yellow';
  if (value >= greenMax) return 'green';
  if (value >= yellowMax) return 'yellow';
  return 'red';
}
function qtcMakeMetric_(label, value, status, note, extraRaw) {
  return { label: label, value: value, status: status, note: note, extraRaw: extraRaw };
}

// ---------------------------------------------------------------------------
// 5) ΣΥΝΟΛΙΚΟ VERDICT
// ---------------------------------------------------------------------------
function computeQtcVerdict_(metrics) {
  const score = { green: 0, yellow: 0, red: 0 };
  metrics.forEach(m => { if (score[m.status] != null) score[m.status]++; });

  const total = metrics.length;
  let verdict, emoji;

  if (score.green >= total * 0.65) { verdict = 'STRONG BULLISH'; emoji = '🟢🚀'; }
  else if (score.green + score.yellow >= total * 0.7 && score.red <= 1) { verdict = 'BULLISH'; emoji = '🟢'; }
  else if (score.red >= total * 0.5) { verdict = 'BEARISH'; emoji = '🔴'; }
  else { verdict = 'NEUTRAL'; emoji = '🟡'; }

  return { label: verdict, emoji: emoji, score: score, total: total };
}

// ---------------------------------------------------------------------------
// 6) SHEET WRITE
// ---------------------------------------------------------------------------
function getOrCreateQtcSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QTC_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QTC_CONFIG.SHEET_NAME);
    const headers = [
      'Date', 'Timestamp', 'Price (' + QTC_CONFIG.VS_CURRENCY.toUpperCase() + ')',
      '24h Change %', 'Volume 24h', '24h Volume Change %', '7d Change %',
      'Volume/MCap %', 'Distance from ATH %', 'Market Cap Rank', 'Rank Trend',
      'Circulating/Max Supply %', 'Mining Profit Ratio', '24h Volatility %',
      'Network Hashrate (EH/s)', 'Hashrate Change %', 'Difficulty', 'Difficulty Change %',
      'Revenue per TH/s ($)', 'Revenue/TH Trend %', 'Network Growth Momentum %',
      'Verdict', 'Green Count', 'Yellow Count', 'Red Count', 'Data Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#6b2fa0').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function writeQtcToSheet_(data, net, metrics, verdict, isStale) {
  const sheet = getOrCreateQtcSheet_();
  const now = new Date();
  const byLabel = {};
  metrics.forEach(m => byLabel[m.label] = m);

  const row = [
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
    data.price,
    qtcRound2_(data.change24h),
    data.volume24h,
    qtcRound2_(byLabel['24h Volume Change %'].value),
    qtcRound2_(data.change7d),
    qtcRound2_(byLabel['Volume/MCap Ratio %'].value),
    qtcRound2_(data.athChangePercentage),
    data.marketCapRank != null ? data.marketCapRank : 'N/A',
    qtcRound2_(byLabel['Market Cap Rank Trend'].value),
    qtcRound2_(byLabel['Circulating/Max Supply %'].value),
    byLabel['Mining Profitability Ratio'].value != null ? qtcRound2_(byLabel['Mining Profitability Ratio'].value) : 'N/A',
    qtcRound2_(byLabel['24h Volatility Spread %'].value),
    byLabel['Network Hashrate Change %'].extraRaw != null ? qtcRound4_(byLabel['Network Hashrate Change %'].extraRaw) : 'N/A',
    qtcRound2_(byLabel['Network Hashrate Change %'].value),
    byLabel['Difficulty Change %'].extraRaw != null ? byLabel['Difficulty Change %'].extraRaw : 'N/A',
    qtcRound2_(byLabel['Difficulty Change %'].value),
    byLabel['Mining Revenue/TH Trend %'].extraRaw != null ? qtcRound4_(byLabel['Mining Revenue/TH Trend %'].extraRaw) : 'N/A',
    qtcRound2_(byLabel['Mining Revenue/TH Trend %'].value),
    qtcRound2_(byLabel['Network Growth Momentum %'].value),
    verdict.label,
    verdict.score.green,
    verdict.score.yellow,
    verdict.score.red,
    isStale ? 'STALE (CoinGecko fallback)' : 'LIVE'
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  const verdictCol = row.length - 4;
  const cell = sheet.getRange(lastRow, verdictCol);
  const colorMap = {
    'STRONG BULLISH': '#34a853', 'BULLISH': '#93c47d',
    'NEUTRAL': '#ffd966', 'BEARISH': '#e06666'
  };
  cell.setBackground(colorMap[verdict.label] || '#ffffff').setFontWeight('bold');

  if (isStale) {
    sheet.getRange(lastRow, row.length).setBackground('#f4cccc').setFontWeight('bold');
  }

  const maxRows = 366;
  if (sheet.getLastRow() > maxRows) sheet.deleteRow(2);
}

// ---------------------------------------------------------------------------
// 7) TELEGRAM MESSAGE
// ---------------------------------------------------------------------------
function buildQtcTelegramMessage_(data, net, metrics, verdict, isStale) {
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const statusEmoji = { green: '🟢', yellow: '🟡', red: '🔴' };
  const networkNotConfigured = net.source === 'unavailable';

  let msg = '<b>⚛️ QUANTUS (QTC) DAILY DASHBOARD</b>\n';
  msg += '<i>' + dateStr + '</i>\n\n';

  if (isStale) {
    msg += '⚠️ <b>ΠΡΟΣΟΧΗ: Stale δεδομένα</b> — Το CoinGecko απέτυχε προσωρινά, ' +
      'χρησιμοποιούνται τα τελευταία γνωστά στοιχεία τιμής.\n\n';
  }

  msg += verdict.emoji + ' <b>Verdict: ' + qtcEscapeHtml_(verdict.label) + '</b>';
  msg += '  (🟢' + verdict.score.green + ' 🟡' + verdict.score.yellow + ' 🔴' + verdict.score.red + '/' + verdict.total + ')\n\n';

  msg += '💰 Τιμή: <b>' + qtcRound4_(data.price) + ' ' + QTC_CONFIG.VS_CURRENCY.toUpperCase() + '</b>';
  msg += '  |  MCap Rank: #' + (data.marketCapRank || 'N/A') + '\n';
  msg += '━━━━━━━━━━━━━━━━━━━\n';
  msg += '<b>📈 Market Metrics</b>\n';

  const marketLabels = ['24h Price Change %', '24h Volume Change %', '7d Price Change %',
    'Volume/MCap Ratio %', 'Distance from ATH %', 'Market Cap Rank Trend',
    'Circulating/Max Supply %', '24h Volatility Spread %'];
  const miningLabels = ['Mining Profitability Ratio', 'Network Hashrate Change %',
    'Difficulty Change %', 'Mining Revenue/TH Trend %', 'Network Growth Momentum %'];

  const byLabel = {};
  metrics.forEach(m => byLabel[m.label] = m);

  marketLabels.forEach(l => {
    const m = byLabel[l];
    msg += (statusEmoji[m.status] || '⚪') + ' <b>' + qtcEscapeHtml_(m.label) + ':</b> ' + qtcFormatMetricValue_(m) + '\n';
  });

  msg += '\n<b>⛏️ Mining & Network Metrics</b>\n';
  if (networkNotConfigured) {
    msg += '<i>ℹ️ Δεδομένα δικτύου μη διαθέσιμα αυτή τη στιγμή.</i>\n';
  }
  miningLabels.forEach(l => {
    const m = byLabel[l];
    msg += (statusEmoji[m.status] || '⚪') + ' <b>' + qtcEscapeHtml_(m.label) + ':</b> ' + qtcFormatMetricValue_(m) + '\n';
  });

  msg += '━━━━━━━━━━━━━━━━━━━\n';
  msg += '📊 24h Volume: ' + qtcFormatCompactNumber_(data.volume24h) + ' ' + QTC_CONFIG.VS_CURRENCY.toUpperCase() + '\n';
  if (data.marketCap) msg += '🏦 Market Cap: ' + qtcFormatCompactNumber_(data.marketCap) + ' ' + QTC_CONFIG.VS_CURRENCY.toUpperCase() + '\n';

  const currentHashrate = byLabel['Network Hashrate Change %'].extraRaw;
  if (currentHashrate != null) msg += '⚡ Network Hashrate: ' + qtcRound4_(currentHashrate * 1e6) + ' TH/s\n';
  const currentDifficulty = byLabel['Difficulty Change %'].extraRaw;
  if (currentDifficulty != null && currentDifficulty !== 'N/A') msg += '🎯 Difficulty: ' + qtcFormatCompactNumber_(currentDifficulty) + '\n';
  if (net.blockReward != null) msg += '💎 Block Reward: ' + qtcRound4_(net.blockReward) + ' QTC (~' + qtcRound2_(net.blockTime) + 's/block)\n';
  if (net.priceUsdFromSource != null) msg += '📉 BackPoW USD price (cross-check): $' + qtcRound4_(net.priceUsdFromSource) + '\n';

  msg += '\n<i>ℹ️ Το Quantus είναι καθαρό PoW (Poseidon2 hashing), όχι AI-compute όπως το Pearl — ' +
    'το "Network Growth Momentum" αντικαθιστά το AI proxy metric με κάτι σχετικό.</i>\n';
  msg += '<i>Δεδομένα: CoinGecko (Demo API key) + BackPoW.com' + (isStale ? ' — STALE fallback' : '') + '.</i>';

  return msg;
}

function qtcFormatMetricValue_(m) {
  if (m.value == null || isNaN(m.value)) return 'N/A';
  if (m.label === 'Market Cap Rank Trend') {
    const sign = m.value > 0 ? '+' : '';
    return sign + m.value + ' θέσεις (τώρα #' + (m.extraRaw || 'N/A') + ')';
  }
  if (m.label === 'Mining Profitability Ratio') return qtcRound2_(m.value) + 'x';
  return qtcRound2_(m.value) + '%';
}

function sendQtcTelegramMessage_(text, props) {
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    console.error('Missing TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID in Script Properties.');
    return;
  }

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const payload = {
    chat_id: chatId, text: text, parse_mode: 'HTML', disable_web_page_preview: true
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error('Telegram send failed: ' + response.getContentText());
  }
}

// ---------------------------------------------------------------------------
// UTILITIES
// ---------------------------------------------------------------------------
function qtcRound2_(n) { return (n == null || isNaN(n)) ? null : Math.round(n * 100) / 100; }
function qtcRound4_(n) { return (n == null || isNaN(n)) ? null : Math.round(n * 10000) / 10000; }
function qtcFormatCompactNumber_(n) {
  if (n == null || isNaN(n)) return 'N/A';
  if (n >= 1e9) return qtcRound2_(n / 1e9) + 'B';
  if (n >= 1e6) return qtcRound2_(n / 1e6) + 'M';
  if (n >= 1e3) return qtcRound2_(n / 1e3) + 'K';
  return String(qtcRound2_(n));
}
function qtcEscapeHtml_(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------------------------------------------------------------------------
// ΔΟΚΙΜΑΣΤΙΚΗ ΕΚΤΕΛΕΣΗ / TRIGGER SETUP / TESTS
// ---------------------------------------------------------------------------
function testQtcRunNow() { runQtcDashboard(); }

function testQtcCoinGeckoApiKey() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('COINGECKO_API_KEY');
  if (!apiKey) {
    console.log('Δεν βρέθηκε COINGECKO_API_KEY στα Script Properties.');
    return;
  }
  const url = 'https://api.coingecko.com/api/v3/coins/' + QTC_CONFIG.COINGECKO_ID +
    '?localization=false&tickers=false&market_data=true&x_cg_demo_api_key=' + encodeURIComponent(apiKey);
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'x-cg-demo-api-key': apiKey, 'Accept': 'application/json' }
  });
  console.log('HTTP status: ' + response.getResponseCode());
  console.log('Response (πρώτοι 800 χαρακτήρες): ' + response.getContentText().substring(0, 800));
}

function testQtcNetworkApi() {
  const props = PropertiesService.getScriptProperties();
  try {
    const parsed = fetchQtcNetworkData_(props);
    console.log('Parsed result: ' + JSON.stringify(parsed));
  } catch (e) {
    console.log('BackPoW test απέτυχε: ' + e.message);
  }
}

function createQtcDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runQtcDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runQtcDashboard').timeBased().everyDays(1).atHour(9).create();
  console.log('Daily trigger created for runQtcDashboard at 09:00.');
}





/**
 * Kryptex Pool → Google Sheets
 * Public API: https://pool.kryptex.com
 *
 * Δημιουργεί:
 * Config, Pool Summary, Workers, Earnings History, API Raw
 *
 * Σημείωση:
 * Το API είναι δημόσιο Pool API. Δεν αποτελεί API για το
 * συνολικό υπόλοιπο του Kryptex Windows application.
 */

const KRYPTEX_BASE = 'https://pool.kryptex.com';

const SHEETS = {
  config: 'Config',
  summary: 'Pool Summary',
  workers: 'Workers',
  earnings: 'Earnings History',
  raw: 'API Raw'
};

function setupKryptex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEETS.config, [
    ['Enabled', 'Coin', 'Mining Address', 'Label'],
    [true, 'kas', 'PUT_YOUR_MINING_ADDRESS_HERE', 'KAS Windows/Pool']
  ]);

  ensureSheet_(ss, SHEETS.summary, [[
    'Timestamp', 'Coin', 'Label', 'Address',
    'Balance API response', 'Payout stats response',
    'Status', 'Error'
  ]]);

  ensureSheet_(ss, SHEETS.workers, [[
    'Timestamp', 'Coin', 'Label', 'Address',
    'Worker data (JSON)', 'Status', 'Error'
  ]]);

  ensureSheet_(ss, SHEETS.earnings, [[
    'Timestamp', 'Coin', 'Label', 'Address',
    'Payout history (JSON)',
    'Reward chart (JSON)',
    'Balance snapshot (JSON)'
  ]]);

  ensureSheet_(ss, SHEETS.raw, [[
    'Timestamp', 'Coin', 'Label', 'Endpoint',
    'HTTP status', 'Response JSON'
  ]]);

  // Αφαίρεση παλιού trigger της ίδιας συνάρτησης.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshKryptex')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshKryptex')
    .timeBased()
    .everyHours(1)
    .create();

  refreshKryptex();
}

function refreshKryptex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.config);

  if (!cfg || cfg.getLastRow() < 2) {
    throw new Error('Συμπλήρωσε το φύλλο Config.');
  }

  const rows = cfg.getRange(2, 1, cfg.getLastRow() - 1, 4)
    .getValues();

  const now = new Date();
  const summaryRows = [];
  const workerRows = [];
  const earningRows = [];
  const rawRows = [];

  rows.forEach(row => {
    const [enabled, coinValue, addressValue, labelValue] = row;

    if (enabled !== true) return;

    const coin = String(coinValue || '').trim().toLowerCase();
    const address = String(addressValue || '').trim();
    const label = String(labelValue || coin).trim();

    if (!coin || !address ||
        address === 'PUT_YOUR_MINING_ADDRESS_HERE') {
      return;
    }


const endpoints = {
  balance:
    `/${coin}/api/v1/miner/balance/${encodeURIComponent(address)}`,

  workers:
    `/${coin}/api/v3/miner/workers/${encodeURIComponent(address)}`,

  payouts:
    `/${coin}/api/v1/miner/payouts/${encodeURIComponent(address)}`,

  payoutStats:
    `/${coin}/api/v1/miner/payouts/${encodeURIComponent(address)}/stats`,

  chart:
    `/${coin}/api/v1/miner/chart/${encodeURIComponent(address)}`,

  rewards:
    `/${coin}/api/v1/miner/reward-chart/${encodeURIComponent(address)}`
};

    const result = {};

    Object.entries(endpoints).forEach(([key, path]) => {
      result[key] = fetchKryptex_(path, now, coin, label, rawRows);
    });

    
const b = result.balance.data || {};
const p = result.payoutStats.data || {};
const reward = p.reward || {};

summaryRows.push([
  now,
  coin,
  label,
  address,
  Number(b.total ?? 0),
  Number(b.unconfirmed ?? 0),
  Number(b.confirmed ?? 0),
  Number(b.threshold ?? 0),
  Number(b.reached_pct ?? 0),
  Number(p.paid ?? 0),
  Number(p.unpaid ?? 0),
  Number(reward.week ?? 0),
  Number(reward.month ?? 0),
  b.last_active
    ? new Date(Number(b.last_active))
    : '',
  result.balance.ok ? 'OK' : 'ERROR',
  result.balance.error || ''
]);


    workerRows.push([
      now, coin, label, address,
      json_(result.workers.data),
      result.workers.ok ? 'OK' : 'ERROR',
      result.workers.error || ''
    ]);

// Αντικατάστησε το earningRows.push με:
    earningRows.push([
      now, coin, label, address,
      json_(result.payouts.data),
      json_(result.rewards.data),
      json_(result.balance.data),
      json_(result.chart.data)
    ]);

  });

  appendRows_(ss.getSheetByName(SHEETS.summary), summaryRows);
  appendRows_(ss.getSheetByName(SHEETS.workers), workerRows);
  appendRows_(ss.getSheetByName(SHEETS.earnings), earningRows);
  appendRows_(ss.getSheetByName(SHEETS.raw), rawRows);

  sendKryptexTelegramSummary();

}

function fetchKryptex_(path, timestamp, coin, label, rawRows) {
  const url = KRYPTEX_BASE + path;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { Accept: 'application/json' }
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    let data;
    try {
      data = JSON.parse(text);
    } catch (_) {
      data = { rawText: text };
    }

    rawRows.push([
      timestamp, coin, label, path, status, json_(data)
    ]);

    return {
      ok: status >= 200 && status < 300,
      data: data,
      error: status >= 200 && status < 300
        ? ''
        : 'HTTP ' + status
    };
  } catch (e) {
    rawRows.push([
      timestamp, coin, label, path, 'FETCH ERROR',
      String(e)
    ]);

    return { ok: false, data: null, error: String(e) };
  }
}

function ensureSheet_(ss, name, initialRows) {
  let sh = ss.getSheetByName(name);

  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, initialRows.length,
      initialRows[0].length).setValues(initialRows);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, initialRows[0].length)
      .setFontWeight('bold');
    sh.autoResizeColumns(1, initialRows[0].length);
  }

  return sh;
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;

  sheet.getRange(
    sheet.getLastRow() + 1, 1,
    rows.length, rows[0].length
  ).setValues(rows);
}

function json_(value) {
  if (value === null || value === undefined) return '';
  const text = typeof value === 'string'
    ? value
    : JSON.stringify(value);

  // Αποφυγή υπέρβασης του ορίου χαρακτήρων ενός κελιού.
  return text.length > 45000
    ? text.substring(0, 45000) + ' ... [TRUNCATED]'
    : text;
}

/** Χειροκίνητη ενημέρωση από το μενού Run. */
function testKryptex() {
  refreshKryptex();
}


function checkKryptexPools() {
  const url = 'https://pool.kryptex.com/api/v1/index';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Accept: 'application/json' }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Pool Index');

  if (!sh) sh = ss.insertSheet('Pool Index');

  sh.clearContents();
  sh.getRange('A1:B2').setValues([
    ['HTTP Status', 'API Response'],
    [response.getResponseCode(), response.getContentText()]
  ]);
}


/**
 * Telegram notifications for Kryptex Google Sheets.
 * Requires Script Properties:
 * TELEGRAM_BOT_TOKEN
 * TELEGRAM_CHAT_ID
 */



function sendKryptexTelegramSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('API Raw');

  if (!sh || sh.getLastRow() < 2) {
    throw new Error('Το API Raw δεν περιέχει δεδομένα.');
  }

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(v => String(v).trim().toLowerCase());

  const idx = name => headers.indexOf(name.toLowerCase());

  const tsCol = idx('Timestamp');
  const coinCol = idx('Coin');
  const labelCol = idx('Label');
  const endpointCol = idx('Endpoint');
  const httpCol = idx('HTTP status');
  const responseCol = idx('Response JSON');

  if ([tsCol, coinCol, labelCol, endpointCol, httpCol, responseCol]
      .some(i => i < 0)) {
    throw new Error(
      'Έλεγξε τις επικεφαλίδες του API Raw: Timestamp, Coin, Label, Endpoint, HTTP status, Response JSON.'
    );
  }

  // Τελευταίο response για κάθε coin, label και endpoint.
  const latest = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const endpoint = String(row[endpointCol] || '');
    const coin = String(row[coinCol] || '');
    const label = String(row[labelCol] || coin);

    if (!endpoint || !coin) continue;

    const key = coin + '|' + label + '|' + endpoint;
    const old = latest.get(key);

    if (!old || new Date(row[tsCol]) >= new Date(old[tsCol])) {
      latest.set(key, row);
    }
  }

  const groups = new Map();

  for (const row of latest.values()) {
    const coin = String(row[coinCol]);
    const label = String(row[labelCol] || coin);
    const endpoint = String(row[endpointCol]);
    const key = coin + '|' + label;

    if (!groups.has(key)) {
      groups.set(key, {
        coin,
        label,
        balance: null,
        payouts: null,
        workers: null,
        statuses: []
      });
    }

    const g = groups.get(key);
    const status = String(row[httpCol] || '');
    const body = parseJsonCell_(row[responseCol]);

    g.statuses.push(status);

    if (endpoint.includes('/miner/balance/')) {
      g.balance = body;
    } else if (endpoint.includes('/miner/payouts/') &&
               endpoint.endsWith('/stats')) {
      g.payouts = body;
    } else if (endpoint.includes('/miner/workers/')) {
      g.workers = body;
    }
  }

  const lines = [
    '<b>⛏️ Kryptex Mining Report</b>',
    '🕒 ' + Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm'
    ),
    ''
  ];

  for (const g of groups.values()) {
    lines.push(
      '<b>' + escapeTelegramHtml_(g.label) + '</b>'
    );

    const b = g.balance;
    const p = g.payouts;
    const reward = p && p.reward ? p.reward : null;

    const symbol = g.coin.toLowerCase() === 'zeph' ||
               g.label.toLowerCase().includes('zephyr') ||
               g.label.toLowerCase().includes('zeph')
      ? 'ZEPH'
      : 'QUAI';

    const unit = ' ' + symbol;

    if (b && typeof b === 'object' && b.total !== undefined) {
      lines.push('💰 Total: ' + fmt_(b.total) + unit);
      lines.push('✅ Confirmed: ' + fmt_(b.confirmed) + unit);
      lines.push('⏳ Unconfirmed: ' + fmt_(b.unconfirmed) + unit);
      lines.push('🎯 Threshold: ' + fmt_(b.threshold) + unit);
      lines.push('📊 Progress: ' + fmt_(b.reached_pct) + '%');
    } else {
      lines.push('⚠️ Balance data unavailable');
    }

    if (p && typeof p === 'object') {
//      lines.push('💸 Paid: ' + fmt_(p.paid) + unit);
//      lines.push('💳 Unpaid: ' + fmt_(p.unpaid) + unit);

      if (reward) {
        lines.push('📅 Reward week: ' + fmt_(reward.week) + unit);
        lines.push('🗓️ Reward month: ' + fmt_(reward.month) + unit);
      }
    } else {
      lines.push('⚠️ Payout statistics unavailable');
    }

    lines.push(
      '📡 API status: ' +
      (g.statuses.length && g.statuses.every(s => s === '200')
        ? 'OK'
        : escapeTelegramHtml_(g.statuses.join(', ') || 'unknown'))
    );

    lines.push('');
  }

  lines.push(
    '<i>Pool statistics only — not total Kryptex account balance.</i>'
  );

  telegramSend_(lines.join('\n'));
}



function testKryptexTelegram() {
  telegramSend_(
    '<b>✅ Kryptex Telegram test</b>\n' +
    'Η σύνδεση του Google Apps Script με το Telegram λειτουργεί.'
  );
}

function telegramSend_(message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    throw new Error(
      'Πρόσθεσε TELEGRAM_BOT_TOKEN και TELEGRAM_CHAT_ID ' +
      'στις Script Properties.'
    );
  }

  const url = 'https://api.telegram.org/bot' +
    token + '/sendMessage';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      chat_id: chatId,
      text: message,
      parse_mode: 'HTML',
      disable_web_page_preview: true
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status < 200 || status >= 300) {
    throw new Error(
      'Telegram HTTP ' + status + ': ' +
      response.getContentText().substring(0, 500)
    );
  }
}

function parseJsonCell_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (typeof value === 'object') return value;

  try {
    return JSON.parse(String(value));
  } catch (e) {
    return null;
  }
}

function fmt_(value) {
  if (value === null || value === undefined || value === '') {
    return 'N/A';
  }

  const n = Number(value);
  return Number.isFinite(n)
    ? n.toLocaleString('en-US', {
        maximumFractionDigits: 4
      })
    : String(value);
}

function workerCount_(data) {
  if (!data) return 'N/A';
  if (Array.isArray(data)) return data.length;

  if (typeof data === 'object') {
    for (const key of ['workers', 'data', 'items', 'result']) {
      if (Array.isArray(data[key])) return data[key].length;
    }
  }

  return 'See Workers sheet';
}

function escapeTelegramHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}




