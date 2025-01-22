import yfinance as yf
import pandas as pd
import csv
from ta import trend, momentum, volume  # 引入技術指標相關模塊
import datetime

# 基本面指標閾值
THRESHOLDS = {
    "PE": {"low": 10, "high": 25},
    "PB": {"low": 1,  "high": 5},
    "Beta": {"low": 0.8, "high": 1.2},
    "ROE": {"low": 0.1, "high": 0.2},
    "DividendYield": {"low": 2, "high": 6},

    "PEG": {"low": 1, "high": 2},
    "OperatingMargin": {"low": 0.1, "high": 0.3},
    "DebtToEquity": {"low": 50, "high": 100},
    "CurrentRatio": {"low": 1, "high": 2},
    "QuickRatio": {"low": 1, "high": 2},
    "RevenueGrowth": {"low": 0.05, "high": 0.2},
    "EarningsGrowth": {"low": 0.05, "high": 0.2}
}

# ETF 指標閾值
ETF_THRESHOLDS = {
    "ETF_PE":     {"low": 10, "high": 25},
    "ETF_Yield":  {"low": 2,  "high": 6},
    "ETF_Beta3Y": {"low": 0.8, "high": 1.2},
    "threeYearAverageReturn": {"low": 0.03, "high": 0.1},
    "fiveYearAverageReturn":  {"low": 0.03, "high": 0.1},
    "totalAssets": {"low": 1e11, "high": 5e11},
}

# 技術指標閾值
TECHNICAL_THRESHOLDS = {
    "RSI": {"low": 30, "high": 70},  # RSI <30 超賣，>70 超買
    "MACD": {"signal_positive": True},  # MACD 線與信號線的關係
    "SMA_50_vs_SMA_200": {"relation": "above"},  # 短期均線在長期均線之上
    "ADX": {"low": 20, "high": 40},  # 趨勢強度
    "Volume_Trend": {"trend": "increasing"},  # 成交量趨勢
}

def compare_with_threshold(value, metric_key):
    thresholds = THRESHOLDS.get(metric_key, None)
    if not thresholds or value is None:
        return None, "", 0

    low, high = thresholds["low"], thresholds["high"]
    less_is_better = {"PE", "PB", "PEG", "Beta", "DebtToEquity"}
    more_is_better = {"ROE", "OperatingMargin", "DividendYield", "CurrentRatio",
                      "QuickRatio", "RevenueGrowth", "EarningsGrowth"}

    if metric_key in less_is_better:
        if value < low:
            return "LOW", f"明顯低於參考值({low})，對投資人相對有利", 10
        elif value > high:
            return "HIGH", f"高於參考值({high})，需留意高估風險", 2
        else:
            return "MID", f"介於 {low} ~ {high} 的區間", 5
    elif metric_key in more_is_better:
        if value < low:
            display_val = f"{low*100:.2f}%" if metric_key not in {'CurrentRatio','QuickRatio'} else low
            return "LOW", f"低於參考值({display_val})", 2
        elif value > high:
            display_val = f"{high*100:.2f}%" if metric_key not in {'CurrentRatio','QuickRatio'} else high
            return "HIGH", f"高於參考值({display_val})，值得肯定", 10
        else:
            display_low = f"{low*100:.2f}%" if metric_key not in {'CurrentRatio','QuickRatio'} else low
            display_high = f"{high*100:.2f}%" if metric_key not in {'CurrentRatio','QuickRatio'} else high
            return "MID", f"在 {display_low} ~ {display_high} 的合理範圍", 5
    else:
        return None, "", 0

def etf_compare_with_threshold(value, metric_key):
    thresholds = ETF_THRESHOLDS.get(metric_key, None)
    if not thresholds or value is None:
        return None, "", 0

    low, high = thresholds["low"], thresholds["high"]

    less_is_better = {"ETF_PE"}
    more_is_better = {"ETF_Yield", "totalAssets", "threeYearAverageReturn", "fiveYearAverageReturn"}

    if metric_key == "ETF_Beta3Y":
        if value < low:
            return "LOW", f"波動低於市場平均({low})", 6
        elif value > high:
            return "HIGH", f"波動高於市場平均({high})，風險也較高", 4
        else:
            return "MID", f"在 {low} ~ {high} 的波動區間", 8

    if metric_key in less_is_better:
        if value < low:
            return "LOW", f"明顯低於參考值({low})，對投資人相對有利", 10
        elif value > high:
            return "HIGH", f"高於參考值({high})，需留意高估風險", 2
        else:
            return "MID", f"介於 {low}~{high} 的區間", 5
    elif metric_key in more_is_better:
        if value < low:
            return "LOW", f"低於參考值({low})", 2
        elif value > high:
            return "HIGH", f"高於參考值({high})，表現不錯", 10
        else:
            return "MID", f"在 {low}~{high} 的區間", 5
    else:
        return None, "", 0

def technical_compare_with_threshold(indicators: dict) -> (list, int, float):
    """
    評價技術指標並返回評價結果和總分
    indicators: 包含所有技術指標的字典
    返回: (技術指標評價字串列表, 技術指標總分, 技術指標平均分)
    """
    comments = []
    total_score = 0
    valid_count = 0

    # RSI 評價
    rsi = indicators.get("RSI", None)
    if rsi is not None:
        if rsi < TECHNICAL_THRESHOLDS["RSI"]["low"]:
            comments.append(f"- RSI: {rsi:.2f}（超賣，可能反彈） (score=8)")
            total_score += 8
        elif rsi > TECHNICAL_THRESHOLDS["RSI"]["high"]:
            comments.append(f"- RSI: {rsi:.2f}（超買，可能回調） (score=6)")
            total_score += 6
        else:
            comments.append(f"- RSI: {rsi:.2f}（正常範圍） (score=7)")
            total_score += 7
        valid_count += 1
    else:
        comments.append("- RSI: 資料缺失")

    # MACD 評價
    macd = indicators.get("MACD", None)
    macd_signal = indicators.get("MACD_Signal", None)
    if macd is not None and macd_signal is not None:
        if macd > macd_signal:
            comments.append("- MACD: 正值（買入信號） (score=8)")
            total_score += 8
        else:
            comments.append("- MACD: 負值（賣出信號） (score=6)")
            total_score += 6
        valid_count += 1
    else:
        comments.append("- MACD: 資料缺失")

    # 移動平均線評價
    sma_50 = indicators.get("SMA_50", None)
    sma_200 = indicators.get("SMA_200", None)
    if sma_50 is not None and sma_200 is not None:
        if sma_50 > sma_200:
            comments.append("- SMA50 > SMA200（黃金交叉，趨勢向上） (score=9)")
            total_score += 9
        else:
            comments.append("- SMA50 < SMA200（死亡交叉，趨勢向下） (score=5)")
            total_score += 5
        valid_count += 1
    else:
        comments.append("- SMA50/SMA200: 資料缺失")

    # ADX 評價
    adx = indicators.get("ADX", None)
    if adx is not None:
        if adx < TECHNICAL_THRESHOLDS["ADX"]["low"]:
            comments.append(f"- ADX: {adx:.2f}（趨勢弱） (score=5)")
            total_score += 5
        elif adx > TECHNICAL_THRESHOLDS["ADX"]["high"]:
            comments.append(f"- ADX: {adx:.2f}（趨勢強） (score=8)")
            total_score += 8
        else:
            comments.append(f"- ADX: {adx:.2f}（趨勢中等） (score=6)")
            total_score += 6
        valid_count += 1
    else:
        comments.append("- ADX: 資料缺失")

    # 成交量趨勢評價
    volume_trend = indicators.get("Volume_Trend", None)
    if volume_trend is not None:
        if volume_trend == "increasing":
            comments.append("- 成交量趨勢: 增加中（趨勢確認） (score=7)")
            total_score += 7
        elif volume_trend == "decreasing":
            comments.append("- 成交量趨勢: 減少中（趨勢可能弱化） (score=5)")
            total_score += 5
        else:
            comments.append("- 成交量趨勢: 無明顯趨勢 (score=6)")
            total_score += 6
        valid_count += 1
    else:
        comments.append("- 成交量趨勢: 資料缺失")

    if valid_count > 0:
        avg_score = total_score / valid_count
    else:
        avg_score = 0

    return comments, total_score, avg_score

def advanced_equity_analysis(info: dict) -> dict:
    """
    回傳包含基本面與技術面分析的 dict。
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")

    # 基本面指標
    pe = info.get("trailingPE", None)
    pb = info.get("priceToBook", None)
    beta = info.get("beta", None)
    roe = info.get("returnOnEquity", None)
    dividend_yield = info.get("dividendYield", None)
    if dividend_yield is not None:
        dividend_yield *= 100

    peg = info.get("trailingPegRatio", None) or info.get("pegRatio", None)
    operating_margin = info.get("operatingMargins", None)
    d2e = info.get("debtToEquity", None)
    current_ratio = info.get("currentRatio", None)
    quick_ratio = info.get("quickRatio", None)
    revenue_growth = info.get("revenueGrowth", None)
    earnings_growth = info.get("earningsQuarterlyGrowth", None) or info.get("earningsGrowth", None)

    metrics = {
        "PE": pe,
        "PB": pb,
        "Beta": beta,
        "ROE": roe,
        "DividendYield": dividend_yield,
        "PEG": peg,
        "OperatingMargin": operating_margin,
        "DebtToEquity": d2e,
        "CurrentRatio": current_ratio,
        "QuickRatio": quick_ratio,
        "RevenueGrowth": revenue_growth,
        "EarningsGrowth": earnings_growth
    }

    # 基本面分析
    fundamental_comments = [f"**[基本面分析] {symbol} / {short_name}**"]
    fundamental_total_score = 0
    fundamental_valid_count = 0

    for k, v in metrics.items():
        if v is None:
            fundamental_comments.append(f"- {k}: 資料缺失")
            continue

        level, note, score = compare_with_threshold(v, k)

        if k in ["OperatingMargin", "RevenueGrowth", "EarningsGrowth", "ROE"]:
            display_value = f"{v*100:.2f}%"
        elif k == "DividendYield":
            display_value = f"{v:.2f}%"
        else:
            display_value = f"{v:.2f}"

        fundamental_comments.append(f"- {k}: {display_value} => {note} (score={score})")

        fundamental_total_score += score
        fundamental_valid_count += 1

    # 計算基本面平均分數與評級
    if fundamental_valid_count > 0:
        fundamental_avg_score = fundamental_total_score / fundamental_valid_count
        if fundamental_avg_score >= 8:
            fundamental_rating = "A"
        elif fundamental_avg_score >= 5:
            fundamental_rating = "B"
        elif fundamental_avg_score >= 3:
            fundamental_rating = "C"
        else:
            fundamental_rating = "D"
        fundamental_comments.append(f"\n【基本面評級】平均分數 {fundamental_avg_score:.1f} → 等級 {fundamental_rating}")
    else:
        fundamental_avg_score = 0
        fundamental_rating = "無法評級"
        fundamental_comments.append("\n【基本面評級】指標不足，無法評級")

    # 技術面分析
    technical_comments = []
    technical_total_score = 0
    technical_valid_count = 0

    try:
        ticker_obj = yf.Ticker(symbol)
        end_date = datetime.datetime.today()
        start_date = end_date - datetime.timedelta(days=365)  # 過去一年的數據
        hist = ticker_obj.history(start=start_date, end=end_date)

        if hist.empty:
            technical_comments.append("- 技術指標: 無歷史價格數據，無法計算技術指標")
        else:
            # 計算技術指標
            indicators = {}
            indicators["RSI"] = momentum.RSIIndicator(hist['Close']).rsi()[-1]
            indicators["MACD"] = trend.MACD(hist['Close']).macd()[-1]
            indicators["MACD_Signal"] = trend.MACD(hist['Close']).macd_signal()[-1]
            indicators["SMA_50"] = trend.SMAIndicator(hist['Close'], window=50).sma_indicator()[-1]
            indicators["SMA_200"] = trend.SMAIndicator(hist['Close'], window=200).sma_indicator()[-1]
            indicators["ADX"] = trend.ADXIndicator(hist['High'], hist['Low'], hist['Close']).adx()[-1]
            # 成交量趨勢判斷（簡單版：最近20天平均大於前20天）
            recent_vol = hist['Volume'][-20:].mean()
            previous_vol = hist['Volume'][-40:-20].mean()
            if recent_vol > previous_vol:
                indicators["Volume_Trend"] = "increasing"
            elif recent_vol < previous_vol:
                indicators["Volume_Trend"] = "decreasing"
            else:
                indicators["Volume_Trend"] = "stable"

            # 評價技術指標
            tech_comments, tech_total_score, tech_avg_score = technical_compare_with_threshold(indicators)
            technical_comments.extend(tech_comments)
    except Exception as e:
        technical_comments.append(f"- 技術指標: 計算失敗 ({e})")

    # 計算技術面平均分數與評級
    if technical_comments and not all("資料缺失" in comment or "無歷史價格數據" in comment for comment in technical_comments):
        # 計算有效的技術分數和計數
        tech_scores = [int(s.split("(score=")[-1].rstrip(")")) for s in technical_comments if "score=" in s]
        if tech_scores:
            technical_total_score = sum(tech_scores)
            technical_valid_count = len(tech_scores)
            technical_avg_score = technical_total_score / technical_valid_count

            if technical_avg_score >= 8:
                technical_rating = "A"
            elif technical_avg_score >= 5:
                technical_rating = "B"
            elif technical_avg_score >= 3:
                technical_rating = "C"
            else:
                technical_rating = "D"

            technical_comments.append(f"\n【技術面評級】平均分數 {technical_avg_score:.1f} → 等級 {technical_rating}")
        else:
            technical_avg_score = 0
            technical_rating = "無法評級"
            technical_comments.append("\n【技術面評級】指標不足，無法評級")
    else:
        technical_avg_score = 0
        technical_rating = "無法評級"
        technical_comments.append("\n【技術面評級】指標不足，無法評級")

    # 最終分析說明
    analysis_text = "\n".join(fundamental_comments + technical_comments)

    return {
        "symbol": symbol,
        "shortName": short_name,
        # 基本面分數與評級
        "fundamental_total_score": fundamental_total_score,
        "fundamental_avg_score": fundamental_avg_score,
        "fundamental_rating": fundamental_rating,
        # 技術面分數與評級
        "technical_total_score": technical_total_score,
        "technical_avg_score": technical_avg_score,
        "technical_rating": technical_rating,
        "analysis_text": analysis_text
    }

def advanced_etf_analysis(info: dict) -> dict:
    """
    回傳包含基本面與技術面分析的 dict（ETF 特有）。
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")

    # ETF 基本面指標
    etf_pe = info.get("trailingPE", None)
    etf_yield = info.get("yield", None)
    if etf_yield is not None:
        etf_yield *= 100

    total_assets = info.get("totalAssets", None)
    beta_3y = info.get("beta3Year", None)
    three_y_avg_return = info.get("threeYearAverageReturn", None)
    five_y_avg_return = info.get("fiveYearAverageReturn", None)

    etf_metrics = {
        "ETF_PE": etf_pe,
        "ETF_Yield": etf_yield,
        "totalAssets": total_assets,
        "ETF_Beta3Y": beta_3y,
        "threeYearAverageReturn": three_y_avg_return,
        "fiveYearAverageReturn": five_y_avg_return
    }

    # 基本面分析
    fundamental_comments = [f"**[基本面分析] {symbol} / {short_name}**"]
    fundamental_total_score = 0
    fundamental_valid_count = 0

    for k, v in etf_metrics.items():
        if v is None:
            fundamental_comments.append(f"- {k}: 資料缺失")
            continue

        level, note, score = etf_compare_with_threshold(v, k)

        if k in ["ETF_Yield"]:
            display_value = f"{v:.2f}%"
        elif k in ["threeYearAverageReturn", "fiveYearAverageReturn"]:
            display_value = f"{v*100:.2f}%"
        elif k == "totalAssets":
            display_value = f"{v:,.0f}"
        else:
            display_value = f"{v:.2f}"

        fundamental_comments.append(f"- {k}: {display_value} => {note} (score={score})")

        fundamental_total_score += score
        fundamental_valid_count += 1

    # 計算基本面平均分數與評級
    if fundamental_valid_count > 0:
        fundamental_avg_score = fundamental_total_score / fundamental_valid_count
        if fundamental_avg_score >= 8:
            fundamental_rating = "A"
        elif fundamental_avg_score >= 5:
            fundamental_rating = "B"
        elif fundamental_avg_score >= 3:
            fundamental_rating = "C"
        else:
            fundamental_rating = "D"
        fundamental_comments.append(f"\n【基本面評級】平均分數 {fundamental_avg_score:.1f} → 等級 {fundamental_rating}")
    else:
        fundamental_avg_score = 0
        fundamental_rating = "無法評級"
        fundamental_comments.append("\n【基本面評級】指標不足，無法評級")

    # 技術面分析
    technical_comments = []
    technical_total_score = 0
    technical_valid_count = 0

    try:
        ticker_obj = yf.Ticker(symbol)
        end_date = datetime.datetime.today()
        start_date = end_date - datetime.timedelta(days=365)  # 過去一年的數據
        hist = ticker_obj.history(start=start_date, end=end_date)

        if hist.empty:
            technical_comments.append("- 技術指標: 無歷史價格數據，無法計算技術指標")
        else:
            # 計算技術指標
            indicators = {}
            indicators["RSI"] = momentum.RSIIndicator(hist['Close']).rsi()[-1]
            indicators["MACD"] = trend.MACD(hist['Close']).macd()[-1]
            indicators["MACD_Signal"] = trend.MACD(hist['Close']).macd_signal()[-1]
            indicators["SMA_50"] = trend.SMAIndicator(hist['Close'], window=50).sma_indicator()[-1]
            indicators["SMA_200"] = trend.SMAIndicator(hist['Close'], window=200).sma_indicator()[-1]
            indicators["ADX"] = trend.ADXIndicator(hist['High'], hist['Low'], hist['Close']).adx()[-1]
            # 成交量趨勢判斷（簡單版：最近20天平均大於前20天）
            recent_vol = hist['Volume'][-20:].mean()
            previous_vol = hist['Volume'][-40:-20].mean()
            if recent_vol > previous_vol:
                indicators["Volume_Trend"] = "increasing"
            elif recent_vol < previous_vol:
                indicators["Volume_Trend"] = "decreasing"
            else:
                indicators["Volume_Trend"] = "stable"

            # 評價技術指標
            tech_comments, tech_total_score, tech_avg_score = technical_compare_with_threshold(indicators)
            technical_comments.extend(tech_comments)
    except Exception as e:
        technical_comments.append(f"- 技術指標: 計算失敗 ({e})")

    # 計算技術面平均分數與評級
    if technical_comments and not all("資料缺失" in comment or "無歷史價格數據" in comment for comment in technical_comments):
        # 計算有效的技術分數和計數
        tech_scores = [int(s.split("(score=")[-1].rstrip(")")) for s in technical_comments if "score=" in s]
        if tech_scores:
            technical_total_score = sum(tech_scores)
            technical_valid_count = len(tech_scores)
            technical_avg_score = technical_total_score / technical_valid_count

            if technical_avg_score >= 8:
                technical_rating = "A"
            elif technical_avg_score >= 5:
                technical_rating = "B"
            elif technical_avg_score >= 3:
                technical_rating = "C"
            else:
                technical_rating = "D"

            technical_comments.append(f"\n【技術面評級】平均分數 {technical_avg_score:.1f} → 等級 {technical_rating}")
        else:
            technical_avg_score = 0
            technical_rating = "無法評級"
            technical_comments.append("\n【技術面評級】指標不足，無法評級")
    else:
        technical_avg_score = 0
        technical_rating = "無法評級"
        technical_comments.append("\n【技術面評級】指標不足，無法評級")

    # 最終分析說明
    analysis_text = "\n".join(fundamental_comments + technical_comments)

    return {
        "symbol": symbol,
        "shortName": short_name,
        # 基本面分數與評級
        "fundamental_total_score": fundamental_total_score,
        "fundamental_avg_score": fundamental_avg_score,
        "fundamental_rating": fundamental_rating,
        # 技術面分數與評級
        "technical_total_score": technical_total_score,
        "technical_avg_score": technical_avg_score,
        "technical_rating": technical_rating,
        "analysis_text": analysis_text
    }

def analyze_ticker(ticker_symbol: str) -> dict:
    """
    根據 quoteType 決定回傳哪種 dict 分析結果。
    """
    ticker_obj = yf.Ticker(ticker_symbol)
    info = ticker_obj.info

    quote_type = info.get("quoteType", "")
    if quote_type == "EQUITY":
        return advanced_equity_analysis(info)
    elif quote_type == "ETF":
        return advanced_etf_analysis(info)
    else:
        # 若不是 EQUITY / ETF，就回傳一個簡易結構
        return {
            "symbol": ticker_symbol,
            "shortName": info.get("shortName", ""),
            "fundamental_total_score": 0,
            "fundamental_avg_score": 0,
            "fundamental_rating": "無法評級",
            "technical_total_score": 0,
            "technical_avg_score": 0,
            "technical_rating": "無法評級",
            "analysis_text": f"{ticker_symbol} quoteType={quote_type}，暫不支援進階分析。"
        }

if __name__ == "__main__":
    watch_list = [
        "2330.TW",
        "2317.TW",
        "2454.TW",
        "2881.TW",
        "2308.TW",
        "2382.TW",
        "2882.TW",
        "2412.TW",
        "2891.TW",
        "3711.TW",
        "2886.TW",
        "2303.TW",
        "6669.TW",
        "2603.TW",
        "1216.TW",
        "2357.TW",
        "2885.TW",
        "2345.TW",
        "2884.TW",
        "3045.TW",
        "2892.TW",
        "5880.TW",
        "2880.TW",
        "3008.TW",
        "2207.TW",
        "6505.TW",
        "4904.TW",
        "2002.TW",
        "3034.TW",
        "3231.TW",
        "2395.TW",
        "2379.TW",
        "2890.TW",
        "2883.TW",
        "2327.TW",
        "2912.TW",
        "2609.TW",
        "3661.TW",
        "4938.TW",
        "2618.TW",
        "1101.TW",
        "3017.TW",
        "1303.TW",
        "2301.TW",
        "2615.TW",
        "1301.TW",
        "2887.TW",
        "3533.TW",
        "3653.TW",
        "3037.TW",
        "0050.TW",
        "0056.TW",
        "00878.TW",
        "00919.TW",
        "00679B.TWO",
        "00937B.TWO",
        "00687B.TWO",
        "00929.TW",
        "006208.TW",
        "00751B.TWO",
        "00720B.TWO",
        "00940.TW",
        "00725B.TWO",
        "00772B.TWO",
        "00713.TW",
        "00761B.TWO",
        "00773B.TWO",
        "00724B.TWO",
        "00746B.TWO",
        "00933B.TWO",
    ]

    # 1) 收集分析結果 (dict)
    resultList = []
    for tkr in watch_list:
        print(f"分析 {tkr} 中...")
        result = analyze_ticker(tkr)
        resultList.append(result)

    # 2) 依「平均分數」(基礎分數與技術分數的加權平均或其他邏輯) 由高至低排序
    # 這裡假設我們按技術面與基本面平均分數的簡單加權平均進行排序
    for result in resultList:
        # 簡單加權平均，可以根據需求調整權重
        result["overall_avg_score"] = (result["fundamental_avg_score"] + result["technical_avg_score"]) / 2
        # 簡單加權總分
        result["overall_total_score"] = result["fundamental_total_score"] + result["technical_total_score"]
        # 簡單綜合評級
        if result["overall_avg_score"] >= 8:
            result["overall_rating"] = "A"
        elif result["overall_avg_score"] >= 5:
            result["overall_rating"] = "B"
        elif result["overall_avg_score"] >= 3:
            result["overall_rating"] = "C"
        else:
            result["overall_rating"] = "D"

    # 依「overall_avg_score」排序
    sorted_results = sorted(resultList, key=lambda x: x.get("overall_avg_score", 0), reverse=True)

    # 3) 轉換成 DataFrame 格式，方便儲存為 CSV
    df = pd.DataFrame(sorted_results)

    # 4) 儲存為 CSV 檔案
    output_file = "stock_analysis_results.csv"
    df.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"分析結果已儲存至 {output_file}")
