import yfinance as yf
import pandas as pd
import csv

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

ETF_THRESHOLDS = {
    "ETF_PE":     {"low": 10, "high": 25},
    "ETF_Yield":  {"low": 2,  "high": 6},
    "ETF_Beta3Y": {"low": 0.8, "high": 1.2},
    "threeYearAverageReturn": {"low": 0.03, "high": 0.1},
    "fiveYearAverageReturn":  {"low": 0.03, "high": 0.1},
    "totalAssets": {"low": 1e11, "high": 5e11},
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
            return "LOW", f"低於參考值({low*100 if metric_key not in {'CurrentRatio','QuickRatio'} else low})", 2
        elif value > high:
            return "HIGH", f"高於參考值({high*100 if metric_key not in {'CurrentRatio','QuickRatio'} else high})，值得肯定", 10
        else:
            return "MID", f"在 {low*100 if metric_key not in {'CurrentRatio','QuickRatio'} else low} ~ {high*100 if metric_key not in {'CurrentRatio','QuickRatio'} else high} 的合理範圍", 5
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

def advanced_equity_analysis(info: dict) -> dict:
    """
    改為回傳 dict，包含：
    - symbol, shortName
    - total_score, avg_score, rating
    - analysis_text (最後完整的分析說明)
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")
    
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
    
    comments = [f"**[進階個股分析] {symbol} / {short_name}**"]
    total_score = 0
    valid_count = 0
    
    for k, v in metrics.items():
        if v is None:
            comments.append(f"- {k}: 資料缺失")
            continue
        
        level, note, score = compare_with_threshold(v, k)
        
        if k in ["OperatingMargin", "RevenueGrowth", "EarningsGrowth", "ROE"]:
            display_value = f"{v*100:.2f}%"
        elif k == "DividendYield":
            display_value = f"{v:.2f}%"
        else:
            display_value = f"{v:.2f}"
        
        comments.append(f"- {k}: {display_value} => {note} (score={score})")
        
        total_score += score
        valid_count += 1
    
    if valid_count > 0:
        avg_score = total_score / valid_count
        if avg_score >= 8:
            rating = "A"
        elif avg_score >= 5:
            rating = "B"
        elif avg_score >= 3:
            rating = "C"
        else:
            rating = "D"
        comments.append(f"\n【最終評級】平均分數 {avg_score:.1f} → 等級 {rating}")
    else:
        avg_score = 0
        rating = "無法評級"
        comments.append("\n【最終評級】指標不足，無法評級")
    
    analysis_text = "\n".join(comments)
    
    return {
        "symbol": symbol,
        "shortName": short_name,
        "total_score": total_score,
        "avg_score": avg_score,
        "rating": rating,
        "analysis_text": analysis_text
    }

def advanced_etf_analysis(info: dict) -> dict:
    """
    與 advanced_equity_analysis 類似，回傳 dict。
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")
    
    etf_pe   = info.get("trailingPE", None)
    etf_yield= info.get("yield", None)
    if etf_yield is not None:
        etf_yield *= 100
    
    total_assets = info.get("totalAssets", None)
    beta_3y = info.get("beta3Year", None)
    three_y_avg_return = info.get("threeYearAverageReturn", None)
    five_y_avg_return  = info.get("fiveYearAverageReturn", None)
    
    metrics = {
        "ETF_PE": etf_pe,
        "ETF_Yield": etf_yield,
        "totalAssets": total_assets,
        "ETF_Beta3Y": beta_3y,
        "threeYearAverageReturn": three_y_avg_return,
        "fiveYearAverageReturn": five_y_avg_return
    }
    
    comments = [f"**[進階ETF分析] {symbol} / {short_name}**"]
    total_score = 0
    valid_count = 0
    
    for k, v in metrics.items():
        if v is None:
            comments.append(f"- {k}: 資料缺失")
            continue
        
        level, note, score = etf_compare_with_threshold(v, k)
        
        if k in ["ETF_Yield"]:
            display_value = f"{v:.2f}%"
        elif k in ["threeYearAverageReturn", "fiveYearAverageReturn"]:
            display_value = f"{v*100:.2f}%"
        elif k == "ETF_Beta3Y":
            display_value = f"{v:.2f}"
        elif k == "totalAssets":
            display_value = f"{v:,.0f}"
        else:
            display_value = f"{v:.2f}"
        
        comments.append(f"- {k}: {display_value} => {note} (score={score})")
        total_score += score
        valid_count += 1
    
    if valid_count > 0:
        avg_score = total_score / valid_count
        if avg_score >= 8:
            rating = "A"
        elif avg_score >= 5:
            rating = "B"
        elif avg_score >= 3:
            rating = "C"
        else:
            rating = "D"
        
        comments.append(f"\n【最終評級】平均分數 {avg_score:.1f} → 等級 {rating}")
    else:
        avg_score = 0
        rating = "無法評級"
        comments.append("\n【最終評級】資料不足，無法評級")
    
    if "0050" in symbol:
        comments.append("→ 追蹤台灣前 50 大權值股，大盤連動度高。")
    elif "0056" in symbol:
        comments.append("→ 高股息策略，殖利率偏高但股價波動與大盤略有差異。")
    
    final_advice = "ETF 可作為分散投資的核心或衛星標的，建議定期定額並留意成分股調整。"
    comments.append(f"\n【綜合建議】{final_advice}")
    
    analysis_text = "\n".join(comments)
    return {
        "symbol": symbol,
        "shortName": short_name,
        "total_score": total_score,
        "avg_score": avg_score,
        "rating": rating,
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
            "total_score": 0,
            "avg_score": 0,
            "rating": "不支援",
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
        result = analyze_ticker(tkr)
        resultList.append(result)
    
    # 2) 依「平均分數」(或總分)由高至低排序
    sorted_results = sorted(resultList, key=lambda x: x["avg_score"], reverse=True)
    
    # 3) 轉換成 DataFrame 格式，方便儲存為 CSV
    df = pd.DataFrame(sorted_results)
    
    # 4) 儲存為 CSV 檔案
    output_file = "stock_analysis_results.csv"
    df.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"分析結果已儲存至 {output_file}")
