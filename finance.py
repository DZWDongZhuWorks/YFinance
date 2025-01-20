import yfinance as yf
import pandas as pd

# 這裡我們整合先前的一些指標 + 新增指標的閾值
THRESHOLDS = {
    # PE, PB, Beta, ROE, DividendYield 與前例相同，不再贅述
    "PE": {"low": 10, "high": 25},
    "PB": {"low": 1,  "high": 5},
    "Beta": {"low": 0.8, "high": 1.2},
    "ROE": {"low": 0.1, "high": 0.2},  # 10% ~ 20%
    "DividendYield": {"low": 2, "high": 6},  
    
    # 以下為新增指標
    "PEG": {"low": 1, "high": 2},                # PEG <1 或 <1.5 常被視為合理區或偏低
    "OperatingMargin": {"low": 0.1, "high": 0.3},# 10% ~ 30%
    "DebtToEquity": {"low": 50, "high": 100},    # <50% 穩健, >100% 負債偏高
    "CurrentRatio": {"low": 1, "high": 2},       # >1 一般較健康
    "QuickRatio": {"low": 1, "high": 2},         # >1 一般較健康
    "RevenueGrowth": {"low": 0.05, "high": 0.2}, # 5% ~ 20% 
    "EarningsGrowth": {"low": 0.05, "high": 0.2} # 5% ~ 20%
}
# ETF用的閾值（示範）：
ETF_THRESHOLDS = {
    "ETF_PE":     {"low": 10, "high": 25},   # ETF P/E
    "ETF_Yield":  {"low": 2,  "high": 6},    # ETF殖利率(%) 2~6%屬於一般區間
    "ETF_Beta3Y": {"low": 0.8, "high": 1.2}, # 3年Beta
    "threeYearAverageReturn": {"low": 0.03, "high": 0.1},  # 3%~10% 年化
    "fiveYearAverageReturn":  {"low": 0.03, "high": 0.1},  # 3%~10% 年化
    # totalAssets可自行分級，例如 <100 億、100~500 億、>500 億
    # 這邊用 "numeric" 分段示範
    "totalAssets": {"low": 1e11, "high": 5e11},  
    # 若有 expenseRatio, trackingError 等可再加
}

def compare_with_threshold(value, metric_key):
    """
    根據 THRESHOLDS 做閾值比較，傳回 (level, note, score)。
    level: "LOW", "MID", or "HIGH"
    note: 額外敘述，例如 '高於一般參考值(2)'
    score: 根據 level 給予一個分數(自行設定)，以方便後續彙整計算。
    """
    thresholds = THRESHOLDS.get(metric_key, None)
    if not thresholds or value is None:
        return None, "", 0
    
    low, high = thresholds["low"], thresholds["high"]
    # 這裡 score 的給法僅示範：LOW ~ HIGH 給不同分數
    # 實務可根據投資人偏好或指標特性 (如越小越好、越大越好) 作不同設計
    # 以 PEG 為例，越小越好；以 ProfitMargin 為例，越大越好。請依實際情況調整邏輯
    
    # 先分辨指標是 "越低越好" 還是 "越高越好"
    # PEG, PE, DebtToEquity, Beta, P/B 等通常「越低越好」
    # ProfitMargin, ROE, CurrentRatio, RevenueGrowth, OperatingMargin 等通常「越高越好」
    
    # 我們先定義一個方向性
    less_is_better = {"PE", "PB", "PEG", "Beta", "DebtToEquity"}
    more_is_better = {"ROE", "OperatingMargin", "DividendYield", "CurrentRatio",
                      "QuickRatio", "RevenueGrowth", "EarningsGrowth"}
    
    if metric_key in less_is_better:
        # value < low → 好 (LOW = "估值偏低")
        if value < low:
            return "LOW", f"明顯低於參考值({low})，對投資人相對有利", 10
        elif value > high:
            return "HIGH", f"高於參考值({high})，需留意高估風險", 2
        else:
            return "MID", f"介於 {low} ~ {high} 的區間", 5
    elif metric_key in more_is_better:
        # value < low → 不好
        if value < low:
            return "LOW", f"低於參考值({low*100 if metric_key not in {'CurrentRatio','QuickRatio'} else low})", 2
        elif value > high:
            return "HIGH", f"高於參考值({high*100 if metric_key not in {'CurrentRatio','QuickRatio'} else high})，值得肯定", 10
        else:
            return "MID", f"在 {low*100 if metric_key not in {'CurrentRatio','QuickRatio'} else low} ~ {high*100 if metric_key not in {'CurrentRatio','QuickRatio'} else high} 的合理範圍", 5
    else:
        return None, "", 0

def advanced_equity_analysis(info: dict) -> str:
    """
    分析個股(EQUITY)的進階指標，並給出更完整評價與最終評級 (A/B/C...等)。
    根據 info 中取得：
    - PEG (info.get("trailingPegRatio") 或 "pegRatio")
    - OperatingMargins, DebtToEquity, CurrentRatio, QuickRatio
    - RevenueGrowth, EarningsGrowth
    - 其他前面已有的 P/E, P/B, ROE, DividendYield ...
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")
    
    # 收集指標
    pe = info.get("trailingPE", None)
    pb = info.get("priceToBook", None)
    beta = info.get("beta", None)
    roe = info.get("returnOnEquity", None)       # already in ratio form (e.g. 0.28 => 28%)
    dividend_yield = info.get("dividendYield", None)
    if dividend_yield is not None:
        dividend_yield *= 100
    
    # 進階指標
    peg = info.get("trailingPegRatio", None) or info.get("pegRatio", None)  # yfinance常見鍵有時是 pegRatio
    operating_margin = info.get("operatingMargins", None)  # e.g. 0.30 => 30%
    d2e = info.get("debtToEquity", None)                    # 若 info 有提供
    current_ratio = info.get("currentRatio", None)
    quick_ratio = info.get("quickRatio", None)
    revenue_growth = info.get("revenueGrowth", None)        # e.g. 0.39 => 39%
    earnings_growth = info.get("earningsQuarterlyGrowth", None) or info.get("earningsGrowth", None)
    
    # 方便後續統一處理
    metrics = {
        "PE": pe,
        "PB": pb,
        "Beta": beta,
        "ROE": roe,  # 0.28 => 28%
        "DividendYield": dividend_yield, 
        "PEG": peg,
        "OperatingMargin": operating_margin, # 0.30 => 30%
        "DebtToEquity": d2e,
        "CurrentRatio": current_ratio,
        "QuickRatio": quick_ratio,
        "RevenueGrowth": revenue_growth,   # 0.39 => 39%
        "EarningsGrowth": earnings_growth  # 0.54 => 54%
    }
    
    comments = [f"**[進階個股分析] {symbol} / {short_name}**"]
    total_score = 0
    valid_count = 0
    
    # 一一對指標做比較
    for k, v in metrics.items():
        if v is None:
            comments.append(f"- {k}: 資料缺失")
            continue
        # 比較
        level, note, score = compare_with_threshold(v, k)
        
        # 格式化輸出：若是百分比類型的指標就轉成%
        # Beta, DebtToEquity, CurrentRatio, QuickRatio等 不需 *100
        # 但 OperatingMargin, RevenueGrowth, EarningsGrowth, ROE 需要 *100 做輸出
        if k in ["OperatingMargin", "RevenueGrowth", "EarningsGrowth", "ROE"]:
            display_value = f"{v*100:.2f}%"
        elif k == "DividendYield":
            display_value = f"{v:.2f}%"
        else:
            display_value = f"{v:.2f}"
        
        # 附註 + Score
        comments.append(f"- {k}: {display_value} => {note} (score={score})")
        
        total_score += score
        valid_count += 1
    
    # 綜合評級 (Demo)
    # 這裡假設平均分數 >=8 給 A, 5-8 給 B, 2-5 給 C, <2 給 D，可依需求微調
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
        comments.append("\n【最終評級】指標不足，無法評級")
    
    return "\n".join(comments)


def etf_compare_with_threshold(value, metric_key):
    """
    與 ETF_THRESHOLDS 做閾值比較，回傳 (level, note, score)。
    level: "LOW", "MID", or "HIGH"
    note: 額外敘述
    score: 用於最終計算評級的分數
    """
    thresholds = ETF_THRESHOLDS.get(metric_key, None)
    if not thresholds or value is None:
        return None, "", 0
    
    low, high = thresholds["low"], thresholds["high"]
    
    # 判斷指標是「越低越好」還是「越高越好」
    # ETF_PE: 越低越好
    # ETF_Yield: 越高越好
    # totalAssets: 越高越好
    # threeYearAverageReturn, fiveYearAverageReturn: 越高越好
    # Beta3Y: ~1.0 為與市場波動相近，<0.8 偏穩健，>1.2 較高波動
    
    less_is_better = {"ETF_PE"}
    more_is_better = {"ETF_Yield", "totalAssets", "threeYearAverageReturn", "fiveYearAverageReturn"}
    
    # Beta 為特殊：中間最好，太低或太高都可能有不同風險意義
    if metric_key == "ETF_Beta3Y":
        # <0.8 => LOW(波動小), >1.2 => HIGH(波動大)
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

def advanced_etf_analysis(info: dict) -> str:
    """
    針對 quoteType='ETF' 的標的，進行進階分析。
    收集以下指標:
    - trailingPE => ETF_PE
    - yield => ETF_Yield
    - totalAssets
    - beta3Year => ETF_Beta3Y (yfinance 可能是 'beta3Year')
    - threeYearAverageReturn, fiveYearAverageReturn
    """
    symbol = info.get("symbol", "UNKNOWN")
    short_name = info.get("shortName", "")
    
    # 取得
    etf_pe   = info.get("trailingPE", None)
    etf_yield= info.get("yield", None)
    if etf_yield is not None:
        etf_yield *= 100  # 轉成百分比
    
    total_assets = info.get("totalAssets", None)
    
    beta_3y = info.get("beta3Year", None)
    
    three_y_avg_return = info.get("threeYearAverageReturn", None) # 0.12 => 12%
    five_y_avg_return  = info.get("fiveYearAverageReturn", None)  # 0.18 => 18%
    
    # 假設可能還有 expenseRatio, trackingError (此處yfinance不一定有)
    # expense_ratio = info.get("annualReportExpenseRatio", None)
    # tracking_error = ... # 可能無法直接抓到
    
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
        
        # 數值輸出格式 (若是百分比就 x100)
        if k in ["ETF_Yield"]:
            display_value = f"{v:.2f}%"
        elif k in ["threeYearAverageReturn", "fiveYearAverageReturn"]:
            display_value = f"{v*100:.2f}%"
        elif k == "ETF_PE":
            display_value = f"{v:.2f}"
        elif k == "ETF_Beta3Y":
            display_value = f"{v:.2f}"
        elif k == "totalAssets":
            display_value = f"{v:,.0f}"  # 例如 "435,205,963,776"
        else:
            display_value = str(v)
        
        comments.append(f"- {k}: {display_value} => {note} (score={score})")
        total_score += score
        valid_count += 1
    
    # 簡單計算平均分數 => 評級
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
        comments.append("\n【最終評級】資料不足，無法評級")
    
    # 額外補充: 若看出 ETF 是 0050 或 0056，可再加描述
    if "0050" in symbol:
        comments.append("→ 追蹤台灣前 50 大權值股，大盤連動度高。")
    elif "0056" in symbol:
        comments.append("→ 高股息策略，殖利率偏高但股價波動與大盤略有差異。")
    
    # 綜合建議 (範例)
    final_advice = "ETF 可作為分散投資的核心或衛星標的，建議定期定額並留意成分股調整。"
    comments.append(f"\n【綜合建議】{final_advice}")
    
    return "\n".join(comments)

def analyze_ticker(ticker_symbol: str) -> str:
    """
    根據 quoteType 決定調用 advanced_equity_analysis or advanced_etf_analysis。
    """
    ticker_obj = yf.Ticker(ticker_symbol)
    info = ticker_obj.info
    
    quote_type = info.get("quoteType", "")
    
    if quote_type == "EQUITY":
        # 這裡假設你有先前寫好的 advanced_equity_analysis(info)
        from typing import Callable
        # 這裡示範直接 inline import (當然你也可以把 advanced_equity_analysis 的程式碼放在同一份檔案裡)
        return advanced_equity_analysis(info)
    elif quote_type == "ETF":
        return advanced_etf_analysis(info)
    else:
        return f"{ticker_symbol} quoteType={quote_type}，暫不支援進階分析。"

        
if __name__ == "__main__":
    watch_list = ["2330.TW", "0050.TW", "0056.TW"]
    for tkr in watch_list:
        result = analyze_ticker(tkr)
        print(result)
        print("-----")
