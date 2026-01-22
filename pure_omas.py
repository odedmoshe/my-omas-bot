import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import os
from colorama import Fore, Style, init

# אתחול צבעים לטרמינל
init(autoreset=True)

# ============================================================================
# 1. CONFIGURATION & PARAMETERS
# ============================================================================

class Config:
    # Strategy Params (Pure OMAS Logic)
    MA_LENGTH = 150
    SLOPE_THRESHOLD = 0.01        # סף מינימלי בדולרים (לסינון ראשוני בלבד)
    ENTRY_BUFFER = 1.01           # כניסה ב-1% מעל הממוצע
    EXIT_BUFFER = 0.99            # יציאה ב-1% מתחת לממוצע
    HARD_STOP_PCT = 0.15          # הגנת קטסטרופה: יציאה בהפסד 15% מהכניסה
    
    # Portfolio Params
    INITIAL_CAPITAL = 100000.0
    MAX_POSITIONS = 20
    POSITION_SIZE_PCT = 0.05      # 5% לכל פוזיציה
    
    # Files
    EXCEL_FILE = "portfolio_log.xlsx"

# ============================================================================
# 2. DATA MANAGER
# ============================================================================

def get_sp500_tickers():
    # רשימה מייצגת לדמו. בשימוש אמיתי ניתן להרחיב לכל ה-500
    return [
        'AAPL', 'MSFT', 'AMZN', 'NVDA', 'GOOGL', 'META', 'TSLA', 'BRK-B', 'UNH', 'XOM',
        'JNJ', 'JPM', 'V', 'PG', 'MA', 'HD', 'CVX', 'MRK', 'ABBV', 'LLY',
        'PEP', 'COST', 'KO', 'AVGO', 'WMT', 'MCD', 'CSCO', 'TMO', 'ACN', 'ABT',
        'DHR', 'NEE', 'LIN', 'ADBE', 'CRM', 'TXN', 'NKE', 'PM', 'RTX', 'ORCL',
        'AMD', 'INTC', 'NFLX', 'UPS', 'QCOM', 'BA', 'HON', 'UNP', 'IBM', 'CAT'
    ]

def download_data(tickers):
    print(f"{Fore.CYAN}Downloading data for {len(tickers)} stocks...{Style.RESET_ALL}")
    start_date = datetime.now() - timedelta(days=300) # מספיק היסטוריה ל-MA150
    data = yf.download(tickers, start=start_date, progress=False, group_by='ticker', auto_adjust=True)
    return data

# ============================================================================
# 3. SIGNAL GENERATOR (THE MATH)
# ============================================================================

def calculate_indicators(df):
    if len(df) < 160: return None
    
    data = df.copy()
    
    # 1. The Trend (MA150)
    data['MA150'] = data['Close'].rolling(window=Config.MA_LENGTH).mean()
    
    # 2. The Slope (MA150 - SMA(MA150, 5))
    data['MA150_Smooth'] = data['MA150'].rolling(window=5).mean()
    data['Slope'] = data['MA150'] - data['MA150_Smooth']
    
    # 3. Normalized Slope (For Ranking) -> Slope / Price
    data['Slope_Norm'] = (data['Slope'] / data['Close']) * 10000
    
    # 4. Entry/Exit Levels
    data['Entry_Threshold'] = data['MA150'] * Config.ENTRY_BUFFER
    data['Exit_Threshold'] = data['MA150'] * Config.EXIT_BUFFER
    
    return data.iloc[-1] # מחזיר רק את השורה האחרונה (היום)

# ============================================================================
# 4. RANKING SYSTEM (PURE OMAS)
# ============================================================================

def rank_candidates(candidates_list):
    """
    מדרג את המועמדות לפי הנוסחה המשולבת:
    Score = 70% Trend Strength - 30% Extension
    """
    if not candidates_list: return []
    
    df = pd.DataFrame(candidates_list)
    
    # חישוב Extension (כמה המניה ברחה מהכניסה)
    # יחס בין מחיר נוכחי למחיר הכניסה האידיאלי
    df['Extension'] = df['Close'] / df['Entry_Threshold']
    
    # נרמול בין 0 ל-100
    # Trend Strength (Slope_Norm): גבוה = טוב
    min_slope = df['Slope_Norm'].min()
    max_slope = df['Slope_Norm'].max()
    if max_slope != min_slope:
        df['Score_Trend'] = (df['Slope_Norm'] - min_slope) / (max_slope - min_slope)
    else:
        df['Score_Trend'] = 0.5

    # Extension: נמוך = טוב (קרוב ל-1.0)
    min_ext = df['Extension'].min()
    max_ext = df['Extension'].max()
    if max_ext != min_ext:
        # הופכים את הסקאלה (1 מינוס...) כי אנחנו רוצים extension נמוך
        df['Score_Ext'] = 1 - ((df['Extension'] - min_ext) / (max_ext - min_ext))
    else:
        df['Score_Ext'] = 0.5
        
    # ציון סופי
    df['Final_Score'] = (0.7 * df['Score_Trend']) + (0.3 * df['Score_Ext'])
    
    return df.sort_values('Final_Score', ascending=False)

# ============================================================================
# 5. EXCEL PORTFOLIO MANAGER (SIMULATION)
# ============================================================================

def load_portfolio():
    if not os.path.exists(Config.EXCEL_FILE):
        # יצירת קובץ חדש אם לא קיים
        df = pd.DataFrame(columns=['Ticker', 'Entry_Date', 'Entry_Price', 'Shares', 'Current_Price', 'PnL', 'Status'])
        # הון התחלתי כדמיון - נשמור בקובץ נפרד או פשוט נחשב דינמית
        return df, Config.INITIAL_CAPITAL
    
    df = pd.read_excel(Config.EXCEL_FILE)
    active_df = df[df['Status'] == 'Open']
    
    # חישוב הון פנוי (פשטני לדמו: הון התחלתי + רווחים סגורים - עלות פוזיציות פתוחות)
    closed_pnl = df[df['Status'] == 'Closed']['PnL'].sum()
    invested_capital = (active_df['Entry_Price'] * active_df['Shares']).sum()
    cash = Config.INITIAL_CAPITAL + closed_pnl - invested_capital
    
    return df, cash + invested_capital # מחזירים את ה-Total Equity

def save_trade(df, ticker, action, price, shares, pnl=0):
    date_str = datetime.now().strftime('%Y-%m-%d')
    
    if action == 'BUY':
        new_row = {
            'Ticker': ticker, 'Entry_Date': date_str, 'Entry_Price': price, 
            'Shares': shares, 'Current_Price': price, 'PnL': 0, 'Status': 'Open'
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        print(f"{Fore.GREEN} [EXECUTE] BOUGHT {shares} shares of {ticker} at {price:.2f}")
        
    elif action == 'SELL':
        # עדכון שורה קיימת
        idx = df[(df['Ticker'] == ticker) & (df['Status'] == 'Open')].index[0]
        df.at[idx, 'Status'] = 'Closed'
        df.at[idx, 'Current_Price'] = price # מחיר יציאה
        df.at[idx, 'PnL'] = (price - df.at[idx, 'Entry_Price']) * df.at[idx, 'Shares']
        print(f"{Fore.RED} [EXECUTE] SOLD {ticker} at {price:.2f} | PnL: {df.at[idx, 'PnL']:.2f}")
        
    df.to_excel(Config.EXCEL_FILE, index=False)
    return df

# ============================================================================
# 6. MAIN ENGINE
# ============================================================================

def run_daily_scan():
    print(f"\n{Fore.YELLOW}=== PURE OMAS DAILY SCAN ({datetime.now().strftime('%Y-%m-%d')}) ==={Style.RESET_ALL}")
    
    # 1. טעינת תיק
    portfolio_df, total_equity = load_portfolio()
    open_positions = portfolio_df[portfolio_df['Status'] == 'Open']
    held_tickers = open_positions['Ticker'].tolist()
    
    print(f"Total Equity: ${total_equity:,.2f}")
    print(f"Open Positions: {len(held_tickers)} / {Config.MAX_POSITIONS}")
    
    # 2. הורדת נתונים
    tickers = get_sp500_tickers()
    raw_data = download_data(tickers)
    
    candidates = []
    
    # 3. סריקת השוק
    print("\nScanning market...")
    for ticker in tickers:
        try:
            if ticker in raw_data.columns.levels[0]:
                df = raw_data[ticker].dropna()
                latest = calculate_indicators(df)
                
                if latest is None: continue
                current_price = latest['Close']
                
                # --- A. בדיקת יציאות (למניות מוחזקות) ---
                if ticker in held_tickers:
                    position_data = open_positions[open_positions['Ticker'] == ticker].iloc[0]
                    entry_price = position_data['Entry_Price']
                    
                    # Hard Stop Check
                    hard_stop_price = entry_price * (1 - Config.HARD_STOP_PCT)
                    
                    exit_reason = None
                    if current_price < hard_stop_price:
                        exit_reason = "HARD STOP LOSS (-15%)"
                    elif current_price < latest['Exit_Threshold']:
                        exit_reason = "Price Below Buffer (1%)"
                    elif latest['Slope'] <= 0:
                        exit_reason = "Slope Turned Negative"
                        
                    if exit_reason:
                        print(f"{Fore.RED}EXIT SIGNAL for {ticker}: {exit_reason}{Style.RESET_ALL}")
                        portfolio_df = save_trade(portfolio_df, ticker, 'SELL', current_price, 0)
                        held_tickers.remove(ticker) # מתפנה מקום מיד
                
                # --- B. בדיקת כניסות (למניות לא מוחזקות) ---
                else:
                    # תנאי כניסה: מעל Buffer וגם שיפוע חיובי (0.01$)
                    if (current_price > latest['Entry_Threshold']) and (latest['Slope'] > Config.SLOPE_THRESHOLD):
                        candidates.append({
                            'Ticker': ticker,
                            'Close': current_price,
                            'Entry_Threshold': latest['Entry_Threshold'],
                            'Slope_Norm': latest['Slope_Norm']
                        })
                        
        except Exception as e:
            continue

    # 4. ניהול כניסות חדשות
    open_slots = Config.MAX_POSITIONS - len(held_tickers)
    
    if open_slots > 0 and candidates:
        print(f"\nFound {len(candidates)} candidates for {open_slots} slots.")
        ranked_df = rank_candidates(candidates)
        
        print("\nTop Candidates (Ranked):")
        print(ranked_df[['Ticker', 'Final_Score', 'Score_Trend', 'Score_Ext']].head(5))
        
        # ביצוע קניות
        to_buy = ranked_df.head(open_slots)
        
        for _, row in to_buy.iterrows():
            allocation = total_equity * Config.POSITION_SIZE_PCT
            shares = int(allocation / row['Close'])
            if shares > 0:
                portfolio_df = save_trade(portfolio_df, row['Ticker'], 'BUY', row['Close'], shares)
    else:
        print("\nNo new entries (Portfolio full or no candidates).")

    print(f"\n{Fore.YELLOW}=== SCAN COMPLETE. Check portfolio_log.xlsx ==={Style.RESET_ALL}")

if __name__ == "__main__":
    run_daily_scan()