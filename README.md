# 🇱🇰 CSE Stock Market Trend Analyzer

An automated stock market tracker for the **Colombo Stock Exchange (CSE), Sri Lanka**.  
Built with Python and Microsoft Excel — fetches real live data from cse.lk every day with zero manual work.

---

## 📸 Preview

| Sheet | Description |
|---|---|
| 🇱🇰 CSE Live Prices | Real-time prices for 10 CSE stocks |
| 📊 Gainers & Losers | Today's top movers |
| 📊 RAW DATA | Auto-growing price history |
| 🔬 Analysis | Moving averages & BUY/SELL signals |
| 📈 Dashboard | Interactive charts & KPI cards |

---

## 🚀 What It Does

- Fetches **live CSE stock prices** directly from cse.lk API every day
- **Automatically adds one new row** to the price history each trading day — no copy-paste ever
- Calculates **7-Day & 20-Day Moving Averages** using Excel formulas
- Generates **BUY ▲ / SELL ▼ signals** based on moving average crossovers
- Shows **top gainers, top losers & most active** trades of the day
- Displays everything in a clean **Excel Dashboard** with charts that update automatically
- Smart duplicate protection — running the script twice in one day won't add duplicate rows

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|---|---|
| Python 3 | Fetching data from CSE API & writing to Excel |
| `requests` library | Connecting to cse.lk API |
| `openpyxl` library | Reading and writing the Excel file |
| Microsoft Excel | Dashboard, charts, formulas |
| CSE API (cse.lk) | Live stock price data source |

---

## 📁 Project Structure

```
CSE Project/
├── CSE_Stock_Trend_Analyzer.xlsx   # Main Excel file (dashboard + data)
├── cse_updater.py                  # Python script (fetches & updates data)
└── README.md                       # This file
```

---

## ⚙️ How To Run

### 1. Install Python
Download from [python.org](https://www.python.org/downloads/)  
During install — make sure to tick **"Add Python to PATH"**

### 2. Install required libraries
Open Command Prompt and run:
```bash
python -m pip install requests openpyxl
```

### 3. Set up your project folder
Put these two files in the same folder:
- `CSE_Stock_Trend_Analyzer.xlsx`
- `cse_updater.py`

### 4. Run the script
Open Command Prompt in your project folder and run:
```bash
python cse_updater.py
```

### 5. Open Excel
Open `CSE_Stock_Trend_Analyzer.xlsx` and go to the **📈 DASHBOARD** sheet.

### 6. Run every trading day
Run the script every weekday after **2:30 PM** (when CSE closes).  
Each run adds one new row of price history automatically.

---

## 📊 How The Analysis Works

```
RAW DATA sheet
     ↓  (grows by 1 row every day you run the script)
ANALYSIS sheet  →  calculates automatically:
     • Daily Return %     (how much price changed today)
     • 7-Day Moving Average   (average of last 7 closing prices)
     • 20-Day Moving Average  (average of last 20 closing prices)
     • BUY ▲ Signal       (when 7-Day MA crosses above 20-Day MA)
     • SELL ▼ Signal      (when 7-Day MA crosses below 20-Day MA)
     ↓
DASHBOARD sheet  →  charts & KPI cards update automatically
```

---

## 📈 Understanding BUY / SELL Signals

| Signal | Meaning |
|---|---|
| **BUY ▲** (green) | 7-Day MA is above 20-Day MA → upward momentum → possible good time to buy |
| **SELL ▼** (red) | 7-Day MA is below 20-Day MA → downward momentum → possible time to be cautious |
| **WAIT** | Not enough data yet (need at least 20 days) |

> ⚠️ These signals are for educational purposes only. Always do your own research before making investment decisions.

---

## ➕ How To Add More Stocks

1. Open `cse_updater.py` in Notepad
2. Find the `STOCKS_TO_TRACK` list near the top
3. Add any CSE stock symbol:
```python
STOCKS_TO_TRACK = [
    "COMB.N0000",   # Commercial Bank of Ceylon
    "LOLC.N0000",   # LOLC Holdings
    "DIAL.N0000",   # Dialog Axiata
    # Add more here...
    "LION.N0000",   # Lion Brewery
]
```
4. Find symbols at [www.cse.lk](https://www.cse.lk) — format is always `SYMBOL.N0000`

---

## ⏰ Optional — Automate With Windows Task Scheduler

Want the script to run itself every day without you doing anything?

1. Search **"Task Scheduler"** in Windows Start menu
2. Click **Create Basic Task**
3. Set trigger: **Daily** at **2:45 PM** (weekdays)
4. Set action: **Start a program** → `python`
5. Add arguments: `cse_updater.py`
6. Set start in: your project folder path (e.g. `D:\Project\CSE Project`)
7. Click Finish — it runs itself every trading day! ✅

---

## ⚠️ Important Notes

- Run the script on **weekdays only** (CSE trades Mon–Fri)
- Run **after 2:30 PM** for the most accurate closing prices
- Do **NOT** manually edit the RAW DATA sheet — the script manages it
- Do **NOT** change the column order in RAW DATA
- The script will **not add duplicate rows** if you run it twice in one day

---

## 👤 Author

**[Thanuka Sachith]**  
Aspiring Data Analyst | Sri Lanka  


---

## 📄 License

This project is open source and free to use for learning and personal purposes.