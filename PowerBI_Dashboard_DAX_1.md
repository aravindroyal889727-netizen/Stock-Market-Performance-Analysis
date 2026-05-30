# Power BI Dashboard Design & DAX Measures
## Stock Market Performance Analysis

---

## Data Cleaning Steps (Power Query M)

```m
// Step 1: Load & clean date column
#"Parsed Date" = Table.TransformColumnTypes(
    Source, {{"Date", type date}},
    "en-US"
),

// Step 2: Strip $ from price columns
#"Clean Close" = Table.TransformColumns(
    #"Parsed Date",
    {{"Close/Last", each Number.From(Text.Replace(_, "$", "")), type number}}
),
#"Clean Prices" = Table.TransformColumns(
    #"Clean Close",
    {
        {"Open",  each Number.From(Text.Replace(_, "$", "")), type number},
        {"High",  each Number.From(Text.Replace(_, "$", "")), type number},
        {"Low",   each Number.From(Text.Replace(_, "$", "")), type number}
    }
),

// Step 3: Rename column
#"Renamed" = Table.RenameColumns(
    #"Clean Prices",
    {{"Close/Last", "Close"}}
),

// Step 4: Sort ascending
#"Sorted" = Table.Sort(#"Renamed", {{"Date", Order.Ascending}})
```

---

## Dashboard Layout Design

### Page 1 — Executive Overview

```
┌─────────────────────────────────────────────────────────────────┐
│  STOCK PERFORMANCE DASHBOARD          Jan 21 – Dec 2, 2026      │
├──────────┬──────────┬──────────┬──────────┬──────────┬──────────┤
│  Latest  │  Period  │  Period  │  Return  │  Avg Vol │ Bull Days│
│  Close   │  High    │   Low    │  (Total) │          │  vs Bear │
│ $261.73  │ $280.91  │ $244.68  │  +5.69%  │  54.9M   │  13 / 8  │
├──────────┴──────────┴──────────┴──────────┴──────────┴──────────┤
│                                                                   │
│  LINE CHART: Close Price + MA5 + MA10                            │
│  (Full width, Jan–Dec 2026, dual Y-axis: price left, vol right)  │
│                                                                   │
├───────────────────────────┬───────────────────────────────────────┤
│  BAR CHART: Daily Volume  │  BAR CHART: Daily Returns (%)         │
│  (Conditional color:      │  (Green = positive, Red = negative)   │
│   purple > 80M, else blue)│                                       │
└───────────────────────────┴───────────────────────────────────────┘
```

### Page 2 — Risk & Volatility

```
┌─────────────────────────────────────────────────────────────────┐
│  RISK ANALYSIS                                                   │
├──────────┬──────────┬──────────┬──────────────────────────────── │
│ Max Gain │ Max Loss │ Avg Ret  │  Peak Volatility                 │
│  +4.06%  │  -5.27%  │  +0.29% │  3.75% (Feb 19)                  │
├──────────┴──────────┴──────────┴──────────────────────────────── │
│                                                                   │
│  AREA CHART: 5-Day Rolling Volatility                            │
│                                                                   │
├───────────────────────────┬───────────────────────────────────────┤
│  BAR CHART: Price Range   │  HISTOGRAM: Return Distribution       │
│  (High − Low per session) │  (Bucketed by 1% intervals)           │
└───────────────────────────┴───────────────────────────────────────┘
```

### Page 3 — Trading Intelligence

```
┌─────────────────────────────────────────────────────────────────┐
│  TRADING SIGNALS & PATTERN ANALYSIS                              │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  COMBO CHART: Candlestick (OHLC) + Volume bars                  │
│  (Requires custom visual: "Candlestick" from AppSource)          │
│                                                                   │
├───────────────────────────┬───────────────────────────────────────┤
│  SCATTER: Volume vs       │  TABLE: Top 5 High-Volume Sessions    │
│  Absolute Return          │  with date, close, volume, return     │
│  (Bubble = range size)    │                                       │
└───────────────────────────┴───────────────────────────────────────┘
```

---

## DAX Measures

### Core Price Measures

```dax
// Latest closing price (most recent date)
Latest Close =
CALCULATE(
    LASTNONBLANK(StockData[Close], 1),
    FILTER(StockData, StockData[Date] = MAX(StockData[Date]))
)

// Period high (highest intraday high)
Period High =
MAX(StockData[High])

// Period low (lowest intraday low)
Period Low =
MIN(StockData[Low])

// First close in selected period
First Close =
CALCULATE(
    FIRSTNONBLANK(StockData[Close], 1),
    FILTER(StockData, StockData[Date] = MIN(StockData[Date]))
)

// Total return over selected period (%)
Total Return % =
DIVIDE(
    [Latest Close] - [First Close],
    [First Close]
) * 100

// Price range (close max - close min)
Price Range =
MAX(StockData[Close]) - MIN(StockData[Close])
```

### Daily Return Measures

```dax
// Daily return % (requires sorted index column in data)
Daily Return % =
VAR CurrentDate = StockData[Date]
VAR PreviousClose =
    CALCULATE(
        MAX(StockData[Close]),
        FILTER(
            ALL(StockData),
            StockData[Date] = MAXX(
                FILTER(ALL(StockData), StockData[Date] < CurrentDate),
                StockData[Date]
            )
        )
    )
RETURN
DIVIDE(StockData[Close] - PreviousClose, PreviousClose) * 100

// Average daily return
Avg Daily Return % =
AVERAGEX(
    FILTER(StockData, NOT(ISBLANK(StockData[Daily_Return]))),
    StockData[Daily_Return]
)

// Best single-day return
Best Day Return % =
MAX(StockData[Daily_Return])

// Worst single-day return
Worst Day Return % =
MIN(StockData[Daily_Return])

// Count of positive return days
Positive Days =
COUNTROWS(
    FILTER(StockData, StockData[Daily_Return] > 0)
)

// Count of negative return days
Negative Days =
COUNTROWS(
    FILTER(StockData, StockData[Daily_Return] < 0)
)

// Bull/Bear ratio
Bull Bear Ratio =
DIVIDE([Positive Days], [Positive Days] + [Negative Days])
```

### Volume Measures

```dax
// Average daily trading volume
Avg Daily Volume =
AVERAGE(StockData[Volume])

// Maximum volume session
Max Volume =
MAX(StockData[Volume])

// Volume vs average (% above/below)
Volume vs Avg % =
DIVIDE(
    StockData[Volume] - [Avg Daily Volume],
    [Avg Daily Volume]
) * 100

// High volume flag (>80M shares)
Is High Volume =
IF(StockData[Volume] > 80000000, "High", "Normal")

// Total volume traded in period
Total Volume =
SUM(StockData[Volume])
```

### Moving Average Measures

```dax
// 5-day simple moving average (calculated in DAX)
MA5 =
AVERAGEX(
    TOPN(
        5,
        FILTER(
            ALL(StockData),
            StockData[Date] <= MAX(StockData[Date])
        ),
        StockData[Date], DESC
    ),
    StockData[Close]
)

// 10-day simple moving average
MA10 =
AVERAGEX(
    TOPN(
        10,
        FILTER(
            ALL(StockData),
            StockData[Date] <= MAX(StockData[Date])
        ),
        StockData[Date], DESC
    ),
    StockData[Close]
)

// Price above MA5 (bullish signal)
Above MA5 =
IF([Latest Close] > [MA5], "Bullish", "Bearish")

// MA crossover signal (MA5 vs MA10)
MA Crossover Signal =
IF([MA5] > [MA10], "Golden Cross ↑", "Death Cross ↓")
```

### Volatility Measures

```dax
// 5-day rolling standard deviation of returns
Rolling Volatility 5D =
STDEVX(
    TOPN(
        5,
        FILTER(
            ALL(StockData),
            StockData[Date] <= MAX(StockData[Date])
                && NOT(ISBLANK(StockData[Daily_Return]))
        ),
        StockData[Date], DESC
    ),
    StockData[Daily_Return]
)

// Daily high-low range
Daily Range =
StockData[High] - StockData[Low]

// Average daily range
Avg Daily Range =
AVERAGEX(StockData, StockData[High] - StockData[Low])

// Volatility classification
Volatility Level =
VAR v = [Rolling Volatility 5D]
RETURN
SWITCH(
    TRUE(),
    v >= 3,     "High",
    v >= 1.5,   "Moderate",
    v >= 0,     "Low",
    "N/A"
)
```

### Conditional Formatting Expressions

```dax
// Color rule for daily return bars (use in conditional formatting)
Return Color =
IF(StockData[Daily_Return] >= 0, "#1D9E75", "#D85A30")

// Color rule for volume bars
Volume Color =
IF(StockData[Volume] > 80000000, "#534AB7", "#AFA9EC")

// KPI icon for total return
Return Icon =
IF([Total Return %] >= 0, "▲ " & FORMAT([Total Return %], "0.00") & "%",
   "▼ " & FORMAT(ABS([Total Return %]), "0.00") & "%")
```

### Advanced Analytics

```dax
// Cumulative return from start of period
Cumulative Return % =
VAR StartPrice = [First Close]
VAR CurrentPrice =
    CALCULATE(MAX(StockData[Close]), FILTER(ALL(StockData), StockData[Date] <= MAX(StockData[Date])))
RETURN
DIVIDE(CurrentPrice - StartPrice, StartPrice) * 100

// Sharpe-like ratio (avg return / volatility, annualized proxy)
Risk Adjusted Return =
DIVIDE(
    [Avg Daily Return %],
    [Rolling Volatility 5D]
)

// Sessions with volume > 1 std dev above average
Abnormal Volume Sessions =
VAR AvgVol = [Avg Daily Volume]
VAR StdVol = STDEVX(ALL(StockData), StockData[Volume])
RETURN
COUNTROWS(
    FILTER(StockData, StockData[Volume] > AvgVol + StdVol)
)
```

---

## Recommended Visuals & Settings

| Chart | Visual Type | Key Settings |
|-------|-------------|-------------|
| Price + MA lines | Line chart | Dual Y-axis, date on X, enable tooltips with all fields |
| Volume bars | Clustered bar | Conditional color by threshold |
| Daily returns | Clustered bar | Conditional color by positive/negative |
| Volatility | Area chart | Semi-transparent fill |
| OHLC candles | Custom visual: "Candlestick by OKViz" | AppSource install required |
| Return distribution | Custom visual: "Histogram Chart" | Bin width = 1% |
| Volume vs Return | Scatter chart | Bubble size = Daily_Range |
| KPI cards | Card visual | Conditional icon using Return Icon measure |

---

## Recommended Slicers

- Date range slicer (between)
- Month selector (dropdown)
- Volume tier filter (High / Normal)
- Volatility level filter (High / Moderate / Low)

---

## Data Model Notes

- Single table model is sufficient for this dataset
- Add a separate **Date dimension table** for time intelligence:

```dax
DateTable =
CALENDAR(DATE(2026,1,1), DATE(2026,12,31))
```

- Mark as Date Table, set active relationship on `StockData[Date]`
- Add Month, Quarter, WeekNum, DayOfWeek columns to DateTable for slicers
