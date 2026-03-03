# Event-Based Result Checker - Implementation Summary

## Overview
Implemented an intelligent, event-based result checker that validates sports betting picks at sport-specific times rather than polling at fixed intervals.

## Key Features

### 1. Sport-Specific Delays (CST Timezone)
- **Soccer & NBA**: Check 1h 20min after pick is placed
- **MLB, NFL, NHL**: Check 3 hours after pick is placed
- Automatic scheduling when pick is logged to Excel
- No manual configuration needed

### 2. How It Works

When a pick is added:
```
1. Pick is parsed and added to Excel
2. scheduleResultCheck() is called with pick details
3. A timer is created with sport-specific delay
4. After delay expires:
   - Fetch scores from Odds API for that sport
   - Match pick against game scores
   - Calculate result (Win/Loss/Push)
   - Update Excel sheet with result
   - Send Telegram alert with result details
5. Clean up timer from memory
```

### 3. API Efficiency
- **Cost**: ~1-2 API calls per pick (only when result is ready)
- **Monthly budget**: 8 picks/day × 1-2 calls = 8-16 credits/day = ~240-480 credits/month
- **Your limit**: 500 credits/month ✅ **Safe!**
- **Benefit**: No wasted calls checking incomplete games

### 4. Result Detection

Supports all pick types:
- **ML (Moneyline)**: Home/Away team wins
- **Over/Under**: Total goals/points above/below threshold
- **BTTS**: Both teams to score (soccer)
- **AH (Asian Handicap)**: Spread betting
- **Pushes**: Automatic detection for ties/draws

### 5. Excel Integration

Automatically updates when game finishes:
```
Tipster Sheet:
- Result column: W/L/P
- Profit/Loss: Auto-calculated based on odds
- Balance: Running total

General Sheet:
- All picks from all tipsters
- Same result/profit tracking

Summary Sheet:
- Win rate, yield, running stats
```

### 6. Telegram Alerts

When result is determined, sends:
```
✅ Result: Win
Tipster: Cristian Rey
Match: Boston Celtics vs Miami Heat
Pick: Boston ML
Odds: 1.85
P&L: $1700
Balance: $8500
```

## Code Changes

### New Files
- `results.js` - Event-based result checker (replaced old polling version)

### Modified Files
- `excel.js` - Added `updatePickResult()` for result updates
- `index.js` - Calls `scheduleResultCheck()` after each pick is logged

## Timezone Handling

All delays are calculated in **CST (Central Standard Time)** as specified:
- Delays use milliseconds: `80 * 60 * 1000` = 80 minutes = 1h 20min
- System maintains no timezone conversion (works locally on your CST machine)
- Excel dates are stored as display dates (e.g., "Mar 2")

## Active Timers Management

- Timers stored in `Map<pickId, timeoutId>`
- Prevents duplicate checks if pick data changes
- `cancelAllChecks()` available on app shutdown
- Auto-cleanup when result is determined

## Error Handling

If game not finished when checked:
- Automatic retry after 5 minutes
- Prevents "not completed" errors from failing the check
- Retries until result is available

If no matching score found:
- Logs warning but doesn't crash
- Allows manual review of edge cases

## Future Enhancements

Possible additions:
1. Load pending picks on app startup and reschedule their checks
2. Webhook notifications instead of polling
3. Support for live odds updates
4. Historical performance analytics dashboard

## Testing

To test with a pick:
1. Send a WhatsApp pick image/message
2. Watch logs for:
   ```
   [Results] ⏰ Scheduled check for nba (80min): Boston Celtics vs Miami Heat
   ... (wait 1h 20min) ...
   [Results] Checking: Boston Celtics vs Miami Heat
   [Results] Score: Boston Celtics 110 - 105 Miami Heat (total 215)
   [Results] ✅ Boston Celtics vs Miami Heat | Boston ML → W ($1700)
   ```
3. Check Excel for updated result row
4. Check Telegram for alert

## API Cost Breakdown

Monthly estimate:
- Initialization (odds cache): 5 credits
- Pick intake: 0 credits (cached from startup)
- Result checking: ~240-480 credits (1-2 per pick × 8/day × 30 days)
- **Total: ~245-485 credits/month (within 500 limit!)**
