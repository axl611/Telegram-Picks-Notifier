# Test Results Script - `test-results-today.js`

## Purpose
Forces immediate API calls to check and verify match results for **only today's picks**. 
Useful for testing the result evaluation logic without waiting for scheduled delays.

## Why Use This
- ✅ Test result evaluation without waiting 1h20min (soccer/NBA) or 3hrs (MLB/NFL/NHL)
- ✅ Verify API integration is working correctly
- ✅ Test pick matching logic (team name normalization)
- ✅ Verify profit/loss calculations
- ✅ Ensure Excel formulas recalculate properly

## Usage

```bash
node test-results-today.js
```

## How It Works

1. **Reads Excel file** - Opens `picks.xlsx`
2. **Finds today's picks** - Filters for picks with today's date (format: "2 Mar")
3. **Filters incomplete picks** - Only checks picks with empty Result column (G)
4. **Groups by sport** - Organizes picks to minimize API calls
5. **Fetches scores** - Calls Odds API once per sport
6. **Evaluates each pick** - Determines W/L/P based on scores
7. **Updates Excel** - Writes result letters to column G
8. **Colors results** - Green (W), Red (L), Yellow (P)
9. **Displays summary** - Shows how many picks were updated

## Example Output

```
🧪 TEST: Checking results for TODAY (2 Mar)
================================================

📋 Roberto Rey: 1 pick(s) today
   Fetching NBA scores...
   ✅ Fetched 10 games from API
      ❌ No matching score: GS Warriors vs. LA Clippers
         (Game not found in API response, may not have started)

📋 Cristian Rey: 2 pick(s) today
   Fetching SOCCER scores...
   ✅ Fetched 15 games from API
      ✅ Manchester United vs. Liverpool | Over 2.5 → W ($360)
      ✅ Barcelona vs. Real Madrid | ML Barcelona → L (-$2000)

📊 Summary:
   Total picks today: 3
   Results found & updated: 2
   Could not process: 1
   Still pending: 0
```

## Output Interpretation

- ✅ **Green check**: Result was found and updated
- ❌ **Red X**: No matching score found (game may not have started or is not in API)
- ⏳ **Hourglass**: Game found but not completed yet
- ❓ **Question mark**: Game completed but pick couldn't be evaluated (unusual pick format)

## What Results Look Like in Excel

After running the script:

```
Date    | Sport | Match              | Pick        | Odds | Result | P&L      | Balance
--------|-------|------------------|-------------|------|--------|----------|--------
2 Mar   | SOCCER| Man Utd vs Liver  | Over 2.5    | 1.68 | W      | $360.00  | $360.00
2 Mar   | SOCCER| Barcelona vs Real | ML Barcelona| 2.10 | L      | -$2000   | -$1640
```

- **Result Column** (G): Contains W, L, or P (colored appropriately)
- **P&L Column** (H): Auto-calculated from formula based on Result and Odds
- **Balance Column** (I): Running total (auto-calculated)

## Notes

- Only processes picks from **today** (current date in format "D Mon", e.g., "2 Mar")
- Skips picks that already have a result (won't overwrite)
- Makes API calls **immediately** (doesn't wait for sport-specific delays)
- Results are saved to Excel automatically
- Does NOT send Telegram notifications (unlike production scheduler)
- Uses same matching/evaluation logic as production code

## Troubleshooting

### "No scores found for SOCCER"
- Game may not have started yet
- Check Odds API for available sports
- Verify API key in `.env` is valid

### "No matching score: [Match Name]"
- Team names might not match exactly
- The normalization function removes special characters and accents
- Try entering team names with more common versions

### "Game not completed yet"
- Game is in progress or hasn't started
- Wait for games to complete and run script again
- Some leagues post final scores with delay

### "Could not evaluate: [Pick]"
- Pick format might be unusual
- Check if pick uses supported formats: Over/Under, ML, AH, BTTS

## Supported Pick Formats

The script can evaluate these pick types:

1. **Over/Under** 
   - "Over 2.5" → Checks if total goals > 2.5
   - "Under 45.5" → Checks if total points < 45.5

2. **Moneyline (ML)**
   - Picks the first team mentioned in match
   - "Man United" → Wins if Man Utd wins
   
3. **Asian Handicap (AH)**
   - "AH +0.5" or "AH -1.0"
   - Applies handicap to score difference

4. **Both Teams to Score (BTTS)**
   - "Ambos equipos" or "BTTS"
   - "No ambos equipos" → Pick wins if NOT both score

## Performance

- API calls: 1 per sport (grouped)
- Example: 6 picks across 2 sports = 2 API calls
- Execution time: 2-5 seconds typically
- No delays: Results checked immediately

## Integration with Production

- Uses same matching/evaluation logic as `results.js`
- Excel formulas handle P&L calculation (not hardcoded)
- Produces same output as scheduled result checker
- Safe to run multiple times (won't duplicate results)

---

**Use this before launching live to verify everything works!** 🚀
