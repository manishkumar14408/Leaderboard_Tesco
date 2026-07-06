# Live Quiz Leaderboard from Excel

This package gives you a live HTML quiz leaderboard backed by your Excel score sheet.

## Files

- `leaderboard.html` - Visual leaderboard dashboard.
- `live_leaderboard_server.py` - Local server that reads Excel and exposes latest scores.
- `Premium_Quiz_Master_Dashboard.xlsx` - Keep your quiz workbook in the same folder.

## How to run

1. Extract the zip.
2. Keep `Premium_Quiz_Master_Dashboard.xlsx` in the same folder as the Python file.
3. Open Terminal or Command Prompt in that folder.
4. Run:

```bash
pip install openpyxl
python live_leaderboard_server.py
```

5. Open this URL in browser:

```text
http://localhost:8000
```

## How refresh works

- Update scores in the Excel file.
- Save the Excel file.
- The browser dashboard polls the local server every 2 seconds.
- The rank, leader, score bars and podium automatically update.

## Important note about Excel formulas

The server recalculates total and rank itself from Round 1 to Round 5. This avoids dependency on Excel's cached formula values.

## Change refresh speed

Open `leaderboard.html` and change:

```js
const REFRESH_MS = 2000;
```

For example, use `1000` for 1 second refresh.
