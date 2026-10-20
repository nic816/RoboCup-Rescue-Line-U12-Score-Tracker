# RCAP Rescue Line Entry, U12 Score Tracker

A lightweight scoring dashboard for **RCAP Rescue Line Entry, U12** training runs. Coaches can log hazards, checkpoint tiles, and missions during live sessions — organised **per team** so you can track each team's progress over time. Everything is saved to a private Google Sheet in real time (or locally in the browser if you prefer zero setup).

![Dashboard Screenshot](screenshot.png)

---

# ✨ Features

- **Multi-Team Tracking** — Manage any number of teams right in the UI. Pick the active team from a dropdown before saving each run. No team names are hardcoded, so any trainer can use it.
- **Team Progress View** — Filter the log by team and see runs logged, best score, average, and the trend between the two most recent runs at a glance.
- **Competition Ranking** — Ranks teams by the mean of normalized field scores (rule 6.1), so you can see standings across rounds the way a real competition would.
- **Live Scoring Dashboard** — 5-minute countdown timer.
- **Hazard Counters** — Gaps (10 pts), Intersections (10 pts), Obstacles (20 pts).
- **Checkpoint Segments** — Dynamic list with tile counts and attempt multipliers.
- **Exit Bonus** — Lack-of-Progress counter and a "reached goal" toggle, computing the official exit bonus automatically.
- **Works Offline** — Runs entirely in the browser. Cloud saving to Google Sheets is optional.
- **No Server Required** — Pure HTML/JS plus optional Google Apps Script.

---

# 📋 Scoring Rules (RCAP Rescue Line Entry, U12 — v2026.0)

| Element | Points |
|---|---|
| Tile with one or more gaps navigated | +10 per tile |
| Intersection navigated | +10 per intersection |
| Obstacle overcome | +20 per obstacle |
| Checkpoint tile (1st attempt) | tiles × 5 |
| Checkpoint tile (2nd attempt) | tiles × 3 |
| Checkpoint tile (3rd attempt) | tiles × 2 |
| Checkpoint tile (beyond 3rd attempt) | tiles × 1 |
| Exit bonus (robot reaches goal and stops 5s) | max(0, 60 − 5 × total LoP) |

> **Exit Bonus (rule 5.6.5):** awarded only when the robot reaches the goal tile and stays stopped for 5+ seconds. Each Lack of Progress (LoP) reduces it by 5, and it never goes below 0. There is **no** separate "red line" bonus — the red tape simply marks the goal tile.

---

# 🚀 Getting Started

This is a single-file web application.

- **Local-only mode (no setup):** Just open `index.html` in a browser. Teams and runs are saved in that browser's local storage. Great for a single laptop at the training table.
- **Cloud mode (recommended for multiple devices):** Link it to your own Google Sheet so data is shared and backed up. See below.

---

## Step 1: Download the Tracker

1. Download the `index.html` file from this repository.
2. Open the file in any web browser to use the tracker locally.
3. To enable permanent cloud saving across devices, continue to Step 2.

---

## Step 2: Set Up Your Google Sheet Database

The database now uses **two tabs**: one for runs and one for your team list.

### 1. Create a New Google Sheet

Open: https://sheets.new/

### 2. Create the `Runs` tab

Rename the first tab to **`Runs`** and add these headers in the **first row** (exactly):

| Timestamp | Date | Team | Round | Gaps | Intersections | Obstacles | Hazards | Tiles | LoP | ExitBonus | Total |
|---|---|---|---|---|---|---|---|---|---|---|---|

> **Why the extra columns?** `Timestamp` is when the row was recorded, while `Date` is the **training session date you chose** (defaults to today, but editable so you can log a past week). Storing the hazard breakdown (gaps / intersections / obstacles) and the Lack-of-Progress count per run — not just the total — lets you spot exactly where a team is losing points and track week-by-week progress.

### 3. Create the `Teams` tab

Add a **second tab** named **`Teams`** with a single header in the first row:

| Team |
|---|

This tab holds your managed team list so it syncs across devices.

### 4. Open Google Apps Script

Go to **Extensions → Apps Script**, delete any existing code, and paste the following:

```javascript
const RUNS_SHEET = 'Runs';
const TEAMS_SHEET = 'Teams';

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// READ: returns { runs: [...], teams: [...] }
function doGet(e) {
  const runsSheet = getSheet(RUNS_SHEET);
  const teamsSheet = getSheet(TEAMS_SHEET);

  // Runs
  const data = runsSheet.getDataRange().getValues();
  const headers = data[0] || [];
  const runs = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  // Teams (first column, skip header)
  let teams = [];
  if (teamsSheet) {
    teams = teamsSheet.getRange(2, 1, Math.max(teamsSheet.getLastRow() - 1, 0), 1)
      .getValues()
      .map(r => r[0])
      .filter(String);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ runs: runs, teams: teams }))
    .setMimeType(ContentService.MimeType.JSON);
}

// WRITE / DELETE
function doPost(e) {
  const action = e.parameter.action;
  const runsSheet = getSheet(RUNS_SHEET);
  const teamsSheet = getSheet(TEAMS_SHEET);

  // Clear all runs
  if (action === 'clear') {
    if (runsSheet.getLastRow() > 1) {
      runsSheet.deleteRows(2, runsSheet.getLastRow() - 1);
    }
    return json({ result: 'cleared' });
  }

  // Add a team
  if (action === 'addTeam') {
    const name = e.parameter.Team;
    if (name && teamsSheet) {
      const existing = teamsSheet.getRange(2, 1, Math.max(teamsSheet.getLastRow() - 1, 0), 1)
        .getValues().map(r => r[0]);
      if (existing.indexOf(name) === -1) teamsSheet.appendRow([name]);
    }
    return json({ result: 'teamAdded' });
  }

  // Remove a team
  if (action === 'removeTeam') {
    const name = e.parameter.Team;
    if (name && teamsSheet) {
      const values = teamsSheet.getRange(2, 1, Math.max(teamsSheet.getLastRow() - 1, 0), 1).getValues();
      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i][0] === name) teamsSheet.deleteRow(i + 2);
      }
    }
    return json({ result: 'teamRemoved' });
  }

  // Default: save a run
  runsSheet.appendRow([
    new Date().toLocaleString(),
    e.parameter.Date || '',
    e.parameter.Team || '',
    e.parameter.Round || '',
    e.parameter.Gaps || 0,
    e.parameter.Intersections || 0,
    e.parameter.Obstacles || 0,
    e.parameter.Hazards || 0,
    e.parameter.Tiles || 0,
    e.parameter.LoP || 0,
    e.parameter.ExitBonus || 0,
    e.parameter.Total || 0
  ]);

  return json({ result: 'success' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 5. Deploy the Script as a Web App

1. Click **Deploy → New Deployment**
2. Select the **gear icon** → **Web App**
3. Under **Execute as**, select `Me`
4. Under **Who has access**, select `Anyone`
5. Click **Deploy**, authorize if prompted, and copy the **Web App URL**

---

## Step 3: Connect the Tracker to Your Google Sheet

1. Open `index.html` in a text editor (VS Code, Notepad++, Sublime, Notepad).
2. Find this line near the top of the `<script>` section:

```javascript
const scriptURL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
```

3. Replace the placeholder with your copied URL:

```javascript
const scriptURL = 'https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXX/exec';
```

4. Save the file and open it in your browser. You're now in cloud mode.

> If you leave the placeholder as-is, the tracker still works fully — it just saves to the local browser instead of the cloud.

---

# 🖥️ Usage

1. **Add your teams** once, using the "Add a New Team" box. They stay in the dropdown.
2. **Set the Session Date** (defaults to today — change it if you're logging a past training week), then **select the active team** and enter the session name / round.
3. Start the timer when a run begins.
4. Record hazards (gaps, intersections, obstacles) as they occur.
5. Add checkpoint segments and tile counts, and pick the attempt multiplier.
6. Increment the **Lack of Progress** counter each time a LoP occurs, and toggle **"Reached goal & stopped 5s"** if the robot finishes — the exit bonus is computed automatically.
7. **Save the run.** It's stored against the selected team.
8. In **Team Progress**, filter by a team to see its runs, best score, average, and trend.
9. Check **Competition Ranking** for the overall standings across rounds (see below).

### How Competition Ranking works (rule 6.1)

Use a consistent **Round / Run** label (e.g. "Run 1", "Run 2") when logging, since ranking groups runs by that label. For each round:

1. Each team's **best run** in that round is taken.
2. That score is normalized against the top score in the round: `normalized = field score / best field score`.
3. A team's ranking metric is the **mean of its normalized scores** across the rounds it took part in.

Teams are ordered by that mean (highest first); ties are broken by best single run. Runs with no round label are grouped under "Unlabelled".

---

# ⚙️ Technologies Used

- HTML5, CSS3, JavaScript (Vanilla)
- Google Apps Script + Google Sheets (optional cloud backend)

---

# 🛠️ Troubleshooting

## Runs are not saving to the cloud
- Deployment access must be **Anyone**.
- Make sure you copied the `/exec` URL.
- The `Runs` and `Teams` tab names and headers must match exactly.
- Runs are always saved locally first, so nothing is lost even if sync fails.

## CORS error in the browser console
- Usually an incorrect URL, non-public deployment, or an old deployment version. Redeploy and update the URL.

## Saved runs do not reload
- Confirm `doGet()` exists and the deployment was updated after any code change, then refresh.

## Upgrading from the old single-sheet version
- The dashboard still reads the old flat format, so existing data will load. To get the new per-team and per-week columns, add the extra headers (including `Date`) to your `Runs` tab (or start a fresh sheet) and create the `Teams` tab. Make sure the `Runs` column order matches the Apps Script's `appendRow` above.

---

# 📜 License

This project is licensed under the MIT License.

You are free to use, modify, and distribute this software for educational purposes, robotics teams, and competitions.
