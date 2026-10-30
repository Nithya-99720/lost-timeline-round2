# Round 2 – Setup (shared storage via Google Sheet)

Browser storage only works on ONE device, so each team's result is POSTed to a Google Sheet that you own.

1. Create a Google Sheet. Open **Extensions → Apps Script**, replace the code with:

```js
const PASS='CHRONO2';            // must match CFG.PASS in the HTML
const COLS=['team_id','team_name','round_id','responses','completion_time','submitted_at','hints_used','build_attempts','wrong_answers','puzzle_times_s','puzzles_solved','fragment_recovered'];
function sh(){const s=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];if(s.getLastRow()==0)s.appendRow(COLS);return s}
function doPost(e){
  const r=JSON.parse(e.postData.contents);
  if(r.action==='clear'){                       // organizer 'Clear test data' button
    if(r.pass!==PASS)return ContentService.createTextOutput('denied');
    const t=sh();if(t.getLastRow()>1)t.deleteRows(2,t.getLastRow()-1);
    return ContentService.createTextOutput('cleared');
  }
  const s=sh(),ids=s.getRange(2,1,Math.max(s.getLastRow()-1,1),1).getValues().flat();
  const row=COLS.map(c=>r[c]??''),i=ids.indexOf(r.team_id);
  if(i>=0&&s.getLastRow()>1)s.getRange(i+2,1,1,COLS.length).setValues([row]);else s.appendRow(row);   // one row per team
  return ContentService.createTextOutput('ok');
}
function doGet(e){
  if(e.parameter.pass!==PASS)return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  const v=sh().getDataRange().getValues().slice(1).map(r=>Object.fromEntries(COLS.map((c,i)=>[c,r[i]])));
  return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);
}
```
2. **Deploy → New deployment → Web app**. Execute as: *Me*. Who has access: *Anyone*. Copy the `/exec` URL.
3. In `round2_architects_blueprint.html`, set `CFG.SHEET_URL='<your /exec URL>'`.
4. Host the HTML (GitHub Pages / Netlify drop / any LAN server) so all teams open the same file.
5. **Export**: click *Organizer export* in the footer, enter the passcode (`CHRONO2` – change it in both places). It downloads `round2_results.csv` with every team (one row per team). You can also use *File → Download → CSV* in the Sheet.

Clearing test data
- Click **Clear test data** on the start screen (passcode needed). It wipes this browser's saved results and the sheet's rows (header stays). If you edit the Apps Script, redeploy: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (the /exec URL stays the same).
- Or simply select the test rows in the Google Sheet and delete them.
- Do this right before the event, after your last test run.

Notes
- Team ID format is `LT-###` (edit `CFG.ID_RE` / `CFG.ID_HINT` to change; use the same one in all four rounds).
- Resubmitting the same Team ID overwrites that team's row (no duplicates).
- With `SHEET_URL=''` the page still works for testing; results stay in that browser only.
- The passcode lives in the page source – it's a convenience gate, not real security.
- Columns: required `team_id, team_name, round_id, responses, completion_time (seconds), submitted_at (ISO)` + extras `hints_used, build_attempts, wrong_answers, puzzle_times_s, puzzles_solved, fragment_recovered`. No points are calculated.
