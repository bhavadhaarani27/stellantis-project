# EHEQ Monthly Review – Quick Start

**Classification: App.** An AI-built tool; no AI is invoked during use. It needs only Node.js – no PowerPoint, SharePoint, Databricks, IT services or AI at runtime.

## Every month
1. Double-click `Start-EHEQ-Review.bat`. The generator opens at http://localhost:8080/.
2. Choose Month and Year.
3. For Session 1 and/or Session 2: select the `.pptx`, enter the **meeting date**, then tick the teams the manager chose. Use ▲▼ to set their order. Click Generate.
4. Open the generated link. Each session is its own page; a missing session is simply not generated.

## Late teams (panel "Late teams and SharePoint links")
Select an existing review, then per team:
- **Late PowerPoint** – upload it; the team is added or replaced, other teams are untouched.
- **SharePoint link** – paste the https link; the page shows an "Open in SharePoint" slide (opens in a new tab).
- **From synced folder** – in SharePoint click "Sync" on the team folder once, paste the resulting Windows folder path into the panel, then pick the team's latest deck from the dropdown (no download/upload, no IT needed).
- **Mark pending** – shows an "Awaiting slides" slide.
- **Remove**, **▲▼ reorder**, and the **meeting date** can be changed at any time; the page rebuilds automatically.

## Notes
- Slides are drawn from the PPTX (text, tables, charts, images). The sidebar lists teams from the deck's PowerPoint sections; teams not in the file are greyed out.
- The cover shows the entered meeting date; the other session keeps a dummy date.
- Slides hidden in PowerPoint are hidden by default; use "Show hidden slides" in the sidebar.
- Only `.pptx` is supported (save old `.ppt` as `.pptx`).
- Output: `monthly-output\simple\<Month>-<Year>\Session1|Session2\`. Regenerating a session replaces only that session.
- Limits: SharePoint contents cannot be read without IT access (link only); SmartArt, EMF/WMF images and exotic charts are approximate.
- Databricks is a possible future data source; not used or required.
