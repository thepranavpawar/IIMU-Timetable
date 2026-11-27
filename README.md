# IIMU Term-IV calendars

Turns the official Term-IV timetable into a personal calendar for each person:
only their own electives, their section's core sessions (CS / ENTR), and the
campus events (Finclave, exams, Audacity, …).

**How it works:** every 3 hours GitHub Actions downloads the timetable from
Google Drive, builds one `.ics` file per person, and publishes them on GitHub
Pages. No Google login or secrets are needed, because the timetable is shared
"anyone with the link".

## Subscribe

Open the Pages site (`https://<username>.github.io/<repo>/`) to see every
link. In Google Calendar: **Other calendars (+) → From URL** → paste e.g.

```
https://<username>.github.io/<repo>/shivali.ics
```

Google re-checks subscribed calendars on its own schedule (often every
8–24 hours), so timetable changes can take that long to appear.

## Add someone / change courses

Edit `config.json` and add an entry to `people`:

```json
{ "slug": "kevin", "calendar_name": "IIMU Term-IV (Kevin)",
  "section": "DEM", "courses": ["CS", "ENTR", "B2BM", "OE", "NiS", "TC"] }
```

Course codes: CS, ENTR, AETBVC, B2BM, DSPAE, EAT, Fintech, NiS, OME, OE, SCC,
SCF, TC. Committing the change rebuilds the calendars in about a minute.

## One-time setup

**Settings → Pages → Build and deployment → Source: GitHub Actions**, then
**Actions → Build calendars → Run workflow**.

## Notes

* Capstone sessions 6–8 are one-on-one group meetings scheduled separately, so they aren't in the timetable.
* If the programme office uploads a *new* timetable file, put its ID in `source_file_id`.
* Local test: `pip install openpyxl && python build_ics.py timetable.xlsx`
