# iNatlookup — iNaturalist metadata lookup for Google Sheets

A bound Google Apps Script for a sample-intake spreadsheet: enter an
iNaturalist observation id (or a voucher/tag/accession value held in an
iNaturalist observation field) and the script fills in the observation's
metadata. Built as companion tooling to the
[AGGATCATTA](https://github.com/mycosomatic/aggatcatta) fungal barcoding
pipeline, but usable with any spreadsheet.

## Setup

1. In the spreadsheet: **Extensions → Apps Script**.
2. Replace the default `Code.gs` with this repo's `Code.gs`.
3. **File → New → HTML**, name it `FieldSelector`, paste in
   `FieldSelector.html`.
4. Save, reload the spreadsheet, authorize on first menu use.
5. **iNaturalist → Auto-lookup: OFF (click to enable)** to install the
   edit trigger.

There is no automatic deployment — after changing files here, update the
spreadsheet's bound script by hand (or wire up
[clasp](https://github.com/google/clasp) if that gets tedious).

## How a row is looked up

Row 1 must be the header row. Any column whose header matches one of the
lookup options — `id`, `iNat URL`, `Observation URL`, `Voucher Number(s)`,
`FUNDIS Tag Number`, `Accession Number`, `GenBank Accession Number` — is a
lookup column. Header matching ignores case, spacing, and punctuation, so
`voucher numbers` matches `Voucher Number(s)`.

- **Auto-lookup (edit trigger):** the column you actually edited is the
  one used. Clearing a cell does nothing (no fallback to another column).
- **Manual runs (menu):** a filled id-style column wins (exact and fast);
  otherwise the first filled lookup column in header order is used.
- The three id-style columns (`id`, `iNat URL`, `Observation URL`) all
  behave identically and accept either a bare observation id or a full
  observation URL (`https://www.inaturalist.org/observations/12345`) —
  QR-scan a label's URL straight into one of them; no splitting or
  number-extraction formulas needed. Formula-computed cells never fire
  edit triggers (a Sheets limitation), so scan into the lookup column
  directly rather than deriving it with SPLIT/REGEXEXTRACT.
- When a lookup goes through an observation field, `id` is always written
  back, so later runs can use the direct id path.
- If several observations share the same observation-field value, the most
  recent is used and the cell gets a note saying so.

Return fields are chosen in **Choose Return Fields (Sidebar)** and stored
per spreadsheet.

## Notes on behavior

- **Rows the trigger can't run immediately are queued, not dropped.** If a
  long run holds the lock while you keep entering ids, those edits are
  stored and processed as soon as the running job finishes.
- **Pasting a whole column of values is fine.** One paste is one job, blank
  cells are skipped, and duplicate values are looked up once. Results are
  written in chunks of ~30 rows as the run goes; if a very large paste
  approaches Apps Script's ~6-minute execution limit, the remainder queues
  itself and continues automatically about a minute later — no re-paste
  needed.
- Lookups by id are batched (100 per request); repeated lookups are served
  from a 30-minute cache; 429/5xx responses are retried with backoff.
- Only the selected return columns are written. Failures put a note on the
  lookup cell and a row in the `_iNatLog` sheet.
- The auto-lookup ON/OFF flag is shared by everyone on the spreadsheet;
  turning it off stops lookups regardless of which account installed the
  trigger.

## If auto-lookup doesn't fire

- Values filled by **formulas or IMPORTRANGE never fire edit triggers** —
  that is a Google Sheets limitation. Use the menu's manual runs for those.
- The header row must be row 1 on the sheet being edited.
- The trigger belongs to the account that enabled it; if the spreadsheet
  was copied, re-enable from the menu in the copy.
- Check the `_iNatLog` sheet for `TriggerError` rows.
