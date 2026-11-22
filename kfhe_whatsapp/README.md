# KFHE screed – WhatsApp extract

Turns WhatsApp chat exports into `KFHE_WhatsApp_extract_<date>.xlsx` and `.md`, with these sheets:
Attendance, Progress, Materials, Plant, Money, Instructions, Other.
Columns: Date | Chat | Sender | Item | Qty | Unit | Detail (verbatim text).
Photos are copied to `images/`. Only messages from 16 Oct 2026 to today are kept.

## 1. Export the chats (phone, read-only)
For each chat (the KFHE / King Faisal site group, Deborah, Prince, Aime, Maxime, Mai, Karen, Ange):
WhatsApp → open chat → ⋮ (or the contact name on iPhone) → **More → Export chat → Include media** → save the .zip to Google Drive.
Exporting does not send, change or mark anything in the chat.

## 2a. Easiest: upload the .zip files to Claude
Upload them in the Claude chat. Claude runs the script, reads each attendance and payroll photo and
types it in. Anything it can't read is marked UNCLEAR.

## 2b. Or run it yourself on Windows
```
pip install openpyxl
python kfhe_extract.py "G:\My Drive\BHI\Shelter africa\King faisal expansion-subcontract\09_Other packages - screeding, IPS, ceiling, drywalls\SCREEDING\WhatsApp extracts" "C:\Users\<you>\Downloads\KFHE exports"
```
The script prints the folder path and a summary of rows per sheet.

## Notes
- Messages are sorted by keyword (e.g. *cement, sand, fibre* → Materials; *grid, screed, WIR* → Progress;
  *payroll, BPR, advance* → Money). Check the sheets and move any row that landed in the wrong place.
- Qty/Unit come from the first number in the message (m2, m3, bags, kg, litres, RWF). The full original text is always kept in Detail.
- Photos are listed as "TO TRANSCRIBE" until the names or amounts are typed in.
