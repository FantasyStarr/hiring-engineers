# Calendar Permission Manager

Local Windows GUI (PowerShell + WPF) for bulk-assigning calendar delegate permissions in a client's Microsoft 365 tenant.
It is the GUI equivalent of looping `Add-MailboxFolderPermission -Identity "<mbx>:\Calendar" -User <delegate> -AccessRights <rights>`.

## Run
```powershell
powershell -ExecutionPolicy Bypass -File .\CalendarPermManager.ps1
```
First run offers to install `ExchangeOnlineManagement` (current user scope).

## Workflow
1. **Connect** – opens Microsoft's own sign-in window (MFA / Conditional Access work normally). The app never sees the password. Optionally enter the admin UPN to pre-fill.
2. **Load mailboxes**, then filter by display name / UPN, **Active users only** (account enabled), and mailbox type.
   Shared/room/equipment mailboxes show `Active = n/a` (their accounts are disabled by design) and are not excluded by the active filter.
3. Check mailboxes, pick a **delegate** and **rights**, click *Add checked mailboxes to queue* – or *Import CSV* with columns `Mailbox,Delegate,AccessRights`.
4. Leave **Preview only (-WhatIf)** ticked for a dry run, then untick and **Apply queue**.

Behaviour notes
- If the delegate already has an entry on the calendar, `Set-MailboxFolderPermission` is used instead of Add (avoids the "existing entry" error).
- The calendar folder is looked up per mailbox, so non-English tenants (`Kalender`, `Calendrier`…) work.
- Every run appends to `CalendarPerm-<timestamp>.csv` next to the script.
- Requires an account with Exchange admin rights in the client tenant.

Not tested on a live tenant yet (built on Linux) – run a `-WhatIf` pass on a test mailbox first.
