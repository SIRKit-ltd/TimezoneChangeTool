# Controls and settings

[Back to README](../README.md) · [Single meetings](non-recurring-meetings.md) · [Recurring meetings](recurring-meetings.md)

## Main window

| Control | What it does |
| --- | --- |
| **Province** | Selects the old/new time-zone pair. Changing it clears the scan and review acknowledgements; scan again. |
| **BC area** | Appears for British Columbia. Choose **Pacific time (most of BC)** or **Mountain time (aligned with Alberta)** according to the user's location. Changing it also clears the scan. |
| **Connect and scan / Scan calendars / Scan again** | Signs in when needed and reads calendar entries. The caption changes with connection state. Rescanning rebuilds the list; it does not write events. |
| **Change Outlook time zone…** | After a scan, asks to set the mailbox time zone to the selected replacement zone. It does not reschedule meetings. Outlook desktop/mobile may use separate settings. |
| **Fix** checkbox on a row | Selects an eligible single meeting for the next repair batch, up to 25. Disabled for ineligible entries and series. Row highlighting is separate from this checkbox. |
| **Add original time to event titles** | Optional, off by default. Adds a suffix such as `(11:00AM)` based on the original zone's clock time. Updates single-meeting previews and newly prepared series plans; it does not modify events merely by being checked. An identical existing suffix is not duplicated. Unchecking does not remove suffixes already saved in Outlook. |
| **Show excluded entries (past, already correct, other zones, or not organized by me)** | Shows excluded rows available in the scan for inspection. It never makes them editable. All-day events, skipped calendars, and handled rows remain omitted. |
| **I reviewed the selected events and want to preserve their original clock times…** | Required acknowledgement for single-meeting repair. Confirms the intended wall-clock time behavior, not preservation of the UTC instant. |
| **I understand meeting updates may be sent…** | Second required single-meeting acknowledgement. Recognizes that corrections and later restoration may notify attendees. |
| **Repair selected** | Requires both main acknowledgements, at least one eligible checked row, and final confirmation. Updates single meetings only. |
| **Adjust recurring series…** | Opens the migration preview for exactly one highlighted recurring-series row. Double-clicking that row does the same thing. Main-window acknowledgements are not required to open it. |

Changing the title option or excluded-entry filter refreshes the list and clears row checkboxes and the two main acknowledgements. Successful repairs also refresh the remaining list. Controls are disabled while work is running; wait for completion before closing.

## List columns and status

| Display | Meaning |
| --- | --- |
| **Event / Calendar** | Meeting title and owning calendar. The title can show the proposed time suffix. |
| **Before (new zone) / After (new zone)** | Existing and proposed start/end times, both expressed in the replacement zone. A recurring row represents an upcoming occurrence, not the whole series. |
| **Shift (min)** | Proposed change to the scheduled instant in minutes. A negative value means earlier. |
| **Review / reason** | Eligibility or review explanation. Hover over a row to read its reason. |
| **Original zone** | Source time-zone identifier for the entry. |
| **Resolve error** | Shows the manual-resolution action when the row has an error requiring review. |
| **Account and counts** | Connected account, visible/excluded counts, and applicable ownership or permission notes. |
| **Progress and status** | Current stage and completion/error messages. An animated bar without a percentage means the total is not yet known. |

## Recurring preview

| Control | What it does |
| --- | --- |
| **Dates** | Read-only sample of up to 100 dates, with existing and replacement times in the new zone. |
| **People and rooms** | Read-only attendee/resource list and location. |
| **Replacement description** | Editable plain text for a newly prepared Teams replacement. Read-only for non-online meetings and resumed migrations. |
| **Original description** | Read-only reference for checking notes that should be retained. |
| **I reviewed the dates, people, rooms and replacement description** | First required recurring-migration acknowledgement. |
| **I understand new invitations/cancellations will be sent…** | Second required acknowledgement, covering new series identity and possible Teams-link changes. |
| **Migrate future series** | Applies the reviewed plan after both acknowledgements. Known future exceptions/cancellations trigger a further Yes/No warning. |
| **Future edits and cancellations: Yes / No** | **Yes** authorizes recreating the regular future schedule, discarding individual future changes. **No** returns to preview without applying. |
| **Close preview** | Cancels the pending preview without applying it. |

There is no date-range picker or exception-override checkbox in the current series-migration flow. Review the calculated first date and retained end condition in the summary. An open-ended series remains open-ended; verification covers a finite horizon.

## Recovery controls

| Control | What it does |
| --- | --- |
| **Last error…** | Displays saved failure details when available. Treat diagnostic text as sensitive before sharing it. |
| **I have manually resolved this** | Asks you to confirm that you checked and resolved the issue in Outlook. Removes the row's active recovery tracking and hides related entries from future scans on this Windows account. Keeps encrypted backups and a local suppression marker. Does not verify or change Outlook. |
| **Clear journals…** | After confirmation, moves all active journals, including unfinished ones and manual-resolution markers, into a timestamped backup folder and clears the scan. Affects all accounts using this Windows journal folder. Does not undo changes or securely erase backups. Retrying afterward can duplicate replacements or notifications. Close other tool instances first. |

Active journals: `%LOCALAPPDATA%\AlbertaCalendarPilot\Journal`.
Clear-all backups: `%LOCALAPPDATA%\AlbertaCalendarPilot\JournalBackups`.
Manual-resolution backups: `%LOCALAPPDATA%\AlbertaCalendarPilot\Journal\ResolvedBackups`.
Backups are not automatically used for recovery. Keep them private.

## Configuration file

Most users should keep the supplied `config.json` unchanged. It contains identifiers, not credentials; never add passwords, tokens or client secrets.

| Setting | Default and purpose |
| --- | --- |
| `AuthenticationMode` | `MicrosoftDefault`: uses Microsoft's existing Graph Command Line Tools app and requests delegated user, calendar-write and mailbox-settings-write access together. `DedicatedApps` uses administrator-provided separate registrations and can require additional sign-ins. |
| `TenantId` | Blank in default mode allows an organizational sign-in. Set a tenant GUID to restrict the connection to that organization. Required in dedicated mode. |
| `ScanClientId` | Blank in default mode; ignored there. In dedicated mode, the administrator-provided scan application's client GUID. |
| `FixClientId` | Blank in default mode; ignored there. In dedicated mode, the repair/settings application's client GUID, which must differ from `ScanClientId`. |

The launcher requires PowerShell 7.2+ already installed. It installs Graph authentication if absent, but does not upgrade an existing version. Its execution-policy bypass lasts for its process and cannot override organization-enforced policy. Neither Windows time-zone updates nor tenant consent are installed automatically.

## Time-zone profiles and scan limits

These are the identifiers configured in this version, not Outlook display labels:

| Profile | Original | Replacement |
| --- | --- | --- |
| Alberta | Mountain Standard Time | Alberta Standard Time |
| British Columbia — Pacific | Pacific Standard Time | British Columbia Standard Time |
| British Columbia — Mountain | Mountain Standard Time | Alberta Standard Time |
| Manitoba | Central Standard Time | Manitoba Standard Time |

The GUI starts scanning from November 1, 2026 and excludes past entries. It does not offer a custom scan-date setting. Single-event inventory has no user-selected upper date limit. Recurring discovery and verification use bounded date windows; preview samples are not a complete list of all future occurrences. All-day events and calendars whose ownership cannot be established are excluded.
