# Non-recurring meetings

[Back to README](../README.md) · [Controls and settings](controls-and-settings.md) · [Recurring meetings](recurring-meetings.md)

Use this workflow for single appointments and meetings. Recurring series use the separate migration window.

## Step by step

1. Launch `Setup-and-Launch.cmd`, choose the **Province** and, for BC, the correct **BC area**. Click **Connect and scan** and sign in as the calendar owner. Check the connected account before proceeding.
2. Wait for **Scan complete**. The scan starts from November 1, 2026 and excludes past events. Only eligible meetings you organize in calendars you own can be changed. All-day events are omitted.
3. Decide whether to enable **Add original time to event titles**. For example, a meeting originally at 11:00 AM can become `Team meeting (11:00AM)`. Changing this option refreshes the preview and clears selections and acknowledgements, so set it before selecting meetings.
4. Review **Event**, **Calendar**, **Original zone**, and **Review / reason**. Both **Before (new zone)** and **After (new zone)** use the replacement zone. For example, an original 11:00 AM meeting might show noon before correction and 11:00 AM afterward. The goal is to preserve the original clock time in the new zone.
5. Tick the **Fix** checkbox beside each eligible single meeting. Select no more than **25 per batch**. Highlighting a row alone does not select it for repair. A disabled checkbox means the row cannot use this workflow; read its reason.
6. Tick both acknowledgements below the list: the one about preserving original clock times and the one about sending meeting updates.
7. Click **Repair selected**. Review the account, region and event count in the confirmation, then choose **Yes** to apply or **No** to cancel. Calendar writes begin only after confirmation.
8. Wait for verification. Completed entries disappear and the remaining entries stay available. Review and select the next batch, then tick the acknowledgements again. A full rescan is normally unnecessary.
9. Check the results in Outlook using the intended display time zone. If needed, use **Change Outlook time zone…** separately; it changes the mailbox setting, not the scheduled times of existing events.

Meeting updates may be sent to attendees. The optional title is included in the same correction. Editing a title already changed by an earlier run is not an undo operation.

## If a meeting is missing or cannot be selected

Enable **Show excluded entries** to see available excluded rows and their reasons. This does not make them eligible. All-day events, skipped shared calendars, and completed or manually resolved entries remain omitted. A scan is not an inventory of every calendar you can access.

Changing the province or BC area clears the current scan. Scan again before applying changes. A single-meeting preview expires after 30 minutes; if prompted, rescan and review the new results.

## If a batch stops

Read the error and use **Last error…** for saved diagnostic details. Earlier verified changes can already be complete; the batch is not automatically rolled back. Completed rows disappear, the failed row requires review, and unattempted rows can remain available.

Keep the encrypted journal at `%LOCALAPPDATA%\AlbertaCalendarPilot\Journal`. Check the affected event in Outlook before retrying. If you resolve it manually, click **I have manually resolved this** on that row and confirm. This changes local recovery tracking only; it does not check, repair or undo anything in Outlook.

Do not use **Clear journals…** as a routine retry step. It forgets active tracking and can allow duplicate operations. See [recovery controls](controls-and-settings.md#recovery-controls).
