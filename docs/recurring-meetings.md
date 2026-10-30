# Recurring meetings

[Back to README](../README.md) · [Controls and settings](controls-and-settings.md) · [Non-recurring meetings](non-recurring-meetings.md)

The tool creates a replacement series for future dates in the new time zone. After verifying it, the tool ends the original before the replacement starts, or removes the original if the entire series is future. It then invites attendees to the replacement. This is not Outlook's native “this and all future events” command.

## Step by step

1. Launch the tool, select the correct **Province** and **BC area** if applicable, then **Connect and scan**. Check the connected account.
2. Set **Add original time to event titles** if wanted before opening the series preview. The replacement uses the regular series' original clock time in its title. A resumed migration uses its saved plan.
3. **Double-click the recurring-series row**, or highlight exactly one series and click **Adjust recurring series…**. Do not use the main list's **Fix** checkboxes or **Repair selected** for a series.
4. In **Migrate future recurring series**, review the summary: account, replacement start date, end condition, destination zone and what will happen to the original. Earlier meetings remain in the original series when it is shortened.
5. Review the **Dates** tab. Both columns use the new zone. **Before (new zone)** shows the existing schedule in that zone; **Replacement (new zone)** shows the proposed schedule. The window displays up to 100 sample dates. The replacement keeps the series end condition, including no end date; the samples do not limit how many future dates will be replaced.
6. Review **People and rooms**. Check attendees, resource addresses and location. This tab is read-only. Rooms must accept the replacement invitation; the tool does not guarantee availability or acceptance.
7. Review **Replacement description** against **Original description**. For a new Teams migration, edit the replacement's plain text as needed to retain useful notes. The old join block is removed and a new Teams link is generated; formatting and embedded images are not preserved. Non-online descriptions are copied with their formatting, and their preview is read-only. Resumed migration descriptions are also read-only.
8. Tick both acknowledgements: **I reviewed the dates, people, rooms and replacement description** and **I understand new invitations/cancellations will be sent and the series identity and Teams link can change**.
9. Click **Migrate future series**. If the **Future edits and cancellations** warning appears, read it before choosing **Yes**. Choose **No** to return to the preview without starting migration, or close the preview to cancel.
10. Wait for verification. On success, the completed series disappears and you can open the next series without rescanning. Check the replacement and original in Outlook, including dates, description, attendees, rooms and joining details.

## Future exceptions and cancellations

If you proceed with the warning, cancelled future dates can be recreated. Individually edited future occurrences revert to the regular series' times, duration, notes, attendees and locations. For open-ended series, this applies beyond the sample dates shown. There is no separate “ignore exceptions” checkbox.

Earlier identifiable edits and cancellations remain in the original series. Unidentifiable dates, unsupported exceptions, attachments, or other unsupported configurations can still block migration. Handle those cases in Outlook; the warning is not an override for every error. All-day series are excluded from the list.

## Invitations and interruption

The replacement is prepared before attendees are added. New invitations and old-series updates/cancellations can still produce multiple emails. Attendee responses reset, and Teams joining details can change.

The migration is several server operations, not one transaction. An error can leave a replacement, an ended original, or another partially completed state. Do not create another replacement manually or clear journals just to retry.

Keep `%LOCALAPPDATA%\AlbertaCalendarPilot\Journal`, inspect Outlook, and reopen the same row for a fresh preview. When the journal identifies an existing replacement, the preview indicates a resume rather than a new creation. If the tool cannot safely reconcile the state, stop and investigate the error. Use only one instance for the mailbox, including across different computers.

A recurring preview expires after 30 minutes. Reopen the series to refresh it; a full calendar rescan is not normally required. If you resolve the problem in Outlook yourself, use **I have manually resolved this** only after checking both the original and replacement. See [recovery controls](controls-and-settings.md#recovery-controls).
