# Timezone Change Tool

This tool helps migrate Outlook calendar items to the new Western Canadian time zones for Manitoba, Alberta, and British Columbia. It is designed for IT professionals using Exchange Online.

## Quickstart

1. **Prepare Windows.** Install PowerShell 7.2 or later and the applicable Windows time-zone updates. Confirm the new regional zone is available in Exchange Online. Test with a non-production mailbox before wider deployment.
2. **Extract the entire package** into one folder and close any older copies of the tool.
3. **Run `Setup-and-Launch.cmd`.** It installs Microsoft Graph authentication for the current user if missing, then opens the tool. Installation requires internet access. The launcher uses execution-policy bypass for this session only; organization-enforced policies still apply.
4. **Connect with the account whose calendar you want to update.** If Microsoft requests administrator approval, have your tenant administrator approve access before continuing.
5. **Select the province and scan.** The default start date is November 1, 2026. Review the remaining eligible meetings.
6. **Preview and apply.** Review single-meeting changes, or double-click a recurring series to open its migration preview. Optionally enable the title checkbox to append the original time, such as `(11:00AM)`. Confirm the changes you want to apply.
7. **Check the results in Outlook.** Updating the Outlook mailbox time-zone setting is a separate, optional action in the tool.

Sign-in, scanning, and previewing do not change events. If script execution is blocked by organizational policy, contact IT for approved signing or deployment.

## User guides

- [Non-recurring meetings: step-by-step](docs/non-recurring-meetings.md)
- [Recurring meetings: preview, migration and recovery](docs/recurring-meetings.md)
- [Every checkbox, button and setting](docs/controls-and-settings.md)

## What to expect

- The tool preserves the original clock time in the new zone: 11:00 AM stays 11:00 AM, although its UTC instant changes. Outlook displaying the old zone may show a different time.
- By default, past and already-correct entries are hidden. Only meetings you organize in calendars you own are eligible. All-day events are excluded.
- Recurring meetings are migrated by creating a replacement for future occurrences and ending or removing the original as appropriate. This can send invitations and cancellations, change Teams links, and lose description formatting. Review each preview carefully.
- If you confirm recreation of future exceptions, cancelled dates may return and custom occurrence details may be reset. Unsupported series can still be blocked. Room invitations do not guarantee acceptance.
- A migration can stop partway through. Run only one instance per mailbox and retain journals after an error.

## Permissions

No new Entra app registration is required with the default configuration. The tool uses Microsoft's shared Graph Command Line Tools application with delegated `User.Read`, `Calendars.ReadWrite`, and `MailboxSettings.ReadWrite` permissions. Tenant administrator consent may be required, and existing shared-app consent can include additional scopes. The tool reuses sign-in during the session, subject to tenant policies.

## **After completion: remove access you no longer need**

1. **Close the tool once all migrations and recovery work are complete.** Closing it or uninstalling Graph does not revoke consent.
2. **Have your tenant administrator review the granted permissions** in Entra admin center → Enterprise apps → All applications → Microsoft Graph Command Line Tools → Permissions. Verify application ID `14d82eec-204b-4c2f-b7e8-296a70dab67e`. Revoke consent granted for this migration that is no longer needed, checking both admin and user consent. User grants may require Graph or PowerShell; see Microsoft's [revocation instructions](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-application-permissions).
3. **Remove migration-only enterprise applications when no longer needed.** If you used dedicated registrations, also remove those registrations when appropriate. **Do not delete the shared Microsoft Graph Command Line Tools enterprise application or revoke shared grants without checking other dependencies:** this can disrupt other users, scripts, and administration tools across the tenant.

## If something goes wrong

Keep the journal and inspect the original and replacement meetings in Outlook before retrying. Encrypted journals are stored at:

`%LOCALAPPDATA%\AlbertaCalendarPilot\Journal`

Use **I have manually resolved this** only after resolving the issue in Outlook; it suppresses the entry locally without verifying server state. **Clear journals** moves active records to backups and resets tracking; it does not undo changes. Clearing records can allow duplicate meetings or repeated invitations on retry.

Do not post journals or unredacted calendar details publicly. Ask a maintainer for a private channel before sharing sensitive diagnostic information.

## Disclaimer

Use this tool at your own risk. It is provided "as is," without warranties. To the maximum extent permitted by applicable law, SIRKit is not liable for any loss, damage, or disruption arising from its use, including altered or lost calendar data, unintended invitations or cancellations, scheduling errors, or business interruption. You are responsible for reviewing changes, testing, and maintaining appropriate backups and recovery procedures.

## License

Code and documentation use the [MIT License](LICENSE), including its warranty disclaimer and limitation of liability. SIRKit branding is excluded; see [NOTICE](NOTICE). This is an independent utility, not a Microsoft product.
