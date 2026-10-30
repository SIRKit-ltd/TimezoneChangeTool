#requires -Version 7.2
# Offline tests only. Graph commands are replaced by in-memory fakes inside this module instance.
[CmdletBinding()]
param([string]$TestDirectory = (Join-Path $env:TEMP ('AlbertaCalendarPilotTests-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
$module = Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
& $module {
    param($TestDirectory)
    $script:checks = 0
    function Assert($Condition, [string]$Message) {
        if (-not $Condition) { throw "FAILED: $Message" }
        $script:checks++
        Write-Output "PASS: $Message"
    }
    function Expect-Throw([scriptblock]$Action, [string]$Message) {
        $threw = $false
        try { & $Action | Out-Null } catch { $threw = $true }
        Assert $threw $Message
    }
    function Make-Event {
        @{
            id = [guid]::NewGuid().ToString(); subject = 'Synthetic test event'; changeKey = 'v1'; '@odata.etag' = 'W/"v1"'
            type = 'singleInstance'; isOrganizer = $true; isAllDay = $false; isCancelled = $false
            originalStartTimeZone = 'Mountain Standard Time'; originalEndTimeZone = 'Mountain Standard Time'
            start = @{ dateTime='2099-01-10T16:00:00.0000000'; timeZone='UTC' }
            end = @{ dateTime='2099-01-10T17:00:00.0000000'; timeZone='UTC' }
        }
    }
    $script:testZone = [TimeZoneInfo]::CreateCustomTimeZone('Test Alberta',[timespan]::FromHours(-6),'Test Alberta','Test Alberta')
    $now = [datetimeoffset]'2026-10-29T00:00:00Z'
    function Propose($e) { New-ABProposal $e $script:testZone 'Test Alberta' $now }
    $p = Propose (Make-Event)
    Assert ($p.Eligible -and $p.ShiftMinutes -eq -60 -and $p.Patch.start.dateTime.StartsWith('2099-01-10T09:00:00')) 'Winter correction preserves 09:00 local time and moves UTC earlier by one hour'
    Assert (-not $p.Selected) 'No candidate is selected by default'
    Assert $p.DefaultVisible 'Future organized event in the old zone is visible by default'
    Assert (($p.Patch.Keys | Sort-Object) -join ',' -eq 'end,start') 'Patch contains only start and end'
    $titled=New-ABProposal (Make-Event) $script:testZone 'Test Alberta' $now -IncludeTimeInTitle
    Assert ($titled.Subject -eq 'Synthetic test event (09:00AM)' -and $titled.Patch.subject -eq $titled.Subject) 'Title option uses the original zone wall time, not the shifted display time'
    Assert ((Add-TZTimeToTitle 'Meeting' ([datetime]'2027-01-01T11:00:00')) -eq 'Meeting (11:00AM)') 'Title time format matches the requested example'
    Assert ((Add-TZTimeToTitle 'Meeting' ([datetime]'2027-01-01T00:00:00')) -eq 'Meeting (12:00AM)' -and (Add-TZTimeToTitle 'Meeting' ([datetime]'2027-01-01T12:00:00')) -eq 'Meeting (12:00PM)') 'Midnight and noon title labels are correct'
    Assert ((Add-TZTimeToTitle 'Meeting (11:00AM)' ([datetime]'2027-01-01T11:00:00')) -eq 'Meeting (11:00AM)') 'Existing identical time suffix is not duplicated'
    $e = Make-Event; $e.start.dateTime='2099-07-10T15:00:00'; $e.end.dateTime='2099-07-10T16:00:00'
    $p = Propose $e
    Assert ($p.Eligible -and $p.ShiftMinutes -eq 0) 'Summer event changes zone without shifting its UTC time'
    foreach ($kind in 'occurrence','exception','seriesMaster') {
        $e = Make-Event; $e.type=$kind
        Assert (-not (Propose $e).Eligible) "Recurring $kind is excluded from repair"
    }
    $e = Make-Event; $e.isOrganizer=$false
    Assert (-not (Propose $e).Eligible) 'Attendee copy is excluded'
    Assert (-not (Propose $e).DefaultVisible) 'Attendee copy is hidden by the default view'
    $e = Make-Event; $e.isAllDay=$true
    Assert (-not (Propose $e).Eligible) 'All-day event with non-midnight boundaries is excluded'
    foreach($kind in 'singleInstance','seriesMaster','occurrence','exception'){
        $e=Make-Event;$e.isAllDay=$true;$e.type=$kind;$e.start.dateTime='2099-01-10T07:00:00';$e.end.dateTime='2099-01-14T07:00:00'
        $p=Propose $e
        Assert (-not $p.Eligible -and -not $p.DefaultVisible -and $null -eq $p.Patch) "All-day $kind is excluded without a proposed change"
    }
    $e = Make-Event; $e.isCancelled=$true
    Assert (-not (Propose $e).Eligible) 'Cancelled event is excluded'
    $e = Make-Event; $e.originalStartTimeZone='Eastern Standard Time'
    Assert (-not (Propose $e).Eligible) 'Mixed or non-Mountain source zones are excluded'
    $e = Make-Event; $e.originalStartTimeZone='America/Edmonton'; $e.originalEndTimeZone='America/Edmonton'
    Assert (-not (Propose $e).Eligible) 'IANA Edmonton rules are not guessed'
    $e = Make-Event; $e.originalStartTimeZone='Test Alberta'; $e.originalEndTimeZone='Test Alberta'
    Assert (-not (Propose $e).Eligible) 'Already migrated event is excluded'
    Assert (-not (Propose $e).DefaultVisible) 'Already-correct event is hidden by default'
    $e = Make-Event; $e.Remove('@odata.etag')
    Assert (-not (Propose $e).Eligible) 'Missing ETag prevents repair'
    $e = Make-Event; $e.start.timeZone='Mountain Standard Time'
    Assert (-not (Propose $e).Eligible) 'Non-UTC response fails closed'
    $e = Make-Event; $e.start.dateTime='nonsense'
    Assert (-not (Propose $e).Eligible) 'Malformed date fails closed'
    $e = Make-Event; $e.start.dateTime='2026-11-01T07:00:00'; $e.end.dateTime='2026-11-01T07:30:00'
    Assert (-not (Propose $e).Eligible) 'Event before the actual fall-back boundary is excluded'
    $e = Make-Event; $e.start.dateTime='2026-11-01T07:30:00'; $e.end.dateTime='2026-11-01T09:00:00'
    Assert (-not (Propose $e).Eligible) 'Transition-spanning event is excluded'
    $e = Make-Event
    $past = New-ABProposal $e $script:testZone 'Test Alberta' ([datetimeoffset]'2099-01-10T15:30:00Z')
    Assert (-not $past.Eligible) 'Proposed start in the past is excluded'
    $old=Make-Event; $old.start.dateTime='2026-12-01T16:00:00'; $old.end.dateTime='2026-12-01T17:00:00'
    Assert (-not (New-ABProposal $old $script:testZone 'Test Alberta' ([datetimeoffset]'2027-01-01T00:00:00Z')).DefaultVisible) 'Past entries after November 1 are also hidden by default'
    Expect-Throw { Assert-ABUri 'https://attacker.invalid/v1.0/me/events' } 'Off-host request is blocked'
    Expect-Throw { Assert-ABUri 'https://graph.microsoft.com/v1.0/users/another-user/events' } 'Other-user endpoint is blocked'
    Expect-Throw { Assert-ABUri 'https://graph.microsoft.com:444/v1.0/me/events' } 'Unexpected port is blocked'
    Assert-ABUri 'https://graph.microsoft.com/v1.0/me/calendarView?$top=250'
    Assert $true 'Own default calendar endpoint is accepted'

    $profiles=@(
        @('Alberta','Mountain Standard Time','Alberta Standard Time',-6,'2026-11-01T08:00:00Z'),
        @('British Columbia','Pacific Standard Time','British Columbia Standard Time',-7,'2026-11-01T09:00:00Z'),
        @('British Columbia (Mountain)','Mountain Standard Time','Alberta Standard Time',-6,'2026-11-01T08:00:00Z'),
        @('Manitoba','Central Standard Time','Manitoba Standard Time',-5,'2026-11-01T07:00:00Z')
    )
    foreach($case in $profiles) {
        $definition=Get-TZProfile $case[0]
        Assert ($definition.OldId -eq $case[1] -and $definition.NewId -eq $case[2]) "$($case[0]): correct old/new Windows identifiers"
        Assert ($definition.Cutover -eq [datetimeoffset]$case[4]) "$($case[0]): cutover is 02:00 local time in its own offset"
        Assert ($definition.Start.Hour -eq 0 -and $definition.Start.Offset.TotalHours -eq $case[3]) "$($case[0]): inventory starts at provincial midnight"
        $zone=[TimeZoneInfo]::CreateCustomTimeZone($definition.NewId,[timespan]::FromHours($case[3]),$definition.Name,$definition.Name)
        $event=Make-Event; $event.originalStartTimeZone=$definition.OldId; $event.originalEndTimeZone=$definition.OldId
        $start=[datetimeoffset]::new([datetime]'2099-01-10T09:00:00',[timespan]::FromHours($case[3]-1))
        $event.start.dateTime=$start.UtcDateTime.ToString('s'); $event.end.dateTime=$start.AddHours(1).UtcDateTime.ToString('s')
        $proposal=New-ABProposal $event $zone $definition.NewId $now $definition.Key
        Assert ($proposal.Eligible -and $proposal.ShiftMinutes -eq -60 -and $proposal.Patch.start.dateTime.StartsWith('2099-01-10T09:00:00') -and $proposal.Patch.start.timeZone -eq $definition.NewId) "$($case[0]): winter proposal preserves 09:00 in the correct target zone"
        $start=[datetimeoffset]::new([datetime]'2099-07-10T09:00:00',[timespan]::FromHours($case[3]))
        $event.start.dateTime=$start.UtcDateTime.ToString('s'); $event.end.dateTime=$start.AddHours(1).UtcDateTime.ToString('s')
        $proposal=New-ABProposal $event $zone $definition.NewId $now $definition.Key
        Assert ($proposal.Eligible -and $proposal.ShiftMinutes -eq 0) "$($case[0]): summer conversion changes the zone without a UTC shift"
        $event.originalStartTimeZone='UTC'; $event.originalEndTimeZone='UTC'
        Assert (-not (New-ABProposal $event $zone $definition.NewId $now $definition.Key).Eligible) "$($case[0]): unrelated source zone cannot become a candidate"
        $event=Make-Event; $event.originalStartTimeZone=$definition.OldId; $event.originalEndTimeZone=$definition.OldId
        $event.start.dateTime=$definition.Cutover.AddMinutes(-30).UtcDateTime.ToString('s')
        $event.end.dateTime=$definition.Cutover.UtcDateTime.ToString('s')
        Assert (-not (New-ABProposal $event $zone $definition.NewId $now $definition.Key).Eligible) "$($case[0]): entry before that province's transition cannot be repaired"
        $demo=Get-ABDemo $definition.Key
        Assert ($demo.Profile -eq $definition.Key -and $demo.Target -eq $definition.NewId -and $demo.Rows[0].Eligible) "$($case[0]): demo uses the selected regional pair"
    }
    Expect-Throw { Get-TZProfile 'Ontario' } 'Unsupported province is rejected'

    $configPath=Join-Path $TestDirectory 'config-test.json'
    New-Item -ItemType Directory -Path $TestDirectory -Force | Out-Null
    [IO.File]::WriteAllText($configPath,'{"AuthenticationMode":"MicrosoftDefault","TenantId":"","ScanClientId":"","FixClientId":""}')
    $defaultConfig=Read-ABConfig $configPath
    Assert ($defaultConfig.AuthenticationMode -eq 'MicrosoftDefault') 'Default connection requires no tenant or application IDs'
    $boundary=[datetimeoffset]'2026-11-01T00:00:00-06:00'
    $e=Make-Event; $e.start.dateTime='2040-01-10T16:00:00'; $e.end.dateTime='2040-01-10T17:00:00'
    Assert (Test-ABInventoryDate $e $boundary) 'Inventory has no upper date cutoff'
    $e=Make-Event; $e.start.dateTime='2026-10-31T15:00:00'; $e.end.dateTime='2026-10-31T16:00:00'
    Assert (-not (Test-ABInventoryDate $e $boundary)) 'One-off event before November 1 is excluded from the inventory'
    $e=Make-Event; $e.start.dateTime='2026-11-01T05:30:00'; $e.end.dateTime='2026-11-01T06:30:00'
    Assert (Test-ABInventoryDate $e $boundary) 'Event overlapping November 1 midnight remains visible'
    $e=Make-Event; $e.type='seriesMaster'; $e.start.dateTime='2020-01-01T15:00:00'; $e.end.dateTime='2020-01-01T16:00:00'
    Assert (Test-ABInventoryDate $e $boundary) 'Older recurring master remains visible for future-occurrence review'

    # Exercise the real authentication policy with SDK calls replaced by fakes.
    function script:Import-Module { [CmdletBinding()]param([string]$Name) }
    function script:Disconnect-MgGraph { }
    function script:Connect-MgGraph {
        param($TenantId,$Scopes,$ContextScope,[switch]$NoWelcome,$ClientId)
        $script:authArguments=@{TenantId=$TenantId;Scopes=$Scopes;ContextScope=$ContextScope;ClientId=$ClientId}
        $script:authCalls++
    }
    function script:Get-MgContext { $script:authContext }
    function script:Invoke-ABGet($Path) {
        if ($Path -eq '/me/outlook/supportedTimeZones') { return @{value=$script:fakeZones} }
        @{id='auth-user';userPrincipalName='auth@example.invalid'}
    }
    $script:authCalls=0
    $script:authContext=@{AuthType='Delegated';TenantId='11111111-1111-1111-1111-111111111111';ClientId='14d82eec-204b-4c2f-b7e8-296a70dab67e';Scopes=@('User.Read','Calendars.Read')}
    $progressQueue=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $auth=Connect-ABGraph $defaultConfig Scan $progressQueue
    Assert (@($progressQueue.ToArray() | Where-Object Stage -eq 'SignedIn').Count -eq 1) 'Successful authentication immediately emits a sign-in-complete notification'
    Assert ($script:authArguments.TenantId -eq 'organizations' -and -not $script:authArguments.ClientId) 'Default connection selects Microsoft app and organization sign-in'
    Assert ('Calendars.ReadWrite' -in $script:authArguments.Scopes -and 'MailboxSettings.ReadWrite' -in $script:authArguments.Scopes -and $script:authArguments.ContextScope -eq 'Process') 'First sign-in requests combined permissions in a process-scoped session'
    $script:authContext.Scopes=@('User.Read','Calendars.Read','Calendars.ReadWrite')
    $auth=Connect-ABGraph $defaultConfig Scan
    Assert ([bool]$auth.AccessNote) 'Prior write consent on Microsoft shared app is disclosed'
    $auth=Connect-ABGraph $defaultConfig Fix
    Assert ('Calendars.ReadWrite' -in $script:authArguments.Scopes) 'Confirmed repair requests calendar-write access'
    $script:authContext.Scopes=@('User.Read','Calendars.Read','Directory.Read.All','Mail.ReadWrite')
    $auth=Connect-ABGraph $defaultConfig Scan
    Assert ($auth.UserId -eq 'auth-user' -and $auth.AccessNote -match '2 additional existing permissions') 'Shared Microsoft app can connect with unrelated existing grants and discloses them'
    Assert (($script:authArguments.Scopes | Sort-Object) -join ',' -eq 'Calendars.ReadWrite,MailboxSettings.ReadWrite,User.Read') 'Combined sign-in never requests unrelated existing grants'
    $script:authContext.Scopes=@('User.Read','Calendars.ReadWrite','Directory.Read.All')
    $auth=Connect-ABGraph $defaultConfig Scan
    Assert ($auth.AccessNote -match 'scan only reads' -and $auth.AccessNote -match '1 additional existing permissions') 'A shared app with prior write and directory consent can still run a read-only scan'
    $auth=Connect-ABGraph $defaultConfig Fix
    Assert (($script:authArguments.Scopes | Sort-Object) -join ',' -eq 'Calendars.ReadWrite,MailboxSettings.ReadWrite,User.Read') 'Repair uses the combined app permissions without requesting unrelated grants'
    $script:authContext.Scopes=@('https://graph.microsoft.com/User.Read','https://graph.microsoft.com/Calendars.Read')
    $auth=Connect-ABGraph $defaultConfig Scan
    Assert ($auth.UserId -eq 'auth-user' -and -not $auth.AccessNote) 'Graph-qualified calendar scope names are recognized without a false permissions error'
    $script:authContext.Scopes=@('User.Read','Directory.Read.All')
    $missingMessage=''
    try { Connect-ABGraph $defaultConfig Scan | Out-Null } catch { $missingMessage=$_.Exception.Message }
    Assert ($missingMessage -match 'required Calendars.Read calendar access') 'Missing calendar access produces a specific calendar-consent error'
    $script:authContext.Scopes=@('User.Read','Calendars.Read')
    Expect-Throw { Connect-ABGraph $defaultConfig Fix } 'Calendar-read access cannot authorize repair'
    Expect-Throw { Connect-ABGraph $defaultConfig Settings } 'Calendar access alone cannot authorize an Outlook preference change'
    $script:authContext.Scopes=@('User.Read','MailboxSettings.ReadWrite')
    $auth=Connect-ABGraph $defaultConfig Settings
    Assert (($script:authArguments.Scopes | Sort-Object) -join ',' -eq 'Calendars.ReadWrite,MailboxSettings.ReadWrite,User.Read') 'Mailbox settings use the same combined sign-in permissions'
    $script:authContext.Scopes=@('User.Read','Calendars.ReadWrite','MailboxSettings.ReadWrite')
    $callsBefore=$script:authCalls
    $auth=Connect-ABGraph $defaultConfig Scan
    $auth=Connect-ABGraph $defaultConfig Fix
    $auth=Connect-ABGraph $defaultConfig Settings
    Assert ($script:authCalls -eq $callsBefore+1) 'Scanning, repairing and settings reuse one authentication request per session'
    $dedicatedConfig=@{AuthenticationMode='DedicatedApps';TenantId='11111111-1111-1111-1111-111111111111';ScanClientId='22222222-2222-2222-2222-222222222222';FixClientId='33333333-3333-3333-3333-333333333333'}
    $script:authContext.ClientId=$dedicatedConfig.ScanClientId
    $script:authContext.Scopes=@('User.Read','Calendars.Read','Directory.Read.All')
    Expect-Throw { Connect-ABGraph $dedicatedConfig Scan } 'Dedicated scan app still rejects unrelated permissions'
    $script:authContext.Scopes=@('User.Read','Calendars.Read','Calendars.ReadWrite')
    Expect-Throw { Connect-ABGraph $dedicatedConfig Scan } 'Dedicated scan app still rejects calendar-write access'
    $script:authContext.Scopes=@('User.Read','Calendars.Read')
    $auth=Connect-ABGraph $dedicatedConfig Scan
    Assert ($auth.UserId -eq 'auth-user' -and $script:authArguments.ClientId -eq $dedicatedConfig.ScanClientId) 'Dedicated scan with its intended scopes continues to work'
    $script:authContext.ClientId=$dedicatedConfig.FixClientId
    $script:authContext.Scopes=@('User.Read','Calendars.ReadWrite','Directory.Read.All')
    Expect-Throw { Connect-ABGraph $dedicatedConfig Fix } 'Dedicated repair app still rejects unrelated permissions'
    $script:authContext.Scopes=@('User.Read','Calendars.Read')
    $script:authContext.ClientId='unexpected-app'
    Expect-Throw { Connect-ABGraph $defaultConfig Scan } 'Unexpected client app is rejected'
    $script:fakeZones=@(
        @{alias='Alberta Standard Time';displayName='(UTC-06:00) Mountain Time (Alberta)'},
        @{alias='British Columbia Standard Time';displayName='(UTC-07:00) Pacific Time (British Columbia)'},
        @{alias='Manitoba Standard Time';displayName='(UTC-05:00) Manitoba'})
    foreach($case in $profiles) {
        $target=Get-ABTarget $case[0]
        Assert ($target.Alias -eq $case[2]) "$($case[0]): mailbox zone is resolved by exact Windows alias"
    }
    $script:fakeZones=@(@{alias='Alberta Standard Time';displayName='British Columbia misleading label'})
    Expect-Throw { Get-ABTarget 'British Columbia' } 'Wrong alias is rejected even when its display name contains the selected province'

    New-Item -ItemType Directory -Path $TestDirectory -Force | Out-Null
    $identity = @{ UserId='test-user'; TenantId='test-tenant'; Account='test@example.invalid' }
    $ownerIdentity=@{Account='login@example.invalid';Mail='primary@example.invalid'}
    Assert (Test-TZCalendarOwned @{owner=@{address='PRIMARY@example.invalid'};canEdit=$true} $ownerIdentity) 'Calendar ownership recognizes a primary email different from the sign-in address'
    Assert (Test-TZCalendarOwned @{owner=@{address='login@example.invalid'}} $ownerIdentity) 'Calendar ownership recognizes the sign-in address'
    Assert (-not(Test-TZCalendarOwned @{owner=@{address='delegate@example.invalid'};canEdit=$true} $ownerIdentity)) 'Editable delegated calendar is not treated as owned'
    Assert (-not(Test-TZCalendarOwned @{canEdit=$true} $ownerIdentity)) 'Missing calendar owner fails closed'
    Assert (-not(Test-TZCalendarOwned @{owner=@{address=''};canEdit=$true} $ownerIdentity)) 'Empty calendar owner fails closed'
    $journalPath = Join-Path $TestDirectory 'roundtrip.dpapi'
    Save-ABJournal @{ Identity=$identity; Items=@(@{Id='test-id';Status='Verified'}); Secret='synthetic-sensitive-title' } $journalPath
    Assert (-not (Get-Content -LiteralPath $journalPath -Raw).Contains('synthetic-sensitive-title')) 'Journal is encrypted on disk'
    $index = Get-ABJournalIndex $identity $TestDirectory
    Assert ($index['test-id'] -eq 'Verified') 'Encrypted journal can be read by this Windows user'
    $largeText=('Synthetic calendar content — été 🍁 ' * 8000)
    $largePath=Join-Path $TestDirectory 'large-roundtrip.dpapi'
    Save-ABJournal @{Identity=$identity;Items=@();Content=$largeText} $largePath
    $roundtrip=Read-ABJournal $largePath
    Assert ($largeText.Length -gt 113204 -and $roundtrip.Content -ceq $largeText) 'Large Unicode journals exceed the reported size and round-trip without truncation'
    $encryptedText=[IO.File]::ReadAllText($largePath)
    Assert ($encryptedText.StartsWith('TZJ2:') -and -not $encryptedText.Contains('Synthetic calendar content')) 'Large journal uses current-user DPAPI without plaintext content'
    $legacyPath=Join-Path $TestDirectory 'legacy-format.dpapi'
    (@{Identity=$identity;Items=@(@{Id='legacy-format';Status='Verified'})}|ConvertTo-Json -Depth 5)|ConvertTo-SecureString -AsPlainText -Force|ConvertFrom-SecureString|Set-Content -LiteralPath $legacyPath
    Assert ((Read-ABJournal $legacyPath).Items[0].Id -eq 'legacy-format') 'Previous SecureString journal format remains readable'
    Assert ((Get-ABJournalIndex $identity $TestDirectory).ContainsKey('legacy-format')) 'Mixed old and new journal formats work in the same directory'
    $other = Get-ABJournalIndex @{ UserId='other';TenantId='test-tenant' } $TestDirectory
    Assert ($other.Count -eq 0) 'Journal lookup separates signed-in mailbox identities'

    # All functions below are fakes. No Graph module is imported and no network request is made.
    function script:Connect-ABGraph($Config,$Mode) { return $script:fakeIdentity }
    function script:Disconnect-MgGraph { }
    $script:regionalTarget=$false
    function script:Get-ABTarget($Profile='Alberta') {
        if ($script:regionalTarget) {
            $definition=Get-TZProfile $Profile
            $zone=[TimeZoneInfo]::CreateCustomTimeZone($definition.NewId,[timespan]::FromHours($definition.OffsetHours),$definition.Name,$definition.Name)
            $script:currentTarget=[pscustomobject]@{Alias=$definition.NewId;Zone=$zone}
        } else { $script:currentTarget=[pscustomobject]@{Alias='Test Alberta';Zone=$script:testZone} }
        return $script:currentTarget
    }
    function script:Invoke-ABGet($Path) { return $script:serverEvent }
    function script:Invoke-MgGraphRequest {
        param($Method,$Uri,$Headers,$Body,$ContentType,$OutputType)
        if ($Method -ne 'PATCH') { throw 'Fake only permits PATCH' }
        $script:writes++
        if ($script:simulateFailure) { throw 'Synthetic uncertain network outcome' }
        if ($Headers['If-Match'] -ne 'W/"v1"') { throw 'Missing precondition' }
        $patch = $Body | ConvertFrom-Json -AsHashtable
        $allowed=if($script:serverEvent.isAllDay){'end,isAllDay,start'}else{'end,start'}
        if($patch.ContainsKey('subject')){$allowed+=',subject';$script:serverEvent.subject=$patch.subject}
        if (($patch.Keys | Sort-Object) -join ',' -ne $allowed) { throw 'Unexpected payload' }
        if($script:serverEvent.isAllDay -and $patch.isAllDay -ne $true){throw 'All-day flag must be retained'}
        foreach($field in 'start','end') {
            $local=[datetime]::Parse($patch[$field].dateTime,[Globalization.CultureInfo]::InvariantCulture)
            $utc=[datetimeoffset]::new($local,$script:currentTarget.Zone.GetUtcOffset($local))
            $script:serverEvent[$field].dateTime=$utc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff')
        }
        $script:serverEvent.originalStartTimeZone=$patch.start.timeZone; $script:serverEvent.originalEndTimeZone=$patch.end.timeZone
        $script:serverEvent.changeKey='v2'
        return $script:serverEvent
    }
    $script:fakeIdentity = $identity; $script:writes=0; $script:simulateFailure=$false
    function Make-Scan($row) { [pscustomobject]@{Identity=$identity; Target='Test Alberta';Rows=@($row);ScannedAt=[datetimeoffset]::UtcNow} }
    $script:serverEvent=Make-Event
    $row=Propose $script:serverEvent; $scan=Make-Scan $row
    $scan | Add-Member -NotePropertyName Profile -NotePropertyValue 'Manitoba'
    Expect-Throw { Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'province-changed') } 'Changing province invalidates old repair selection'
    Assert ($script:writes -eq 0) 'Province mismatch results in zero writes'
    $scan.Profile='Alberta'
    $script:serverEvent.changeKey='changed'
    Expect-Throw { Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'changed') } 'Change after preview stops repair'
    Assert ($script:writes -eq 0) 'Changed event results in zero writes'
    $script:serverEvent=Make-Event; $row=Propose $script:serverEvent; $scan=Make-Scan $row
    $script:fakeIdentity=@{UserId='other-user';TenantId='test-tenant'}
    Expect-Throw { Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'identity') } 'Repair with a different signed-in user is blocked'
    Assert ($script:writes -eq 0) 'Identity mismatch results in zero writes'
    $script:fakeIdentity=$identity
    $blockedDirectory = Join-Path $TestDirectory 'not-a-directory'
    [IO.File]::WriteAllText($blockedDirectory, 'synthetic')
    Expect-Throw { Invoke-ABApply @{} $scan @($row) $blockedDirectory } 'Journal storage failure stops repair before any write'
    Assert ($script:writes -eq 0) 'Journal failure results in zero writes'
    $scan.ScannedAt=[datetimeoffset]::UtcNow.AddHours(-1)
    Expect-Throw { Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'expired') } 'Expired preview is blocked'
    $scan=Make-Scan $row
    $successDir=Join-Path $TestDirectory 'success'
    $result=Invoke-ABApply @{} $scan @($row) $successDir
    Assert ($result.Count -eq 1 -and $script:writes -eq 1) 'Selected event is patched once and verified'
    $saved=Get-ABJournalIndex $identity $successDir
    Assert ($saved[$row.Id] -eq 'Verified') 'Verified result is durably journaled'
    Expect-Throw { Invoke-ABApply @{} $scan @($row) $successDir } 'Previously attempted event is blocked from another repair'
    Assert ($script:writes -eq 1) 'Repeat repair makes no additional writes'
    $script:serverEvent=Make-Event; $row=Propose $script:serverEvent; $scan=Make-Scan $row
    $script:simulateFailure=$true
    $failedDir=Join-Path $TestDirectory 'uncertain'
    Expect-Throw { Invoke-ABApply @{} $scan @($row) $failedDir } 'Uncertain write outcome stops the batch'
    $saved=Get-ABJournalIndex $identity $failedDir
    Assert ($saved[$row.Id] -eq 'Sending - outcome may be unknown') 'Uncertain outcome remains in the encrypted journal'
    $before=$script:writes
    Expect-Throw { Invoke-ABApply @{} $scan @($row) $failedDir } 'Uncertain event cannot be blindly retried'
    Assert ($script:writes -eq $before) 'Blocked uncertain retry makes zero additional writes'

    $script:simulateFailure=$false
    $script:serverEvent=Make-Event;$script:serverEvent.isAllDay=$true
    $row=Propose $script:serverEvent;$scan=Make-Scan $row;$before=$script:writes
    Expect-Throw {Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'all-day-excluded')} 'All-day update is rejected before writing'
    Assert ($script:writes -eq $before) 'All-day update sends zero writes'
    $script:regionalTarget=$true; $script:simulateFailure=$false
    $script:serverEvent=Make-Event;$before=$script:writes
    $row=New-ABProposal $script:serverEvent $script:testZone 'Test Alberta' $now -IncludeTimeInTitle
    $scan=Make-Scan $row;$scan|Add-Member IncludeTimeInTitle $true
    $script:regionalTarget=$false
    $result=Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory 'title-in-same-update')
    Assert ($result.Count -eq 1 -and $script:writes -eq $before+1 -and $script:serverEvent.subject -eq 'Synthetic test event (09:00AM)') 'Time and title are applied and verified together in one update'
    $script:regionalTarget=$true
    foreach($case in $profiles) {
        $definition=Get-TZProfile $case[0]
        $target=Get-ABTarget $definition.Key
        $script:serverEvent=Make-Event
        $script:serverEvent.originalStartTimeZone=$definition.OldId; $script:serverEvent.originalEndTimeZone=$definition.OldId
        $start=[datetimeoffset]::new([datetime]'2099-01-10T09:00:00',[timespan]::FromHours($definition.OffsetHours-1))
        $script:serverEvent.start.dateTime=$start.UtcDateTime.ToString('s'); $script:serverEvent.end.dateTime=$start.AddHours(1).UtcDateTime.ToString('s')
        $row=New-ABProposal $script:serverEvent $target.Zone $target.Alias $now $definition.Key
        $scan=[pscustomobject]@{Identity=$identity;Target=$definition.NewId;Profile=$definition.Key;Rows=@($row);ScannedAt=[datetimeoffset]::UtcNow}
        $before=$script:writes
        $result=Invoke-ABApply @{} $scan @($row) (Join-Path $TestDirectory ('regional-'+$definition.NewId.Replace(' ','')))
        Assert ($result.Count -eq 1 -and $script:writes -eq $before+1 -and $script:serverEvent.originalStartTimeZone -eq $definition.NewId) "$($definition.Key): full simulated repair verifies the selected target zone"
    }
    $script:regionalTarget=$false

    # Scan the full owned-calendar inventory over paginated reads; never PATCH.
    function script:Invoke-ABGet($Path) {
        if($Path -match 'isSharedWithMe'){throw 'RequestBroker-ParseUri: unsupported calendar property'}
        if ($Path.StartsWith('/me/calendars?')) { return @{value=@(
            @{id='owned';name='Calendar';owner=@{address='test@example.invalid'};canEdit=$true},
            @{id='shared';name='Shared';owner=@{address='other@example.invalid'};canEdit=$true})} }
        if ($Path -like '/me/calendars/owned/events?*') {
            $past=Make-Event; $past.start.dateTime='2020-01-10T16:00:00'; $past.end.dateTime='2020-01-10T17:00:00'
            $future=Make-Event
            $allDayRows=@(foreach($kind in 'singleInstance','seriesMaster','occurrence','exception'){
                $allDay=Make-Event;$allDay.type=$kind;$allDay.isAllDay=$true;$allDay.end.dateTime='2099-01-14T07:00:00';$allDay
            })
            return @{value=(@($past,$future)+$allDayRows);'@odata.nextLink'='https://graph.microsoft.com/v1.0/me/calendars/owned/events?page=2'}
        }
        if ($Path -eq 'https://graph.microsoft.com/v1.0/me/calendars/owned/events?page=2') {
            $series=Make-Event; $series.type='seriesMaster'; $series.start.dateTime='2020-01-10T16:00:00'; $series.end.dateTime='2020-01-10T17:00:00'
            return @{value=@($series)}
        }
        if($Path -like '/me/events/*/instances?*'){
            $occurrence=Make-Event; $occurrence.type='occurrence'
            $when=[datetimeoffset]::UtcNow.AddMonths(3)
            if($when -lt [datetimeoffset]'2027-01-01T00:00:00Z'){$when=[datetimeoffset]'2027-01-01T16:00:00Z'}
            $occurrence.start.dateTime=$when.UtcDateTime.ToString('s'); $occurrence.end.dateTime=$when.AddHours(1).UtcDateTime.ToString('s')
            return @{value=@($occurrence)}
        }
        throw "Unexpected read in scan fixture: $Path"
    }
    $before=$script:writes
    $scanProgress=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $inventory=Get-ABScan -Config $defaultConfig -Start $boundary -JournalDirectory (Join-Path $TestDirectory 'inventory') -ProgressQueue $scanProgress
    Assert ($inventory.Rows.Count -eq 2) 'Full inventory follows pages, keeps far-future events and old series, and excludes old singles'
    Assert ($inventory.SkippedCalendars.Count -eq 1 -and $inventory.Rows[0].Calendar -eq 'Calendar') 'Owned-calendar scope is labeled and shared calendars are excluded'
    Assert ($script:writes -eq $before) 'Full inventory scan performs zero writes'
    $seriesRow=@($inventory.Rows | Where-Object Kind -eq 'seriesMaster')[0]
    Assert ($seriesRow.DefaultVisible -and $seriesRow.DisplayStartUtc -gt [datetimeoffset]::UtcNow -and $seriesRow.CurrentLocal -notmatch '2020') 'A 2020 series displays a verified future occurrence instead of its original start'
    Assert (@($scanProgress.ToArray() | Where-Object Stage -eq 'Complete').Count -eq 1) 'Scanning emits an explicit completion notification'
    Assert (@($scanProgress.ToArray() | Where-Object Stage -eq 'Reading').Count -ge 1) 'Scanning emits page-by-page progress before completion'

    # Mailbox preference writes use a separate endpoint and never send an event patch.
    $script:regionalTarget=$true; $script:settingsWrites=0; $script:settingsPersist=$true
    $script:mailboxZone='Mountain Standard Time'
    function script:Invoke-ABGet($Path){
        if($Path -eq '/me/mailboxSettings?$select=timeZone'){return @{timeZone=$script:mailboxZone}}
        throw "Unexpected settings read: $Path"
    }
    function script:Invoke-MgGraphRequest {
        param($Method,$Uri,$Headers,$Body,$ContentType,$OutputType)
        if($Method -ne 'PATCH' -or $Uri -ne 'https://graph.microsoft.com/v1.0/me/mailboxSettings'){throw 'Settings must not write events or other endpoints'}
        $patch=$Body | ConvertFrom-Json -AsHashtable
        if(@($patch.Keys).Count -ne 1 -or -not $patch.ContainsKey('timeZone')){throw 'Only the timeZone preference may be patched'}
        $script:settingsWrites++
        if($script:settingsPersist){$script:mailboxZone=$patch.timeZone}
        return @{timeZone=$patch.timeZone}
    }
    $settingsScan=[pscustomobject]@{Identity=$identity;Profile='Alberta';Target='Alberta Standard Time'}
    $settingsDir=Join-Path $TestDirectory 'settings'
    Expect-Throw {Set-TZOutlookTimezone @{} $settingsScan $settingsDir} 'Outlook time-zone changes require explicit confirmation'
    Assert ($script:settingsWrites -eq 0) 'No settings confirmation means zero writes'
    $script:fakeIdentity=@{UserId='other-user';TenantId='test-tenant'}
    Expect-Throw {Set-TZOutlookTimezone @{} $settingsScan $settingsDir -Confirmed} 'Outlook setting cannot be changed for a different signed-in account'
    Assert ($script:settingsWrites -eq 0) 'Wrong settings account results in zero writes'
    $script:fakeIdentity=$identity
    $settingsProgress=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $result=Set-TZOutlookTimezone @{} $settingsScan $settingsDir -Confirmed -ProgressQueue $settingsProgress
    Assert ($result.Changed -and $script:mailboxZone -eq 'Alberta Standard Time' -and $script:settingsWrites -eq 1) 'The confirmed Outlook preference change is applied and read back'
    Assert ($script:writes -eq $before) 'Changing the Outlook preference sends no calendar-event writes'
    Assert ((Get-ABJournalIndex $identity $settingsDir).Count -eq 0) 'Preference journals do not interfere with event repair history'
    $result=Set-TZOutlookTimezone @{} $settingsScan $settingsDir -Confirmed
    Assert (-not $result.Changed -and $script:settingsWrites -eq 1) 'An already-correct Outlook preference is not written again'
    $script:mailboxZone='Mountain Standard Time'; $script:settingsPersist=$false
    Expect-Throw {Set-TZOutlookTimezone @{} $settingsScan $settingsDir -Confirmed} 'A settings change that fails read-back verification is reported as a failure'
    Assert (@($settingsProgress.ToArray() | Where-Object Stage -eq 'Complete').Count -eq 1) 'A verified Outlook preference change emits a completion notification'
    Write-Output "All $script:checks offline checks passed. Synthetic encrypted test journals: $TestDirectory"
} $TestDirectory
Remove-Module $module
