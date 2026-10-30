#requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:GraphRoot = 'https://graph.microsoft.com/v1.0'
$script:Fields = 'id,subject,start,end,type,seriesMasterId,isOrganizer,isAllDay,isCancelled,originalStartTimeZone,originalEndTimeZone,changeKey,webLink,iCalUId,recurrence'
$script:Headers = @{ Prefer = 'outlook.timezone="UTC", IdType="ImmutableId"' }
$script:SessionIdentity=$null
$script:SessionFingerprint=''

function Get-TZContextFingerprint($Context){
    if($null -eq $Context){return ''}
    $scopeText=(@(Get-Field $Context 'Scopes' @())|Sort-Object)-join ' '
    @([string](Get-Field $Context 'AuthType'),[string](Get-Field $Context 'ClientId'),[string](Get-Field $Context 'TenantId'),
        [string](Get-Field $Context 'Account'),$scopeText)-join '|'
}

function Disconnect-TZSession {
    if(Get-Module Microsoft.Graph.Authentication){Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null}
    $script:SessionIdentity=$null; $script:SessionFingerprint=''
}

function Connect-TZSession($Config,$ProgressQueue=$null){Connect-ABGraph $Config Scan $ProgressQueue}

function Get-TZProfile {
    param([ValidateSet('Alberta','British Columbia','British Columbia (Mountain)','Manitoba')][string]$Name='Alberta')
    $values = switch ($Name) {
        'Alberta' { @('Mountain Standard Time','Alberta Standard Time',-6,'Alberta') }
        'British Columbia' { @('Pacific Standard Time','British Columbia Standard Time',-7,'British Columbia') }
        'British Columbia (Mountain)' { @('Mountain Standard Time','Alberta Standard Time',-6,'British Columbia (Alberta-aligned areas)') }
        'Manitoba' { @('Central Standard Time','Manitoba Standard Time',-5,'Manitoba') }
    }
    $offset = [timespan]::FromHours($values[2])
    $labels = switch ($Name) {
        'British Columbia' { @('Pacific Time (US & Canada)','(UTC-07:00) Pacific Time (British Columbia)') }
        'Manitoba' { @('Central Time (US & Canada)','(UTC-05:00) Manitoba') }
        default { @('Mountain Time (US & Canada)','(UTC-06:00) Mountain Time (Alberta)') }
    }
    [pscustomobject]@{
        Key=$Name; Name=$values[3]; OldId=$values[0]; NewId=$values[1]; OffsetHours=$values[2]
        OldLabel=$labels[0]; NewLabel=$labels[1]
        Start=[datetimeoffset]::new([datetime]'2026-11-01T00:00:00',$offset)
        Cutover=[datetimeoffset]::new([datetime]'2026-11-01T02:00:00',$offset)
    }
}

function Get-Field($Object, [string]$Name, $Default = $null) {
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    } elseif ($Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}

function Send-TZProgress($Queue, [string]$Stage, [string]$Message, [int]$Completed=0, [int]$Total=0, [string]$Account='') {
    if ($null -ne $Queue) {
        $Queue.Enqueue([pscustomobject]@{Stage=$Stage;Message=$Message;Completed=$Completed;Total=$Total;Account=$Account})
    }
}

function Read-ABConfig([string]$Path) {
    $c = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable
    if (-not $c.ContainsKey('AuthenticationMode')) { $c.AuthenticationMode = 'DedicatedApps' }
    if ($c.AuthenticationMode -notin 'MicrosoftDefault','DedicatedApps') { throw 'AuthenticationMode must be MicrosoftDefault or DedicatedApps.' }
    if ($c.AuthenticationMode -eq 'MicrosoftDefault') {
        if ($c.TenantId) {
            $id = [guid]::Empty
            if (-not [guid]::TryParse([string]$c.TenantId,[ref]$id) -or $id -eq [guid]::Empty) { throw 'TenantId must be blank or a valid organization tenant ID.' }
        }
        return $c
    }
    foreach ($key in 'TenantId','ScanClientId','FixClientId') {
        $id = [guid]::Empty
        if (-not [guid]::TryParse([string](Get-Field $c $key), [ref]$id) -or $id -eq [guid]::Empty) {
            throw "Set $key in config.json to the ID provided by your Microsoft 365 administrator."
        }
    }
    if ($c.ScanClientId -eq $c.FixClientId) { throw 'Use separate scan and repair app registrations.' }
    return $c
}

function Connect-ABGraph($Config, [ValidateSet('Scan','Fix','Settings')][string]$Mode, $ProgressQueue=$null) {
    Send-TZProgress $ProgressQueue 'Loading' 'Loading Microsoft sign-in…'
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
    catch { throw 'Microsoft Graph sign-in is not installed. In PowerShell 7 run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
    $dedicated = (Get-Field $Config 'AuthenticationMode' 'DedicatedApps') -eq 'DedicatedApps'
    $client = if ($dedicated) {
        if ($Mode -eq 'Scan') { $Config.ScanClientId } else { $Config.FixClientId }
    } else { '14d82eec-204b-4c2f-b7e8-296a70dab67e' }
    $calendarScope = switch($Mode) { 'Scan' {'Calendars.Read'} 'Fix' {'Calendars.ReadWrite'} 'Settings' {'MailboxSettings.ReadWrite'} }
    $tenant = if (Get-Field $Config 'TenantId' '') { $Config.TenantId } else { 'organizations' }
    $existing=Get-MgContext
    $existingScopes=@(Get-Field $existing 'Scopes' @() | ForEach-Object {([string]$_).Replace('https://graph.microsoft.com/','')})
    $canReuse=($calendarScope -in $existingScopes -or ($Mode -eq 'Scan' -and 'Calendars.ReadWrite' -in $existingScopes))
    if($null -ne $script:SessionIdentity -and $existing -and $existing.ClientId -eq $client -and
        ($tenant -eq 'organizations' -or $existing.TenantId -eq $tenant) -and $canReuse -and
        (Get-TZContextFingerprint $existing) -ceq $script:SessionFingerprint){
        Send-TZProgress $ProgressQueue 'SignedIn' 'Using your existing signed-in session.' -Account $script:SessionIdentity.Account
        return $script:SessionIdentity
    }
    $requestedScopes=if($dedicated){@('User.Read',$calendarScope)}else{@('User.Read','Calendars.ReadWrite','MailboxSettings.ReadWrite')}
    $arguments = @{TenantId=$tenant; Scopes=$requestedScopes; ContextScope='Process'; NoWelcome=$true}
    if ($dedicated) { $arguments.ClientId = $client }
    # Omitting ClientId selects Microsoft's existing Graph PowerShell app. No new registration is created.
    Send-TZProgress $ProgressQueue 'SignIn' 'Waiting for Microsoft sign-in…'
    Connect-MgGraph @arguments | Out-Null
    $context = Get-MgContext
    if ($context.AuthType -ne 'Delegated' -or ($tenant -ne 'organizations' -and $context.TenantId -ne $tenant) -or
        $context.ClientId -ne $client -or $context.TenantId -eq '9188040d-6c67-4c5b-b112-36a304b66dad') {
        Disconnect-MgGraph | Out-Null
        throw 'The signed-in account does not match the configured tenant and app.'
    }
    # The SDK may report previously consented scopes for Microsoft's shared app.
    # Normalize Graph-qualified scopes before checking capabilities. Never request those extra scopes.
    $graphPrefix='https://graph.microsoft.com/'
    $scopes=@($context.Scopes | ForEach-Object {
        $scope=([string]$_).Trim()
        if ($scope.StartsWith($graphPrefix,[StringComparison]::OrdinalIgnoreCase)) { $scope=$scope.Substring($graphPrefix.Length) }
        if ($scope) { $scope }
    } | Sort-Object -Unique)
    $allowed = @('User.Read', $calendarScope, 'openid','profile','email','offline_access')
    if ($Mode -eq 'Fix') { $allowed += @('Calendars.Read','MailboxSettings.ReadWrite') }
    elseif ($Mode -eq 'Settings') { $allowed += @('Calendars.Read','Calendars.ReadWrite','MailboxSettings.Read') }
    elseif (-not $dedicated) { $allowed += 'Calendars.ReadWrite' } # Shared Microsoft app may retain prior calendar consent.
    if(-not $dedicated){$allowed+=@('Calendars.ReadWrite','MailboxSettings.ReadWrite')}
    $extra = @($scopes | Where-Object { $_ -notin $allowed })
    $hasCalendarScope = $calendarScope -in $scopes -or ($Mode -eq 'Scan' -and -not $dedicated -and 'Calendars.ReadWrite' -in $scopes)
    if (-not $hasCalendarScope) {
        Disconnect-MgGraph | Out-Null
        $access=if($Mode -eq 'Settings'){'mailbox-settings'}else{'calendar'}
        throw "Sign-in completed but the session does not have the required $calendarScope $access access. No request was made. Sign in again and approve the requested access; your organization may require administrator consent."
    }
    if ($dedicated -and $extra.Count -gt 0) {
        Disconnect-MgGraph | Out-Null
        throw ('The dedicated app returned permissions outside its configured purpose: '+($extra -join ', ')+'. No calendar request was made. Ask IT to review that dedicated app consent.')
    }
    $me = Invoke-ABGet '/me?$select=id,userPrincipalName,displayName,mail'
    Send-TZProgress $ProgressQueue 'SignedIn' 'Sign-in complete. Connected to Microsoft 365.' -Account $me.userPrincipalName
    $notes=[System.Collections.Generic.List[string]]::new()
    if (-not $dedicated -and $extra.Count -gt 0) {
        $notes.Add("Microsoft's shared app has $($extra.Count) additional existing permissions. This tool does not use them.")
    }
    if ($Mode -eq 'Scan' -and 'Calendars.ReadWrite' -in $scopes) {
        $notes.Add('Your session includes calendar-write access. This scan only reads; repairs still require confirmation.')
    }
    $note=$notes -join ' '
    $script:SessionIdentity=[pscustomobject]@{ UserId = $me.id; Account = $me.userPrincipalName; Mail = [string](Get-Field $me 'mail' ''); TenantId = $context.TenantId; AccessNote=$note }
    $script:SessionFingerprint=Get-TZContextFingerprint $context
    return $script:SessionIdentity
}

function Assert-ABUri([string]$Uri) {
    $u = [uri]$Uri
    if ($u.Scheme -ne 'https' -or $u.Host -ne 'graph.microsoft.com' -or -not $u.IsDefaultPort -or
        (-not $u.AbsolutePath.StartsWith('/v1.0/me/', [StringComparison]::OrdinalIgnoreCase) -and $u.AbsolutePath -ne '/v1.0/me') -or
        $u.UserInfo -or $u.Fragment) { throw 'Blocked a request outside the signed-in user Graph endpoint.' }
}

function Invoke-ABGet([string]$Path) {
    $uri = if ($Path.StartsWith('/')) { $script:GraphRoot + $Path } else { $Path }
    Assert-ABUri $uri
    # General writes are not retried. Confirmed recurring 412 rejections have a bounded revalidation path.
    Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $script:Headers -OutputType PSObject
}

function Get-ABCollection([string]$Path, $ProgressQueue=$null, [string]$Context='Reading calendar data') {
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $pageNumber=0
    while ($Path) {
        if (-not $seen.Add($Path)) { throw 'Graph returned a repeated page link; scan stopped.' }
        $page = Invoke-ABGet $Path
        $pageNumber++
        Send-TZProgress $ProgressQueue 'Reading' "$Context · page $pageNumber"
        foreach ($item in @(Get-Field $page 'value' @())) { $item }
        $Path = [string](Get-Field $page '@odata.nextLink' '')
    }
}

function Get-ABTarget([string]$Profile='Alberta') {
    $definition = Get-TZProfile $Profile
    $zones = @(Get-ABCollection '/me/outlook/supportedTimeZones')
    $matches = @($zones | Where-Object { $_.alias -eq $definition.NewId })
    if ($matches.Count -ne 1) { throw "The mailbox does not advertise $($definition.NewId). Check Microsoft updates before continuing." }
    try { $zone = [TimeZoneInfo]::FindSystemTimeZoneById($matches[0].alias) }
    catch { throw "Install the Windows time-zone update for $($definition.Name) before using this tool." }
    foreach ($date in '2026-11-02T12:00:00Z','2027-01-15T12:00:00Z','2027-07-15T12:00:00Z') {
        if ($zone.GetUtcOffset([datetimeoffset]$date) -ne [timespan]::FromHours($definition.OffsetHours)) {
            throw "Local time-zone rules for $($definition.Name) failed validation. Install Windows updates."
        }
    }
    [pscustomobject]@{ Alias = $matches[0].alias; DisplayName = $matches[0].displayName; Zone = $zone; Profile=$definition }
}

function Convert-ABUtc($Value) {
    if ((Get-Field $Value 'timeZone') -ne 'UTC') { throw 'Expected event times in UTC. No conversion attempted.' }
    [datetimeoffset]::Parse([string]$Value.dateTime, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
}

function Test-TZNeedsReview($Event, $Definition, [datetimeoffset]$From) {
    if(Get-Field $Event 'isAllDay' $false){return $false}
    if ((Get-Field $Event 'isOrganizer') -ne $true -or (Get-Field $Event 'isCancelled' $false)) { return $false }
    if ((Get-Field $Event 'originalStartTimeZone') -ne $Definition.OldId -or
        (Get-Field $Event 'originalEndTimeZone') -ne $Definition.OldId) { return $false }
    try { return (Convert-ABUtc $Event.start) -ge $From } catch { return $false }
}

function Get-TZUpcomingOccurrence($Master, $Definition, [datetimeoffset]$From, $ProgressQueue=$null,$JournalIndex=$null) {
    # Confirm a future occurrence instead of showing the original series date.
    $anchor=$From
    $first=Convert-ABUtc $Master.start
    if ($first -gt $anchor) { $anchor=$first }
    $recurrence=Get-Field $Master 'recurrence'
    $pattern=Get-Field $recurrence 'pattern'
    $range=Get-Field $recurrence 'range'
    $interval=[math]::Max(1,[int](Get-Field $pattern 'interval' 1))
    $until=switch([string](Get-Field $pattern 'type' '')) {
        {$_ -in 'absoluteYearly','relativeYearly'} { $anchor.AddYears([math]::Max(2,$interval*2)) }
        {$_ -in 'absoluteMonthly','relativeMonthly'} { $anchor.AddMonths([math]::Max(24,$interval*2)) }
        'weekly' { $anchor.AddDays([math]::Max(730,$interval*14)) }
        default { $anchor.AddDays([math]::Max(730,$interval*2)) }
    }
    $endDate=[string](Get-Field $range 'endDate' '')
    if ($endDate) {
        $finiteEnd=[datetimeoffset]::new([datetime]::ParseExact($endDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture),[timespan]::FromHours($Definition.OffsetHours)).AddDays(2)
        if ($finiteEnd -gt $From) { $until=$finiteEnd }
    }
    $path='/me/events/'+[uri]::EscapeDataString($Master.id)+'/instances?startDateTime='+[uri]::EscapeDataString($From.ToString('o'))+
        '&endDateTime='+[uri]::EscapeDataString($until.ToString('o'))+'&$top=50&$orderby=start/dateTime&$select='+$script:Fields
    $seen=[System.Collections.Generic.HashSet[string]]::new()
    while($path) {
        if(-not $seen.Add($path)){throw 'Repeated occurrence page link.'}
        $page=Invoke-ABGet $path
        foreach($instance in @((Get-Field $page 'value' @()) | Sort-Object { (Convert-ABUtc $_.start).UtcTicks })) {
            if($null -ne $JournalIndex -and $JournalIndex.ContainsKey([string]$instance.id) -and $JournalIndex[[string]$instance.id] -eq 'Verified'){continue}
            if(Test-TZNeedsReview $instance $Definition $From) {
                return [pscustomobject]@{Event=$instance;CheckedUntil=$until}
            }
        }
        $path=[string](Get-Field $page '@odata.nextLink' '')
        Send-TZProgress $ProgressQueue 'Recurring' 'Checking upcoming recurring meetings…'
    }
    [pscustomobject]@{Event=$null;CheckedUntil=$until}
}

function Add-TZTimeToTitle([string]$Title,[datetime]$OriginalWallTime){
    $suffix=' ('+$OriginalWallTime.ToString('hh:mmtt',[Globalization.CultureInfo]::InvariantCulture)+')'
    if($Title.TrimEnd().EndsWith($suffix,[StringComparison]::OrdinalIgnoreCase)){return $Title.TrimEnd()}
    return $Title.TrimEnd()+$suffix
}
function New-ABProposal($Event, [TimeZoneInfo]$TargetZone, [string]$TargetAlias, [datetimeoffset]$Now = [datetimeoffset]::UtcNow,
    [string]$Profile='Alberta', [switch]$AllowOccurrence,[switch]$IncludeTimeInTitle) {
    $definition = Get-TZProfile $Profile
    $row = [pscustomobject]@{
        Selected = $false; Eligible = $false; Subject = [string](Get-Field $Event 'subject' '(untitled)')
        Kind = [string](Get-Field $Event 'type' 'unknown'); SourceZone = [string](Get-Field $Event 'originalStartTimeZone' '')
        CurrentLocal = ''; ProposedLocal = ''; ShiftMinutes = 0; Decision = ''; Result = ''; Calendar = ''; Profile=$definition.Key
        Id = [string](Get-Field $Event 'id' ''); ChangeKey = [string](Get-Field $Event 'changeKey' '')
        ETag = [string](Get-Field $Event '@odata.etag' ''); Patch = $null; Original = $Event
        DefaultVisible=$false; DisplayStartUtc=$null; MasterId=''; SeriesRecurrence=$null;CanResolve=$false;Handled=$false
    }
    if(Get-Field $Event 'isAllDay' $false){$row.Decision='All-day event - excluded';return $row}
    try {
        $start = Convert-ABUtc (Get-Field $Event 'start')
        $end = Convert-ABUtc (Get-Field $Event 'end')
        $from=$definition.Start
        if($Now -gt $from){$from=$Now}
        $row.DefaultVisible=Test-TZNeedsReview $Event $definition $from
        $row.DisplayStartUtc=$start
        $current = [TimeZoneInfo]::ConvertTime($start, $TargetZone)
        $row.CurrentLocal = $current.ToString('yyyy-MM-dd HH:mm') + ' to ' + [TimeZoneInfo]::ConvertTime($end, $TargetZone).ToString('yyyy-MM-dd HH:mm')
        if (Get-Field $Event 'isCancelled' $false) { $row.Decision = 'Cancelled - excluded'; return $row }
        if ($row.Kind -ne 'singleInstance' -and -not ($AllowOccurrence -and $row.Kind -in 'occurrence','exception')) { $row.Decision = 'Recurring series · select row and use Adjust recurring series'; return $row }
        if ($start -lt $Now -or $start -lt $definition.Cutover) { $row.Decision = 'Before cutover or already started - excluded'; return $row }
        if(Get-Field $Event 'isAllDay' $false){$row.DefaultVisible=$false;$row.Decision='All-day event - excluded';return $row}
        if ((Get-Field $Event 'isOrganizer') -ne $true) { $row.Decision = 'Attendee copy - organizer must correct'; return $row }
        if ($row.SourceZone -eq $TargetAlias -and (Get-Field $Event 'originalEndTimeZone') -eq $TargetAlias) {
            $row.Decision = 'Already uses the selected regional zone'; return $row
        }
        # Old US/Canada Windows zones still serve unaffected locations; migration requires user review.
        if ($row.SourceZone -ne $definition.OldId -or (Get-Field $Event 'originalEndTimeZone') -ne $definition.OldId) {
            $row.Decision = 'Other, mixed, or unrecognized source zones - manual review'; return $row
        }
        $source = [TimeZoneInfo]::FindSystemTimeZoneById($definition.OldId)
        $wallStart = [TimeZoneInfo]::ConvertTime($start, $source).DateTime
        $wallEnd = [TimeZoneInfo]::ConvertTime($end, $source).DateTime
        $newStart = [datetimeoffset]::new($wallStart, $TargetZone.GetUtcOffset($wallStart))
        $newEnd = [datetimeoffset]::new($wallEnd, $TargetZone.GetUtcOffset($wallEnd))
        $row.ProposedLocal = $wallStart.ToString('yyyy-MM-dd HH:mm') + ' to ' + $wallEnd.ToString('yyyy-MM-dd HH:mm')
        $row.ShiftMinutes = ($newStart - $start).TotalMinutes
        if (($newEnd - $newStart) -ne ($end - $start) -or $newEnd -le $newStart) {
            $row.Decision = 'Duration changes across transition - manual review'; return $row
        }
        if ($newStart -lt $Now) { $row.Decision = 'Proposed start is in the past - excluded'; return $row }
        if (-not $row.Id -or -not $row.ChangeKey -or -not $row.ETag) { $row.Decision = 'Missing version information - excluded'; return $row }
        $row.Patch = @{
            start = @{ dateTime = $wallStart.ToString('yyyy-MM-ddTHH:mm:ss.fffffff'); timeZone = $TargetAlias }
            end = @{ dateTime = $wallEnd.ToString('yyyy-MM-ddTHH:mm:ss.fffffff'); timeZone = $TargetAlias }
        }

        $row.Eligible = $true
        if($IncludeTimeInTitle){$row.Subject=Add-TZTimeToTitle $row.Subject $wallStart;$row.Patch.subject=$row.Subject}
        $row.Decision = 'Candidate - confirm the intended local time'
    } catch { $row.Decision = 'Invalid or unsupported time data - manual review' }
    return $row
}

function Read-TZJournalRecords($Identity,[string]$JournalDirectory){
    if(-not(Test-Path -LiteralPath $JournalDirectory)){return}
    $records=[System.Collections.Generic.List[object]]::new()
    foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.dpapi' -File | Sort-Object LastWriteTimeUtc -Descending){
        try {
            $saved = Read-ABJournal $file.FullName
            if ($saved.Identity.UserId -eq $Identity.UserId -and $saved.Identity.TenantId -eq $Identity.TenantId) {
                foreach ($entry in $saved.Items) {
                    $records.Add([pscustomobject]@{Path=$file.FullName;Journal=$saved;Entry=$entry})
                }
            }
        } catch { throw 'A local journal could not be read. Have IT reconcile the journal before scanning again.' }
    }
    $records | Sort-Object { [datetimeoffset](Get-Field $_.Entry 'At' '1970-01-01T00:00:00Z') } -Descending
}

function Get-ABJournalIndex($Identity, [string]$JournalDirectory) {
    $index=[System.Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
    $seen=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($record in @(Read-TZJournalRecords $Identity $JournalDirectory)){
        $entry=$record.Entry
        foreach($id in @([string]$entry.Id,[string](Get-Field $entry 'AfterId' ''))){
            if($id -and $seen.Add($id) -and $entry.Status -ne 'Not applied - ready for a new preview'){$index[$id]=[string]$entry.Status}
        }
    }
    return ,$index
}

function Find-TZJournalRecord($Records,$Event){
    foreach($record in $Records){
        $entry=$record.Entry; $before=Get-Field $entry 'Before'
        if([string]$entry.Id -ceq [string]$Event.id -or [string](Get-Field $entry 'AfterId' '') -ceq [string]$Event.id){return $record}
        $oldUid=[string](Get-Field $before 'iCalUId' '');$currentUid=[string](Get-Field $Event 'iCalUId' '')
        if($oldUid -and $currentUid -and $oldUid -ceq $currentUid){return $record}
    }
    return $null
}

function Test-TZPatchApplied($Event,$Patch,[TimeZoneInfo]$Zone){
    try{
        foreach($field in 'start','end'){
            $wall=[datetime]::Parse($Patch[$field].dateTime,[Globalization.CultureInfo]::InvariantCulture)
            $expected=[datetimeoffset]::new($wall,$Zone.GetUtcOffset($wall))
            if((Convert-ABUtc (Get-Field $Event $field)) -ne $expected){return $false}
        }
        return $true
    }catch{return $false}
}

function Test-TZSameOccurrence($Event,$Before){
    try{
        if((Convert-ABUtc $Event.start) -ne (Convert-ABUtc $Before.start) -or (Convert-ABUtc $Event.end) -ne (Convert-ABUtc $Before.end)){return $false}
        foreach($field in 'subject','originalStartTimeZone','originalEndTimeZone','isOrganizer','isAllDay','isCancelled','seriesMasterId'){
            if([string](Get-Field $Event $field) -cne [string](Get-Field $Before $field)){return $false}
        }
        return $true
    }catch{return $false}
}

function Resolve-TZJournalRecord($Record,$Event,[TimeZoneInfo]$Zone){
    if($null -eq $Record){return 'New'}
    $entry=$Record.Entry
    $patch=Get-Field $entry 'Patch'
    if($null -eq $patch){return 'Unresolved'}
    $applied=Test-TZPatchApplied $Event $patch $Zone
    $unchanged=Test-TZSameOccurrence $Event $entry.Before
    $responseReceived=([string]$entry.Status -like 'Verified*' -or [string]$entry.Status -eq 'Response received - verifying' -or
        [bool](Get-Field $entry 'AfterId' ''))
    if($applied -and ($responseReceived -or -not $unchanged)){
        $entry.Status='Verified';$entry.AfterId=$Event.id;$entry.AfterChangeKey=Get-Field $Event 'changeKey';$entry.ReconciledAt=[datetimeoffset]::UtcNow
        Save-ABJournal $Record.Journal $Record.Path
        return 'Applied'
    }
    if($entry.Status -notlike 'Verified*' -and $unchanged -and (-not $applied -or $entry.Status -eq 'Prepared' -or $entry.Status -eq 'Not applied - ready for a new preview')){
        $entry.Status='Not applied - ready for a new preview'
        Save-ABJournal $Record.Journal $Record.Path
        return 'NotApplied'
    }
    return 'Unresolved'
}

function Test-ABInventoryDate($Event, [datetimeoffset]$Start) {
    if ((Get-Field $Event 'type') -eq 'seriesMaster') {
        # Include masters regardless of first date: moved exceptions can fall outside the base range.
        # A series row is a review item, not a promise that every series has future instances.
        return $true
    }
    if ((Get-Field $Event 'type') -ne 'singleInstance') { return $true }
    try { return (Convert-ABUtc $Event.end) -gt $Start -or (Convert-ABUtc $Event.start) -ge $Start }
    catch { return $true } # Keep malformed entries visible for review rather than silently dropping them.
}
function Test-TZCalendarOwned($Calendar,$Identity){
    # canEdit also applies to delegated calendars, so it cannot establish ownership.
    # Some Exchange deployments reject isSharedWithMe in $select. Use owner instead.
    $owner=[string](Get-Field (Get-Field $Calendar 'owner' @{}) 'address' '')
    if([string]::IsNullOrWhiteSpace($owner)){return $false}
    foreach($address in @((Get-Field $Identity 'Account' ''),(Get-Field $Identity 'Mail' ''))){
        if(-not [string]::IsNullOrWhiteSpace([string]$address) -and $owner.Trim() -ieq ([string]$address).Trim()){return $true}
    }
    return $false
}

function Get-ABScan($Config, [datetimeoffset]$Start,
    [string]$JournalDirectory = (Join-Path $env:LOCALAPPDATA 'AlbertaCalendarPilot\Journal'), [string]$Profile='Alberta', $ProgressQueue=$null,[switch]$IncludeTimeInTitle) {
    $definition = Get-TZProfile $Profile
    if (-not $PSBoundParameters.ContainsKey('Start')) { $Start = $definition.Start }
    $identity = Connect-ABGraph $Config Scan $ProgressQueue
    try {
        Send-TZProgress $ProgressQueue 'TimeZones' 'Connected. Checking the selected time zone…' -Account $identity.Account
        $target = Get-ABTarget $definition.Key
        $previous = Get-ABJournalIndex $identity $JournalDirectory
        $resolved=Get-TZResolvedIds $identity $JournalDirectory
        $rows = [System.Collections.Generic.List[object]]::new()
        $calendars = @(Get-ABCollection '/me/calendars?$select=id,name,owner,canEdit&$top=100' $ProgressQueue 'Reading calendars')
        $skipped = [System.Collections.Generic.List[string]]::new()
        $now=[datetimeoffset]::UtcNow
        $from=if($now -gt $Start){$now}else{$Start}
        $examined=0; $seriesUnverified=0; $calendarNumber=0
        foreach ($calendar in $calendars) {
            $calendarNumber++
            if (-not(Test-TZCalendarOwned $calendar $identity)) { $skipped.Add([string]$calendar.name); continue }
            $path = '/me/calendars/' + [uri]::EscapeDataString($calendar.id) + '/events?$top=250&$select=' + $script:Fields
            Send-TZProgress $ProgressQueue 'Reading' "Scanning calendar $calendarNumber of $($calendars.Count): $($calendar.name)"
            foreach ($e in Get-ABCollection $path $ProgressQueue "Reading $($calendar.name)") {
                $examined++
                if($resolved.Contains([string]$e.id)){continue}
                if(Get-Field $e 'isAllDay' $false){continue}
                if($examined % 25 -eq 0){Send-TZProgress $ProgressQueue 'Filtering' "$examined entries checked…"}
                if (-not (Test-ABInventoryDate $e $Start)) { continue }
                $row = New-ABProposal -Event $e -TargetZone $target.Zone -TargetAlias $target.Alias -Profile $definition.Key -IncludeTimeInTitle:$IncludeTimeInTitle
                if((Get-Field $e 'type') -eq 'seriesMaster') {
                    $row.DefaultVisible=$false
                    if((Get-Field $e 'isOrganizer') -eq $true -and -not (Get-Field $e 'isCancelled' $false) -and
                        (Get-Field $e 'originalStartTimeZone') -eq $definition.OldId -and (Get-Field $e 'originalEndTimeZone') -eq $definition.OldId) {
                        Send-TZProgress $ProgressQueue 'Recurring' "Checking upcoming occurrence · $examined entries checked…"
                        try {
                            $next=Get-TZUpcomingOccurrence $e $definition $from $ProgressQueue $previous
                            if($null -ne $next.Event) {
                                $row=New-ABProposal -Event $next.Event -TargetZone $target.Zone -TargetAlias $target.Alias -Profile $definition.Key -IncludeTimeInTitle:$IncludeTimeInTitle
                                $row.Kind='seriesMaster'; $row.MasterId=$e.id; $row.Eligible=$false
                                $row.Decision='Recurring series · select row to adjust future occurrences'
                            } else { $row.Decision='No upcoming affected occurrence in the checked range'; $seriesUnverified++ }
                        } catch { $row.Decision='Upcoming occurrence could not be verified · manual review'; $seriesUnverified++ }
                    }
                    $row.MasterId=$e.id; $row.SeriesRecurrence=Get-Field $e 'recurrence'
                }
                $row.Calendar = [string]$calendar.name
                if ((Get-Field $calendar 'canEdit') -ne $true) { $row.Eligible=$false; $row.DefaultVisible=$false; $row.Decision='Calendar is read-only - review only' }
                if ($row.Kind -ne 'seriesMaster' -and $previous.ContainsKey($row.Id)) {
                    $row.Eligible = $false
                    $row.Decision = 'Previously attempted by this pilot - reconcile journal before further changes'
                    $row.CanResolve=$true
                    if($previous[$row.Id] -like 'Verified*'){$row.DefaultVisible=$false;$row.CanResolve=$false}
                }
                $rows.Add($row)
            }
        }
        foreach($recovery in @(Get-TZMigrationRecoveryRows $identity $JournalDirectory $definition.Key)){
            if($resolved.Contains([string]$recovery.MasterId)){continue}
            if(Get-Field $recovery.Original 'isAllDay' $false){continue}
            $existing=@($rows|Where-Object {$_.Kind -eq 'seriesMaster' -and $_.MasterId -ceq $recovery.MasterId})
            if($existing.Count){foreach($item in $existing){$item.DefaultVisible=$true;$item.CanResolve=$true;$item.Decision=$recovery.Decision;$item.CurrentLocal=$recovery.CurrentLocal;$item.ProposedLocal=$recovery.ProposedLocal;$item.DisplayStartUtc=$recovery.DisplayStartUtc}}
            else{$rows.Add($recovery)}
        }
        $visible=@($rows | Where-Object DefaultVisible).Count
        Send-TZProgress $ProgressQueue 'Complete' "Scan complete · $visible entries need review · $examined entries checked" 1 1 -Account $identity.Account
        [pscustomobject]@{ Identity = $identity; Target = $target.Alias; Rows = @($rows.ToArray() | Sort-Object DisplayStartUtc,Calendar,Subject)
            Start=$Start; IncludeTimeInTitle=[bool]$IncludeTimeInTitle; Profile=$definition.Key; SkippedCalendars=$skipped.ToArray(); ScannedAt = [datetimeoffset]::UtcNow
            ExaminedCount=$examined; UnverifiedSeriesCount=$seriesUnverified }
    } finally { } # Keep the app session until the window closes.
}

function Save-ABJournal($Journal, [string]$Path) {
    # Protect bytes directly: SecureString's 65,536-character limit is too small for series snapshots.
    if (-not $IsWindows) { throw 'Encrypted journals require Windows.' }
    $plain=$null;$bytes=$null;$protected=$null
    try{
        $plain=$Journal|ConvertTo-Json -Depth 40 -WarningAction Stop
        $bytes=[Text.Encoding]::UTF8.GetBytes($plain)
        $protected=[Security.Cryptography.ProtectedData]::Protect($bytes,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
        $encrypted='TZJ2:'+ [Convert]::ToBase64String($protected)
    }finally{
        if($null -ne $bytes){[Array]::Clear($bytes,0,$bytes.Length)}
        $plain=$null
    }
    $temp = $Path + '.tmp'
    [IO.File]::WriteAllText($temp, $encrypted, [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function Read-ABJournal([string]$Path){
    if(-not $IsWindows){throw 'Encrypted journals require Windows.'}
    $saved=[IO.File]::ReadAllText($Path).Trim()
    $bytes=$null;$plain=$null;$secure=$null
    try{
        if($saved.StartsWith('TZJ2:',[StringComparison]::Ordinal)){
            $protected=[Convert]::FromBase64String($saved.Substring(5))
            $bytes=[Security.Cryptography.ProtectedData]::Unprotect($protected,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
            $plain=[Text.Encoding]::UTF8.GetString($bytes)
        }else{
            # Read previous releases' DPAPI-protected SecureString journals without replacing them.
            $secure=$saved|ConvertTo-SecureString
            $plain=[Net.NetworkCredential]::new('',$secure).Password
        }
        $plain|ConvertFrom-Json -AsHashtable
    }finally{
        if($null -ne $bytes){[Array]::Clear($bytes,0,$bytes.Length)}
        if($null -ne $secure){$secure.Dispose()}
        $plain=$null
    }
}

function Get-TZErrorDetails($Record) {
    # Save only diagnostic fields, never invocation arguments or request/authentication headers.
    $details=@{At=[datetimeoffset]::UtcNow;Message=$Record.Exception.Message;ExceptionType=$Record.Exception.GetType().FullName;
        ErrorId=$Record.FullyQualifiedErrorId;ServerDetails=[string](Get-Field $Record.ErrorDetails 'Message' '')}
    $exception=$Record.Exception
    for($i=0;$null -ne $exception -and $i -lt 4;$i++){
        foreach($name in 'ResponseStatusCode','StatusCode'){
            $value=Get-Field $exception $name
            if($null -ne $value){$details.HttpStatus=[string]$value}
        }
        $response=Get-Field $exception 'Response'
        if($response){
            $status=Get-Field $response 'StatusCode'
            if($null -ne $status){$details.HttpStatus=[string]$status}
        }
        $exception=$exception.InnerException
    }
    try{
        $json=$details.ServerDetails
        # Invoke-MgGraphRequest may wrap the JSON in an HTTP response transcript.
        if($json -match '(?s)\r?\n\r?\n(\{.*\})\s*$'){$json=$Matches[1]}
        $server=$json|ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if($server.error){
            $details.GraphCode=[string](Get-Field $server.error 'code' '')
            $details.GraphMessage=[string](Get-Field $server.error 'message' '')
            $inner=Get-Field $server.error 'innerError'
            if($inner){$details.RequestId=[string](Get-Field $inner 'request-id' '')}
            for($depth=0;$null -ne $inner -and $depth -lt 5;$depth++){
                $reason=[string](Get-Field $inner 'message' '')
                if($reason){$details.GraphReason=$reason}
                $inner=Get-Field $inner 'internalexception' (Get-Field $inner 'innererror')
            }
        }
    }catch{} # A plain-text SDK error is still retained verbatim inside the encrypted journal.
    if(-not $details.ContainsKey('RequestId') -and $details.ServerDetails -match '(?im)^request-id:\s*([^\r\n]+)'){
        $details.RequestId=$Matches[1].Trim()
    }
    $details
}

function Get-TZLastFailure([string]$JournalDirectory) {
    if(-not(Test-Path -LiteralPath $JournalDirectory)){return 'No saved repair error is available.'}
    foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.dpapi' -File | Sort-Object LastWriteTimeUtc -Descending){
        try{
            $saved=Read-ABJournal $file.FullName
            if(Get-Field $saved 'Failure'){
                return "Repair stopped at occurrence $($saved.Failure.ItemNumber).`n$($saved.Failure.VerifiedCount) occurrences verified before stopping.`n`n$($saved.Failure | ConvertTo-Json -Depth 8)`n`nEncrypted journal: $($file.FullName)"
            }
        }catch{return 'A saved journal could not be read. Ask IT to review it.'}
    }
    'No detailed repair error was saved. Earlier versions did not retain the Microsoft error response. Keep the existing journals; they still record which updates were verified.'
}

function Invoke-ABApply($Config, $Scan, [object[]]$Rows, [string]$JournalDirectory, $ProgressQueue=$null,
    [switch]$AllowOccurrences, [switch]$ConfirmedSeries) {
    $definition = Get-TZProfile ([string](Get-Field $Scan 'Profile' 'Alberta'))
    if($AllowOccurrences -and (-not $ConfirmedSeries -or -not (Get-Field $Scan 'MasterId' ''))){throw 'Confirm the future-series preview before updating occurrences.'}
    if ($Rows.Count -eq 0) { throw 'Select at least one eligible event.' }
    $allowedKinds=if($AllowOccurrences){@('occurrence','exception')}else{@('singleInstance')}
    if (@($Rows | Where-Object { -not $_.Eligible -or $_.Kind -notin $allowedKinds }).Count) { throw 'Selection includes an ineligible event.' }
    if (@($Rows | Where-Object { $_.Profile -ne $definition.Key }).Count) { throw 'The selected province no longer matches the preview. Scan again.' }
    if (([datetimeoffset]::UtcNow - [datetimeoffset]$Scan.ScannedAt).TotalMinutes -gt 30) { throw 'Preview expired. Scan again before fixing.' }
    $identity = Connect-ABGraph $Config Fix $ProgressQueue
    $journal = $null; $entry=$null; $completed=0; $alreadyApplied=0; $itemNumber=0; $stage='Validating repair'
    try {
        if ($identity.UserId -ne $Scan.Identity.UserId -or $identity.TenantId -ne $Scan.Identity.TenantId) {
            throw 'Repair sign-in must be the same user and tenant as the preview.'
        }
        $target = Get-ABTarget $definition.Key
        if ($target.Alias -ne $Scan.Target) { throw 'The target time zone changed. Scan again.' }
        if($AllowOccurrences){
            $master=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($Scan.MasterId)+'?$select='+$script:Fields)
            if($master.type -ne 'seriesMaster' -or $master.isOrganizer -ne $true -or $master.changeKey -cne $Scan.MasterChangeKey -or
                (Get-Field $master '@odata.etag') -cne $Scan.MasterETag){throw 'The recurring series changed after preview. Preview it again.'}
        }
        $previous = Get-ABJournalIndex $identity $JournalDirectory
        if (-not $AllowOccurrences -and @($Rows | Where-Object { $previous.ContainsKey($_.Id) }).Count) { throw 'A selected event already has a local repair journal. Reconcile it before further changes.' }
        $records=if($AllowOccurrences){@(Read-TZJournalRecords $identity $JournalDirectory)}else{@()}
        New-Item -ItemType Directory -Path $JournalDirectory -Force | Out-Null
        $journalPath = Join-Path $JournalDirectory (([guid]::NewGuid().ToString()) + '.dpapi')
        $journal = @{ Version = 2; Created = [datetimeoffset]::UtcNow; Identity = $identity; Profile=$definition.Key; Items = [System.Collections.Generic.List[object]]::new() }
        Save-ABJournal $journal $journalPath
        $completed=0; $alreadyApplied=0
        foreach ($row in $Rows) {
            $itemNumber++; $entry=$null; $stage='Reading occurrence'
            Send-TZProgress $ProgressQueue 'Repairing' "Repairing event $($completed+1) of $($Rows.Count)…" $completed $Rows.Count
            $record=if($AllowOccurrences){Find-TZJournalRecord $records $row.Original}else{$null}
            $readId=if($record -and (Get-Field $record.Entry 'AfterId' '')){[string]$record.Entry.AfterId}else{$row.Id}
            $path = '/me/events/' + [uri]::EscapeDataString($readId)
            $fresh = Invoke-ABGet ($path + '?$select=' + $script:Fields)
            if($AllowOccurrences -and ((Get-Field $fresh 'seriesMasterId') -cne $Scan.MasterId -or $fresh.type -notin 'occurrence','exception')){
                throw 'An occurrence no longer belongs to the selected series.'
            }
            if($AllowOccurrences -and $record){
                $resolution=Resolve-TZJournalRecord $record $fresh $target.Zone
                if($resolution -eq 'Applied'){$alreadyApplied++;continue}
                if($resolution -eq 'Unresolved'){throw 'A prior attempt cannot yet be reconciled with this occurrence. No repeat update was sent.'}
            }
            if ((Get-Field $fresh 'changeKey') -cne $row.ChangeKey -or (Get-Field $fresh '@odata.etag') -cne $row.ETag) {
                if(-not $AllowOccurrences -or -not (Test-TZSameOccurrence $fresh $row.Original)){
                    throw 'An event changed after preview. No further events will be updated; scan again.'
                }
            }
            $proposal = New-ABProposal -Event $fresh -TargetZone $target.Zone -TargetAlias $target.Alias -Profile $definition.Key -AllowOccurrence:$AllowOccurrences -IncludeTimeInTitle:([bool](Get-Field $Scan 'IncludeTimeInTitle' $false))
            if (-not $proposal.Eligible -or $proposal.Patch.start.dateTime -ne $row.Patch.start.dateTime -or
                $proposal.Patch.end.dateTime -ne $row.Patch.end.dateTime -or $proposal.Patch.start.timeZone -ne $row.Patch.start.timeZone -or
                $proposal.Patch.end.timeZone -ne $row.Patch.end.timeZone) { throw 'The proposal no longer matches the preview. Scan again.' }
            if([string](Get-Field $proposal.Patch 'subject' '') -cne [string](Get-Field $row.Patch 'subject' '')){throw 'The proposed title no longer matches the preview. Scan again.'}
            $entry = @{ Id = $row.Id; Before = $fresh; Patch = $proposal.Patch; Status = 'Prepared'; At = [datetimeoffset]::UtcNow }
            $journal.Items.Add($entry)
            Save-ABJournal $journal $journalPath
            $entry.Status = 'Sending - outcome may be unknown'
            Save-ABJournal $journal $journalPath
            $headers = @{ Prefer = $script:Headers.Prefer; 'If-Match' = [string](Get-Field $fresh '@odata.etag') }
            $uri = $script:GraphRoot + $path
            Assert-ABUri $uri
            # Preserve the same event ID. All-day events are excluded;
            # no per-day creation, recurrence conversion, or delete is performed.
            $stage='Updating occurrence'
            for($attempt=0;$attempt -lt 4;$attempt++){
                try{
                    $updated = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Headers $headers -Body ($proposal.Patch | ConvertTo-Json -Depth 5) -ContentType 'application/json' -OutputType PSObject
                    break
                }catch{
                    $rejected=Get-TZErrorDetails $_
                    $status=[string](Get-Field $rejected 'HttpStatus' '')
                    $is412=$status -in '412','PreconditionFailed','Precondition Failed'
                    # Retry only a definite server rejection, never a timeout, 5xx or uncertain result.
                    if(-not $AllowOccurrences -or -not $is412){throw}
                    if(-not $entry.ContainsKey('RejectedAttempts')){$entry.RejectedAttempts=[System.Collections.Generic.List[object]]::new()}
                    $entry.RejectedAttempts.Add($rejected)
                    $entry.Status='Rejected - version conflict'
                    Save-ABJournal $journal $journalPath
                    if($attempt -ge 3){throw}
                    Send-TZProgress $ProgressQueue 'Repairing' "Exchange version conflict on occurrence $itemNumber. Rechecking before retry $($attempt+1) of 3…" $completed $Rows.Count
                    Start-Sleep -Milliseconds (500*[math]::Pow(2,$attempt))
                    $candidate=Invoke-ABGet ($path+'?$select='+$script:Fields)
                    if($candidate.type -notin 'occurrence','exception' -or -not(Test-TZSameOccurrence $candidate $entry.Before)){
                        throw 'The occurrence changed while Exchange was processing the series. No repeat update was sent. Preview the series again.'
                    }
                    $retryProposal=New-ABProposal -Event $candidate -TargetZone $target.Zone -TargetAlias $target.Alias -Profile $definition.Key -AllowOccurrence -IncludeTimeInTitle:([bool](Get-Field $Scan 'IncludeTimeInTitle' $false))
                    if(-not $retryProposal.Eligible -or $retryProposal.Patch.start.dateTime -cne $row.Patch.start.dateTime -or
                        $retryProposal.Patch.end.dateTime -cne $row.Patch.end.dateTime -or
                        $retryProposal.Patch.start.timeZone -cne $row.Patch.start.timeZone -or
                        $retryProposal.Patch.end.timeZone -cne $row.Patch.end.timeZone){
                        throw 'The occurrence no longer matches the confirmed preview. No repeat update was sent.'
                    }
                    $etag=[string](Get-Field $candidate '@odata.etag' '')
                    if(-not $etag){throw 'Exchange did not return an occurrence version. No repeat update was sent.'}
                    $headers['If-Match']=$etag
                    $entry.LastAttemptBefore=$candidate
                    $entry.Status='Sending - outcome may be unknown'
                    Save-ABJournal $journal $journalPath
                }
            }
            $entry.Status = 'Response received - verifying'
            Save-ABJournal $journal $journalPath
            $afterId=[string](Get-Field $updated 'id' $row.Id)
            $entry.AfterId=$afterId
            Save-ABJournal $journal $journalPath
            $stage='Verifying occurrence'
            $verified = Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($afterId)+'?$select='+$script:Fields)
            $expectedStart = [datetimeoffset]::new([datetime]::Parse($proposal.Patch.start.dateTime, [Globalization.CultureInfo]::InvariantCulture),
                $target.Zone.GetUtcOffset([datetime]::Parse($proposal.Patch.start.dateTime, [Globalization.CultureInfo]::InvariantCulture)))
            $expectedEnd = [datetimeoffset]::new([datetime]::Parse($proposal.Patch.end.dateTime, [Globalization.CultureInfo]::InvariantCulture),
                $target.Zone.GetUtcOffset([datetime]::Parse($proposal.Patch.end.dateTime, [Globalization.CultureInfo]::InvariantCulture)))
            $entry.Verification=@{Start=$verified.start;End=$verified.end;OriginalStartTimeZone=(Get-Field $verified 'originalStartTimeZone');IsAllDay=(Get-Field $verified 'isAllDay')}
            if ((Convert-ABUtc $verified.start) -ne $expectedStart -or (Convert-ABUtc $verified.end) -ne $expectedEnd) {
                Save-ABJournal $journal $journalPath
                throw 'Server verification failed: the scheduled start/end do not match the preview. The actual values are saved in the encrypted journal.'
            }
            if((Get-Field $proposal.Patch 'subject') -and [string](Get-Field $verified 'subject') -cne $proposal.Patch.subject){throw 'Server verification failed: the event title does not match the preview. Keep the journal.'}
            $entry.Status = 'Verified'; $entry.AfterChangeKey = $verified.changeKey
            Save-ABJournal $journal $journalPath
            $completed++
            Send-TZProgress $ProgressQueue 'Repairing' "$completed of $($Rows.Count) events verified" $completed $Rows.Count
        }
        [pscustomobject]@{ Count = $completed; AlreadyAppliedCount=$alreadyApplied; JournalPath = $journalPath }
    } catch {
        $failure=Get-TZErrorDetails $_
        $message = $_.Exception.Message
        if ($journal) {
            $failure.Stage=$stage; $failure.ItemNumber=$itemNumber; $failure.VerifiedCount=$completed
            $failure.AlreadyAppliedCount=$alreadyApplied
            $journal.Failure=$failure
            if($entry){$entry.Failure=$failure}
            $savedError=$false
            try{Save-ABJournal $journal $journalPath;$savedError=$true}catch{}
            $detail=if($failure.ServerDetails){" Microsoft details: $($failure.ServerDetails)"}else{''}
            $savedMessage=if($savedError){"Error details saved in encrypted journal $journalPath. Use Last error to reopen them."}else{"Error details could not be saved. Existing journal: $journalPath. Copy this message before closing it."}
            $nextStep=if($AllowOccurrences){'Open a fresh series preview before further changes.'}else{'Rescan before further changes.'}
            throw "Stopped while $($stage.ToLowerInvariant()) ($itemNumber of $($Rows.Count)): $message$detail $completed verified; $alreadyApplied previously applied. $savedMessage $nextStep Earlier items may already be fixed."
        }
        throw
    } finally { } # Keep the app session until the window closes.
}

function Get-TZSeriesPlan($Config, $Scan, $Row, [Nullable[datetime]]$ThroughDate=$null, $ProgressQueue=$null,
    [string]$JournalDirectory=(Join-Path $env:LOCALAPPDATA 'AlbertaCalendarPilot\Journal')) {
    if($Row.Kind -ne 'seriesMaster'){throw 'Select a recurring-series row.'}
    $definition=Get-TZProfile $Scan.Profile
    $masterId=if($Row.MasterId){$Row.MasterId}else{$Row.Id}
    $identity=Connect-ABGraph $Config Scan $ProgressQueue
    try {
        if($identity.UserId -ne $Scan.Identity.UserId -or $identity.TenantId -ne $Scan.Identity.TenantId){throw 'Sign in as the same account used for the calendar preview.'}
        $target=Get-ABTarget $definition.Key
        if($target.Alias -ne $Scan.Target){throw 'Selected time zone changed. Scan again.'}
        $master=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($masterId)+'?$select='+$script:Fields)
        if($master.type -ne 'seriesMaster' -or $master.isOrganizer -ne $true -or (Get-Field $master 'isCancelled' $false) -or
            (Get-Field $master 'isAllDay' $false) -or $master.originalStartTimeZone -ne $definition.OldId -or
            $master.originalEndTimeZone -ne $definition.OldId){throw 'Only timed recurring meetings you organize in the selected old zone can be adjusted.'}
        if(-not $master.changeKey -or -not (Get-Field $master '@odata.etag')){throw 'The series is missing version information.'}
        $range=$master.recurrence.range
        $source=[TimeZoneInfo]::FindSystemTimeZoneById($definition.OldId)
        $from=[datetimeoffset]::UtcNow
        if($from -lt $definition.Cutover){$from=$definition.Cutover}
        $until=switch($range.type){
            'endDate' {
                $date=[datetime]::ParseExact($range.endDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(1)
                [datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc($date,$source)
            }
            'numbered' {
                $first=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $master.start),$source).DateTime
                $pattern=$master.recurrence.pattern
                $cycles=[long]$range.numberOfOccurrences*[long]$pattern.interval
                $end=switch($pattern.type){
                    {$_ -in 'absoluteYearly','relativeYearly'}{$first.AddYears($cycles+1)}
                    {$_ -in 'absoluteMonthly','relativeMonthly'}{$first.AddMonths($cycles+1)}
                    'weekly' {$first.AddDays($cycles*7+7)}
                    'daily' {$first.AddDays($cycles+2)}
                    default {throw 'Unsupported recurrence pattern.'}
                }
                [datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc($end,$source)
            }
            'noEnd' {
                if($null -eq $ThroughDate){throw 'This series repeats forever. Choose an end date for this update.'}
                [datetimeoffset]::new($ThroughDate.Date.AddDays(1),[timespan]::FromHours($definition.OffsetHours))
            }
            default {throw 'The recurrence end condition could not be read.'}
        }
        if($until -le $from){throw 'This series has no remaining date range to update.'}
        $previous=Get-ABJournalIndex $identity $JournalDirectory
        $records=@(Read-TZJournalRecords $identity $JournalDirectory)
        $path='/me/events/'+[uri]::EscapeDataString($masterId)+'/instances?startDateTime='+[uri]::EscapeDataString($from.ToString('o'))+
            '&endDateTime='+[uri]::EscapeDataString($until.ToString('o'))+'&$top=250&$select='+$script:Fields
        $rows=[System.Collections.Generic.List[object]]::new(); $skipped=0; $last=$null; $alreadyApplied=0; $unresolved=0
        foreach($instance in Get-ABCollection $path $ProgressQueue 'Previewing remaining series occurrences'){
            $start=Convert-ABUtc $instance.start
            if($null -eq $last -or $start -gt $last){$last=$start}
            $record=Find-TZJournalRecord $records $instance
            if($record){
                $resolution=Resolve-TZJournalRecord $record $instance $target.Zone
                if($resolution -eq 'Applied'){$alreadyApplied++;continue}
                if($resolution -eq 'Unresolved'){$unresolved++;continue}
            }
            $proposal=New-ABProposal -Event $instance -TargetZone $target.Zone -TargetAlias $target.Alias -Profile $definition.Key -AllowOccurrence -IncludeTimeInTitle:([bool](Get-Field $Scan 'IncludeTimeInTitle' $false))
            if($proposal.Eligible -and (Get-Field $instance 'seriesMasterId') -ceq $masterId){
                $proposal.MasterId=$masterId; $rows.Add($proposal)
            }else{$skipped++}
        }
        Send-TZProgress $ProgressQueue 'Complete' "Series preview ready · $($rows.Count) future occurrences · $skipped excluded" 1 1
        [pscustomobject]@{Identity=$identity;Target=$target.Alias;Profile=$definition.Key;MasterId=$masterId;MasterChangeKey=$master.changeKey
            MasterETag=(Get-Field $master '@odata.etag');Subject=$master.subject;Rows=@($rows.ToArray()|Sort-Object DisplayStartUtc);SkippedCount=$skipped
            WebLink=(Get-Field (Get-Field $Row 'Original') 'webLink' (Get-Field $master 'webLink' ''))
            AlreadyAppliedCount=$alreadyApplied;UnresolvedCount=$unresolved
            RangeType=$range.type;EndDate=(Get-Field $range 'endDate' '');From=$from;Until=$until;LastOccurrence=$last;ScannedAt=[datetimeoffset]::UtcNow}
    } finally { }
}

function Set-TZOutlookTimezone($Config, $Scan, [string]$JournalDirectory, [switch]$Confirmed, $ProgressQueue=$null) {
    if(-not $Confirmed){throw 'Confirm the Outlook time-zone change before continuing.'}
    $definition=Get-TZProfile $Scan.Profile
    $identity=Connect-ABGraph $Config Settings $ProgressQueue
    $journalPath=''
    try {
        if($identity.UserId -ne $Scan.Identity.UserId -or $identity.TenantId -ne $Scan.Identity.TenantId){throw 'Sign in as the same account used for the preview.'}
        $target=Get-ABTarget $definition.Key
        if($target.Alias -ne $Scan.Target){throw 'The selected time zone changed. Scan again.'}
        Send-TZProgress $ProgressQueue 'Settings' 'Reading your Outlook mailbox time zone…'
        $before=Invoke-ABGet '/me/mailboxSettings?$select=timeZone'
        if($before.timeZone -eq $target.Alias){return [pscustomobject]@{Changed=$false;TimeZone=$target.Alias;JournalPath=''}}
        New-Item -ItemType Directory -Path $JournalDirectory -Force | Out-Null
        $journalPath=Join-Path $JournalDirectory (([guid]::NewGuid().ToString())+'.dpapi')
        $journal=@{Version=3;Kind='OutlookTimezone';Identity=$identity;Profile=$definition.Key;Items=@();Before=$before.timeZone;Target=$target.Alias;Status='Sending - outcome may be unknown';At=[datetimeoffset]::UtcNow}
        Save-ABJournal $journal $journalPath
        Send-TZProgress $ProgressQueue 'Settings' 'Updating your Outlook mailbox time zone…'
        $uri=$script:GraphRoot+'/me/mailboxSettings'
        Assert-ABUri $uri
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body (@{timeZone=$target.Alias}|ConvertTo-Json -Compress) -ContentType 'application/json' -OutputType PSObject | Out-Null
        $after=Invoke-ABGet '/me/mailboxSettings?$select=timeZone'
        if($after.timeZone -ne $target.Alias){throw 'The Outlook time-zone change could not be verified.'}
        $journal.Status='Verified'; Save-ABJournal $journal $journalPath
        Send-TZProgress $ProgressQueue 'Complete' 'Outlook mailbox time zone updated and verified. No meetings were rescheduled.' 1 1
        [pscustomobject]@{Changed=$true;TimeZone=$target.Alias;JournalPath=$journalPath}
    } catch {
        if($journalPath){throw "The settings operation stopped. Check your Outlook setting and journal $journalPath before retrying. $($_.Exception.Message)"}
        throw
    } finally { }
}

function Get-ABDemo([string]$Profile='Alberta') {
    $definition = Get-TZProfile $Profile
    $zone = [TimeZoneInfo]::CreateCustomTimeZone($definition.NewId, [timespan]::FromHours($definition.OffsetHours), $definition.Name, $definition.Name)
    $events = @(
        @{ subject='Local team workshop'; type='singleInstance'; isOrganizer=$true; originalStartTimeZone='Mountain Standard Time' },
        @{ subject='Weekly operations'; type='seriesMaster'; isOrganizer=$true; originalStartTimeZone='Mountain Standard Time' },
        @{ subject='Customer invitation'; type='singleInstance'; isOrganizer=$false; originalStartTimeZone='Mountain Standard Time' },
        @{ subject='Toronto project call'; type='singleInstance'; isOrganizer=$true; originalStartTimeZone='Eastern Standard Time' },
        @{ subject='Summer planning - zone change, no clock shift'; type='singleInstance'; isOrganizer=$true; originalStartTimeZone='Mountain Standard Time'; summer=$true },
        @{ subject='Already corrected'; type='singleInstance'; isOrganizer=$true; originalStartTimeZone=$definition.NewId }
    )
    $rows = foreach ($e in $events) {
        if ($e.originalStartTimeZone -eq 'Mountain Standard Time') { $e.originalStartTimeZone=$definition.OldId }
        $e.id = [guid]::NewGuid().ToString(); $e.changeKey = 'demo'; $e.'@odata.etag' = 'W/"demo"'
        $e.originalEndTimeZone = $e.originalStartTimeZone; $e.isAllDay = $false; $e.isCancelled = $false
        $wall=[datetime]'2026-11-10T09:00:00'
        $offset=[timespan]::FromHours($definition.OffsetHours-1)
        if (Get-Field $e 'summer' $false) { $wall=[datetime]'2027-07-10T09:00:00'; $offset=[timespan]::FromHours($definition.OffsetHours) }
        if ($e.originalStartTimeZone -eq $definition.NewId) { $offset=[timespan]::FromHours($definition.OffsetHours) }
        $utc=[datetimeoffset]::new($wall,$offset)
        $e.start = @{ dateTime = $utc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff'); timeZone = 'UTC' }
        $e.end = @{ dateTime = $utc.AddHours(1).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff'); timeZone = 'UTC' }
        $row = New-ABProposal $e $zone $definition.NewId ([datetimeoffset]'2026-10-29T00:00:00Z') $definition.Key
        $row.Calendar='Calendar'
        if($e.type -eq 'seriesMaster'){
            $row.MasterId=$e.id
            $row.SeriesRecurrence=@{pattern=@{type='weekly';interval=1};range=@{type='endDate';endDate='2027-11-10'}}
        }
        $row
    }
    [pscustomobject]@{ Identity = @{ UserId='demo'; TenantId='demo'; Account='demo@example.invalid' }; Target=$definition.NewId; Profile=$definition.Key; Rows=@($rows); ScannedAt=[datetimeoffset]::UtcNow }
}

. (Join-Path $PSScriptRoot 'SeriesMigration.ps1')
. (Join-Path $PSScriptRoot 'GraphRequestContract.ps1')
function Get-TZResolvedIds($Identity,[string]$JournalDirectory){
    $ids=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    if(Test-Path -LiteralPath $JournalDirectory){
        foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.resolved.dpapi' -File){
            $j=Read-ABJournal $file.FullName
            if($j.Identity.UserId -eq $Identity.UserId -and $j.Identity.TenantId -eq $Identity.TenantId){
                foreach($id in $j.ResolvedIds){[void]$ids.Add([string]$id)}
            }
        }
    }
    return ,$ids
}
function Resolve-TZJournalRow($Identity,$Row,[string]$JournalDirectory,[switch]$Confirmed){
    if(-not $Confirmed){throw 'Confirm that you manually resolved this meeting in Outlook.'}
    $ids=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    if($Row.Id){[void]$ids.Add([string]$Row.Id)}
    if($Row.MasterId){[void]$ids.Add([string]$Row.MasterId)}
    if(-not $ids.Count){throw 'No meeting ID is available.'}
    New-Item -ItemType Directory -Path $JournalDirectory -Force|Out-Null
    $changes=[Collections.Generic.List[object]]::new()
    foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.dpapi' -File){
        $j=Read-ABJournal $file.FullName
        if($j.Identity.UserId -ne $Identity.UserId -or $j.Identity.TenantId -ne $Identity.TenantId -or (Get-Field $j 'Kind' '') -eq 'ManualResolution'){continue}
        if((Get-Field $j 'Kind' '') -eq 'SeriesMigration'){
            if(-not $ids.Contains([string]$j.MasterId)){continue}
            if(Get-Field $j 'NewId' ''){[void]$ids.Add([string]$j.NewId)}
            $changes.Add(@{Path=$file.FullName;Journal=$j;Remove=$true})
        }else{
            $items=@(Get-TZArray $j 'Items');$remaining=@($items|Where-Object {
                $sameSeries=$Row.MasterId -and ([string](Get-Field (Get-Field $_ 'Before' @{}) 'seriesMasterId' '') -ceq $Row.MasterId)
                -not $sameSeries -and -not $ids.Contains([string]$_.Id) -and -not $ids.Contains([string](Get-Field $_ 'AfterId' ''))
            })
            if($items.Count -eq $remaining.Count){continue}
            foreach($item in $items){
                if(-not @($remaining|Where-Object {$_.Id -ceq $item.Id}).Count){
                    [void]$ids.Add([string]$item.Id)
                    if(Get-Field $item 'AfterId' ''){[void]$ids.Add([string]$item.AfterId)}
                }
            }
            $j.Items=$remaining
            $failure=Get-Field $j 'Failure'
            if($failure){
                $number=[int](Get-Field $failure 'ItemNumber' 0)
                if($number -gt 0 -and $number -le $items.Count){
                    $failedItem=$items[$number-1]
                    if($ids.Contains([string]$failedItem.Id) -or $ids.Contains([string](Get-Field $failedItem 'AfterId' ''))){$j.Remove('Failure')}
                    else{
                        for($i=0;$i -lt $remaining.Count;$i++){if($remaining[$i].Id -ceq $failedItem.Id){$failure.ItemNumber=$i+1;break}}
                    }
                }
            }
            $changes.Add(@{Path=$file.FullName;Journal=$j;Remove=($remaining.Count -eq 0)})
        }
    }
    # Preserve original encrypted evidence outside the active journal before edits.
    $backup=Join-Path (Join-Path $JournalDirectory 'ResolvedBackups') ([guid]::NewGuid().ToString('N'))
    if($changes.Count){New-Item -ItemType Directory -Path $backup -Force|Out-Null}
    foreach($change in $changes){Copy-Item -LiteralPath $change.Path -Destination (Join-Path $backup ([IO.Path]::GetFileName($change.Path))) -ErrorAction Stop}
    foreach($change in $changes){
        if($change.Remove){Remove-Item -LiteralPath $change.Path -ErrorAction Stop}
        else{Save-ABJournal $change.Journal $change.Path}
    }
    $marker=@{Version=1;Kind='ManualResolution';Identity=$Identity;Items=@();ResolvedIds=@($ids);ResolvedAt=[datetimeoffset]::UtcNow}
    Save-ABJournal $marker (Join-Path $JournalDirectory ([guid]::NewGuid().ToString()+'.resolved.dpapi'))
    [pscustomobject]@{ResolvedIds=@($ids);ChangedJournals=$changes.Count;BackupPath=$backup}
}
function Clear-TZJournal([string]$JournalDirectory,[switch]$Confirmed){
    if(-not $Confirmed){throw 'Confirm clearing all active journals first.'}
    $root=[IO.Path]::GetFullPath($JournalDirectory).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if(-not(Test-Path -LiteralPath $root)){return [pscustomobject]@{Count=0;BackupPath=''}}
    $rootItem=Get-Item -LiteralPath $root
    if(-not $rootItem.PSIsContainer -or ($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Journal folder must be a regular local directory.'}
    $files=@(Get-ChildItem -LiteralPath $root -Filter '*.dpapi' -File)
    if(-not $files.Count){return [pscustomobject]@{Count=0;BackupPath=''}}
    # Move encrypted files without decrypting: even corrupt or unreadable failures
    # can be cleared, while preserving the original bytes for manual recovery.
    $backupRoot=Join-Path (Split-Path -Parent $root) 'JournalBackups'
    if(Test-Path -LiteralPath $backupRoot){
        if((Get-Item -LiteralPath $backupRoot).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Backup folder cannot be a directory link.'}
    }
    $backup=Join-Path $backupRoot ((Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $backup -Force -ErrorAction Stop|Out-Null
    $count=0
    try{
        foreach($file in $files){
            if([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($file.FullName)) -cne $root -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Unexpected journal file path.'}
            Move-Item -LiteralPath $file.FullName -Destination (Join-Path $backup $file.Name) -ErrorAction Stop
            $count++
        }
    }catch{throw "Journal clearing stopped after $count files. Backup: $backup. Remaining files stay active. $($_.Exception.Message)"}
    [pscustomobject]@{Count=$count;BackupPath=$backup}
}
Export-ModuleMember -Function Read-ABConfig,Get-ABScan,Invoke-ABApply,Get-ABDemo,New-ABProposal,Get-TZProfile,Set-TZOutlookTimezone,Get-TZSeriesPlan,Disconnect-TZSession,Connect-TZSession,Get-TZLastFailure,Get-TZMigrationPlan,Invoke-TZMigration,Clear-TZJournal,Resolve-TZJournalRow,Get-ABJournalIndex
