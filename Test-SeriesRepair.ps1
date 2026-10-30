#requires -Version 7.2
param([string]$TestDirectory=(Join-Path $env:TEMP ('TimezoneSeriesTests-'+[guid]::NewGuid())))
$ErrorActionPreference='Stop'
$module=Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
& $module {
    param($TestDirectory)
    $script:checks=0; $script:writes=[System.Collections.Generic.List[string]]::new()
    $script:conflictsRemaining=0;$script:rejectCount=0;$script:waitCount=0;$script:semanticOnConflict=$false
    $script:conflictId='future-2';$script:failureStatus=412
    function script:Start-Sleep {param($Milliseconds);$script:waitCount++}
    function Assert($condition,$message){if(-not $condition){throw "FAILED: $message"};$script:checks++;"PASS: $message"}
    function Must-Fail([scriptblock]$action,$message){$failed=$false;try{&$action|Out-Null}catch{$failed=$true};Assert $failed $message}
    function Copy-Data($data){$data|ConvertTo-Json -Depth 15|ConvertFrom-Json -AsHashtable}
    $script:identity=@{UserId='series-user';TenantId='series-tenant';Account='series@example.invalid'}
    function script:Connect-ABGraph($Config,$Mode,$ProgressQueue){$script:identity}
    function script:Disconnect-MgGraph{}
    function script:Get-ABTarget($Profile){$d=Get-TZProfile $Profile;[pscustomobject]@{Alias=$d.NewId;Zone=[TimeZoneInfo]::FindSystemTimeZoneById($d.NewId)}}
    function Make-Occurrence($id,$date,$kind='occurrence'){
        $d=Get-TZProfile 'Alberta'
        $wall=[datetime]$date
        $utc=[TimeZoneInfo]::ConvertTimeToUtc($wall,[TimeZoneInfo]::FindSystemTimeZoneById($d.OldId))
        @{id=$id;subject=$id;type=$kind;seriesMasterId='master-id';isOrganizer=$true;isAllDay=$false;isCancelled=$false
            originalStartTimeZone=$d.OldId;originalEndTimeZone=$d.OldId;changeKey='v1';'@odata.etag'='W/"v1"'
            start=@{dateTime=$utc.ToString('s');timeZone='UTC'};end=@{dateTime=$utc.AddHours(1).ToString('s');timeZone='UTC'}}
    }
    $script:master=Make-Occurrence 'master-id' '2020-01-10T09:00:00' 'seriesMaster'
    $script:master.changeKey='m1';$script:master.'@odata.etag'='W/"m1"'
    $script:master.recurrence=@{pattern=@{type='weekly';interval=1;daysOfWeek=@('Saturday')};range=@{type='endDate';startDate='2020-01-10';endDate='2099-02-01';recurrenceTimeZone='Mountain Standard Time'}}
    $script:events=@{}
    $script:events['past']=Make-Occurrence 'past' '2020-01-10T09:00:00'
    $script:events['future-1']=Make-Occurrence 'future-1' '2099-01-10T09:00:00'
    $script:events['future-2']=Make-Occurrence 'future-2' '2099-01-17T11:00:00' 'exception'
    $script:events['cancelled']=Make-Occurrence 'cancelled' '2099-01-24T09:00:00';$script:events.cancelled.isCancelled=$true
    $script:events['correct']=Make-Occurrence 'correct' '2099-01-25T09:00:00';$script:events.correct.originalStartTimeZone='Alberta Standard Time';$script:events.correct.originalEndTimeZone='Alberta Standard Time'
    function script:Invoke-ABGet($Path){
        if($Path.StartsWith('/me/events/master-id?')){return Copy-Data $script:master}
        if($Path.StartsWith('/me/events/master-id/instances?')){
            $script:lastRange=$Path
            $first=if($script:events.ContainsKey('future-1')){$script:events.'future-1'}else{$script:events.'materialized-future-1'}
            return @{value=@((Copy-Data $script:events.past),(Copy-Data $first));'@odata.nextLink'='https://graph.microsoft.com/v1.0/me/events/master-id/instances?page=2'}
        }
        if($Path -eq 'https://graph.microsoft.com/v1.0/me/events/master-id/instances?page=2'){
            return @{value=@((Copy-Data $script:events.'future-2'),(Copy-Data $script:events.cancelled),(Copy-Data $script:events.correct))}
        }
        if($Path -match '^/me/events/([^?]+)\?'){
            $id=[uri]::UnescapeDataString($Matches[1]);if($script:events.ContainsKey($id)){return Copy-Data $script:events[$id]}
        }
        throw "Unexpected series read: $Path"
    }
    function script:Invoke-MgGraphRequest {
        param($Method,$Uri,$Headers,$Body,$ContentType,$OutputType)
        if($Method -ne 'PATCH' -or $Uri -notmatch '^https://graph.microsoft.com/v1.0/me/events/([^?]+)$'){throw 'Unexpected mutation'}
        $id=[uri]::UnescapeDataString($Matches[1]);if($id -eq 'master-id' -or $id -eq 'past' -or $id -eq 'cancelled'){throw 'Must not change master, past or cancelled events'}
        if($id -eq $script:conflictId -and ($script:failSecond -or $script:conflictsRemaining -gt 0)){
            $script:conflictsRemaining--;$script:rejectCount++
            if($script:failureStatus -ne 412){
                $other=[InvalidOperationException]::new('Synthetic uncertain failure')
                $other|Add-Member -NotePropertyName ResponseStatusCode -NotePropertyValue $script:failureStatus
                throw $other
            }
            $script:events[$id].changeKey="conflict-$script:rejectCount"
            $script:events[$id].'@odata.etag'='W/"conflict-'+$script:rejectCount+'"'
            if($script:semanticOnConflict){$script:events[$id].start.dateTime='2099-01-17T20:00:00'}
            $ex=[InvalidOperationException]::new('Synthetic second occurrence failure')
            $ex|Add-Member -NotePropertyName ResponseStatusCode -NotePropertyValue 412
            $errorRecord=[Management.Automation.ErrorRecord]::new($ex,'SyntheticGraphError',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
            $errorRecord.ErrorDetails=[Management.Automation.ErrorDetails]::new('{"error":{"code":"ErrorIrresolvableConflict","message":"Synthetic version conflict","innerError":{"request-id":"test-request"}}}')
            throw $errorRecord
        }
        $patch=$Body|ConvertFrom-Json -AsHashtable
        if(($patch.Keys|Sort-Object)-join ',' -ne 'end,start'){throw 'Only occurrence start/end may be patched'}
        $event=$script:events[$id]
        if($Headers['If-Match'] -cne $event.'@odata.etag'){throw 'Must use the latest occurrence precondition'}
        foreach($field in 'start','end'){
            $local=[datetime]::Parse($patch[$field].dateTime,[Globalization.CultureInfo]::InvariantCulture)
            $utc=[TimeZoneInfo]::ConvertTimeToUtc($local,[TimeZoneInfo]::FindSystemTimeZoneById($patch[$field].timeZone))
            $event[$field].dateTime=$utc.ToString('s')
        }
        # Exchange can preserve the original series zone metadata after applying the requested times.
        $event.changeKey='v2';$event.type='exception'
        $script:writes.Add($id)
        if($id -eq 'future-1'){
            $event.id='materialized-future-1';$script:events[$event.id]=$event;$script:events.Remove($id)
            $script:events.'future-2'.changeKey='shared-v2';$script:events.'future-2'.'@odata.etag'='W/"shared-v2"'
        }
        Copy-Data $event
    }
    $scan=[pscustomobject]@{Identity=$script:identity;Target='Alberta Standard Time';Profile='Alberta'}
    $row=[pscustomobject]@{Kind='seriesMaster';MasterId='master-id';Id='future-1'}
    New-Item -ItemType Directory -Path $TestDirectory -Force|Out-Null
    $plan=Get-TZSeriesPlan -Config @{} -Scan $scan -Row $row -JournalDirectory $TestDirectory
    Assert ($plan.Rows.Count -eq 2 -and $plan.SkippedCount -eq 3) 'Preview excludes past, cancelled and corrected occurrences and keeps edited future exceptions'
    Assert ($plan.Until -eq [datetimeoffset]'2099-02-02T07:00:00Z') 'Finite series preview runs through its own end date'
    Assert ($plan.Rows[1].Patch.start.dateTime.StartsWith('2099-01-17T11:00:00')) 'Edited occurrence keeps its individual 11:00 clock time'
    Assert ($script:writes.Count -eq 0) 'Building the series preview makes zero writes'
    $savedRecurrence=$script:master.recurrence|ConvertTo-Json -Compress -Depth 6
    $script:master.recurrence.range=@{type='numbered';startDate='2099-01-10';numberOfOccurrences=3;recurrenceTimeZone='Mountain Standard Time'}
    $script:master.start=(Make-Occurrence 'x' '2099-01-10T09:00:00').start
    $numbered=Get-TZSeriesPlan -Config @{} -Scan $scan -Row $row -JournalDirectory $TestDirectory
    Assert ($numbered.RangeType -eq 'numbered' -and $numbered.Until -gt [datetimeoffset]'2099-01-24T17:00:00Z') 'Numbered recurrence range includes the final scheduled occurrence'
    $script:master.recurrence.range=@{type='noEnd';startDate='2099-01-10';recurrenceTimeZone='Mountain Standard Time'}
    Must-Fail {Get-TZSeriesPlan -Config @{} -Scan $scan -Row $row -JournalDirectory $TestDirectory} 'Repeat-forever series requires an explicit end date'
    $open=Get-TZSeriesPlan -Config @{} -Scan $scan -Row $row -ThroughDate ([datetime]'2099-02-01') -JournalDirectory $TestDirectory
    Assert ($open.Until -eq [datetimeoffset]'2099-02-02T06:00:00Z') 'Open-ended series uses the chosen inclusive regional end date'
    $script:master.recurrence=$savedRecurrence|ConvertFrom-Json -AsHashtable
    $script:master.start=(Make-Occurrence 'x' '2020-01-10T09:00:00').start
    Must-Fail {Invoke-ABApply @{} $plan $plan.Rows $TestDirectory -AllowOccurrences} 'Updating future occurrences requires the separate series confirmation'
    $script:master.changeKey='changed'
    Must-Fail {Invoke-ABApply @{} $plan $plan.Rows $TestDirectory -AllowOccurrences -ConfirmedSeries} 'A series changed after preview is blocked'
    Assert ($script:writes.Count -eq 0) 'Unconfirmed or stale series produces zero writes'
    $script:master.changeKey='m1'
    $script:failSecond=$true
    Must-Fail {Invoke-ABApply @{} $plan $plan.Rows $TestDirectory -AllowOccurrences -ConfirmedSeries} 'A second-occurrence server failure stops the batch'
    Assert ($script:writes.Count -eq 1) 'A failed second update leaves the first verified update intact'
    Assert ($script:rejectCount -eq 4 -and $script:waitCount -eq 3) 'Persistent 412 conflicts stop after three delayed revalidation retries'
    $failureText=Get-TZLastFailure $TestDirectory
    Assert ($failureText.Contains('ErrorIrresolvableConflict') -and $failureText.Contains('test-request') -and $failureText.Contains('412')) 'Encrypted diagnostics retain server reason, request ID and HTTP status'
    Assert ($failureText.Contains('Updating occurrence') -and $failureText.Contains('"VerifiedCount": 1')) 'Reopened error identifies the failing stage and prior verified count'
    $script:failSecond=$false
    $script:conflictsRemaining=2
    $result=Invoke-ABApply @{} $plan $plan.Rows $TestDirectory -AllowOccurrences -ConfirmedSeries
    Assert ($result.Count -eq 1 -and $result.AlreadyAppliedCount -eq 1 -and $script:writes.Count -eq 2) 'Resuming verifies the remaining occurrence without rewriting the first'
    Assert ($script:rejectCount -eq 6 -and $script:waitCount -eq 5) 'Transient 412 conflicts re-read and use fresh versions before succeeding'
    Assert (($script:master.recurrence|ConvertTo-Json -Compress -Depth 6) -eq $savedRecurrence) 'Master recurrence pattern and end conditions remain unchanged'
    Assert ($script:events.past.originalStartTimeZone -eq 'Mountain Standard Time' -and $script:events.cancelled.isCancelled) 'Past meetings and cancellations remain untouched'
    Assert (Test-TZPatchApplied $script:events.'future-2' $plan.Rows[1].Patch ([TimeZoneInfo]::FindSystemTimeZoneById('Alberta Standard Time'))) 'Edited future occurrence is verified by actual scheduled times'
    Assert ($script:events.'future-2'.originalStartTimeZone -eq 'Mountain Standard Time') 'Retained original-zone metadata does not produce a false verification failure'
    $index=Get-ABJournalIndex $script:identity $TestDirectory
    Assert ($index.ContainsKey('future-1') -and $index.ContainsKey('materialized-future-1')) 'Journal protects both original and materialized exception IDs against repeat repairs'
    $retry=Invoke-ABApply @{} $plan $plan.Rows $TestDirectory -AllowOccurrences -ConfirmedSeries
    Assert ($retry.Count -eq 0 -and $retry.AlreadyAppliedCount -eq 2 -and $script:writes.Count -eq 2) 'Retry reconciles previously applied occurrences without sending duplicate updates'
    $remaining=Get-TZSeriesPlan -Config @{} -Scan $scan -Row $row -JournalDirectory $TestDirectory
    Assert ($remaining.Rows.Count -eq 0 -and $remaining.AlreadyAppliedCount -eq 2) 'A new preview excludes already-correct occurrences even when original zone metadata is retained'
    $legacyDir=Join-Path $TestDirectory 'legacy';New-Item -ItemType Directory -Path $legacyDir -Force|Out-Null
    $before=Make-Occurrence 'legacy-event' '2099-01-10T09:00:00'
    $zone=[TimeZoneInfo]::FindSystemTimeZoneById('Alberta Standard Time')
    $legacyProposal=New-ABProposal $before $zone 'Alberta Standard Time' ([datetimeoffset]::UtcNow) 'Alberta' -AllowOccurrence
    $entry=@{Id='legacy-event';Before=(Copy-Data $before);Patch=$legacyProposal.Patch;Status='Response received - verifying';At=[datetimeoffset]::UtcNow}
    $journal=@{Identity=$script:identity;Items=@($entry)};$journalPath=Join-Path $legacyDir 'partial.dpapi';Save-ABJournal $journal $journalPath
    $live=Copy-Data $before;$live.start.dateTime='2099-01-10T15:00:00';$live.end.dateTime='2099-01-10T16:00:00'
    $record=[pscustomobject]@{Entry=$entry;Journal=$journal;Path=$journalPath}
    Assert ((Resolve-TZJournalRecord $record $live $zone) -eq 'Applied' -and $entry.Status -eq 'Verified') 'Old response-received verification failures reconcile as applied when live times match the saved patch'
    Assert ($script:writes.Count -eq 2) 'Reconciling a legacy partial result makes no event writes'
    Assert ((Resolve-TZJournalRecord $record $before $zone) -eq 'Unresolved') 'A previously verified update later reverted by the user is not silently reapplied'
    $entry.Status='Sending - outcome may be unknown';$entry.Remove('AfterId')
    Assert ((Resolve-TZJournalRecord $record $before $zone) -eq 'NotApplied') 'An unchanged occurrence from an uncompleted attempt becomes eligible for a fresh preview'
    Assert ((Get-ABJournalIndex $script:identity $legacyDir).Count -eq 0) 'A proven not-applied attempt no longer blocks the next preview'
    $caseDir=Join-Path $TestDirectory 'case-sensitive';New-Item -ItemType Directory -Path $caseDir -Force|Out-Null
    Save-ABJournal @{Identity=$script:identity;Items=@(@{Id='CaseEvent';Status='Prepared'},@{Id='caseevent';Status='Prepared'})} (Join-Path $caseDir 'case.dpapi')
    Assert ((Get-ABJournalIndex $script:identity $caseDir).Count -eq 2) 'Case-sensitive Graph event IDs cannot falsely block a different occurrence'
    $script:events.'future-2'=Make-Occurrence 'future-2' '2099-01-17T11:00:00' 'exception'
    $script:failSecond=$true;$script:semanticOnConflict=$true
    $rejectsBefore=$script:rejectCount
    Must-Fail {Invoke-ABApply @{} $plan @($plan.Rows[1]) (Join-Path $TestDirectory 'changed-during-conflict') -AllowOccurrences -ConfirmedSeries} 'A real time change during a conflict blocks retry'
    Assert ($script:rejectCount -eq $rejectsBefore+1 -and $script:writes.Count -eq 2) 'A materially changed occurrence receives no repeat PATCH'
    $transcriptException=[InvalidOperationException]::new('Precondition Failed')
    $transcriptRecord=[Management.Automation.ErrorRecord]::new($transcriptException,'SyntheticTranscript',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
    $transcriptRecord.ErrorDetails=[Management.Automation.ErrorDetails]::new("PATCH https://graph.microsoft.com/v1.0/me/events/test`r`nHTTP/2.0 412 Precondition Failed`r`nrequest-id: response-request`r`n`r`n"+'{"error":{"code":"ErrorIrresolvableConflict","message":"Synthetic conflict"}}')
    $parsed=Get-TZErrorDetails $transcriptRecord
    Assert ($parsed.GraphCode -eq 'ErrorIrresolvableConflict' -and $parsed.RequestId -eq 'response-request') 'Real SDK HTTP transcript errors expose Graph code and request ID'
    foreach($failureStatus in 500,0){
        $script:failureStatus=$failureStatus;$script:semanticOnConflict=$false
        $script:events.'future-2'=Make-Occurrence 'future-2' '2099-01-17T11:00:00' 'exception'
        $rejectsBefore=$script:rejectCount;$waitsBefore=$script:waitCount
        Must-Fail {Invoke-ABApply @{} $plan @($plan.Rows[1]) (Join-Path $TestDirectory "uncertain-$failureStatus") -AllowOccurrences -ConfirmedSeries} "Uncertain failure $failureStatus stops the batch"
        Assert ($script:rejectCount -eq $rejectsBefore+1 -and $script:waitCount -eq $waitsBefore) "Uncertain failure $failureStatus is never automatically retried"
    }
    $script:events.'future-1'=Make-Occurrence 'future-1' '2099-01-10T09:00:00'
    $script:events.'future-2'=Make-Occurrence 'future-2' '2099-01-17T11:00:00' 'exception'
    $script:events.'future-3'=Make-Occurrence 'future-3' '2099-01-24T09:00:00'
    $third=New-ABProposal $script:events.'future-3' $zone 'Alberta Standard Time' ([datetimeoffset]::UtcNow) 'Alberta' -AllowOccurrence
    $script:conflictId='future-3';$script:failureStatus=412;$script:failSecond=$false;$script:conflictsRemaining=1
    $writesBefore=$script:writes.Count
    $three=Invoke-ABApply @{} $plan @($plan.Rows[0],$plan.Rows[1],$third) (Join-Path $TestDirectory 'third-conflict') -AllowOccurrences -ConfirmedSeries
    Assert ($three.Count -eq 3 -and $script:writes.Count -eq $writesBefore+3) 'A third-occurrence version conflict recovers and verifies all three without duplicate writes'
    "All $script:checks series-repair checks passed. No network requests were made."
} $TestDirectory
Remove-Module $module
