#requires -Version 7.2
param([string]$TestDirectory=(Join-Path $env:TEMP ('TimezoneMigrationTests-'+[guid]::NewGuid())))
$ErrorActionPreference='Stop'
$module=Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
& $module {
    param($TestDirectory)
    $script:checks=0
    function Assert($condition,$message){if(-not $condition){throw "FAILED: $message"};$script:checks++;"PASS: $message"}
    function Must-Fail([scriptblock]$action,$message){$failed=$false;try{&$action|Out-Null}catch{$failed=$true};Assert $failed $message}
    function Throw-Http($Status){$ex=[InvalidOperationException]::new("Synthetic HTTP $Status");$ex|Add-Member ResponseStatusCode $Status;throw $ex}
    $script:identity=@{UserId='migration-user';TenantId='tenant';Account='migration@example.invalid'}
    function script:Connect-ABGraph($Config,$Mode,$ProgressQueue){$script:identity}
    function script:Get-ABTarget($Profile){[pscustomobject]@{Alias='Alberta Standard Time';Zone=[TimeZoneInfo]::FindSystemTimeZoneById('Alberta Standard Time')}}
    function script:Start-Sleep {param($Seconds,$Milliseconds)}
    function Reset-Fixture {
        $script:writes=[Collections.Generic.List[string]]::new();$script:newEvent=$null;$script:deleted=$false
        $script:failEnd=$false;$script:postError=0;$script:unknownPost=$false
        $script:wrongNewZone=$false;$script:wrongNewDates=$false
        $script:unknownEnd=$false
        $script:replacementInstances=$null
        $script:teamsReady=$true;$script:inviteLost=$false;$script:invitationWrites=0;$script:createBody=$null
        $script:master=@{id='master';type='seriesMaster';subject='Synthetic weekly meeting';isOrganizer=$true;isAllDay=$false;isCancelled=$false;
            originalStartTimeZone='Mountain Standard Time';originalEndTimeZone='Mountain Standard Time';'@odata.etag'='W/"v1"';changeKey='v1';
            start=@{dateTime='2020-01-04T16:00:00';timeZone='UTC'};end=@{dateTime='2020-01-04T17:00:00';timeZone='UTC'};
            recurrence=@{pattern=@{type='weekly';interval=1;daysOfWeek=@('saturday');firstDayOfWeek='sunday';month=0;dayOfMonth=0};range=@{type='endDate';startDate='2020-01-04';endDate='2099-01-24';recurrenceTimeZone='Mountain Standard Time'}};
            cancelledOccurrences=@();exceptionOccurrences=@();hasAttachments=$false;body=@{contentType='text';content='Synthetic description'};
            attendees=@(@{emailAddress=@{address='person@example.invalid';name='Test Person'};type='required'},@{emailAddress=@{address='room@example.invalid';name='Test Room'};type='resource'});
            location=@{displayName='Test Room'};isOnlineMeeting=$false;onlineMeetingProvider='unknown'}
        $script:instances=@('2099-01-10','2099-01-17','2099-01-24'|ForEach-Object {
            $e=Copy-TZData $script:master;$e.id=$_;$e.type='occurrence';$e.seriesMasterId='master';$e.start=@{dateTime=$_+'T16:00:00';timeZone='UTC'};$e.end=@{dateTime=$_+'T17:00:00';timeZone='UTC'};$e.originalStart=$_+'T16:00:00Z';$e.recurrence=$null;$e
        })
        $script:calendar=@{id='calendar';name='Calendar';owner=@{address='migration@example.invalid'};canEdit=$true}
    }
    function script:Invoke-ABGet($Path){
        if($Path -match 'isSharedWithMe'){throw 'RequestBroker-ParseUri: unsupported calendar property'}
        if($Path -match '^/me/events/new-series/instances\?'){
            $serverInstances=if($null -ne $script:replacementInstances){$script:replacementInstances}else{$script:instances}
            $items=@($serverInstances|ForEach-Object {
                $e=Copy-TZData $_;$day=([datetime]$e.originalStart).ToString('yyyy-MM-dd')
                if($script:wrongNewDates){$day=([datetime]$day).AddDays(1).ToString('yyyy-MM-dd')}
                $e.start=@{dateTime=$day+'T15:00:00';timeZone='UTC'};$e.end=@{dateTime=$day+'T16:00:00';timeZone='UTC'};$e
            })
            return @{value=$items}
        }
        if($Path -match '^/me/events/master/calendar\?'){return Copy-TZData $script:calendar}
        if($Path -match '^/me/events/master/instances\?'){return @{value=(Copy-TZData $script:instances)}}
        if($Path -match '^/me/events/master\?'){if($script:deleted){Throw-Http 404};return Copy-TZData $script:master}
        if($Path -match '^/me/events/new-series\?'){
            if(-not $script:newEvent){Throw-Http 404}
            $copy=Copy-TZData $script:newEvent
            if($Path -match 'onlineMeeting,body' -and $script:teamsReady){
                $copy.onlineMeeting=@{joinUrl='https://teams.microsoft.com/meet/synthetic-test'}
                $copy.body.content+="`nhttps://teams.microsoft.com/meet/synthetic-test"
            }
            return $copy
        }
        if($Path -match '^/me/calendars/calendar/events\?'){return @{value=@(if($script:newEvent){Copy-TZData $script:newEvent})}}
        throw "Unexpected read: $Path"
    }
    function script:Invoke-MgGraphRequest {
        param($Method,$Uri,$Headers,$Body,$ContentType,$OutputType)
        $script:writes.Add($Method)
        if($Method -eq 'POST'){
            if($script:postError){Throw-Http $script:postError}
            if($script:newEvent){throw 'A duplicate POST was attempted'}
            $script:newEvent=$Body|ConvertFrom-Json -AsHashtable
            $script:createBody=Copy-TZData $script:newEvent
            if($script:newEvent.attendees.Count -ne 0){throw 'Replacement preparation must not invite anyone'}
            if($script:newEvent.recurrence.range.type -eq 'noEnd'){$script:newEvent.recurrence.range.endDate='0001-01-01'}
            $script:newEvent.recurrence.range.numberOfOccurrences=0
            if($script:newEvent.recurrence.pattern.type -in 'weekly','relativeMonthly','relativeYearly' -and $script:newEvent.recurrence.pattern.daysOfWeek -isnot [array]){throw 'Exchange contract: daysOfWeek must be an array, not a primitive value'}
            $script:newEvent.id='new-series';$script:newEvent.type='seriesMaster';$script:newEvent.isOrganizer=$true
            $script:newEvent.'@odata.etag'='W/"new-v1"'
            $script:newEvent.originalStartTimeZone=$script:newEvent.start.timeZone;$script:newEvent.originalEndTimeZone=$script:newEvent.end.timeZone
            if($script:wrongNewZone){$script:newEvent.originalStartTimeZone='Mountain Standard Time'}
            foreach($field in 'start','end'){
                $wall=[datetime]$script:newEvent[$field].dateTime
                $utc=[TimeZoneInfo]::ConvertTimeToUtc($wall,[TimeZoneInfo]::FindSystemTimeZoneById($script:newEvent[$field].timeZone))
                $script:newEvent[$field]=@{dateTime=$utc.ToString('s');timeZone='UTC'}
            }
            if($script:unknownPost){Throw-Http 500}
            return Copy-TZData $script:newEvent
        }
        if($Method -eq 'PATCH' -and $Uri -eq 'https://graph.microsoft.com/v1.0/me/events/new-series'){
            if(-not $script:deleted -and $script:master.recurrence.range.endDate -ne '2099-01-09'){throw 'Rooms must be added only after ending the original bookings'}
            $patch=$Body|ConvertFrom-Json -AsHashtable
            if(($patch.Keys -join ',') -ne 'attendees' -or $Headers['If-Match'] -cne $script:newEvent.'@odata.etag'){throw 'Room update must only patch attendees with a fresh version'}
            $script:invitationWrites++
            $script:newEvent.attendees=$patch.attendees
            if($script:inviteLost){Throw-Http 500}
            return Copy-TZData $script:newEvent
        }
        if($Uri -ne 'https://graph.microsoft.com/v1.0/me/events/master' -or $Headers['If-Match'] -cne $script:master.'@odata.etag'){throw 'Original mutation must be conditional and address only the old master'}
        if($script:failEnd){$script:failEnd=$false;Throw-Http 412}
        if($Method -eq 'PATCH'){
            $patch=$Body|ConvertFrom-Json -AsHashtable
            if(($patch.Keys -join ',') -ne 'recurrence'){throw 'Only the recurrence end may change on a historical series'}
            $script:master.recurrence=$patch.recurrence;if($script:unknownEnd){Throw-Http 500};return Copy-TZData $script:master
        }
        if($Method -eq 'DELETE'){$script:deleted=$true;if($script:unknownEnd){Throw-Http 500};return}
        throw 'Unexpected mutation'
    }
    New-Item -ItemType Directory -Path $TestDirectory -Force|Out-Null
    $scan=[pscustomobject]@{Identity=$script:identity;Profile='Alberta';Target='Alberta Standard Time'}
    $row=[pscustomobject]@{Kind='seriesMaster';MasterId='master';Id='master'}
    Reset-Fixture
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'normal')
    Assert ($script:writes.Count -eq 0) 'Series migration preview is read-only'
    Assert ($plan.Operation -eq 'SplitSeries' -and $plan.OldEnd -eq '2099-01-09') 'Historical series ends before the first replacement date'
    Assert ($plan.NewPayload.start.dateTime -eq '2099-01-10T09:00:00' -and $plan.NewPayload.recurrence.range.recurrenceTimeZone -eq 'Alberta Standard Time') 'Replacement preserves the 09:00 clock time and uses the new recurrence zone'
    Assert ($plan.NewPayload.attendees.Count -eq 2 -and $plan.NewPayload.attendees[1].type -eq 'resource') 'Replacement preserves people and resource attendee types'
    Assert (-not $plan.NewPayload.recurrence.pattern.ContainsKey('month')) 'Only applicable recurrence pattern fields are sent'
    Must-Fail {Invoke-TZMigration @{} $plan (Join-Path $TestDirectory 'normal')} 'Migration requires explicit confirmation'
    Assert ($script:writes.Count -eq 0) 'Unconfirmed migration produces no writes'
    $result=Invoke-TZMigration @{} $plan (Join-Path $TestDirectory 'normal') -Confirmed
    Assert (($script:writes -join ',') -eq 'POST,PATCH,PATCH' -and $result.MigratedSeries) 'Confirmed split verifies replacement, ends old bookings, then adds resource attendees'
    Assert ($script:master.start.dateTime -eq '2020-01-04T16:00:00') 'Historical master start remains unchanged'
    Assert ($script:newEvent.originalStartTimeZone -eq 'Alberta Standard Time' -and $script:newEvent.recurrence.range.recurrenceTimeZone -eq 'Alberta Standard Time') 'Replacement verifies both original zone labels and recurrence zone'
    Reset-Fixture;$titleScan=[pscustomobject]@{Identity=$script:identity;Profile='Alberta';Target='Alberta Standard Time';IncludeTimeInTitle=$true}
    $dir=Join-Path $TestDirectory 'series-title';$plan=Get-TZMigrationPlan @{} $titleScan $row $dir
    Assert ($plan.Subject -eq 'Synthetic weekly meeting (09:00AM)' -and $plan.Rows[0].Subject -eq $plan.NewPayload.subject) 'Series confirmation previews the original regular start time in the replacement title'
    $null=Invoke-TZMigration @{} $plan $dir -Confirmed
    Assert ($script:createBody.subject -eq 'Synthetic weekly meeting (09:00AM)' -and $script:invitationWrites -eq 1 -and $script:master.subject -eq 'Synthetic weekly meeting') 'Replacement includes the title before invitations; historical original title is unchanged'
    Reset-Fixture;$script:master.start.dateTime='2099-01-10T16:00:00';$script:master.end.dateTime='2099-01-10T17:00:00';$script:master.recurrence.range.startDate='2099-01-10'
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'future-only')
    Assert ($plan.Operation -eq 'ReplaceFutureSeries') 'Entirely future series is identified for replacement and cancellation'
    $null=Invoke-TZMigration @{} $plan (Join-Path $TestDirectory 'future-only') -Confirmed
    Assert (($script:writes -join ',') -eq 'POST,DELETE,PATCH' -and $script:deleted) 'Entirely future replacement cancels only the confirmed original series before adding rooms'
    Reset-Fixture;$script:master.cancelledOccurrences=@('cancelled-id')
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'cancelled')} 'Cancelled occurrences block automatic splitting'
    Assert ($script:writes.Count -eq 0) 'Cancellation protection produces zero writes'
    Reset-Fixture;$script:master.hasAttachments=$true
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'attachments')} 'Attachments require Outlook handling'
    Reset-Fixture;$script:calendar.owner.address='other@example.invalid'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'shared')} 'Shared calendars cannot be split'
    Reset-Fixture
    $exception=Copy-TZData $script:instances[0];$exception.type='exception';$exception.id='historical-edit'
    $exception.originalStart='2020-01-11T16:00:00Z'
    $exception.start.dateTime='2020-01-11T17:00:00';$exception.end.dateTime='2020-01-11T18:00:00'
    $exception.subject='Custom historical subject';$exception.body.content='Historical notes';$exception.location.displayName='Different historical location'
    $script:master.exceptionOccurrences=@($exception)
    $historicalSnapshot=ConvertTo-Json -InputObject $exception -Depth 30 -Compress
    $dir=Join-Path $TestDirectory 'historical-exception'
    $plan=Get-TZMigrationPlan @{} $scan $row $dir
    Assert ($plan.RetainedExceptionCount -eq 1 -and $script:writes.Count -eq 0) 'Past edits with custom content and location allow a read-only preview'
    $null=Invoke-TZMigration @{} $plan $dir -Confirmed
    Assert (($script:writes -join ',') -eq 'POST,PATCH,PATCH' -and (ConvertTo-Json -InputObject $script:master.exceptionOccurrences[0] -Depth 30 -Compress) -ceq $historicalSnapshot) 'Historical split keeps past exception data and sends no exception edits'
    Reset-Fixture;$script:master.exceptionOccurrences=@(Copy-TZData $exception)
    $script:master.exceptionOccurrences[0].start.dateTime='2099-01-10T16:00:00';$script:master.exceptionOccurrences[0].end.dateTime='2099-01-10T17:00:00'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'past-moved-future')} 'A historical slot moved into the future still blocks migration'
    Assert ($script:writes.Count -eq 0) 'Boundary-crossing exception is blocked before writes'
    Reset-Fixture;$script:master.exceptionOccurrences=@(Copy-TZData $exception);$script:master.exceptionOccurrences[0].Remove('originalStart')
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'missing-original')} 'A missing original exception date still blocks migration'
    $script:master.exceptionOccurrences[0].originalStart='not-a-date'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'invalid-original')} 'An invalid original exception date still blocks migration'
    Reset-Fixture
    $script:replacementInstances=Copy-TZData $script:instances
    $changed=Copy-TZData $script:instances[1];$changed.type='exception';$changed.subject='Custom future subject'
    $changed.body.content='Custom future notes';$changed.location.displayName='Custom venue';$changed.attendees=@()
    $changed.start.dateTime='2099-01-17T18:00:00';$changed.end.dateTime='2099-01-17T20:00:00'
    $script:instances=@($changed);$script:master.exceptionOccurrences=@($changed)
    $script:master.cancelledOccurrences=@('OID.master.2099-01-10','OID.master.2099-01-24')
    $dir=Join-Path $TestDirectory 'recreate-future'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row $dir} 'Default mode still blocks future custom edits and cancellations'
    $plan=Get-TZMigrationPlan @{} $scan $row $dir -RecreateFutureChanges
    Assert ($plan.FirstDate -eq '2099-01-10' -and $plan.ExpectedStarts.Count -eq 3 -and $plan.RecreatedCancellationCount -eq 2 -and $plan.DiscardedExceptionCount -eq 1) 'Opt-in preview restores cancelled first and last dates and resets a custom future exception'
    Assert ($plan.Rows[0].CurrentLocal -like 'Cancelled*' -and $plan.NewPayload.body.content -eq 'Synthetic description' -and $plan.NewPayload.attendees.Count -eq 2 -and $script:writes.Count -eq 0) 'Preview marks restored dates and uses master content and attendees without writing'
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'General confirmation alone cannot discard future changes'
    Assert ($script:writes.Count -eq 0) 'Missing discard acknowledgement produces no writes'
    $plan.DiscardChangesConfirmed=$true
    $null=Invoke-TZMigration @{} $plan $dir -Confirmed
    Assert (($script:writes -join ',') -eq 'POST,PATCH,PATCH') 'Explicit discard confirmation survives fresh validation and completes the migration'
    $file=Get-ChildItem -LiteralPath $dir -Filter *.dpapi|Select-Object -First 1
    Assert ((Read-ABJournal $file.FullName).Plan.DiscardChangesConfirmed) 'Recovery journal records explicit discard consent'
    Reset-Fixture;$script:master.cancelledOccurrences=@('unknown')
    $allDayException=Copy-TZData $script:instances[0];$allDayException.type='exception';$allDayException.isAllDay=$true
    $script:master.exceptionOccurrences=@($allDayException)
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'recreate-all-day-exception') -RecreateFutureChanges} 'Warning-based recreation never overrides the all-day exclusion'
    Reset-Fixture;$script:master.cancelledOccurrences=@('unknown')
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'recreate-unknown') -RecreateFutureChanges} 'Recreate mode does not guess unidentifiable cancellation dates'
    Reset-Fixture;$script:replacementInstances=Copy-TZData $script:instances
    $script:master.cancelledOccurrences=@('OID.master.2099-01-10','OID.master.2099-01-17','OID.master.2099-01-24');$script:instances=@()
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'all-cancelled') -RecreateFutureChanges
    Assert ($plan.ExpectedStarts.Count -eq 3 -and $plan.RecreatedCancellationCount -eq 3) 'All cancelled future dates can be previewed even with no surviving instances'
    Reset-Fixture;$exception=Copy-TZData $script:instances[0];$exception.type='exception';$script:master.exceptionOccurrences=@($exception)
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'custom')} 'Unknown custom exceptions block splitting'
    Reset-Fixture
    $dir=Join-Path $TestDirectory 'resume';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:failEnd=$true
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'A failure ending the original is saved after replacement verification'
    Assert (@(Get-TZMigrationRecoveryRows $script:identity $dir 'Alberta').Count -eq 1) 'An incomplete migration remains available for recovery after a fresh scan or restart'
    $resume=Get-TZMigrationPlan @{} $scan $row $dir
    Assert ([bool]$resume.ResumeJournalPath) 'A new preview identifies the existing replacement for resumption'
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert (@($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1) 'Resuming never creates a second replacement or duplicate invitation set'
    Assert (@(Get-TZMigrationRecoveryRows $script:identity $dir 'Alberta').Count -eq 0) 'Completed migrations leave the recovery list'
    Reset-Fixture
    $dir=Join-Path $TestDirectory 'unknown-post';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:unknownPost=$true
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'An uncertain POST stops without ending the old series'
    Assert (($script:writes -join ',') -eq 'POST' -and $script:master.recurrence.range.endDate -eq '2099-01-24') 'Unknown creation leaves old recurrence intact'
    $resume=Get-TZMigrationPlan @{} $scan $row $dir;$script:unknownPost=$false
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert (@($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1) 'Transaction discovery reconciles a lost creation response without another POST'
    Reset-Fixture
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'stale');$script:master.subject='Changed after preview'
    Must-Fail {Invoke-TZMigration @{} $plan (Join-Path $TestDirectory 'stale') -Confirmed} 'Changed series content requires a new preview'
    Assert ($script:writes.Count -eq 0) 'A stale migration preview creates no replacement'
    Reset-Fixture;$script:master.recurrence.range.type='noEnd'
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'forever')
    Assert ($plan.NewPayload.recurrence.range.type -eq 'noEnd' -and -not $plan.NewPayload.recurrence.range.ContainsKey('endDate')) 'Repeat-forever series stays open-ended in the new zone'
    $dir=Join-Path $TestDirectory 'noend-server-defaults'
    $script:failEnd=$true
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'Open-ended replacement verifies through server defaults before a simulated interruption ending the old series'
    Assert ($script:newEvent.recurrence.range.endDate -eq '0001-01-01' -and ($script:writes -join ',') -eq 'POST,PATCH') 'Server sentinel end date does not cause a false time-zone rejection'
    $resume=Get-TZMigrationPlan @{} $scan $row $dir
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert (@($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1) 'Open-ended interrupted migration resumes without creating a second replacement'
    $script:newEvent.recurrence.range.type='endDate'
    Assert (-not(Test-TZReplacement $script:newEvent $plan.NewPayload (Get-ABTarget 'Alberta').Zone)) 'An actually changed recurrence range type still fails verification'
    Reset-Fixture;$script:master.recurrence.range=@{type='numbered';startDate='2020-01-04';numberOfOccurrences=5000;recurrenceTimeZone='Mountain Standard Time'}
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'numbered')
    Assert ($plan.NewPayload.recurrence.range.endDate -eq '2099-01-24') 'Numbered series retains its final scheduled date'
    Reset-Fixture
    $dir=Join-Path $TestDirectory 'prior-adjustment';New-Item -ItemType Directory -Path $dir -Force|Out-Null
    $before=Copy-TZData $script:instances[0];$adjusted=Copy-TZData $before;$adjusted.type='exception';$adjusted.start.dateTime='2099-01-10T15:00:00';$adjusted.end.dateTime='2099-01-10T16:00:00'
    $script:instances[0]=$adjusted;$script:master.exceptionOccurrences=@($adjusted)
    Save-ABJournal @{Identity=$script:identity;Items=@(@{Id=$before.id;Before=$before;Status='Verified';Patch=@{start=@{dateTime='2099-01-10T09:00:00';timeZone='Alberta Standard Time'};end=@{dateTime='2099-01-10T10:00:00';timeZone='Alberta Standard Time'}}})} (Join-Path $dir 'prior.dpapi')
    $plan=Get-TZMigrationPlan @{} $scan $row $dir
    Assert ($plan.NewPayload.start.dateTime -eq '2099-01-10T09:00:00') 'A journal-verified prior occurrence move retains 09:00 without a second one-hour shift'
    $script:master.isOnlineMeeting=$true;$script:master.onlineMeetingProvider='teamsForBusiness';$script:master.exceptionOccurrences=@();$script:instances[0]=$before
    $script:master.body=@{contentType='html';content='<p>Agenda to retain</p><p>Microsoft Teams meeting</p><p>Old joining details</p>'}
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'online')
    Assert ($plan.ReplacementDescription.Trim() -eq 'Agenda to retain' -and $plan.NewPayload.isOnlineMeeting) 'Teams replacement drafts omit old joining details and request a new online meeting'
    $plan.ReplacementDescription='Reviewed replacement agenda'
    $null=Invoke-TZMigration @{} $plan (Join-Path $TestDirectory 'online') -Confirmed
    Assert ($script:newEvent.body.content -eq 'Reviewed replacement agenda') 'Confirmed online description edits are included in replacement creation'
    Assert ($script:createBody.attendees.Count -eq 0 -and $script:createBody.body.content -eq 'Reviewed replacement agenda' -and $script:invitationWrites -eq 1 -and $script:newEvent.attendees.Count -eq 2) 'Description is prepared without recipients, then people and rooms are invited in one attendee-only update'
    Reset-Fixture;$script:master.isOnlineMeeting=$true;$script:master.onlineMeetingProvider='teamsForBusiness';$script:teamsReady=$false
    $dir=Join-Path $TestDirectory 'teams-pending';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'Unfinished Teams provisioning stops before ending the original or inviting recipients'
    Assert (($script:writes -join ',') -eq 'POST' -and $script:invitationWrites -eq 0 -and $script:newEvent.attendees.Count -eq 0) 'Teams wait timeout leaves an uninvited replacement and unchanged original'
    $script:teamsReady=$true;$resume=Get-TZMigrationPlan @{} $scan $row $dir
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert (@($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1 -and $script:invitationWrites -eq 1) 'Teams-ready resume reuses the prepared replacement and invites once'
    Reset-Fixture;$script:inviteLost=$true
    $dir=Join-Path $TestDirectory 'lost-invitation-response';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'A lost invitation response is journaled for reconciliation'
    $script:inviteLost=$false;$resume=Get-TZMigrationPlan @{} $scan $row $dir
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert ($script:invitationWrites -eq 1) 'Resume checks existing attendees and does not resend an already-applied invitation update'
    Reset-Fixture;$script:failEnd=$true
    $dir=Join-Path $TestDirectory 'legacy-invited-resume';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'Legacy recovery fixture stops after preparation'
    $file=Get-ChildItem -LiteralPath $dir -Filter *.dpapi|Select-Object -First 1
    $saved=Read-ABJournal $file.FullName;$saved.Remove('InitialAttendeesMode');$saved.Version=4;Save-ABJournal $saved $file.FullName
    $script:newEvent.attendees=@($plan.NewPayload.attendees|Where-Object type -ne 'resource')
    $resume=Get-TZMigrationPlan @{} $scan $row $dir
    $null=Invoke-TZMigration @{} $resume $dir -Confirmed
    Assert (@($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1 -and $script:newEvent.attendees.Count -eq 2) 'Older journals with people already invited can resume without replacement recreation'
    Reset-Fixture;$dir=Join-Path $TestDirectory 'wrong-zone';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:wrongNewZone=$true
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'An old-zone replacement fails verification'
    Assert (($script:writes -join ',') -eq 'POST') 'The original is not ended when replacement zone verification fails'
    Reset-Fixture;$dir=Join-Path $TestDirectory 'wrong-dates';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:wrongNewDates=$true
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'Replacement recurrence dates must match the reviewed schedule'
    Assert (($script:writes -join ',') -eq 'POST') 'A shifted replacement recurrence never triggers the original end operation'
    Reset-Fixture;$dir=Join-Path $TestDirectory 'rejected-create';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:postError=400
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'A definite creation rejection is saved before any original change'
    Assert ($null -eq (Get-TZPendingMigration $script:identity 'master' $dir)) 'A definite non-creation does not force an uncertain-outcome lock'
    Reset-Fixture;$dir=Join-Path $TestDirectory 'expired';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$plan.ScannedAt=[datetimeoffset]::UtcNow.AddMinutes(-31)
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'Expired migration plans require a fresh preview'
    Assert ($script:writes.Count -eq 0) 'Expired plans make zero writes'
    foreach($futureOnly in $false,$true){
        Reset-Fixture
        if($futureOnly){$script:master.start.dateTime='2099-01-10T16:00:00';$script:master.end.dateTime='2099-01-10T17:00:00';$script:master.recurrence.range.startDate='2099-01-10'}
        $dir=Join-Path $TestDirectory "lost-end-$futureOnly";$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:unknownEnd=$true
        Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} "A lost original-end response is retained (future-only: $futureOnly)"
        $resume=Get-TZMigrationPlan @{} $scan $row $dir;$script:unknownEnd=$false
        $null=Invoke-TZMigration @{} $resume $dir -Confirmed
        Assert ($script:writes.Count -eq 3 -and @($script:writes|Where-Object {$_ -eq 'POST'}).Count -eq 1) "Reconciliation skips the already-completed end/cancellation operation (future-only: $futureOnly)"
    }
    Reset-Fixture;$dir=Join-Path $TestDirectory 'missing-unknown-create';$plan=Get-TZMigrationPlan @{} $scan $row $dir;$script:postError=500
    Must-Fail {Invoke-TZMigration @{} $plan $dir -Confirmed} 'A server failure during creation remains uncertain'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row $dir} 'An unidentified creation blocks another POST'
    Assert ($script:writes.Count -eq 1) 'An unknown creation never sends a second invitation set automatically'
    Reset-Fixture;$script:master.cancelledOccurrences=$null;$script:master.exceptionOccurrences=$null;$script:master.attendees=$null
    $plan=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'null-collections')
    Assert ($plan.NewPayload.attendees.Count -eq 0 -and $script:writes.Count -eq 0) 'Null cancellations, exceptions and attendees are treated as empty collections in a read-only preview'
    Reset-Fixture;$script:master.Remove('cancelledOccurrences');$script:master.Remove('exceptionOccurrences')
    $null=Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'missing-collections')
    Assert ($script:writes.Count -eq 0) 'Missing optional cancellation and exception lists do not create a false cancellation'
    Reset-Fixture
    $dir=Join-Path $TestDirectory 'equivalent-room';New-Item -ItemType Directory -Path $dir -Force|Out-Null
    $before=Copy-TZData $script:instances[0];$adjusted=Copy-TZData $before
    $script:master.location=@{displayName='';uniqueId='master-only-id';locationType='default'}
    $adjusted.type='exception';$adjusted.start.dateTime='2099-01-10T15:00:00';$adjusted.end.dateTime='2099-01-10T16:00:00'
    $adjusted.location=@{displayName='Microsoft Teams Meeting; Test Room';locationEmailAddress='room@example.invalid';uniqueId='exception-only-id';locationType='conferenceRoom'}
    $adjusted.body=[ordered]@{content='Synthetic description';contentType='TEXT'}
    $script:instances[0]=$adjusted;$script:master.exceptionOccurrences=@($adjusted)
    Save-ABJournal @{Identity=$script:identity;Items=@(@{Id=$before.id;Before=$before;Status='Verified';Patch=@{start=@{dateTime='2099-01-10T09:00:00';timeZone='Alberta Standard Time'};end=@{dateTime='2099-01-10T10:00:00';timeZone='Alberta Standard Time'}}})} (Join-Path $dir 'prior.dpapi')
    $plan=Get-TZMigrationPlan @{} $scan $row $dir
    Assert ($plan.NewPayload.location.displayName -eq 'Test Room') 'Equivalent room labels and metadata do not block splitting and the replacement retains the room label'
    Assert ($script:writes.Count -eq 0) 'Semantic content/location comparison sends no writes during preview'
    $adjusted.location.displayName='Different offsite venue'
    Must-Fail {Get-TZMigrationPlan @{} $scan $row $dir} 'A real custom location still blocks the split'
    $adjusted.location.displayName='Microsoft Teams Meeting; Test Room';$adjusted.body.content='Important custom agenda'
    $reason='';try{Get-TZMigrationPlan @{} $scan $row $dir|Out-Null}catch{$reason=$_.Exception.Message}
    Assert ($reason -like '*2099-01-10*' -and $reason -like '*body*') 'A real content difference reports the occurrence date and differing field'
    $adjusted.body.content='Synthetic description';$adjusted.location.address=@{street='Different street'}
    Must-Fail {Get-TZMigrationPlan @{} $scan $row $dir} 'A changed physical address is not hidden by an equivalent room display label'
    Reset-Fixture;$script:master.cancelledOccurrences=@('OID.master.2020-02-01','OID.alternate-id.2099-01-03')
    $dir=Join-Path $TestDirectory 'past-cancellations';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    Assert ($plan.RetainedCancellationCount -eq 2 -and $script:writes.Count -eq 0) 'Cancellations before the replacement start do not block a read-only preview'
    $null=Invoke-TZMigration @{} $plan $dir -Confirmed
    Assert ($script:master.cancelledOccurrences.Count -eq 2 -and $script:master.recurrence.range.endDate -eq '2099-01-09') 'Historical cancellations remain on the original series when it is shortened'
    Reset-Fixture;$script:master.cancelledOccurrences=@('OID.master.2099-01-17')
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'future-cancel')} 'A future cancelled date inside the replacement range still blocks splitting'
    Reset-Fixture;$script:master.cancelledOccurrences=@('OID.master.2026-02-30')
    Must-Fail {Get-TZMigrationPlan @{} $scan $row (Join-Path $TestDirectory 'invalid-cancel-date')} 'An invalid cancellation date is not treated as a past cancellation'
    Reset-Fixture;$script:master.body.content=('Synthetic long meeting description ' * 6000)
    $dir=Join-Path $TestDirectory 'large-plan';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    Assert (($plan|ConvertTo-Json -Depth 30).Length -gt 113204) 'Migration fixture exceeds the reported encrypted-journal size'
    $null=Invoke-TZMigration @{} $plan $dir -Confirmed
    $file=Get-ChildItem -LiteralPath $dir -Filter *.dpapi|Select-Object -First 1
    Assert ((Read-ABJournal $file.FullName).Status -eq 'Complete') 'A full large-plan migration completes with an encrypted recovery journal'
    Reset-Fixture;$dir=Join-Path $TestDirectory 'journal-fails';$plan=Get-TZMigrationPlan @{} $scan $row $dir
    $saveFunction=${function:Save-ABJournal}
    try{
        function script:Save-ABJournal($Journal,$Path){throw 'Synthetic disk write failure'}
        $reason='';try{Invoke-TZMigration @{} $plan $dir -Confirmed|Out-Null}catch{$reason=$_.Exception.Message}
        Assert ($script:writes.Count -eq 0 -and $reason -like '*No calendar change*' -and $reason -notlike '*replacement may already exist*') 'A failed first journal save sends no calendar writes and reports that accurately'
    }finally{Set-Item -Path function:script:Save-ABJournal -Value $saveFunction}
    "All $script:checks migration checks passed. No network requests were made."
} $TestDirectory
Remove-Module $module
