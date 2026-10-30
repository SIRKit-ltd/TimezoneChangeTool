# Internal functions loaded by TimezoneEngine.psm1. All writes require a confirmed, fresh plan.
$script:MigrationFields=$script:Fields+',originalStart,body,attendees,location,locations,isOnlineMeeting,onlineMeetingProvider,hasAttachments,categories,sensitivity,importance,showAs,responseRequested,allowNewTimeProposals,isReminderOn,reminderMinutesBeforeStart,transactionId'

function Copy-TZData($Value){
    # Both pipeline stages can unwrap single-item arrays. Keep the input and output as one object.
    $json=ConvertTo-Json -InputObject $Value -Depth 40 -Compress -WarningAction Stop
    $args=@{InputObject=$json;AsHashtable=$true;NoEnumerate=$true}
    if((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')){$args.DateKind='String'}
    $copy=ConvertFrom-Json @args
    return ,$copy
}
function Get-TZArray($Object,[string]$Name){Get-Field $Object $Name @()|Where-Object {$null -ne $_}}
function Get-TZBodySignature($Body){
    if($null -eq $Body){return ''}
    ([string](Get-Field $Body 'contentType' '')).ToLowerInvariant()+'|'+([string](Get-Field $Body 'content' '') -replace "`r`n","`n")
}
function Get-TZLocationSignature($Event){
    $location=Get-Field $Event 'location'
    $rooms=@(Get-TZArray $Event 'attendees'|Where-Object type -eq 'resource')
    $roomIds=@($rooms|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()}|Sort-Object)
    $aliases=@($rooms|ForEach-Object {([string]$_.emailAddress.name).Trim().ToLowerInvariant();([string]$_.emailAddress.address).ToLowerInvariant()})
    $display=([string](Get-Field $location 'displayName' '')).Trim().ToLowerInvariant()
    $parts=@($display.Split(';')|ForEach-Object {$_.Trim()}|Where-Object {$_ -and $_ -notin 'microsoft teams meeting','teams meeting'})
    $resourceLabel=($roomIds.Count -gt 0 -and @($parts|Where-Object {$_ -notin $aliases}).Count -eq 0)
    if($resourceLabel){$label='rooms:'+($roomIds -join ',')}
    else{$label=$display}
    # IDs and display/type metadata can differ in a projected exception. Preserve actual address/coordinates/URLs.
    $detail=[ordered]@{label=$label}
    foreach($key in 'locationEmailAddress','locationUri'){
        $value=[string](Get-Field $location $key '')
        $normalized=$value.Trim().ToLowerInvariant() -replace '^mailto:',''
        if($resourceLabel -and $normalized -in $roomIds){continue}
        if($value){$detail[$key]=$normalized}
    }
    foreach($key in 'address','coordinates'){
        $obj=Get-Field $location $key
        if($null -ne $obj){
            $fields=[ordered]@{}
            foreach($property in 'street','city','state','countryOrRegion','postalCode','latitude','longitude','altitude','accuracy','altitudeAccuracy'){
                $value=Get-Field $obj $property
                if($null -ne $value -and [string]$value -ne ''){$fields[$property]=$value}
            }
            if($fields.Count){$detail[$key]=$fields}
        }
    }
    $detail|ConvertTo-Json -Depth 8 -Compress
}
function Copy-TZPattern($Pattern){
    $copy=@{type=$Pattern.type;interval=$Pattern.interval}
    $keys=switch($Pattern.type){
        'daily'{@()}
        'weekly'{@('daysOfWeek','firstDayOfWeek')}
        'absoluteMonthly'{@('dayOfMonth')}
        'relativeMonthly'{@('daysOfWeek','index')}
        'absoluteYearly'{@('dayOfMonth','month')}
        'relativeYearly'{@('daysOfWeek','index','month')}
        default{throw 'Unsupported recurrence pattern.'}
    }
    foreach($key in $keys){
        if($key -eq 'daysOfWeek'){$copy[$key]=@(Get-TZArray $Pattern $key|ForEach-Object {([string]$_).ToLowerInvariant()});continue}
        $value=Get-Field $Pattern $key
        if($null -ne $value){$copy[$key]=Copy-TZData $value}
    }
    if($Pattern.type -eq 'weekly' -and -not $copy.ContainsKey('firstDayOfWeek')){$copy.firstDayOfWeek='sunday'}
    $copy
}
function Get-TZMigrationMaster([string]$Id){
    Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($Id)+'?$select='+$script:MigrationFields+',cancelledOccurrences,exceptionOccurrences&$expand=exceptionOccurrences($select='+$script:MigrationFields+')')
}
function Get-TZMigrationContent($Event){
    $protected=[ordered]@{}
    foreach($key in @($script:MigrationFields.Split(',')|Sort-Object -Unique)){
        if($key -in 'changeKey','transactionId'){continue}
        $value=Get-Field $Event $key
        if($key -eq 'attendees'){
            $value=@(Get-TZArray $Event 'attendees'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
        }
        $protected[$key]=$value
    }
    $protected['cancelledOccurrences']=@(Get-TZArray $Event 'cancelledOccurrences')
    $protected['exceptions']=@(Get-TZArray $Event 'exceptionOccurrences'|ForEach-Object {([string]$_.id)+'|'+[string](Get-Field $_ '@odata.etag' '')}|Sort-Object)
    $protected|ConvertTo-Json -Depth 30 -Compress
}
function Get-TZPendingMigration($Identity,$MasterId,$JournalDirectory){
    if(-not(Test-Path -LiteralPath $JournalDirectory)){return}
    foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.dpapi' -File|Sort-Object LastWriteTimeUtc -Descending){
        $saved=Read-ABJournal $file.FullName
        if((Get-Field $saved 'Kind') -eq 'SeriesMigration' -and $saved.Identity.UserId -eq $Identity.UserId -and
            $saved.Identity.TenantId -eq $Identity.TenantId -and $saved.MasterId -ceq $MasterId -and $saved.Status -notin 'Complete','Rejected before creation'){
            return [pscustomobject]@{Journal=$saved;Path=$file.FullName}
        }
    }
}
function Get-TZMigrationRecoveryRows($Identity,$JournalDirectory,$Profile){
    if(-not(Test-Path -LiteralPath $JournalDirectory)){return}
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($file in Get-ChildItem -LiteralPath $JournalDirectory -Filter '*.dpapi' -File|Sort-Object LastWriteTimeUtc -Descending){
        $saved=Read-ABJournal $file.FullName
        if((Get-Field $saved 'Kind') -ne 'SeriesMigration' -or $saved.Identity.UserId -ne $Identity.UserId -or
            $saved.Identity.TenantId -ne $Identity.TenantId -or $saved.Plan.Profile -ne $Profile -or -not $seen.Add($saved.MasterId)){continue}
        if($saved.Status -in 'Complete','Rejected before creation'){continue}
        $row=New-ABProposal $saved.Plan.Master ([TimeZoneInfo]::FindSystemTimeZoneById($saved.Plan.Target)) $saved.Plan.Target -Profile $Profile
        $row.MasterId=$saved.MasterId;$row.Kind='seriesMaster';$row.SeriesRecurrence=$saved.Plan.Master.recurrence
        $row.DefaultVisible=$true;$row.Eligible=$false;$row.CanResolve=$true;$row.Calendar='Pending migration'
        $row.CurrentLocal='Future migration starting '+$saved.Plan.FirstDate
        $row.ProposedLocal=$saved.Plan.Target
        $date=[datetime]$saved.Plan.NewPayload.start.dateTime
        $zone=[TimeZoneInfo]::FindSystemTimeZoneById($saved.Plan.Target)
        $row.DisplayStartUtc=[datetimeoffset]::new($date,$zone.GetUtcOffset($date))
        $row.Decision='Incomplete series migration · double-click to reconcile and resume'
        $row
    }
}
function Test-TZReplacement($Event,$Payload,$Zone){
    if([string](Get-Field $Event 'subject') -cne [string](Get-Field $Payload 'subject')){return $false}
    if($Event.type -ne 'seriesMaster' -or $Event.isOrganizer -ne $true -or (Get-Field $Event 'isCancelled' $false)){return $false}
    if((Get-Field $Event 'transactionId' '') -cne $Payload.transactionId){return $false}
    if($Event.recurrence.range.recurrenceTimeZone -cne $Payload.start.timeZone){return $false}
    if($Event.originalStartTimeZone -cne $Payload.start.timeZone -or $Event.originalEndTimeZone -cne $Payload.end.timeZone){return $false}
    if(-not(Test-TZPatchApplied $Event $Payload $Zone)){return $false}
    # Exchange returns sentinel endDate/zero count even when unused. Compare only
    # the fields that define this range, while still checking its type exactly.
    $rangeFields=@('type','startDate')
    switch($Payload.recurrence.range.type){
        'endDate'{$rangeFields+='endDate'}
        'numbered'{$rangeFields+='numberOfOccurrences'}
        'noEnd'{}
        default{return $false}
    }
    foreach($field in $rangeFields){
        if([string](Get-Field $Event.recurrence.range $field '') -cne [string](Get-Field $Payload.recurrence.range $field '')){
            return $false
        }
    }
    foreach($field in 'type','interval','month','dayOfMonth','index','firstDayOfWeek','daysOfWeek'){
        $wanted=Get-Field $Payload.recurrence.pattern $field
        if($null -ne $wanted -and (($wanted|ConvertTo-Json -Compress) -cne ((Get-Field $Event.recurrence.pattern $field)|ConvertTo-Json -Compress))){return $false}
    }
    return $true
}
function Test-TZReplacementDates($Id,$Plan,$Zone,$ProgressQueue){
    $expected=@($Plan.ExpectedStarts|ForEach-Object {([datetime]$_).ToString('s')}|Sort-Object)
    if(-not $expected.Count){return $false}
    $first=[datetime]$expected[0];$last=[datetime]$expected[-1]
    $from=[datetimeoffset]::new($first.AddDays(-1),$Zone.GetUtcOffset($first.AddDays(-1)))
    $until=[datetimeoffset]::new($last.AddDays(1),$Zone.GetUtcOffset($last.AddDays(1)))
    $path='/me/events/'+[uri]::EscapeDataString($Id)+'/instances?startDateTime='+[uri]::EscapeDataString($from.ToString('o'))+
        '&endDateTime='+[uri]::EscapeDataString($until.ToString('o'))+'&$top=250&$select=id,start,end,isCancelled'
    $events=@(Get-ABCollection $path $ProgressQueue 'Verifying replacement recurrence dates')
    $actual=@($events|ForEach-Object {
        $start=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $_.start),$Zone).DateTime
        if($start.Date -ge $first.Date -and $start.Date -le $last.Date){
            if(Get-Field $_ 'isCancelled' $false){throw 'A replacement occurrence was cancelled during migration.'}
            $end=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $_.end),$Zone).DateTime
            if(($end-$start) -ne ([datetime]$Plan.NewPayload.end.dateTime-[datetime]$Plan.NewPayload.start.dateTime)){throw 'A replacement occurrence has an unexpected duration.'}
            $start.ToString('s')
        }
    }|Sort-Object)
    return (($actual -join ',') -ceq ($expected -join ','))
}
function Wait-TZReplacementOnlineDetails($Id,$Payload,$ProgressQueue){
    if(-not(Get-Field $Payload 'isOnlineMeeting' $false)){return}
    $path='/me/events/'+[uri]::EscapeDataString($Id)+'?$select=id,isOnlineMeeting,onlineMeeting,body'
    for($attempt=0;$attempt -lt 6;$attempt++){
        $event=Invoke-ABGet $path
        $join=[string](Get-Field (Get-Field $event 'onlineMeeting' @{}) 'joinUrl' '')
        $body=[Net.WebUtility]::HtmlDecode([string](Get-Field (Get-Field $event 'body' @{}) 'content' ''))
        if((Get-Field $event 'isOnlineMeeting' $false) -and $join -and $body -match '(?i)https://teams\.microsoft\.com/(?:meet/|l/meetup-join/)'){return}
        if($attempt -lt 5){
            Send-TZProgress $ProgressQueue 'Recurring' 'Waiting for Teams joining details before sending invitations…'
            Start-Sleep -Seconds 2
        }
    }
    throw 'Teams joining details are not ready yet. The original series was not changed by this attempt and no new invitations were requested. Reopen the series to resume preparing the existing replacement.'
}
function Convert-TZBodyText($Body){
    $text=[string](Get-Field $Body 'content' '')
    if((Get-Field $Body 'contentType' '') -eq 'html'){
        $text=[regex]::Replace($text,'(?is)<(script|style)\b[^>]*>.*?</\1>','')
        $text=$text -replace '(?i)<br\s*/?>|</p>|</div>',"`n"
        $text=[Net.WebUtility]::HtmlDecode(($text -replace '<[^>]+>',''))
    }
    $text.Trim()
}
function Get-TZDescriptionDraft($Master){
    $text=Convert-TZBodyText (Get-Field $Master 'body')
    if(Get-Field $Master 'isOnlineMeeting' $false){
        $marker=[regex]::Match($text,'(?i)Microsoft Teams(?: meeting)?')
        if(-not $marker.Success){return ''} # User reviews the original alongside the new draft.
        $text=$text.Substring(0,$marker.Index).Trim().TrimEnd('_','-',' ')
    }
    $text
}

function Get-TZMigrationPlan($Config,$Scan,$Row,[string]$JournalDirectory,$ProgressQueue=$null,[switch]$RecreateFutureChanges){
    $definition=Get-TZProfile $Scan.Profile
    if($Row.Kind -ne 'seriesMaster'){throw 'Select a recurring series.'}
    $identity=Connect-ABGraph $Config Scan $ProgressQueue
    if($identity.UserId -ne $Scan.Identity.UserId -or $identity.TenantId -ne $Scan.Identity.TenantId){throw 'Sign in as the preview account.'}
    $target=Get-ABTarget $definition.Key
    if($target.Alias -cne $Scan.Target){throw 'The target zone changed. Scan again.'}
    $masterId=if($Row.MasterId){$Row.MasterId}else{$Row.Id}
    $pending=Get-TZPendingMigration $identity $masterId $JournalDirectory
    if($pending){
        $saved=$pending.Journal
        if(-not(Get-Field $saved 'NewId' '')){
            # An interrupted POST can have succeeded. Discover by the saved transaction ID, never repeat it blindly.
            $matches=@(Get-ABCollection ('/me/calendars/'+[uri]::EscapeDataString($saved.Plan.CalendarId)+'/events?$top=250&$select=id,transactionId') $ProgressQueue 'Checking interrupted replacement'|
                Where-Object { (Get-Field $_ 'transactionId' '') -ceq $saved.Plan.NewPayload.transactionId })
            if($matches.Count -ne 1){throw 'A prior replacement creation has an uncertain outcome. No repeat invitations were sent. Review the saved journal and calendar with IT.'}
            $saved.NewId=$matches[0].id;Save-ABJournal $saved $pending.Path
        }
        $replacement=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($saved.NewId)+'?$select='+$script:MigrationFields)
        if(-not(Test-TZReplacement $replacement $saved.Plan.NewPayload $target.Zone)){throw 'The previous replacement cannot be verified. Review it with IT before further changes.'}
        $plan=[pscustomobject](Copy-TZData $saved.Plan)
        $plan|Add-Member -Force -NotePropertyName ResumeJournalPath -NotePropertyValue $pending.Path
        $plan.ScannedAt=[datetimeoffset]::UtcNow
        return $plan
    }
    Send-TZProgress $ProgressQueue 'Recurring' 'Checking the series, prior repairs and exceptions…'
    $master=Get-TZMigrationMaster $masterId
    if(Get-Field $master 'isAllDay' $false){throw 'All-day events and recurring series are excluded from this tool. No replacement was created.'}
    if($master.type -ne 'seriesMaster' -or $master.isOrganizer -ne $true -or $master.isAllDay -or $master.isCancelled -or
        $master.originalStartTimeZone -cne $definition.OldId -or $master.originalEndTimeZone -cne $definition.OldId){throw 'Only timed series you organize in the selected old zone can be split.'}
    if(-not(Get-Field $master '@odata.etag' '')){throw 'The series version is unavailable.'}
    if(Get-Field $master 'hasAttachments' $false){throw 'This series has attachments. Use Outlook so attachments are preserved.'}
    $cancelled=@(Get-TZArray $master 'cancelledOccurrences')
    if(Get-Field $master 'exceptionOccurrences@odata.nextLink' ''){throw 'The complete exception list is unavailable. Use Outlook for this series.'}
    $calendar=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($masterId)+'/calendar?$select=id,name,owner,canEdit')
    if(-not(Test-TZCalendarOwned $calendar $identity) -or (Get-Field $calendar 'canEdit') -ne $true){throw 'Automatic splitting requires a calendar you own and can edit.'}
    $from=[datetimeoffset]::UtcNow
    if($from -lt $definition.Cutover){$from=$definition.Cutover}
    $source=[TimeZoneInfo]::FindSystemTimeZoneById($definition.OldId)
    $seed=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $master.start),$source).DateTime
    $seedEnd=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $master.end),$source).DateTime
    if(($seedEnd-$seed) -ne ((Convert-ABUtc $master.end)-(Convert-ABUtc $master.start))){throw 'The meeting spans a daylight-saving transition. Use Outlook to review its duration.'}
    $range=$master.recurrence.range;$pattern=Copy-TZPattern $master.recurrence.pattern
    if((Get-Field $range 'recurrenceTimeZone' $definition.OldId) -notin '',$definition.OldId){throw 'The recurrence zone differs from the source zone. Use Outlook to review this series.'}
    $until=switch($range.type){
        'endDate'{[datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc(([datetime]::ParseExact($range.endDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(1)),$source)}
        'numbered'{
            $cycles=[long]$range.numberOfOccurrences*[long]$pattern.interval
            $bound=switch($pattern.type){
                {$_ -in 'absoluteYearly','relativeYearly'}{$seed.AddYears($cycles+1)}
                {$_ -in 'absoluteMonthly','relativeMonthly'}{$seed.AddMonths($cycles+1)}
                'weekly'{$seed.AddDays($cycles*7+7)}
                'daily'{$seed.AddDays($cycles+2)}
                default{throw 'Unsupported recurrence pattern.'}
            }
            [datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc($bound,$source)
        }
        'noEnd'{
            $years=2+[math]::Ceiling([int]$pattern.interval/1.0)
            $from.AddYears([int]$years)
        }
        default{throw 'Unsupported series end condition.'}
    }
    if($until -le $from){throw 'There are no future occurrences to migrate.'}
    $records=@(Read-TZJournalRecords $identity $JournalDirectory)
    $retainedExceptions=0;$discardedExceptions=0
    $retainedBoundaryDate=[TimeZoneInfo]::ConvertTime($from,$source).Date
    $retainedBoundary=[datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc($retainedBoundaryDate,$source)
    foreach($exception in @(Get-TZArray $master 'exceptionOccurrences')){
        $original=Get-Field $exception 'originalStart'
        $originalTime=[datetimeoffset]::MinValue
        if(-not $original -or -not [datetimeoffset]::TryParse([string]$original,[ref]$originalTime)){
            throw 'An edited occurrence has no identifiable original date. Use Outlook to preserve that edit.'
        }
        if([TimeZoneInfo]::ConvertTime($originalTime,$source).Date -lt $retainedBoundaryDate){
            # These slots remain in the original recurrence, including their custom content.
            # Do not silently retain a past slot that was moved into the future range.
            try{$actualStart=Convert-ABUtc $exception.start;$actualEnd=Convert-ABUtc $exception.end}
            catch{throw 'An earlier edited occurrence has an unreadable scheduled time. Use Outlook to preserve that edit.'}
            if($actualEnd -le $actualStart -or $actualStart -ge $retainedBoundary -or $actualEnd -gt $retainedBoundary){
                throw 'An earlier edited occurrence extends into the migration date range. Use Outlook to preserve that moved meeting.'
            }
            $retainedExceptions++
            continue
        }
        if(Get-Field $exception 'isAllDay' $false){throw 'This timed series contains an all-day exception. All-day events are excluded; this series requires manual review to leave that occurrence unchanged.'}
        if($RecreateFutureChanges){$discardedExceptions++;continue}
        $record=Find-TZJournalRecord $records $exception
        if(-not $record -or $record.Entry.Before.type -ne 'occurrence' -or
            -not((Test-TZPatchApplied $exception $record.Entry.Patch $target.Zone) -or (Test-TZSameOccurrence $exception $record.Entry.Before))){
            throw 'This series has custom future exceptions. Use Outlook so those edits are preserved.'
        }
        $differences=[Collections.Generic.List[string]]::new()
        foreach($key in 'subject','isAllDay','hasAttachments','isOnlineMeeting','onlineMeetingProvider'){
            $value=Get-Field $exception $key
            # Omitted/null fields in an expanded occurrence inherit the series value.
            if($null -ne $value -and [string]$value -cne [string](Get-Field $master $key)){$differences.Add($key)}
        }
        $body=Get-Field $exception 'body'
        if($null -ne $body -and (Get-TZBodySignature $body) -cne (Get-TZBodySignature (Get-Field $master 'body'))){$differences.Add('body')}
        if((Get-TZLocationSignature $exception) -cne (Get-TZLocationSignature $master)){$differences.Add('location')}
        if($differences.Count){throw "Adjusted occurrence $(([datetimeoffset]$original).ToString('yyyy-MM-dd')) differs from the series in: $($differences -join ', '). The split is blocked to preserve those changes; use Outlook for this series."}
        $a=@(Get-TZArray $exception 'attendees'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
        $b=@(Get-TZArray $master 'attendees'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
        if(($a -join ',') -cne ($b -join ',')){throw 'An adjusted occurrence has custom attendees. Use Outlook to preserve them.'}
    }
    $path='/me/events/'+[uri]::EscapeDataString($masterId)+'/instances?startDateTime='+[uri]::EscapeDataString($from.AddDays(-2).ToString('o'))+
        '&endDateTime='+[uri]::EscapeDataString($until.ToString('o'))+'&$top=250&$select='+$script:MigrationFields
    $instances=@(Get-ABCollection $path $ProgressQueue 'Previewing future series dates')
    $recreatedCancellations=0
    if($RecreateFutureChanges){
        # Reintroduce the original cancelled slots; Graph's instances collection omits them.
        foreach($cancelledId in $cancelled){
            $match=[regex]::Match([string]$cancelledId,'^OID\..+\.(?<date>\d{4}-\d{2}-\d{2})$')
            $date=[datetime]::MinValue
            if(-not $match.Success -or -not [datetime]::TryParseExact($match.Groups['date'].Value,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)){
                throw 'A cancelled occurrence has no identifiable original date. It cannot be safely included in the replacement preview.'
            }
            $wall=$date.Add($seed.TimeOfDay)
            $utc=[datetimeoffset][TimeZoneInfo]::ConvertTimeToUtc($wall,$source)
            $desired=[datetimeoffset]::new($wall,$target.Zone.GetUtcOffset($wall))
            if($utc -lt $from -or $utc -ge $until -or $desired -lt [datetimeoffset]::UtcNow){continue}
            $instances+=@{id=[string]$cancelledId;type='occurrence';originalStart=$utc.ToString('o');RecreatedCancellation=$true;
                start=@{dateTime=$utc.UtcDateTime.ToString('s');timeZone='UTC'};end=@{dateTime=$utc.UtcDateTime.Add($seedEnd-$seed).ToString('s');timeZone='UTC'}}
            $recreatedCancellations++
        }
        # An exception moved outside the queried window still has an original future slot.
        foreach($exception in @(Get-TZArray $master 'exceptionOccurrences')){
            if(-not @($instances|Where-Object {$_.id -ceq $exception.id}).Count){$instances+=,$exception}
        }
    }
    $slots=@($instances|ForEach-Object {
        $event=$_;$original=Get-Field $event 'originalStart'
        if(-not $original -and $event.type -eq 'exception'){throw 'An exception has no original date. Use Outlook.'}
        $slot=if($original){[datetimeoffset]$original}else{Convert-ABUtc $event.start}
        $date=[TimeZoneInfo]::ConvertTime($slot,$source).Date
        $wall=$date.Add($seed.TimeOfDay)
        $desired=[datetimeoffset]::new($wall,$target.Zone.GetUtcOffset($wall))
        if($slot -ge $from -and (-not $RecreateFutureChanges -or $slot -lt $until) -and $desired -ge [datetimeoffset]::UtcNow){[pscustomobject]@{Event=$event;Date=$date;Wall=$wall}}
    }|Sort-Object Date)
    if(-not $slots.Count){throw 'No future dates could be confirmed in this series.'}
    # Reject unknown exceptions even if Exchange omitted them from its expanded list.
    foreach($slot in $slots|Where-Object {$_.Event.type -eq 'exception'}){
        if(-not @(Get-TZArray $master 'exceptionOccurrences'|Where-Object {$_.id -ceq $slot.Event.id}).Count){throw 'The exception inventory is incomplete. Use Outlook for this series.'}
    }
    $first=$slots[0];$last=$slots[-1]
    $retainedCancellations=0;$futureCancellations=0;$unknownCancellations=0
    foreach($cancelledId in $cancelled){
        # Graph occurrence IDs encode the original start DATE in the recurrence zone.
        # The ID segment can use a different ID representation; membership comes from the master's list.
        $match=[regex]::Match([string]$cancelledId,'^OID\..+\.(?<date>\d{4}-\d{2}-\d{2})$')
        $cancelledDate=[datetime]::MinValue
        if(-not $match.Success -or -not [datetime]::TryParseExact($match.Groups['date'].Value,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$cancelledDate)){
            $unknownCancellations++;continue
        }
        if($cancelledDate.Date -lt $first.Date){$retainedCancellations++}else{$futureCancellations++}
    }
    if(($futureCancellations -and -not $RecreateFutureChanges) -or $unknownCancellations){
        throw "The original series has $retainedCancellations cancellations before the replacement starts (ignored), $futureCancellations in the replacement date range, and $unknownCancellations with unidentifiable dates. Automatic splitting is blocked only for the future/unidentifiable cancellations, to avoid recreating cancelled meetings."
    }
    $wallEnd=$first.Wall.Add($seedEnd-$seed)
    if($wallEnd -le $first.Wall){throw 'Invalid meeting duration.'}
    $newRange=@{type='endDate';startDate=$first.Date.ToString('yyyy-MM-dd');recurrenceTimeZone=$target.Alias;endDate=$last.Date.ToString('yyyy-MM-dd')}
    if($range.type -eq 'endDate'){$newRange.endDate=$range.endDate}
    if($range.type -eq 'noEnd'){$newRange=@{type='noEnd';startDate=$first.Date.ToString('yyyy-MM-dd');recurrenceTimeZone=$target.Alias}}
    $includeTimeInTitle=[bool](Get-Field $Scan 'IncludeTimeInTitle' $false)
    $payload=@{subject=$master.subject;start=@{dateTime=$first.Wall.ToString('s');timeZone=$target.Alias};end=@{dateTime=$wallEnd.ToString('s');timeZone=$target.Alias};
        recurrence=@{pattern=(Copy-TZData $pattern);range=$newRange};transactionId=[guid]::NewGuid().ToString();
        attendees=@(Get-TZArray $master 'attendees'|ForEach-Object {@{emailAddress=@{address=$_.emailAddress.address;name=$_.emailAddress.name};type=$_.type}})}
    if($includeTimeInTitle){$payload.subject=Add-TZTimeToTitle $master.subject $seed}
    foreach($key in 'location','locations','categories','sensitivity','importance','showAs','responseRequested','allowNewTimeProposals','isReminderOn','reminderMinutesBeforeStart'){
        if($key -in 'locations','categories'){
            $values=@(Get-TZArray $master $key)
            if($values.Count){$payload[$key]=Copy-TZData $values}
            continue
        }
        $value=Get-Field $master $key
        if($null -ne $value){$payload[$key]=Copy-TZData $value}
    }
    if([string]::IsNullOrWhiteSpace([string](Get-Field (Get-Field $master 'location') 'displayName' ''))){
        $roomNames=@($payload.attendees|Where-Object type -eq 'resource'|ForEach-Object {$_.emailAddress.name})
        if($roomNames.Count){$payload.location=@{displayName=($roomNames -join '; ')}}
    }
    $online=Get-Field $master 'isOnlineMeeting' $false
    if(-not $online -and [string](Get-Field (Get-Field $master 'body') 'content' '') -match '(?i)teams\.microsoft\.com(?:/|%2f)(?:meet|l(?:/|%2f)meetup-join)'){
        throw 'This series contains a Teams join link but is not identified as a supported online meeting. Use Outlook so its meeting link is handled correctly.'
    }
    if($online){
        if((Get-Field $master 'onlineMeetingProvider') -ne 'teamsForBusiness'){throw 'This online meeting provider is not supported for automatic replacement. Use Outlook.'}
        $payload.isOnlineMeeting=$true;$payload.onlineMeetingProvider='teamsForBusiness'
    }
    $description=Get-TZDescriptionDraft $master
    if($online){$payload.body=@{contentType='text';content=$description}}else{$payload.body=Copy-TZData (Get-Field $master 'body' @{contentType='text';content=''})}
    $action=if($seed.Date -ge $first.Date){'ReplaceFutureSeries'}else{'SplitSeries'}
    $oldEnd=$first.Date.AddDays(-1).ToString('yyyy-MM-dd')
    $endPatch=@{recurrence=@{pattern=(Copy-TZData $pattern);range=@{type='endDate';startDate=$range.startDate;endDate=$oldEnd;recurrenceTimeZone=(Get-Field $range 'recurrenceTimeZone' $definition.OldId)}}}
    $preview=@($slots|Select-Object -First 100|ForEach-Object {
        $current=[TimeZoneInfo]::ConvertTime((Convert-ABUtc $_.Event.start),$target.Zone)
        $before=if(Get-Field $_.Event 'RecreatedCancellation' $false){'Cancelled — will be recreated'}else{$current.ToString('yyyy-MM-dd HH:mm')}
        [pscustomobject]@{Subject=$payload.subject;CurrentLocal=$before;ProposedLocal=$_.Wall.ToString('yyyy-MM-dd HH:mm')}
    })
    [pscustomobject]@{MigrationVersion=1;Identity=$identity;Target=$target.Alias;Profile=$definition.Key;MasterId=$masterId;Master=$master;MasterETag=(Get-Field $master '@odata.etag');
        MasterContent=(Get-TZMigrationContent $master);CalendarId=$calendar.id;Subject=$payload.subject;IncludeTimeInTitle=$includeTimeInTitle;Rows=$preview;RangeType=$range.type;EndDate=(Get-Field $newRange 'endDate' '');
        From=$from;Until=$until;FirstDate=$first.Date.ToString('yyyy-MM-dd');LastOccurrence=$last.Date;ScannedAt=[datetimeoffset]::UtcNow;Operation=$action;
        OldEnd=$oldEnd;EndPatch=$endPatch;NewPayload=$payload;ReplacementDescription=$description;IsOnlineMeeting=$online;OriginalDescription=(Convert-TZBodyText (Get-Field $master 'body'));
        RecreateFutureChanges=[bool]$RecreateFutureChanges;DiscardedExceptionCount=$discardedExceptions;RecreatedCancellationCount=$recreatedCancellations;FutureCancellationCount=$futureCancellations;DiscardChangesConfirmed=$false;
        AlreadyAppliedCount=0;UnresolvedCount=0;SkippedCount=0;FutureCount=$slots.Count;RetainedCancellationCount=$retainedCancellations;RetainedExceptionCount=$retainedExceptions;ExpectedStarts=@($slots|ForEach-Object {$_.Wall.ToString('s')});ResumeJournalPath=''}
}

function Invoke-TZMigration($Config,$Plan,[string]$JournalDirectory,[switch]$Confirmed,$ProgressQueue=$null){
    if(-not $Confirmed){throw 'Confirm the series replacement preview before making changes.'}
    if((Get-Field $Plan 'RecreateFutureChanges' $false) -and -not (Get-Field $Plan 'DiscardChangesConfirmed' $false)){throw 'Explicitly confirm restoring cancellations and discarding future edits before migrating.'}
    if((Get-Field $Plan 'MigrationVersion' 0) -ne 1){throw 'Open a new series migration preview.'}
    if(([datetimeoffset]::UtcNow-[datetimeoffset]$Plan.ScannedAt).TotalMinutes -gt 30){throw 'Series preview expired. Open it again; no full rescan is required.'}
    $identity=Connect-ABGraph $Config Fix $ProgressQueue
    if($identity.UserId -ne $Plan.Identity.UserId -or $identity.TenantId -ne $Plan.Identity.TenantId){throw 'The repair account must match the preview.'}
    $target=Get-ABTarget $Plan.Profile
    if($target.Alias -cne $Plan.Target){throw 'The target time zone changed.'}
    $pending=Get-TZPendingMigration $identity $Plan.MasterId $JournalDirectory
    $journal=$null;$journalPath='';$stage='Checking series';$creationMayExist=$false
    try{
        if($pending){
            if(-not $Plan.ResumeJournalPath -or $Plan.ResumeJournalPath -cne $pending.Path){throw 'A prior migration needs a fresh resume preview. No new series was created.'}
            $journal=$pending.Journal;$journalPath=$pending.Path
            $creationMayExist=$true
            if(-not(Get-Field $journal 'NewId' '')){throw 'The replacement outcome is still uncertain. No repeat invitations were sent.'}
            $payload=$journal.Plan.NewPayload
            $replacement=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($journal.NewId)+'?$select='+$script:MigrationFields)
            if(-not(Test-TZReplacement $replacement $payload $target.Zone)){throw 'The replacement series changed or cannot be verified. Ask IT to review it.'}
        }else{
            $row=[pscustomobject]@{Kind='seriesMaster';MasterId=$Plan.MasterId;Id=$Plan.MasterId}
            $freshPlan=Get-TZMigrationPlan $Config $Plan $row $JournalDirectory $ProgressQueue -RecreateFutureChanges:([bool](Get-Field $Plan 'RecreateFutureChanges' $false))
            if($freshPlan.MasterETag -cne $Plan.MasterETag -or $freshPlan.MasterContent -cne $Plan.MasterContent -or $freshPlan.FirstDate -cne $Plan.FirstDate){throw 'The series changed after preview. Open a fresh preview.'}
            $payload=Copy-TZData $freshPlan.NewPayload
            $payload.transactionId=$Plan.NewPayload.transactionId
            if($Plan.IsOnlineMeeting){$payload.body=@{contentType='text';content=[string]$Plan.ReplacementDescription}}
            New-Item -ItemType Directory -Path $JournalDirectory -Force|Out-Null
            $journalPath=Join-Path $JournalDirectory ([guid]::NewGuid().ToString()+'.dpapi')
            $savedPlan=Copy-TZData $Plan;$savedPlan.NewPayload=$payload
            $journal=@{Version=5;InitialAttendeesMode='none';Kind='SeriesMigration';Identity=$identity;MasterId=$Plan.MasterId;Plan=$savedPlan;Items=@();Status='Prepared';Created=[datetimeoffset]::UtcNow}
            Save-ABJournal $journal $journalPath
            $stage='Creating replacement series';$journal.Status='Creating - outcome may be unknown';Save-ABJournal $journal $journalPath
            Send-TZProgress $ProgressQueue 'Recurring' 'Creating the confirmed replacement series…' 0 3
            $uri=$script:GraphRoot+'/me/calendars/'+[uri]::EscapeDataString($Plan.CalendarId)+'/events';Assert-ABUri $uri
            # Prepare the final description and Teams details without recipients.
            # People and rooms are added together only after verification and old-series end.
            $createPayload=Copy-TZData $payload;$createPayload.attendees=@()
            $requestJson=ConvertTo-TZGraphRequestJson $createPayload Create
            $journal.CreateRequestJson=$requestJson
            Save-ABJournal $journal $journalPath
            $creationMayExist=$true
            try{$created=Invoke-MgGraphRequest -Method POST -Uri $uri -Headers $script:Headers -Body $requestJson -ContentType 'application/json' -OutputType PSObject}
            catch{
                $errorInfo=Get-TZErrorDetails $_
                if([string](Get-Field $errorInfo 'HttpStatus' '') -in '400','403','404','BadRequest','Forbidden','NotFound'){$journal.Status='Rejected before creation';$creationMayExist=$false}
                throw
            }
            $journal.NewId=[string](Get-Field $created 'id' '')
            Save-ABJournal $journal $journalPath
            if(-not $journal.NewId){throw 'The creation response had no event ID. The outcome needs reconciliation.'}
            $replacement=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($journal.NewId)+'?$select='+$script:MigrationFields)
            if(-not(Test-TZReplacement $replacement $payload $target.Zone)){throw 'The replacement series did not verify in the new time zone. The old series was not changed.'}
            $journal.Status='Replacement verified';Save-ABJournal $journal $journalPath
        }
        $stage='Verifying replacement recurrence dates'
        if(-not(Test-TZReplacementDates $journal.NewId $journal.Plan $target.Zone $ProgressQueue)){throw 'Replacement recurrence dates do not match the preview. The original series has not been changed by this attempt.'}
        if((Get-Field $journal 'InitialAttendeesMode' '') -eq 'none'){
            $stage='Preparing joining details'
            Wait-TZReplacementOnlineDetails $journal.NewId $payload $ProgressQueue
        }
        $stage='Ending old series'
        Send-TZProgress $ProgressQueue 'Recurring' 'Replacement verified. Ending the original future series…' 1 3
        $old=$null
        try{$old=Get-TZMigrationMaster $Plan.MasterId}catch{
            $info=Get-TZErrorDetails $_
            if([string](Get-Field $info 'HttpStatus' '') -notin '404','NotFound'){throw}
        }
        $alreadyEnded=($Plan.Operation -eq 'ReplaceFutureSeries' -and ($null -eq $old -or (Get-Field $old 'isCancelled' $false)))
        if($old -and $Plan.Operation -eq 'SplitSeries'){
            $alreadyEnded=($old.recurrence.range.type -eq 'endDate' -and $old.recurrence.range.endDate -ceq $Plan.OldEnd)
        }
        if(-not $alreadyEnded){
            if($null -eq $old -or (Get-TZMigrationContent $old) -cne $Plan.MasterContent){throw 'The original series changed before it could be ended. The replacement exists; reopen this series to review and resume. No second replacement will be created.'}
            $calendar=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($Plan.MasterId)+'/calendar?$select=id,name,owner,canEdit')
            if($calendar.id -cne $Plan.CalendarId -or -not(Test-TZCalendarOwned $calendar $identity) -or (Get-Field $calendar 'canEdit') -ne $true){throw 'The original calendar ownership or location changed. No original-series change was sent.'}
            $headers=@{Prefer=$script:Headers.Prefer;'If-Match'=[string](Get-Field $old '@odata.etag')}
            $uri=$script:GraphRoot+'/me/events/'+[uri]::EscapeDataString($Plan.MasterId);Assert-ABUri $uri
            $journal.Status='Ending original - outcome may be unknown';Save-ABJournal $journal $journalPath
            if($Plan.Operation -eq 'ReplaceFutureSeries'){
                Invoke-MgGraphRequest -Method DELETE -Uri $uri -Headers $headers|Out-Null
            }else{
                $requestJson=ConvertTo-TZGraphRequestJson $Plan.EndPatch EndRecurrence
                Invoke-MgGraphRequest -Method PATCH -Uri $uri -Headers $headers -Body $requestJson -ContentType 'application/json' -OutputType PSObject|Out-Null
            }
        }
        $stage='Verifying original series end'
        if($Plan.Operation -eq 'SplitSeries'){
            $ended=Get-TZMigrationMaster $Plan.MasterId
            if($ended.recurrence.range.type -ne 'endDate' -or $ended.recurrence.range.endDate -cne $Plan.OldEnd -or
                (Convert-ABUtc $ended.start) -ne (Convert-ABUtc $Plan.Master.start) -or (Convert-ABUtc $ended.end) -ne (Convert-ABUtc $Plan.Master.end)){
                throw 'The original series end could not be verified. Review the saved journal before continuing.'
            }
        }else{
            $stillExists=$false
            try{$remaining=Get-TZMigrationMaster $Plan.MasterId;$stillExists=-not (Get-Field $remaining 'isCancelled' $false)}catch{
                $info=Get-TZErrorDetails $_;if([string](Get-Field $info 'HttpStatus' '') -notin '404','NotFound'){throw}
            }
            if($stillExists){throw 'Cancellation of the original future series could not be verified.'}
        }
        $stage='Sending replacement invitations'
        Send-TZProgress $ProgressQueue 'Recurring' 'Sending the prepared replacement to people and rooms together…' 2 3
        $replacement=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($journal.NewId)+'?$select='+$script:MigrationFields)
        $expected=@($payload.attendees|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
        $actual=@(Get-TZArray $replacement 'attendees'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
        if(($actual -join ',') -cne ($expected -join ',')){
            $initial=if((Get-Field $journal 'InitialAttendeesMode' '') -eq 'none'){@()}else{@($payload.attendees|Where-Object type -ne 'resource'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)}
            if(($actual -join ',') -cne ($initial -join ',')){throw 'Replacement attendees changed. No attendee update was sent; review the replacement in Outlook.'}
            $etag=[string](Get-Field $replacement '@odata.etag' '')
            if(-not $etag){throw 'The replacement version is unavailable. No attendee update was sent.'}
            $journal.Status='Inviting attendees - outcome may be unknown';Save-ABJournal $journal $journalPath
            $uri=$script:GraphRoot+'/me/events/'+[uri]::EscapeDataString($journal.NewId);Assert-ABUri $uri
            $headers=@{Prefer=$script:Headers.Prefer;'If-Match'=$etag}
            $requestJson=ConvertTo-TZGraphRequestJson @{attendees=$payload.attendees} Attendees
            Invoke-MgGraphRequest -Method PATCH -Uri $uri -Headers $headers -Body $requestJson -ContentType 'application/json' -OutputType PSObject|Out-Null
            $replacement=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($journal.NewId)+'?$select='+$script:MigrationFields)
            $actual=@(Get-TZArray $replacement 'attendees'|ForEach-Object {([string]$_.emailAddress.address).ToLowerInvariant()+'|'+$_.type}|Sort-Object)
            if(($actual -join ',') -cne ($expected -join ',')){throw 'The replacement attendee list could not be verified.'}
        }
        if(-not(Test-TZReplacement $replacement $payload $target.Zone)){throw 'Final replacement zone verification failed.'}
        $journal.Status='Complete';$journal.Completed=[datetimeoffset]::UtcNow;Save-ABJournal $journal $journalPath
        Send-TZProgress $ProgressQueue 'Complete' 'Future series migrated and verified in the new time zone.' 3 3
        [pscustomobject]@{Count=1;AlreadyAppliedCount=0;MigratedSeries=$true;NewId=$journal.NewId;JournalPath=$journalPath}
    }catch{
        if($journal){
            $failure=Get-TZErrorDetails $_;$failure.Stage=$stage;$failure.ItemNumber=1;$failure.VerifiedCount=0
            $journal.Failure=$failure
            $errorSaved=$false
            try{Save-ABJournal $journal $journalPath;$errorSaved=$true}catch{}
            $outcome=if($creationMayExist){'A replacement may already exist. Reopen this series for reconciliation; do not create a replacement manually without reviewing it.'}elseif($stage -eq 'Creating replacement series'){'Exchange rejected replacement creation. No original-series change was sent. Resolve the error and open a fresh preview.'}else{'No calendar change was requested by this attempt. Resolve the error and open a fresh preview.'}
            $journalNote=if($errorSaved){"Keep journal $journalPath."}else{'The encrypted error journal could not be saved. Copy this error before closing it.'}
            $serverReason=[string](Get-Field $failure 'GraphReason' (Get-Field $failure 'GraphMessage' ''))
            $code=[string](Get-Field $failure 'GraphCode' '')
            $serverNote=if($code -or $serverReason){"Microsoft response: $code $serverReason"}else{''}
            throw "Series migration stopped while $($stage.ToLowerInvariant()). $($failure.Message) $serverNote $outcome $journalNote"
        }
        throw
    }
}
