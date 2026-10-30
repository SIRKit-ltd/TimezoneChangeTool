#requires -Version 7.2
[CmdletBinding()]
param(
    [Parameter(Mandatory)][switch]$Confirmed,
    [Parameter(Mandatory)][string]$TestDirectory,
    [string]$ExpectedAccount='',
    [ValidateSet('Alberta','British Columbia','British Columbia (Mountain)','Manitoba')][string]$Province='Alberta'
)
$ErrorActionPreference='Stop'
if(-not $Confirmed){throw 'Explicit authorization for temporary no-attendee Exchange tests is required.'}
$module=Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
try{
    & $module {
        param($TestDirectory,$ExpectedAccount,$Province)
        $config=Read-ABConfig (Join-Path $PSScriptRoot 'config.json')
        Write-Output 'Connecting to Microsoft Graph for the authorized no-attendee test…'
        $identity=Connect-ABGraph $config Fix
        Write-Output 'Microsoft Graph connection completed.'
        if($ExpectedAccount -and $identity.Account -ine $ExpectedAccount){throw 'The signed-in account is not the account authorized for this test.'}
        $target=Get-ABTarget $Province;$definition=Get-TZProfile $Province
        $calendar=Invoke-ABGet '/me/calendar?$select=id,name,owner,canEdit'
        if(-not(Test-TZCalendarOwned $calendar $identity) -or (Get-Field $calendar 'canEdit') -ne $true){throw 'Live tests require an owned, editable default calendar.'}
        New-Item -ItemType Directory -Path $TestDirectory -Force|Out-Null
        $probePath=Join-Path $TestDirectory 'live-probe.dpapi'
        $marker='Timezone Change Tool TEST '+[guid]::NewGuid().ToString()
        $ownedIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $transactions=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $probe=@{Kind='RecurringLiveProbe';Identity=$identity;Items=@();Marker=$marker;Created=[datetimeoffset]::UtcNow;Results=@();Ids=@();Transactions=@();Status='Prepared'}
        Save-ABJournal $probe $probePath
        $testError=$null;$cleanupErrors=[Collections.Generic.List[string]]::new()
        try{
            foreach($scenario in 'HistoricalSplit','EntirelyFuture'){
                Write-Output "Testing Exchange scenario: $scenario"
                $first=[datetime]::SpecifyKind([datetime]::UtcNow.Date.AddDays(7),[DateTimeKind]::Unspecified)
                while($first.DayOfWeek -ne [DayOfWeek]::Tuesday){$first=$first.AddDays(1)}
                if($first -lt $definition.Start.Date){$first=[datetime]::SpecifyKind($definition.Start.Date,[DateTimeKind]::Unspecified);while($first.DayOfWeek -ne [DayOfWeek]::Tuesday){$first=$first.AddDays(1)}}
                $seed=if($scenario -eq 'HistoricalSplit'){$first.AddDays(-35)}else{$first}
                $transaction=[guid]::NewGuid().ToString();[void]$transactions.Add($transaction)
                $payload=@{subject="$marker $scenario";start=@{dateTime=$seed.AddHours(8).AddMinutes(30).ToString('s');timeZone=$definition.OldId};
                    end=@{dateTime=$seed.AddHours(9).AddMinutes(15).ToString('s');timeZone=$definition.OldId};
                    recurrence=@{pattern=@{type='weekly';interval=1;daysOfWeek=@('tuesday');firstDayOfWeek='sunday'};
                        range=@{type='endDate';startDate=$seed.ToString('yyyy-MM-dd');endDate=$first.AddDays(28).ToString('yyyy-MM-dd');recurrenceTimeZone=$definition.OldId}};
                    transactionId=$transaction;attendees=@();isOnlineMeeting=$true;onlineMeetingProvider='teamsForBusiness';
                    body=@{contentType='text';content='Authorized temporary test appointment. No attendees, invitations or room bookings.'};
                    showAs='free';sensitivity='private';isReminderOn=$false;categories=@('Timezone Change Tool Test');
                    location=@{displayName='Temporary test - no room booking'};locations=@(@{displayName='Temporary test - no room booking'})}
                $probe.Transactions=@($transactions);Save-ABJournal $probe $probePath
                $uri=$script:GraphRoot+'/me/calendars/'+[uri]::EscapeDataString($calendar.id)+'/events';Assert-ABUri $uri
                $json=ConvertTo-TZGraphRequestJson $payload Create
                $seedEvent=Invoke-MgGraphRequest -Method POST -Uri $uri -Headers $script:Headers -Body $json -ContentType 'application/json' -OutputType PSObject
                [void]$ownedIds.Add([string]$seedEvent.id);$probe.Ids=@($ownedIds);Save-ABJournal $probe $probePath
                $scan=[pscustomobject]@{Identity=$identity;Target=$target.Alias;Profile=$Province}
                $row=[pscustomobject]@{Kind='seriesMaster';Id=$seedEvent.id;MasterId=$seedEvent.id}
                $scenarioDirectory=Join-Path $TestDirectory $scenario
                $plan=Get-TZMigrationPlan $config $scan $row $scenarioDirectory
                if($plan.NewPayload.attendees.Count -ne 0){throw 'The temporary test unexpectedly acquired attendees. No migration will be performed.'}
                [void]$transactions.Add($plan.NewPayload.transactionId);$probe.Transactions=@($transactions);Save-ABJournal $probe $probePath
                $result=Invoke-TZMigration $config $plan $scenarioDirectory -Confirmed
                [void]$ownedIds.Add($result.NewId);$probe.Ids=@($ownedIds)
                $new=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($result.NewId)+'?$select='+$script:MigrationFields)
                if(@(Get-TZArray $new 'attendees').Count -ne 0 -or -not(Test-TZReplacement $new $plan.NewPayload $target.Zone)){throw 'Live replacement verification failed.'}
                $probe.Results+=@{Scenario=$scenario;Passed=$true;Operation=$plan.Operation;Target=$target.Alias;FirstDate=$plan.FirstDate;SampleDates=$plan.ExpectedStarts.Count;VerifiedAt=[datetimeoffset]::UtcNow}
                Save-ABJournal $probe $probePath
                Write-Output "PASS: $scenario; new zone labels, recurrence dates and original-series end verified by Exchange."
            }
        }catch{
            $testError=$_;$probe.Failure=Get-TZErrorDetails $_
        }finally{
            # Find only this run's exact transaction IDs and marker, including lost POST responses.
            try{
                foreach($event in Get-ABCollection ('/me/calendars/'+[uri]::EscapeDataString($calendar.id)+'/events?$top=250&$select=id,subject,transactionId')){
                    if($transactions.Contains([string](Get-Field $event 'transactionId' '')) -and ([string]$event.subject).StartsWith($marker,[StringComparison]::Ordinal)){[void]$ownedIds.Add($event.id)}
                }
            }catch{$cleanupErrors.Add('Could not discover all temporary test events: '+$_.Exception.Message)}
            foreach($id in $ownedIds){
                try{
                    $event=$null
                    try{$event=Invoke-ABGet ('/me/events/'+[uri]::EscapeDataString($id)+'?$select='+$script:MigrationFields)}catch{
                        $details=Get-TZErrorDetails $_
                        if([string](Get-Field $details 'HttpStatus' '') -in '404','NotFound'){continue};throw
                    }
                    if(-not ([string]$event.subject).StartsWith($marker,[StringComparison]::Ordinal) -or -not $transactions.Contains([string](Get-Field $event 'transactionId' '')) -or
                        $event.isOrganizer -ne $true -or @(Get-TZArray $event 'attendees').Count){throw 'Cleanup identity/attendee check failed. No deletion sent.'}
                    $uri=$script:GraphRoot+'/me/events/'+[uri]::EscapeDataString($id);Assert-ABUri $uri
                    Invoke-MgGraphRequest -Method DELETE -Uri $uri -Headers $script:Headers|Out-Null
                }catch{$cleanupErrors.Add($_.Exception.Message)}
            }
            $probe.Ids=@($ownedIds);$probe.CleanupErrors=@($cleanupErrors)
            $probe.Status=if($cleanupErrors.Count){'Cleanup needs review'}elseif($testError){'Failed - temporary events removed'}else{'Passed - temporary events removed'}
            $probe.Completed=[datetimeoffset]::UtcNow;Save-ABJournal $probe $probePath
        }
        if($cleanupErrors.Count){throw "Temporary test cleanup needs review. Encrypted report: $probePath"}
        if($testError){throw $testError}
        Write-Output "Live Exchange test passed; temporary appointments removed. Encrypted report: $probePath"
    } $TestDirectory $ExpectedAccount $Province
}finally{Disconnect-TZSession;Remove-Module $module}
