# Validate actual serialized JSON at the transport boundary, not the shape of a mock server object.
function ConvertTo-TZGraphRequestJson($Payload,[ValidateSet('Create','EndRecurrence','Attendees')][string]$Purpose){
    $copy=Copy-TZData $Payload
    foreach($field in 'start','end'){
        $part=Get-Field $copy $field
        if($part -and $part.dateTime -is [datetime]){$part.dateTime=$part.dateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffffff',[Globalization.CultureInfo]::InvariantCulture)}
    }
    $json=ConvertTo-Json -InputObject $copy -Depth 40 -Compress -WarningAction Stop
    $document=[Text.Json.JsonDocument]::Parse($json)
    try{
        $root=$document.RootElement
        if($root.ValueKind -ne [Text.Json.JsonValueKind]::Object){throw 'Graph request must be a JSON object.'}
        $allowed=switch($Purpose){
            'Create'{@('subject','start','end','recurrence','transactionId','attendees','body','location','locations','categories','sensitivity','importance','showAs','responseRequested','allowNewTimeProposals','isReminderOn','reminderMinutesBeforeStart','isOnlineMeeting','onlineMeetingProvider')}
            'EndRecurrence'{@('recurrence')}
            'Attendees'{@('attendees')}
        }
        foreach($property in $root.EnumerateObject()){
            if($property.Name -notin $allowed){throw "Unsupported property in $Purpose request: $($property.Name)."}
        }
        foreach($field in 'attendees','locations','categories'){
            $node=[Text.Json.JsonElement]::new()
            if($root.TryGetProperty($field,[ref]$node)){
                if($node.ValueKind -ne [Text.Json.JsonValueKind]::Array){throw "Graph request field '$field' must be a JSON array, even for one item."}
                foreach($item in $node.EnumerateArray()){
                    if($field -eq 'categories'){
                        if($item.ValueKind -ne [Text.Json.JsonValueKind]::String){throw 'Categories must contain strings.'}
                    }else{
                        if($item.ValueKind -ne [Text.Json.JsonValueKind]::Object){throw "$field must contain objects."}
                        if($field -eq 'attendees'){
                            if($item.GetProperty('emailAddress').GetProperty('address').ValueKind -ne [Text.Json.JsonValueKind]::String -or
                                [string]::IsNullOrWhiteSpace($item.GetProperty('emailAddress').GetProperty('address').GetString())){throw 'An attendee needs an email address.'}
                            if($item.GetProperty('type').GetString() -notin 'required','optional','resource'){throw 'Unsupported attendee type.'}
                        }
                    }
                }
            }elseif($field -eq 'attendees' -and $Purpose -in 'Create','Attendees'){throw 'Explicit attendee array is required.'}
        }
        if($Purpose -ne 'Attendees'){
            $recurrence=$root.GetProperty('recurrence');$pattern=$recurrence.GetProperty('pattern');$range=$recurrence.GetProperty('range')
            $type=$pattern.GetProperty('type').GetString()
            if($type -notin 'daily','weekly','absoluteMonthly','relativeMonthly','absoluteYearly','relativeYearly'){throw 'Unsupported recurrence type.'}
            if($pattern.GetProperty('interval').ValueKind -ne [Text.Json.JsonValueKind]::Number -or $pattern.GetProperty('interval').GetInt32() -lt 1){throw 'Recurrence interval must be a positive integer.'}
            if($type -in 'weekly','relativeMonthly','relativeYearly'){
                $days=$pattern.GetProperty('daysOfWeek')
                if($days.ValueKind -ne [Text.Json.JsonValueKind]::Array -or $days.GetArrayLength() -lt 1){throw 'Graph recurrence.pattern.daysOfWeek must be a nonempty JSON array.'}
                foreach($day in $days.EnumerateArray()){
                    if($day.ValueKind -ne [Text.Json.JsonValueKind]::String -or $day.GetString() -notin 'sunday','monday','tuesday','wednesday','thursday','friday','saturday'){throw 'Invalid recurrence weekday.'}
                }
            }
            if($type -eq 'weekly' -and $pattern.GetProperty('firstDayOfWeek').GetString() -notin 'sunday','monday','tuesday','wednesday','thursday','friday','saturday'){throw 'Weekly recurrence needs firstDayOfWeek.'}
            $patternKeys=switch($type){
                'daily'{@('type','interval')}
                'weekly'{@('type','interval','daysOfWeek','firstDayOfWeek')}
                'absoluteMonthly'{@('type','interval','dayOfMonth')}
                'relativeMonthly'{@('type','interval','daysOfWeek','index')}
                'absoluteYearly'{@('type','interval','dayOfMonth','month')}
                'relativeYearly'{@('type','interval','daysOfWeek','index','month')}
            }
            foreach($property in $pattern.EnumerateObject()){if($property.Name -notin $patternKeys){throw "Unsupported field for $type recurrence: $($property.Name)."}}
            if($type -in 'absoluteMonthly','absoluteYearly'){
                $day=$pattern.GetProperty('dayOfMonth')
                if($day.ValueKind -ne [Text.Json.JsonValueKind]::Number -or $day.GetInt32() -lt 1 -or $day.GetInt32() -gt 31){throw 'Invalid recurrence dayOfMonth.'}
            }
            if($type -in 'absoluteYearly','relativeYearly'){
                $month=$pattern.GetProperty('month')
                if($month.ValueKind -ne [Text.Json.JsonValueKind]::Number -or $month.GetInt32() -lt 1 -or $month.GetInt32() -gt 12){throw 'Invalid recurrence month.'}
            }
            if($type -in 'relativeMonthly','relativeYearly' -and $pattern.GetProperty('index').GetString() -notin 'first','second','third','fourth','last'){throw 'Invalid relative recurrence index.'}
            if($range.GetProperty('type').GetString() -notin 'endDate','noEnd','numbered'){throw 'Unsupported recurrence range.'}
            $date=[datetime]::MinValue
            if(-not [datetime]::TryParseExact($range.GetProperty('startDate').GetString(),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)){throw 'Invalid recurrence start date.'}
            if([string]::IsNullOrWhiteSpace($range.GetProperty('recurrenceTimeZone').GetString())){throw 'Explicit recurrence time zone is required.'}
            if($range.GetProperty('type').GetString() -eq 'endDate'){
                $last=[datetime]::MinValue
                if(-not [datetime]::TryParseExact($range.GetProperty('endDate').GetString(),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$last) -or $last -lt $date){throw 'Invalid recurrence end date.'}
            }
            if($range.GetProperty('type').GetString() -eq 'numbered'){
                $count=$range.GetProperty('numberOfOccurrences')
                if($count.ValueKind -ne [Text.Json.JsonValueKind]::Number -or $count.GetInt32() -lt 1){throw 'Invalid recurrence occurrence count.'}
            }
            if($Purpose -eq 'Create'){
                foreach($field in 'start','end'){
                    $part=$root.GetProperty($field)
                    if($part.GetProperty('dateTime').ValueKind -ne [Text.Json.JsonValueKind]::String -or $part.GetProperty('dateTime').GetString() -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?$'){throw "$field.dateTime must be a local ISO date/time string without a UTC offset."}
                    if($part.GetProperty('timeZone').GetString() -cne $range.GetProperty('recurrenceTimeZone').GetString()){throw 'Start/end and recurrence zones must agree.'}
                }
                if($root.GetProperty('start').GetProperty('dateTime').GetString().Substring(0,10) -cne $range.GetProperty('startDate').GetString()){throw 'Event start and recurrence start dates must agree.'}
                if([string]::IsNullOrWhiteSpace($root.GetProperty('subject').GetString())){throw 'Replacement subject is required.'}
                $guid=[guid]::Empty
                if(-not [guid]::TryParse($root.GetProperty('transactionId').GetString(),[ref]$guid) -or $guid -eq [guid]::Empty){throw 'A unique creation transaction ID is required.'}
            }
        }
    }finally{$document.Dispose()}
    $json
}
