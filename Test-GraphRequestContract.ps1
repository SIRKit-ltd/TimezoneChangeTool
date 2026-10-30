#requires -Version 7.2
$ErrorActionPreference='Stop'
$module=Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
& $module {
    $script:checks=0
    function Assert($condition,$message){if(-not $condition){throw "FAILED: $message"};$script:checks++;"PASS: $message"}
    function Must-Fail([scriptblock]$action,$message){$failed=$false;try{&$action|Out-Null}catch{$failed=$true};Assert $failed $message}
    function Fixture {
        @{subject='Wire contract fixture';start=@{dateTime='2026-11-03T08:30:00';timeZone='Alberta Standard Time'};end=@{dateTime='2026-11-03T09:15:00';timeZone='Alberta Standard Time'};
            recurrence=@{pattern=@{type='weekly';interval=1;daysOfWeek=@('tuesday');firstDayOfWeek='sunday'};
                range=@{type='endDate';startDate='2026-11-03';endDate='2027-04-27';recurrenceTimeZone='Alberta Standard Time'}};
            transactionId=[guid]::NewGuid().ToString();attendees=@();categories=@();locations=@()}
    }
    foreach($count in 0,1,2){
        $array=@(for($i=0;$i -lt $count;$i++){"value-$i"})
        $copy=Copy-TZData $array
        Assert ($copy -is [array] -and $copy.Count -eq $count) "Root arrays retain cardinality $count through the copy helper"
        $payload=Fixture
        $payload.attendees=@(for($i=0;$i -lt $count;$i++){@{emailAddress=@{address="person-$i@example.invalid";name="Person $i"};type='required'}})
        $payload.categories=@(for($i=0;$i -lt $count;$i++){"Category $i"})
        $payload.locations=@(for($i=0;$i -lt $count;$i++){@{displayName="Location $i"}})
        $copy=Copy-TZData $payload
        $json=ConvertTo-TZGraphRequestJson $copy Create
        $wire=[Text.Json.JsonDocument]::Parse($json)
        try{
            foreach($field in 'attendees','categories','locations'){
                $node=$wire.RootElement.GetProperty($field)
                Assert ($node.ValueKind -eq [Text.Json.JsonValueKind]::Array -and $node.GetArrayLength() -eq $count) "Serialized $field stays an array with $count entries"
            }
            $days=$wire.RootElement.GetProperty('recurrence').GetProperty('pattern').GetProperty('daysOfWeek')
            Assert ($days.ValueKind -eq [Text.Json.JsonValueKind]::Array -and $days.GetArrayLength() -eq 1 -and $days[0].GetString() -eq 'tuesday') 'CS TL Checkin one-Tuesday recurrence has the exact Graph array form'
        }finally{$wire.Dispose()}
    }
    $payload=Fixture;$payload.recurrence.pattern.daysOfWeek='tuesday'
    Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} 'The exact scalar-weekday shape rejected by Exchange is blocked before transport'
    foreach($field in 'attendees','categories','locations'){
        $payload=Fixture;$payload[$field]='scalar'
        Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} "Scalar $field is blocked before transport"
    }
    $patterns=@(
        @{type='daily';interval=1},
        @{type='weekly';interval=2;daysOfWeek=@('tuesday','thursday');firstDayOfWeek='sunday'},
        @{type='absoluteMonthly';interval=1;dayOfMonth=3},
        @{type='relativeMonthly';interval=1;daysOfWeek=@('tuesday');index='first'},
        @{type='absoluteYearly';interval=1;dayOfMonth=3;month=11},
        @{type='relativeYearly';interval=1;daysOfWeek=@('tuesday');index='first';month=11})
    foreach($pattern in $patterns){
        $payload=Fixture;$payload.recurrence.pattern=Copy-TZPattern $pattern
        $json=ConvertTo-TZGraphRequestJson $payload Create
        Assert ($json -like '*"recurrence"*') "$($pattern.type) is serialized with its documented fields"
    }
    foreach($type in 'weekly','relativeMonthly','relativeYearly'){
        $source=@{type=$type;interval=1;daysOfWeek=@('tuesday');firstDayOfWeek='sunday';index='first';month=11}
        $pattern=Copy-TZPattern $source
        Assert ($pattern.daysOfWeek -is [array] -and $pattern.daysOfWeek.Count -eq 1) "$type pattern builder retains a one-day collection"
    }
    $payload=Fixture;$payload.id='read-only-id'
    Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} 'Read-only event IDs cannot leak into creation requests'
    $payload=Fixture;$payload.start.dateTime='2026-11-03T08:30:00Z'
    Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} 'Local wall time cannot be accidentally serialized as a UTC instant'
    $payload=Fixture;$payload.recurrence.pattern.month=0
    Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} 'Irrelevant default recurrence fields are rejected'
    $payload=Fixture;$payload.recurrence.range.endDate='2026-01-01'
    Must-Fail {ConvertTo-TZGraphRequestJson $payload Create} 'An invalid end range is rejected before transport'
    "All $script:checks wire-contract checks passed. No network requests were made."
}
Remove-Module $module
