#requires -Version 7.2
param([string]$TestDirectory=(Join-Path $env:TEMP ('ManualResolutionTests-'+[guid]::NewGuid())))
$ErrorActionPreference='Stop'
$m=Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force -PassThru
& $m {
    param($TestDirectory)
    $identity=@{UserId='u';TenantId='t';Account='test@example.invalid'}
    $row=[pscustomobject]@{Id='failed';MasterId='';Kind='singleInstance'}
    $root=Join-Path $TestDirectory 'Journal';New-Item -ItemType Directory -Path $root -Force|Out-Null
    $path=Join-Path $root 'batch.dpapi'
    Save-ABJournal @{Identity=$identity;Items=@(@{Id='completed';Status='Verified'},@{Id='failed';Status='Sending - outcome may be unknown'});Failure=@{ItemNumber=2;Message='failure'}} $path
    $blocked=$false;try{Resolve-TZJournalRow $identity $row $root|Out-Null}catch{$blocked=$true}
    if(-not $blocked -or (Read-ABJournal $path).Items.Count -ne 2){throw 'Confirmation is required before journal edits'}
    $result=Resolve-TZJournalRow $identity $row $root -Confirmed
    $saved=Read-ABJournal $path
    if($saved.Items.Count -ne 1 -or $saved.Items[0].Id -cne 'completed'){throw 'Unrelated batch entries must remain'}
    if((Get-ABJournalIndex $identity $root).ContainsKey('failed')){throw 'Active failure entry must be removed'}
    if(-not(Get-TZResolvedIds $identity $root).Contains('failed')){throw 'Resolved meeting must remain hidden after restarting'}
    if((Get-TZResolvedIds @{UserId='other';TenantId='t'} $root).Contains('failed')){throw 'Resolution must be scoped to the matching account'}
    if(-not(Test-Path (Join-Path $result.BackupPath 'batch.dpapi'))){throw 'Original journal backup is missing'}
    $row=[pscustomobject]@{Id='occurrence';MasterId='master';Kind='seriesMaster'}
    $path=Join-Path $root 'series.dpapi'
    Save-ABJournal @{Kind='SeriesMigration';Identity=$identity;MasterId='master';NewId='replacement';Status='Stopped';Items=@()} $path
    $null=Resolve-TZJournalRow $identity $row $root -Confirmed
    $ids=Get-TZResolvedIds $identity $root
    if((Test-Path $path) -or -not $ids.Contains('master') -or -not $ids.Contains('replacement')){throw 'Both original and replacement must leave active recovery'}
    'Manual-resolution checks passed: confirmation, selective removal, unrelated entries, backup, account isolation and persistent original/replacement suppression. No network calls.'
} $TestDirectory
