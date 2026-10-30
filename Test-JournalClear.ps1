#requires -Version 7.2
param([string]$TestDirectory=(Join-Path $env:TEMP ('JournalClearTests-'+[guid]::NewGuid())))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force
$root=Join-Path $TestDirectory 'Journal'
New-Item -ItemType Directory -Path $root -Force|Out-Null
foreach($name in 'completed','failed','unfinished','unreadable'){
    [IO.File]::WriteAllBytes((Join-Path $root ($name+'.dpapi')),[byte[]](0,1,2,3,255))
}
[IO.File]::WriteAllText((Join-Path $root 'unrelated.txt'),'leave this file')
$blocked=$false;try{Clear-TZJournal $root|Out-Null}catch{$blocked=$true}
if(-not $blocked -or @(Get-ChildItem $root -Filter '*.dpapi').Count -ne 4){throw 'Unconfirmed clearing must leave files intact'}
$result=Clear-TZJournal $root -Confirmed
if($result.Count -ne 4 -or @(Get-ChildItem $root -Filter '*.dpapi').Count -ne 0){throw 'All journal types must leave the active directory'}
if(-not(Test-Path (Join-Path $root 'unrelated.txt'))){throw 'Unrelated files must remain'}
foreach($file in Get-ChildItem -LiteralPath $result.BackupPath){
    if(([IO.File]::ReadAllBytes($file.FullName) -join ',') -ne '0,1,2,3,255'){throw 'Backup bytes changed'}
}
if(@(Get-ChildItem -LiteralPath $result.BackupPath).Count -ne 4){throw 'All files must be backed up'}
if((Clear-TZJournal $root -Confirmed).Count -ne 0){throw 'Clearing an empty folder must be harmless'}
if((Clear-TZJournal (Join-Path $TestDirectory 'missing') -Confirmed).Count -ne 0){throw 'Missing journal folder must be harmless'}
'Journal clear checks passed: confirmation, all statuses, unreadable files, exact backups, unrelated files, empty and missing folders. No production journals were touched.'
