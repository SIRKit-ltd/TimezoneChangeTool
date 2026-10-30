#requires -Version 7.2
[CmdletBinding()]
param(
    [switch]$Demo,
    [switch]$Connect,
    [switch]$ValidateUi,
    [ValidateSet('Alberta','British Columbia','British Columbia (Mountain)','Manitoba')][string]$Province='Alberta',
    [string]$PreviewImagePath,
    [string]$SeriesPreviewImagePath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json')
)
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This preview window requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with: pwsh -STA -File .\Start-TimezoneChangeTool.ps1 -Demo' }
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Import-Module (Join-Path $PSScriptRoot 'TimezoneEngine.psm1') -Force
$script:scan = $null
$script:operation = $null
$script:journalDirectory = Join-Path $env:LOCALAPPDATA 'AlbertaCalendarPilot\Journal'
$script:config = $null
$script:connectedIdentity=$null
$script:launchSignInStarted=$false
$script:autoConnect = [bool]$Connect
if (-not $Demo -and -not $ValidateUi) { $script:config = Read-ABConfig $ConfigPath }

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Timezone Change Tool | Sirkit" Width="1320" Height="860" MinWidth="1120" MinHeight="740"
        WindowStartupLocation="CenterScreen" Background="White" Foreground="#002849" FontFamily="Segoe UI" FontSize="14">
 <Grid Margin="24" Background="White">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Border BorderBrush="#EDF0F2" BorderThickness="0,0,0,1" Padding="0,0,0,12">
   <DockPanel LastChildFill="True">
    <Image x:Name="SirkitLogo" DockPanel.Dock="Left" Width="180" Height="66" Stretch="Uniform" Margin="-12,0,28,0"/>
    <TextBlock Text="Timezone Change Tool" FontSize="28" FontWeight="SemiBold" VerticalAlignment="Center"/>
   </DockPanel>
  </Border>
  <DockPanel Grid.Row="1" Margin="0,18,0,16">
   <Button x:Name="ScanButton" DockPanel.Dock="Right" Content="Connect and scan" Padding="22,10" Margin="20,0,0,0" Background="#0252C6" BorderBrush="#0252C6" Foreground="White" VerticalAlignment="Bottom"/>
   <Button x:Name="OutlookTimezoneButton" DockPanel.Dock="Right" Content="Change Outlook time zone…" Padding="14,10" Margin="12,0,0,0" Background="White" BorderBrush="#0252C6" Foreground="#0252C6" VerticalAlignment="Bottom" IsEnabled="False"/>
   <StackPanel Orientation="Horizontal">
    <StackPanel><TextBlock Text="Province" FontSize="12" Margin="0,0,0,6"/><ComboBox x:Name="ProvinceBox" Width="200" Padding="8,6">
     <ComboBoxItem Tag="British Columbia" Content="British Columbia"/><ComboBoxItem Tag="Alberta" Content="Alberta"/><ComboBoxItem Tag="Manitoba" Content="Manitoba"/>
    </ComboBox></StackPanel>
    <StackPanel x:Name="RegionPanel" Margin="20,0,0,0" Visibility="Collapsed"><TextBlock Text="BC area" FontSize="12" Margin="0,0,0,6"/><ComboBox x:Name="RegionBox" Width="290" Padding="8,6">
     <ComboBoxItem Tag="British Columbia" Content="Pacific time (most of BC)"/><ComboBoxItem Tag="British Columbia (Mountain)" Content="Mountain time (aligned with Alberta)"/>
    </ComboBox></StackPanel>
   </StackPanel>
  </DockPanel>
  <Border Grid.Row="2" Background="#EDF0F2" Padding="14" Margin="0,0,0,14" CornerRadius="4">
   <StackPanel>
    <TextBlock x:Name="TimezonePair" FontWeight="SemiBold" TextWrapping="Wrap"/>
    <TextBlock Text="November 1, 2026 onward · Future entries you organize · Recurring series show an upcoming occurrence" FontSize="12" Margin="0,5,0,0"/>
    <TextBlock x:Name="AccountLabel" Text="Not connected" TextWrapping="Wrap" FontSize="12" Margin="0,6,0,0"/>
    <ProgressBar x:Name="OperationProgress" Height="7" Margin="0,12,0,0" Minimum="0" Maximum="100" Foreground="#0252C6" Background="#D5DFE7" Visibility="Collapsed"/>
    <TextBlock x:Name="ProgressText" Text="" FontSize="12" Margin="0,6,0,0" TextWrapping="Wrap"/>
   </StackPanel>
  </Border>
  <DataGrid x:Name="PreviewGrid" Grid.Row="3" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
            SelectionMode="Extended" HeadersVisibility="Column" GridLinesVisibility="Horizontal" RowHeight="58" Background="White" BorderBrush="#EDF0F2" HorizontalGridLinesBrush="#EDF0F2">
   <DataGrid.Resources>
    <Style x:Key="CellText" TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="Margin" Value="3"/><Setter Property="ToolTip" Value="{Binding Text,RelativeSource={RelativeSource Self}}"/></Style>
   </DataGrid.Resources>
   <DataGrid.RowStyle><Style TargetType="DataGridRow"><Setter Property="ToolTip" Value="{Binding Decision}"/></Style></DataGrid.RowStyle>
   <DataGrid.Columns>
    <DataGridTemplateColumn Header="Fix" Width="45">
     <DataGridTemplateColumn.CellTemplate><DataTemplate>
      <CheckBox IsChecked="{Binding Selected,Mode=TwoWay,UpdateSourceTrigger=PropertyChanged}" IsEnabled="{Binding Eligible}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
     </DataTemplate></DataGridTemplateColumn.CellTemplate>
    </DataGridTemplateColumn>
    <DataGridTextColumn Header="Event" Binding="{Binding Subject}" Width="200" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Calendar" Binding="{Binding Calendar}" Width="90" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Before (new zone)" Binding="{Binding CurrentLocal}" Width="185" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="After (new zone)" Binding="{Binding ProposedLocal}" Width="185" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Shift (min)" Binding="{Binding ShiftMinutes}" Width="70" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Review / reason" Binding="{Binding Decision}" Width="*" MinWidth="210" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTextColumn Header="Original zone" Binding="{Binding SourceZone}" Width="130" IsReadOnly="True" ElementStyle="{StaticResource CellText}"/>
    <DataGridTemplateColumn Header="Resolve error" Width="180"><DataGridTemplateColumn.CellTemplate><DataTemplate>
      <Button Tag="ResolveJournalRow" Content="I have manually resolved this" Padding="5,6"><Button.Style><Style TargetType="Button"><Setter Property="Visibility" Value="Collapsed"/><Style.Triggers><DataTrigger Binding="{Binding CanResolve}" Value="True"><Setter Property="Visibility" Value="Visible"/></DataTrigger></Style.Triggers></Style></Button.Style></Button>
    </DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
   </DataGrid.Columns>
  </DataGrid>
  <StackPanel Grid.Row="4" Margin="0,16,0,0">
   <CheckBox x:Name="TitleTimeCheck" Content="Add original time to event titles (for example, Team meeting (11:00AM))" Margin="0,0,0,10"/>
   <CheckBox x:Name="ShowExcludedCheck" Content="Show excluded entries (past, already correct, other zones, or not organized by me)" Margin="0,0,0,10"/>
   <TextBlock Text="Double-click a recurring meeting to preview replacing its future series in the new time zone." FontSize="12" Margin="0,0,0,10"/>
   <CheckBox x:Name="IntentCheck" Content="I reviewed the selected events and want to preserve their original clock times in the selected regional time zone."/>
   <CheckBox x:Name="NoticeCheck" Margin="0,8,0,0" Content="I understand meeting updates may be sent, and restoring old times can send further updates."/>
  </StackPanel>
  <DockPanel Grid.Row="5" Margin="0,18,0,0" LastChildFill="True">
   <Button x:Name="FixButton" DockPanel.Dock="Right" Content="Repair selected" Padding="22,10" IsEnabled="False" Margin="20,0,0,0" Background="#0252C6" BorderBrush="#0252C6" Foreground="White"/>
   <Button x:Name="SeriesButton" DockPanel.Dock="Right" Content="Adjust recurring series…" Padding="18,10" IsEnabled="False" Margin="16,0,0,0" Background="White" BorderBrush="#0252C6" Foreground="#0252C6"/>
   <Button x:Name="LastErrorButton" DockPanel.Dock="Left" Content="Last error…" Padding="10,8" Margin="0,0,12,0" Background="White" BorderBrush="#0252C6" Foreground="#0252C6"/>
   <Button x:Name="ClearJournalButton" DockPanel.Dock="Left" Content="Clear journals…" Padding="10,8" Margin="0,0,12,0" Background="White" BorderBrush="#0252C6" Foreground="#0252C6"/>
   <TextBlock x:Name="StatusLabel" Text="Ready" TextWrapping="Wrap" VerticalAlignment="Center" FontSize="12"/>
  </DockPanel>
 </Grid>
</Window>
'@
$reader = [Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$reader.Close()
$window.FindName('ClearJournalButton').Add_Click({
    if($Demo -or $ValidateUi -or $null -ne $script:operation){return}
    $message="Clear ALL active journal files, including failed or unfinished changes?`n`nFolder: $script:journalDirectory`n`nThis does not undo any calendar changes. The tool will forget previous attempts, so retrying can create duplicate meetings or send invitations again. Review existing replacements in Outlook before retrying.`n`nEncrypted files will be moved to a timestamped JournalBackups folder. Backups are not used automatically. Close any other copies of this tool first. All accounts using this Windows journal folder are affected.`n`nContinue?"
    if([Windows.MessageBox]::Show($window,$message,'Clear all active journals','YesNo','Warning','No') -ne 'Yes'){return}
    try{
        $result=Clear-TZJournal -JournalDirectory $script:journalDirectory -Confirmed
        $StatusLabel.Text="Cleared $($result.Count) active journals. Scan again before making changes."
        [void][Windows.MessageBox]::Show($window,"Cleared $($result.Count) active journals.`n`nBackup: $($result.BackupPath)`n`nCalendar entries were not changed.",'Journals cleared','OK','Information')
    }catch{[void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Journal clearing stopped','OK','Error')}
    finally{
        $script:scan=$null;$PreviewGrid.ItemsSource=$null;$IntentCheck.IsChecked=$false;$NoticeCheck.IsChecked=$false
        Update-FixButton
    }
})
$window.FindName('LastErrorButton').Add_Click({
    $text=Get-TZLastFailure -JournalDirectory $script:journalDirectory
    $dialog=[Windows.Window]::new();$dialog.Title='Last repair error';$dialog.Owner=$window
    $dialog.Width=760;$dialog.Height=470;$dialog.WindowStartupLocation='CenterOwner'
    $box=[Windows.Controls.TextBox]::new();$box.Text=$text;$box.IsReadOnly=$true
    $box.AcceptsReturn=$true;$box.TextWrapping='Wrap';$box.VerticalScrollBarVisibility='Auto'
    $box.Margin='16';$dialog.Content=$box
    [void]$dialog.ShowDialog()
})
foreach ($name in 'SirkitLogo','ProvinceBox','RegionBox','RegionPanel','TimezonePair','ScanButton','OutlookTimezoneButton','SeriesButton','OperationProgress','ProgressText','TitleTimeCheck','ShowExcludedCheck','PreviewGrid','AccountLabel','IntentCheck','NoticeCheck','FixButton','StatusLabel') {
    Set-Variable -Name $name -Value $window.FindName($name) -Scope Script
}
$logoReader=[Xml.XmlReader]::Create((Join-Path $PSScriptRoot 'assets\sirkit-logo.xaml'))
try { $SirkitLogo.Source=[Windows.Markup.XamlReader]::Load($logoReader) } finally { $logoReader.Dispose() }
$ProvinceBox.SelectedIndex=switch ($Province) { 'Alberta' {1} 'Manitoba' {2} default {0} }
$RegionBox.SelectedIndex=if ($Province -eq 'British Columbia (Mountain)') {1} else {0}

function Get-SelectedProfile {
    $key=[string]$ProvinceBox.SelectedItem.Tag
    if ($key -eq 'British Columbia') { $key=[string]$RegionBox.SelectedItem.Tag }
    Get-TZProfile $key
}

function Update-ProvinceUi {
    $RegionPanel.Visibility=if ($ProvinceBox.SelectedItem.Tag -eq 'British Columbia') {'Visible'} else {'Collapsed'}
    $profile=Get-SelectedProfile
    $TimezonePair.Text=$profile.OldLabel+'  →  '+$profile.NewLabel
    $script:scan=$null; $PreviewGrid.ItemsSource=$null
    $IntentCheck.IsChecked=$false; $NoticeCheck.IsChecked=$false
    $AccountLabel.Text=if($null -ne $script:connectedIdentity){'Connected: '+$script:connectedIdentity.Account}else{'Not connected'}
    $StatusLabel.Text='Ready'
    $ProgressText.Text=''; $OperationProgress.Visibility='Collapsed'; $OperationProgress.IsIndeterminate=$false; $OperationProgress.Value=0
    $ShowExcludedCheck.IsChecked=$false
    if ($Demo -or $ValidateUi) {
        Set-Scan (Get-ABDemo $profile.Key)
        $StatusLabel.Text='Demo data · Calendar access and repairs disabled'
    }
    Update-FixButton
}

function Update-FixButton {
    $FixButton.IsEnabled = (-not $Demo -and -not $ValidateUi -and $null -eq $script:operation -and $null -ne $script:scan -and
        $IntentCheck.IsChecked -eq $true -and $NoticeCheck.IsChecked -eq $true -and
        $script:scan.Profile -eq (Get-SelectedProfile).Key)
    $OutlookTimezoneButton.IsEnabled=(-not $Demo -and -not $ValidateUi -and $null -eq $script:operation -and
        $null -ne $script:scan -and $script:scan.Profile -eq (Get-SelectedProfile).Key)
    $SeriesButton.IsEnabled=(-not $Demo -and -not $ValidateUi -and $null -eq $script:operation -and $null -ne $script:scan -and
        $PreviewGrid.SelectedItems.Count -eq 1 -and $PreviewGrid.SelectedItem.Kind -eq 'seriesMaster')
}
function Set-Busy([bool]$Busy) {
    $window.FindName('ClearJournalButton').IsEnabled=(-not $Busy -and -not $Demo -and -not $ValidateUi)
    $ScanButton.IsEnabled = -not $Busy
    $ProvinceBox.IsEnabled = -not $Busy; $RegionBox.IsEnabled = -not $Busy
    $PreviewGrid.IsEnabled = -not $Busy
    $ShowExcludedCheck.IsEnabled = -not $Busy; $TitleTimeCheck.IsEnabled = -not $Busy
    $IntentCheck.IsEnabled = -not $Busy; $NoticeCheck.IsEnabled = -not $Busy
    Update-FixButton
}
function Set-Scan($Value) {
    if ($Value.Profile -ne (Get-SelectedProfile).Key) { throw 'The province changed. Scan again before reviewing or repairing.' }
    $script:scan = $Value
    $script:connectedIdentity=$Value.Identity
    Refresh-Preview
    Update-FixButton
}

function Refresh-Preview {
    if($null -eq $script:scan){$PreviewGrid.ItemsSource=$null;return}
    foreach($row in $script:scan.Rows){$row.Selected=$false}
    $Value=$script:scan
    $shown=@($Value.Rows | Where-Object { -not $_.Handled -and -not $_.Original.isAllDay -and ($ShowExcludedCheck.IsChecked -eq $true -or $_.DefaultVisible) })
    $PreviewGrid.ItemsSource=$shown
    $AccountLabel.Text = "Connected: $($Value.Identity.Account) · $($shown.Count) shown · $(@($Value.Rows).Count-$shown.Count) excluded"
    if ($Value.PSObject.Properties['SkippedCalendars'] -and @($Value.SkippedCalendars).Count) {
        $AccountLabel.Text += " | $(@($Value.SkippedCalendars).Count) shared or ownership-unknown calendars excluded"
    }
    if ($Value.Identity.PSObject.Properties['AccessNote'] -and $Value.Identity.AccessNote) {
        $AccountLabel.Text += ' | ' + $Value.Identity.AccessNote
    }
    $IntentCheck.IsChecked = $false; $NoticeCheck.IsChecked = $false
    Update-FixButton
}

function Update-SingleReviewState($Rows,[switch]$Succeeded){
    if($null -eq $script:scan){return}
    $states=$null
    $failureMarked=$false
    if(-not $Succeeded){try{$states=Get-ABJournalIndex $script:scan.Identity $script:journalDirectory}catch{}}
    foreach($row in $Rows){
        if($Succeeded -or ($null -ne $states -and $states.ContainsKey($row.Id) -and $states[$row.Id] -like 'Verified*')){
            $row.Handled=$true;$row.CanResolve=$false;$row.Eligible=$false;$row.DefaultVisible=$false;$row.Result='Verified'
        }elseif(-not $failureMarked -or ($null -ne $states -and $states.ContainsKey($row.Id))){
            $failureMarked=$true
            $row.CanResolve=$true;$row.Eligible=$false;$row.Result='Stopped - review before retrying';$row.DefaultVisible=$true
        }
    }
    Refresh-Preview
}
$PreviewGrid.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent,[Windows.RoutedEventHandler]{
    param($sender,$eventArgs)
    $button=$eventArgs.OriginalSource
    if($button -isnot [Windows.Controls.Button] -or $button.Tag -ne 'ResolveJournalRow'){return}
    $eventArgs.Handled=$true
    if($Demo -or $ValidateUi -or $null -ne $script:operation -or $null -eq $script:scan){return}
    $row=$button.DataContext
    if(-not $row.CanResolve){return}
    $message="Mark '$($row.Subject)' as manually resolved?`n`nConfirm you have checked the meeting and any replacement in Outlook. This removes its active failure/recovery records and hides it from future scans on this Windows account. It does not change or undo calendar entries. A backup is retained."
    if([Windows.MessageBox]::Show($window,$message,'Manually resolved','YesNo','Warning','No') -ne 'Yes'){return}
    try{
        $resolved=Resolve-TZJournalRow $script:scan.Identity $row $script:journalDirectory -Confirmed
        foreach($item in $script:scan.Rows){if($item.Id -cin $resolved.ResolvedIds -or ($item.MasterId -and $item.MasterId -cin $resolved.ResolvedIds)){$item.Handled=$true;$item.CanResolve=$false;$item.Eligible=$false}}
        Refresh-Preview;$StatusLabel.Text='Marked manually resolved. No calendar changes were made.'
    }catch{[void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Could not mark resolved','OK','Error')}
})
function Update-ProgressDisplay($Update) {
    $OperationProgress.Visibility='Visible'
    if($Update.Total -gt 0){$OperationProgress.IsIndeterminate=$false;$OperationProgress.Value=100*$Update.Completed/$Update.Total}
    else{$OperationProgress.IsIndeterminate=$true}
    $ProgressText.Text=$Update.Message
    $StatusLabel.Text=$Update.Message
    if($Update.Account){$AccountLabel.Text='Connected: '+$Update.Account;$ScanButton.Content='Scan again'}
}

function Update-SeriesReviewState($Plan,$Result,[switch]$Stopped) {
    if($null -eq $script:scan -or $script:scan.Profile -ne $Plan.Profile){return}
    $seriesRows=@($PreviewGrid.Items|Where-Object Kind -eq 'seriesMaster')
    $current=@($seriesRows|Where-Object {$_.MasterId -ceq $Plan.MasterId})|Select-Object -First 1
    $index=[array]::IndexOf($seriesRows,$current)
    foreach($row in $script:scan.Rows|Where-Object {$_.Kind -eq 'seriesMaster' -and $_.MasterId -ceq $Plan.MasterId}){
        $row.Selected=$false;$row.Eligible=$false
        if($Stopped){
            $row.Handled=$false
            $row.CanResolve=$true;$row.Result='Stopped · preview again to reconcile prior changes'
            $row.Decision='Repair stopped · double-click for a fresh preview'
            $row.DefaultVisible=$true
        }else{
            $row.Result="$($Result.Count) verified; $($Result.AlreadyAppliedCount) already correct"
            if($Result.PSObject.Properties['MigratedSeries'] -and $Result.MigratedSeries){
                $row.Handled=$true;$row.CanResolve=$false;$row.Decision='Future series migrated and verified · original series ended'
                $row.DefaultVisible=$false
            }elseif($Plan.UnresolvedCount -gt 0){
                $row.Decision='Reviewed changes completed · unresolved occurrences need review'
                $row.DefaultVisible=$true
            }elseif($Plan.RangeType -eq 'noEnd'){
                $row.Decision='Chosen date range completed · double-click to review another range'
                $row.DefaultVisible=$true
            }else{
                $row.Decision='Future occurrences completed · double-click to review again'
                $row.DefaultVisible=$false
            }
        }
    }
    Refresh-Preview
    # Move to the next visible recurring row, wrapping around when needed.
    if($seriesRows.Count -gt 1){
        for($offset=1;$offset -le $seriesRows.Count;$offset++){
            $next=$seriesRows[($index+$offset)%$seriesRows.Count]
            if($next.MasterId -cne $Plan.MasterId -and $PreviewGrid.Items.Contains($next)){
                $PreviewGrid.SelectedItem=$next;$PreviewGrid.ScrollIntoView($next);break
            }
        }
    }
    Update-FixButton
}

function Get-OpenEndedSeriesDate {
    [xml]$markup=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Choose a future-occurrence end date" Width="540" Height="270" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" Background="White" Foreground="#002849">
 <StackPanel Margin="24"><TextBlock Text="This series repeats forever" FontSize="20" FontWeight="SemiBold"/><TextBlock Text="Choose the last date to update for this run. Meetings after this date will keep their existing time zone." TextWrapping="Wrap" Margin="0,12,0,14"/><DatePicker x:Name="EndPicker"/><Button x:Name="PreviewButton" Content="Preview future occurrences" Background="#0252C6" Foreground="White" Padding="14,9" Margin="0,18,0,0"/></StackPanel>
</Window>
'@
    $r=[Xml.XmlNodeReader]::new($markup);$dialog=[Windows.Markup.XamlReader]::Load($r);$r.Close();$dialog.Owner=$window
    $picker=$dialog.FindName('EndPicker');$picker.SelectedDate=[datetime]::UtcNow.AddYears(2).Date
    $preview=$dialog.FindName('PreviewButton')
    $preview.Add_Click({if($null -ne $picker.SelectedDate){$dialog.DialogResult=$true}})
    if($dialog.ShowDialog() -eq $true){return $picker.SelectedDate}
    return $null
}

function Test-SeriesDiscardWarning($Plan){
    if(-not $Plan.PSObject.Properties['RecreateFutureChanges'] -or -not $Plan.RecreateFutureChanges){return $false}
    # Older recovery plans did not count cancellations beyond the preview horizon.
    if(-not $Plan.PSObject.Properties['FutureCancellationCount']){return $true}
    return ($Plan.FutureCancellationCount -gt 0 -or $Plan.DiscardedExceptionCount -gt 0)
}
function Show-SeriesConfirmation($Plan,[switch]$ValidateOnly) {
    [xml]$markup=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Migrate future recurring series | Sirkit" Width="1100" Height="780" MinWidth="900" MinHeight="650" WindowStartupLocation="CenterOwner" Background="White" Foreground="#002849" FontFamily="Segoe UI" FontSize="14">
 <Grid Margin="24" Background="White"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <StackPanel><TextBlock Text="Migrate future recurring series" FontSize="24" FontWeight="SemiBold"/><TextBlock x:Name="Summary" TextWrapping="Wrap" Margin="0,10,0,12"/></StackPanel>
  <Border Grid.Row="1" Background="#FFF4E6" BorderBrush="#E98607" BorderThickness="1" Padding="12" Margin="0,0,0,14"><TextBlock TextWrapping="Wrap" Text="A replacement series will be created with the new time zone, then the original future series will be ended. This sends new invitations and updates/cancellations, rebooks rooms, resets attendee responses, and generates a new Teams link for online meetings. Creating the replacement and ending the original are separate operations: if interrupted, both may temporarily exist. Reopen the same row to reconcile and resume; do not create another series manually. When future edits or cancellations exist, a warning asks you to confirm recreating the regular schedule before changes are applied. Unidentifiable dates and attachments still require Outlook handling. Earlier edits and cancellations stay in the original series."/></Border>
  <TabControl Grid.Row="2"><TabItem Header="Dates"><DataGrid x:Name="Occurrences" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" HeadersVisibility="Column" Background="White"><DataGrid.Columns>
   <DataGridTextColumn Header="Meeting" Binding="{Binding Subject}" Width="*" MinWidth="180"/>
   <DataGridTextColumn Header="Before (new zone)" Binding="{Binding CurrentLocal}" Width="285"/>
   <DataGridTextColumn Header="Replacement (new zone)" Binding="{Binding ProposedLocal}" Width="285"/>
  </DataGrid.Columns></DataGrid></TabItem>
  <TabItem Header="People and rooms"><TextBox x:Name="People" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" Margin="12"/></TabItem>
  <TabItem Header="Replacement description"><Grid Margin="12"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><TextBlock x:Name="DescriptionNote" TextWrapping="Wrap" Margin="0,0,0,8"/><TextBox x:Name="DescriptionDraft" Grid.Row="1" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/></Grid></TabItem>
  <TabItem Header="Original description"><TextBox x:Name="OriginalDescription" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" Margin="12"/></TabItem></TabControl>
  <StackPanel Grid.Row="3" Margin="0,16,0,0"><CheckBox x:Name="Acknowledge" Content="I reviewed the dates, people, rooms and replacement description."/><CheckBox x:Name="Notice" Content="I understand new invitations/cancellations will be sent and the series identity and Teams link can change." Margin="0,8,0,0"/><DockPanel Margin="0,16,0,0"><Button x:Name="ApplyButton" DockPanel.Dock="Right" Content="Migrate future series" Padding="18,10" Background="#0252C6" Foreground="White" IsEnabled="False"/><TextBlock x:Name="Excluded" VerticalAlignment="Center" TextWrapping="Wrap"/></DockPanel></StackPanel>
 </Grid>
</Window>
'@
    $r=[Xml.XmlNodeReader]::new($markup);$dialog=[Windows.Markup.XamlReader]::Load($r);$r.Close()
    if(-not $ValidateOnly){$dialog.Owner=$window}
    $profile=Get-TZProfile $Plan.Profile
    $end=if($Plan.RangeType -eq 'noEnd'){'no end date'}else{$Plan.EndDate}
    $action=if($Plan.Operation -eq 'ReplaceFutureSeries'){'All occurrences are future: the original series will be cancelled after replacement verification.'}else{"The original series will end $($Plan.OldEnd); earlier meetings stay in the old series."}
    $resume=if($Plan.ResumeJournalPath){'Resume: the replacement already exists; it will not be created again.'}else{''}
    $dialog.FindName('Summary').Text="$($Plan.Subject)`n$($Plan.Identity.Account) · starts $($Plan.FirstDate) · ends $end`n$($profile.NewLabel) · all preview times use this zone`n$action $resume"
    if($Plan.PSObject.Properties['RetainedExceptionCount'] -and $Plan.RetainedExceptionCount -gt 0){$dialog.FindName('Summary').Text+="`n$($Plan.RetainedExceptionCount) earlier edited occurrences stay in the original series and do not block this migration."}
    if($Plan.PSObject.Properties['RetainedCancellationCount'] -and $Plan.RetainedCancellationCount -gt 0){$dialog.FindName('Summary').Text+="`n$($Plan.RetainedCancellationCount) earlier cancellations stay in the original series and do not block this migration."}
    $dialog.FindName('Occurrences').ItemsSource=@($Plan.Rows)
    $dialog.FindName('Excluded').Text="Showing $($Plan.Rows.Count) sample dates. The replacement keeps the series end condition."
    $dialog.FindName('People').Text=(@($Plan.NewPayload.attendees|ForEach-Object {"$($_.type): $($_.emailAddress.name) <$($_.emailAddress.address)>"}) -join "`n")+"`n`nLocation: $($Plan.NewPayload.location.displayName)"
    $draft=$dialog.FindName('DescriptionDraft');$draft.Text=$Plan.ReplacementDescription
    $draft.IsReadOnly=(-not $Plan.IsOnlineMeeting -or [bool]$Plan.ResumeJournalPath)
    $dialog.FindName('OriginalDescription').Text=$Plan.OriginalDescription
    $dialog.FindName('DescriptionNote').Text=if($Plan.IsOnlineMeeting){'Review and edit this plain-text replacement description. The old join block was removed; a new Teams link will be generated. Formatting and embedded images are not copied. Compare the original description for notes to retain.'}else{'The original description and formatting will be copied. This tab shows a plain-text preview.'}
    $recreate=($Plan.PSObject.Properties['RecreateFutureChanges'] -and $Plan.RecreateFutureChanges)
    if($recreate -and (Test-SeriesDiscardWarning $Plan)){
        $dialog.FindName('Summary').Text+="`nRECREATE MODE: $($Plan.RecreatedCancellationCount) cancelled dates will return; $($Plan.DiscardedExceptionCount) edited occurrences will use the regular series details."
    }
    $ack=$dialog.FindName('Acknowledge');$notice=$dialog.FindName('Notice');$apply=$dialog.FindName('ApplyButton')
    $ack.Add_Click({$apply.IsEnabled=($ack.IsChecked -eq $true -and $notice.IsChecked -eq $true -and $Plan.Rows.Count -gt 0)})
    $notice.Add_Click({$apply.IsEnabled=($ack.IsChecked -eq $true -and $notice.IsChecked -eq $true -and $Plan.Rows.Count -gt 0)})
    $apply.Add_Click({
        if($ack.IsChecked -ne $true -or $notice.IsChecked -ne $true -or $Plan.Rows.Count -eq 0){return}
        if($recreate){
            if(Test-SeriesDiscardWarning $Plan){
                $warning="This replacement will recreate cancelled future dates and discard edits made to individual future occurrences. Their times, duration, notes, attendees and locations will revert to the regular series. New invitations and updates/cancellations may be sent.`n`nFor an open-ended series, this also applies beyond the sample dates shown.`n`nProceed with recreating the regular schedule?"
                if([Windows.MessageBox]::Show($dialog,$warning,'Future edits and cancellations','YesNo','Warning','No') -ne 'Yes'){return}
            }
            $Plan.DiscardChangesConfirmed=$true
        }
        if(-not $draft.IsReadOnly){$Plan.ReplacementDescription=$draft.Text}
        $dialog.DialogResult=$true
    })
    if($ValidateOnly){return $dialog}
    return ($dialog.ShowDialog() -eq $true)
}

# A single background runspace keeps sign-in/network requests off the window thread.
$worker = [runspacefactory]::CreateRunspace()
$worker.ApartmentState = 'STA'; $worker.ThreadOptions = 'ReuseThread'; $worker.Open()
$timer = [Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [timespan]::FromMilliseconds(200)
function Start-Work([string]$Kind, [object[]]$Arguments) {
    $ps = [powershell]::Create()
    $ps.Runspace = $worker
    $code = @'
param($modulePath,$kind,$config,$profile,$scan,$rows,$journalDirectory,$progressQueue)
$ErrorActionPreference = 'Stop'
Import-Module $modulePath
if($kind -eq 'Connect'){Connect-TZSession -Config $config -ProgressQueue $progressQueue}
elseif ($kind -eq 'Scan') { Get-ABScan -Config $config -Profile $profile -JournalDirectory $journalDirectory -ProgressQueue $progressQueue -IncludeTimeInTitle:$rows.IncludeTimeInTitle }
else {
    if ($scan.Profile -ne $profile) { throw 'The province changed. Scan again before repairing.' }
    if($kind -eq 'Settings'){Set-TZOutlookTimezone -Config $config -Scan $scan -JournalDirectory $journalDirectory -Confirmed -ProgressQueue $progressQueue}
    elseif($kind -eq 'SeriesPreview'){Get-TZMigrationPlan -Config $config -Scan $scan -Row $rows.Row -RecreateFutureChanges:$rows.RecreateFutureChanges -ProgressQueue $progressQueue -JournalDirectory $journalDirectory}
    elseif($kind -eq 'SeriesFix'){Invoke-TZMigration -Config $config -Plan $scan -JournalDirectory $journalDirectory -ProgressQueue $progressQueue -Confirmed}
    else{Invoke-ABApply $config $scan $rows $journalDirectory $progressQueue}
}
'@
    [void]$ps.AddScript($code).AddArgument((Join-Path $PSScriptRoot 'TimezoneEngine.psm1')).AddArgument($Kind)
    foreach ($arg in $Arguments) { [void]$ps.AddArgument($arg) }
    $queue=[System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    [void]$ps.AddArgument($queue)
    $script:operation = @{ PowerShell = $ps; Handle = $ps.BeginInvoke(); Kind = $Kind; Queue=$queue; Review=$Arguments[2]; Rows=$Arguments[3] }
    Set-Busy $true
    $starting=switch($Kind){'Connect'{'Starting Microsoft sign-in…'} 'Scan'{'Scanning calendars…'} 'SeriesPreview'{'Preparing series migration preview…'} 'SeriesFix'{'Starting confirmed series migration…'} default{'Preparing confirmed change…'}}
    Update-ProgressDisplay ([pscustomobject]@{Stage='Starting';Message=$starting;Completed=0;Total=0;Account=''})
    $timer.Start()
}
function Receive-WorkResult($Op) {
    $result=$Op.PowerShell.EndInvoke($Op.Handle)
    # HadErrors stays true even for a caught 412 that was subsequently repaired.
    # EndInvoke surfaces terminating failures; only remaining stream errors are failures here.
    if($Op.PowerShell.Streams.Error.Count -gt 0){throw ($Op.PowerShell.Streams.Error|Out-String)}
    return ,$result
}
$timer.Add_Tick({
    if ($null -eq $script:operation) { return }
    $update=$null
    while($script:operation.Queue.TryDequeue([ref]$update)){Update-ProgressDisplay $update}
    if (-not $script:operation.Handle.IsCompleted) { return }
    $timer.Stop()
    $op = $script:operation
    $seriesPlan=$null
    try {
        $result = Receive-WorkResult $op
        if($op.Kind -eq 'Connect'){
            $script:connectedIdentity=$result[0]
            $AccountLabel.Text='Connected: '+$script:connectedIdentity.Account
            $ScanButton.Content='Scan calendars'
            $StatusLabel.Text='Signed in. This session will be reused while this window stays open. Choose a province and scan.'
        } elseif ($op.Kind -eq 'Scan') {
            Set-Scan $result[0]
            $StatusLabel.Text = "Scan complete. $($PreviewGrid.Items.Count) entries need review. No events changed."
            if($result[0].UnverifiedSeriesCount -gt 0){$StatusLabel.Text+=" $($result[0].UnverifiedSeriesCount) series had no verified upcoming affected occurrence; see excluded entries."}
        } elseif($op.Kind -eq 'SeriesPreview') {
            $seriesPlan=$result[0]
            $StatusLabel.Text="Series migration preview ready. No changes made."
        } elseif($op.Kind -eq 'Settings') {
            $StatusLabel.Text=if($result[0].Changed){"Outlook mailbox time zone changed to $($result[0].TimeZone). No meetings rescheduled."}else{'Your Outlook mailbox already uses the selected time zone.'}
            $IntentCheck.IsChecked=$false; $NoticeCheck.IsChecked=$false
        } elseif($op.Kind -eq 'SeriesFix') {
            Update-SeriesReviewState $op.Review $result[0]
            $StatusLabel.Text='Future series migrated and verified in the new time zone. Select the next series; no rescan needed.'
        } else {
            $StatusLabel.Text = "$($result[0].Count) events verified."
            if($result[0].PSObject.Properties['AlreadyAppliedCount']){$StatusLabel.Text+=" $($result[0].AlreadyAppliedCount) already correct; no repeat update sent."}
            Update-SingleReviewState $op.Rows -Succeeded
            $StatusLabel.Text+=" Remaining meetings are ready for review."
        }
        $OperationProgress.IsIndeterminate=$false; $OperationProgress.Value=100
        $ProgressText.Text=$StatusLabel.Text
    } catch {
        $OperationProgress.IsIndeterminate=$false; $OperationProgress.Value=0
        $StatusLabel.Text = 'Operation stopped. See the error details.'; $ProgressText.Text=$StatusLabel.Text
        if($op.Kind -eq 'SeriesFix'){
            Update-SeriesReviewState $op.Review $null -Stopped
            $StatusLabel.Text='Series stopped. Other entries are still available. Open this series again for a fresh preview before continuing.'
            $ProgressText.Text=$StatusLabel.Text
        }elseif($op.Kind -eq 'Fix'){
            Update-SingleReviewState $op.Rows
        }elseif($op.Kind -eq 'SeriesPreview'){
            $op.Rows.Row.CanResolve=$true;$op.Rows.Row.Result='Preview failed';$PreviewGrid.Items.Refresh()
        }
        [void][Windows.MessageBox]::Show($window, $_.Exception.Message, 'Timezone Change Tool', 'OK', 'Warning')
    } finally {
        $op.PowerShell.Dispose(); $script:operation = $null
        Set-Busy $false
    }
    if($null -ne $seriesPlan -and (Show-SeriesConfirmation $seriesPlan)){
        Start-Work SeriesFix @($script:config,$seriesPlan.Profile,$seriesPlan,$seriesPlan.Rows,$script:journalDirectory)
    }
})
$PreviewGrid.Add_SelectionChanged({Update-FixButton})
function Open-SelectedRecurringSeries {
    try {
        if($Demo -or $ValidateUi -or $null -ne $script:operation -or $null -eq $script:scan){return}
        if($PreviewGrid.SelectedItems.Count -ne 1 -or $PreviewGrid.SelectedItem.Kind -ne 'seriesMaster'){throw 'Select one recurring-series row first.'}
        $row=$PreviewGrid.SelectedItem; $through=$null
        Start-Work SeriesPreview @($script:config,(Get-SelectedProfile).Key,$script:scan,[pscustomobject]@{Row=$row;ThroughDate=$through;RecreateFutureChanges=$true},$script:journalDirectory)
    } catch {[void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Cannot preview recurring series','OK','Warning')}
}
$SeriesButton.Add_Click({Open-SelectedRecurringSeries})
$PreviewGrid.Add_MouseDoubleClick({
    param($sender,$eventArgs)
    if($null -ne $script:operation){return}
    $container=[Windows.Controls.ItemsControl]::ContainerFromElement($PreviewGrid,$eventArgs.OriginalSource)
    if($container -isnot [Windows.Controls.DataGridRow] -or $container.Item.Kind -ne 'seriesMaster'){return}
    $eventArgs.Handled=$true
    $PreviewGrid.SelectedItems.Clear();$PreviewGrid.SelectedItem=$container.Item
    Open-SelectedRecurringSeries
})
$TitleTimeCheck.Add_Click({
    if($null -ne $script:scan){
        $enabled=($TitleTimeCheck.IsChecked -eq $true)
        $script:scan|Add-Member -Force IncludeTimeInTitle $enabled
        $zone=[TimeZoneInfo]::FindSystemTimeZoneById($script:scan.Target)
        foreach($row in $script:scan.Rows|Where-Object {$_.Eligible -and -not $_.Handled}){
            $fresh=New-ABProposal -Event $row.Original -TargetZone $zone -TargetAlias $script:scan.Target -Profile $script:scan.Profile -IncludeTimeInTitle:$enabled
            $row.Subject=$fresh.Subject;$row.Patch=$fresh.Patch;$row.Eligible=$fresh.Eligible;$row.Decision=$fresh.Decision
        }
        Refresh-Preview
    }
    $StatusLabel.Text='Title preview updated. Remaining meetings are ready for review.'
})
$ShowExcludedCheck.Add_Click({Refresh-Preview})
$ScanButton.Add_Click({
    try {
        $profile=Get-SelectedProfile
        if ($Demo) { Set-Scan (Get-ABDemo $profile.Key); $StatusLabel.Text = 'Demo data · Calendar access and repairs disabled'; return }
        $script:scan = $null; $PreviewGrid.ItemsSource = $null
        Start-Work Scan @($script:config,$profile.Key,$null,[pscustomobject]@{IncludeTimeInTitle=($TitleTimeCheck.IsChecked -eq $true)},$script:journalDirectory)
    } catch { [void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Cannot scan','OK','Warning') }
})
$IntentCheck.Add_Click({ Update-FixButton }); $NoticeCheck.Add_Click({ Update-FixButton })
$OutlookTimezoneButton.Add_Click({
    try {
        if($Demo -or $ValidateUi -or $null -eq $script:scan){return}
        $profile=Get-SelectedProfile
        $message="Set the Outlook mailbox time zone for $($script:scan.Identity.Account) to $($profile.NewLabel)?`n`nThis changes the default mailbox time zone used by Outlook on the web. Desktop/mobile clients may use their own settings.`n`nIt does not reschedule existing meetings. Microsoft will request mailbox-settings permission if needed."
        if([Windows.MessageBox]::Show($window,$message,'Change Outlook time zone','YesNo','Question') -ne 'Yes'){return}
        Start-Work Settings @($script:config,$profile.Key,$script:scan,$null,$script:journalDirectory)
    } catch {[void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Cannot change Outlook time zone','OK','Warning')}
})
$FixButton.Add_Click({
    try {
        if ($Demo -or $ValidateUi -or $null -eq $script:scan) { return }
        [void]$PreviewGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true)
        [void]$PreviewGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
        $selected = @($script:scan.Rows | Where-Object { $_.Selected -and $_.Eligible })
        if ($selected.Count -eq 0) { throw 'Tick the Fix checkbox beside at least one eligible event.' }
        if ($selected.Count -gt 25) { throw 'Select at most 25 events per repair.' }
        $profile=Get-SelectedProfile
        $message = "Update $($selected.Count) selected events for $($script:scan.Identity.Account)?`n`nRegion: $($profile.Name)`n$($profile.OldId) → $($profile.NewId)`n`nThis can send meeting updates to attendees. Recurring series are excluded. An encrypted local journal will be saved first."
        if ([Windows.MessageBox]::Show($window,$message,'Confirm selected corrections','YesNo','Warning') -ne 'Yes') { return }
        Start-Work Fix @($script:config,$profile.Key,$script:scan,$selected,$script:journalDirectory)
    } catch { [void][Windows.MessageBox]::Show($window,$_.Exception.Message,'Cannot repair','OK','Warning') }
})
$window.Add_Closing({
    param($sender,$eventArgs)
    if ($null -ne $script:operation) {
        $eventArgs.Cancel = $true
        $StatusLabel.Text = 'Wait for the current operation to finish before closing.'
    }
})
if ($Demo -or $ValidateUi) {
    $ScanButton.Content = 'Reload demo'
}
$ProvinceBox.Add_SelectionChanged({ Update-ProvinceUi })
$RegionBox.Add_SelectionChanged({ Update-ProvinceUi })
Update-ProvinceUi
$window.Add_ContentRendered({
    if (-not $script:launchSignInStarted -and -not $Demo -and -not $ValidateUi) {
        $script:launchSignInStarted=$true
        if($script:autoConnect){$ScanButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))}
        else{Start-Work Connect @($script:config,(Get-SelectedProfile).Key,$null,$null,$script:journalDirectory)}
    }
})
try {
    if ($ValidateUi) {
        foreach($validationProvince in 'Manitoba','British Columbia','British Columbia (Mountain)','Alberta') {
            if ((Get-SelectedProfile).Key -eq $validationProvince) { continue }
            # Exercise the actual selection events, including invalidating an acknowledged preview.
            $script:scan.Rows[0].Selected=$true
            $IntentCheck.IsChecked=$true; $NoticeCheck.IsChecked=$true
            $ProvinceBox.SelectedIndex=switch ($validationProvince) {'Manitoba' {2} 'Alberta' {1} default {0}}
            $RegionBox.SelectedIndex=if($validationProvince -eq 'British Columbia (Mountain)'){1}else{0}
            if ($script:scan.Profile -ne $validationProvince -or @($script:scan.Rows | Where-Object Selected).Count -gt 0 -or
                $IntentCheck.IsChecked -or $NoticeCheck.IsChecked -or $FixButton.IsEnabled) {
                throw 'Province selection failed to invalidate the prior preview and acknowledgements.'
            }
        }
        $ProvinceBox.SelectedIndex=switch ($Province) {'Manitoba' {2} 'Alberta' {1} default {0}}
        $RegionBox.SelectedIndex=if($Province -eq 'British Columbia (Mountain)'){1}else{0}
        Update-ProvinceUi
        $defaultCount=$PreviewGrid.Items.Count
        if(@($PreviewGrid.Items | Where-Object {-not $_.DefaultVisible}).Count){throw 'An excluded entry appeared in the default view.'}
        $ShowExcludedCheck.IsChecked=$true; Refresh-Preview
        if($PreviewGrid.Items.Count -ne $script:scan.Rows.Count){throw 'Excluded entries toggle did not reveal the full scanned set.'}
        $ShowExcludedCheck.IsChecked=$false; Refresh-Preview
        if($PreviewGrid.Items.Count -ne $defaultCount){throw 'Default filter did not restore the expected entries.'}
        Update-ProgressDisplay ([pscustomobject]@{Stage='SignedIn';Message='Sign-in complete. Scanning calendars…';Completed=0;Total=0;Account='demo@example.invalid'})
        if($AccountLabel.Text -notmatch 'Connected:' -or -not $OperationProgress.IsIndeterminate){throw 'Sign-in notification/progress display failed.'}
        Update-ProgressDisplay ([pscustomobject]@{Stage='Complete';Message="Demo scan complete · $defaultCount entries need review";Completed=1;Total=1;Account=''})
        if($OperationProgress.Value -ne 100 -or $OperationProgress.IsIndeterminate){throw 'Completion progress display failed.'}
        if($OutlookTimezoneButton.IsEnabled){throw 'Outlook preference changes must be disabled in demo/validation mode.'}
        if($SeriesButton.IsEnabled){throw 'Recurring-series changes must be disabled in demo/validation mode.'}
        $demoPlan=[pscustomobject]@{Profile=$Province;Subject='Weekly team meeting';Identity=@{Account='demo@example.invalid'}
            Rows=@($script:scan.Rows|Where-Object Eligible);RangeType='endDate';EndDate='2027-11-10';SkippedCount=1;LastOccurrence=$null;Until=[datetimeoffset]'2027-11-11T00:00:00Z';
            Operation='SplitSeries';OldEnd='2026-11-02';FirstDate='2026-11-03';ResumeJournalPath='';IsOnlineMeeting=$true;
            NewPayload=@{attendees=@(@{type='required';emailAddress=@{name='Demo attendee';address='attendee@example.invalid'}});location=@{displayName='Meeting room'}};
            ReplacementDescription='Weekly team check-in';OriginalDescription='Weekly team check-in. Microsoft Teams meeting [original join details]'}
        $seriesDialog=Show-SeriesConfirmation $demoPlan -ValidateOnly
        if($seriesDialog.FindName('ApplyButton').IsEnabled -or $seriesDialog.FindName('Acknowledge').IsChecked -or $seriesDialog.FindName('Notice').IsChecked){throw 'Series migration must require two unchecked acknowledgements before applying.'}
        if($window.FindName('RecreateFutureCheck')){throw 'The main recreate checkbox must not be present.'}
        $demoPlan|Add-Member RecreateFutureChanges $true
        $demoPlan|Add-Member RecreatedCancellationCount 2
        $demoPlan|Add-Member DiscardedExceptionCount 1
        $demoPlan|Add-Member DiscardChangesConfirmed $false
        $recreateDialog=Show-SeriesConfirmation $demoPlan -ValidateOnly
        if($recreateDialog.FindName('DiscardChanges') -or $recreateDialog.FindName('ApplyButton').IsEnabled -or $demoPlan.DiscardChangesConfirmed -or -not (Test-SeriesDiscardWarning $demoPlan)){throw 'Recreate preview must require confirmation through a warning, without a discard checkbox.'}
        if($recreateDialog.FindName('Summary').Text -notmatch '2 cancelled dates will return; 1 edited occurrences'){throw 'Recreate preview must explain affected dates and edits.'}
        $warningCase=[pscustomobject]@{RecreateFutureChanges=$true;FutureCancellationCount=0;DiscardedExceptionCount=0}
        if(Test-SeriesDiscardWarning $warningCase){throw 'An unchanged regular series should not prompt to discard changes.'}
        $warningCase.FutureCancellationCount=1
        if(-not(Test-SeriesDiscardWarning $warningCase)){throw 'Future cancellations beyond displayed dates must still trigger the warning.'}
        $warningCase.FutureCancellationCount=0;$warningCase.DiscardedExceptionCount=1
        if(-not(Test-SeriesDiscardWarning $warningCase)){throw 'Custom future edits must trigger the warning.'}
        if($SeriesPreviewImagePath){
            $seriesDialog.Measure([Windows.Size]::new(1100,780));$seriesDialog.Arrange([Windows.Rect]::new(0,0,1100,780));$seriesDialog.UpdateLayout()
            $surface=$seriesDialog.Content
            $surface.Measure([Windows.Size]::new(1052,732));$surface.Arrange([Windows.Rect]::new(0,0,1052,732));$surface.UpdateLayout()
            $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new(1052,732,96,96,[Windows.Media.PixelFormats]::Pbgra32);$bitmap.Render($surface)
            $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new();$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream=[IO.File]::Create($SeriesPreviewImagePath);try{$encoder.Save($stream)}finally{$stream.Dispose()}
        }
        $PreviewGrid.EnableRowVirtualization=$false
        $window.Measure([Windows.Size]::new(1320,860)); $window.Arrange([Windows.Rect]::new(0,0,1320,860)); $window.UpdateLayout()
        $window.Content.Measure([Windows.Size]::new(1272,812));$window.Content.Arrange([Windows.Rect]::new(0,0,1272,812));$window.Content.UpdateLayout()
        $seriesItem=@($PreviewGrid.Items|Where-Object Kind -eq 'seriesMaster')[0]
        $seriesContainer=$PreviewGrid.ItemContainerGenerator.ContainerFromItem($seriesItem)
        if($null -eq $seriesContainer){throw 'A recurring row was not generated for double-click validation.'}
        $doubleClick=[Windows.Input.MouseButtonEventArgs]::new([Windows.Input.Mouse]::PrimaryDevice,0,[Windows.Input.MouseButton]::Left)
        $doubleClick.RoutedEvent=[Windows.Controls.Control]::MouseDoubleClickEvent
        $doubleClick.Source=$seriesContainer
        $PreviewGrid.RaiseEvent($doubleClick)
        if(-not $doubleClick.Handled -or $PreviewGrid.SelectedItem -ne $seriesItem){throw 'Double-click did not route to the recurring-series command.'}
        if ($PreviewImagePath) {
            $content = $window.Content
            $content.Measure([Windows.Size]::new(1272,812)); $content.Arrange([Windows.Rect]::new(0,0,1272,812)); $content.UpdateLayout()
            $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new(1272,812,96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($content)
            $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
            $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream = [IO.File]::Create($PreviewImagePath)
            try { $encoder.Save($stream) } finally { $stream.Dispose() }
        }
        $scanBefore=$script:scan;$timeBefore=$script:scan.ScannedAt;$savedRows=$script:scan.Rows
        $savedDecision=$seriesItem.Decision;$savedVisible=$seriesItem.DefaultVisible
        $another=$seriesItem.PSObject.Copy();$another.MasterId='second-demo-series';$another.Id='second-demo-series';$another.Subject='Next recurring series'
        $script:scan.Rows=@($savedRows)+@($another);Refresh-Preview
        $PreviewGrid.SelectedItem=$seriesItem
        $completedPlan=[pscustomobject]@{Profile=$Province;MasterId=$seriesItem.MasterId;UnresolvedCount=0;RangeType='noEnd'}
        $completedResult=[pscustomobject]@{Count=1;AlreadyAppliedCount=0;MigratedSeries=$true}
        Update-SeriesReviewState $completedPlan $completedResult
        if(-not [object]::ReferenceEquals($script:scan,$scanBefore) -or $script:scan.ScannedAt -ne $timeBefore){throw 'Series completion replaced or extended the calendar scan.'}
        if($PreviewGrid.Items.Contains($seriesItem) -or -not $PreviewGrid.Items.Contains($another) -or $PreviewGrid.SelectedItem -ne $another){throw 'Series completion did not retain the list and select the next series.'}
        Update-SeriesReviewState $completedPlan $null -Stopped
        if(-not $PreviewGrid.Items.Contains($seriesItem) -or -not $PreviewGrid.Items.Contains($another)){throw 'A stopped series discarded the remaining calendar list.'}
        $seriesItem.Decision=$savedDecision;$seriesItem.DefaultVisible=$savedVisible;$seriesItem.Result='';$script:scan.Rows=$savedRows;Refresh-Preview
        $single=@($script:scan.Rows|Where-Object {$_.Eligible -and $_.Kind -eq 'singleInstance'})[0]
        $singleCopy=$single.PSObject.Copy();$singleCopy.Id='remaining-single';$singleCopy.Subject='Remaining single'
        $script:scan.Rows=@($savedRows)+@($singleCopy);Refresh-Preview
        $visibleBefore=$PreviewGrid.Items.Count
        $TitleTimeCheck.IsChecked=$true
        $TitleTimeCheck.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        if($null -eq $script:scan -or $PreviewGrid.Items.Count -ne $visibleBefore){throw 'Title checkbox must retain the meeting list.'}
        Update-SingleReviewState @($single) -Succeeded
        if($PreviewGrid.Items.Contains($single) -or -not $PreviewGrid.Items.Contains($singleCopy)){throw 'Single completion must hide only the completed row.'}
        $single.Handled=$false;$single.Eligible=$true;$script:scan.Rows=$savedRows;Refresh-Preview
        $fake=[powershell]::Create();$fake.Runspace=$worker
        try{
            [void]$fake.AddScript('try{Write-Error "Handled synthetic conflict" -ErrorAction Stop}catch{};[pscustomobject]@{Count=1}')
            $fakeOp=@{PowerShell=$fake;Handle=$fake.BeginInvoke()}
            $recovered=Receive-WorkResult $fakeOp
            if($recovered[0].Count -ne 1 -or -not $fake.HadErrors -or $fake.Streams.Error.Count -ne 0){throw 'A recovered conflict was incorrectly reported as a failed operation.'}
        }finally{$fake.Dispose()}
        Write-Output 'UI verified: migration warning and acknowledgements, retained list/next-series selection, stopped-series recovery, handled-conflict completion, double-click, filters and progress. No calendar writes.'
    } else { [void]$window.ShowDialog() }
} finally {
    $timer.Stop()
    $logout=[powershell]::Create();$logout.Runspace=$worker
    try{[void]$logout.AddScript('if(Get-Module TimezoneEngine){Disconnect-TZSession}');[void]$logout.Invoke()}finally{$logout.Dispose()}
    $worker.Close(); $worker.Dispose()
}
