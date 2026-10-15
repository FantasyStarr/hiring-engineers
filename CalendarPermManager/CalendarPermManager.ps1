<#
.SYNOPSIS
  Local GUI for bulk-assigning calendar permissions in a client's Microsoft 365 tenant.

.DESCRIPTION
  - Signs in with Connect-ExchangeOnline (interactive, Microsoft's own MFA sign-in, run in a background session so the UI stays responsive).
    This app never sees or stores the password.
  - Lists mailboxes, filterable by display name / UPN / type / active (enabled) accounts.
  - Build a queue of Mailbox -> Delegate -> AccessRights from checked mailboxes or a CSV
    (columns: Mailbox, Delegate, AccessRights - same as the original script).
  - Applies with Add-MailboxFolderPermission, or Set-MailboxFolderPermission when the
    delegate already has an entry. Supports -WhatIf and writes a CSV log.

  Requirements: Windows, PowerShell 5.1+ (or 7 on Windows), ExchangeOnlineManagement module,
  and an account with Exchange admin rights (e.g. Exchange Administrator / Global Admin).
#>

#Requires -Version 5.1
Set-StrictMode -Off
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Data

if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) { throw 'This app needs Windows (WPF).' }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    # WPF needs STA; relaunch if needed (PowerShell 7 defaults to STA on Windows, 5.1 console does not always).
    & (Get-Process -Id $PID).Path -STA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    return
}

$AccessRightsList = 'Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor',
                    'Reviewer','Contributor','AvailabilityOnly','LimitedDetails'
$LogPath = Join-Path $PSScriptRoot ("CalendarPerm-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Calendar Permission Manager" Width="1200" Height="780" WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="170"/>
    </Grid.RowDefinitions>

    <!-- Connection -->
    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Admin UPN (optional):" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="AdminUpn" Width="260" Margin="0,0,10,0"/>
      <CheckBox x:Name="DeviceCode" Content="Use device code" VerticalAlignment="Center" Margin="0,0,10,0" ToolTip="Fallback if the browser sign-in does not open"/>
      <Button x:Name="BtnConnect" Content="Connect (MFA sign-in)" Padding="12,4"/>
      <Button x:Name="BtnDisconnect" Content="Disconnect" Padding="12,4" Margin="6,0,0,0" IsEnabled="False"/>
      <TextBlock x:Name="ConnStatus" Text="Not connected" VerticalAlignment="Center" Margin="12,0,0,0" FontWeight="Bold"/>
    </StackPanel>

    <!-- Filters -->
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Filter name/UPN:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="FilterText" Width="220" Margin="0,0,12,0"/>
      <CheckBox x:Name="ActiveOnly" Content="Active users only" IsChecked="True" VerticalAlignment="Center" Margin="0,0,12,0"/>
      <TextBlock Text="Type:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="TypeFilter" Width="140" SelectedIndex="0" Margin="0,0,12,0">
        <ComboBoxItem Content="All"/><ComboBoxItem Content="UserMailbox"/>
        <ComboBoxItem Content="SharedMailbox"/><ComboBoxItem Content="RoomMailbox"/>
        <ComboBoxItem Content="EquipmentMailbox"/>
      </ComboBox>
      <Button x:Name="BtnLoad" Content="Load / refresh mailboxes" Padding="10,4" IsEnabled="False" Margin="0,0,12,0"/>
      <Button x:Name="BtnCheckVisible" Content="Check visible" Padding="8,4" Margin="0,0,4,0"/>
      <Button x:Name="BtnUncheckAll" Content="Uncheck all" Padding="8,4" Margin="0,0,12,0"/>
      <TextBlock x:Name="CountText" VerticalAlignment="Center"/>
    </StackPanel>

    <!-- Grids -->
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions><ColumnDefinition Width="1.2*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <DockPanel Grid.Column="0">
        <TextBlock DockPanel.Dock="Top" Text="1. Mailboxes (check the ones to grant access ON)" FontWeight="Bold" Margin="0,0,0,4"/>
        <DataGrid x:Name="MbxGrid" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column" GridLinesVisibility="Horizontal">
          <DataGrid.Columns>
            <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="30"/>
            <DataGridTextColumn Header="Display name" Binding="{Binding DisplayName}" IsReadOnly="True" Width="*"/>
            <DataGridTextColumn Header="UPN" Binding="{Binding UPN}" IsReadOnly="True" Width="*"/>
            <DataGridTextColumn Header="Type" Binding="{Binding Type}" IsReadOnly="True" Width="110"/>
            <DataGridTextColumn Header="Active" Binding="{Binding Active}" IsReadOnly="True" Width="50"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>

      <DockPanel Grid.Column="2">
        <StackPanel DockPanel.Dock="Top">
          <TextBlock Text="2. Delegate + rights" FontWeight="Bold" Margin="0,0,0,4"/>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
            <TextBlock Text="Delegate:" VerticalAlignment="Center" Margin="0,0,6,4"/>
            <ComboBox x:Name="DelegateBox" Grid.Column="1" IsEditable="True" Margin="0,0,0,4"/>
            <TextBlock Text="Rights:" Grid.Row="1" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="RightsBox" Grid.Row="1" Grid.Column="1" SelectedIndex="6"/>
          </Grid>
          <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
            <Button x:Name="BtnAddQueue" Content="Add checked mailboxes to queue" Padding="10,4" Margin="0,0,6,0"/>
            <Button x:Name="BtnImport" Content="Import CSV" Padding="10,4" Margin="0,0,6,0"/>
            <Button x:Name="BtnClearQueue" Content="Clear queue" Padding="10,4"/>
          </StackPanel>
          <TextBlock Text="3. Queue (Mailbox / Delegate / AccessRights)" FontWeight="Bold" Margin="0,0,0,4"/>
        </StackPanel>
        <DataGrid x:Name="QueueGrid" AutoGenerateColumns="False" CanUserAddRows="False" HeadersVisibility="Column" IsReadOnly="True">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="*"/>
            <DataGridTextColumn Header="Delegate" Binding="{Binding Delegate}" Width="*"/>
            <DataGridTextColumn Header="Rights" Binding="{Binding AccessRights}" Width="110"/>
            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="*"/>
          </DataGrid.Columns>
        </DataGrid>
      </DockPanel>
    </Grid>

    <!-- Apply -->
    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,8">
      <CheckBox x:Name="WhatIfBox" Content="Preview only (-WhatIf)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,16,0"/>
      <Button x:Name="BtnApply" Content="Apply queue" Padding="16,6" FontWeight="Bold" IsEnabled="False"/>
      <TextBlock x:Name="ApplyStatus" VerticalAlignment="Center" Margin="12,0,0,0"/>
    </StackPanel>

    <TextBox x:Name="LogBox" Grid.Row="4" IsReadOnly="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" TextWrapping="Wrap"/>
  </Grid>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $xaml))
$ui = @{}
foreach ($n in 'DeviceCode','AdminUpn','BtnConnect','BtnDisconnect','ConnStatus','FilterText','ActiveOnly','TypeFilter','BtnLoad',
               'BtnCheckVisible','BtnUncheckAll','CountText','MbxGrid','DelegateBox','RightsBox','BtnAddQueue',
               'BtnImport','BtnClearQueue','QueueGrid','WhatIfBox','BtnApply','ApplyStatus','LogBox') {
    $ui[$n] = $win.FindName($n)
}
$ui.RightsBox.ItemsSource = $AccessRightsList
$ui.RightsBox.SelectedItem = 'Reviewer'

# Data tables ---------------------------------------------------------------
$mbxTable = New-Object Data.DataTable
[void]$mbxTable.Columns.Add('Selected',[bool]);
foreach ($c in 'DisplayName','UPN','Type','Active') { [void]$mbxTable.Columns.Add($c,[string]) }
$mbxTable.Columns['Selected'].DefaultValue = $false
$ui.MbxGrid.ItemsSource = $mbxTable.DefaultView

$queueTable = New-Object Data.DataTable
foreach ($c in 'Mailbox','Delegate','AccessRights','Status') { [void]$queueTable.Columns.Add($c,[string]) }
$ui.QueueGrid.ItemsSource = $queueTable.DefaultView

# Helpers -------------------------------------------------------------------
function DoEvents {
    $f = New-Object Windows.Threading.DispatcherFrame
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke('Background',[Action]{ $f.Continue = $false })
    [Windows.Threading.Dispatcher]::PushFrame($f)
}
function Log([string]$m) {
    $ui.LogBox.AppendText(("[{0:HH:mm:ss}] {1}`r`n" -f (Get-Date), $m)); $ui.LogBox.ScrollToEnd(); DoEvents
}
function Escape-Like([string]$s) { ($s -replace "'","''") -replace '([\[\]%*])','[$1]' }

function Update-Filter {
    $parts = @()
    $t = $ui.FilterText.Text.Trim()
    if ($t) { $e = Escape-Like $t; $parts += "(DisplayName LIKE '%$e%' OR UPN LIKE '%$e%')" }
    if ($ui.ActiveOnly.IsChecked) { $parts += "Active <> 'No'" }
    $type = $ui.TypeFilter.SelectedItem.Content
    if ($type -ne 'All') { $parts += "Type = '$type'" }
    $mbxTable.DefaultView.RowFilter = ($parts -join ' AND ')
    $checked = @($mbxTable.Select('Selected = true')).Count
    $ui.CountText.Text = "{0} shown / {1} loaded / {2} checked" -f $mbxTable.DefaultView.Count, $mbxTable.Rows.Count, $checked
}

# All Exchange work runs in a separate MTA runspace so the UI thread never blocks
# (interactive sign-in on the WPF thread freezes the window).
$script:exoRs = [runspacefactory]::CreateRunspace()
$script:exoRs.ApartmentState = 'MTA'
$script:exoRs.Open()
function Invoke-Exo([scriptblock]$Script, [object[]]$ArgList = @()) {
    $ps = [powershell]::Create()
    $ps.Runspace = $script:exoRs
    [void]$ps.AddScript('$ErrorActionPreference = "Stop"; ' + $Script.ToString())
    foreach ($a in $ArgList) { [void]$ps.AddArgument($a) }
    $h = $ps.BeginInvoke(); $seen = 0
    try {
        while (-not $h.IsCompleted) {
            DoEvents; Start-Sleep -Milliseconds 40
            # surface Write-Host output (e.g. device-code instructions)
            while ($seen -lt $ps.Streams.Information.Count) { Log ([string]$ps.Streams.Information[$seen].MessageData); $seen++ }
        }
        while ($seen -lt $ps.Streams.Information.Count) { Log ([string]$ps.Streams.Information[$seen].MessageData); $seen++ }
        $ps.EndInvoke($h)
    } finally { $ps.Dispose() }
}

function Add-ToQueue($mailbox,$delegate,$rights) {
    $r = $queueTable.NewRow(); $r.Mailbox=$mailbox; $r.Delegate=$delegate; $r.AccessRights=$rights; $r.Status='Pending'
    $queueTable.Rows.Add($r)
}

# Events --------------------------------------------------------------------
$ui.FilterText.Add_TextChanged({ Update-Filter })
$ui.ActiveOnly.Add_Click({ Update-Filter })
$ui.TypeFilter.Add_SelectionChanged({ if ($mbxTable) { Update-Filter } })
$ui.MbxGrid.Add_CellEditEnding({ $win.Dispatcher.BeginInvoke('Background',[Action]{ Update-Filter }) | Out-Null })
$ui.BtnCheckVisible.Add_Click({ foreach ($v in $mbxTable.DefaultView) { $v.Row['Selected'] = $true }; Update-Filter })
$ui.BtnUncheckAll.Add_Click({ foreach ($r in $mbxTable.Rows) { $r['Selected'] = $false }; Update-Filter })

$ui.BtnConnect.Add_Click({
    try {
        $ui.BtnConnect.IsEnabled = $false
        $ui.ConnStatus.Text = 'Signing in... complete the Microsoft sign-in (browser window)'
        Log 'Checking ExchangeOnlineManagement module...'
        $have = Invoke-Exo { [bool](Get-Module -ListAvailable ExchangeOnlineManagement) }
        if (-not $have) {
            $a = [Windows.MessageBox]::Show('ExchangeOnlineManagement module not found. Install it for the current user now?','Missing module','YesNo')
            if ($a -ne 'Yes') { $ui.ConnStatus.Text = 'Not connected'; return }
            Log 'Installing ExchangeOnlineManagement...'
            Invoke-Exo { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber } | Out-Null
        }
        Log 'Opening sign-in...'
        $info = Invoke-Exo {
            param($upn, $device)
            Import-Module ExchangeOnlineManagement
            $p = @{ ShowBanner = $false }
            if ($upn) { $p.UserPrincipalName = $upn }
            if ($device) { $p.Device = $true }
            elseif ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $p.DisableWAM = $true }
            Connect-ExchangeOnline @p
            $c = Get-ConnectionInformation | Select-Object -First 1
            "$($c.UserPrincipalName)|$($c.TenantID)"
        } @($ui.AdminUpn.Text.Trim(), [bool]$ui.DeviceCode.IsChecked)
        $u,$t = ([string]($info | Select-Object -Last 1)) -split '\|'
        $ui.ConnStatus.Text = "Connected: $u @ $t"
        $ui.BtnLoad.IsEnabled = $true; $ui.BtnDisconnect.IsEnabled = $true; $ui.BtnApply.IsEnabled = $true
        Log "Connected as $u (tenant $t)"
    } catch { $ui.ConnStatus.Text = 'Not connected'; Log "Connect failed: $($_.Exception.Message)" }
    finally { $ui.BtnConnect.IsEnabled = $true }
})

$ui.BtnDisconnect.Add_Click({
    try { Invoke-Exo { Disconnect-ExchangeOnline -Confirm:$false } | Out-Null } catch {}
    $ui.ConnStatus.Text = 'Not connected'
    $ui.BtnLoad.IsEnabled = $false; $ui.BtnDisconnect.IsEnabled = $false; $ui.BtnApply.IsEnabled = $false
    Log 'Disconnected.'
})

$ui.BtnLoad.Add_Click({
    try {
        $ui.BtnLoad.IsEnabled = $false; Log 'Loading mailboxes (large tenants can take a minute)...'
        $types = 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox'
        # Get-User carries AccountDisabled, so one call gives name + active state.
        $users = Invoke-Exo {
            param($types)
            Get-User -ResultSize Unlimited -RecipientTypeDetails $types | ForEach-Object {
                [pscustomobject]@{ DisplayName = [string]$_.DisplayName; UserPrincipalName = [string]$_.UserPrincipalName
                                   RecipientTypeDetails = [string]$_.RecipientTypeDetails; AccountDisabled = [bool]$_.AccountDisabled }
            }
        } @(,$types)
        $mbxTable.Clear()
        foreach ($u in $users | Sort-Object DisplayName) {
            $r = $mbxTable.NewRow()
            $r.Selected = $false; $r.DisplayName = $u.DisplayName; $r.UPN = $u.UserPrincipalName
            $r.Type = [string]$u.RecipientTypeDetails
            # Shared/room/equipment accounts are disabled by design, so "active" only applies to user mailboxes.
            $r.Active = if ($u.RecipientTypeDetails -eq 'UserMailbox') { if ($u.AccountDisabled) {'No'} else {'Yes'} } else { 'n/a' }
            $mbxTable.Rows.Add($r)
        }
        $ui.DelegateBox.ItemsSource = @($users | Sort-Object UserPrincipalName | ForEach-Object UserPrincipalName)
        Update-Filter; Log "Loaded $($mbxTable.Rows.Count) mailboxes."
    } catch { Log "Load failed: $($_.Exception.Message)" }
    finally { $ui.BtnLoad.IsEnabled = $true }
})

$ui.BtnAddQueue.Add_Click({
    $d = $ui.DelegateBox.Text.Trim(); $rights = [string]$ui.RightsBox.SelectedItem
    if (-not $d) { [void][Windows.MessageBox]::Show('Enter or pick a delegate first.'); return }
    $rows = @($mbxTable.Select('Selected = true'))
    if (-not $rows) { [void][Windows.MessageBox]::Show('Check at least one mailbox.'); return }
    foreach ($r in $rows) { Add-ToQueue $r['UPN'] $d $rights }
    Log "Queued $($rows.Count) entr$(if($rows.Count -eq 1){'y'}else{'ies'}) for $d ($rights)."
})

$ui.BtnImport.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = 'CSV (*.csv)|*.csv'
    if (-not $dlg.ShowDialog()) { return }
    try {
        $csv = Import-Csv $dlg.FileName; $bad = 0
        foreach ($row in $csv) {
            if (-not ($row.Mailbox -and $row.Delegate -and $row.AccessRights) -or $row.AccessRights -notin $AccessRightsList) {
                $bad++; Log "Skipped invalid row: $($row | ConvertTo-Json -Compress)"; continue }
            Add-ToQueue $row.Mailbox.Trim() $row.Delegate.Trim() $row.AccessRights.Trim()
        }
        Log "Imported $($csv.Count - $bad) rows ($bad skipped) from $($dlg.FileName)."
    } catch { Log "Import failed: $($_.Exception.Message)" }
})

$ui.BtnClearQueue.Add_Click({ $queueTable.Clear() })

$ui.BtnApply.Add_Click({
    $whatIf = [bool]$ui.WhatIfBox.IsChecked
    $n = $queueTable.Rows.Count
    if ($n -eq 0) { [void][Windows.MessageBox]::Show('Queue is empty.'); return }
    if (-not $whatIf) {
        $a = [Windows.MessageBox]::Show("Apply $n calendar permission change(s) to the live tenant?",'Confirm','YesNo','Warning')
        if ($a -ne 'Yes') { return }
    }
    $ui.BtnApply.IsEnabled = $false; $ok = 0; $fail = 0; $i = 0; $results = @()
    foreach ($row in $queueTable.Rows) {
        $i++; $ui.ApplyStatus.Text = "Processing $i / $n"; DoEvents
        try {
            $verb = Invoke-Exo {
                param($mbx, $delegate, $rights, $whatIf)
                $f = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
                     Where-Object FolderType -eq 'Calendar' | Select-Object -First 1
                if (-not $f) { throw "No calendar folder found for $mbx" }
                $id = "{0}:{1}" -f $mbx, ($f.FolderPath -replace '/','\')
                $existing = $null
                try { $existing = Get-MailboxFolderPermission -Identity $id -User $delegate } catch {}
                $p = @{ Identity = $id; User = $delegate; AccessRights = $rights; Confirm = $false }
                if ($whatIf) { $p.WhatIf = $true }
                if ($existing) { Set-MailboxFolderPermission @p | Out-Null; 'Updated' }
                else           { Add-MailboxFolderPermission @p | Out-Null; 'Added' }
            } @($row.Mailbox, $row.Delegate, $row.AccessRights, $whatIf)
            $verb = [string]($verb | Select-Object -Last 1)
            $row.Status = if ($whatIf) { "WhatIf: would have $($verb.ToLower())" } else { $verb }
            $ok++
        } catch { $row.Status = 'FAILED: ' + $_.Exception.Message; $fail++ }
        Log "$($row.Mailbox) <- $($row.Delegate) [$($row.AccessRights)]: $($row.Status)"
        $results += [pscustomobject]@{ Time=Get-Date -Format s; Mailbox=$row.Mailbox; Delegate=$row.Delegate
                                       AccessRights=$row.AccessRights; WhatIf=$whatIf; Result=$row.Status }
    }
    $results | Export-Csv -Path $LogPath -NoTypeInformation -Append
    $ui.ApplyStatus.Text = "Done: $ok ok, $fail failed. Log: $LogPath"
    Log $ui.ApplyStatus.Text; $ui.BtnApply.IsEnabled = $true
})

$win.Add_Closing({ try { Invoke-Exo { if (Get-ConnectionInformation) { Disconnect-ExchangeOnline -Confirm:$false } } | Out-Null } catch {}; try { $script:exoRs.Dispose() } catch {} })
Update-Filter
[void]$win.ShowDialog()
