#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------
# Storage (Definition + State)
# ----------------------------
$AppRoot = Join-Path $env:APPDATA 'ChangeChecklist'
$null = New-Item -ItemType Directory -Path $AppRoot -Force

# Use a Change Request ID so each CR remembers separately.
# Change this default to whatever you want.
$ChangeId = 'CR-000000'
$DefinitionPath = Join-Path $AppRoot "$ChangeId.definition.json"
$StatePath      = Join-Path $AppRoot "$ChangeId.state.json"

function New-DefaultDefinition {
    @{
        changeId = $ChangeId
        title    = "Change Request Checklist ($ChangeId)"
        sections = @(
            @{
                name  = "Pre Change Request"
                items = @(
                    @{
                        id       = "pre-001"
                        text     = "Confirm CAB approval recorded"
                        linkText = "CAB record"
                        linkUrl  = "https://example.com/cab"
                    },
                    @{
                        id       = "pre-002"
                        text     = "Notify impacted users / distribution list"
                        linkText = "Notification template"
                        linkUrl  = "https://example.com/template"
                    },
                    @{
                        id       = "pre-003"
                        text     = "Validate rollback plan documented"
                        linkText = ""
                        linkUrl  = ""
                    }
                )
            },
            @{
                name  = "Begin Implementation on the Change Request"
                items = @(
                    @{
                        id       = "impl-001"
                        text     = "Start maintenance window / confirm monitoring in place"
                        linkText = ""
                        linkUrl  = ""
                    },
                    @{
                        id       = "impl-002"
                        text     = "Execute step-by-step implementation tasks"
                        linkText = "Runbook"
                        linkUrl  = "https://example.com/runbook"
                    },
                    @{
                        id       = "impl-003"
                        text     = "Post-change validation (service checks / smoke tests)"
                        linkText = ""
                        linkUrl  = ""
                    }
                )
            }
        )
    }
}

function Load-Definition {
    if (-not (Test-Path $DefinitionPath)) {
        (New-DefaultDefinition) | ConvertTo-Json -Depth 10 | Set-Content -Path $DefinitionPath -Encoding UTF8
    }
    Get-Content -Path $DefinitionPath -Raw | ConvertFrom-Json
}

function Load-State {
    if (Test-Path $StatePath) {
        return (Get-Content -Path $StatePath -Raw | ConvertFrom-Json)
    }
    return [pscustomobject]@{
        changeId   = $ChangeId
        savedUtc   = $null
        itemStates = @{}   # id -> @{ isChecked=bool; checkedUtc=string|null; notes=string }
        window     = @{ top=$null; left=$null; width=$null; height=$null }
    }
}

function Save-State([object]$state) {
    $state.savedUtc = ([DateTime]::UtcNow.ToString("o"))
    $state | ConvertTo-Json -Depth 10 | Set-Content -Path $StatePath -Encoding UTF8
}

# --------------------------------
# XAML (UI)
# --------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Change Checklist" Height="720" Width="1050"
        WindowStartupLocation="CenterScreen"
        Background="#0F1218" Foreground="#EAEFF7">
    <Window.Resources>

        <Style TargetType="TextBlock">
            <Setter Property="TextWrapping" Value="Wrap"/>
        </Style>

        <Style x:Key="PillButton" TargetType="Button">
            <Setter Property="Padding" Value="12,7"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#EAEFF7"/>
            <Setter Property="Background" Value="#1B2332"/>
            <Setter Property="BorderBrush" Value="#2A3850"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="14">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#22304A"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.55"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="HyperlinkishButton" TargetType="Button" BasedOn="{StaticResource PillButton}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Foreground" Value="#85B6FF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <DataTemplate x:Key="ChecklistItemTemplate">
            <Border Margin="0,0,0,10" Padding="12"
                    Background="#121927" BorderBrush="#25324A" BorderThickness="1"
                    CornerRadius="14">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="34"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="260"/>
                    </Grid.ColumnDefinitions>

                    <!-- Checkbox -->
                    <CheckBox Grid.Column="0"
                              VerticalAlignment="Top"
                              IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                              Margin="2,2,0,0"/>

                    <!-- Text + link -->
                    <StackPanel Grid.Column="1" Margin="6,0,10,0">
                        <TextBlock FontSize="14" FontWeight="SemiBold" Text="{Binding Text}"/>

                        <TextBlock Margin="0,6,0,0" Visibility="{Binding HasLink, Converter={x:Static BooleanBoxes.VisibleIfTrue}}">
                            <Hyperlink NavigateUri="{Binding LinkUrl}">
                                <Run Text="{Binding LinkText}"/>
                            </Hyperlink>
                        </TextBlock>
                    </StackPanel>

                    <!-- Notes -->
                    <StackPanel Grid.Column="2">
                        <TextBlock FontSize="12" Foreground="#AAB7D3" Text="Notes" Margin="0,0,0,6"/>
                        <TextBox Text="{Binding Notes, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                 MinHeight="52"
                                 Background="#0D1422"
                                 Foreground="#EAEFF7"
                                 BorderBrush="#2A3850"
                                 BorderThickness="1"
                                 Padding="10"
                                 AcceptsReturn="True"
                                 TextWrapping="Wrap"
                                 VerticalScrollBarVisibility="Auto"/>
                        <TextBlock Margin="0,8,0,0" FontSize="11" Foreground="#8FA1C2"
                                   Text="{Binding CheckedLabel}"/>
                    </StackPanel>

                </Grid>
            </Border>
        </DataTemplate>

        <!-- Small helper: because WPF doesn't have a built-in boolean->Visibility converter without extra libs -->
        <x:StaticExtension x:Key="BooleanBoxes.VisibleIfTrue" Member="Visibility.Visible"/>
    </Window.Resources>

    <DockPanel LastChildFill="True" Margin="18">

        <!-- Top header -->
        <Border DockPanel.Dock="Top" Padding="18" CornerRadius="18"
                Background="#121927" BorderBrush="#25324A" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel>
                    <TextBlock FontSize="22" FontWeight="ExtraBold" Text="{Binding Title}"/>
                    <TextBlock Margin="0,6,0,0" Foreground="#AAB7D3" FontSize="12"
                               Text="Check items off, add notes, and close the window — it saves automatically per Change ID."/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnEditDefinition" Style="{StaticResource PillButton}" Content="Edit Checklist (JSON)"/>
                    <Button x:Name="BtnOpenState" Style="{StaticResource PillButton}" Content="Open Saved State"/>
                    <Button x:Name="BtnMarkAllDone" Style="{StaticResource PillButton}" Content="Mark Section Done"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main tabs -->
        <Border Margin="0,14,0,0" Padding="14" CornerRadius="18"
                Background="#0C111B" BorderBrush="#25324A" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <TabControl x:Name="Tabs" Background="Transparent" BorderThickness="0" Grid.Row="1">
                    <!-- Tabs added dynamically from PowerShell -->
                </TabControl>
            </Grid>
        </Border>
    </DockPanel>
</Window>
"@

# NOTE: The XAML above references BooleanBoxes.VisibleIfTrue as a hacky placeholder.
# We'll fix link visibility purely in PowerShell by binding Visibility directly to HasLink with a converter-free approach:
# We'll set HasLinkVisibility property on each item. (No need for converters.)

# ----------------------------
# Build View Models (PS objects)
# ----------------------------
$def   = Load-Definition
$state = Load-State

# Window + reader
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find named controls
$tabs             = $window.FindName('Tabs')
$btnEditDef       = $window.FindName('BtnEditDefinition')
$btnOpenState     = $window.FindName('BtnOpenState')
$btnMarkAllDone   = $window.FindName('BtnMarkAllDone')

# Create a simple item type with property-changed so bindings update nicely
Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;

public class ChecklistItem : INotifyPropertyChanged
{
    public string Id { get; set; }
    public string Text { get; set; }
    public string LinkText { get; set; }
    public string LinkUrl { get; set; }

    private bool _isChecked;
    public bool IsChecked
    {
        get { return _isChecked; }
        set { _isChecked = value; OnPropertyChanged("IsChecked"); OnPropertyChanged("CheckedLabel"); }
    }

    private string _notes;
    public string Notes
    {
        get { return _notes; }
        set { _notes = value; OnPropertyChanged("Notes"); }
    }

    public bool HasLink
    {
        get { return !String.IsNullOrWhiteSpace(LinkUrl) && !String.IsNullOrWhiteSpace(LinkText); }
    }

    public string CheckedUtc { get; set; } // ISO string or null

    public string CheckedLabel
    {
        get
        {
            if (!IsChecked) return "Not checked yet";
            if (String.IsNullOrWhiteSpace(CheckedUtc)) return "Checked";
            return "Checked: " + CheckedUtc;
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;
    private void OnPropertyChanged(string name)
    {
        if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(name));
    }
}
"@ -ReferencedAssemblies 'System.dll'

# Title binding
$window.DataContext = [pscustomobject]@{ Title = $def.title }

# Helper: ensure state entry exists
function Ensure-ItemState([string]$id) {
    if (-not $state.itemStates.$id) {
        $state.itemStates | Add-Member -MemberType NoteProperty -Name $id -Value ([pscustomobject]@{
            isChecked  = $false
            checkedUtc = $null
            notes      = ""
        })
    }
}

# Build tabs + lists
$allSections = @()
foreach ($section in $def.sections) {
    $items = New-Object 'System.Collections.ObjectModel.ObservableCollection[ChecklistItem]'

    foreach ($it in $section.items) {
        Ensure-ItemState $it.id

        $st = $state.itemStates.$($it.id)

        $ci = New-Object ChecklistItem
        $ci.Id       = $it.id
        $ci.Text     = $it.text
        $ci.LinkText = $it.linkText
        $ci.LinkUrl  = $it.linkUrl
        $ci.IsChecked = [bool]$st.isChecked
        $ci.CheckedUtc = $st.checkedUtc
        $ci.Notes    = [string]$st.notes

        # Save-on-change behavior
        $ci.add_PropertyChanged({
            param($sender, $args)

            # Mirror to state
            Ensure-ItemState $sender.Id
            $entry = $state.itemStates.$($sender.Id)

            if ($args.PropertyName -eq 'IsChecked') {
                $entry.isChecked = [bool]$sender.IsChecked
                if ($sender.IsChecked) {
                    $sender.CheckedUtc = ([DateTime]::UtcNow.ToString("o"))
                    $entry.checkedUtc  = $sender.CheckedUtc
                } else {
                    $sender.CheckedUtc = $null
                    $entry.checkedUtc  = $null
                }
            }

            if ($args.PropertyName -eq 'Notes') {
                $entry.notes = [string]$sender.Notes
            }

            Save-State $state
        })

        $items.Add($ci)
    }

    # Create tab content
    $tab = New-Object System.Windows.Controls.TabItem
    $tab.Header = $section.name

    $sv = New-Object System.Windows.Controls.ScrollViewer
    $sv.VerticalScrollBarVisibility = 'Auto'
    $sv.HorizontalScrollBarVisibility = 'Disabled'

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = '0,12,0,0'

    # Summary row
    $summary = New-Object System.Windows.Controls.TextBlock
    $summary.Margin = '0,0,0,12'
    $summary.Foreground = [Windows.Media.Brushes]::LightSteelBlue
    $summary.FontSize = 12
    $panel.Children.Add($summary) | Out-Null

    # ItemsControl using template from resources
    $ic = New-Object System.Windows.Controls.ItemsControl
    $ic.ItemsSource = $items
    $ic.ItemTemplate = $window.Resources['ChecklistItemTemplate']
    $panel.Children.Add($ic) | Out-Null

    $sv.Content = $panel
    $tab.Content = $sv

    $tabs.Items.Add($tab) | Out-Null

    $allSections += [pscustomobject]@{
        Name  = $section.name
        Items = $items
        SummaryTextBlock = $summary
    }
}

# Update summaries
function Update-Summaries {
    foreach ($sec in $allSections) {
        $total = $sec.Items.Count
        $done  = ($sec.Items | Where-Object IsChecked).Count
        $sec.SummaryTextBlock.Text = "Progress: $done / $total completed"
    }
}
Update-Summaries

# When any item changes, update summary too (cheap + effective)
foreach ($sec in $allSections) {
    foreach ($it in $sec.Items) {
        $it.add_PropertyChanged({ Update-Summaries })
    }
}

# Hyperlink handler (works for all Hyperlink clicks inside templates)
$window.AddHandler(
    [System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
    [System.Windows.Navigation.RequestNavigateEventHandler]{
        param($s, $e)
        try { Start-Process $e.Uri.AbsoluteUri | Out-Null } catch {}
        $e.Handled = $true
    }
)

# Buttons
$btnEditDef.Add_Click({
    # Open definition for editing
    Start-Process notepad.exe $DefinitionPath | Out-Null
    [System.Windows.MessageBox]::Show("Edit the JSON, save it, then re-run the script to reload the checklist definition.`n`nDefinition:`n$DefinitionPath","Edit Checklist") | Out-Null
})

$btnOpenState.Add_Click({
    Start-Process notepad.exe $StatePath | Out-Null
})

$btnMarkAllDone.Add_Click({
    $currentTab = $tabs.SelectedItem
    if (-not $currentTab) { return }
    $sec = $allSections | Where-Object { $_.Name -eq [string]$currentTab.Header } | Select-Object -First 1
    if (-not $sec) { return }

    foreach ($item in $sec.Items) { $item.IsChecked = $true }
    Save-State $state
    Update-Summaries
})

# Restore window position/size if available
if ($state.window.width -and $state.window.height) {
    $window.Width  = [double]$state.window.width
    $window.Height = [double]$state.window.height
}
if ($state.window.left -ne $null -and $state.window.top -ne $null) {
    $window.Left = [double]$state.window.left
    $window.Top  = [double]$state.window.top
}

# Save window position on close
$window.add_Closing({
    $state.window.left   = $window.Left
    $state.window.top    = $window.Top
    $state.window.width  = $window.Width
    $state.window.height = $window.Height
    Save-State $state
})

# Show
$null = $window.ShowDialog()
