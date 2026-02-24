#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------
# Paths
# ----------------------------
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$XamlPath   = Join-Path $ScriptRoot 'UI\ChangeChecklist.xaml'
if (-not (Test-Path $XamlPath)) { throw "XAML not found: $XamlPath" }

# ----------------------------
# Storage layout (maintainable)
# ----------------------------
$AppRoot        = Join-Path $env:APPDATA 'ChangeChecklist'
$DefinitionsDir = Join-Path $AppRoot 'definitions'
$StateDir       = Join-Path $AppRoot 'state'
$null = New-Item -ItemType Directory -Path $DefinitionsDir -Force
$null = New-Item -ItemType Directory -Path $StateDir -Force

$DefaultDefinitionPath = Join-Path $DefinitionsDir 'default.definition.json'

# Change Request URL template: replace with your real system.
# Use "{id}" placeholder.
$ChangeUrlTemplate = "https://example.service-now.com/nav_to.do?uri=change_request.do?sysparm_query=number={id}"

function New-DefaultDefinition {
    @{
        title    = "Change Checklist"
        sections = @(
            @{
                name  = "Pre Change Request"
                items = @(
                    @{ id="pre-001"; text="Confirm approval / authorization recorded"; linkText=""; linkUrl="" },
                    @{ id="pre-002"; text="Notify impacted users / distribution list";  linkText="Template"; linkUrl="https://example.com/notify-template" },
                    @{ id="pre-003"; text="Rollback plan documented and verified";     linkText=""; linkUrl="" }
                )
            },
            @{
                name  = "Begin Implementation on the Change Request"
                items = @(
                    @{ id="impl-001"; text="Start maintenance window / confirm monitoring in place"; linkText=""; linkUrl="" },
                    @{ id="impl-002"; text="Execute implementation steps per runbook";             linkText="Runbook"; linkUrl="https://example.com/runbook" },
                    @{ id="impl-003"; text="Post-change validation / smoke tests complete";        linkText=""; linkUrl="" }
                )
            }
        )
    }
}

function Ensure-DefaultDefinition {
    if (-not (Test-Path $DefaultDefinitionPath)) {
        (New-DefaultDefinition) | ConvertTo-Json -Depth 12 | Set-Content -Path $DefaultDefinitionPath -Encoding UTF8
    }
}

function Load-Definition {
    Ensure-DefaultDefinition
    Get-Content -Path $DefaultDefinitionPath -Raw | ConvertFrom-Json
}

function Get-StatePath([string]$ChangeId) {
    $safe = ($ChangeId -replace '[^\w\-]', '_')
    Join-Path $StateDir "$safe.state.json"
}

function New-BlankState([string]$ChangeId) {
    [pscustomobject]@{
        changeId   = $ChangeId
        savedUtc   = $null
        itemStates = @{}
        window     = @{ top=$null; left=$null; width=$null; height=$null }
    }
}

function Load-State([string]$ChangeId) {
    $path = Get-StatePath $ChangeId
    if (Test-Path $path) { return (Get-Content -Path $path -Raw | ConvertFrom-Json) }
    New-BlankState -ChangeId $ChangeId
}

function Save-State([object]$state) {
    $state.savedUtc = ([DateTime]::UtcNow.ToString("o"))
    $path = Get-StatePath $state.changeId
    $state | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding UTF8
}

function Open-ChangeLink([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) { return }
    $url = $ChangeUrlTemplate.Replace('{id}', [uri]::EscapeDataString($ChangeId.Trim()))
    try { Start-Process $url | Out-Null } catch {}
}

# ----------------------------
# ViewModel item
# ----------------------------
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

    public string CheckedUtc { get; set; }

    public bool HasLink
    {
        get { return !String.IsNullOrWhiteSpace(LinkUrl) && !String.IsNullOrWhiteSpace(LinkText); }
    }

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

# ----------------------------
# Load XAML from file
# ----------------------------
[xml]$xaml = Get-Content -Path $XamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Named controls
$tabs         = $window.FindName('Tabs')
$txtTitle     = $window.FindName('TxtTitle')
$txtChangeId  = $window.FindName('TxtChangeId')
$txtStatus    = $window.FindName('TxtStatus')
$btnLoad      = $window.FindName('BtnLoad')
$btnSave      = $window.FindName('BtnSave')
$btnOpenCR    = $window.FindName('BtnOpenCR')
$btnEditDef   = $window.FindName('BtnEditDefinition')
$btnOpenState = $window.FindName('BtnOpenState')

function Set-Status([string]$msg) { $txtStatus.Text = $msg }

$definition      = Load-Definition
$currentState    = $null
$allSections     = @()

function Ensure-ItemState([object]$state, [string]$id) {
    if (-not $state.itemStates.$id) {
        $state.itemStates | Add-Member -MemberType NoteProperty -Name $id -Value ([pscustomobject]@{
            isChecked  = $false
            checkedUtc = $null
            notes      = ""
        })
    }
}

function Clear-Tabs {
    $tabs.Items.Clear()
    $allSections = @()
}

function Update-SectionProgress {
    foreach ($sec in $allSections) {
        $total = $sec.Items.Count
        $done  = ($sec.Items | Where-Object IsChecked).Count
        $sec.Progress.Text = "Progress: $done / $total completed"
    }
}

function Build-Tabs([object]$state) {
    Clear-Tabs

    foreach ($section in $definition.sections) {
        $items = New-Object 'System.Collections.ObjectModel.ObservableCollection[ChecklistItem]'

        foreach ($it in $section.items) {
            Ensure-ItemState $state $it.id
            $st = $state.itemStates.$($it.id)

            $ci = New-Object ChecklistItem
            $ci.Id         = $it.id
            $ci.Text       = $it.text
            $ci.LinkText   = $it.linkText
            $ci.LinkUrl    = $it.linkUrl
            $ci.IsChecked  = [bool]$st.isChecked
            $ci.CheckedUtc = $st.checkedUtc
            $ci.Notes      = [string]$st.notes

            $ci.add_PropertyChanged({
                param($sender, $args)
                if (-not $currentState) { return }

                Ensure-ItemState $currentState $sender.Id
                $entry = $currentState.itemStates.$($sender.Id)

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

                Save-State $currentState
                Set-Status "Saved: $($currentState.changeId)"
                Update-SectionProgress
            })

            $items.Add($ci)
        }

        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = $section.name

        $sv = New-Object System.Windows.Controls.ScrollViewer
        $sv.VerticalScrollBarVisibility = 'Auto'
        $sv.HorizontalScrollBarVisibility = 'Disabled'

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = '0,12,0,0'

        $progress = New-Object System.Windows.Controls.TextBlock
        $progress.Margin = '0,0,0,12'
        $progress.Foreground = [Windows.Media.Brushes]::LightSteelBlue
        $progress.FontSize = 12
        $panel.Children.Add($progress) | Out-Null

        $ic = New-Object System.Windows.Controls.ItemsControl
        $ic.ItemsSource = $items
        $ic.ItemTemplate = $window.Resources['ChecklistItemTemplate']
        $panel.Children.Add($ic) | Out-Null

        $sv.Content = $panel
        $tab.Content = $sv
        $tabs.Items.Add($tab) | Out-Null

        $allSections += [pscustomobject]@{ Name=$section.name; Items=$items; Progress=$progress }
    }

    Update-SectionProgress
}

# Hyperlink handler
$window.AddHandler(
    [System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
    [System.Windows.Navigation.RequestNavigateEventHandler]{
        param($s, $e)
        try { Start-Process $e.Uri.AbsoluteUri | Out-Null } catch {}
        $e.Handled = $true
    }
)

function Load-Change([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) { Set-Status "Enter a ChangeID first."; return }
    $id = $ChangeId.Trim()

    $currentState = Load-State $id
    $txtTitle.Text = "$($definition.title) — $id"

    Build-Tabs $currentState

    # Restore window geometry per ChangeID
    if ($currentState.window.width -and $currentState.window.height) {
        $window.Width  = [double]$currentState.window.width
        $window.Height = [double]$currentState.window.height
    }
    if ($currentState.window.left -ne $null -and $currentState.window.top -ne $null) {
        $window.Left = [double]$currentState.window.left
        $window.Top  = [double]$currentState.window.top
    }

    Save-State $currentState
    Set-Status "Loaded: $id"
}

function Save-Current {
    if (-not $currentState) { Set-Status "Nothing loaded yet."; return }
    Save-State $currentState
    Set-Status "Saved: $($currentState.changeId)"
}

$btnLoad.Add_Click({ Load-Change $txtChangeId.Text })
$btnSave.Add_Click({ Save-Current })
$btnOpenCR.Add_Click({ Open-ChangeLink $txtChangeId.Text })

$btnEditDef.Add_Click({
    Start-Process notepad.exe $DefaultDefinitionPath | Out-Null
    [System.Windows.MessageBox]::Show(
        "Edit the template JSON and save it. Then click Load again to refresh.",
        "Edit Template"
    ) | Out-Null
    try { $definition = Load-Definition; Set-Status "Template reloaded. Click Load." } catch { Set-Status "Template edited. Click Load." }
})

$btnOpenState.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtChangeId.Text)) {
        [System.Windows.MessageBox]::Show("Enter a ChangeID first, then click Open Saved State.","Open Saved State") | Out-Null
        return
    }
    $path = Get-StatePath $txtChangeId.Text
    if (-not (Test-Path $path)) {
        [System.Windows.MessageBox]::Show("No saved state yet for this ChangeID. Click Load first.","Open Saved State") | Out-Null
        return
    }
    Start-Process notepad.exe $path | Out-Null
})

$window.add_Closing({
    if ($currentState) {
        $currentState.window.left   = $window.Left
        $currentState.window.top    = $window.Top
        $currentState.window.width  = $window.Width
        $currentState.window.height = $window.Height
        Save-State $currentState
    }
})

# Default starter
$txtChangeId.Text = "CHG0000000"
Set-Status "Enter ChangeID and click Load."

$null = $window.ShowDialog()
