#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# CONFIG (keep these variable names consistent)
# ============================================================
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# XAML is separate
$XamlPath = Join-Path $ScriptRoot 'UI\ChangeChecklist.xaml'

# Definition JSON is separate
$DefinitionPath = Join-Path $ScriptRoot 'data\default.definition.json'

# State is per ChangeID in AppData
$StateDir = Join-Path $env:APPDATA 'ChangeChecklist\state'
$null = New-Item -ItemType Directory -Path $StateDir -Force

# Change Request link template (ChangeID replaces {id})
$ChangeUrlTemplate = "https://example.service-now.com/nav_to.do?uri=change_request.do?sysparm_query=number={id}"

# ============================================================
# Safety checks
# ============================================================
if (-not (Test-Path $XamlPath)) {
    throw "XAML not found: $XamlPath`nExpected: UI\ChangeChecklist.xaml (relative to script)"
}
if (-not (Test-Path $DefinitionPath)) {
    throw "Definition JSON not found: $DefinitionPath`nExpected: data\default.definition.json (relative to script)"
}

# ============================================================
# Helpers: Definition + State
# ============================================================
function Load-Definition {
    # Always loads from $DefinitionPath
    $definition = Get-Content -Path $DefinitionPath -Raw | ConvertFrom-Json

    # Optional: If JSON includes changeUrlTemplate, we will copy it into $ChangeUrlTemplate
    # (Variable name stays the same.)
    if ($null -ne $definition.changeUrlTemplate -and -not [string]::IsNullOrWhiteSpace([string]$definition.changeUrlTemplate)) {
        $script:ChangeUrlTemplate = [string]$definition.changeUrlTemplate
    }

    return $definition
}

function Get-StatePath([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) { return $null }
    $safe = ($ChangeId.Trim() -replace '[^\w\-]', '_')
    return (Join-Path $StateDir "$safe.state.json")
}

function New-BlankState([string]$ChangeId, [object]$definition) {
    [pscustomobject]@{
        changeId          = $ChangeId
        definitionId      = if ($definition.definitionId) { [string]$definition.definitionId } else { "default" }
        definitionVersion = if ($definition.definitionVersion) { [int]$definition.definitionVersion } else { 1 }
        savedUtc          = $null
        itemStates        = @{}   # id -> @{ isChecked=bool; checkedUtc=string|null; notes=string }
        window            = @{ top=$null; left=$null; width=$null; height=$null }
    }
}

function Load-State([string]$ChangeId, [object]$definition) {
    $path = Get-StatePath $ChangeId
    if ($path -and (Test-Path $path)) {
        try {
            $state = Get-Content -Path $path -Raw | ConvertFrom-Json
        } catch {
            # If state got corrupted, start fresh rather than crashing
            $state = New-BlankState -ChangeId $ChangeId -definition $definition
        }

        # Ensure these exist and are up-to-date
        if (-not $state.PSObject.Properties.Name.Contains('definitionId')) {
            $state | Add-Member -MemberType NoteProperty -Name definitionId -Value (if ($definition.definitionId) { [string]$definition.definitionId } else { "default" })
        }
        if (-not $state.PSObject.Properties.Name.Contains('definitionVersion')) {
            $state | Add-Member -MemberType NoteProperty -Name definitionVersion -Value (if ($definition.definitionVersion) { [int]$definition.definitionVersion } else { 1 })
        }
        if (-not $state.PSObject.Properties.Name.Contains('itemStates') -or $null -eq $state.itemStates) {
            $state | Add-Member -MemberType NoteProperty -Name itemStates -Value (@{})
        }
        if (-not $state.PSObject.Properties.Name.Contains('window') -or $null -eq $state.window) {
            $state | Add-Member -MemberType NoteProperty -Name window -Value (@{ top=$null; left=$null; width=$null; height=$null })
        }

        return $state
    }

    return (New-BlankState -ChangeId $ChangeId -definition $definition)
}

function Save-State([object]$state) {
    if (-not $state) { return }
    $state.savedUtc = ([DateTime]::UtcNow.ToString("o"))
    $path = Get-StatePath $state.changeId
    if (-not $path) { return }

    $state | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding UTF8
}

function Ensure-ItemState([object]$state, [string]$id) {
    if (-not $state) { return }
    if ([string]::IsNullOrWhiteSpace($id)) { return }

    # itemStates is a PSCustomObject from JSON; we can add NoteProperty dynamically
    if (-not $state.itemStates.$id) {
        $state.itemStates | Add-Member -MemberType NoteProperty -Name $id -Value ([pscustomobject]@{
            isChecked  = $false
            checkedUtc = $null
            notes      = ""
        })
    }
}

function Open-ChangeLink([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) { return }
    $id = $ChangeId.Trim()
    $url = $ChangeUrlTemplate.Replace('{id}', [uri]::EscapeDataString($id))
    try { Start-Process $url | Out-Null } catch {}
}

# ============================================================
# ViewModel Item (INotifyPropertyChanged)
# ============================================================
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

    public string CheckedUtc { get; set; } // ISO string or null

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

# ============================================================
# Load XAML (separate file) and find controls
# ============================================================
[xml]$xaml = Get-Content -Path $XamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

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

# ============================================================
# Runtime objects
# ============================================================
$definition   = Load-Definition
$currentState = $null
$allSections  = @()

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
            $ci.Id         = [string]$it.id
            $ci.Text       = [string]$it.text
            $ci.LinkText   = [string]$it.linkText
            $ci.LinkUrl    = [string]$it.linkUrl
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
        $tab.Header = [string]$section.name

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

        $allSections += [pscustomobject]@{
            Name     = [string]$section.name
            Items    = $items
            Progress = $progress
        }
    }

    Update-SectionProgress
}

# Hyperlink handler (for checklist item links)
$window.AddHandler(
    [System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
    [System.Windows.Navigation.RequestNavigateEventHandler]{
        param($s, $e)
        try { Start-Process $e.Uri.AbsoluteUri | Out-Null } catch {}
        $e.Handled = $true
    }
)

function Load-Change([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) {
        Set-Status "Enter a ChangeID first."
        return
    }

    # Reload definition each time in case you edited the JSON
    $script:definition = Load-Definition

    $id = $ChangeId.Trim()
    $currentState = Load-State -ChangeId $id -definition $definition

    # Ensure any new items in the template get state entries
    foreach ($section in $definition.sections) {
        foreach ($it in $section.items) {
            Ensure-ItemState $currentState $it.id
        }
    }

    # Title
    $txtTitle.Text = "$([string]$definition.title) — $id"

    # Build UI
    Build-Tabs $currentState

    # Restore window geometry (per ChangeID)
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
    if (-not $currentState) {
        Set-Status "Nothing loaded yet."
        return
    }
    Save-State $currentState
    Set-Status "Saved: $($currentState.changeId)"
}

# Buttons
$btnLoad.Add_Click({ Load-Change $txtChangeId.Text })
$btnSave.Add_Click({ Save-Current })
$btnOpenCR.Add_Click({ Open-ChangeLink $txtChangeId.Text })

$btnEditDef.Add_Click({
    try {
        Start-Process notepad.exe $DefinitionPath | Out-Null
        [System.Windows.MessageBox]::Show(
            "Edit the template JSON and save it. Then click Load to refresh a ChangeID.",
            "Edit Template"
        ) | Out-Null
        Set-Status "Template opened: $DefinitionPath"
    } catch {
        Set-Status "Could not open template."
    }
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

# Save window geometry on close
$window.add_Closing({
    if ($currentState) {
        $currentState.window.left   = $window.Left
        $currentState.window.top    = $window.Top
        $currentState.window.width  = $window.Width
        $currentState.window.height = $window.Height
        Save-State $currentState
    }
})

# Default starter (edit if you want)
$txtChangeId.Text = "CHG0000000"
Set-Status "Enter ChangeID and click Load."

$null = $window.ShowDialog()
