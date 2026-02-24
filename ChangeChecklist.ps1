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
$ChangeUrlTemplate = "https://yourit.va.gov/nav_to.do?uri=change_request.do?sysparm_query=number={id}"

# Safe em-dash (avoids encoding problems)
$Dash = [char]0x2014

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
# Admin detection
# ============================================================
function Test-IsAdmin {
    try {
        $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
        $wp = New-Object Security.Principal.WindowsPrincipal($wi)

        # Local admin
        if ($wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { return $true }

        # Optional: AD group check (uncomment + set group if you want)
        # $AdminAdGroup = 'DOMAIN\YourChecklistAdmins'
        # $groups = $wi.Groups | ForEach-Object { $_.Translate([Security.Principal.NTAccount]).Value }
        # if ($groups -contains $AdminAdGroup) { return $true }

        return $false
    } catch {
        return $false
    }
}

# ============================================================
# Helpers: Definition + State
# ============================================================
function Load-Definition {
    $definition = Get-Content -Path $DefinitionPath -Raw | ConvertFrom-Json

    # Optional: allow JSON to override the template for the change URL
    # (Variable name stays the same.)
    if ($null -ne $definition.changeUrlTemplate -and -not [string]::IsNullOrWhiteSpace([string]$definition.changeUrlTemplate)) {
        $script:ChangeUrlTemplate = [string]$definition.changeUrlTemplate
    }

    return $definition
}

function Get-StatePath([string]$ChangeId) {
    if ([string]::IsNullOrWhiteSpace($ChangeId)) { return $null }
    $safe = ($ChangeId.Trim() -replace '[^\w\-]', '_')
    Join-Path $StateDir "$safe.state.json"
}

function New-BlankState([string]$ChangeId, [object]$definition) {
    [pscustomobject]@{
        changeId          = $ChangeId
        definitionId      = if ($definition.definitionId) { [string]$definition.definitionId } else { "default" }
        definitionVersion = if ($definition.definitionVersion) { [int]$definition.definitionVersion } else { 1 }
        savedUtc          = $null
        itemStates        = [pscustomobject]@{}  # MUST be object so we can safely add NoteProperties with hyphen keys
        window            = @{ top=$null; left=$null; width=$null; height=$null }
    }
}

function Load-State([string]$ChangeId, [object]$definition) {
    $path = Get-StatePath $ChangeId
    if ($path -and (Test-Path $path)) {
        try {
            $state = Get-Content -Path $path -Raw | ConvertFrom-Json
        } catch {
            $state = New-BlankState -ChangeId $ChangeId -definition $definition
        }

        if (-not $state.PSObject.Properties.Name.Contains('definitionId')) {
            $state | Add-Member -MemberType NoteProperty -Name definitionId -Value (if ($definition.definitionId) { [string]$definition.definitionId } else { "default" })
        }
        if (-not $state.PSObject.Properties.Name.Contains('definitionVersion')) {
            $state | Add-Member -MemberType NoteProperty -Name definitionVersion -Value (if ($definition.definitionVersion) { [int]$definition.definitionVersion } else { 1 })
        }

        if (-not $state.PSObject.Properties.Name.Contains('itemStates') -or $null -eq $state.itemStates) {
            $state | Add-Member -MemberType NoteProperty -Name itemStates -Value ([pscustomobject]@{})
        } elseif ($state.itemStates -is [hashtable]) {
            # normalize to PSCustomObject so hyphen keys work consistently
            $tmp = [pscustomobject]@{}
            foreach ($k in $state.itemStates.Keys) {
                $tmp | Add-Member -MemberType NoteProperty -Name $k -Value $state.itemStates[$k]
            }
            $state.itemStates = $tmp
        }

        if (-not $state.PSObject.Properties.Name.Contains('window') -or $null -eq $state.window) {
            $state | Add-Member -MemberType NoteProperty -Name window -Value (@{ top=$null; left=$null; width=$null; height=$null })
        }

        return $state
    }

    New-BlankState -ChangeId $ChangeId -definition $definition
}

function Save-State([object]$state) {
    if (-not $state) { return }
    $state.savedUtc = ([DateTime]::UtcNow.ToString("o"))
    $path = Get-StatePath $state.changeId
    if (-not $path) { return }
    $state | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding UTF8
}

# --- itemStates accessors (safe for keys like "pre-001")
function Get-ItemStateEntry([object]$state, [string]$id) {
    if (-not $state -or [string]::IsNullOrWhiteSpace($id)) { return $null }
    $prop = $state.itemStates.PSObject.Properties[$id]
    if ($null -eq $prop) { return $null }
    $prop.Value
}

function Set-ItemStateEntry([object]$state, [string]$id, [object]$value) {
    if (-not $state -or [string]::IsNullOrWhiteSpace($id)) { return }
    $prop = $state.itemStates.PSObject.Properties[$id]
    if ($null -eq $prop) {
        $state.itemStates | Add-Member -MemberType NoteProperty -Name $id -Value $value
    } else {
        $prop.Value = $value
    }
}

function Ensure-ItemState([object]$state, [string]$id) {
    if (-not $state -or [string]::IsNullOrWhiteSpace($id)) { return }
    if ($null -eq (Get-ItemStateEntry -state $state -id $id)) {
        Set-ItemStateEntry -state $state -id $id -value ([pscustomobject]@{
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

# ============================================================
# Load XAML (separate file) and find controls
# ============================================================
[xml]$xaml = Get-Content -Path $XamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$txtTitle       = $window.FindName('TxtTitle')
$txtChangeId    = $window.FindName('TxtChangeId')
$txtStatus      = $window.FindName('TxtStatus')
$btnLoad        = $window.FindName('BtnLoad')
$btnSave        = $window.FindName('BtnSave')
$btnOpenCR      = $window.FindName('BtnOpenCR')
$btnEditDef     = $window.FindName('BtnEditDefinition')
$btnOpenState   = $window.FindName('BtnOpenState')
$adminPanel     = $window.FindName('AdminPanel')
$checklistHost  = $window.FindName('ChecklistHost')

function Set-Status([string]$msg) { $txtStatus.Text = $msg }

# ============================================================
# Runtime objects
# ============================================================
$IsAdmin = Test-IsAdmin
if ($adminPanel) {
    $adminPanel.Visibility = if ($IsAdmin) { 'Visible' } else { 'Collapsed' }
}

# Load definition once (we reload on each Load too)
$definition   = Load-Definition
$currentState = $null

# List for progress tracking (do NOT use +=)
$allSections  = New-Object 'System.Collections.Generic.List[object]'

function Clear-ChecklistHost {
    $checklistHost.Children.Clear()
    $script:allSections = New-Object 'System.Collections.Generic.List[object]'
}

function Update-SectionProgress {
    foreach ($sec in $script:allSections) {
        $total = ($sec.Items | Measure-Object).Count
        $done  = @($sec.Items | Where-Object { $_.IsChecked }).Count
        $sec.Progress.Text = "Progress: $done / $total completed"
    }
}

function Build-ChecklistPage([object]$state) {
    Clear-ChecklistHost

    foreach ($section in $definition.sections) {

        # Section header
        $hdr = New-Object System.Windows.Controls.TextBlock
        $hdr.Text = [string]$section.name
        $hdr.FontSize = 16
        $hdr.FontWeight = 'Bold'
        $hdr.Margin = '0,10,0,6'
        $hdr.Foreground = [Windows.Media.Brushes]::White
        $checklistHost.Children.Add($hdr) | Out-Null

        # Progress line
        $progress = New-Object System.Windows.Controls.TextBlock
        $progress.Margin = '0,0,0,10'
        $progress.Foreground = [Windows.Media.Brushes]::LightSteelBlue
        $progress.FontSize = 12
        $checklistHost.Children.Add($progress) | Out-Null

        # Items collection
        $items = New-Object 'System.Collections.ObjectModel.ObservableCollection[ChecklistItem]'

        foreach ($it in $section.items) {
            $itemId = [string]$it.id

            Ensure-ItemState $state $itemId
            $st = Get-ItemStateEntry -state $state -id $itemId

            $ci = New-Object ChecklistItem
            $ci.Id         = $itemId
            $ci.Text       = [string]$it.text
            $ci.LinkText   = [string]$it.linkText
            $ci.LinkUrl    = [string]$it.linkUrl
            $ci.IsChecked  = [bool]$st.isChecked
            $ci.CheckedUtc = $st.checkedUtc
            $ci.Notes      = [string]$st.notes

            $ci.add_PropertyChanged({
                param($sender, $args)

                if (-not $currentState) { return }

                $sid = [string]$sender.Id
                Ensure-ItemState $currentState $sid
                $entry = Get-ItemStateEntry -state $currentState -id $sid

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

                Set-ItemStateEntry -state $currentState -id $sid -value $entry
                Save-State $currentState
                Set-Status "Saved: $($currentState.changeId)"
                Update-SectionProgress
            })

            $items.Add($ci)
        }

        # ItemsControl for this section
        $ic = New-Object System.Windows.Controls.ItemsControl
        $ic.ItemsSource = $items
        $ic.ItemTemplate = $window.Resources['ChecklistItemTemplate']
        $checklistHost.Children.Add($ic) | Out-Null

        # Track for progress updates
        $script:allSections.Add([pscustomobject]@{
            Name     = [string]$section.name
            Items    = $items
            Progress = $progress
        }) | Out-Null
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

    # Reload definition each time (admins may update JSON)
    $script:definition = Load-Definition

    $id = $ChangeId.Trim()
    $currentState = Load-State -ChangeId $id -definition $definition

    # Ensure every template item has state
    foreach ($section in $definition.sections) {
        foreach ($it in $section.items) {
            Ensure-ItemState $currentState ([string]$it.id)
        }
    }

    $txtTitle.Text = "$([string]$definition.title) $Dash $id"

    Build-ChecklistPage $currentState

    # Restore window geometry
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

# Admin-only buttons: safe even if hidden
if ($btnEditDef) {
    $btnEditDef.Add_Click({
        if (-not $IsAdmin) { return }
        try {
            Start-Process notepad.exe $DefinitionPath | Out-Null
            Set-Status "Template opened: $DefinitionPath"
        } catch {
            Set-Status "Could not open template."
        }
    })
}

if ($btnOpenState) {
    $btnOpenState.Add_Click({
        if (-not $IsAdmin) { return }
        if ([string]::IsNullOrWhiteSpace($txtChangeId.Text)) {
            [System.Windows.MessageBox]::Show("Enter a ChangeID first.","Open Saved State") | Out-Null
            return
        }
        $path = Get-StatePath $txtChangeId.Text
        if (-not (Test-Path $path)) {
            [System.Windows.MessageBox]::Show("No saved state yet for this ChangeID. Click Load first.","Open Saved State") | Out-Null
            return
        }
        Start-Process notepad.exe $path | Out-Null
    })
}

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

# Default starter
$txtChangeId.Text = "CHG0000000"
Set-Status "Enter ChangeID and click Load."

$null = $window.ShowDialog()
