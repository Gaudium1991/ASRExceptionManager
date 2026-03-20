Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================
# Global state
# =========================
$script:AccessToken = $null
$script:Headers = $null
$script:CurrentEvents = @()
$script:CurrentPolicies = @()
$script:CurrentPolicyCandidates = @()
$script:SelectedEvent = $null
$script:TenantId = ""
$script:ClientId = ""
$script:ClientSecret = ""

# =========================
# Helper functions
# =========================

function New-DefenderLogoBitmap {
    param(
        [int]$Width = 32,
        [int]$Height = 32
    )

    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)

    $blueDark = [System.Drawing.Color]::FromArgb(0, 90, 158)
    $blueLight = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $white = [System.Drawing.Color]::White

    $shield = New-Object System.Drawing.Drawing2D.GraphicsPath
    $shield.AddBezier(
        (New-Object System.Drawing.Point(16, 2)),
        (New-Object System.Drawing.Point(25, 4)),
        (New-Object System.Drawing.Point(28, 6)),
        (New-Object System.Drawing.Point(28, 11))
    )
    $shield.AddLine(28,11,28,18)
    $shield.AddBezier(
        (New-Object System.Drawing.Point(28, 18)),
        (New-Object System.Drawing.Point(28, 25)),
        (New-Object System.Drawing.Point(22, 29)),
        (New-Object System.Drawing.Point(16, 31))
    )
    $shield.AddBezier(
        (New-Object System.Drawing.Point(16, 31)),
        (New-Object System.Drawing.Point(10, 29)),
        (New-Object System.Drawing.Point(4, 25)),
        (New-Object System.Drawing.Point(4, 18))
    )
    $shield.AddLine(4,18,4,11)
    $shield.AddBezier(
        (New-Object System.Drawing.Point(4, 11)),
        (New-Object System.Drawing.Point(4, 6)),
        (New-Object System.Drawing.Point(7, 4)),
        (New-Object System.Drawing.Point(16, 2))
    )
    $shield.CloseFigure()

    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle 0,0,$Width,$Height),
        $blueLight,
        $blueDark,
        90
    )
    $g.FillPath($brush, $shield)

    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(220,255,255,255)), 1.5
    $g.DrawPath($pen, $shield)

    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70,255,255,255))), 15, 5, 2, 22)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70,255,255,255))), 7, 13, 18, 2)

    $g.FillRectangle((New-Object System.Drawing.SolidBrush $white), 14, 4, 4, 24)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $white), 7, 13, 18, 4)

    $brush.Dispose()
    $pen.Dispose()
    $g.Dispose()

    return $bmp
}

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $txtLog.AppendText("[$ts] $Message`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-ErrorDialog {
    param(
        [string]$Title,
        [object]$ErrorRecord
    )

    $msg = $null

    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $msg = $ErrorRecord.Exception.Message

        try {
            $resp = $ErrorRecord.Exception.Response
            if ($null -ne $resp) {
                $stream = $resp.GetResponseStream()
                if ($null -ne $stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $graphBody = $reader.ReadToEnd()
                    if (-not [string]::IsNullOrWhiteSpace($graphBody)) {
                        $msg += "`n`nGraph response:`n$graphBody"
                    }
                }
            }
        }
        catch {
            try {
                if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
                    $msg += "`n`nDetails:`n$($ErrorRecord.ErrorDetails.Message)"
                }
            }
            catch {}
        }
    }
    else {
        $msg = [string]$ErrorRecord
    }

    [System.Windows.Forms.MessageBox]::Show(
        $msg,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function ConvertTo-JsonCompat {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )

    try {
        return ($InputObject | ConvertTo-Json -Depth 100 -Compress)
    }
    catch {
        Add-Type -AssemblyName System.Web.Extensions
        $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $js.MaxJsonLength = 67108864
        return $js.Serialize($InputObject)
    }
}


function Update-ModeChart {
    param(
        [array]$Events
    )

    if ($null -eq $chartModes) { return }

    $chartModes.Series.Clear()
    $chartModes.Titles.Clear()

    $auditCount = @($Events | Where-Object { [string]$_.Mode -eq 'Audit' }).Count
    $blockCount = @($Events | Where-Object { [string]$_.Mode -eq 'Block' }).Count

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.Name = 'Modalita'
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $series.ChartArea = 'MainArea'
    $series.IsValueShownAsLabel = $true

    $pointAudit = $series.Points.AddXY('Audit', [int]$auditCount)
    $pointBlock = $series.Points.AddXY('Blocco', [int]$blockCount)

    $chartModes.Series.Add($series)
    [void]$chartModes.Titles.Add('Eventi ASR per modalità')

    $chartModes.ChartAreas['MainArea'].RecalculateAxesScale()
    $chartModes.Invalidate()
    $chartModes.Update()
}

function Update-UiLayout {
    $margin = 10
    $gap = 10
    $topRowY = 10
    $searchY = 45
    $searchH = 70
    $gridY = 125
    $statsWidth = 350
    $deviceH = 220
    $applyH = 95
    $minGridH = 240
    $minLogH = 72

    $clientW = $form.ClientSize.Width
    $clientH = $form.ClientSize.Height

    $btnConnect.Location = New-Object System.Drawing.Point(($clientW - $margin - $btnConnect.Width), 8)

    $txtSecret.Location = New-Object System.Drawing.Point(900, $topRowY)
    $txtSecret.Width = [Math]::Max(180, $btnConnect.Left - 15 - $txtSecret.Left)

    $grpSearch.Location = New-Object System.Drawing.Point($margin, $searchY)
    $grpSearch.Size = New-Object System.Drawing.Size(($clientW - ($margin * 2)), $searchH)
    $lblHint.Width = [Math]::Max(100, $grpSearch.ClientSize.Width - $lblHint.Left - 15)

    $bottomReserved = $gap + $deviceH + $gap + $applyH + $gap + $minLogH + $margin
    $gridH = [Math]::Max($minGridH, $clientH - $gridY - $bottomReserved)

    $gridWidth = [Math]::Max(620, $clientW - ($margin * 2) - $statsWidth - $gap)
    $dgvEvents.Location = New-Object System.Drawing.Point($margin, $gridY)
    $dgvEvents.Size = New-Object System.Drawing.Size($gridWidth, $gridH)

    $grpStats.Location = New-Object System.Drawing.Point(($dgvEvents.Right + $gap), $gridY)
    $grpStats.Size = New-Object System.Drawing.Size(($clientW - $grpStats.Left - $margin), $gridH)

    $deviceY = $dgvEvents.Bottom + $gap
    $grpDevice.Location = New-Object System.Drawing.Point($margin, $deviceY)
    $grpDevice.Size = New-Object System.Drawing.Size(($clientW - ($margin * 2)), $deviceH)

    $txtDevice.Width = [Math]::Max(200, $grpDevice.ClientSize.Width - 150 - 845)
    $lblManaged.Left = $txtDevice.Right + 20
    $txtManaged.Left = $lblManaged.Right + 5
    $txtManaged.Width = 300
    $lblEntra.Left = $txtManaged.Right + 15
    $txtEntra.Left = $lblEntra.Right + 5
    $txtEntra.Width = [Math]::Max(180, $grpDevice.ClientSize.Width - $txtEntra.Left - 15)

    $txtRule.Width = [Math]::Max(300, $grpDevice.ClientSize.Width - 150 - 400)
    $lblMode.Left = $txtRule.Right + 20
    $txtMode.Left = $lblMode.Right + 5
    $txtMode.Width = [Math]::Max(80, $grpDevice.ClientSize.Width - $txtMode.Left - 220)

    $txtAction.Width = [Math]::Max(300, $grpDevice.ClientSize.Width - $txtAction.Left - 15)
    $lstGroups.Width = [Math]::Max(250, [int](($grpDevice.ClientSize.Width - 45) / 2))
    $lstPolicies.Left = $lstGroups.Right + 20
    $lstPolicies.Width = [Math]::Max(250, $grpDevice.ClientSize.Width - $lstPolicies.Left - 15)

    $applyY = $grpDevice.Bottom + $gap
    $grpApply.Location = New-Object System.Drawing.Point($margin, $applyY)
    $grpApply.Size = New-Object System.Drawing.Size(($clientW - ($margin * 2)), $applyH)

    $btnApply.Left = $grpApply.ClientSize.Width - $btnApply.Width - 20
    $btnDump.Left = $grpApply.ClientSize.Width - $btnDump.Width - 20
    $chkAlsoGlobal.Left = $btnApply.Left - $chkAlsoGlobal.Width - 15
    $chkRuleSpecific.Left = $btnApply.Left - $chkRuleSpecific.Width - 15
    $chkRuleSpecific.Top = 12
    $chkAlsoGlobal.Top = 42
    $txtExclusion.Width = [Math]::Max(300, $chkRuleSpecific.Left - $txtExclusion.Left - 15)

    $txtLog.Location = New-Object System.Drawing.Point($margin, ($grpApply.Bottom + 8))
    $txtLog.Size = New-Object System.Drawing.Size(($clientW - ($margin * 2)), [Math]::Max($minLogH, $clientH - $txtLog.Top - $margin))
}

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $resp.access_token
}

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers
    )

    return Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers -ContentType 'application/json'
}

function Invoke-GraphPost {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][object]$Body
    )

    $json = ConvertTo-JsonCompat -InputObject $Body
    return Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -Body $json -ContentType 'application/json'
}

function Invoke-GraphPatch {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][object]$Body
    )

    $json = ConvertTo-JsonCompat -InputObject $Body
    return Invoke-RestMethod -Method Patch -Uri $Uri -Headers $Headers -Body $json -ContentType 'application/json'
}

function Invoke-GraphPut {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][object]$Body
    )

    $json = ConvertTo-JsonCompat -InputObject $Body
    return Invoke-RestMethod -Method Put -Uri $Uri -Headers $Headers -Body $json -ContentType 'application/json'
}

function Get-AllPages {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers
    )

    $items = New-Object System.Collections.ArrayList
    $next = $Uri

    while ($next) {
        $res = Invoke-GraphGet -Uri $next -Headers $Headers
        if ($res.value) {
            foreach ($i in $res.value) { [void]$items.Add($i) }
        }
        else {
            [void]$items.Add($res)
            break
        }
        $next = $res.'@odata.nextLink'
    }

    return $items
}

function Run-AsrHuntingQuery {
    param(
        [hashtable]$Headers,
        [int]$HoursBack = 72
    )

    $query = @"
DeviceEvents
| where Timestamp > ago(${HoursBack}h)
| where ActionType startswith "Asr"
| project Timestamp, DeviceName, DeviceId, ActionType, ReportId, InitiatingProcessFileName, InitiatingProcessFolderPath, FileName, FolderPath, SHA1, AccountName, AdditionalFields
| order by Timestamp desc
"@

    $body = @{ Query = $query }
    $uri = 'https://graph.microsoft.com/v1.0/security/runHuntingQuery'
    $res = Invoke-GraphPost -Uri $uri -Headers $Headers -Body $body

    $output = @()
    foreach ($r in $res.results) {
        $mode = 'Unknown'
        if ($r.ActionType -match 'Audit') { $mode = 'Audit' }
        elseif ($r.ActionType -match 'Block') { $mode = 'Block' }

        $ruleHint = $r.ActionType
        try {
            if ($r.AdditionalFields) {
                $af = $r.AdditionalFields
                if ($af -is [string]) {
                    $afObj = $af | ConvertFrom-Json -ErrorAction Stop
                    if ($afObj.RuleName) { $ruleHint = $afObj.RuleName }
                    elseif ($afObj.RuleId) { $ruleHint = $afObj.RuleId }
                }
            }
        }
        catch {}

        $output += [pscustomobject]@{
            Timestamp                    = [datetime]$r.Timestamp
            DeviceName                   = $r.DeviceName
            DeviceId                     = $r.DeviceId
            ActionType                   = $r.ActionType
            Mode                         = $mode
            RuleHint                     = $ruleHint
            InitiatingProcessFileName    = $r.InitiatingProcessFileName
            InitiatingProcessFolderPath  = $r.InitiatingProcessFolderPath
            FileName                     = $r.FileName
            FolderPath                   = $r.FolderPath
            SHA1                         = $r.SHA1
            AccountName                  = $r.AccountName
            ReportId                     = $r.ReportId
            AdditionalFields             = $r.AdditionalFields
        }
    }

    return $output
}

function Get-AsrRuleDisplay {
    param($Event)

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($Event.PSObject.Properties.Match('RuleHint').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Event.RuleHint)) {
        $candidates.Add([string]$Event.RuleHint)
    }
    if ($Event.PSObject.Properties.Match('ActionType').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Event.ActionType)) {
        $candidates.Add([string]$Event.ActionType)
    }

    $af = $null
    if ($Event.PSObject.Properties.Match('AdditionalFields').Count -gt 0) { $af = $Event.AdditionalFields }

    $afObj = $null
    try {
        if ($af -is [string] -and -not [string]::IsNullOrWhiteSpace($af)) {
            $afObj = $af | ConvertFrom-Json -Depth 100 -ErrorAction Stop
        }
        elseif ($af -and $af -isnot [string]) {
            $afObj = $af
        }
    }
    catch {}

    if ($afObj) {
        foreach ($prop in @('RuleName','RuleFriendlyName','AttackSurfaceReductionRuleName','Title','RuleId','AttackSurfaceReductionRuleId')) {
            try {
                $v = $afObj.$prop
                if (-not [string]::IsNullOrWhiteSpace([string]$v)) { $candidates.Add([string]$v) }
            }
            catch {}
        }
    }

    foreach ($c in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($c) -and $c -notmatch '^Asr[A-Za-z]+$') {
            return $c
        }
    }

    foreach ($c in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($c)) { return $c }
    }

    return 'N/D'
}

function Initialize-EventsGrid {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.AutoGenerateColumns = $false
    $Grid.Columns.Clear()
    $Grid.Rows.Clear()

    $defs = @(
        @{ Name='Timestamp'; Header='Timestamp'; Width=140; Visible=$true },
        @{ Name='Device'; Header='Device'; Width=140; Visible=$true },
        @{ Name='Mode'; Header='Mode'; Width=70; Visible=$true },
        @{ Name='RegolaASR'; Header='Regola ASR'; Width=280; Visible=$true },
        @{ Name='ActionType'; Header='ActionType'; Width=180; Visible=$true },
        @{ Name='Processo'; Header='Processo'; Width=130; Visible=$true },
        @{ Name='PathProcesso'; Header='Path Processo'; Width=240; Visible=$true },
        @{ Name='File'; Header='File'; Width=130; Visible=$true },
        @{ Name='PathFile'; Header='Path File'; Width=220; Visible=$true },
        @{ Name='Utente'; Header='Utente'; Width=120; Visible=$true },
        @{ Name='_Index'; Header='_Index'; Width=60; Visible=$false }
    )

    foreach ($d in $defs) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $d.Name
        $col.HeaderText = $d.Header
        $col.Width = $d.Width
        $col.Visible = $d.Visible
        $col.ReadOnly = $true
        [void]$Grid.Columns.Add($col)
    }

    $Grid.RowHeadersVisible = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.AutoSizeRowsMode = 'None'
    $Grid.ColumnHeadersVisible = $true
    $Grid.EnableHeadersVisualStyles = $true
}

function Populate-EventsGrid {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [array]$Events
    )

    if ($Grid.Columns.Count -eq 0) {
        Initialize-EventsGrid -Grid $Grid
    }

    $Grid.SuspendLayout()
    try {
        $Grid.Rows.Clear()
        for ($i = 0; $i -lt $Events.Count; $i++) {
            $e = $Events[$i]
            $rule = Get-AsrRuleDisplay -Event $e
            $idx = $Grid.Rows.Add()
            $row = $Grid.Rows[$idx]
            $row.Cells['Timestamp'].Value = if ($e.Timestamp) { ([datetime]$e.Timestamp).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            $row.Cells['Device'].Value = [string]$e.DeviceName
            $row.Cells['Mode'].Value = [string]$e.Mode
            $row.Cells['RegolaASR'].Value = [string]$rule
            $row.Cells['ActionType'].Value = [string]$e.ActionType
            $row.Cells['Processo'].Value = [string]$e.InitiatingProcessFileName
            $row.Cells['PathProcesso'].Value = [string]$e.InitiatingProcessFolderPath
            $row.Cells['File'].Value = [string]$e.FileName
            $row.Cells['PathFile'].Value = [string]$e.FolderPath
            $row.Cells['Utente'].Value = [string]$e.AccountName
            $row.Cells['_Index'].Value = $i
            $row.Tag = $e
        }
        if ($Grid.Rows.Count -gt 0) {
            $Grid.ClearSelection()
            $Grid.Rows[0].Selected = $true
            $Grid.CurrentCell = $Grid.Rows[0].Cells['Timestamp']
        }
    }
    finally {
        $Grid.ResumeLayout()
        $Grid.Refresh()
    }
}

function Get-ManagedDeviceByName {
    param(
        [string]$DeviceName,
        [hashtable]$Headers
    )

    if ([string]::IsNullOrWhiteSpace($DeviceName)) { return $null }

    # 1. Tentativo con nome completo
    $escaped = $DeviceName.Replace("'","''")
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$escaped'"
    $res = Invoke-GraphGet -Uri $uri -Headers $Headers

    if ($res.value -and $res.value.Count -gt 0) {
        Write-Log "Device trovato con nome completo: $DeviceName"
        return $res.value[0]
    }

    # 2. Fallback → rimuovo il suffisso di dominio
    $shortName = $DeviceName.Split('.')[0]

    if ($shortName -ne $DeviceName) {
        Write-Log "Fallback hostname senza dominio: $shortName"

        $escapedShort = $shortName.Replace("'","''")
        $uri2 = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$escapedShort'"
        $res2 = Invoke-GraphGet -Uri $uri2 -Headers $Headers

        if ($res2.value -and $res2.value.Count -gt 0) {
            Write-Log "Device trovato con hostname corto: $shortName"
            return $res2.value[0]
        }
    }

    Write-Log "Device NON trovato in Intune (nome: $DeviceName)"
    return $null
}

function Get-EntraDeviceByAzureAdDeviceId {
    param(
        [string]$AzureAdDeviceId,
        [hashtable]$Headers
    )

    if ([string]::IsNullOrWhiteSpace($AzureAdDeviceId)) { return $null }

    $escaped = $AzureAdDeviceId.Replace("'","''")
    $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$escaped'"
    $res = Invoke-GraphGet -Uri $uri -Headers $Headers
    if ($res.value -and $res.value.Count -gt 0) {
        return $res.value[0]
    }
    return $null
}

function Get-DeviceGroups {
    param(
        [string]$EntraObjectId,
        [hashtable]$Headers
    )

    if ([string]::IsNullOrWhiteSpace($EntraObjectId)) { return @() }

    $uri = "https://graph.microsoft.com/v1.0/devices/$EntraObjectId/transitiveMemberOf/microsoft.graph.group?`$select=id,displayName"
    return Get-AllPages -Uri $uri -Headers $Headers
}

function Get-ConfigurationPolicies {
    param([hashtable]$Headers)

    $uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies'
    return Get-AllPages -Uri $uri -Headers $Headers
}

function Get-PolicyAssignments {
    param(
        [string]$PolicyId,
        [hashtable]$Headers
    )

    $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$PolicyId/assignments"
    $res = Invoke-GraphGet -Uri $uri -Headers $Headers
    if ($res.value) { return $res.value }
    return @()
}

function Get-PolicySettings {
    param(
        [string]$PolicyId,
        [hashtable]$Headers
    )

    $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$PolicyId/settings?`$expand=settingDefinitions"
    $res = Invoke-GraphGet -Uri $uri -Headers $Headers
    if ($res.value) { return $res.value }
    return @()
}


function Get-TargetGroupIdsFromAssignment {
    param($Assignment)

    $ids = @()
    try {
        $target = $Assignment.target
        if ($null -eq $target) { return @() }

        foreach ($name in @('groupId','entraObjectId')) {
            if ($target.PSObject.Properties.Name -contains $name) {
                $value = $target.$name
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $ids += $value
                }
            }
        }
    }
    catch {}

    return $ids | Select-Object -Unique
}

function Test-PolicyLooksAsr {
    param(
        $Policy,
        [array]$Settings
    )

    $tokens = @()
    foreach ($name in @('name','description','templateReference')) {
        try {
            $value = $Policy.$name
            if ($value) { $tokens += ($value | Out-String) }
        }
        catch {}
    }

    foreach ($s in $Settings) {
        try {
            if ($s.settingDefinitions) {
                foreach ($d in $s.settingDefinitions) {
                    foreach ($field in @('displayName','description','id')) {
                        if ($d.$field) { $tokens += [string]$d.$field }
                    }
                }
            }
            if ($s.settingInstance -and $s.settingInstance.settingDefinitionId) {
                $tokens += [string]$s.settingInstance.settingDefinitionId
            }
        }
        catch {}
    }

    $blob = ($tokens -join ' | ')
    if ($blob -match 'attack surface reduction' -or $blob -match '\bASR\b' -or $blob -match 'asr') {
        return $true
    }
    return $false
}

function Find-AsrExclusionsSetting {
    param([array]$Settings)

    foreach ($s in $Settings) {
        $tokens = @()

        try {
            if ($s.settingDefinitions) {
                foreach ($d in $s.settingDefinitions) {
                    foreach ($field in @('displayName','description','id')) {
                        if ($d.$field) { $tokens += [string]$d.$field }
                    }
                }
            }

            if ($s.settingInstance) {
                foreach ($field in @('settingDefinitionId')) {
                    if ($s.settingInstance.$field) { $tokens += [string]$s.settingInstance.$field }
                }
            }
        }
        catch {}

        $blob = ($tokens -join ' | ')
        if (
            ($blob -match 'exclude' -or $blob -match 'exclusion') -and
            ($blob -match 'attack surface reduction' -or $blob -match '\bASR\b' -or $blob -match 'asr')
        ) {
            return $s
        }

        if ($blob -match 'exclude files and paths from asr rules') {
            return $s
        }
    }

    return $null
}


function Find-AsrGlobalExclusionsSetting {
    param([array]$Settings)

    foreach ($s in $Settings) {
        $tokens = @()
        try {
            foreach ($d in @($s.settingDefinitions)) {
                foreach ($field in @('displayName','description','id','name')) {
                    if ($d.$field) { $tokens += [string]$d.$field }
                }
            }
            if ($s.settingInstance -and $s.settingInstance.settingDefinitionId) {
                $tokens += [string]$s.settingInstance.settingDefinitionId
            }
        } catch {}

        $blob = ($tokens -join ' | ')
        $isGlobalByText = $blob -match '(?i)attack surface reduction only exclusions'
        $isSimpleCollection = $false
        try {
            $isSimpleCollection = ($null -ne $s.settingInstance.simpleSettingCollectionValue)
        } catch {}

        if ($isGlobalByText -or ($isSimpleCollection -and $blob -match '(?i)exclude|exclusion' -and $blob -match '(?i)attack surface reduction|\bASR\b|asr')) {
            return $s
        }
    }

    return $null
}

function Find-AsrPerRuleExclusionsSetting {
    param([array]$Settings)

    foreach ($s in $Settings) {
        $tokens = @()
        try {
            foreach ($d in @($s.settingDefinitions)) {
                foreach ($field in @('displayName','description','id','name')) {
                    if ($d.$field) { $tokens += [string]$d.$field }
                }
                foreach ($opt in @($d.options)) {
                    foreach ($dob in @($opt.dependedOnBy)) {
                        if ($dob.dependedOnBy) { $tokens += [string]$dob.dependedOnBy }
                    }
                }
            }
            if ($s.settingInstance -and $s.settingInstance.settingDefinitionId) {
                $tokens += [string]$s.settingInstance.settingDefinitionId
            }
        } catch {}

        $blob = ($tokens -join ' | ')
        $isGroupCollection = $false
        try {
            $isGroupCollection = ($s.settingInstance.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance')
        } catch {}
        $hasPerRuleMarker = $blob -match '(?i)_perruleexclusions|per[- ]rule exclusions'

        if ($isGroupCollection -or $hasPerRuleMarker) {
            return $s
        }
    }

    return $null
}

function Test-EventUnknownAsrRule {
    param($Event)
    $display = Get-AsrRuleDisplay -Event $Event
    if ([string]::IsNullOrWhiteSpace($display)) { return $true }
    if ($display -match '(?i)^unknown asr rule$|^n/d$|^unknown$') { return $true }

    $guid = Get-EventRuleGuid -Event $Event
    $candidates = @(Get-EventRuleCandidates -Event $Event)
    if (-not $guid -and $candidates.Count -le 1 -and [string]$Event.ActionType -eq [string]$display) { return $true }

    return $false
}

function Update-GlobalExclusionToggle {
    try {
        if (-not $script:SelectedEvent) { return }
        $isUnknown = Test-EventUnknownAsrRule -Event $script:SelectedEvent
        if ($isUnknown) {
            $chkAlsoGlobal.Checked = $true
            $chkAlsoGlobal.Enabled = $false
            $chkAlsoGlobal.Text = 'Applica anche Global Exclusion (auto: Unknown ASR Rule)'
            Write-Log 'Toggle automatico: Unknown ASR Rule rilevata, Global Exclusion attivata automaticamente.'
        }
        else {
            $chkAlsoGlobal.Enabled = $true
            $chkAlsoGlobal.Text = 'Applica anche Global Exclusion'
        }
    }
    catch {}
}

function Test-GlobalExclusionPresent {
    param(
        [Parameter(Mandatory=$true)]$Setting,
        [Parameter(Mandatory=$true)][string]$Value
    )

    try {
        foreach ($sv in @($Setting.settingInstance.simpleSettingCollectionValue)) {
            if ([string]$sv.value -eq $Value) { return $true }
        }
    } catch {}
    return $false
}

function New-StringSettingValueObject {
    param([string]$Value)
    return [pscustomobject]@{
        '@odata.type' = '#microsoft.graph.deviceManagementConfigurationStringSettingValue'
        settingValueTemplateReference = $null
        value = $Value
    }
}

function Normalize-RuleText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $t = $Text.ToLowerInvariant()
    $t = [regex]::Replace($t, '[^a-z0-9]+', ' ')
    $t = [regex]::Replace($t, '\s+', ' ').Trim()
    return $t
}

function Get-EventAdditionalFieldsObject {
    param($Event)
    try {
        if ($null -eq $Event) { return $null }
        if ($Event.PSObject.Properties.Match('AdditionalFields').Count -eq 0) { return $null }
        $af = $Event.AdditionalFields
        if ($af -is [string] -and -not [string]::IsNullOrWhiteSpace($af)) {
            return ($af | ConvertFrom-Json -Depth 100 -ErrorAction Stop)
        }
        elseif ($af -and $af -isnot [string]) {
            return $af
        }
    }
    catch {}
    return $null
}

function Get-EventRuleGuid {
    param($Event)
    $afObj = Get-EventAdditionalFieldsObject -Event $Event
    foreach ($prop in @('RuleId','AttackSurfaceReductionRuleId')) {
        try {
            $v = [string]$afObj.$prop
            if ($v -match '^[0-9a-fA-F-]{36}$') { return $v.ToUpperInvariant() }
        }
        catch {}
    }
    return $null
}

function Get-EventRuleCandidates {
    param($Event)
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($prop in @('RuleHint','ActionType')) {
        try {
            $v = [string]$Event.$prop
            if (-not [string]::IsNullOrWhiteSpace($v)) { $items.Add($v) }
        }
        catch {}
    }
    $afObj = Get-EventAdditionalFieldsObject -Event $Event
    foreach ($prop in @('RuleName','RuleFriendlyName','AttackSurfaceReductionRuleName','Title','RuleId','AttackSurfaceReductionRuleId')) {
        try {
            $v = [string]$afObj.$prop
            if (-not [string]::IsNullOrWhiteSpace($v)) { $items.Add($v) }
        }
        catch {}
    }
    return @($items | Select-Object -Unique)
}

function Resolve-PerRuleTargetFromSetting {
    param(
        [Parameter(Mandatory=$true)]$Setting,
        [Parameter(Mandatory=$true)]$Event
    )

    $defs = @($Setting.settingDefinitions)
    if ($defs.Count -eq 0) { return $null }

    $eventRuleGuid = Get-EventRuleGuid -Event $Event
    if ($eventRuleGuid) {
        foreach ($d in $defs) {
            if ($d.'@odata.type' -notmatch 'ChoiceSettingDefinition$') { continue }
            foreach ($opt in @($d.options)) {
                $val = [string]$opt.optionValue.value
                if ([string]::IsNullOrWhiteSpace($val) -or $val -notmatch '=') { continue }
                $guid = ($val.Split('=')[0]).ToUpperInvariant()
                if ($guid -ne $eventRuleGuid) { continue }
                foreach ($dob in @($opt.dependedOnBy)) {
                    $perRuleId = [string]$dob.dependedOnBy
                    if ($perRuleId -like '*_perruleexclusions') {
                        return [pscustomobject]@{
                            ParentDefinitionId = [string]$d.id
                            PerRuleDefinitionId = $perRuleId
                            MatchType = 'guid'
                        }
                    }
                }
            }
        }
    }

    $candidates = @(Get-EventRuleCandidates -Event $Event | ForEach-Object { Normalize-RuleText $_ } | Where-Object { $_ })
    if ($candidates.Count -eq 0) { return $null }

    $best = $null
    $bestScore = -1
    foreach ($d in $defs) {
        if ($d.'@odata.type' -notmatch 'ChoiceSettingDefinition$') { continue }
        $texts = @([string]$d.id,[string]$d.displayName,[string]$d.name,[string]$d.description)
        foreach ($opt in @($d.options)) {
            $texts += @([string]$opt.itemId,[string]$opt.displayName,[string]$opt.name,[string]$opt.description)
        }
        $blob = Normalize-RuleText ($texts -join ' ')
        $score = 0
        foreach ($cand in $candidates) {
            if (-not $cand) { continue }
            if ($blob.Contains($cand)) { $score += ($cand.Length + 50) }
            else {
                foreach ($token in ($cand -split ' ' | Where-Object { $_.Length -ge 4 })) {
                    if ($blob.Contains($token)) { $score += $token.Length }
                }
            }
        }
        if ($score -gt $bestScore) {
            $perRule = $null
            foreach ($opt in @($d.options)) {
                foreach ($dob in @($opt.dependedOnBy)) {
                    if ([string]$dob.dependedOnBy -like '*_perruleexclusions') { $perRule = [string]$dob.dependedOnBy; break }
                }
                if ($perRule) { break }
            }
            if ($perRule) {
                $bestScore = $score
                $best = [pscustomobject]@{
                    ParentDefinitionId = [string]$d.id
                    PerRuleDefinitionId = $perRule
                    MatchType = 'fuzzy'
                    Score = $score
                }
            }
        }
    }

    if ($bestScore -gt 0) { return $best }
    return $null
}

function Add-UniqueStringSettingValue {
    param(
        [Parameter(Mandatory=$true)][array]$Existing,
        [Parameter(Mandatory=$true)][string]$NewValue
    )
    $already = @($Existing | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'value' -and [string]$_.value -eq $NewValue }).Count -gt 0
    if ($already) { return ,$Existing }
    return @($Existing + (New-StringSettingValueObject -Value $NewValue))
}

function Ensure-PerRuleExclusionChildOnChoice {
    param(
        [Parameter(Mandatory=$true)]$ChoiceInstance,
        [Parameter(Mandatory=$true)][string]$PerRuleDefinitionId,
        [Parameter(Mandatory=$true)][string]$NewValue
    )

    if ($null -eq $ChoiceInstance.choiceSettingValue) {
        throw 'choiceSettingValue non presente sulla regola ASR selezionata.'
    }

    $children = @()
    if ($ChoiceInstance.choiceSettingValue.children) {
        $children = @($ChoiceInstance.choiceSettingValue.children)
    }

    $existingChild = $children | Where-Object { [string]$_.settingDefinitionId -eq $PerRuleDefinitionId } | Select-Object -First 1
    if ($existingChild) {
        $existingValues = @()
        if ($existingChild.simpleSettingCollectionValue) { $existingValues = @($existingChild.simpleSettingCollectionValue) }
        $existingChild.simpleSettingCollectionValue = Add-UniqueStringSettingValue -Existing $existingValues -NewValue $NewValue
    }
    else {
        $newChild = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance'
            settingDefinitionId = $PerRuleDefinitionId
            settingInstanceTemplateReference = $null
            simpleSettingCollectionValue = @((New-StringSettingValueObject -Value $NewValue))
        }
        $children += $newChild
    }

    $ChoiceInstance.choiceSettingValue.children = $children
}

function Add-ValueToSettingInstance {
    param(
        [Parameter(Mandatory=$true)]$Setting,
        [Parameter(Mandatory=$true)]$Event,
        [Parameter(Mandatory=$true)][string]$NewValue,
        [switch]$PreferPerRule
    )

    $clone = $Setting | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $inst = $clone.settingInstance
    if ($null -eq $inst) {
        throw 'settingInstance non trovato nel setting.'
    }

    if ($PreferPerRule -or $inst.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance') {
        $target = Resolve-PerRuleTargetFromSetting -Setting $clone -Event $Event
        if (-not $target) {
            throw 'Impossibile identificare la regola ASR corretta nel setting. Usa Dump setting JSON se vuoi una diagnosi più profonda.'
        }

        foreach ($group in @($inst.groupSettingCollectionValue)) {
            foreach ($child in @($group.children)) {
                if ([string]$child.settingDefinitionId -eq [string]$target.ParentDefinitionId) {
                    Ensure-PerRuleExclusionChildOnChoice -ChoiceInstance $child -PerRuleDefinitionId $target.PerRuleDefinitionId -NewValue $NewValue
                    return [pscustomobject]@{
                        id = $clone.id
                        settingInstance = $clone.settingInstance
                    }
                }
            }
        }

        throw "Regola ASR trovata nel mapping ma non presente nell'istanza della policy: $($target.ParentDefinitionId)"
    }

    if ($inst.PSObject.Properties.Name -contains 'simpleSettingCollectionValue') {
        $existing = @()
        if ($inst.simpleSettingCollectionValue) { $existing = @($inst.simpleSettingCollectionValue) }
        $inst.simpleSettingCollectionValue = Add-UniqueStringSettingValue -Existing $existing -NewValue $NewValue
        return [pscustomobject]@{ id = $clone.id; settingInstance = $clone.settingInstance }
    }

    throw 'Formato del setting exclusions non riconosciuto per questo tenant.'
}

function Convert-SettingForPolicyPut {
    param([Parameter(Mandatory=$true)]$Setting)

    $clean = $Setting | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    if ($clean.PSObject.Properties.Name -contains 'settingDefinitions') {
        $clean.PSObject.Properties.Remove('settingDefinitions')
    }
    if ($clean.PSObject.Properties.Name -contains 'id') {
        $clean.PSObject.Properties.Remove('id')
    }
    if (-not ($clean.PSObject.Properties.Name -contains '@odata.type')) {
        $clean | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue '#microsoft.graph.deviceManagementConfigurationSetting' -Force
    }
    return $clean
}

function Build-PolicyPutBody {
    param(
        [Parameter(Mandatory=$true)]$Policy,
        [Parameter(Mandatory=$true)][array]$AllSettings,
        [Parameter(Mandatory=$true)]$UpdatedSetting
    )

    $settingsForPut = @()
    foreach ($s in $AllSettings) {
        if ([string]$s.id -eq [string]$UpdatedSetting.id) {
            $settingsForPut += ,(Convert-SettingForPolicyPut -Setting $UpdatedSetting)
        }
        else {
            $settingsForPut += ,(Convert-SettingForPolicyPut -Setting $s)
        }
    }

    return [ordered]@{
        name = [string]$Policy.RawPolicy.name
        description = [string]$Policy.RawPolicy.description
        creationSource = $Policy.RawPolicy.creationSource
        platforms = [string]$Policy.RawPolicy.platforms
        technologies = [string]$Policy.RawPolicy.technologies
        roleScopeTagIds = @($Policy.RawPolicy.roleScopeTagIds)
        settings = $settingsForPut
        templateReference = [ordered]@{
            templateId = [string]$Policy.RawPolicy.templateReference.templateId
        }
    }
}

function Update-PolicySetting {
    param(
        [Parameter(Mandatory=$true)]$Policy,
        [Parameter(Mandatory=$true)]$Setting,
        [Parameter(Mandatory=$true)]$BodyObject,
        [Parameter(Mandatory=$true)][hashtable]$Headers
    )

    $putBody = Build-PolicyPutBody -Policy $Policy -AllSettings @($Policy.Settings) -UpdatedSetting $BodyObject
    $policyId = [string]$Policy.PolicyId
    if ([string]::IsNullOrWhiteSpace($policyId)) {
        throw 'PolicyId non valido.'
    }

    $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$policyId')"
    Write-Log "Aggiorno la configurationPolicy intera via PUT: $policyId"
    return Invoke-GraphPut -Uri $uri -Headers $Headers -Body $putBody
}



function Refresh-CurrentPolicyCandidateSettings {
    param(
        [int]$PolicyIndex,
        [hashtable]$Headers
    )

    $freshSettings = Get-PolicySettings -PolicyId $script:CurrentPolicyCandidates[$PolicyIndex].PolicyId -Headers $Headers
    $script:CurrentPolicyCandidates[$PolicyIndex].Settings = $freshSettings
    return $freshSettings
}

function Test-PerRuleExclusionPresent {
    param(
        [Parameter(Mandatory=$true)]$Setting,
        [Parameter(Mandatory=$true)]$Event,
        [Parameter(Mandatory=$true)][string]$Value
    )

    $target = Resolve-PerRuleTargetFromSetting -Setting $Setting -Event $Event
    if (-not $target) { return $false }

    foreach ($group in @($Setting.settingInstance.groupSettingCollectionValue)) {
        foreach ($child in @($group.children)) {
            if ([string]$child.settingDefinitionId -ne [string]$target.ParentDefinitionId) { continue }
            foreach ($gc in @($child.choiceSettingValue.children)) {
                if ([string]$gc.settingDefinitionId -ne [string]$target.PerRuleDefinitionId) { continue }
                foreach ($sv in @($gc.simpleSettingCollectionValue)) {
                    if ([string]$sv.value -eq $Value) { return $true }
                }
            }
        }
    }
    return $false
}

function Save-JsonDump {
    param(
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)][string]$FileName
    )

    $folder = Join-Path $env:TEMP 'ASR-Exception-Manager'
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }
    $path = Join-Path $folder $FileName
    ($Object | ConvertTo-Json -Depth 100) | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Find-CandidatePoliciesForGroups {
    param(
        [array]$Groups,
        [hashtable]$Headers
    )

    $groupIds = @($Groups | ForEach-Object { $_.id } | Select-Object -Unique)
    $allPolicies = Get-ConfigurationPolicies -Headers $Headers
    $candidates = @()

    $count = 0
    foreach ($policy in $allPolicies) {
        $count++
        Write-Log "Analizzo policy $count/$($allPolicies.Count): $($policy.name)"

        $assignments = @()
        try { $assignments = Get-PolicyAssignments -PolicyId $policy.id -Headers $Headers } catch { continue }

        $assignedGroupIds = @()
        foreach ($a in $assignments) {
            $assignedGroupIds += Get-TargetGroupIdsFromAssignment -Assignment $a
        }
        $assignedGroupIds = $assignedGroupIds | Select-Object -Unique

        $intersect = @($assignedGroupIds | Where-Object { $_ -in $groupIds })
        if ($intersect.Count -eq 0) { continue }

        $settings = @()
        try { $settings = Get-PolicySettings -PolicyId $policy.id -Headers $Headers } catch { $settings = @() }

        if (Test-PolicyLooksAsr -Policy $policy -Settings $settings) {
            $matchedNames = @($Groups | Where-Object { $_.id -in $intersect } | ForEach-Object { $_.displayName }) -join '; '
            $candidates += [pscustomobject]@{
                PolicyId          = $policy.id
                PolicyName        = $policy.name
                Description       = $policy.description
                MatchedGroupNames = $matchedNames
                Settings          = $settings
                RawPolicy         = $policy
            }
        }
    }

    return $candidates
}

function Validate-ExclusionPath {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw 'Inserisci un path di esclusione.'
    }

    if ($Value -notmatch '^[A-Za-z]:\\' -and $Value -notmatch '^\\\\') {
        throw 'Inserisci un percorso completo, ad esempio C:\Program Files\App\app.exe oppure \\\\server\\share\\file.exe'
    }
}

function Export-GridToCsv {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Path
    )

    $rows = @()
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $obj = [ordered]@{}
        foreach ($col in $Grid.Columns) {
            if ($col.Visible) { $obj[$col.HeaderText] = $row.Cells[$col.Index].Value }
        }
        $rows += [pscustomobject]$obj
    }

    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

# =========================
# UI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'ASR Exception Manager'
$form.Size = New-Object System.Drawing.Size(1420, 880)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $false
$form.MinimumSize = New-Object System.Drawing.Size(1200, 820)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(10, 14)
$lblTenant.Size = New-Object System.Drawing.Size(70, 20)
$lblTenant.Text = 'Tenant ID'
$form.Controls.Add($lblTenant)

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Location = New-Object System.Drawing.Point(85, 10)
$txtTenant.Size = New-Object System.Drawing.Size(320, 24)
$form.Controls.Add($txtTenant)

$lblClient = New-Object System.Windows.Forms.Label
$lblClient.Location = New-Object System.Drawing.Point(420, 14)
$lblClient.Size = New-Object System.Drawing.Size(60, 20)
$lblClient.Text = 'Client ID'
$form.Controls.Add($lblClient)

$txtClient = New-Object System.Windows.Forms.TextBox
$txtClient.Location = New-Object System.Drawing.Point(485, 10)
$txtClient.Size = New-Object System.Drawing.Size(320, 24)
$form.Controls.Add($txtClient)

$lblSecret = New-Object System.Windows.Forms.Label
$lblSecret.Location = New-Object System.Drawing.Point(820, 14)
$lblSecret.Size = New-Object System.Drawing.Size(75, 20)
$lblSecret.Text = 'Client Secret'
$form.Controls.Add($lblSecret)

$txtSecret = New-Object System.Windows.Forms.TextBox
$txtSecret.Location = New-Object System.Drawing.Point(900, 10)
$txtSecret.Size = New-Object System.Drawing.Size(320, 24)
$txtSecret.UseSystemPasswordChar = $true
$form.Controls.Add($txtSecret)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(1235, 35)
$btnConnect.Size = New-Object System.Drawing.Size(150, 28)
$btnConnect.Text = 'Connetti'
$form.Controls.Add($btnConnect)

$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = 'Ricerca eventi ASR'
$grpSearch.Location = New-Object System.Drawing.Point(10, 75)
$grpSearch.Size = New-Object System.Drawing.Size(1375, 70)
$form.Controls.Add($grpSearch)

$lblHours = New-Object System.Windows.Forms.Label
$lblHours.Location = New-Object System.Drawing.Point(15, 30)
$lblHours.Size = New-Object System.Drawing.Size(95, 20)
$lblHours.Text = 'Ultime ore'
$grpSearch.Controls.Add($lblHours)

$numHours = New-Object System.Windows.Forms.NumericUpDown
$numHours.Location = New-Object System.Drawing.Point(110, 28)
$numHours.Size = New-Object System.Drawing.Size(80, 24)
$numHours.Minimum = 1
$numHours.Maximum = 720
$numHours.Value = 72
$grpSearch.Controls.Add($numHours)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Location = New-Object System.Drawing.Point(210, 26)
$btnSearch.Size = New-Object System.Drawing.Size(180, 28)
$btnSearch.Text = 'Cerca ASR'
$grpSearch.Controls.Add($btnSearch)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(405, 26)
$btnExport.Size = New-Object System.Drawing.Size(180, 28)
$btnExport.Text = 'Esporta CSV'
$grpSearch.Controls.Add($btnExport)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Location = New-Object System.Drawing.Point(610, 30)
$lblHint.Size = New-Object System.Drawing.Size(740, 20)
$lblHint.Text = 'Seleziona una riga evento per caricare device, gruppi e policy ASR candidate.'
$grpSearch.Controls.Add($lblHint)

$dgvEvents = New-Object System.Windows.Forms.DataGridView
$dgvEvents.Location = New-Object System.Drawing.Point(10, 125)
$dgvEvents.Size = New-Object System.Drawing.Size(1015, 290)
$dgvEvents.ReadOnly = $true
$dgvEvents.SelectionMode = 'FullRowSelect'
$dgvEvents.MultiSelect = $false
$dgvEvents.AutoSizeColumnsMode = 'None'
$dgvEvents.ScrollBars = 'Both'
$dgvEvents.BackgroundColor = [System.Drawing.Color]::White
$dgvEvents.AllowUserToAddRows = $false
$form.Controls.Add($dgvEvents)
Initialize-EventsGrid -Grid $dgvEvents

$grpStats = New-Object System.Windows.Forms.GroupBox
$grpStats.Text = 'Istogramma Audit / Blocco'
$grpStats.Location = New-Object System.Drawing.Point(1035, 125)
$grpStats.Size = New-Object System.Drawing.Size(350, 290)
$form.Controls.Add($grpStats)

$chartModes = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chartModes.Dock = 'Fill'
$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartArea.Name = 'MainArea'
$chartArea.AxisX.Interval = 1
$chartArea.AxisY.MajorGrid.Enabled = $true
$chartArea.AxisY.Minimum = 0
$chartModes.ChartAreas.Add($chartArea)
$grpStats.Controls.Add($chartModes)
Update-ModeChart -Events @()

$grpDevice = New-Object System.Windows.Forms.GroupBox
$grpDevice.Text = 'Dettagli device / gruppi / policy'
$grpDevice.Location = New-Object System.Drawing.Point(10, 425)
$grpDevice.Size = New-Object System.Drawing.Size(1375, 220)
$form.Controls.Add($grpDevice)

$lblDevice = New-Object System.Windows.Forms.Label
$lblDevice.Location = New-Object System.Drawing.Point(15, 28)
$lblDevice.Size = New-Object System.Drawing.Size(130, 20)
$lblDevice.Text = 'Device selezionato'
$grpDevice.Controls.Add($lblDevice)

$txtDevice = New-Object System.Windows.Forms.TextBox
$txtDevice.Location = New-Object System.Drawing.Point(150, 25)
$txtDevice.Size = New-Object System.Drawing.Size(380, 24)
$txtDevice.ReadOnly = $true
$grpDevice.Controls.Add($txtDevice)

$lblManaged = New-Object System.Windows.Forms.Label
$lblManaged.Location = New-Object System.Drawing.Point(550, 28)
$lblManaged.Size = New-Object System.Drawing.Size(110, 20)
$lblManaged.Text = 'ManagedDeviceId'
$grpDevice.Controls.Add($lblManaged)

$txtManaged = New-Object System.Windows.Forms.TextBox
$txtManaged.Location = New-Object System.Drawing.Point(665, 25)
$txtManaged.Size = New-Object System.Drawing.Size(300, 24)
$txtManaged.ReadOnly = $true
$grpDevice.Controls.Add($txtManaged)

$lblEntra = New-Object System.Windows.Forms.Label
$lblEntra.Location = New-Object System.Drawing.Point(980, 28)
$lblEntra.Size = New-Object System.Drawing.Size(90, 20)
$lblEntra.Text = 'EntraObjectId'
$grpDevice.Controls.Add($lblEntra)

$txtEntra = New-Object System.Windows.Forms.TextBox
$txtEntra.Location = New-Object System.Drawing.Point(1075, 25)
$txtEntra.Size = New-Object System.Drawing.Size(285, 24)
$txtEntra.ReadOnly = $true
$grpDevice.Controls.Add($txtEntra)

$lblRule = New-Object System.Windows.Forms.Label
$lblRule.Location = New-Object System.Drawing.Point(15, 60)
$lblRule.Size = New-Object System.Drawing.Size(120, 20)
$lblRule.Text = 'Regola ASR'
$grpDevice.Controls.Add($lblRule)

$txtRule = New-Object System.Windows.Forms.TextBox
$txtRule.Location = New-Object System.Drawing.Point(150, 57)
$txtRule.Size = New-Object System.Drawing.Size(810, 24)
$txtRule.ReadOnly = $true
$grpDevice.Controls.Add($txtRule)

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Location = New-Object System.Drawing.Point(980, 60)
$lblMode.Size = New-Object System.Drawing.Size(45, 20)
$lblMode.Text = 'Mode'
$grpDevice.Controls.Add($lblMode)

$txtMode = New-Object System.Windows.Forms.TextBox
$txtMode.Location = New-Object System.Drawing.Point(1030, 57)
$txtMode.Size = New-Object System.Drawing.Size(110, 24)
$txtMode.ReadOnly = $true
$grpDevice.Controls.Add($txtMode)

$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Location = New-Object System.Drawing.Point(15, 90)
$lblAction.Size = New-Object System.Drawing.Size(120, 20)
$lblAction.Text = 'ActionType'
$grpDevice.Controls.Add($lblAction)

$txtAction = New-Object System.Windows.Forms.TextBox
$txtAction.Location = New-Object System.Drawing.Point(150, 87)
$txtAction.Size = New-Object System.Drawing.Size(1210, 24)
$txtAction.ReadOnly = $true
$grpDevice.Controls.Add($txtAction)

$lstGroups = New-Object System.Windows.Forms.ListBox
$lstGroups.Location = New-Object System.Drawing.Point(15, 120)
$lstGroups.Size = New-Object System.Drawing.Size(510, 125)
$grpDevice.Controls.Add($lstGroups)

$lstPolicies = New-Object System.Windows.Forms.ListBox
$lstPolicies.Location = New-Object System.Drawing.Point(545, 120)
$lstPolicies.Size = New-Object System.Drawing.Size(815, 125)
$grpDevice.Controls.Add($lstPolicies)

$grpApply = New-Object System.Windows.Forms.GroupBox
$grpApply.Text = 'Applica esclusione ASR nella policy selezionata'
$grpApply.Location = New-Object System.Drawing.Point(10, 685)
$grpApply.Size = New-Object System.Drawing.Size(1375, 95)
$form.Controls.Add($grpApply)

$lblExcl = New-Object System.Windows.Forms.Label
$lblExcl.Location = New-Object System.Drawing.Point(15, 30)
$lblExcl.Size = New-Object System.Drawing.Size(120, 20)
$lblExcl.Text = 'Path esclusione'
$grpApply.Controls.Add($lblExcl)

$txtExclusion = New-Object System.Windows.Forms.TextBox
$txtExclusion.Location = New-Object System.Drawing.Point(140, 26)
$txtExclusion.Size = New-Object System.Drawing.Size(760, 24)
$grpApply.Controls.Add($txtExclusion)

$chkRuleSpecific = New-Object System.Windows.Forms.CheckBox
$chkRuleSpecific.Location = New-Object System.Drawing.Point(880, 12)
$chkRuleSpecific.Size = New-Object System.Drawing.Size(290, 24)
$chkRuleSpecific.Text = 'Tenta per-rule exclusion (best effort)'
$chkRuleSpecific.Checked = $false
$grpApply.Controls.Add($chkRuleSpecific)

$chkAlsoGlobal = New-Object System.Windows.Forms.CheckBox
$chkAlsoGlobal.Location = New-Object System.Drawing.Point(880, 42)
$chkAlsoGlobal.Size = New-Object System.Drawing.Size(280, 24)
$chkAlsoGlobal.Text = 'Applica anche Global Exclusion'
$chkAlsoGlobal.Checked = $false
$grpApply.Controls.Add($chkAlsoGlobal)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Location = New-Object System.Drawing.Point(1195, 24)
$btnApply.Size = New-Object System.Drawing.Size(160, 30)
$btnApply.Text = 'Applica policy'
$grpApply.Controls.Add($btnApply)

$btnDump = New-Object System.Windows.Forms.Button
$btnDump.Location = New-Object System.Drawing.Point(1195, 58)
$btnDump.Size = New-Object System.Drawing.Size(160, 26)
$btnDump.Text = 'Dump setting JSON'
$grpApply.Controls.Add($btnDump)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 788)
$txtLog.Size = New-Object System.Drawing.Size(1375, 72)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# =========================
# Event handlers
# =========================
$btnConnect.Add_Click({
    try {
        $script:TenantId = $txtTenant.Text.Trim()
        $script:ClientId = $txtClient.Text.Trim()
        $script:ClientSecret = $txtSecret.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($script:TenantId) -or
            [string]::IsNullOrWhiteSpace($script:ClientId) -or
            [string]::IsNullOrWhiteSpace($script:ClientSecret)) {
            throw 'Compila Tenant ID, Client ID e Client Secret.'
        }

        Write-Log 'Richiesta token Graph...'
        $script:AccessToken = Get-GraphToken -TenantId $script:TenantId -ClientId $script:ClientId -ClientSecret $script:ClientSecret
        $script:Headers = @{ Authorization = "Bearer $($script:AccessToken)" }

        Write-Log 'Connessione Graph riuscita.'
        [System.Windows.Forms.MessageBox]::Show('Connessione riuscita.', 'OK', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch {
        Write-Log "Errore connessione: $($_.Exception.Message)"
        Show-ErrorDialog -Title 'Errore connessione' -ErrorRecord $_
    }
})

$btnSearch.Add_Click({
    try {
        if (-not $script:Headers) { throw 'Connettiti prima al tenant.' }

        Write-Log 'Avvio ricerca eventi ASR...'
        $hours = [int]$numHours.Value
        $script:CurrentEvents = @(Run-AsrHuntingQuery -Headers $script:Headers -HoursBack $hours)
        Populate-EventsGrid -Grid $dgvEvents -Events $script:CurrentEvents
        Update-ModeChart -Events $script:CurrentEvents
        Write-Log "Trovati $($script:CurrentEvents.Count) eventi ASR."
        if ($script:CurrentEvents.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Nessun evento ASR trovato nel periodo selezionato.', 'Informazione', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    }
    catch {
        Write-Log "Errore ricerca ASR: $($_.Exception.Message)"
        Show-ErrorDialog -Title 'Errore ricerca ASR' -ErrorRecord $_
    }
})

$dgvEvents.Add_SelectionChanged({
    try {
        if ($dgvEvents.SelectedRows.Count -eq 0) { return }
        $row = $dgvEvents.SelectedRows[0]
        $event = $row.Tag
        if ($null -eq $event) {
            $idxCell = $row.Cells['_Index'].Value
            if ($null -ne $idxCell) {
                $idx = [int]$idxCell
                if ($idx -ge 0 -and $idx -lt $script:CurrentEvents.Count) { $event = $script:CurrentEvents[$idx] }
            }
        }
        if ($null -eq $event) { return }

        $script:SelectedEvent = $event
        $txtDevice.Text = [string]$event.DeviceName
        $txtManaged.Text = ''
        $txtRule.Text = Get-AsrRuleDisplay -Event $event
        $txtAction.Text = [string]$event.ActionType
        $txtMode.Text = [string]$event.Mode
        $txtEntra.Text = ''
        $lstGroups.Items.Clear()
        $lstPolicies.Items.Clear()
        $script:CurrentPolicyCandidates = @()

        Write-Log "Carico contesto per device $($event.DeviceName)..."

        $md = Get-ManagedDeviceByName -DeviceName $event.DeviceName -Headers $script:Headers
        if ($null -eq $md) {
            Write-Log 'Managed device non trovato in Intune.'
            return
        }

        $txtManaged.Text = [string]$md.id

        $entra = Get-EntraDeviceByAzureAdDeviceId -AzureAdDeviceId $md.azureADDeviceId -Headers $script:Headers
        if ($null -eq $entra) {
            Write-Log 'Device Entra non trovato.'
            return
        }

        $txtEntra.Text = [string]$entra.id

        $groups = @(Get-DeviceGroups -EntraObjectId $entra.id -Headers $script:Headers)

        $lstGroups.Items.Clear()

        $cleanGroups = @(
            $groups |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.displayName) } |
            Sort-Object displayName -Unique
        )

        foreach ($g in $cleanGroups) {
            $rowText = "{0} [{1}]" -f [string]$g.displayName, [string]$g.id
            [void]$lstGroups.Items.Add($rowText)
        }
        Write-Log "Gruppi univoci trovati: $($cleanGroups.Count)"

        $script:CurrentPolicyCandidates = Find-CandidatePoliciesForGroups -Groups $groups -Headers $script:Headers
        foreach ($p in $script:CurrentPolicyCandidates) {
            [void]$lstPolicies.Items.Add("$($p.PolicyName) | Gruppi: $($p.MatchedGroupNames)")
        }
        Write-Log "Policy ASR candidate trovate: $($script:CurrentPolicyCandidates.Count)"

        if (-not [string]::IsNullOrWhiteSpace($event.InitiatingProcessFolderPath) -and -not [string]::IsNullOrWhiteSpace($event.InitiatingProcessFileName)) {
            $txtExclusion.Text = (Join-Path $event.InitiatingProcessFolderPath $event.InitiatingProcessFileName)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($event.FolderPath) -and -not [string]::IsNullOrWhiteSpace($event.FileName)) {
            $txtExclusion.Text = (Join-Path $event.FolderPath $event.FileName)
        }
        Update-GlobalExclusionToggle
    }
    catch {
        Write-Log "Errore caricamento dettagli evento: $($_.Exception.Message)"
    }
})

$btnDump.Add_Click({
    try {
        if ($lstPolicies.SelectedIndex -lt 0) { throw 'Seleziona una policy candidata.' }
        $policy = $script:CurrentPolicyCandidates[$lstPolicies.SelectedIndex]
        $setting = Find-AsrExclusionsSetting -Settings $policy.Settings
        if (-not $setting) { throw 'Setting exclusions non trovato nella policy selezionata.' }

        $dumpPath = Save-JsonDump -Object $setting -FileName ("setting_{0}.json" -f $policy.PolicyId)
        Write-Log "Dump salvato in: $dumpPath"
        [System.Windows.Forms.MessageBox]::Show("Dump salvato in:`n$dumpPath", 'Dump completato', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch {
        Write-Log "Errore dump JSON: $($_.Exception.Message)"
        Show-ErrorDialog -Title 'Errore dump JSON' -ErrorRecord $_
    }
})

$btnApply.Add_Click({
    try {
        if (-not $script:Headers) { throw 'Connettiti prima al tenant.' }
        if ($lstPolicies.SelectedIndex -lt 0) { throw 'Seleziona una policy candidata dalla lista.' }

        $exclusion = $txtExclusion.Text.Trim()
        Validate-ExclusionPath -Value $exclusion

        $policyIndex = $lstPolicies.SelectedIndex
        $policy = $script:CurrentPolicyCandidates[$policyIndex]
        Write-Log "Uso policy: $($policy.PolicyName)"

        $settings = @($policy.Settings)
        $globalSetting = Find-AsrGlobalExclusionsSetting -Settings $settings
        $perRuleSetting = Find-AsrPerRuleExclusionsSetting -Settings $settings

        $ruleIsUnknown = Test-EventUnknownAsrRule -Event $script:SelectedEvent
        $mustApplyGlobal = $chkAlsoGlobal.Checked -or $ruleIsUnknown
        $canTryPerRule = $chkRuleSpecific.Checked -and (-not $ruleIsUnknown)

        if ($ruleIsUnknown) {
            Write-Log 'Regola ASR non determinabile (Unknown ASR Rule): userò Attack Surface Reduction Only Exclusions.'
        }

        if (-not $globalSetting -and -not $perRuleSetting) {
            throw 'Non trovo né il setting ASR Global Exclusions né il setting per-rule nella policy selezionata. Usa "Dump setting JSON" per analizzare la struttura.'
        }

        $appliedScopes = New-Object System.Collections.Generic.List[string]

        if ($canTryPerRule) {
            if (-not $perRuleSetting) {
                Write-Log 'Setting per-rule non trovato nella policy selezionata: salto applicazione specifica.'
            }
            else {
                Write-Log 'Opzione per-rule exclusion richiesta: provo a scrivere l''esclusione sulla regola ASR specifica.'
                $updatedPerRule = Add-ValueToSettingInstance -Setting $perRuleSetting -Event $script:SelectedEvent -NewValue $exclusion -PreferPerRule:$true
                $null = Update-PolicySetting -Policy $policy -Setting $perRuleSetting -BodyObject $updatedPerRule -Headers $script:Headers
                Start-Sleep -Seconds 2
                $freshSettings = Refresh-CurrentPolicyCandidateSettings -PolicyIndex $policyIndex -Headers $script:Headers
                $policy = $script:CurrentPolicyCandidates[$policyIndex]
                $freshPerRuleSetting = Find-AsrPerRuleExclusionsSetting -Settings $freshSettings
                if (-not $freshPerRuleSetting) {
                    throw 'Policy aggiornata ma non riesco a rileggere il setting per-rule ASR.'
                }
                if (-not (Test-PerRuleExclusionPresent -Setting $freshPerRuleSetting -Event $script:SelectedEvent -Value $exclusion)) {
                    $verifyDump = Save-JsonDump -Object $freshPerRuleSetting -FileName ('verify-perrule-' + [guid]::NewGuid().ToString() + '.json')
                    throw "Graph ha risposto OK, ma al refresh l'esclusione per-rule non risulta salvata. Dump: $verifyDump"
                }
                [void]$appliedScopes.Add('Per-Rule')
                $settings = @($freshSettings)
                $globalSetting = Find-AsrGlobalExclusionsSetting -Settings $settings
            }
        }

        if ($mustApplyGlobal) {
            if (-not $globalSetting) {
                throw 'Non trovo il setting "Attack Surface Reduction Only Exclusions" nella policy selezionata.'
            }

            Write-Log 'Applico l''esclusione al setting globale: Attack Surface Reduction Only Exclusions.'
            $updatedGlobal = Add-ValueToSettingInstance -Setting $globalSetting -Event $script:SelectedEvent -NewValue $exclusion
            $null = Update-PolicySetting -Policy $policy -Setting $globalSetting -BodyObject $updatedGlobal -Headers $script:Headers
            Start-Sleep -Seconds 2
            $freshSettings = Refresh-CurrentPolicyCandidateSettings -PolicyIndex $policyIndex -Headers $script:Headers
            $policy = $script:CurrentPolicyCandidates[$policyIndex]
            $freshGlobalSetting = Find-AsrGlobalExclusionsSetting -Settings $freshSettings
            if (-not $freshGlobalSetting) {
                throw 'Policy aggiornata ma non riesco a rileggere il setting globale ASR exclusions.'
            }
            if (-not (Test-GlobalExclusionPresent -Setting $freshGlobalSetting -Value $exclusion)) {
                $verifyDump = Save-JsonDump -Object $freshGlobalSetting -FileName ('verify-global-' + [guid]::NewGuid().ToString() + '.json')
                throw "Graph ha risposto OK, ma al refresh l'esclusione globale non risulta salvata. Dump: $verifyDump"
            }
            [void]$appliedScopes.Add('Global')
        }

        if ($appliedScopes.Count -eq 0) {
            throw 'Nessuna applicazione eseguita. Se vuoi la regola specifica abilita la checkbox per-rule; per la globale abilita "Applica anche Global Exclusion".'
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Esclusione applicata e verificata con successo.`n`nPolicy: $($policy.PolicyName)`nPath: $exclusion`nScope: $([string]::Join(', ', $appliedScopes))",
            'Operazione completata',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-Log "Errore applicazione esclusione: $($_.Exception.Message)"
        Show-ErrorDialog -Title 'Errore applicazione esclusione' -ErrorRecord $_
    }
})


$btnExport.Add_Click({
    try {
        if ($dgvEvents.Rows.Count -eq 0) { throw 'Non ci sono dati da esportare.' }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV (*.csv)|*.csv'
        $sfd.FileName = 'asr-events.csv'
        if ($sfd.ShowDialog() -eq 'OK') {
            Export-GridToCsv -Grid $dgvEvents -Path $sfd.FileName
            Write-Log "CSV esportato in $($sfd.FileName)"
        }
    }
    catch {
        Write-Log "Errore export CSV: $($_.Exception.Message)"
        Show-ErrorDialog -Title 'Errore export CSV' -ErrorRecord $_
    }
})

# =========================
# Startup message
# =========================
Write-Log 'Prerequisiti consigliati Graph application permissions:'
Write-Log '- ThreatHunting.Read.All'
Write-Log '- DeviceManagementManagedDevices.Read.All'
Write-Log '- DeviceManagementConfiguration.ReadWrite.All'
Write-Log '- Group.Read.All / Directory.Read.All'
Write-Log 'Nota: la modifica diretta della policy impatta tutti i device assegnati a quella policy/gruppo.'
Write-Log 'Per troubleshooting usa il pulsante "Dump setting JSON".'
Write-Log 'V1.7: toggle automatico Global Exclusion + fix lookup device Intune con fallback hostname senza dominio.'
Write-Log 'Per Endpoint Security / ASR lo script usa PUT su configurationPolicies, in linea con il payload pubblico mostrato da Microsoft per questi profili.'

$form.Add_Shown({ Update-UiLayout })
$form.Add_Resize({ Update-UiLayout })

[void]$form.ShowDialog()
