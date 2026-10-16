<#
.SYNOPSIS
    Searches an OpenText Content Manager (ex-HP TRIM) dataset and downloads every
    retrievable electronic document into a folder, with a CSV manifest.

.DESCRIPTION
    A read-only command line client over the Content Manager .NET SDK, intended to
    stage document content on disk so that an agent (or a human) can read it.

    Works with either SDK generation:
      * CM 10.x - 23.x   TRIM.SDK.dll        namespace TRIM.SDK
      * RM8 / CM9        HP.HPTRIM.SDK.dll   namespace HP.HPTRIM.SDK

    The script never checks out, modifies or deletes a record. Documents are
    extracted with checkDocumentOut = $false.

    Requires Windows PowerShell 5.1 and a Content Manager client installation whose
    bitness matches the PowerShell host (see README.md).

.PARAMETER Term
    A simple search term. Combined with -SearchField into "<SearchField>:<Term>".

.PARAMETER SearchField
    The clause ("index") the term is matched against, e.g. content, title, notes.
    Run with -ListSearchFields to see the real clause names for your dataset.

.PARAMETER SearchString
    A raw Content Manager search string, e.g. 'content:budget and registeredOn:this year'.
    Used verbatim; -Term and -SearchField are ignored.

.PARAMETER MaxFiles
    Stop after this many documents have been downloaded successfully.

.PARAMETER MaxScan
    Hard ceiling on how many records are examined. 0 means unlimited.

.PARAMETER MaxAgeDays
    Only download documents whose record was registered within this many days.

.PARAMETER SortBy
    Content Manager sort string. A trailing '-' sorts descending.

.EXAMPLE
    .\Invoke-TrimSearchDownload.ps1 -TestConnection -Verbose

.EXAMPLE
    .\Invoke-TrimSearchDownload.ps1 -ListSearchFields | Format-Table

.EXAMPLE
    .\Invoke-TrimSearchDownload.ps1 -Term budget -OutputFolder .\out -MaxFiles 25 -MaxAgeDays 90

.EXAMPLE
    .\Invoke-TrimSearchDownload.ps1 -SearchString 'title:minutes' -OutputFolder D:\agent\inbox -IncludeExtension pdf,docx
#>
[CmdletBinding(DefaultParameterSetName = 'TermSearch', SupportsShouldProcess = $true)]
param(
    # ---------------- what to search ----------------
    [Parameter(ParameterSetName = 'TermSearch', Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Term,

    [Parameter(ParameterSetName = 'TermSearch')]
    [ValidateNotNullOrEmpty()]
    [string]$SearchField = 'content',

    [Parameter(ParameterSetName = 'RawSearch', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchString,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [string]$FilterString,

    [string]$ObjectType = 'Record',

    # ---------------- what to keep ----------------
    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [ValidateRange(1, 1000000)]
    [int]$MaxFiles = 100,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [ValidateRange(0, 10000000)]
    [int]$MaxScan = 0,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [ValidateRange(0, 1000000)]
    [int]$MaxAgeDays = 0,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [string]$SortBy = 'registeredOn-',

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [ValidateRange(0, 1000000)]
    [int]$MaxFileSizeMB = 0,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [string[]]$IncludeExtension,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [string[]]$ExcludeExtension,

    # ---------------- where it goes ----------------
    [Parameter(ParameterSetName = 'TermSearch', Mandatory = $true)]
    [Parameter(ParameterSetName = 'RawSearch', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [string]$ManifestPath,

    [Parameter(ParameterSetName = 'TermSearch')]
    [Parameter(ParameterSetName = 'RawSearch')]
    [switch]$Overwrite,

    [string]$LogPath,

    # ---------------- connection ----------------
    [string]$Dataset,
    [string]$WorkgroupServer,
    [int]$WorkgroupServerPort = 0,
    [string]$TrustedUser,
    [string]$SdkPath,

    # ---------------- discovery ----------------
    [Parameter(ParameterSetName = 'ListFields', Mandatory = $true)]
    [switch]$ListSearchFields,

    [Parameter(ParameterSetName = 'ListSort', Mandatory = $true)]
    [switch]$ListSortFields,

    [Parameter(ParameterSetName = 'TestConn', Mandatory = $true)]
    [switch]$TestConnection
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$script:TrimAsm        = $null
$script:TrimNamespaces = @('TRIM.SDK', 'HP.HPTRIM.SDK')
$script:TrimTypeCache  = @{}
$script:LogFile        = $null

#region Infrastructure

function Write-TrimLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $line = '{0:yyyy-MM-dd HH:mm:ss}  {1,-5}  {2}' -f (Get-Date), $Level, $Message

    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Verbose $Message }
    }

    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 }
        catch { Write-Warning "Could not write to log file $($script:LogFile): $($_.Exception.Message)" }
    }
}

function Resolve-TrimSdk {
    <#
        Locates and loads the Content Manager SDK assembly and works out which
        namespace it uses. Returns the path that was loaded.
    #>
    [CmdletBinding()]
    param([string]$SdkPath)

    if ($script:TrimAsm) { return $script:TrimAsm.Location }

    $names = @('TRIM.SDK.dll', 'HP.HPTRIM.SDK.dll')
    $candidates = New-Object System.Collections.Generic.List[string]

    # An explicit hint wins. Accept either the dll itself or the folder holding it.
    foreach ($hint in @($SdkPath, $env:TRIM_SDK_PATH, $PSScriptRoot)) {
        if (-not $hint) { continue }
        if (Test-Path -LiteralPath $hint -PathType Leaf) {
            $candidates.Add((Resolve-Path -LiteralPath $hint).Path)
        }
        elseif (Test-Path -LiteralPath $hint -PathType Container) {
            foreach ($n in $names) {
                $p = Join-Path $hint $n
                if (Test-Path -LiteralPath $p -PathType Leaf) { $candidates.Add($p) }
            }
        }
    }

    # Otherwise probe the usual client install locations.
    if ($candidates.Count -eq 0) {
        $roots = @(
            (Join-Path $env:ProgramFiles 'OpenText'),
            (Join-Path $env:ProgramFiles 'Micro Focus'),
            (Join-Path $env:ProgramFiles 'Hewlett-Packard'),
            (Join-Path ${env:ProgramFiles(x86)} 'OpenText'),
            (Join-Path ${env:ProgramFiles(x86)} 'Micro Focus'),
            (Join-Path ${env:ProgramFiles(x86)} 'Hewlett-Packard')
        )
        foreach ($root in $roots) {
            if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }
            foreach ($n in $names) {
                $hits = Get-ChildItem -LiteralPath $root -Filter $n -Recurse -Depth 3 -File -ErrorAction SilentlyContinue
                foreach ($h in $hits) { $candidates.Add($h.FullName) }
            }
        }
    }

    if ($candidates.Count -eq 0) {
        throw @"
Could not find the Content Manager SDK assembly (TRIM.SDK.dll or HP.HPTRIM.SDK.dll).
Pass -SdkPath pointing at the dll, set the TRIM_SDK_PATH environment variable, or
copy the dll next to this script.
"@
    }

    $dll = $candidates[0]
    Write-TrimLog "Loading SDK assembly: $dll"

    # The managed assembly binds to native siblings, so its folder must be on PATH.
    $dir = Split-Path -Parent $dll
    if (($env:PATH -split ';') -notcontains $dir) {
        $env:PATH = "$dir;$env:PATH"
        Write-TrimLog "Prepended SDK folder to PATH: $dir"
    }

    try {
        $script:TrimAsm = [System.Reflection.Assembly]::LoadFrom($dll)
    }
    catch {
        $base = $_.Exception.GetBaseException()
        if ($base -is [System.BadImageFormatException]) {
            if ([Environment]::Is64BitProcess) {
                $hostBits  = '64-bit'
                $otherHost = Join-Path $env:WinDir 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
            }
            else {
                $hostBits  = '32-bit'
                $otherHost = Join-Path $env:WinDir 'System32\WindowsPowerShell\v1.0\powershell.exe'
            }
            throw @"
Bitness mismatch loading $dll
This PowerShell host is $hostBits but the SDK is the other architecture.
Re-run the script with:
    $otherHost -File "$PSCommandPath" ...
"@
        }
        throw "Failed to load SDK assembly $dll : $($base.Message)"
    }

    # Work out the namespace by looking for the Database type, then put the winner first.
    $found = $null
    foreach ($ns in $script:TrimNamespaces) {
        if ($script:TrimAsm.GetType("$ns.Database", $false)) { $found = $ns; break }
    }
    if (-not $found) {
        throw "Loaded $dll but found neither TRIM.SDK.Database nor HP.HPTRIM.SDK.Database. Is this the Content Manager SDK?"
    }
    $script:TrimNamespaces = @($found) + ($script:TrimNamespaces | Where-Object { $_ -ne $found })
    Write-TrimLog "SDK namespace: $found (assembly version $($script:TrimAsm.GetName().Version))"

    return $dll
}

function Get-TrimType {
    <#
        The only namespace-aware code in the script. Everything else relies on
        PowerShell late binding, which does not care which namespace a type came from.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)

    if ($script:TrimTypeCache.ContainsKey($Name)) { return $script:TrimTypeCache[$Name] }

    $type = $null
    foreach ($ns in $script:TrimNamespaces) {
        $type = $script:TrimAsm.GetType("$ns.$Name", $false)
        if ($type) { break }
    }
    if (-not $type) {
        throw "SDK type $Name not found in $($script:TrimNamespaces -join ' or '). The SDK version may be incompatible."
    }

    $script:TrimTypeCache[$Name] = $type
    return $type
}

function Get-TrimProp {
    <#
        Reads the first property that exists, tolerating naming drift between SDK
        versions (Number vs RecordNumber, DateRegistered vs RegisteredOn) and
        swallowing SDK-side access errors such as a permission denial.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Name,
        $Default = $null
    )

    foreach ($n in $Name) {
        try {
            $prop = $InputObject.PSObject.Properties[$n]
            if ($prop) { return $prop.Value }
        }
        catch {
            Write-TrimLog "Could not read property $n : $($_.Exception.Message)"
        }
    }
    return $Default
}

function Get-TrimDisplay {
    <# Flattens an SDK object (RecordType, Location, TrimURI) to a string for the manifest. #>
    [CmdletBinding()]
    param($InputObject)

    if ($null -eq $InputObject) { return '' }
    if ($InputObject -is [string]) { return $InputObject }

    foreach ($n in @('NameString', 'Name', 'Title')) {
        try {
            $prop = $InputObject.PSObject.Properties[$n]
            if ($prop -and $prop.Value) { return [string]$prop.Value }
        }
        catch { }
    }
    try { return [string]$InputObject.ToString() } catch { return '' }
}

function ConvertFrom-TrimDate {
    <# TrimDateTime -> [datetime], or $null when the SDK value is unset ("clear"). #>
    [CmdletBinding()]
    param($InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [datetime]) { return $InputObject }

    try { if ($InputObject.IsClear) { return $null } } catch { }
    try { return [datetime]$InputObject.ToDateTime() } catch { }
    try { return [datetime]$InputObject.ToString() } catch { }
    return $null
}

function Format-TrimExtensionList {
    <# Normalises pdf / .PDF / .pdf to a lowercase, dotless set. #>
    [CmdletBinding()]
    param([string[]]$Extension)

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($e in @($Extension)) {
        if (-not $e) { continue }
        $out.Add($e.Trim().TrimStart('.').ToLowerInvariant())
    }
    return $out.ToArray()
}

#endregion

#region SDK operations

function Get-TrimObjectType {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ObjectType)

    $enumType = Get-TrimType 'BaseObjectTypes'
    try { return [enum]::Parse($enumType, $ObjectType, $true) }
    catch {
        $valid = ([enum]::GetNames($enumType) | Sort-Object) -join ', '
        throw "Unknown -ObjectType $ObjectType. Valid values: $valid"
    }
}

function Connect-TrimDatabase {
    <#
        With no -Dataset and no -WorkgroupServer the SDK connects to the dataset the
        logged-on user is configured for, using integrated Windows authentication.
    #>
    [CmdletBinding()]
    param(
        [string]$Dataset,
        [string]$WorkgroupServer,
        [int]$WorkgroupServerPort = 0,
        [string]$TrustedUser
    )

    $dbType = Get-TrimType 'Database'
    $db = [Activator]::CreateInstance($dbType)

    try {
        if ($Dataset)             { $db.Id                  = $Dataset }
        if ($WorkgroupServer)     { $db.WorkgroupServerName = $WorkgroupServer }
        if ($WorkgroupServerPort) { $db.WorkgroupServerPort = $WorkgroupServerPort }
        if ($TrustedUser)         { $db.TrustedUser         = $TrustedUser }

        if (-not $Dataset -and -not $WorkgroupServer) {
            # Prefer whatever the installed client is pointed at.
            if ($dbType.GetMethod('SetAsClientDatabase', [Type[]]@())) {
                try { $db.SetAsClientDatabase() }
                catch { Write-TrimLog "SetAsClientDatabase() failed, falling back to the SDK default: $($_.Exception.GetBaseException().Message)" 'WARN' }
            }
        }

        Write-TrimLog ('Connecting (dataset={0}, server={1})' -f $db.Id, $db.WorkgroupServerName)
        $db.Connect()

        if (-not $db.IsConnected) { throw 'Database.Connect() returned without establishing a connection.' }

        Write-TrimLog ('Connected to dataset {0} {1}' -f $db.Id, (Get-TrimDisplay (Get-TrimProp $db @('Name', 'DatabaseName'))))
        return $db
    }
    catch {
        try { $db.Dispose() } catch { }
        throw
    }
}

function Get-TrimSearchClause {
    <#
        Enumerates the connected dataset's real search clauses, so clause names never
        have to be guessed. InternalName is what SetSearchString expects.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Database,
        [Parameter(Mandatory = $true)][string]$ObjectType
    )

    $defType  = Get-TrimType 'SearchClauseDef'
    $objDefId = Get-TrimObjectType -ObjectType $ObjectType

    $method = $defType.GetMethod('GetAllSearchClauseIds')
    if (-not $method) {
        throw 'SearchClauseDef.GetAllSearchClauseIds was not found; this SDK version cannot enumerate clauses.'
    }

    # Documented order is (objectDefId, database); try the reverse before giving up.
    $ids = $null
    foreach ($argSet in @(, @($objDefId, $Database)) + @(, @($Database, $objDefId))) {
        try { $ids = $method.Invoke($null, $argSet); break }
        catch { Write-TrimLog "GetAllSearchClauseIds rejected an argument order: $($_.Exception.GetBaseException().Message)" }
    }
    if ($null -eq $ids) {
        throw 'Could not call SearchClauseDef.GetAllSearchClauseIds with either argument order.'
    }

    foreach ($id in $ids) {
        try {
            $def = [Activator]::CreateInstance($defType, @($id, $Database))

            $canSort = $false
            try { $canSort = [bool]$def.CanSort($objDefId) } catch { }

            [pscustomobject]@{
                InternalName    = [string](Get-TrimProp $def @('InternalName') '')
                Name            = [string](Get-TrimProp $def @('Name', 'Caption') '')
                ParameterFormat = Get-TrimDisplay (Get-TrimProp $def @('ParameterFormat'))
                CanSort         = $canSort
            }
        }
        catch {
            Write-TrimLog "Skipped search clause $id : $($_.Exception.GetBaseException().Message)" 'WARN'
        }
    }
}

function New-TrimSearch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Database,
        [Parameter(Mandatory = $true)][string]$ObjectType,
        [Parameter(Mandatory = $true)][string]$SearchString,
        [string]$FilterString,
        [string]$SortBy
    )

    $searchType = Get-TrimType 'TrimMainObjectSearch'
    $objDefId   = Get-TrimObjectType -ObjectType $ObjectType
    $search     = [Activator]::CreateInstance($searchType, @($Database, $objDefId))

    try {
        Write-TrimLog "Search string: $SearchString"
        $search.SetSearchString($SearchString)
    }
    catch {
        $msg = $_.Exception.GetBaseException().Message
        throw "Content Manager rejected the search string: $msg$([Environment]::NewLine)Run with -ListSearchFields to see the clause names this dataset accepts."
    }

    if ($FilterString) {
        try { $search.SetFilterString($FilterString) }
        catch { throw "Content Manager rejected the filter string: $($_.Exception.GetBaseException().Message)" }
    }

    if ($SortBy) {
        try {
            $search.SetSortString($SortBy)
            Write-TrimLog "Sort: $SortBy"
        }
        catch {
            Write-TrimLog "Sort string $SortBy was rejected; using the dataset default order. Run -ListSortFields for sortable clauses." 'WARN'
        }
    }

    $info = Get-TrimProp $search @('QueryInfoMessage') ''
    if ($info) { Write-TrimLog "QueryInfoMessage: $info" }

    return $search
}

function Get-TrimSearchCount {
    <# FastCount first (cheap, may be -1), falling back to the accurate Count. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Search)

    try {
        $fast = [int64](Get-TrimProp $Search @('FastCount') -1)
        if ($fast -ge 0) { return $fast }
    }
    catch { }

    try { return [int64]$Search.Count } catch { return [int64](-1) }
}

function Get-TrimOutputFileName {
    <#
        Builds a collision-safe file name prefixed with the record number, kept short
        enough that the full path stays inside the 260-char limit that Windows
        PowerShell 5.1 is subject to.
    #>
    [CmdletBinding()]
    param(
        [string]$RecordNumber,
        [string]$SuggestedFileName,
        [string]$Extension,
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][hashtable]$UsedNames
    )

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()

    $clean = {
        param([string]$Value)
        if (-not $Value) { return '' }
        $sb = New-Object System.Text.StringBuilder
        foreach ($ch in $Value.ToCharArray()) {
            if ($invalid -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
        }
        return ($sb.ToString() -replace '\s+', ' ').Trim(' ', '.')
    }

    $suggested = & $clean $SuggestedFileName
    $number    = & $clean $RecordNumber

    $ext = ''
    if ($suggested) { $ext = [System.IO.Path]::GetExtension($suggested) }
    if (-not $ext -and $Extension) { $ext = '.' + $Extension.TrimStart('.') }
    $ext = (& $clean $ext)
    if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

    $stem = ''
    if ($suggested) { $stem = [System.IO.Path]::GetFileNameWithoutExtension($suggested) }
    if (-not $stem) { $stem = 'document' }

    $prefix = ''
    if ($number) { $prefix = $number + '__' }

    # Reserve room for the folder, the extension and a de-duplication suffix.
    $budget = 250 - $Folder.Length - 1 - $prefix.Length - $ext.Length - 6
    if ($budget -lt 8) { $budget = 8 }
    if ($stem.Length -gt $budget) { $stem = $stem.Substring(0, $budget).Trim() }

    $candidate = $prefix + $stem + $ext
    $suffix    = 1
    while ($UsedNames.ContainsKey($candidate.ToLowerInvariant())) {
        $suffix++
        $candidate = '{0}{1}_{2}{3}' -f $prefix, $stem, $suffix, $ext
    }

    $UsedNames[$candidate.ToLowerInvariant()] = $true
    return $candidate
}

function Export-TrimRecordDocument {
    <#
        Extracts the electronic document to $Path.

        checkDocumentOut is hard-coded to $false: this client is read-only and must
        never leave a record checked out.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $written = $Record.GetDocument($Path, $false, $null, $null)

    $actual = $Path
    if ($written -and (Test-Path -LiteralPath $written -PathType Leaf)) { $actual = $written }

    if (-not (Test-Path -LiteralPath $actual -PathType Leaf)) {
        throw "GetDocument() reported success but no file was written to $actual"
    }

    $info = Get-Item -LiteralPath $actual
    if ($info.Length -eq 0) {
        throw "GetDocument() wrote a zero-byte file to $actual"
    }

    return $info
}

#endregion

#region Main

$summary = $null
$db      = $null

try {
    if ($LogPath) {
        $logDir = Split-Path -Parent $LogPath
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $script:LogFile = $LogPath
        Write-TrimLog "=== Invoke-TrimSearchDownload started (parameter set: $($PSCmdlet.ParameterSetName)) ==="
    }

    Resolve-TrimSdk -SdkPath $SdkPath | Out-Null

    $db = Connect-TrimDatabase -Dataset $Dataset -WorkgroupServer $WorkgroupServer `
        -WorkgroupServerPort $WorkgroupServerPort -TrustedUser $TrustedUser

    # ---------------- discovery modes ----------------

    if ($TestConnection) {
        [pscustomobject]@{
            SdkAssembly    = $script:TrimAsm.Location
            SdkVersion     = $script:TrimAsm.GetName().Version.ToString()
            SdkNamespace   = $script:TrimNamespaces[0]
            HostBitness    = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
            PSVersion      = $PSVersionTable.PSVersion.ToString()
            Dataset        = [string]$db.Id
            DatasetName    = Get-TrimDisplay (Get-TrimProp $db @('Name', 'DatabaseName'))
            WorkgroupServer= [string]$db.WorkgroupServerName
            CurrentUser    = Get-TrimDisplay (Get-TrimProp $db @('CurrentUser'))
            IsConnected    = [bool]$db.IsConnected
        }
        return
    }

    if ($ListSearchFields) {
        Get-TrimSearchClause -Database $db -ObjectType $ObjectType | Sort-Object InternalName
        return
    }

    if ($ListSortFields) {
        Get-TrimSearchClause -Database $db -ObjectType $ObjectType |
            Where-Object { $_.CanSort } |
            Sort-Object InternalName
        return
    }

    # ---------------- download mode ----------------

    if ($PSCmdlet.ParameterSetName -eq 'TermSearch') {
        # Quote the term so phrases and punctuation survive.
        $effectiveSearch = '{0}:"{1}"' -f $SearchField, ($Term -replace '"', '')
    }
    else {
        $effectiveSearch = $SearchString
    }

    if (-not $ManifestPath) { $ManifestPath = Join-Path $OutputFolder 'manifest.csv' }

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        if ($WhatIfPreference) {
            Write-Host "What if: would create output folder $OutputFolder"
        }
        else {
            New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
            Write-TrimLog "Created output folder: $OutputFolder"
        }
    }
    $resolvedFolder = $OutputFolder
    if (Test-Path -LiteralPath $OutputFolder) {
        $resolvedFolder = (Resolve-Path -LiteralPath $OutputFolder).Path
    }

    $search = New-TrimSearch -Database $db -ObjectType $ObjectType `
        -SearchString $effectiveSearch -FilterString $FilterString -SortBy $SortBy

    $cutoff = $null
    if ($MaxAgeDays -gt 0) { $cutoff = (Get-Date).Date.AddDays(-$MaxAgeDays) }

    $include  = Format-TrimExtensionList -Extension $IncludeExtension
    $exclude  = Format-TrimExtensionList -Extension $ExcludeExtension
    $maxBytes = [int64]0
    if ($MaxFileSizeMB -gt 0) { $maxBytes = [int64]$MaxFileSizeMB * 1MB }

    $hasClientFilter = ($null -ne $cutoff) -or ($include.Count -gt 0) -or ($exclude.Count -gt 0) -or ($maxBytes -gt 0)

    # Only cap the result set server-side when nothing is filtered client-side.
    # Otherwise the cap would truncate before filtering and silently under-deliver.
    if (-not $hasClientFilter) {
        try { $search.LimitOnRowsReturned = $MaxFiles }
        catch { Write-TrimLog "Could not set LimitOnRowsReturned: $($_.Exception.GetBaseException().Message)" }
    }

    # Newest-first sorting lets us stop at the first record older than the cutoff.
    $earlyExitOnAge = ($null -ne $cutoff) -and ($SortBy -and $SortBy.Trim() -eq 'registeredOn-')
    if ($null -ne $cutoff -and -not $earlyExitOnAge) {
        Write-TrimLog "-MaxAgeDays with a custom -SortBy: every result must be examined (no early exit)." 'WARN'
    }

    $total = Get-TrimSearchCount -Search $search
    if ($total -ge 0) { Write-TrimLog "Search matched $total record(s)." }
    else { Write-TrimLog 'Result count unavailable for this search.' }

    $rows       = New-Object System.Collections.Generic.List[object]
    $usedNames  = @{}
    $scanned    = 0
    $downloaded = 0
    $failed     = 0
    $stopEarly  = $false

    foreach ($rec in $search) {
        $scanned++

        if ($MaxScan -gt 0 -and $scanned -gt $MaxScan) {
            $scanned--
            Write-TrimLog "Reached -MaxScan ($MaxScan); stopping."
            break
        }

        $uri      = Get-TrimDisplay (Get-TrimProp $rec @('Uri'))
        $recNum   = Get-TrimDisplay (Get-TrimProp $rec @('Number', 'RecordNumber', 'LongNumber'))
        $title    = Get-TrimDisplay (Get-TrimProp $rec @('Title', 'TypedTitle'))
        $recType  = Get-TrimDisplay (Get-TrimProp $rec @('RecordType'))
        $container= Get-TrimDisplay (Get-TrimProp $rec @('Container'))
        $author   = Get-TrimDisplay (Get-TrimProp $rec @('Author', 'Creator'))
        $extRaw   = [string](Get-TrimProp $rec @('Extension') '')
        $ext      = $extRaw.Trim().TrimStart('.').ToLowerInvariant()
        $regDate  = ConvertFrom-TrimDate (Get-TrimProp $rec @('DateRegistered', 'RegisteredOn'))
        $creDate  = ConvertFrom-TrimDate (Get-TrimProp $rec @('DateCreated', 'CreatedOn'))

        $size = [int64]0
        try { $size = [int64](Get-TrimProp $rec @('DocumentSize') 0) } catch { }

        $isElectronic = $false
        try { $isElectronic = [bool](Get-TrimProp $rec @('IsElectronic') $false) } catch { }

        $status   = ''
        $errText  = ''
        $fileName = ''

        if (-not $isElectronic) {
            $status = 'SkippedNotElectronic'
        }
        elseif ($null -ne $cutoff -and $null -ne $regDate -and $regDate -lt $cutoff) {
            $status = 'SkippedTooOld'
            if ($earlyExitOnAge) { $stopEarly = $true }
        }
        elseif ($include.Count -gt 0 -and $include -notcontains $ext) {
            $status = 'SkippedExtension'
        }
        elseif ($exclude.Count -gt 0 -and $exclude -contains $ext) {
            $status = 'SkippedExtension'
        }
        elseif ($maxBytes -gt 0 -and $size -gt $maxBytes) {
            $status  = 'SkippedTooLarge'
            $errText = "$size bytes exceeds -MaxFileSizeMB $MaxFileSizeMB"
        }
        else {
            $suggested = [string](Get-TrimProp $rec @('SuggestedFileName') '')
            $fileName  = Get-TrimOutputFileName -RecordNumber $recNum -SuggestedFileName $suggested `
                -Extension $ext -Folder $resolvedFolder -UsedNames $usedNames
            $target = Join-Path $resolvedFolder $fileName

            if ((Test-Path -LiteralPath $target -PathType Leaf) -and -not $Overwrite) {
                $status = 'SkippedExists'
            }
            elseif (-not $PSCmdlet.ShouldProcess($target, 'Download document')) {
                $status = 'WouldDownload'
                $downloaded++
            }
            else {
                try {
                    $info = Export-TrimRecordDocument -Record $rec -Path $target
                    $fileName = $info.Name
                    $size     = $info.Length
                    $status   = 'Downloaded'
                    $downloaded++
                    Write-TrimLog "Downloaded $recNum -> $fileName ($size bytes)"
                }
                catch {
                    $status  = 'Failed'
                    $errText = $_.Exception.GetBaseException().Message
                    $failed++
                    Write-TrimLog "Record $recNum could not be retrieved: $errText" 'WARN'
                }
            }
        }

        $rows.Add([pscustomobject]@{
            FileName          = $fileName
            Uri               = $uri
            RecordNumber      = $recNum
            Title             = $title
            RecordType        = $recType
            Extension         = $ext
            DocumentSizeBytes = $size
            DateRegistered    = if ($regDate) { $regDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            DateCreated       = if ($creDate) { $creDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            Container         = $container
            Author            = $author
            Status            = $status
            Error             = $errText
        })

        $activity = "Downloaded $downloaded of $MaxFiles  (examined $scanned)"
        if ($total -gt 0) {
            $pct = [Math]::Min(100, [int](($scanned / $total) * 100))
            Write-Progress -Activity 'Retrieving Content Manager documents' -Status $activity -PercentComplete $pct
        }
        else {
            Write-Progress -Activity 'Retrieving Content Manager documents' -Status $activity
        }

        if ($stopEarly) {
            Write-TrimLog 'Reached a record older than -MaxAgeDays in a newest-first sort; stopping.'
            break
        }
        if ($downloaded -ge $MaxFiles) {
            Write-TrimLog "Reached -MaxFiles ($MaxFiles); stopping."
            break
        }
    }

    Write-Progress -Activity 'Retrieving Content Manager documents' -Completed

    if ($rows.Count -gt 0) {
        if ($WhatIfPreference) {
            Write-Host "What if: would write manifest with $($rows.Count) row(s) to $ManifestPath"
        }
        else {
            $rows | Export-Csv -LiteralPath $ManifestPath -NoTypeInformation -Encoding UTF8
            Write-TrimLog "Manifest written: $ManifestPath"
        }
    }
    else {
        Write-TrimLog 'Search returned no records; no manifest written.' 'WARN'
    }

    $skipCounts = @{}
    foreach ($r in $rows) {
        if ($r.Status -like 'Skipped*') {
            if (-not $skipCounts.ContainsKey($r.Status)) { $skipCounts[$r.Status] = 0 }
            $skipCounts[$r.Status]++
        }
    }

    $summary = [pscustomobject]@{
        SearchString   = $effectiveSearch
        Matched        = $total
        Examined       = $scanned
        Downloaded     = $downloaded
        Failed         = $failed
        Skipped        = ($skipCounts.Values | Measure-Object -Sum).Sum
        SkippedDetail  = ($skipCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
        OutputFolder   = $resolvedFolder
        ManifestPath   = $ManifestPath
    }

    Write-Host ''
    Write-Host 'Content Manager retrieval summary' -ForegroundColor Cyan
    $summary | Format-List | Out-String | Write-Host

    if ($failed -gt 0) {
        Write-Host "$failed record(s) matched but could not be retrieved - see the Error column in $ManifestPath" -ForegroundColor Yellow
    }

    $summary
}
catch {
    Write-TrimLog $_.Exception.Message 'ERROR'
    throw
}
finally {
    if ($db) {
        try { $db.Dispose(); Write-TrimLog 'Disconnected.' } catch { }
    }
    if ($script:LogFile) { Write-TrimLog '=== Invoke-TrimSearchDownload finished ===' }
}

#endregion
