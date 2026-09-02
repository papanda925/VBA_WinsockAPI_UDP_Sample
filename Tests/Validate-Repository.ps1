$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$sourcePath = Join-Path $repositoryRoot 'VBA_WinsockAPI_UDP_Sample.bas'
$readmePath = Join-Path $repositoryRoot 'README.md'

[System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
$source = [System.Text.Encoding]::GetEncoding(932).GetString(
    [System.IO.File]::ReadAllBytes($sourcePath)
)
$readme = [System.IO.File]::ReadAllText($readmePath)
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-Match {
    param(
        [Parameter(Mandatory)]
        [string]$Text,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if ($Text -notmatch $Pattern) {
        $failures.Add($Message)
    }
}

function Assert-NotMatch {
    param(
        [Parameter(Mandatory)]
        [string]$Text,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if ($Text -match $Pattern) {
        $failures.Add($Message)
    }
}

Assert-Match $source '(?im)^Option Explicit\s*$' 'Option Explicit is required.'
Assert-Match $source '(?im)Function SOCKET .+ As LongPtr\s*$' 'socket must return LongPtr.'
Assert-Match $source '(?im)Function closesocket .+LongPtr.+ As Long\s*$' 'closesocket must receive a LongPtr handle.'
Assert-Match $source '(?im)Function recvfrom .+ByRef fromAddrSize As Long\) As Long\s*$' 'recvfrom must receive the address length by reference.'
Assert-NotMatch $source '(?i)wsock32\.dll' 'Use ws2_32.dll consistently.'
Assert-NotMatch $source '(?im)^\s*(Application\.Quit|ThisWorkbook\.Close)\s*$' "The sample must not close the user's Excel session."
Assert-Match $readme '(?m)^## 使い方\s*$' 'README must contain setup and usage instructions.'
Assert-Match $readme '(?m)^## 制限事項\s*$' 'README must describe limitations.'
Assert-Match $readme '(?m)^## トラブルシューティング\s*$' 'README must contain troubleshooting guidance.'

$subStarts = ([regex]::Matches($source, '(?im)^\s*(Public\s+|Private\s+)?Sub\s+\w+')).Count
$subEnds = ([regex]::Matches($source, '(?im)^\s*End Sub\s*$')).Count
$functionStarts = ([regex]::Matches($source, '(?im)^\s*(Public\s+|Private\s+)?Function\s+\w+')).Count
$functionEnds = ([regex]::Matches($source, '(?im)^\s*End Function\s*$')).Count

if ($subStarts -ne $subEnds) {
    $failures.Add("Sub/End Sub count mismatch: $subStarts/$subEnds")
}
if ($functionStarts -ne $functionEnds) {
    $failures.Add("Function/End Function count mismatch: $functionStarts/$functionEnds")
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Host 'Repository validation passed.'
