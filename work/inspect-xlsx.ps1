param(
  [Parameter(Mandatory=$true)][string]$Path,
  [string]$Sheets = "",
  [int]$MaxRows = 30
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText($zip, [string]$entryName) {
  $entry = $zip.GetEntry($entryName)
  if (-not $entry) { return $null }
  $stream = $entry.Open()
  try {
    $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally { $stream.Dispose() }
}

function Col-Index([string]$ref) {
  $letters = ([regex]::Match($ref, '^[A-Z]+')).Value
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n - 1
}

$zip = [IO.Compression.ZipFile]::OpenRead($Path)
try {
  [xml]$workbook = Read-ZipText $zip 'xl/workbook.xml'
  [xml]$rels = Read-ZipText $zip 'xl/_rels/workbook.xml.rels'
  $relMap = @{}
  foreach ($rel in $rels.Relationships.Relationship) {
    $target = [string]$rel.Target
    if (-not $target.StartsWith('xl/')) { $target = 'xl/' + $target.TrimStart('/') }
    $relMap[[string]$rel.Id] = $target
  }

  $shared = @()
  [xml]$sst = Read-ZipText $zip 'xl/sharedStrings.xml'
  if ($sst) {
    foreach ($si in $sst.sst.si) {
      $parts = @()
      if ($si.t) { $parts += [string]$si.t }
      foreach ($r in $si.r) { if ($r.t) { $parts += [string]$r.t } }
      $shared += ($parts -join '')
    }
  }

  $sheetFilter = @()
  if ($Sheets) {
    $sheetFilter = $Sheets.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
  }

  foreach ($sheet in $workbook.workbook.sheets.sheet) {
    $name = [string]$sheet.name
    if ($sheetFilter.Count -gt 0 -and $name -notin $sheetFilter) { continue }
    $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $entryName = $relMap[$rid]
    [xml]$xml = Read-ZipText $zip $entryName
    Write-Output "### SHEET: $name ($entryName)"
    $count = 0
    foreach ($row in $xml.worksheet.sheetData.row) {
      if ($count -ge $MaxRows) { break }
      $values = @{}
      $maxCol = 0
      foreach ($c in $row.c) {
        $idx = Col-Index([string]$c.r)
        $maxCol = [Math]::Max($maxCol, $idx)
        $value = ''
        if ($c.t -eq 's') { $value = $shared[[int]$c.v] }
        elseif ($c.t -eq 'inlineStr') { $value = [string]$c.is.t }
        else { $value = [string]$c.v }
        $values[$idx] = $value
      }
      $line = for ($i = 0; $i -le $maxCol; $i++) {
        if ($values.ContainsKey($i)) { $values[$i] } else { '' }
      }
      Write-Output (($line) -join ' | ')
      $count += 1
    }
    Write-Output ''
  }
} finally {
  $zip.Dispose()
}
