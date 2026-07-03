param(
  [Parameter(Mandatory=$true)][string]$Path,
  [Parameter(Mandatory=$true)][string]$OutFile
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

function Col-Name([int]$index) {
  $name = ""
  $n = $index + 1
  while ($n -gt 0) {
    $rem = ($n - 1) % 26
    $name = [char]([int]([int][char]'A' + $rem)) + $name
    $n = [int][math]::Floor(($n - 1) / 26)
  }
  return $name
}

function Ref-Max-Col([string]$ref) {
  if (-not $ref) { return -1 }
  $parts = $ref.Split(':')
  $last = $parts[$parts.Count - 1]
  return Col-Index $last
}

function Cell-Value($cell, $shared) {
  $type = [string]$cell.t
  $raw = [string]$cell.v
  if ($type -eq 's') {
    $idx = 0
    if ([int]::TryParse($raw, [ref]$idx) -and $idx -lt $shared.Count) { return $shared[$idx] }
    return ""
  }
  if ($type -eq 'inlineStr') {
    $parts = @()
    if ($cell.is.t) { $parts += [string]$cell.is.t }
    foreach ($r in $cell.is.r) { if ($r.t) { $parts += [string]$r.t } }
    return ($parts -join '')
  }
  if ($type -eq 'b') {
    if ($raw -eq '1') { return "TRUE" }
    return "FALSE"
  }
  return $raw
}

function Sheet-Rows($xml, $shared, [int]$sheetMaxCol) {
  $rows = @()
  foreach ($rowNode in $xml.worksheet.sheetData.row) {
    $cells = @()
    $cellMap = @{}
    foreach ($cell in $rowNode.c) {
      $ref = [string]$cell.r
      $idx = Col-Index $ref
      $value = Cell-Value $cell $shared
      $formula = if ($cell.f) { [string]$cell.f.InnerText } else { "" }
      $cellMap[$idx] = [ordered]@{
        ref = $ref
        column = Col-Name $idx
        value = $value
        formula = $formula
      }
    }
    for ($i = 0; $i -le $sheetMaxCol; $i++) {
      if ($cellMap.ContainsKey($i)) {
        $cells += $cellMap[$i]
      } else {
        $cells += [ordered]@{ ref = "$(Col-Name $i)$($rowNode.r)"; column = Col-Name $i; value = ""; formula = "" }
      }
    }
    $rows += [ordered]@{
      number = [int]$rowNode.r
      cells = $cells
    }
  }
  return @($rows)
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

  $sheets = @()
  foreach ($sheet in $workbook.workbook.sheets.sheet) {
    $name = [string]$sheet.name
    $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $entryName = $relMap[$rid]
    [xml]$sheetXml = Read-ZipText $zip $entryName
    $sheetMaxCol = Ref-Max-Col ([string]$sheetXml.worksheet.dimension.ref)
    foreach ($rowNode in $sheetXml.worksheet.sheetData.row) {
      foreach ($cell in $rowNode.c) {
        $sheetMaxCol = [Math]::Max($sheetMaxCol, (Col-Index ([string]$cell.r)))
      }
    }
    [object[]]$rows = if ($sheetMaxCol -ge 0) { @(Sheet-Rows $sheetXml $shared $sheetMaxCol) } else { @() }
    $nonEmpty = 0
    $formulaCount = 0
    foreach ($row in $rows) {
      foreach ($cell in $row.cells) {
        if ([string]$cell.value -ne "" -or [string]$cell.formula -ne "") { $nonEmpty += 1 }
        if ([string]$cell.formula -ne "") { $formulaCount += 1 }
      }
    }
    $sheets += [ordered]@{
      name = $name
      sheetId = [string]$sheet.sheetId
      path = $entryName
      columnCount = if ($sheetMaxCol -ge 0) { $sheetMaxCol + 1 } else { 0 }
      importedColumnCount = if ($sheetMaxCol -ge 0) { $sheetMaxCol + 1 } else { 0 }
      rows = $rows
      nonEmptyCellCount = $nonEmpty
      formulaCount = $formulaCount
    }
  }

  $payload = [ordered]@{
    sourceFile = [IO.Path]::GetFileName($Path)
    generatedAt = (Get-Date).ToString("s")
    sheetCount = $sheets.Count
    sheets = $sheets
  }
  $json = $payload | ConvertTo-Json -Depth 100 -Compress
  $content = "window.EMBEDDED_WORKBOOK = $json;"
  [IO.File]::WriteAllText($OutFile, $content, [Text.UTF8Encoding]::new($false))
  Write-Output "Exported $($sheets.Count) sheets to $OutFile"
} finally {
  $zip.Dispose()
}
