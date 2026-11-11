# custodump.ps1 (10.11.2026): vollstaendiger Auszug von custobase.mdb (alle Tabellen als CSV, dazu der Aufbau
# der Tabellen) und eine Liste aller Dateien unter C:\Datenbanken (labor3), nur lesend.
# 11.11.2026: zusaetzlich alle alten Staende: custobase*.mdb* auf labor3 (Ordner L3_<Dateiname>) und die aus den
# Monatsarchiven gezogenen in P:\Datenbanken\custodump\alt (Ordner alt_<Dateiname>); schon vorhandene Ordner
# werden uebersprungen.
# Ergebnis: P:\Datenbanken\custodump\  (<Ordner>\*.csv, <Ordner>_aufbau.txt, dateien_C.csv)
# Enthaelt Patientendaten, bleibt also auf P: (Patientendokumente) und wird nach der Auswertung geloescht.
# Aufruf auf labor3, mit der 32-Bit-PowerShell:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File U:\Programmierung\Custo\custodump.ps1
param(
  [string]$Aus = '\\linux1\daten\Patientendokumente\Datenbanken\custodump'
)
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { Write-Host 'Bitte mit der 32-Bit-PowerShell starten (SysWOW64), siehe Kopf der Datei.'; exit 1 }
Add-Type -AssemblyName System.Data
New-Item -ItemType Directory -Force -Path $Aus | Out-Null
function W($x) {
  if ($x -is [DBNull] -or $null -eq $x) { return '' }
  if ($x -is [byte[]]) { return ($x | ForEach-Object { $_.ToString('x2') }) -join '' }
  if ($x -is [datetime]) { return $x.ToString('yyyy-MM-dd HH:mm:ss') }
  return ([string]$x) -replace '"', '""' -replace "`r?`n", ' | '
}
$liste = [ordered]@{ 'C' = 'C:\Datenbanken\custobase.mdb' }
Get-ChildItem 'C:\Datenbanken' -Filter 'custobase*.mdb*' -File | Where-Object { $_.Name -ne 'custobase.mdb' } | ForEach-Object { $liste['L3_' + ($_.Name -replace '[^A-Za-z0-9_.-]', '_')] = $_.FullName }
Get-ChildItem (Join-Path $Aus 'alt') -Filter '*.mdb' -File | ForEach-Object { $liste['alt_' + $_.BaseName] = $_.FullName }
foreach ($name in $liste.Keys) {
  $pfad = $liste[$name]
  if ($name -ne 'C' -and (Test-Path (Join-Path $Aus ($name + '_aufbau.txt')))) { continue }
  $ziel = Join-Path $Aus $name
  New-Item -ItemType Directory -Force -Path $ziel | Out-Null
  $con = New-Object System.Data.OleDb.OleDbConnection "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$pfad;Mode=Read;Jet OLEDB:Database Password=CDS"
  try { $con.Open() } catch { Write-Host "$name nicht zu oeffnen: $($_.Exception.Message)"; continue }
  $aufbau = New-Object System.Collections.Generic.List[string]
  $tabs = $con.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null) | Where-Object { $_.TABLE_TYPE -eq 'TABLE' } | ForEach-Object { $_.TABLE_NAME }
  $spalten = $con.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns, $null)
  $idx = $con.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Indexes, $null)
  foreach ($t in $tabs) {
    $da = New-Object System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [$t]", $con)
    $dt = New-Object System.Data.DataTable
    try { [void]$da.Fill($dt) } catch { $aufbau.Add("== $t : FEHLER $($_.Exception.Message)"); continue }
    $aufbau.Add("== $t ($($dt.Rows.Count) Zeilen)")
    $spalten | Where-Object { $_.TABLE_NAME -eq $t } | Sort-Object ORDINAL_POSITION | ForEach-Object {
      $aufbau.Add(("   {0,-32} Typ {1,-4} Laenge {2,-6} Null {3}" -f $_.COLUMN_NAME, $_.DATA_TYPE, $_.CHARACTER_MAXIMUM_LENGTH, $_.IS_NULLABLE)) }
    $idx | Where-Object { $_.TABLE_NAME -eq $t } | ForEach-Object {
      $aufbau.Add(("   Index {0}: {1} primaer={2} eindeutig={3}" -f $_.INDEX_NAME, $_.COLUMN_NAME, $_.PRIMARY_KEY, $_.UNIQUE)) }
    $zeilen = New-Object System.Collections.Generic.List[string]
    $zeilen.Add((($dt.Columns | ForEach-Object { '"' + $_.ColumnName + '"' }) -join ';'))
    foreach ($r in $dt.Rows) { $zeilen.Add((($r.ItemArray | ForEach-Object { '"' + (W $_) + '"' }) -join ';')) }
    [System.IO.File]::WriteAllLines((Join-Path $ziel ($t + '.csv')), $zeilen, (New-Object System.Text.UTF8Encoding $false))
    Write-Host "$name $t $($dt.Rows.Count)"
  }
  $con.Close()
  [System.IO.File]::WriteAllLines((Join-Path $Aus ($name + '_aufbau.txt')), $aufbau, (New-Object System.Text.UTF8Encoding $false))
}
$basis = 'C:\Datenbanken\'
$liste = New-Object System.Collections.Generic.List[string]
$liste.Add('"Pfad";"Bytes";"Geaendert";"Erstellt"')
Get-ChildItem -Path $basis -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
  $liste.Add(('"{0}";"{1}";"{2}";"{3}"' -f $_.FullName.Substring($basis.Length), $_.Length, $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'), $_.CreationTime.ToString('yyyy-MM-dd HH:mm:ss'))) }
[System.IO.File]::WriteAllLines((Join-Path $Aus 'dateien_C.csv'), $liste, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Fertig: $Aus"
