# custovergleich.ps1 (10.11.2026): vergleicht C:\Datenbanken\custobase.mdb (labor3) mit
# P:\Datenbanken\custobase.mdb. Nur lesend (Jet OLEDB 4.0, Mode=Read, Share Deny None), daher auch bei
# laufendem custo. Gibt keine Namen aus, nur Nummern, Datumsangaben und Anzahlen.
# Ergebnis: P:\Datenbanken\custovergleich.txt (und auf dem Bildschirm).
# Aufruf auf labor3, unbedingt mit der 32-Bit-PowerShell (Jet gibt es nur 32-bittig):
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File U:\Programmierung\Custo\custovergleich.ps1
param(
  [string]$C = 'C:\Datenbanken\custobase.mdb',
  [string]$P = '\\linux1\daten\Patientendokumente\Datenbanken\custobase.mdb',
  [string]$Aus = '\\linux1\daten\Patientendokumente\Datenbanken\custovergleich.txt'
)
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { Write-Host 'Bitte mit der 32-Bit-PowerShell starten (SysWOW64), siehe Kopf der Datei.'; exit 1 }
Add-Type -AssemblyName System.Data
$out = New-Object System.Collections.Generic.List[string]
function Zeile($s) { $out.Add($s); Write-Host $s }

# Verbindungsvarianten: die erste, die sich oeffnen laesst, wird fuer alle weiteren Abfragen behalten
$script:csOk = @{}
function Oeffne($pfad) {
  $basis = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$pfad"
  $var = @("$basis;Mode=Read;Jet OLEDB:Database Password=CDS", "$basis;Mode=Share Deny None;Jet OLEDB:Database Password=CDS",
           "$basis;Jet OLEDB:Database Password=CDS", "$basis;Mode=Read")
  if ($script:csOk.ContainsKey($pfad)) { $var = @($script:csOk[$pfad]) }
  $fehler = @()
  foreach ($cs in $var) {
    $con = New-Object System.Data.OleDb.OleDbConnection $cs
    try { $con.Open(); if (-not $script:csOk.ContainsKey($pfad)) { $script:csOk[$pfad] = $cs; Zeile ("   geoeffnet mit: " + ($cs -replace '^.*?mdb;?', '')) }; return ,$con }
    catch { $fehler += (($cs -replace '^.*?mdb;?', '') + ' -> ' + $_.Exception.Message) }
  }
  $fehler | ForEach-Object { Zeile "   Fehler: $_" }
  throw "custobase.mdb nicht zu oeffnen: $pfad"
}
function Lies($pfad, $sql) {
  $con = Oeffne $pfad
  try {
    $da = New-Object System.Data.OleDb.OleDbDataAdapter($sql, $con)
    $t = New-Object System.Data.DataTable
    [void]$da.Fill($t)
    return ,$t
  } finally { $con.Close() }
}
function D($x) { if ($x -is [DBNull] -or $null -eq $x) { '' } else { ([datetime]$x).ToString('yyyy-MM-dd HH:mm') } }

$db = @{}
foreach ($name in 'C', 'P') {
  $pfad = Get-Variable -Name $name -ValueOnly
  $fi = Get-Item $pfad
  Zeile ("== {0}: {1}  ({2} Bytes, geaendert {3})" -f $name, $pfad, $fi.Length, $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))
  $A = Lies $pfad 'SELECT AuswertungsNummer, PatientenNummer, Auswertungsdatum, Auswertungstyp, Auswertungsstatus FROM Auswertungsdaten'
  $Pt = Lies $pfad 'SELECT PatientenNummer, Anlagedatum, PraxEdvNum FROM Patientenstammdaten'
  $E = Lies $pfad 'SELECT Nummer, Typ, SubTyp, DatumBeginn, DatumEnde, BenutzerNummer, ProjektNummer, ErrorCode FROM Ereignisse'
  $V = Lies $pfad 'SELECT Versionsnummer, UpdateDatum FROM Versionen'
  $S = Lies $pfad 'SELECT * FROM Systemdaten'
  $ad = @($A.Rows | ForEach-Object { $_.Auswertungsdatum } | Where-Object { $_ -isnot [DBNull] } | Sort-Object)
  Zeile ("   Auswertungen: {0}, Nummern {1}..{2}, Datum {3} .. {4}" -f $A.Rows.Count,
    ($A.Rows | Measure-Object AuswertungsNummer -Minimum).Minimum, ($A.Rows | Measure-Object AuswertungsNummer -Maximum).Maximum,
    (D $ad[0]), (D $ad[-1]))
  Zeile ("   Patienten:    {0}, hoechste Nummer {1}, letzte Anlage {2}" -f $Pt.Rows.Count,
    ($Pt.Rows | Measure-Object PatientenNummer -Maximum).Maximum,
    (D ($Pt.Rows | ForEach-Object { $_.Anlagedatum } | Where-Object { $_ -isnot [DBNull] } | Sort-Object | Select-Object -Last 1)))
  Zeile ('   Auswertungen je Jahr: ' + (($ad | Group-Object { ([datetime]$_).Year } | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ' '))
  Zeile ('   Auswertungen je Typ:  ' + (($A.Rows | Group-Object Auswertungstyp | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ' '))
  Zeile ('   Versionen: ' + (($V.Rows | ForEach-Object { "$($_.Versionsnummer) ($(D $_.UpdateDatum))" }) -join ', '))
  Zeile ('   Systemdaten: ' + (($S.Columns | ForEach-Object { $c = $_.ColumnName; $w = $S.Rows[0][$c]; if ($w -is [datetime]) { $w = D $w }; "$c=$w" }) -join '; '))
  Zeile ("   Ereignisse ({0}):" -f $E.Rows.Count)
  $E.Rows | Sort-Object DatumBeginn | ForEach-Object {
    Zeile ("     {0}  bis {1}  Typ {2}/{3}  Benutzer {4}  Projekt {5}  Fehler {6}" -f (D $_.DatumBeginn), (D $_.DatumEnde), $_.Typ, $_.SubTyp, $_.BenutzerNummer, $_.ProjektNummer, $_.ErrorCode)
  }
  $h = @{}; foreach ($r in $A.Rows) { $h[[string]$r.AuswertungsNummer] = "{0}|{1}|{2}|{3}" -f $r.PatientenNummer, (D $r.Auswertungsdatum), $r.Auswertungstyp, $r.Auswertungsstatus }
  $hp = @{}; foreach ($r in $Pt.Rows) { $hp[[string]$r.PatientenNummer] = "{0}|{1}" -f (D $r.Anlagedatum), $r.PraxEdvNum }
  $db[$name] = @{ A = $h; P = $hp }
}

Zeile ''
Zeile '== Vergleich der Auswertungen (Nummer: Patient|Datum|Typ|Status)'
$nurC = @($db.C.A.Keys | Where-Object { -not $db.P.A.ContainsKey($_) } | Sort-Object { [int]$_ })
$nurP = @($db.P.A.Keys | Where-Object { -not $db.C.A.ContainsKey($_) } | Sort-Object { [int]$_ })
$versch = @($db.C.A.Keys | Where-Object { $db.P.A.ContainsKey($_) -and $db.P.A[$_] -ne $db.C.A[$_] } | Sort-Object { [int]$_ })
Zeile ("   nur auf C (labor3): {0}" -f $nurC.Count); $nurC | ForEach-Object { Zeile ("     {0}: {1}" -f $_, $db.C.A[$_]) }
Zeile ("   nur auf P:          {0}" -f $nurP.Count); $nurP | ForEach-Object { Zeile ("     {0}: {1}" -f $_, $db.P.A[$_]) }
Zeile ("   gleiche Nummer, verschiedener Inhalt: {0}" -f $versch.Count); $versch | ForEach-Object { Zeile ("     {0}: C {1}  /  P {2}" -f $_, $db.C.A[$_], $db.P.A[$_]) }
Zeile ''
Zeile '== Vergleich der Patienten (Nummer: Anlage|PraxEdvNum)'
$nurC = @($db.C.P.Keys | Where-Object { -not $db.P.P.ContainsKey($_) } | Sort-Object { [int]$_ })
$nurP = @($db.P.P.Keys | Where-Object { -not $db.C.P.ContainsKey($_) } | Sort-Object { [int]$_ })
$versch = @($db.C.P.Keys | Where-Object { $db.P.P.ContainsKey($_) -and $db.P.P[$_] -ne $db.C.P[$_] } | Sort-Object { [int]$_ })
Zeile ("   nur auf C (labor3): {0}" -f $nurC.Count); $nurC | ForEach-Object { Zeile ("     {0}: {1}" -f $_, $db.C.P[$_]) }
Zeile ("   nur auf P:          {0}" -f $nurP.Count); $nurP | ForEach-Object { Zeile ("     {0}: {1}" -f $_, $db.P.P[$_]) }
Zeile ("   gleiche Nummer, verschiedener Inhalt: {0}" -f $versch.Count); $versch | ForEach-Object { Zeile ("     {0}: C {1}  /  P {2}" -f $_, $db.C.P[$_], $db.P.P[$_]) }

$out | Set-Content -Path $Aus -Encoding UTF8
Write-Host "Geschrieben: $Aus"
