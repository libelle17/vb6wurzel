# custoneuaufbau.ps1 (11.11.2026): baut die durch das fruehere Zuruecksetzen verlorenen oder falsch zugeordneten
# custo-Eintraege neu auf, nach den auf linux1 erzeugten Listen in P:\Datenbanken\custodump\neuaufbau\:
#   patienten_neu.csv        fehlende Patienten (Angaben aus den Koepfen der Messdateien), Nummern ab 2001
#   auswertungen_neu.csv     fehlende Auswertungen, Nummern ab 3001 (Abstand zu den laufend vergebenen)
#   auswertungen_loeschen.csv Eintraege, deren eigene Messdatei durch die einer anderen Person ueberschrieben ist
#   dateien.csv              kopie (Messdatei unter neuem Namen), zurueck (richtige Datei von P:), weg (alter Name)
#   umlaute_alt.csv          Umlaute bestehender Patienten: DOS-Schreibweise (z.B. „ statt ae) -> richtige Zeichen;
#                            nur wenn das Feld noch genau den alten Wert hat
# Aufruf auf labor3 mit der 32-Bit-PowerShell:
#   Test (Standard): baut P:\Datenbanken\custotest\ aus einer Kopie, C:\Datenbanken bleibt unberuehrt
#     C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File V:\custoneuaufbau.ps1
#   Echt: aendert C:\Datenbanken; nur wenn custo ueberall geschlossen ist (keine custobase.ldb); sichert vorher
#   custobase.mdb als custobase_vor_neuaufbau_<Datum>.mdb und verschiebt alte Dateinamen nach C:\Datenbanken\vor_neuaufbau\
#     ... -File V:\custoneuaufbau.ps1 -Echt
# 11.11.26: -Quelle \\labor3\Datenbanken erlaubt den Aufruf auch von einem anderen PC (z.B. szn4), in der eigenen Sitzung
param([switch]$Echt, [string]$Quelle = 'C:\Datenbanken')
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { Write-Host 'Bitte mit der 32-Bit-PowerShell starten (SysWOW64), siehe Kopf der Datei.'; exit 1 }
Add-Type -AssemblyName System.Data
$L = '\\linux1\daten\Patientendokumente\Datenbanken\custodump\neuaufbau'
$quelle = $Quelle.TrimEnd('\')
$ziel = if ($Echt) { $quelle } else { '\\linux1\daten\Patientendokumente\Datenbanken\custotest' }
$log = Join-Path $L ($(if ($Echt) { 'echt' } else { 'test' }) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
function Prot($s) { Write-Host $s; Add-Content -Path $log -Value $s -Encoding UTF8 }
function Hex($s) { $b = New-Object byte[] ([int]($s.Length / 2)); for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = [Convert]::ToByte($s.Substring(2 * $i, 2), 16) }; return ,$b }
function Dat($s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return [DBNull]::Value }
  return [datetime]::ParseExact($s.Trim(), [string[]]@('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd'), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None)
}
$pat = Import-Csv (Join-Path $L 'patienten_neu.csv') -Delimiter ';' -Encoding UTF8
$aus = Import-Csv (Join-Path $L 'auswertungen_neu.csv') -Delimiter ';' -Encoding UTF8
$weg = Import-Csv (Join-Path $L 'auswertungen_loeschen.csv') -Delimiter ';' -Encoding UTF8
$dat = Import-Csv (Join-Path $L 'dateien.csv') -Delimiter ';' -Encoding UTF8
$korr = @(Import-Csv (Join-Path $L 'umlaute_alt.csv') -Delimiter ';' -Encoding UTF8)
Prot ("{0}: {1} neue Patienten, {2} neue Auswertungen, {3} zu loeschen, {4} Dateiaktionen; Ziel {5}" -f (Get-Date), $pat.Count, $aus.Count, $weg.Count, $dat.Count, $ziel)

if ($Echt) {
  if (Test-Path "$quelle\custobase.ldb") { Prot 'custobase.ldb vorhanden: custo ist noch irgendwo offen. Abbruch, nichts geaendert.'; exit 1 }
  $sich = "$quelle\custobase_vor_neuaufbau_$(Get-Date -Format 'yyyyMMdd_HHmmss').mdb"
  Copy-Item "$quelle\custobase.mdb" $sich
  Prot "Sicherung: $sich"
} else {
  if (Test-Path $ziel) { Prot "$ziel gibt es schon; bitte vorher loeschen oder umbenennen. Abbruch."; exit 1 }
  New-Item -ItemType Directory -Path $ziel | Out-Null
  Copy-Item "$quelle\custobase.mdb" $ziel
  Get-ChildItem $quelle -Filter 'LizenzKey*.txt' | Copy-Item -Destination $ziel
  foreach ($o in 'Ekg', 'Blutdruck', 'LuFu', 'lzekg', 'ulufu', 'ublutdruck') {
    if (Test-Path "$quelle\$o") { Copy-Item "$quelle\$o" (Join-Path $ziel $o) -Recurse }
  }
  Prot 'Testkopie von custobase.mdb und den Messordnern angelegt.'
}

# Datenbank: alles in einer Transaktion, vorher pruefen, dass die Nummern frei und die zu loeschenden Eintraege unveraendert sind
$con = New-Object System.Data.OleDb.OleDbConnection "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$ziel\custobase.mdb;Jet OLEDB:Database Password=CDS"
$con.Open()
function Zahl($sql) { $c = $con.CreateCommand(); $c.CommandText = $sql; if ($script:tr) { $c.Transaction = $script:tr }; return $c.ExecuteScalar() }
$maxA = Zahl 'SELECT MAX(AuswertungsNummer) FROM Auswertungsdaten'
$maxP = Zahl 'SELECT MAX(PatientenNummer) FROM Patientenstammdaten'
Prot "Hoechste Nummern jetzt: Auswertung $maxA, Patient $maxP"
$minNeuA = ($aus | ForEach-Object { [int]$_.AuswertungsNummer } | Measure-Object -Minimum).Minimum
$minNeuP = ($pat | ForEach-Object { [int]$_.PatientenNummer } | Measure-Object -Minimum).Minimum
if ($maxA -ge $minNeuA -or $maxP -ge $minNeuP) { Prot "Neue Nummern nicht mehr frei (ab $minNeuA / $minNeuP). Abbruch, nichts geaendert."; $con.Close(); exit 1 }
foreach ($w in $weg) {
  $n = Zahl ("SELECT COUNT(*) FROM Auswertungsdaten WHERE AuswertungsNummer=$([int]$w.AuswertungsNummer) AND PatientenNummer=$([int]$w.PatientenNummer)")
  if ($n -ne 1) { Prot "Zu loeschender Eintrag $($w.AuswertungsNummer)/$($w.PatientenNummer) nicht (mehr) vorhanden. Abbruch, nichts geaendert."; $con.Close(); exit 1 }
}
$script:tr = $con.BeginTransaction()
try {
  foreach ($w in $weg) {
    foreach ($t in 'Auswertungsdaten', 'AuswStatusList') {
      $c = $con.CreateCommand(); $c.Transaction = $tr
      $c.CommandText = "DELETE FROM $t WHERE AuswertungsNummer=$([int]$w.AuswertungsNummer) AND PatientenNummer=$([int]$w.PatientenNummer)"
      [void]$c.ExecuteNonQuery()
    }
  }
  foreach ($p in $pat) {
    $c = $con.CreateCommand(); $c.Transaction = $tr
    $c.CommandText = 'INSERT INTO Patientenstammdaten (Name, PraxEdvNum, Vorname, Geburtsdatum, Geschlecht, Ethnie, Patientenflags, Anlageart, Anlagedatum, PatientenNummer, Strasse, [PLZ/Ort], Titel, Namenszusatz, Sozialversicherungsnummer, sonstigeInfos, Gewicht, Groesse) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
    $c.Parameters.Add('n', [System.Data.OleDb.OleDbType]::VarWChar).Value = $p.Name
    $c.Parameters.Add('e', [System.Data.OleDb.OleDbType]::VarWChar).Value = $(if ($p.PraxEdvNum) { $p.PraxEdvNum } else { [DBNull]::Value })
    $c.Parameters.Add('v', [System.Data.OleDb.OleDbType]::VarWChar).Value = $p.Vorname
    $c.Parameters.Add('g', [System.Data.OleDb.OleDbType]::Date).Value = (Dat $p.Geburtsdatum)
    $c.Parameters.Add('s', [System.Data.OleDb.OleDbType]::UnsignedTinyInt).Value = [byte]$p.Geschlecht
    $c.Parameters.Add('t', [System.Data.OleDb.OleDbType]::UnsignedTinyInt).Value = [byte]0
    $c.Parameters.Add('f', [System.Data.OleDb.OleDbType]::Binary).Value = (Hex '00000000000000000000')
    $c.Parameters.Add('a', [System.Data.OleDb.OleDbType]::Integer).Value = 0
    $c.Parameters.Add('d', [System.Data.OleDb.OleDbType]::Date).Value = (Dat $p.Anlagedatum)
    $c.Parameters.Add('p', [System.Data.OleDb.OleDbType]::Integer).Value = [int]$p.PatientenNummer
    $c.Parameters.Add('st', [System.Data.OleDb.OleDbType]::VarWChar).Value = $p.Strasse
    $c.Parameters.Add('o', [System.Data.OleDb.OleDbType]::VarWChar).Value = $p.PLZOrt
    # 11.11.26: wie bei von custo selbst angelegten Patienten keine NULL-Felder, sonst verrutschen in custo die
    # Namensfelder und die Suche nach dem Nachnamen findet den Patienten nicht
    $c.Parameters.Add('ti', [System.Data.OleDb.OleDbType]::VarWChar).Value = ''
    $c.Parameters.Add('nz', [System.Data.OleDb.OleDbType]::VarWChar).Value = ''
    $c.Parameters.Add('sv', [System.Data.OleDb.OleDbType]::VarWChar).Value = '-'
    $c.Parameters.Add('si', [System.Data.OleDb.OleDbType]::VarWChar).Value = ''
    $c.Parameters.Add('gw', [System.Data.OleDb.OleDbType]::Single).Value = [single]0
    $c.Parameters.Add('gr', [System.Data.OleDb.OleDbType]::Single).Value = [single]0
    [void]$c.ExecuteNonQuery()
  }
  foreach ($a in $aus) {
    $c = $con.CreateCommand(); $c.Transaction = $tr
    $c.CommandText = 'INSERT INTO Auswertungsdaten (PatientenNummer, AuswertungsBezeichnung, Auswertungsdatum, Auswertungstyp, Auswertungsstatus, AuswertungsNummer) VALUES (?,?,?,?,?,?)'
    $c.Parameters.Add('p', [System.Data.OleDb.OleDbType]::Integer).Value = [int]$a.PatientenNummer
    $c.Parameters.Add('b', [System.Data.OleDb.OleDbType]::VarWChar).Value = '???'
    $c.Parameters.Add('d', [System.Data.OleDb.OleDbType]::Date).Value = (Dat $a.Auswertungsdatum)
    $c.Parameters.Add('t', [System.Data.OleDb.OleDbType]::Binary).Value = (Hex $a.Typ)
    $c.Parameters.Add('s', [System.Data.OleDb.OleDbType]::Binary).Value = (Hex $a.Status)
    $c.Parameters.Add('n', [System.Data.OleDb.OleDbType]::Integer).Value = [int]$a.AuswertungsNummer
    [void]$c.ExecuteNonQuery()
    $c = $con.CreateCommand(); $c.Transaction = $tr
    $c.CommandText = 'INSERT INTO AuswStatusList (AuswertungsNummer, PatientenNummer, Auswertungsstatus, Befundung) VALUES (?,?,?,?)'
    $c.Parameters.Add('n', [System.Data.OleDb.OleDbType]::Integer).Value = [int]$a.AuswertungsNummer
    $c.Parameters.Add('p', [System.Data.OleDb.OleDbType]::Integer).Value = [int]$a.PatientenNummer
    $c.Parameters.Add('s', [System.Data.OleDb.OleDbType]::Binary).Value = (Hex $a.Status)
    $c.Parameters.Add('d', [System.Data.OleDb.OleDbType]::Date).Value = (Dat $a.Auswertungsdatum)
    [void]$c.ExecuteNonQuery()
  }
  $nk = 0; $nu = 0
  foreach ($k in $korr) {
    $f = $k.Feld; $nr = [int]$k.PatientenNummer
    $c = $con.CreateCommand(); $c.Transaction = $tr
    $c.CommandText = "SELECT [$f] FROM Patientenstammdaten WHERE PatientenNummer=$nr"
    $ist = $c.ExecuteScalar()
    if ($ist -is [string] -and $ist -ceq $k.Alt) {
      $c = $con.CreateCommand(); $c.Transaction = $tr
      $c.CommandText = "UPDATE Patientenstammdaten SET [$f]=? WHERE PatientenNummer=$nr"
      $c.Parameters.Add('w', [System.Data.OleDb.OleDbType]::VarWChar).Value = $k.Neu
      [void]$c.ExecuteNonQuery(); $nk++
    } else { $nu++ }
  }
  $tr.Commit()
  Prot ("Datenbank: {0} geloescht, {1} Patienten und {2} Auswertungen eingetragen." -f $weg.Count, $pat.Count, $aus.Count)
  Prot ("Umlaute: {0} Felder bestehender Patienten berichtigt, {1} inzwischen geaendert und uebersprungen." -f $nk, $nu)
} catch {
  $tr.Rollback(); $con.Close()
  Prot "FEHLER in der Datenbank, alles zurueckgenommen: $($_.Exception.Message)"
  exit 1
}
$con.Close()

# Dateien: erst kopieren (aus den unveraenderten Quellen), dann zurueckholen, dann alte Namen wegraeumen
$archiv = Join-Path $ziel 'vor_neuaufbau'
$n = @{ kopie = 0; zurueck = 0; weg = 0 }
foreach ($art in 'kopie', 'zurueck', 'weg') {
  foreach ($d in ($dat | Where-Object { $_.Art -eq $art })) {
    $zp = Join-Path $ziel $d.Ziel
    if ($art -eq 'weg') {
      if (-not (Test-Path $zp)) { Prot "fehlt schon: $($d.Ziel)"; continue }
      if ($Echt) {
        $ad = Join-Path $archiv (Split-Path $d.Ziel -Parent)
        if (-not (Test-Path $ad)) { New-Item -ItemType Directory -Path $ad -Force | Out-Null }
        Move-Item $zp $ad
      } else { Remove-Item $zp }
    } else {
      if ($art -eq 'kopie' -and (Test-Path $zp)) { Prot "Ziel gibt es schon, uebersprungen: $($d.Ziel)"; continue }
      if ($art -eq 'zurueck' -and (Test-Path $zp)) {
        if ($Echt) { $ad = Join-Path $archiv (Split-Path $d.Ziel -Parent); if (-not (Test-Path $ad)) { New-Item -ItemType Directory -Path $ad -Force | Out-Null }; Copy-Item $zp (Join-Path $ad ((Split-Path $d.Ziel -Leaf) + '.ueberschrieben')) }
      }
      $qp = if ($d.Quelle -like 'C:\Datenbanken\*') { $quelle + $d.Quelle.Substring(14) } else { $d.Quelle }
      Copy-Item $qp $zp -Force
    }
    $n[$art]++
  }
}
Prot ("Dateien: {0} kopiert, {1} zurueckgeholt, {2} alte Namen {3}." -f $n.kopie, $n.zurueck, $n.weg, $(if ($Echt) { "nach $archiv verschoben" } else { 'in der Testkopie geloescht' }))
Prot 'Fertig.'
