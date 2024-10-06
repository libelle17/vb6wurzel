Attribute VB_Name = "CSV"
Option Explicit

' in DatenSpeichern_Click
Function machCSVs$(CNs$, dn$)
 Dim rs As New ADODB.Recordset
 If LenB(CNs) = 0 Then CNs = "DRIVER={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=" & MDI.dbv.pwd & ";database=dp;option=11"
 rs.Open "show tables", CNs, adOpenStatic, adLockReadOnly
 machCSVs = dn & IIf(Right(dn, 1) = "\", vNS, "\") & rs.ActiveConnection.DefaultDatabase & Format(Now, "_yyyymmdd_hhmmss")
 On Error Resume Next
 MkDir machCSVs
 On Error GoTo 0
 Do While Not rs.EOF
  Call machCSV(CNs, rs.Fields(0), machCSVs & "\" & rs.Fields(0) & ".csv")
  rs.Move 1
 Loop
End Function ' machcsvs

'"DRIVER={MySQL ODBC 5.1 Driver};server=linux;uid=praxis;pwd=...;database=dp;option=" & opti
' in machCSVs
Function machCSV(CNs$, tn$, dn$)
 If LenB(CNs) = 0 Then CNs = "DRIVER={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=" & MDI.dbv.pwd & ";database=dp;option=11"
 Dim oboffen%, i%, Sg$, Wert$
 Dim rs As New ADODB.Recordset
 Call rs.Open("SELECT * FROM `" & tn & "`", CNs, adOpenDynamic, adLockReadOnly)
 Do While Not rs.EOF
  If Not oboffen Then
   Open dn For Output As #327
   oboffen = True
   For i = 0 To rs.Fields.Count - 1
    Sg = Sg & vNS & rs.Fields(i).name & vNS & ";" ' IIf(i < rs.Fields.Count - 1, ";", vns)
   Next i
   Print #327, Sg
  End If
  Sg = vNS
  For i = 0 To rs.Fields.Count - 1
   Select Case rs.Fields(i).Type
    Case adBoolean ' 11
     If IsNull(rs.Fields(i).Value) Then Wert = "0" Else If rs.Fields(i) Then Wert = "-1" Else Wert = "0"
    Case 16, 17, 2, 18, 3, 19, 4, 5, 20, 21, 131, 139, 6, 14
     If IsNull(rs.Fields(i)) Then Wert = 0 Else Wert = REPLACE(REPLACE(rs.Fields(i), ".", vNS), ",", ".")
    Case 7, 64, 133, 134, 135
     Wert = datformhier(rs.Fields(i).Value)
     If Wert = "##" Or Wert = "''" Then Err.Raise 999, , "Fehler in machCSV mit Wert = ""##"" OR wert = ""''"" in Datumsfeld"
    Case 8, 129, 130, 200, 201, 202, 203, 0, 9, 12, 13, 72, 128, 132, 138, 204, 205
     Wert = Chr(34) & umw(IIf(IsNull(rs.Fields(i).Value), vNS, rs.Fields(i).Value)) & Chr(34)
    Case Else
     Err.Raise 999, , "Fehler in machCSV mit unbekanntem Datentyp: " & rs.Fields(i).Type
   End Select
   Sg = Sg & Wert & ";" ' IIf(i < rs.Fields.Count - 1, ";", vns)
  Next i
  Print #327, Sg
  rs.Move 1
 Loop
 If oboffen Then Close #327
End Function ' machCSV

Function datformhier(fm) As Date
 If IsNull(fm) Then datformhier = 0 Else datformhier = fm
End Function ' datformhier

Function umw$(Sg$)
  umw = Sg
End Function ' umw
