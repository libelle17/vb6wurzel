Attribute VB_Name = "RegComm"
' Klassenmodul cRegistry.cls hinzufügen
Option Explicit
' Modul für die Kommunikation mit der Registry
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, _
         ByVal lpSecurityAttributes As Any, phkResult&, lpdwDisposition&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, lpData As Any, ByVal cbData&)
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" _
       (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData As Any, lpcbData&)
Declare Function RegOpenKey& Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey&, ByVal lpSubKey$, phkResult&)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)
Declare Function RegEnumValue& Lib "advapi32.dll" _
                 Alias "RegEnumValueA" ( _
                 ByVal hKey&, _
                 ByVal dwIndex&, _
                 ByVal lpValueName$, _
                 lpcbValueName&, _
                 ByVal lpReserved&, _
                 lpType&, _
                 ByVal lpData$, _
                 lpcbData&)
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal Bytes As Long)

'Public Const HKEY_CLASSES_ROOT& = &H80000000
Public Const HCR& = &H80000000
'Public Const HKEY_CURRENT_USER& = &H80000001
Public Const HCU& = &H80000001
'Public Const HKEY_LOCAL_MACHINE& = &H80000002
Public Const HLM& = &H80000002
'Public Const HKEY_USERS& = &H80000003
Public Const HU& = &H80000003
'public const HKEY_CURRENT_CONTROLSET = &H80000005
Public Const HCC& = &H80000005

' jetzt in cRegistry.cls
'Public Const REG_NONE = 0
'Public Const REG_SZ = 1
'Public Const REG_EXPAND_SZ = 2
'Public Const REG_BINARY = 3
'Public Const REG_DWORD = 4
'Public Const REG_DWORD_BIG_ENDIAN = 5
'Public Const REG_LINK = 6
'Public Const REG_MULTI_SZ = 7
 
Public Const REG_OPTION_NON_VOLATILE = &H0
Public Const REG_OPTION_VOLATILE = &H1
Public Const REG_OPTION_BACKUP_RESTORE = &H4
Public Const KEY_ALL_Access = &H3F
Const KEY_CREATE_SUB_KEY As Long = &H4
Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Const KEY_QUERY_VALUE As Long = &H1
Const KEY_SET_VALUE As Long = &H2
Const KEY_NOTIFY As Long = &H10
Const ERROR_SUCCESS = &H0

#If False Then
Public wsh As IWshShell_Class ' %windir%\system32\wshom.ocx
#Else
Public wsh As New IWshShell_Class
#End If
Public cReg As New Registry
'Public fpos& ' Fehlerposition -> soll in haupt.bas o.ä. jeweils geschrieben werden
Const RegWurzel$ = "Software\GSProducts\"
Dim ErrNumber&
Dim ErrDescr$, ErrSource$
Dim ErrLastDllError&

Public Function doSetzReg(frm As Form, ctl$, Optional obWord%, Optional RegName$)
 Dim RegStelle$, rVal, obRegFehlt%
 On Error GoTo fehler
 RegStelle = RegWurzel + App.EXEName
 If LenB(RegName) = 0 Then RegName = frm.Controls(ctl).name
 rVal = frm.Controls(ctl)
 Call wsh.RegWrite("HKEY_CURRENT_USER" + "\" + RegStelle + "\" + RegName, rVal, IIf(obWord, "REG_DWORD", "REG_SZ"))
' cReg.WriteKey rVal, RegName, RegStelle, HKEY_CURRENT_USER, IIf(obWord, REG_DWORD, REG_SZ)
 Exit Function
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doSetzReg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doSetzReg

'Public WMIreg AS SWbemObjectEx ' %windir%\system32\wbem\wbemdisp.tlb
Public Function doHolReg(frm As Form, ctl$, Optional Vorgabe, Optional RegName$)
 Dim RegStelle$, rVal, obRegFehlt%
 On Error GoTo fehler
 RegStelle = RegWurzel + App.EXEName
 If wsh Is Nothing Then Set wsh = New IWshShell_Class ' = CreateObject("Wscript.Shell")
 On Error GoTo fehler
 If IsNull(RegName) Then RegName = ctl
 If LenB(RegName) = 0 Then RegName = ctl
 rVal = wsh.RegRead("HKEY_CURRENT_USER" + "\" + RegStelle + "\" + RegName)
 If obRegFehlt Or (IsNull(rVal)) Then
  rVal = Vorgabe
 End If
 frm.Controls(ctl) = rVal
 Exit Function
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
If ErrNumber = -2147024894 Then
  obRegFehlt = -1
  Resume Next ' Reg-Eintrag nicht gefunden
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doHolReg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doHolReg

' Public Function ProgEnde()
'  End
' End Function 'ProgEnde


'Wert (String/Text) für einen bestimmten
'Schlüsselnamen speichern. Sollte der Schlüssel nicht
'existieren, wird dieser autom. erstellt.
'
'Parameterbeschreibung
'---------------------
'hKey (Hauptschlüssel) : z.B. HKEY_CURRENT_USER
'sPath (Schlüsselpfad) : z.B. MeineAnwendung
'sValue (Schlüsselname): z.B. Path
'iData (Schlüsselwert) : z.B. progverz\MeineAnwendung
           
Public Sub fStSpei(hKey&, sPath$, sValue$, iData$)

  Dim vRet, vDisp& ' 1 = neu, 2 = schon da
  Dim erg&
  On Error GoTo fehler
  erg = RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  erg = RegSetValueEx(vRet, sValue, 0, REG_SZ, ByVal iData, Len(iData))
  erg = RegCloseKey(vRet)
  Exit Sub
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fStSpei/" & App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fStSpei


Public Sub fDWSpei(hKey&, sPath$, sValue$, iData)
  Dim vRet As Variant, vDisp&, erg& ' 1 = neu, 2 = schon da
  On Error GoTo fehler
  erg = RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Dim dwFlags, pdwtype&, pvData, pcbdata
'  Dim DataVor, DT&, erg, lpcbData, lpValue
'  erg = RegQueryValueEx(hKey, sValue, 0, DT, DataVor, lpcbData)
'      IF RegQueryValueEx(hKey, sValue, 0, DT, lpValue, lpcbData) = ERROR_MORE_DATA THEN
'         lpValue = Space$(lpcbData)
'        'retrieve the desired value
'         erg = RegQueryValueEx(hKey, sValue, 0&, DT, ByVal lpValue, lpcbData) = ERROR_SUCCESS
'      END IF  'If RegQueryValueEx (first call)
  RegSetValueEx vRet, sValue, 0, REG_DWORD, CLng(iData), 4
  RegCloseKey vRet
  Exit Sub
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fDWSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fDWSpei


Public Sub fBeiSpei(hKey&, sPath$, sValue, iData, iLen&)
  Dim erg
  Dim vRet As Variant, vDisp& ' 1 = neu, 2 = schon da
    
  Call RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Dim dwFlags, pdwtype&, pvData, pcbdata
'  Dim DataVor, DT&, erg, lpcbData, lpValue
'  erg = RegQueryValueEx(hKey, sValue, 0, DT, DataVor, lpcbData)
'      IF RegQueryValueEx(hKey, sValue, 0, DT, lpValue, lpcbData) = ERROR_MORE_DATA THEN
'         lpValue = Space$(lpcbData)
'        'retrieve the desired value
'         erg = RegQueryValueEx(hKey, sValue, 0&, DT, ByVal lpValue, lpcbData) = ERROR_SUCCESS
'      END IF  'If RegQueryValueEx (first call)
  erg = RegSetValueEx(vRet, sValue, 0, REG_BINARY, iData, iLen)
  RegCloseKey vRet
  Exit Sub
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fBeiSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fBeiSpei

Public Sub fBiSpei(hKey&, sPath$, sValue, iData)
  Dim erg
  Dim VarT&
  Dim vRet As Variant, vDisp& ' 1 = neu, 2 = schon da
  On Error GoTo fehler
  Call RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Dim dwFlags, pdwtype&, pvData, pcbdata
'  Dim DataVor, DT&, erg, lpcbData, lpValue
'  erg = RegQueryValueEx(hKey, sValue, 0, DT, DataVor, lpcbData)
'      IF RegQueryValueEx(hKey, sValue, 0, DT, lpValue, lpcbData) = ERROR_MORE_DATA THEN
'         lpValue = Space$(lpcbData)
'        'retrieve the desired value
'         erg = RegQueryValueEx(hKey, sValue, 0&, DT, ByVal lpValue, lpcbData) = ERROR_SUCCESS
'      END IF  'If RegQueryValueEx (first call)
 Dim abData() As Byte
 Dim cData As Long
 VarT = VarType(iData)
 Select Case VarT
  Case 8: abData = StrConv(iData, vbFromUnicode)
  Case 11: ReDim abData(0): abData(0) = -(iData)
  Case 17
   ReDim abData(0)
   abData(0) = iData
  Case 8209: abData = iData
  Case Else
   ReDim abData(0)
   Call CopyMem(abData(0), (iData), 1)
 End Select
 cData = (UBound(abData) - LBound(abData) + 1)
  If VarT = 11 Or VarT = 17 Or VarT = 8209 Then
   erg = RegSetValueEx(vRet, sValue, 0, REG_BINARY, abData(0), cData)
  Else
   erg = RegSetValueEx(vRet, sValue, 0, REG_BINARY, VarPtr(abData(0)), cData)
  End If
  RegCloseKey vRet
  Exit Sub
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fBiSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fBiSpei

'Wert für einen bestimmten
'Schlüsselnamen auslesen.
'
'Parameterbeschreibung
'---------------------
'hKey (Hauptschlüssel) : z.B. HKEY_CURRENT_USER
'sPath (Schlüsselpfad) : z.B. MeineAnwendung
'sValue (Schlüsselname): z.B. Path
'Rückgabewert          : z.B. progverz\MeineAnwendung

Public Function fWertLesen(hKey&, sPath$, sValue$, Optional Länge&)
  Dim vRet
  RegOpenKey hKey, sPath, vRet
  fWertLesen = fRegAbfrageWert(vRet, sValue, Länge)
  RegCloseKey vRet
End Function ' fWertLesen

Function test2()
test2 = fWertLesen(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_AdminToolsTemp")
' test = GetReg(1, "AppEvents\Schemes\Apps\.Default\SystemExit\.Current", vns)
Call fDWSpei(HCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_AdminToolsTemp", 1)
End Function ' test2

'Wird von "fWertLesen" aufgerufen und gibt den Wert
'eines Schlüsselnamens zurück. Hierbei wird autom.
'ermittelt, ob es sich um einen String oder Binärwert
'handelt.
Public Function fRegAbfrageWert(ByVal hKey&, ByVal sValueName$, Optional lBufferSizeData&)
  Dim sBuffer$
  Dim lRes&
  Dim lTypeValue&
  Dim iData&
  On Error GoTo fehler
  lRes = RegQueryValueEx(hKey, sValueName, 0, lTypeValue, ByVal 0, lBufferSizeData)
  If lRes = 0 Then
    If lTypeValue = REG_SZ Then
      sBuffer = String$(lBufferSizeData, vbNullChar)
      lRes = RegQueryValueEx(hKey, sValueName, 0, REG_SZ, ByVal sBuffer, lBufferSizeData)
      If lRes = 0 Then
        Dim pos&
        pos = InStr(1, sBuffer, vbNullChar)
        If pos = 0 Then
         fRegAbfrageWert = Left$(sBuffer, pos)
        Else
         fRegAbfrageWert = Left$(sBuffer, pos - 1)
        End If
      End If
    ElseIf lTypeValue = REG_DWORD Then
      lRes = RegQueryValueEx(hKey, sValueName, 0, REG_DWORD, iData, lBufferSizeData)
      If lRes = 0 Then
        fRegAbfrageWert = iData
      End If
    ElseIf lTypeValue = REG_BINARY Then
     Dim erg() As Byte, ergS$
     ReDim erg(0 To lBufferSizeData - 1)
     lRes = RegQueryValueEx(hKey, sValueName, 0&, REG_BINARY, erg(0), lBufferSizeData)
'     ergS = StrConv(erg, vbUnicode)
     fRegAbfrageWert = erg
    End If
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fRegAbfrageWert/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
  
End Function ' fRegAbfrageWert(ByVal hKey&, ByVal sValueName$, Optional lBufferSizeData&)


'Löschen eines Schlüsselnamens
'
'Parameterbeschreibung
'---------------------
'hKey (Hauptschlüssel) : z.B. HKEY_CURRENT_USER
'sPath (Schlüsselpfad) : z.B. MeineAnwendung
'sValue (Schlüsselname): z.B. Path

Public Sub fWerteLoeschen(hKey&, sPath$, sValue$)

  Dim vRet, vDisp&  ' 1 = neu, 2 = schon da
  On Error GoTo fehler
  Call RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Call RegDeleteValue(vRet, sValue)
  RegCloseKey vRet
  Exit Sub
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fWerteLoeschen/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fWerteLoeschen

Function getReg(Zuord&, schlüssel$, Wert$) As Variant
Dim erg& ' Ergebnisse von RegOpenKeyEx und RegQueryValueEx
Dim hKey&
Dim lpSubKey$
Dim ulOptions&
Dim samDesired&
Dim phkResult&
Dim lpValueName$
Dim lpReserved&
Dim lpType&
Dim lpData As String * 100
Dim lpcbData&
Dim i%
On Error GoTo fehler
Select Case Zuord
 Case 0, &H80000000
    hKey = &H80000000 ' HKEY_CLASSES_ROOT ' steht in winreg.h
 Case 1, &H80000001
    hKey = &H80000001 ' HKEY_CURRENT_USER ' steht in winreg.h
 Case 2, &H80000002
    hKey = &H80000002 ' HKEY_LOCAL_MACHINE ' steht in winreg.h
 Case 3, &H80000003
    hKey = &H80000003 ' HKEY_USERS ' steht in winreg.h
 Case 4, &H80000004
    hKey = &H80000004 ' HKEY_PERFORMANCE_DATA' steht in winreg.h
 Case 5, &H80000005
    hKey = &H80000005 ' HKEY_CURRENT_CONFIG' steht in winreg.h
 Case 6, &H80000006
    hKey = &H80000006 ' HKEY_DYN_DATA' steht in winreg.h
 Case Else
    hKey = Zuord
End Select
lpSubKey = schlüssel
ulOptions = 0
samDesired = &H20119 ' ' steht in winnt.h -> KEY_READ (alles zusammenzählen)
'Const KEY_READ& = &H20019
'Const KEY_WOW64_64KEY& = &H100
erg = RegOpenKeyEx(hKey, lpSubKey, ulOptions, samDesired, phkResult)
If erg = 0 Then ' Debug.Print "Success=", erg, "phkResult=", phkResult
 lpValueName = Wert
 lpReserved = 0
 lpType = 1
 lpcbData = 100
 erg = RegQueryValueEx(phkResult, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
 If lpType <> 1 Then
  erg = RegQueryValueEx(phkResult, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
 End If
 getReg = RegTrim$(lpData, lpcbData)
End If 'erg = 0
Exit Function
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetReg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetReg

Function RegTrim$(Val$, Optional ByRef lenge&)
 Dim i&
 On Error GoTo fehler
 If lenge = 0 Then
  lenge = Len(Val)
 End If
 For i = lenge To 1 Step -1
   If LenB(Mid$(Val, i, 1)) = 0 Then
    lenge = lenge - 1
   Else
    If Asc(Mid$(Val, i, 1)) = 0 Then
     lenge = lenge - 1
    Else
     Exit For
    End If
   End If
  Next i
  RegTrim = Left$(Val, lenge)
 Exit Function
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "Fehlerposition: " & CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegTrim/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RegTrim
Public Function ReadRegistryGetVALUES$(ByVal Group&, ByVal Section$, idx&, Inhalt$)
Dim lResult&, lKeyValue&, lValueLength&, td As Double, res&, vTyp&, cbValueN&
Dim ValueN As String * 2048, sValue As String * 2048
On Error GoTo fehler
cbValueN = Len(ValueN)
lValueLength = Len(sValue)
'ValueN = space$(cbValueN)
'sValue = space$(lValueLength)
Inhalt = vNS
sValue = vNS
On Error Resume Next
lResult = RegOpenKeyEx(Group, Section, 0, &HF003F, lKeyValue)
If lResult = 0 Then
 lResult = RegEnumValue(lKeyValue, idx, ValueN, cbValueN, 0&, vTyp, sValue, lValueLength)
 If (lResult = 0) And (Err.Number = 0) Then
   sValue = getReg(2, Section, Trim$(ValueN))
'   sValue = LEFT(sValue, InStr(sValue, vbnullchar) - 1)
   ValueN = Left$(ValueN, InStr(ValueN, vbNullChar) - 1)
   ReadRegistryGetVALUES = ValueN
   If (InStrB(sValue, "Path") <> 0) Then
    Inhalt = Mid$(sValue, InStr(sValue, "Path") + 5)
    Inhalt = Left$(Inhalt, InStr(Inhalt, "Permiss") - 2)
   End If
   lResult = RegCloseKey(lKeyValue)
 End If
End If
 Exit Function
fehler:
ErrNumber = Err.Number
ErrDescr = Err.Description
ErrSource = Err.source
ErrLastDllError = Err.LastDllError
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in ReadRegisterGetVALUES/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, Idx&)

'Sub REG_EnumUSBDevices(Schlüssel$)
Sub regEnumSub(Abschn&, schlüssel$, Sch)
Dim Key As String, lCount As Byte, lResult As Long, hKey As Long
ReDim Sch(0)
'lResult = RegOpenKeyEx(abschn, "SYSTEM\CurrentControlSet\Enum", 0, KEY_ENUMERATE_SUB_KEYS, hkey)
lResult = RegOpenKeyEx(Abschn, schlüssel, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
If Not (lResult = ERROR_SUCCESS) Then Exit Sub

lCount = 0
Do
 Key = Space$(1024)
 lResult = RegEnumKeyEx(hKey, lCount, Key, Len(Key), 0, vNS, 0, 0)
 If lResult = ERROR_SUCCESS Then
  ReDim Preserve Sch(UBound(Sch) + 1)
  If Left$(Key, 1) = vbNullChar Then
   Sch(UBound(Sch)) = vNS
  Else
   Sch(UBound(Sch)) = Trim$(Left$(Key, InStr(Key, vbNullChar) - 1))
  End If
 End If
 lCount = lCount + 1
Loop Until Not (lResult = ERROR_SUCCESS)
End Sub ' regEnumSub(Abschn&, Schlüssel$, Sch)

Sub regEnumVal(schlüssel$, Sch, Inh)
Dim Key As String, lCount As Byte, lResult As Long, hKey As Long, lValueLength&
Const vnz% = 2048
Dim ValueN$, cbValueN&, vTyp&, sValue$
cbValueN = Len(ValueN)
lValueLength = Len(sValue)
ReDim Sch(0), Inh(0)
lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, schlüssel, 0, KEY_QUERY_VALUE, hKey)
If Not (lResult = ERROR_SUCCESS) Then Exit Sub

lCount = 1
Do
 ValueN = String$(vnz, 0)
 sValue = String$(vnz, 0)
 lResult = RegEnumValue(hKey, lCount, ValueN, vnz, 0&, vTyp, sValue, vnz)
 If lResult = ERROR_SUCCESS Then
  ReDim Preserve Sch(UBound(Sch) + 1)
  ReDim Preserve Inh(UBound(Inh) + 1)
  Sch(UBound(Sch)) = Left$(ValueN, InStr(ValueN, vbNullChar))
  Inh(UBound(Inh)) = Left$(sValue, InStr(sValue, vbNullChar))
 End If
 lCount = lCount + 1
Loop Until Not (lResult = ERROR_SUCCESS)
Call RegCloseKey(hKey)
End Sub ' regEnumVal(Schlüssel$, Sch, Inh)

