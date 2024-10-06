Attribute VB_Name = "RegComm0"
Option Explicit
'Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HCR = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
Public Const HCU = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HLM = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HU = &H80000003
Private Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, _
         ByVal lpSecurityAttributes As Any, phkResult&, lpdwDisposition&)
Private Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, lpData As Any, ByVal cbData&)
Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Private Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" _
       (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Private Declare Function RegOpenKey& Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey&, ByVal lpSubKey$, phkResult&)
Private Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData As Any, lpcbData&)
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
Const REG_NONE = 0
Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_OPTION_NON_VOLATILE = &H0
Const KEY_ALL_Access = &H3F
Const KEY_CREATE_SUB_KEY As Long = &H4
Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Const KEY_QUERY_VALUE As Long = &H1
Const KEY_SET_VALUE As Long = &H2
Const KEY_NOTIFY As Long = &H10
Const ERROR_SUCCESS = &H0
Function getReg(Zuord%, Schlüssel$, Wert$)
 Dim hKey&
 Select Case Zuord
  Case 0
    hKey = &H80000000 ' HKEY_CLASSES_ROOT ' steht in winreg.h
  Case 1
    hKey = &H80000001 ' HKEY_CURRENT_USER ' steht in winreg.h
  Case 2
    hKey = &H80000002 ' HKEY_LOCAL_MACHINE ' steht in winreg.h
  Case 3
    hKey = &H80000003 ' HKEY_USERS ' steht in winreg.h
  Case 4
    hKey = &H80000004 ' HKEY_PERFORMANCE_DATA' steht in winreg.h
  Case 5
    hKey = &H80000005 ' HKEY_CURRENT_CONFIG' steht in winreg.h
  Case 6
    hKey = &H80000006 ' HKEY_DYN_DATA' steht in winreg.h
  Case Else
    hKey = Zuord
 End Select
 getReg = fWertLesen(hKey, Schlüssel, Wert)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in getReg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getReg
Function fWertLesen(hKey As Long, sPath As String, sValue As String)

  Dim vRet As Variant
    
  RegOpenKey hKey, sPath, vRet
  fWertLesen = fRegAbfrageWert(vRet, sValue)
  RegCloseKey vRet
End Function
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

'Wird von "fWertLesen" aufgerufen und gibt den Wert
'eines Schlüsselnamens zurück. Hierbei wird autom.
'ermittelt, ob es sich um einen String oder Binärwert
'handelt.
Function fRegAbfrageWert(ByVal hKey As Long, ByVal sValueName As String) As String
    
  Dim sBuffer As String
  Dim lRes As Long
  Dim lTypeValue As Long
  Dim lBufferSizeData As Long
  Dim iData As Integer

  lRes = RegQueryValueEx(hKey, sValueName, 0, lTypeValue, ByVal 0, lBufferSizeData)
  If lRes = 0 Then
    If lTypeValue = REG_SZ Then
      sBuffer = String(lBufferSizeData, Chr$(0))
      lRes = RegQueryValueEx(hKey, sValueName, 0, 0, ByVal sBuffer, lBufferSizeData)
      If lRes = 0 Then
        If LenB(sBuffer) = 0 Then
         fRegAbfrageWert = vNS
        Else
         fRegAbfrageWert = Left$(sBuffer, InStr(1, sBuffer, Chr$(0)) - 1)
        End If
      End If
    ElseIf lTypeValue = REG_BINARY Or lTypeValue = REG_DWORD Then
      lRes = RegQueryValueEx(hKey, sValueName, 0, 0, iData, lBufferSizeData)
      If lRes = 0 Then
        fRegAbfrageWert = iData
      End If
    End If
  End If
End Function
           
Public Sub fStSpei(hKey&, sPath$, sValue$, iData$)
  Dim vRet, vDisp& ' 1 = neu, 2 = schon da
  Dim erg&
  On Error GoTo fehler
  erg = RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  erg = RegSetValueEx(vRet, sValue, 0, REG_SZ, ByVal iData, Len(iData))
  erg = RegCloseKey(vRet)
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fStSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fStSpei

Public Sub fdwspei(hKey&, sPath$, sValue$, iData)
  Dim vRet As Variant, vDisp& ' 1 = neu, 2 = schon da
  On Error GoTo fehler
  Call RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Dim dwFlags, pdwtype&, pvdata, pcbdata
'  Dim DataVor, DT&, erg, lpcbData, lpValue
'  erg = RegQueryValueEx(hKey, sValue, 0, DT, DataVor, lpcbData)
'      If RegQueryValueEx(hKey, sValue, 0, DT, lpValue, lpcbData) = ERROR_MORE_DATA Then
'         lpValue = Space$(lpcbData)
'        'retrieve the desired value
'         erg = RegQueryValueEx(hKey, sValue, 0&, DT, ByVal lpValue, lpcbData) = ERROR_SUCCESS
'      End If  'If RegQueryValueEx (first call)
  RegSetValueEx vRet, sValue, 0, REG_DWORD, CLng(iData), 4
  RegCloseKey vRet
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fDWSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fDWSpei

Public Sub fBiSpei(hKey&, sPath$, sValue, iData, iLen&)
  Dim erg
  Dim vRet As Variant, vDisp& ' 1 = neu, 2 = schon da
  On Error GoTo fehler
  Call RegCreateKeyEx(hKey, sPath, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_Access, 0&, vRet, vDisp)
  Dim dwFlags, pdwtype&, pvdata, pcbdata
'  Dim DataVor, DT&, erg, lpcbData, lpValue
'  erg = RegQueryValueEx(hKey, sValue, 0, DT, DataVor, lpcbData)
'      If RegQueryValueEx(hKey, sValue, 0, DT, lpValue, lpcbData) = ERROR_MORE_DATA Then
'         lpValue = Space$(lpcbData)
'        'retrieve the desired value
'         erg = RegQueryValueEx(hKey, sValue, 0&, DT, ByVal lpValue, lpcbData) = ERROR_SUCCESS
'      End If  'If RegQueryValueEx (first call)
  erg = RegSetValueEx(vRet, sValue, 0, REG_BINARY, iData, iLen)
  RegCloseKey vRet
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fBiSpei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fBiSpei

'Sub REG_EnumUSBDevices(Schlüssel$)
Sub regEnumSub(Schlüssel$)
Dim Key As String, lCount As Byte, lResult As Long, hKey As Long

'lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum", 0, KEY_ENUMERATE_SUB_KEYS, hkey)
lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Schlüssel, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
If Not (lResult = ERROR_SUCCESS) Then Exit Sub

lCount = 0
Do
Key = Space(1024)
lResult = RegEnumKeyEx(hKey, lCount, Key, Len(Key), 0, vbNullString, 0, 0)
If lResult = ERROR_SUCCESS Then
Key = Trim(Key)
Debug.Print (Left(Key, Len(Key) - 1))
End If
lCount = lCount + 1
Loop Until Not (lResult = ERROR_SUCCESS)
End Sub
Sub test()
 Dim v1$(), v2$(), i
 Call regEnumVal("SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", v1, v2)
 For i = 0 To UBound(v1)
  Debug.Print v1(i), v2(i)
 Next i
End Sub
Sub regEnumVal(Schlüssel$, Sch, Inh)
Dim Key As String, lCount As Byte, lResult As Long, hKey As Long, lValueLength&
Const vnz% = 2048
Dim ValueN$, cbValueN&, vTyp&, sValue$
cbValueN = Len(ValueN)
lValueLength = Len(sValue)
ReDim Sch(0), Inh(0)
lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Schlüssel, 0, KEY_QUERY_VALUE, hKey)
If Not (lResult = ERROR_SUCCESS) Then Exit Sub

lCount = 1
Do
 ValueN = String(vnz, 0)
 sValue = String(vnz, 0)
 lResult = RegEnumValue(hKey, lCount, ValueN, vnz, 0&, vTyp, sValue, vnz)
 If lResult = ERROR_SUCCESS Then
  ReDim Preserve Sch(UBound(Sch) + 1)
  ReDim Preserve Inh(UBound(Inh) + 1)
  Sch(UBound(Sch)) = Left(ValueN, InStr(ValueN, Chr(0)))
  Inh(UBound(Inh)) = Left(sValue, InStr(sValue, Chr(0)))
 End If
 lCount = lCount + 1
Loop Until Not (lResult = ERROR_SUCCESS)
Call RegCloseKey(hKey)
End Sub

