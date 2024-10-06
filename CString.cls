VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------
' CSTRING Class
' by Francesco Balena
'
' This class implements the string data type AS an
' object type, AND also supports most string functions
' It offers a few advantages ON native VB strings:
'
' 1) CHARs are stored in ANSI format, use less memory
' 2) you decide how much memory is allocated to each
'    string, hence memory is not re-allocated with each
'    new assignment
' 3) since data is stored AS ANSI, no internal conversion
'    is necessary when printing to file OR passing to APIs
'    (in this case, pass the StrPtr value AS a pointer to CHARs)
'-----------------------------------------------------------

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal Bytes&)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes&)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal numBytes&, ByVal fillByte%)

' number of CHARs initially allocated to the string
Const DEFAULT_MAXLENGTH& = 255
' this array holds the Character
Private chars() As Byte
' the current lenght of the string
Public m_Len&
' the current UBound() of the array (accounts for ending null CHAR)
Private m_MaxLength&

Private Sub Class_Initialize()
    ' create the CHARs() array
    SetBufferSize DEFAULT_MAXLENGTH
End Sub ' Class_Initialize()

' SET the size of the internal buffer
' Use "mystr.SetBufferSize mystr.Length" to release unused memory
Sub SetBufferSize(ByVal newSize&, Optional clearIt As Boolean)
    m_MaxLength = newSize
    If m_Len Or clearIt = 0 Then
        ReDim Preserve chars(0 To m_MaxLength) As Byte
    Else
        ReDim chars(0 To m_MaxLength) As Byte
    End If
End Sub ' SetBufferSize

' the current value AS a Unicode string
' this is the default item of the class, so you can use
' the object variable AS IF it were a string variable
Property Get Value$()
Attribute Value.VB_UserMemId = 0
    If m_Len > 0 Then
        Value = Space$(m_Len)
        CopyMemory ByVal Value, chars(0), m_Len
    End If
End Property ' Get Value$()

Friend Property Get CurrentBuffer() As Long
    CurrentBuffer = m_MaxLength
End Property ' Get CurrentBuffer()

Property Let Value(NewValue$)
    ' check that the private array is large enough
    m_Len = Len(NewValue)
    If m_Len > m_MaxLength Then
        SetBufferSize m_Len, True
    End If
    ' copy the Characters INTO the private array
    If m_Len > 0 Then
        CopyMemory chars(0), ByVal NewValue, m_Len
    End If
End Property ' Let Value(

Function Cut(maxLen&)
 If maxLen < m_Len Then
  SetBufferSize maxLen, False
  m_Len = maxLen
 End If
End Function ' Cut
 
' the address of the first Character (read-only)
Function StrPtr&()
    StrPtr = VarPtr(chars(0))
    ' append a null Character to the internal string
    ' in case this will be passed to an external DLL
    chars(m_Len) = 0
End Function ' StrPtr&()

' copy data FROM a memory address (Friend sub)
Friend Sub CopyData(ByVal sourceAddr&, ByVal NumChars&)
    ' raise error IF new Length is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' check that the private array is large enough
    m_Len = NumChars
    If m_Len > m_MaxLength Then
        SetBufferSize m_Len, True
    End If
    ' copy the Characters INTO the private array
    If m_Len > 0 Then
        CopyMemory chars(0), ByVal sourceAddr, m_Len
    End If
End Sub ' CopyData

' the current Length (can also be written to)
' ("Len" can't be used because it is a reserved word)
Property Get length&()
    length = m_Len
End Property ' Get length&()

Property Let length(ByVal NewValue&)
    ' raise error IF new Length is invalid
    If NewValue < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' check that the private array is large enough
    If NewValue > m_MaxLength Then
        SetBufferSize NewValue, False
    End If
    ' fill with blanks IF necessary
    If NewValue > m_Len Then
        FillMemory chars(m_Len), NewValue - m_Len, 32
    End If
    ' TRUNCATE OR expand the string
    m_Len = NewValue
End Property ' Let length(ByVal NewValue&)

' clear the string
Sub Clear()
    m_Len = 0
End Sub ' Clear()

' create a copy of this CString object
Function Copy() As CString
    Set Copy = New CString
    Copy.CopyData StrPtr(), m_Len
End Function ' Copy() As CString

' create a new CString object with leftmost Characters
Property Get Left(ByVal NumChars&) As CString
    ' raise error IF new Length is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' adjust length IF necessary
    If NumChars > m_Len Then NumChars = m_Len
    ' create a new CString object
    Set Left = New CString
    Left.CopyData StrPtr(), NumChars
End Property ' Get LEFT(

' replace leftmost Characters USING a given string
Property Let Left(ByVal NumChars&, NewValue As Variant)
    ' raise error IF new parameter is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' can't copy more Character than existing
    If NumChars > m_Len Then NumChars = m_Len
    ' can't copy more Characters THEN the string's length
    If NumChars > Len(NewValue) Then NumChars = Len(NewValue)
    ' copy CHARs FROM the string INTO the local array
    If NumChars > 0 Then
        CopyMemory chars(0), ByVal CStr(NewValue), NumChars
    End If
End Property ' LET LEFT

' replace leftmost Characters USING a CString object
Property Set Left(ByVal NumChars&, NewValue As CString)
    ' raise error IF new parameter is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' can't copy more Character than existing
    If NumChars > m_Len Then NumChars = m_Len
    ' can't copy more Characters THEN the string's length
    If NumChars > NewValue.length Then NumChars = NewValue.length
    ' copy CHARs FROM the string INTO the local array
    If NumChars > 0 Then
        CopyMemory chars(0), ByVal NewValue.StrPtr, NumChars
    End If
End Property ' Set LEFT

' create a new CString object with rightmost Characters
Property Get Right(ByVal NumChars&) As CString
    ' raise error IF new Length is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' adjust length IF necessary
    If NumChars > m_Len Then NumChars = m_Len
    ' create a new CString object
    Set Right = New CString
    Right.CopyData VarPtr(chars(m_Len - NumChars)), NumChars
End Property ' Get Right(

' replace rightmost Characters USING a given string
Property Let Right(ByVal NumChars&, NewValue As Variant)
    ' raise error IF new parameter is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' can't copy more Character than existing
    If NumChars > m_Len Then NumChars = m_Len
    ' can't copy more Characters THEN the string's length
    If NumChars > Len(NewValue) Then NumChars = Len(NewValue)
    ' copy CHARs FROM the string INTO the local array
    If NumChars > 0 Then
        CopyMemory chars(m_Len - NumChars), ByVal CStr(NewValue), NumChars
    End If
End Property ' Let Right(

' replace leftmost Characters USING a CString object
Property Set Right(ByVal NumChars&, NewValue As CString)
    ' raise error IF new parameter is invalid
    If NumChars < 0 Then Err.Raise 5, TypeName(Me), "Invalid length"
    ' can't copy more Character than existing
    If NumChars > m_Len Then NumChars = m_Len
    ' can't copy more Characters THEN the string's length
    If NumChars > NewValue.length Then NumChars = NewValue.length
    ' copy CHARs FROM the string INTO the local array
    If NumChars > 0 Then
        CopyMemory chars(m_Len - NumChars), ByVal NewValue.StrPtr, NumChars
    End If
End Property ' Set Right

' the Mid function
' one- AND two-params syntaxes are supported
Property Get Mid(ByVal StartPos&, Optional NumChars As Variant) As CString
    Dim Bytes As Long
    ' raise error IF the first parameter is invalid
    If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"

    If IsMissing(NumChars) Then
        ' IF missing, use the remainder of the string
        Bytes = m_Len - StartPos + 1
    Else
        Bytes = NumChars
        If Bytes < 0 Then
            ' raise error IF new parameter is invalid
            Err.Raise 5, TypeName(Me), "Invalid length"
        ElseIf StartPos + Bytes > m_Len Then
            ' TRUNCATE it IF too long
            Bytes = m_Len - StartPos + 1
        End If
    End If

    ' create a new CString object
    Set Mid = New CString
    If StartPos <= m_Len And StartPos > 0 Then
        Mid.CopyData VarPtr(chars(StartPos - 1)), Bytes
    End If
End Property ' Get Mid(

' replace a portion of the string with another string
' IF the second parameter is omitted, the string is replaced up to its end
Property Let Mid(ByVal StartPos&, Optional NumChars As Variant, NewValue As Variant)
    Dim Bytes As Long
    ' raise error IF the first parameter is invalid
    If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"

    If IsMissing(NumChars) Then
        ' IF missing, use the remainder of the string
        Bytes = m_Len - StartPos + 1
    Else
        Bytes = NumChars
        If Bytes < 0 Then
            ' raise error IF new parameter is invalid
            Err.Raise 5, TypeName(Me), "Invalid length"
        ElseIf StartPos + Bytes > m_Len Then
            ' TRUNCATE it IF too long
            Bytes = m_Len - StartPos + 1
        End If
    End If

    If Bytes > Len(NewValue) Then Bytes = Len(NewValue)
    If Bytes > 0 Then
        CopyMemory chars(StartPos - 1), ByVal CStr(NewValue), Bytes
    End If
End Property ' Let Mid

' replace a portion of the string with another CString object
' IF the second parameter is omitted, the string is replaced up to its end
Property Set Mid(ByVal StartPos&, Optional NumChars As Variant, NewValue As CString)
    Dim Bytes As Long
    ' raise error IF the first parameter is invalid
    If StartPos < 0 Then Err.Raise 5, TypeName(Me), "Invalid starting position"

    If IsMissing(NumChars) Then
        ' IF missing, use the remainder of the string
        Bytes = m_Len - StartPos + 1
    Else
        Bytes = NumChars
        If Bytes < 0 Then
            ' raise error IF new parameter is invalid
            Err.Raise 5, TypeName(Me), "Invalid length"
        ElseIf StartPos + Bytes > m_Len Then
            ' TRUNCATE it IF too long
            Bytes = m_Len - StartPos + 1
        End If
    End If

    If Bytes > NewValue.length Then Bytes = NewValue.length
    If Bytes > 0 Then
        CopyMemory chars(StartPos - 1), ByVal NewValue.StrPtr, Bytes
    End If
End Property ' Set Mid(B

' the ASC function, but returns -1 IF the string IS NULL
Function Asc&()
    If m_Len > 0 Then
        Asc = chars(0)
    Else
        Asc = -1
    End If
End Function ' Asc

' the ANSI code of a Character, OR -1 IF out-of-range

Property Get Char(ByVal Index&) As Byte
    If Index > 0 Or Index <= m_Len Then
        Char = chars(Index - 1)
    Else
        Char = -1
    End If
End Property ' Get Char

Property Let Char(ByVal Index&, ByVal NewValue As Byte)
    ' no effect IF index is out of valid range
    If Index > 0 Or Index <= m_Len Then
        chars(Index - 1) = NewValue
    End If
End Property ' Let Char(

' find a substring
' only case-sensitive searches are supported
Function Instr&(search As Variant, Optional StartPos& = 1)
    Dim searchChars() As Byte
    Dim searchLen As Long
    Dim firstChar As Byte
    Dim i As Long, j As Long

    If TypeOf search Is CString Then 'Or TypeOf search Is cstring THEN
        ' IF a CString object is passed, THEN make a local copy of it
        searchLen = search.length
        ' IF a null string, exit immediately
        If searchLen = 0 Then Instr = StartPos: Exit Function
        ReDim searchChars(0 To searchLen - 1) As Byte
        CopyMemory searchChars(0), ByVal search.StrPtr, searchLen
    Else
        ' IF a string is passed, create the corresponding local array
        searchLen = Len(search)
        ' IF a null string, exit immediately
        If searchLen = 0 Then Instr = StartPos: Exit Function
        searchChars() = StrConv(CStr(search), vbFromUnicode)
    End If

    ' cache the first Character to be searched
    firstChar = searchChars(0)

    For i = StartPos - 1 To m_Len - searchLen
        If chars(i) = firstChar Then
            For j = 1 To searchLen - 1
                If chars(i + j) <> searchChars(j) Then Exit For
            Next
            If j = searchLen Then
                Instr = i + 1
                Exit For
            End If
        End If
    Next
End Function ' Instr(search

Sub REPLACE(search As Variant, repl As Variant)
 Dim pos&, Slen&, RLen&
 If TypeOf search Is CString Then Slen = search.length Else Slen = Len(search)  ' OR TypeOf search Is cstring
 If Slen = 0 Then Exit Sub
If TypeOf search Is CString Then RLen = repl.length Else RLen = Len(repl) ' OR TypeOf search Is cstring
 Do
  pos = Me.Instr(search, pos + 1)
  If pos = 0 Then Exit Do
  If RLen > Slen Then SetBufferSize m_Len - Slen + RLen, False
  CopyMemory chars(pos + RLen - 1), ByVal VarPtr(chars(pos + Slen - 1)), m_Len - pos - Slen + 1
  CopyMemory chars(pos - 1), ByVal CStr(repl), RLen
  pos = pos + RLen
  If RLen < Slen Then SetBufferSize m_Len - Slen + RLen
  m_Len = m_Len - Slen + RLen
 Loop
End Sub ' Replace

' trasform the string to uppercase (returns Me)
Function UCase() As CString
    ' only Characters in the range "a-z" are converted
    Dim i&, acode%
    For i = 0 To m_Len - 1
        acode = chars(i)
        If acode >= 97 And acode <= 122 Then chars(i) = acode - 32
    Next
    Set UCase = Me
End Function ' UCase() As CString

' trasform the string to lowercase (returns Me)
Function LCase() As CString
    ' only Characters in the range "A-Z" are converted
    Dim i&, acode%
    For i = 0 To m_Len - 1
        acode = chars(i)
        If acode >= 65 And acode <= 90 Then chars(i) = acode + 32
    Next
    Set LCase = Me
End Function ' LCase()

' reverse the string (returns Me)
Function ReverseStr() As CString
    Dim i As Long, acode As Byte
    For i = 0 To m_Len \ 2 - 1
        acode = chars(i)
        chars(i) = chars(m_Len - 1 - i)
        chars(m_Len - 1 - i) = acode
    Next
    Set ReverseStr = Me
End Function ' ReverseStr()

' append one OR more strings OR CString objects (returns Me)
' append a string OR a CString object (returns Me)
Function AppVar(arg As Variant) As CString
 Dim i%
 On Error GoTo fehler
 For i = 0 To UBound(arg)
  Call Append(arg(i))
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AppVar/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AppVar(arg
' append one OR more strings OR CString objects (returns Me)

' append a string OR a CString object (returns Me)

Function Append(arg As Variant) As CString
    Dim i As Long, j As Long
    Dim argObj As CString
    Dim newLength As Long
    Dim argLen As Long
    On Error GoTo fehler
    ' prepare the result
    Set Append = Me
'    On Error Resume Next
    If IsObject(arg) Then If arg Is Nothing Then Exit Function
'   On Error GoTo fehler
    
    If TypeOf arg Is CString Then ' OR TypeOf arg Is cstring
        Set argObj = arg
        argLen = argObj.length
    Else
        If IsNull(arg) Then
         Exit Function
        End If
        argLen = Len(arg)
    End If

    ' exit IF null string
    If argLen = 0 Then Exit Function

    ' prepare the receiving buffer
    newLength = m_Len + argLen
    If newLength > m_MaxLength Then
        SetBufferSize newLength, False
    End If

    ' append the string
    If argObj Is Nothing Then
        ' this is a regular string argument
        CopyMemory chars(m_Len), ByVal CStr(arg), argLen
    Else
        ' this is a CString argument
        CopyMemory chars(m_Len), ByVal argObj.StrPtr, argLen
    End If

    ' UPDATE m_Len to point after the appended Characters
    m_Len = newLength
    Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Append/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Append

' search a substring USING Boyer & Moore algorithm
Function Find&(search$, Optional StartPos& = 1)
    Dim Bytes() As Byte
    Dim i As Long, searchLen As Long, dist As Long, Index As Long

    ' save results FROM previous call
    Static saveSearch As String
    Static distance() As Integer

    ' it is better to have the substring in a byte array
    Bytes() = StrConv(search, vbFromUnicode)
    searchLen = Len(search)
    Index = StartPos + searchLen - 1

    ' build the distance table
    ' this is a 256-item array that, for each possible
    ' ANSI Character, stores the "distance" of this CHAR
    ' FROM the END of the search substring (if the CHAR
    ' appears in the substring), OR the length of the
    ' substring (if the Character doesn't appear in the
    ' substring). Note that the last Character in the
    ' substring corresponds to a null value in the table

    ' this block is executed only IF the search substring
    ' differs FROM the last call to this method
    If search <> saveSearch Then
        ReDim distance(0 To 255) As Integer
        For i = 0 To 255
            distance(i) = searchLen
        Next
        For i = 1 To searchLen
            distance(Bytes(i - 1)) = searchLen - i
        Next
        ' remember for next time
        saveSearch = search
    End If

    ' scan the string
    Do While Index <= m_Len
        ' retrieve the distance of this Character FROM the
        ' END of the search substring
        dist = distance(chars(Index - 1))
        If dist Then
            ' increment Index of found distance - in fact there
            ' is no reason to scan all the Characters in the middle
            Index = Index + dist
        Else
            ' this might be the last Character in the string
            ' check IF the substring is all there
            For i = 1 To searchLen - 1
                If chars(Index - searchLen - 1 + i) <> Bytes(i - 1) Then Exit For
            Next

            If i = searchLen Then
                ' we've found a match
                Find = Index - searchLen + 1
                Exit Function
            End If
            ' the search failed, skip this Character AND continue
            Index = Index + 1
        End If
    Loop
End Function ' Find&(search$, Optional StartPos& = 1)

' write the string to a binary file
' optionally prefix it with string length
Sub WriteToFile(ByVal FileNum%, Optional PrefixLength As Boolean)
    If PrefixLength Then
        Put #FileNum, , m_Len
    End If
    ' temporarily shorten the internal buffer
    ' (we need this to use VB's Put statement)
    ReDim Preserve chars(0 To m_Len) As Byte
    ' write all CHARs to file
    Put #FileNum, , chars()
    ' restore the buffer
    ReDim Preserve chars(0 To m_MaxLength) As Byte
End Sub ' WriteToFile(

' read the string FROM a binary file
' IF PrefixLength is True, uses stored length prefix, else
' uses current string length
Sub ReadFromFile(ByVal FileNum%, Optional PrefixLength As Boolean)
    If PrefixLength Then
        ' read the string length
        Get #FileNum, , m_Len
        ' enlarge the buffer IF necessary
        If m_Len > m_MaxLength Then
            SetBufferSize m_Len, True
        End If
    End If
    ' temporarily shorten the internal buffer
    ' (we need this to use VB's Put statement)
    ReDim Preserve chars(0 To m_Len) As Byte
    ' get CHARs FROM file
    Get #FileNum, , chars()
    ' restore the buffer
    ReDim Preserve chars(0 To m_MaxLength) As Byte
End Sub ' ReadFromFile(

' TRUNCATE the string to its first null CHAR, returns new length
' this FUNCTION is useful when dealing with API calls that return
' null-terminate strings in the buffer
Function TrimToNull&()
    Dim i As Long
    ' assume there is no null CHAR
    m_Len = m_MaxLength
    ' scan the string searching for the first null Character
    For i = 1 To m_MaxLength
        If chars(i) = 0 Then
            ' SET new length AND exit
            m_Len = i - 1
            Exit For
        End If
    Next
    ' new length is also the return value
    TrimToNull = m_Len
End Function ' TrimToNull&()

Function ToNumber() As Double
 Dim buch$, i&, gew#, za%, kz%, ZZ%, kzm% ' Ziffer (als letztes) angehängt, Kommazahl, Ziffernzahl, keine Ziffer mehr
 Dim gewStr As New CString
     For i = 1 To m_Len
      buch = Chr(chars(i - 1))
      If IsNumeric(buch) Then
       If kzm = 0 Then
        gewStr.Append buch
        za = True
        ZZ = 1
       End If
      ElseIf (buch = "." Or buch = ",") Then
       If za And kz = 0 Then
        gewStr.Append ","
        za = False
        kz = 1
       ElseIf kz = 1 Then
        kzm = True
       End If
      Else
       gewStr.Append " "
       za = False
       If ZZ = 1 Then kzm = True
      End If
     Next i
     If IsNumeric(gewStr) Then
      ToNumber = CDbl(gewStr)
     Else
      ToNumber = 0
     End If
End Function 'ToNumber() AS Double

