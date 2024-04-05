Attribute VB_Name = "JSONPARSE"
''''''''''''''''''  JSON TEXT parser  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    JSON TEXT  ---->  jsonvar Variant
'
'    Conversion is rather straightforward:
'
'    - number             --> translated to VBA Variant/Double
'    - "string"           --> translated to VBA Variant/String
'    - true/false literal --> translated to VBA Variant/Boolean
'    - null literal       --> translated to VBA Variant/Null
'    - [array]            --> translated to VBA Variant/Collection
'    - {object}           --> translated to VBA Variant/Dictionary
'
'    [array] is a comma-separated list of json values - or empty []
'    {object} is a comma-separated list of name/value pairs - or empty {}
'
'    This implementation conforms most with I-JSON RFC 7493:
'      - top-Level: any JSON value
'      - numbers: IEEE 754 double precision (as Excel)
'      - names: SHOULD be unique
'      - comparison: binary on unescaped strings
'
' MIT License
' Copyright (c) 2024 Attila Tarpai https://github.com/Halicery/JSON-VBAProject/
'

Option Private Module

Option Explicit
Option Compare Binary

Public Const ASCW_LBRACE As Long = 123
Public Const ASCW_RBRACE As Long = 125
Public Const ASCW_LBRACK As Long = 91
Public Const ASCW_RBRACK As Long = 93
Public Const ASCW_COLON As Long = 58
Public Const ASCW_QUOT As Long = 34
Public Const ASCW_PERIOD As Long = 46
Public Const ASCW_PLUS As Long = 43
Public Const ASCW_MINUS As Long = 45
Public Const ASCW_COMMA As Long = 44
Public Const ASCW_SOL As Long = 47 ' Solidus aka slash
Public Const ASCW_BSOL As Long = 92 ' Reverse Solidus aka backslash
Public Const ASCW_NUM0 As Long = 48
Public Const ASCW_NUM9 As Long = 57
Public Const ASCW_DOLLAR As Long = 36
Public Const ASCW_AST As Long = 42 ' Asterisk
Public Const ASCW_SPACE As Long = 32

Public Const ESCAPECHARS = "/""\btnfr"
Public Const ESCAPECODES = "/""\" & vbBack & vbTab & vbLf & vbFormFeed & vbCr
Public Const HEXCHARS = "0123456789ABCDEFabcdef"

Public json_string As String ' The JSON TEXT
Public json_pos As Long
Public json_ascw As Long     ' Long. EOF=-1
Public json_char As String

' Error handling: JSON Lib raises a few error numbers at run-time
' For now S=0 SCODE-s in the range of &H800, &H801, ... &H807
Public Enum JERRENUM
    JERRBASE = &H800
    JERRPARSE
    JERRPARSEOBJ
    JERRPATH
    JERRMAX = &H807
End Enum

''''''''''' Public '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub json_parse(json_text As String, ByRef jsonvar_out As Variant)
    If 0 = LenB(json_text) Then Err.Raise JERRPARSE, , "empty JSON input", "json_parse()"
    init_string_reader json_text
    nextToken
    read_json_value jsonvar_out
    json_string = vbNullString ' free BSTR buffer
    If -1 <> json_ascw Then Err.Raise JERRPARSE, , "top-level value not single: there is more input", "json_parse()"
End Sub

' json_error() meant to be called from an error handler
' Puts descriptive JSON error message into Err.Description and also returns the string (I hate #VALUE!)

Public Function json_error() As String
    If Err.Number And Not JERRMAX Then     ' not ours, 0 never raised ok
        ' leave original
    Else
        Select Case Err.Number
            Case JERRPARSE
                Err.Description = "#JSON parse: " & Err.Description
            Case JERRPARSEOBJ
                Err.Description = "#JSON parse object: " & Err.Description
            Case JERRPATH
                Err.Description = "#JSON path: " & Err.Description
        End Select
    End If
    json_error = Err.Description
End Function


''''''''''' Private '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub read_json_value(jsonvar_out)
    Select Case json_ascw
        Case ASCW_QUOT
            jsonvar_out = read_json_string
            nextToken
        Case ASCW_LBRACE
            Set jsonvar_out = read_json_obj
        Case ASCW_LBRACK
            Set jsonvar_out = read_json_arr
        Case 102&, 110&, 116&    ' f n t
            If check_json_string_for_word("true") Then
                jsonvar_out = True
            ElseIf check_json_string_for_word("false") Then
                jsonvar_out = False
            ElseIf check_json_string_for_word("null") Then
                jsonvar_out = Null
            Else
                Err.Raise JERRPARSE, , "invalid literal", "read_json_value()"
            End If
            nextToken
        Case Else  ' suppose number
            If ASCW_MINUS = json_ascw Then
                nextChar
                jsonvar_out = -read_json_number
            Else
                jsonvar_out = read_json_number
            End If
    End Select
End Sub

Private Function read_json_obj() As Scripting.Dictionary
    Set read_json_obj = New Scripting.Dictionary
    nextToken
    If json_ascw <> ASCW_RBRACE Then
        Dim name As String, value
        Do
            Select Case json_ascw
                Case ASCW_QUOT
                    name = read_json_string  ' unescape
                    nextToken
                    If json_ascw <> ASCW_COLON Then Err.Raise JERRPARSEOBJ, , "colon ':' expected after member name"
                    nextToken
                    read_json_value value
                    If read_json_obj.Exists(name) Then ' Note: I-JSON this is an error. AND unescaped is the final key
                        Err.Raise JERRPARSEOBJ, , "member name """ & name & """ already exists"
                    Else
                        read_json_obj.Add name, value
                    End If
                    Select Case json_ascw
                        Case ASCW_RBRACE
                            Exit Do
                        Case ASCW_COMMA
                            nextToken
                        Case Else
                            Err.Raise JERRPARSEOBJ, , "} or , expected"
                    End Select
                Case Else
                    Err.Raise JERRPARSEOBJ, , "member name expected"
            End Select
        Loop
    End If
    nextToken
End Function

Private Function read_json_arr() As Collection
    Set read_json_arr = New Collection
    nextToken
    If json_ascw <> ASCW_RBRACK Then
        Dim value
        Do
            read_json_value value
            read_json_arr.Add value
            Select Case json_ascw
                Case ASCW_RBRACK
                    Exit Do
                Case ASCW_COMMA
                    nextToken
                Case Else
                    Err.Raise JERRPARSE, , " bracket ']' or comma ',' expected in array"
            End Select
        Loop
    End If
    nextToken
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions are "Friend" for PATH

Public Sub init_string_reader(str)
    json_string = str  ' BSTR copy..
    json_pos = 0
End Sub

Public Sub nextChar()
    json_pos = json_pos + 1
    json_char = Mid$(json_string, json_pos, 1)
    If vbNullString = json_char Then
        json_ascw = -1 ' EOF
    Else
        json_ascw = AscW(json_char) And &HFFFF&  ' to 32-bit Long
    End If
End Sub

Public Sub nextToken()
    Do
        nextChar
    Loop While is_json_ws
End Sub

Public Function is_json_ws() As Boolean
    Select Case json_ascw
        Case &H20&, &H9&, &HA&, &HD&
            is_json_ws = True
    End Select
End Function

Public Function not_json_digit() As Boolean   ' Negative logic, possibly many digits. Range-check: no short-cut eval in VBA..
    Select Case json_ascw
        Case ASCW_NUM0 To ASCW_NUM9
        Case Else
            not_json_digit = True
    End Select
End Function

Public Function not_json_letter() As Boolean  ' Path expr. Negative logic, possibly many letters [a..z|A..Z]
    Select Case json_ascw And Not &H20&
        Case 65& To 90&
        Case Else
            not_json_letter = True
    End Select
End Function

Public Function check_json_string_for_word(w As String) As Boolean ' Moves ptr. (Used for null, false, true, lax, strict)
    If json_pos = InStr(json_pos, json_string, w) Then
        json_pos = json_pos + Len(w) - 1
        check_json_string_for_word = True
    End If
End Function

Public Function read_json_unquoted_string() As String ' Path expr
    If not_json_letter Then Err.Raise JERRPATH, , "invalid first character in unquoted string: " & ChrW$(json_ascw)
    Dim pos_start As Long
    pos_start = json_pos
    Do
        nextChar
        Select Case json_ascw
            Case ASCW_LBRACK, ASCW_PERIOD, -1   ' unquoted path string ends
                read_json_unquoted_string = Mid$(json_string, pos_start, json_pos - pos_start)
                json_pos = json_pos - 1 ' ungetc - just for easier code for path
                Exit Function
            Case Else
                If not_json_letter Then If not_json_digit Then Err.Raise JERRPATH, , "invalid character in unquoted string: " & ChrW(json_ascw)
        End Select
    Loop
End Function

Public Function read_json_integer() As Double  ' I-JSON thing. This exits with next non-digit read
    If not_json_digit Then Err.Raise JERRPARSE, , "numeric digit/value expected", "read_json_integer()"
    read_json_integer = json_ascw - ASCW_NUM0  ' first digit
    Do
        nextChar
        If not_json_digit Then Exit Function
        read_json_integer = read_json_integer * 10# + (json_ascw - ASCW_NUM0) ' long numbers silently lose precision
    Loop
End Function


Private Function read_json_number() As Double
    read_json_number = read_json_integer
    If ASCW_PERIOD = json_ascw Then ' fraction
        Dim start_pos As Long
        start_pos = json_pos + 1
        nextChar
        read_json_number = read_json_number + read_json_integer * 10# ^ (start_pos - json_pos) ' just make sure call evaluated first. Seems ok.
    End If
    If 69& = (json_ascw And Not &H20&) Then  ' E or e (69/101 = 45h/65h). VBA: comparison/assignment is higher prec than Logical and Bitwise Operators
        nextChar
        If ASCW_MINUS = json_ascw Then
            nextChar
            read_json_number = read_json_number * 10# ^ -read_json_integer
        Else
            If ASCW_PLUS = json_ascw Then nextChar
            read_json_number = read_json_number * 10# ^ read_json_integer
        End If
    End If
    If is_json_ws Then nextToken
End Function

Public Function read_json_string() As String ' unescaping. Default is null-ptr. So "" reads as 0 and tests for vbNullString
    Dim pos As Long, sub_start As Long
    sub_start = json_pos + 1
    Do
        nextChar
        Select Case json_ascw
            Case Is < ASCW_SPACE
                If -1 = json_ascw Then Err.Raise JERRPARSE, , "unexpected end of input in json string"
                Err.Raise JERRPARSE, , "invalid character in json string"
            Case ASCW_QUOT
                read_json_string = read_json_string & Mid$(json_string, sub_start, json_pos - sub_start)
                Exit Function
            Case ASCW_BSOL
                pos = json_pos
                read_json_string = read_json_string & Mid$(json_string, sub_start, pos - sub_start) & read_json_escape
                sub_start = json_pos + 1
        End Select
    Loop
End Function

Private Function read_json_escape() As String ' translate json to VBA host char, "a s \t d s" -> a s vbTab d s
    Dim i As Long
        nextChar
        If 117& = json_ascw Then    ' u hex hex hex hex
            Dim hexcode As Long
            Dim code As Long
            For i = 3 * 4 To 0 Step -4
                nextChar
                hexcode = InStr(HEXCHARS, json_char) - 1
                If hexcode < 0 Then Err.Raise JERRPARSE, , "invalid escape hex character \uHHHH expected"
                If hexcode > 15 Then hexcode = hexcode - 6
                code = code + hexcode * 2 ^ i
            Next i
            read_json_escape = ChrW$(code)
        Else
            i = InStr(ESCAPECHARS, json_char)
            If 0 = i Then Err.Raise JERRPARSE, , "invalid JSON escape character \" & json_char
            read_json_escape = Mid$(ESCAPECODES, i, 1)
        End If
End Function

