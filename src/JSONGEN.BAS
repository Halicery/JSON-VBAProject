Attribute VB_Name = "JSONGEN"
''''''''''''''''''  JSON TEXT generator  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   jsonvar Variant --> JSON TEXT
'
'   Two options:
'   - write solidus escaped (default false)
'   - write high charcodes \u escaped (like above &H300)
'
' MIT License
' Copyright (c) 2024 Attila Tarpai https://github.com/Halicery/JSON-VBAProject/

Option Private Module

Option Explicit
Option Compare Binary

Private Const ESCAPED_CHARS_WITH_SOL = "/""\"
Private Const ESCAPED_CHARS_WITHOUT_SOL = """\"
Private Const DEFAULT_MAX_ASCW As Long = &H300&

Private Type GENERATOROPTIONS
    escaped_chars As String
    max_ascw As Long
End Type

Private GO As GENERATOROPTIONS  ' make it static


''''''''''' Public '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function json_text(jsonvar, Optional write_solidus_escaped As Boolean, Optional max_ascw_unescaped As Long = -1) As String
    genoption_write_solidus_escaped = write_solidus_escaped
    genoption_max_ascw_unescaped = max_ascw_unescaped
    json_text = write_jsonvar(jsonvar)
End Function



''''''''''' Private/Friend '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Let genoption_write_solidus_escaped(v As Boolean)
    If v Then
        GO.escaped_chars = ESCAPED_CHARS_WITH_SOL
    Else
        GO.escaped_chars = ESCAPED_CHARS_WITHOUT_SOL
    End If
End Property

Property Let genoption_max_ascw_unescaped(ascw_code As Long) ' above write with \uxxxx
    If ascw_code And &HFFFF0000 Then
        GO.max_ascw = DEFAULT_MAX_ASCW
    Else
        GO.max_ascw = ascw_code ' 0-65535
    End If
End Property


Public Function write_jsonvar(jsonvar) As String ' recurs
    If IsObject(jsonvar) Then
        Dim it
        If "Dictionary" = TypeName(jsonvar) Then
            For Each it In jsonvar.Keys               ' keys are Variant/String
                write_jsonvar = write_jsonvar & """" & write_json_string(it) & """:" & write_jsonvar(jsonvar(it)) & ","
            Next it
            quote_list write_jsonvar, "{}"
        ElseIf "Collection" = TypeName(jsonvar) Then
            For Each it In jsonvar
                write_jsonvar = write_jsonvar & write_jsonvar(it) & ","
            Next it
            quote_list write_jsonvar, "[]"
        Else
            write_jsonvar = TypeName(jsonvar) ' should not happen
        End If
    Else
        write_jsonvar = write_json_scalar(jsonvar)
    End If
End Function

Private Sub quote_list(ByRef list As String, quote As String)
    If vbNullString = list Then
        list = quote
    Else
        list = Left$(quote, 1) & Mid$(list, 1, Len(list) - 1) & Right$(quote, 1)
    End If
End Sub

Public Function write_json_scalar(jsonvar) As String
    Select Case VarType(jsonvar)
        Case vbString
            write_json_scalar = """" & write_json_string(jsonvar) & """"
        Case vbNull
            write_json_scalar = "null" ' cannot LCase vbNull
        Case vbBoolean
            write_json_scalar = LCase(jsonvar)
        Case vbDouble
            write_json_scalar = Replace(jsonvar, ",", ".") ' Quick Locale fix. TODO different number formatting options like 1.01e3?
        Case Else   ' vbEmpty (lax)
            write_json_scalar = "empty"
    End Select
End Function

Private Function write_json_string(jsonvar) As String     ' eg. a s vbTab d s --> "a s \t d s"
    Dim ch As String, idx As Long
    Dim pos As Long, sub_start As Long
    sub_start = 1
    For pos = 1 To Len(jsonvar)
        ch = Mid$(jsonvar, pos, 1)
        Select Case AscW(ch) And &HFFFF&
            Case ASCW_SPACE To GO.max_ascw
                If InStr(GO.escaped_chars, ch) Then GoTo 9
            Case Else
                idx = InStr(4, ESCAPECODES, ch)  ' btnfr
                If idx Then
                    ch = Mid$(ESCAPECHARS, idx, 1)
                Else
                    ch = Hex$(AscW(ch))
                    ch = Left$("u0000", 5 - Len(ch)) & ch
                End If
9               write_json_string = write_json_string & Mid$(jsonvar, sub_start, pos - sub_start) & "\" & ch
                sub_start = pos + 1
        End Select
    Next pos
    write_json_string = write_json_string & Mid$(jsonvar, sub_start) ' rest
End Function

