Attribute VB_Name = "JSONPATH"
''''''''''''''''''''''' JSON PATH EXPRESSIONS ''''''''''''''''''''''''''''''''''''''
'
' JSON Path Expression syntax with Syntax Relaxation, automatic wrapping and unwrapping
' Based on Oracle SQL/JSON Path Expressions and MS SQL (strict/lax).
'
' MIT License
' Copyright (c) 2024 Attila Tarpai https://github.com/Halicery/JSON-VBAProject/

Option Private Module
Option Explicit

Private Type PATHRANGER ' Our fake iterator for index-specifier: has first() and next()
    idx As Long
    to As Long
    last As Long
    dir As Long ' for reverse iteration
End Type

Private PR As PATHRANGER

Private strict As Boolean


''''''''''' Public '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Matches Path Expression against a JSON value, the context item, represented by the dollar sign ($).
' The match either fails or results in one or more JSON values.
' The default mode is the relaxed syntax form, lax.
' Automatic wrapping and unwrapping for obj and arr steps against non-obj and non-arr values
' No match returns Empty and processing ends.
' In strict mode a match should be found otherwise error raised
'
' Not implemented yet:
' - filter expressions
' - descendant step (..)
'
' Example
' "$.person[last].age"
' "strict $.person[last].age"
' "lax $.person[last].age"

Public Sub json_match(jsonvar, path_expr As String, ByRef jsonvar_out As Variant)
    strict = False
    init_string_reader path_expr
    nextChar    ' lax, strict or $
    If ASCW_DOLLAR <> json_ascw Then
        If check_json_string_for_word("lax $") Then
        ElseIf check_json_string_for_word("strict $") Then
            strict = True
        Else
            Err.Raise JERRPATH, , "must start with lax, strict or $"
        End If
    End If
    set_variant jsonvar_out, jsonvar
    select_node jsonvar_out
    If Not strict Then
        If IsObject(jsonvar_out) Then
            If 2 > jsonvar_out.Count Then
                If 1 = jsonvar_out.Count Then ' unbox
                    If "Collection" = TypeName(jsonvar_out) Then set_variant jsonvar_out, jsonvar_out(1) ' final unboxing
                Else
                    jsonvar_out = Empty ' match was {} or []
                End If
            End If
        End If
    End If
End Sub


''''''''''' Private '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub set_variant(variable, value)
    If IsObject(value) Then
        Set variable = value
    Else
        Let variable = value
    End If
End Sub


' Index-specifier syntax:
'
' index-specifier = [*] | [idxval|idxrange (,...)]
' idxval          = number|last|last-number
' idxrange        = idxval to|.. idxval
'
' The index iterator supports:
' - wildcard [*]
' - index list with comma
' - index ranges with either 'to' or '..'
' - reverse iteration like [5..1] in that order
' Ex: [2] [*] [2,5,6] [2..6]=[2 to 6] [last] [3 to last-2] [last..0] [last-1, 3, 6]
' NB. Duplicates will be added.

Private Sub PR_first()    ' Reads idxval or idxrange. Inits the "iterator"
    PR.idx = read_idxval
    If check_json_string_for_word("to") Then  ' no short-cut eval in VBA
    ElseIf check_json_string_for_word("..") Then ' we also support 2..3
    Else
        PR.to = PR.idx ' single idx, set end cond for next
        Exit Sub
    End If
    ' range:
    nextToken
    PR.to = read_idxval
    If PR.to < PR.idx Then PR.dir = -1 Else PR.dir = 1  ' set reverse
End Sub

Private Sub PR_next()
    If PR.to = PR.idx Then  ' last, done
        Select Case json_ascw
            Case ASCW_RBRACK
                 PR.idx = &H80000000  ' end-condition
            Case ASCW_COMMA
                nextToken
                PR_first
            Case Else
                Err.Raise JERRPATH, , "invalid array-specifier: ] or , expected"
        End Select
    Else
        PR.idx = PR.idx + PR.dir  ' +/-1
    End If
End Sub

Private Function read_idxval() As Long  ' reads number|last|last-number
    If not_json_digit Then  ' check, read_json_integer may raise error
        If check_json_string_for_word("last") Then
            nextToken
            If ASCW_MINUS = json_ascw Then
                nextToken
                If not_json_digit Then Err.Raise JERRPATH, , "invalid array-specifier, last-number expected"
                read_idxval = PR.last - read_json_integer_ws ' can be minus
                'If read_idxval < -1 Then read_idxval = -1 ' skip excess?
            Else
                read_idxval = PR.last
            End If
        Else
            Err.Raise JERRPATH, , "invalid index-specifier, number or last expected"
        End If
    Else
        read_idxval = read_json_integer_ws
        'If read_idxval > PR.last + 1 Then read_idxval = PR.last + 1 ' skip excess?
    End If
End Function

Private Function read_json_integer_ws() As Long ' read_json_integer exits with first non-digit read
    read_json_integer_ws = read_json_integer
    If is_json_ws Then nextToken
End Function

' 1. object step is a period (.) followed by an object field name or an asterisk (*) wildcard
' 1.1 object step on object
' 1.1.1 name match --> value | empty
' 1.1.2 asterisk match --> [value,..] (relaxation and unwrapping) | value | empty
' 1.2 object step on non-object: syntax relaxation
' 1.2.1 object step on array: syntax relaxation and unwrapping
' 1.2.2 object step on scalar: what now? --> empty
' 2. array step is a left bracket ([) followed by either an asterisk (*) wildcard one or more specific array indexes or range specifications separated by commas
' 2.1 array step on array --> [value,..] | value | empty (NB: [value,..] will be unwrapped if followed by array-step)
' 2.2 array step on non array: implicitly wrapped in a one-element array and relax
' 3. descendant step is two consecutive periods (..) followed by a field name (NOT IMPLEMENTED)

Private Sub select_node(ByRef jsonvar)
    Dim coll As Collection
    Dim it, wrap_coll As Collection ' for unwrapping and for each: a "transposed" vector
    Do
        nextChar
        Select Case json_ascw
            Case -1
                Exit Sub
                
            Case ASCW_LBRACK
                ' 2. array step
                If "Collection" = TypeName(jsonvar) Then
                    ' 2.1 array step on array
                    If wrap_coll Is Nothing Then
                        Set coll = jsonvar
                    Else
                        Exit Do ' unwrap
                    End If
                Else
                    If strict Then Err.Raise JERRPATH, , "array-step on non-array"
                    ' 2.2 array step on non array: wrap and relax
                    Set coll = New Collection
                    coll.Add jsonvar
                End If
                ' read index-specifier
                nextToken
                If ASCW_AST = json_ascw Then
                    nextToken
                    If ASCW_RBRACK <> json_ascw Then Err.Raise JERRPATH, , "invalid index-specifier: [*] expected"
                    Set wrap_coll = coll
                Else
                    Set wrap_coll = New Collection
                    PR.last = coll.Count - 1
                    PR_first
                    Do
                        Select Case PR.idx
                            Case 0 To PR.last
                                wrap_coll.Add coll(PR.idx + 1)
                            Case Else
                                If strict Then Err.Raise JERRPATH, , "index out of bound"
                        End Select
                        PR_next
                    Loop Until PR.idx = &H80000000
                End If
                GoSub unbox
           
            Case ASCW_PERIOD
                ' 1. object step
                If "Dictionary" = TypeName(jsonvar) Then
                    ' 1.1 object step on object
                    Dim dict As Scripting.Dictionary
                    Set dict = jsonvar
                    nextChar
                    If ASCW_AST = json_ascw Then
                        ' 1.1.2 asterisk match
                        Set wrap_coll = New Collection
                        For Each it In dict.Items
                            wrap_coll.Add it
                        Next it
                        GoSub unbox
                        Set wrap_coll = Nothing ' don't unwrap f.ex. $.*[0]
                    Else
                        ' 1.1.1 name match
                        Dim name As String
                        If ASCW_QUOT = json_ascw Then
                            name = read_json_string ' unescaped string
                        Else
                            name = read_json_unquoted_string
                        End If
                        If dict.Exists(name) Then
                            set_variant jsonvar, dict(name)
                        Else
                            If strict Then Err.Raise JERRPATH, , "object member """ & name & """ not exists"
                            GoTo noMatch
                        End If
                    End If
                Else
                    If strict Then Err.Raise JERRPATH, , "object-step on non-object"
                    ' 1.2 object step on non-object: syntax relaxation
                    If "Collection" = TypeName(jsonvar) Then
                        ' 1.2.1 object step on array (obj evolves to array of objects) ALLOW IN strict mode?
                        Exit Do ' unwrap
                    Else
                        ' 1.2.2 object step on scalar: empty? or scalar?
                        GoTo noMatch
                    End If
                End If
                
            Case Else '
                Err.Raise JERRPATH, , "invalid operator or character: " & json_char
        End Select
    Loop
    
    ' unwrapping
    Set coll = jsonvar
    Set wrap_coll = New Collection
    Dim pos As Long
    pos = json_pos - 1 ' ungetc
    For Each it In coll
        json_pos = pos
        select_node it
        If Not IsEmpty(it) Then wrap_coll.Add it
    Next it
    GoSub unbox
    Exit Sub

unbox:
        If 0 = wrap_coll.Count Then
            If strict Then Err.Raise JERRPATH, , "empty array"
            GoTo noMatch
        ElseIf 1 = wrap_coll.Count Then
            set_variant jsonvar, wrap_coll(1) ' single
            Set wrap_coll = Nothing ' unwrapping off
        Else
            Set jsonvar = wrap_coll ' unwrapping on next step
        End If
        Return
        
noMatch:
        jsonvar = Empty
End Sub

