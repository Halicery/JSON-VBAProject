Attribute VB_Name = "JSON4Excel"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Wrapper UDF-s for Excel
' Some example usage, error handling, and always return something meaningful in cells
'
' MIT License
' Copyright (c) 2024 Attila Tarpai https://github.com/Halicery/JSON-VBAProject/

Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A general API call
' Single sync HTTP GET. No VBA error handling
' Returns responseText
'
' Example:
' =jsonapi_get_response("https://dummyjson.com/test")

Function jsonapi_get_response(url As String) As String
    ' Late
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Early: add reference to Microsoft WinHttp Services 5.1 (TypeLib: %SystemRoot%\system32\winhttpcom.dll)
    'Dim objHTTP As WinHttpRequest
    'Set objHTTP = New WinHttpRequest
    objHTTP.Open "GET", url, False  ' sync
    objHTTP.Send
    If 200 = objHTTP.Status Then ' OK
        jsonapi_get_response = objHTTP.responseText
    Else
        jsonapi_get_response = Join(Array("#HTTP", objHTTP.StatusText, objHTTP.Status)) ' return some HTTP error
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JSON API call and return one or multiple values into cells
' For multiple values use array formula or new spill support
'
' Example:
' =jsonapi_and_get_values("https://dummyjson.com/products/categories", "$[3]", "$[last]")

Function jsonapi_and_get_values(url As String, ParamArray json_path_strings())
On Error GoTo dritt
    Dim resp As String
    resp = jsonapi_get_response(url)
    If "#" = Left$(resp, 1) Then
        jsonapi_and_get_values = resp ' return the HTTP error
    Else
        Dim v()
        v = json_path_strings ' arr copy, cannot pass ParamArray
        jsonapi_and_get_values = parse_and_get_values(resp, v)
    End If
    Exit Function
dritt:
    jsonapi_and_get_values = Join(Array(Err.Description, Err.Source), " - ")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Parse String (JSON TEXT) and return one or more Path Expression results
' Array-formula or with spilling
' This works only in horizontal direction of cells.
'
' Example:
' =json_parse_and_get_values(A1, "strict $.longitude", "strict $.latitude")

Function json_parse_and_get_values(json_string As String, ParamArray json_path_strings())
    Dim v()
    v = json_path_strings ' arr copy, cannot pass ParamArray
    json_parse_and_get_values = parse_and_get_values(json_string, v)
End Function


''''''''''' Private '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Returns JSON scalar as Variant value or as JSON TEXT for object/array
' TODO: return json_array as VBA Array?

Private Sub json_query(jsonroot, json_path_string As String, ByRef jsonvar) ' convert some to string and meaningful error message
On Error GoTo dritt
    json_match jsonroot, json_path_string, jsonvar
    If IsObject(jsonvar) Then
        jsonvar = json_text(jsonvar) ' with default options
    ElseIf IsEmpty(jsonvar) Then
        jsonvar = "empty" ' or CVErr(xlErrNA)? ' Empty lax expression: Excel shows 0. What to put here?
    End If
    Exit Sub
dritt:
    jsonvar = json_error
End Sub

' common parse for multiple Path expression results
' returns Variant/Array

Private Function parse_and_get_values(json_string As String, path_strings())
On Error GoTo dritt
    If UBound(path_strings) = -1 Then
        parse_and_get_values = "#path string(s) missing"
    Else
        Dim jsonroot, jsonvar
        json_parse json_string, jsonroot
        Dim i As Long
        For i = LBound(path_strings) To UBound(path_strings)
            json_query jsonroot, CStr(path_strings(i)), path_strings(i)
        Next i
        parse_and_get_values = path_strings ' reuse array for return value
    End If
    Exit Function
dritt:
    parse_and_get_values = json_error ' for parse errors
End Function

