# JSON-VBAProject

VBA code written mainly to parse JSON API responses, query JSON data using JSON Path Expressions and return values into cells for Excel.

### Parsing 

Conversion is rather straightforward and the implementation conforms most with I-JSON RFC 7493:

- top-Level: any JSON value
- numbers: IEEE 754 double precision (as Excel stores numbers)
- names: SHOULD be unique
- comparison: binary on unescaped strings
- RECOMMENDED to encode 64-bit integers in JSON string values
- RECOMMENDED that Time and Date data items be expressed as string values in ISO 8601 format

### JSON Path Expression Syntax

The Path Expression syntax supports basic dot-notation, Syntax Relaxation with automatic wrapping and unwrapping and is based on Oracle SQL/JSON Path Expression syntax. There is also strict mode that raises more errors, idea taken from MS SQL SQL/JSON syntax.

Not implemented yet:

- filter expressions
- descendant step (..)

### Example JSON query

A JSON TEXT (source: json.org):

```json
{"menu": {
  "id": "file",
  "value": "File",
  "popup": {
    "menuitem": [
      {"value": "New", "onclick": "CreateNewDoc()"},
      {"value": "Open", "onclick": "OpenDoc()"},
      {"value": "Close", "onclick": "CloseDoc()"}
    ]
  }
}}
```
In a Sheet A1 contains the above string. To query for the last two "onclick" property from the array of objects of menuitem, use this formula: 

```
=json_parse_and_get_path_value(A1,"$.menu.popup.menuitem[last-1 to last].onclick")
or
=json_parse_and_get_path_value(A1,"$.menu.popup.*[last-1..last].onclick")
```

```
["OpenDoc()","CloseDoc()"]
```



It works with VBA Variants so this version is not for building and/or manipulating a JSON object tree. 


One Public Module that contains some example wrapper UDF with error handling to be used in Excel. 

The three Private Modules are standalone and independent of Excel



+-- Private --+
| JSONPARSE 
| JSONPATH
| JSONGEN
+-------------

Purpose to implement the skeleton structure first for parse and path expressions.



