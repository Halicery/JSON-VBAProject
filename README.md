# JSON-VBAProject

Code written in VBA Standard Modules to parse JSON API responses, query JSON data with *JSON Path Expressions* and return JSON values into cells with UDF formulae in Excel.

### The VBAProject

    VBAProject
    |
    +-- JSON4Excel
    |
    +-- Private "Lib"
        +-- JSONPARSE
        +-- JSONPATH
        +-- JSONGEN

The three Private Modules, meant to be a VBA JSON Library, are standalone and written intentionally in pure VBA independent of Excel. These are not exposed outside of the VBA Project (but are fully accessable for the Host App Excel): 

- [JSONPARSE](doc/JSONPARSE.md)  - parses JSON TEXT and stores the result in *jsonvar* Variant
- [JSONPATH](doc/JSONPATH.md) - to query any *jsonvar* Variant using JSON Path Expression Syntax
- JSONGEN - generates JSON TEXT from *jsonvar* Variant

The frontend Public Module [JSON4Excel](doc/JSON4Excel.md) contains a few wrapper UDF functions with error handling to be used in Excel formulae.

The Project references Scripting.Dictionary. 

It is not possible to distribute separately a VBA Project. The .bas Modules are under [/src](/src), which can be imported into Excel. The Excel file is under [/Workbook](/Workbook) containing the whole VBAProject with the examples below and more. It is a macro-enabled .xlsm file version 16. 



### JSON Path Expression Syntax

The Path Expression syntax supports basic dot-notation, Syntax Relaxation with automatic wrapping and unwrapping and is based on Oracle SQL/JSON Path Expression syntax. There is also strict mode that raises more errors, idea taken from MS SQL JSON syntax.

Not implemented yet:

- filter expressions
- descendant step (..)

### Example JSON query

Cell A1 contains this JSON TEXT (source: json.org):

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

Get the value of the `"id"` property:
| Formula  | `=json_parse_and_get_values(A1,"$.menu.id")`  |
|-|-|
| Cell value | file  |

Query for multiple values and use array-formula (or a simple formula with new spilling support in Excel):

<table>
<thead>
<tr>
<th>Array formula</th>
<th colspan="2"> <code> {=json_parse_and_get_values(A1,"$.menu.id","$.menu.value")} </code> </th>
</tr>
</thead>
<tbody>
<tr>
<td>Cell values</td>
<td>file</td>
<td>File</td>
</tr>
</tbody>
</table>

Query for the last two "onclick" property in reverse order from the array of objects of menuitem: 

| Formula  | `=json_parse_and_get_values(A1,"$.menu.popup.menuitem[last to last-1].onclick")`  |
|-|-|
| Cell value | ["CloseDoc()","OpenDoc()"]  |





### GeoJSON query Example 

Lets say for some reason we need only the latitude values (the second element) from all coordinates of a Polygon Feature. The following GeoJSON FeatureCollection defines two Polygon Features:

```geojson
{
    "type": "FeatureCollection",
    "features": [
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1489", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-109.6875, 63.6328125], [-112.5, 35.5078125], [-85.078125, 34.8046875], [-68.90625, 39.7265625], [-68.203125, 67.1484375], [-109.6875, 63.6328125]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}},
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1668", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-40.78125, 65.0390625], [-40.078125, 34.8046875], [-12.65625, 25.6640625], [21.09375, 17.2265625], [22.5, 58.0078125], [-40.78125, 65.0390625]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}}
    ]
}
```

The following formula returns an array of latitude values from the last Feature's first linear ring, omitting the last latitude value - which is the same as the first by Standard (see RFC 7946): 

| Formula  | `=json_parse_and_get_values(A1,"$.features[last].geometry.coordinates[0][0..last-1][1]")`  |
|-|-|
| Cell value | [65.0390625,34.8046875,25.6640625,17.2265625,58.0078125] |

The GeoJSON (source: https://openlayers.org/):

```json
{
    "type": "FeatureCollection",
    "features": [
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1489", "properties":{},
         "geometry":{"type":"Polygon", "coordinates":[[[-109.6875, 63.6328125], [-112.5, 35.5078125], [-85.078125, 34.8046875], [-68.90625, 39.7265625], [-68.203125, 67.1484375], [-109.6875, 63.6328125]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}},
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1668", "properties":{},
         "geometry":{"type":"Polygon", "coordinates":[[[-40.78125, 65.0390625], [-40.078125, 34.8046875], [-12.65625, 25.6640625], [21.09375, 17.2265625], [22.5, 58.0078125], [-40.78125, 65.0390625]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}}
    ]
}
```


