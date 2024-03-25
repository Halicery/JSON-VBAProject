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
Sheet A1 contains the above string. To query for the last two "onclick" property in reverse order from the array of objects of menuitem, enter this this formula in A2: 

```
=json_parse_and_get_path_value(A1,"$.menu.popup.menuitem[last to last-1].onclick")
or
=json_parse_and_get_path_value(A1,"$.menu.popup.*[last to last-1].onclick")
```

The formula return this value into A2: 

```
["CloseDoc()","OpenDoc()"]
```


### GeoJSON Example

Lets say for some reason we need only the latitude values for all coordinates from a Polygon Feature (the second elements). The following GeoJSON FeatureCollection (source: OpenLayers.org) defines two Polygon Features:

```geojson
{
    "type": "FeatureCollection",
    "features": [
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1489", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-109.6875, 63.6328125], [-112.5, 35.5078125], [-85.078125, 34.8046875], [-68.90625, 39.7265625], [-68.203125, 67.1484375], [-109.6875, 63.6328125]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}},
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1668", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-40.78125, 65.0390625], [-40.078125, 34.8046875], [-12.65625, 25.6640625], [21.09375, 17.2265625], [22.5, 58.0078125], [-40.78125, 65.0390625]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}}
    ]
}
```
```json
{
    "type": "FeatureCollection",
    "features": [
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1489", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-109.6875, 63.6328125], [-112.5, 35.5078125], [-85.078125, 34.8046875], [-68.90625, 39.7265625], [-68.203125, 67.1484375], [-109.6875, 63.6328125]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}},
        {"type":"Feature", "id":"OpenLayers.Feature.Vector_1668", "properties":{}, "geometry":{"type":"Polygon", "coordinates":[[[-40.78125, 65.0390625], [-40.078125, 34.8046875], [-12.65625, 25.6640625], [21.09375, 17.2265625], [22.5, 58.0078125], [-40.78125, 65.0390625]]]}, "crs":{"type":"OGC", "properties":{"urn":"urn:ogc:def:crs:OGC:1.3:CRS84"}}}
    ]
}
```

The following formula returns an array of latitude values from the last Feature's first linear ring, omitting the last latitude value - which is the same as the first by Standard (see RFC 7946): 

```
=json_parse_and_get_path_value(A1,"$.features[last].geometry.coordinates[0][0..last-1][1]")
```

The formula returns: 

```
[65.0390625,34.8046875,25.6640625,17.2265625,58.0078125]
```

