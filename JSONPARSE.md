# Parsing JSON TEXT

Conversion is rather straightforward and the implementation conforms most with I-JSON RFC 7493:

- top-Level: any JSON value
- numbers: IEEE 754 double precision (as Excel stores numbers)
- names: SHOULD be unique
- comparison: binary on unescaped strings
- RECOMMENDED to encode 64-bit integers in JSON string values
- RECOMMENDED that Time and Date data items be expressed as string values in ISO 8601 format

## Tokenization

JSON TEXT is a sequence of TOKENS separated by optional WHITESPACES.

Tokens: 

`[ ] { } : ,` and values of numbers, strings, null/true/false literals. 

Whitespace characters: 

space (0020), horizontal tab (0009), line feed (000A) and carriage return (000D)

## JSON value

	number             --> translated to VBA Variant/Double
	"string"           --> translated to VBA Variant/String
	true/false literal --> translated to VBA Variant/Boolean
	null literal       --> translated to VBA Variant/Null
	[array]            --> translated to VBA Variant/Collection
	{object}           --> translated to VBA Variant/Dictionary

- [array] is a comma-separated list of JSON values - or empty []
- {object} is a comma-separated list of name/value pairs - or empty {}


### JSON string type handling

A JSON string is a double-quoted sequence of unicode characters. The allowed unicode range for characters is 0020..FFFF. With \u escaped characters it can encode the full unicode range. 

VBA uses COM BSTR strings that store 2-byte wide characters. A character can store any 16-bit value in the full range of 0000-FFFF (there is no special null-char for string end). 

When reading and storing characters of the JSON string \ escaped characters will be translated. This can be used f. eg. with \n to break text in a cell (if wrap enabled). 

The translation process for JSON string to VBA String: 

       JSON CHAR          VBA CHAR
                  
       0000..001F  ---->  PARSE ERROR
       0020..FFFF  ---->  OK           but " (end of string) and \ has meaning:
       \"          ---->  34           
       \\          ---->  92           
       \/          ---->  47           Solidus can be escaped or not
       \b          ---->  8            Backspace: vbBack
       \t          ---->  9            Horizontal tab: vbTab
       \n          ---->  10           Line feed: vbLf
       \f          ---->  12           Form feed: vbFormFeed
       \r          ---->  13           Carriage return: vbCr
       \uxxxx      ---->  0000..FFFF


## JSON number type handling

JSONPARSE simply uses Double when parsing numbers conforming to I-JSON. Although 64-bit VBA can handle 64-bit integers using LongLong - Excel cannot. It stores all numbers as Double data type. Double has *only* 52 bits mantissa. When entering large numbers in a cell it will be truncated in a way that looses precision (i.e. keeps its magnitude but the last digits will be zeroes). This is possible because the exponent part of Double can be quite large. 







