## Parse

A JSON TEXT is a sequence of TOKENS separated by optional WHITESPACES.

Tokens: 

`[ ] { } : ,` and values of numbers, strings, null/true/false literals. 

Whitespace characters: 

space (0020), horizontal tab (0009), line feed (000A) and carriage return (000D)

## JSON value string type handling

A JSON value string is a double-quoted sequence of unicode characters. 
When reading and storing in host VBA String format \ escaped characters are translated. This can be used f. eg. with \n to break text in a cell (if wrap enabled).
When writing from host certain characters will be written as \ escaped.

	                           Host VBA
	
	               parse                      toString
	JSON string  --------->  Variant/String  --------->  JSON string
	
	"..C:\\..."              "..C:\..."                  "..C:\\..."


The allowed unicode range for JSON string is 0020..FFFF. 
With escaped characters it can encode the full unicode range. 

       JSON CHAR          VBA CHAR 
       RANGE              RANGE
       0020..FFFF  ---->  0000..FFFF

More precisely:

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

.

When writing a JSON string we implement 2 options for toString (default values in parenthesis): 

- write Solidus escaped (False)
- write characters \u escaped above a certain code point (eg. 0200)

