# About JSON Path Expressions and Syntax Relaxation 

JSONPATH supports:

- basic Dot-notation syntax
- array filtering
- Syntax Relaxation between arrays and non-arrays
- automatic wrapping and unwrapping


## JSON Path Expressions 

<details>
<summary>An example JSON TEXT (source: docs.oracle.com)</summary>

```json
{ "PONumber"             : 1600,
  "Reference"            : "ABULL-20140421",
  "Requestor"            : "Alexis Bull",
  "User"                 : "ABULL",
  "CostCenter"           : "A50",
  "ShippingInstructions" : { "name"   : "Alexis Bull",
                             "Address": { "street"  : "200 Sporting Green",
                                          "city"    : "South San Francisco",
                                          "state"   : "CA",
                                          "zipCode" : 99236,
                                          "country" : "United States of America" },
                             "Phone" : [ { "type"   : "Office", 
                                           "number" : "909-555-7307" },
                                         { "type"   : "Mobile",
                                           "number" : "415-555-1234" } ] },
  "Special Instructions" : null,
  "AllowPartialShipment" : false,
  "LineItems"            : [ { "ItemNumber" : 1,
                               "Part"       : { "Description" : "One Magic Christmas",
                                                "UnitPrice"   : 19.95,
                                                "UPCCode"     : 13131092899 },
                               "Quantity"   : 9.0 },
                             { "ItemNumber" : 2,
                               "Part"       : { "Description" : "Lethal Weapon",
                                                "UnitPrice"   : 19.95,
                                                "UPCCode"     : 85391628927 },
                               "Quantity"   : 5.0 } ] }
```
</details>



## The Dot-notation syntax

The context item is represented by the dollar sign ($), which is the JSON TEXT itself in this case. Dot-notation consists of one or more field names separated by periods (.). The expression *walks* JSON objects by property names: 

<details>
<summary>Every dot-notation values</summary>

```
JSON DATA                              JSON PATH EXPRESSION

object                                 $
 `-- number                            $.PONumber
 `-- string                            $.Reference
 `-- string                            $.Requestor
 `-- string                            $.User
 `-- string                            $.CostCenter
 `-- object                            $.ShippingInstructions
 '    `-- string                       $.ShippingInstructions.name
 '    `-- object                       $.ShippingInstructions.Address
 '    '    `-- string                  $.ShippingInstructions.Address.street
 '    '    `-- string                  $.ShippingInstructions.Address.city
 '    '    `-- string                  $.ShippingInstructions.Address.state
 '    '    `-- number                  $.ShippingInstructions.Address.zipCode
 '    '    `-- string                  $.ShippingInstructions.Address.country
 '    `-- array                        $.ShippingInstructions.Phone
 `-- null                              $.Special Instructions
 `-- false                             $.AllowPartialShipment
 `-- array                             $.LineItems
```
</details>


## Array indexes

Extending Path Expression with array indices on arrays all JSON value has their unique path: 

<details>
<summary>Every absolute path values</summary>

```
JSON DATA                              JSON PATH EXPRESSION

object                                 $
 `-- number                            $.PONumber
 `-- string                            $.Reference
 `-- string                            $.Requestor
 `-- string                            $.User
 `-- string                            $.CostCenter
 `-- object                            $.ShippingInstructions
 '    `-- string                       $.ShippingInstructions.name
 '    `-- object                       $.ShippingInstructions.Address
 '    '    `-- string                  $.ShippingInstructions.Address.street
 '    '    `-- string                  $.ShippingInstructions.Address.city
 '    '    `-- string                  $.ShippingInstructions.Address.state
 '    '    `-- number                  $.ShippingInstructions.Address.zipCode
 '    '    `-- string                  $.ShippingInstructions.Address.country
 '    `-- array                        $.ShippingInstructions.Phone
 '         `-- object                  $.ShippingInstructions.Phone[0]
 '         '    `-- string             $.ShippingInstructions.Phone[0].type
 '         '    `-- string             $.ShippingInstructions.Phone[0].number
 '         `-- object                  $.ShippingInstructions.Phone[1]
 '              `-- string             $.ShippingInstructions.Phone[1].type
 '              `-- string             $.ShippingInstructions.Phone[1].number
 `-- null                              $.Special Instructions
 `-- false                             $.AllowPartialShipment
 `-- array                             $.LineItems
      `-- object                       $.LineItems[0]
      '    `-- number                  $.LineItems[0].ItemNumber
      '    `-- object                  $.LineItems[0].Part
      '    '    `-- string             $.LineItems[0].Part.Description
      '    '    `-- number             $.LineItems[0].Part.UnitPrice
      '    '    `-- number             $.LineItems[0].Part.UPCCode
      '    `-- number                  $.LineItems[0].Quantity
      `-- object                       $.LineItems[1]
           `-- number                  $.LineItems[1].ItemNumber
           `-- object                  $.LineItems[1].Part
           '    `-- string             $.LineItems[1].Part.Description
           '    `-- number             $.LineItems[1].Part.UnitPrice
           '    `-- number             $.LineItems[1].Part.UPCCode
           `-- number                  $.LineItems[1].Quantity

```
</details>


These are basic and absolute path expressions, which return a single JSON value as it is stored in the JSON data.


## Array filtering with Path Expressions

The array step accepts more than only a single index value in brackets `[]` and it is called index-specifier. It can specify one or more array indices and can return multiple values in an array: 

- asterisk (*) wildcard: represents all elements
- index value returns a single element: 0, 5, last, last - 1
- index ranges: 1 to 5 or 1..5 (supported only in JSONPATH)
- a comma separated list of index values and index ranges

```
index-specifier = [*] | [idxval|idxrange (,...)]
idxval          = number|last|last-number
idxrange        = idxval to|.. idxval
```




## Automatic Unwrapping 

Unwrapping means iteration: evaluate the Path Expression for each previous result and return a possible array. 

## Array unwrapping

Especially to handle 2D arrays. Consider an array of arrays:  

	[ [1,2], [3,4], [5,6] ] 

What should be the result of the Path Expression `$[*][0]`? 

To be consistent in syntax, when `$[1][0] = 3` then `$[*][0]` should be the first element of each sub-arrays, i.e. `[1,3,5]` - and not [1,2]. 

This is because `[*]` is not the same array as the original in structure. 


It is implemented by unwrapped arrays. 

[\*] and [multi-index] array specification might return such unwrapped arrays, which can have different meaning in different contexts: 

	$[*] --> [ [1,2], [3,4], [5,6] ]    here no difference

In unwrapped form the same array looks like this: 

	        [1,2]
	[*] = [ [3,4] ]       and not [ [1,2], [3,4], [5,6] ] 
	        [5,6]


***[\*] transposes the array vector***

So the $[*][0] operation will be:
        
	            [1,2]          [1,2][0]       1        
	$[*][0] = [ [3,4] ][0] = [ [3,4][0] ] = [ 3 ] = [1,3,5]
	            [5,6]          [5,6][0]       5        


### Syntax Relaxation - wrapping and unwrapping object example

To get the name property from this object the Path Expression would be `$.name` and the result is `"n1"`: 

	{ "name":"n1" }  
	                    
	$.name = "n1"

When object evolves to an array of objects Syntax Relaxation allows to use the same Path Expression, without breaking code. It will perform unwrapping and return the name value for each object: 

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$.name = [ "n1", "n2","n3" ]       <-- unwrapping

When querying later, one might begin to write new Path Expression-s with the array form:

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$[*].name = [ "n1", "n2","n3" ]

The array form should still work on old elements. Automatic wrapping is performed: 

	{ "name":"n1" }  
	                    
	$[*].name = "n1"                   <-- wrapping

This type of Syntax Relaxation means that [*] can be inserted or can be ommitted before the dot (.) operator without changing the match result:

	$.name  <-- -->  $[*].name

### Unwrapping 

The implementation for unwrapping uses transposed arrays and for each: 

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$.name = $[*].name =
	
	    { "name":"n1" }            { "name":"n1" }.name       "n1"     
	= [ { "name":"n2" } ].name = [ { "name":"n2" }.name ] = [ "n2" ] = [ "n1", "n2","n3" ]
	    { "name":"n3" }            { "name":"n3" }.name       "n3"    

### Wrapping 

The implementation for wrapping is fairly simple: 

	{ "name":"n1" }  
	                    
	$[*].name = [ { "name":"n1" } ].name = [ { "name":"n1" }.name ] = [ "n1" ] = "n1"

As any one-element array result is automatically unwrapped. 

This is to be consistent with single array index Path Expression filtering, f. ex.:

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$[0].name = { "name":"n1" }.name = "n1"



