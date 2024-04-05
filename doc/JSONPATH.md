# JSONPATH.BAS

JSONPATH supports:

- basic dot-notation syntax
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



### The Dot-notation syntax

The context item is represented by the dollar sign ($), which is the JSON TEXT itself in this case. Dot-notation consists of one or more field names separated by periods (.). The expression *walks* JSON objects along property names. The result of the expression is either: 

- another JSON object
- JSON array value
- JSON scalar value (string, number or literal)


<details>
<summary>Every dot-notation expressions</summary>

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


### Array indexes

Extending Path Expression with array indices on arrays every JSON value will have a unique path. The basic Path Expression syntax thus consist of: 

- object-step on objects yields the named member value
- array-step on arrays yields the n<sup>th</sup> array member value 

Every absolute path values:

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



These are basic and absolute path expressions, which return a single JSON value as it is stored in the JSON data.


## Syntax Relaxation

In the relaxed syntax form the purpose of the JSON Path Expression is extended beyond returning a single JSON value. 

It is a matching process: a JSON value ($) is matched against a Path Expression, which can yield: 

- a single JSON value
- multiple JSON values
- no value, empty (i.e. no match is found)

It is also involves:

- array filtering
- unwrapping: to allow object-step on non-objects
- wrapping: to allow array-step on non-arrays
- empty array [] or empty object {} returns empty value
- unboxing: array with one element returns the element itself

The last two is debatable. 

### Array filtering with Path Expressions

The array step accepts more than only a single index value in brackets `[]` and it is called index-specifier. It can specify one or more array indices and can return multiple values in an array: 

- asterisk (*) wildcard: represents all elements
- index value returns a single element: 0, 5, last, last - 1
- index ranges: 1 to 5 or 1..5 (the `..` form is supported only in JSONPATH)
- a comma separated list of index values and index ranges

```
index-specifier = [*] | [idxval|idxrange (,...)]
idxval          = number|last|last-number
idxrange        = idxval to|.. idxval
```

In the relaxed syntax form an index value out of range is simply no-match and the result is empty. 

### Unwrapping 

Unwrapping means iteration: to evaluate the Path Expression for each previous result. It can happen in the relaxed syntax form, when the previous result is an array - but the expression expects a single value. It is an automatic process, which can yield multiple results in an array. 

In addition, JSONPATH implements unwrapping on consecutive array steps as well. 

Especially to handle 2D arrays. Consider an array of arrays:  

	[ [1,2], [3,4], [5,6] ] 

What should be the result of the Path Expression `$[*][0]`? 

To be consistent in syntax, when `$[1][0] = 3` then `$[*][0]` should be the first element of each sub-arrays, i.e. `[1,3,5]` - and not [1,2]. 

This is because `[*]`, or any other multi-index specifier makes a selection and returns the result in an array that is different in structure. 

It is implemented by unwrapped arrays. 

In unwrapped form the array looks like this internally, ***a transposed vector***: 

	        [1,2]
	[*] = [ [3,4] ] 
	        [5,6]

And the $[*][0] operation will be:
        
	            [1,2]          [1,2][0]       1        
	$[*][0] = [ [3,4] ][0] = [ [3,4][0] ] = [ 3 ] = [1,3,5]
	            [5,6]          [5,6][0]       5        

Note that if `[*]` is the last step there is no difference in result: 

	$[*] --> [ [1,2], [3,4], [5,6] ]   


### Syntax Relaxation - wrapping and unwrapping object example

The true purpose of Syntax Relaxation with wrapping/unwrapping is to allow JSON data to evolve, while using the same query Path Expressions without breaking code. 

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

### Unwrapping implementation

The implementation for unwrapping uses transposed vectors and for each: 

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$.name = $[*].name =
	
	    { "name":"n1" }            { "name":"n1" }.name       "n1"     
	= [ { "name":"n2" } ].name = [ { "name":"n2" }.name ] = [ "n2" ] = [ "n1", "n2","n3" ]
	    { "name":"n3" }            { "name":"n3" }.name       "n3"    

### Wrapping implementation

The implementation for wrapping is fairly simple: 

	{ "name":"n1" }  
	                    
	$[*].name = [ { "name":"n1" } ].name = [ { "name":"n1" }.name ] = [ "n1" ] = "n1"

As any one-element array result is automatically unboxed. 

This is to be consistent with single array index Path Expression filtering, f. ex.:

	[ { "name":"n1" }, { "name":"n2" }, { "name":"n3" } ]
	
	$[0].name = { "name":"n1" }.name = "n1"



