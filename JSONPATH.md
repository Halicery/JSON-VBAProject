# About JSON Path Expression and Syntax Relaxation 

## Unwrapping 

Unwrapping means evaluate the Path Expression for each previous result and return a possible array. 

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



