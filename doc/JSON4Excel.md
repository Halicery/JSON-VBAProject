# JSON4Excel.bas

Essentially 3 Public UDF functions:


|`+` means parameter<br>`=` means return value         | URL     |    json_text    |     Path-expr  |      Value|
|:-|-|-|-|-|
|jsonapi_get_response()       | +       |    =            |                | |
|json_parse_and_get_values()  |         |    +            |     +          |      =|
|jsonapi_and_get_values()     | +       |                 |     +          |      =|


**`=jsonapi_get_response(url)`**

Takes a URL and returns HTTP responseText, the JSON TEXT in a cell. 

**`=json_parse_and_get_values(json_text, path, ...)`**

Takes a JSON TEXT (from a cell or as a String parameter f. ex.) and runs parse. Then runs each Path-expr parameter against the JSON value. Can be used to return one or multiple JSON value results. 

For multiple values use array formula or new spill support in Excel.

**`=jsonapi_and_get_values(url, path, ...)`**

Shortcut for the above two functions - without revealing JSON TEXT. Takes a URL and fetches JSON TEXT. Parses JSON TEXT and runs each Path-expr parameter against the JSON value. Returns one or multiple JSON value results. 





Same as the API version but the input is String containing JSON TEXT. 







### Example API

Imagine we have a list of Product id-s in a Sheet and we'd like to get the product's brand name and price using a JSON API. This example is based on **Get a single product** API from https://dummyjson.com/. 

||A|B|C|
|:-|-:|-|-|
|1|*Id*|*Brand name*|*Price*
2|*7*|
3|*10*|
4|*31*|
5|*1233*|

The resource address is https://dummyjson.com/products/{id} and the response contains one JSON object. We query the brand name using `$.brand` and the price with `$.price` Path Expression. 

The formulae (using new spill support of Excel):

||A|B|C|
|:-|-:|-|-|
1|*Id*|*Brand name*|*Price*
2|*7*|=jsonapi_and_get_values("https://dummyjson.com/products/" & A2,"$.brand","$.price")
3|*10*|=jsonapi_and_get_values("https://dummyjson.com/products/" & A3,"$.brand","$.price")
4|*31*| or use fill..
5|*1233*|

And the result on success is: 

||A|B|C|
|:-|-:|-|-|
1|*Id*|*Brand name*|*Price*
2|7|Samsung|	1499
3|10|HP Pavilion	|1099
4|31|Furniture Bed Set|	40
5|1233|#HTTP Not Found 404	

For the last one a non-existent id is used to test HTTP error. 


