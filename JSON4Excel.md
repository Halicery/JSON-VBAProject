

|                             | URL     |    json_text    |     Path-expr  |      Value
|-|-|-|-|-|
|jsonapi_get_response         | +       |    =            |                |
|jsonapi_and_get_values       | +       |                 |     +          |      =
|json_parse_and_get_values    |         |    +            |     +          |      =

`+` means parameter

`=` means return value



                        
jsonapi_get_response
takes a URL and returns responseText, the JSON TEXT 

jsonapi_and_get_values
takes a URL and fetches the JSON TEXT
parses JSON
runs each Path-expr parameter and returns a Variant (single Path-expr) or a Variant Array (multiple). 

For multiple values use array formula or new spill support. 

json_parse_and_get_values
Parses json text

### Example API

Imagine we have a list of Product id-s in a Sheet and we'd like to get the product's brand name using a JSON API. This example is based on **Get a single product** API from https://dummyjson.com/. 

The resource address is https://dummyjson.com/products/{id} and the response contains one JSON object. We query the brand name using `$.brand` Path Expression and the formula is: 

=jsonapi_and_get_values(B$13 & A15,"$.brand")


|Id	|https://dummyjson.com/products/||
|-|-|-|
|1	|Apple
|8	|Microsoft Surface
|31	|Furniture Bed Set
|1233	|#HTTP Not Found 404


![image](https://github.com/Halicery/JSON-VBAProject/assets/419722/9e57d045-8bfd-4772-8435-48f770e7d489)

