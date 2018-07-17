## querPy

An extendable script for executing multiple queries against a SPARQL-endpoint of your choice, returning the result-data either in different data formats (csv, tsv, xml, json, xslx) to be saved locally or uploaded as a google sheets files into a google folder or inserted into existing google sheets file. Additionally anytime it is executed it also creates a summary (as a file if saved locally, or as a page if saved into an xslx or google sheets), wherein the original sparql-queries are included, their execution times, their total number of results, and a few sample result lines.

There is no fancyness at all to this script; it just provides the core logic for the described purpose, in a minimalistic manner in order to be extensible for any kind of interface to be wrapped around it. 


### dependencies

The script was written in python3, no downward compability to python2.x was tested.

The script brings in few dependencies: 
##### three external libaries:
* SPARQLWrapper: https://github.com/RDFLib/sparqlwrapper
* google-api-python-client: https://github.com/google/google-api-python-client
* regex module: https://pypi.org/project/regex/ (not the default one, but with improved capabilities)
##### google OAuth2 credentials (their API requires such)

#### External libaries

The external libraries you can install by running:

```
pip install SPARQLWrapper
pip install google-api-python-client
pip install regex
```

#### Google OAuth2 credentials

Only when writing into google sheets or folders, you need to provide two files for google to process the traffic via its API:
* client_secret.json (basically authenticating the script as a service)
* credentials.json (authenticating the script to act on a user's behalf, and also to write into his/her private google drive)


To obtain a client_secret.json file you must log into the google developer console, register a project, and download the secrets-file, as outlined here:
https://developers.google.com/drive/api/v3/quickstart/python

To obtain a credentials.json file you simply provide the querPy script the client_secrets.json file (either as explicit argument '-s client_secret.json' or just put it into the folder wherer querPy is saved into). Then when running the querPy script, a browser will popup and you will be asked to authorize the script.

If you want to save the results as local files only, you don't need to obtain these credential files. 


### running querPy

To run, you would isse the following command (wherein 'template.py' refers to a file containing sparql-queries)
```
python querPy.py -r template.py
```



### structure of the queries file

To create a template you can run:
```
python querPy.py -t
```

After which you would find a template file in your folder. The file is itself a python module (due to problems having arisen when using other popular formats, such as json doesn't allow multilines content (annoying when writing sparql-queries) and xml can't be used due to '<' being a meta-character but sparql queries can contain such). 

Within the file there are several variables (most of which are actually optional):

#### title
defines the title of the whole set of queries

OPTIONAL, if not set, timestamp will be used

#### description
defines the textual and human-intended description of the purpose of these queries

OPTIONAL, if not set, nothing will be used or displayed

#### output_destination
defines where to save the results, input can be: 

* a local path to a folder 

* a URL for a google sheets document  

* a URL for a google drive folder

NOTE: On windows, folders in a path use backslashes, in such a case it is mandatory to attach a 'r' in front of the quotes, e.g. r"C:\Users\sresch\.."
In the other cases the 'r' is simply ignored; thus best would be to always leave it there.

OPTIONAL, if not set, folder of executed script will be used

#### output_format
defines the format in which the result data shall be saved (currently available: csv, tsv, xml, json, xlsx)

OPTIONAL, if not set, csv will be used

#### summary_sample_limit
defines how many rows shall be displayed in the summary

OPTIONAL, if not set, 5 will be used

#### cooldown_between_queries
defines how many seconds should be waited between execution of individual queries in order to prevent exhaustion of Google API due to too many writes per time-interval

OPTIONAL, if not set, 0 will be used

#### endpoint
defines the SPARQL endpoint against which all the queries are run

MANDATORY

#### queries
defines the set of queries to be run. 

MANDATAORY

##### query

###### title

OPTIONAL, if not set, timestamp will be used

###### description

OPTIONAL, if not set, nothing will be used or displayed

###### query
the sparql query itself

MANDATORY


### multi values in querPy

Additionally, there exists also the feature of inserting multiple values into any field of a query collections file. Such multi values are defined as lists embedded in a bigger list. Then querPy will create from such lists individiual fields for each, e.g. one could define 

```
output_format = ["csv","xml"]
```

And querPy would create then collection files with all their contents identical except for one having csv and the other xml as defined output_format, which can be useful if one wants to have multiple executions of queries which are very similar to each other except for minor differences which can be encoded in such a way. 

Multi values can also be used inside of strings, e.g. inside sparql queries (where again the whole query-text would be a list with some string first, then the list of multi values, then some string again):


```
...
"query" : [ "SELECT * WHERE { ?s <http://www.w3.org/2000/01/rdf-schema#", ["label", "type"], "> ?o}" ]
...
```

Then querPy would create from this two files with identical content except for these two different queries:

file 1:
```
...
"query" : [ "SELECT * WHERE { ?s <http://www.w3.org/2000/01/rdf-schema#label> ?o}" ]
...
```
file 2:
```
...
"query" : [ "SELECT * WHERE { ?s <http://www.w3.org/2000/01/rdf-schema#type> ?o}" ]
...
```

Multi values can also be used multiple times themselves, e.g. one could define the output_format and corresponding queries like this:


```
...
output_format = ["csv","xml"]
...
"query" : [ "SELECT * WHERE { ?s <http://www.w3.org/2000/01/rdf-schema#", ["label", "type"], "> ?o}" ]
...
```

querPy would always construct new files by joing the elements in the lists with the same index together, e.g. the first element of the first list with the first element of the second list, et cetera. Thus querPy creates two new files in the example above, where the first has 'csv' as output_format and 'label' used inside its query, while the second would have 'xml' and 'type'. 

Should there be a mismatch between the number of elements of the lists used, querPy will abort.

Additionally, since the query collection file is itself a python module, instead of defining bare lists without identifiers, one could also create them beforehand with and save it as a variable so that it can be reused whenever needed. E.g.

```

...

var_predicates = ["label", "type"]

...

{ 
"title" : "query 1"
"query" : [ 
    "SELECT * WHERE { 
        ?s <http://www.w3.org/2000/01/rdf-schema#", var_predicates, "> ?o 
    }" 
]
},

{ 
"title" : "query 2"
"query" : [ 
    "SELECT COUNT(?p) WHERE { 
        ?s ?p ?o . 
        VALUES ?p { \"http://www.w3.org/2000/01/rdf-schema#", var_predicates, "\" }
    }"
]
},

...
```
