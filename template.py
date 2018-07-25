
    
    
# -------------------- OPTIONAL SETTINGS -------------------- 

# title
# defines the title of the whole set of queries
# OPTIONAL, if not set, timestamp will be used
title = "TEST QUERIES"


# description
# defines the textual and human-intended description of the purpose of these queries
# OPTIONAL, if not set, nothing will be used or displayed
description = "This set of queries is used as a template for showcasing a valid query collection file."


# output_destination
# defines where to save the results, input can be: 
# * a local path to a folder 
# * a URL for a google sheets document  
# * a URL for a google folder
# NOTE: On windows, folders in a path use backslashes, in such a case it is mandatory to attach a 'r' in front of the quotes, e.g. r"C:\Users\sresch\.."
# In the other cases the 'r' is simply ignored; thus best would be to always leave it there.
# OPTIONAL, if not set, folder of executed script will be used
output_destination = r"."


# output_format
# defines the format in which the result data shall be saved (currently available: csv, tsv, xml, json, xlsx)
# OPTIONAL, if not set, csv will be used
output_format = "csv"


# summary_sample_limit
# defines how many rows shall be displayed in the summary
# OPTIONAL, if not set, 5 will be used
summary_sample_limit = 3


# cooldown_between_queries
# defines how many seconds should be waited between execution of individual queries in order to prevent exhaustion of Google API due to too many writes per time-interval
# OPTIONAL, if not set, 0 will be used
cooldown_between_queries = 0


# write_empty_results
# Should empty results be written to summary files? Possible values are python boolean values: True, False
# OPTIONAL, if not set, False will be used
write_empty_results = False


# -------------------- MANDATORY SETTINGS -------------------- 

# endpoint
# defines the SPARQL endpoint against which all the queries are run
# MANDATORY
endpoint = "http://dbpedia.org/sparql"

# queries
# defines the set of queries to be run. 
# MANDATAORY
queries = [


    {
        # title
        # OPTIONAL, if not set, timestamp will be used
        "title" : "Optional title of first query" ,

        # description
        # OPTIONAL, if not set, nothing will be used or displayed
        "description" : "Optional description of first query, used to describe the purpose of the query, which in this case is of mere demonstration." ,

        # query
        # the sparql query itself
        # NOTE: best practise is to attach a 'r' before the string so that python would not interpret some characters as metacharacters, e.g. "\n"
        # MANDATORY
        "query" : r"""
            SELECT * WHERE {
                ?s ?p ?o
            }
        """
    },   
    {    
        "title" : "Second query" , 
        "description" : "This query returns all triples which have a rdf:type associated" , 
        "query" : r"""
            SELECT * WHERE {
                ?s <http://www.w3.org/2000/01/rdf-schema#label> ?o
            }
        """
    },
    {    
        "query" : r"""
            SELECT COUNT (?s) AS ?count_of_subjects_with_type WHERE {
                ?s <http://www.w3.org/1999/02/22-rdf-syntax-ns#type> ?o
            }
        """
    },
]

# Each query is itself encoded as a python dictionary, and together these dictionaries are collected in a python list. Beginner's note on such syntax as follows:
# * the set of queries is enclosed by '[' and ']'
# * individual queries are enclosed by '{' and '},'
# * All elements of a query (title, description, query) need to be defined using quotes as well as their contents, and both need to be separated by ':'
# * All elements of a query (title, description, query) need to be separated from each other using quotes ','
# * The content of a query needs to be defined using triple quotes, e.g. """ SELECT * WHERE .... """
# * Any indentation (tabs or spaces) do not influence the queries-syntax, they are merely syntactic sugar.


