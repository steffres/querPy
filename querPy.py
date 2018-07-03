#!/usr/bin/python3

from __future__ import print_function
import argparse
import imp
import csv
import json
import logging
import sys
import time
import os
import regex
from httplib2 import Http
import xlsxwriter
from pathlib import Path
from SPARQLWrapper import *
from googleapiclient import discovery
from oauth2client import client, tools, file
from oauth2client.client import GoogleCredentials


def main():

    # argument parser
    parser = argparse.ArgumentParser()
    parser.add_argument("-r", help="runs all queries in the specified file. (To create a template for such a file, use '-t'.)")
    parser.add_argument("-s", help="reads in a provided client_secret json file. If no client_secret.json is provided as argument, querPy will search the current folder for one. (A client_secret can be obtained by logging into the Google Developer Console where a projects needs to be registered.)")
    parser.add_argument("-c", help="reads in a provided credentials json file. If no credentials.json is provided as argument, querPy will search the current folder for one. If there does not exist a credentials file yet, you can create one by providing a client_secret, after which you should be directed to a google-login, the resulting credentials file will be saved in the current folder.")
    parser.add_argument("-t", action='store_true', help="creates a template file for showcasing the queries-layout")

    if len(sys.argv) == 1:
        print("Invalid arguments!")
        parser.print_help()
        sys.exit()

    args = parser.parse_args()

    # user wants to run a queries file and does not want to create a template file
    if args.r and not args.t:

        logging.basicConfig(filename="querPy_log.log", filemode="w", level=logging.INFO)

        # read user configuration file
        conf = imp.load_source('conf', args.r)

        # extract and validate data from the configuration
        data = read_input(conf, args.r)

        ## google authentication cases

        # user provides a credentials.json file
        if args.c:
            data['credentials_path'] = args.c
            data['client_secret_path'] = False

        # user provides a client_secret.json file
        elif args.s:
            data['client_secret_path'] = args.s
            data['credentials_path'] = False

        # user did not provide any file. Search local folder for files and load them
        else:
            files_list = os.listdir('./')

            if "credentials.json" in files_list:
                data['credentials_path'] = "credentials.json"
                data['client_secret_path'] = False

            elif "client_secret.json" in files_list:
                data['client_secret_path'] = "client_secret.json"
                data['credentials_path'] = False

            else:
                data['credentials_path'] = False
                data['client_secret_path'] = False

        # create OutputWriter object from the given output-configuration
        output_writer = OutputWriter(data)

        # run all the queries and let them write using the OutputWriter object
        execute_queries(data, output_writer)

        # Close xlsl writer
        output_writer.close()


    # user wants to create a template file and does not run a queries-file
    elif args.t and not args.r:

        create_template()

    # invalid arguments, print help
    else:
        print("Invalid arguments!")
        parser.print_help()
        sys.exit()




def read_input(conf, conf_filename):
    """Reads input from config file and convert into usable data structure available throughout the entire program lifecycle"""

    def main(conf):

        data = {}
        data['timestamp_start'] = time.strftime('%y%m%d_%H%M%S')

        message = \
            "\n################################\n" + \
            "READING CONFIG FILE: " + conf_filename + \
            "\n################################\n" + \
            "\ntimestamp: " + str(data['timestamp_start'])
        logging.info(message)
        print(message)

        # title
        try:
            data['title'] = conf.title
        except AttributeError:
            message = "Did not find title in config file, using timestamp instead"
            logging.info(message)
            print(message)
            data['title'] = data['timestamp_start']
        logging.info("title: " + data['title'])


        # description
        try:
            data['description'] = conf.description
        except AttributeError:
            message = "Did not find description in config file, ignoring instead"
            logging.info(message)
            print(message)
            data['description'] = ""
        logging.info("description: " + data['description'])


        # output_destination
        try:
            if conf.output_destination == "":
                data['output_destination'] = "."
            else:
                data['output_destination'] = conf.output_destination
        except AttributeError:
            message = "Did not find output_destination in config file, using local instead"
            logging.info(message)
            print(message)
            data['output_destination'] = "."
        message = "output_destination: " + data['output_destination']
        logging.info(message)
        print(message)


        # output_format
        try:
            if conf.output_format.upper() == "CSV":
                data['output_format'] = CSV
            elif conf.output_format.upper() == "TSV":
                data['output_format'] = TSV
            elif conf.output_format.upper() == "XML":
                data['output_format'] = XML
            elif conf.output_format.upper() == "JSON":
                data['output_format'] = JSON
            elif conf.output_format.upper() == "XLSX":
                data['output_format'] = "XLSX"
            elif conf.output_format.upper() == "" or conf.output_format is None:
                message = "Did not find output_format, using CSV instead"
                logging.info(message)
                print(message)
                data['output_format'] = CSV
            else:
                message = "INVALID INPUT! output_format not recognized: '" + conf.output_format + "'. Available formats are: CSV, TSV, XML, JSON, XLSX"
                logging.error(message)
                sys.exit(message)
        except AttributeError:
            message = "Did not find output_format in config file, using csv instead"
            logging.info(message)
            print(message)
            data['output_format'] = CSV
        logging.info("output_format: " + data['output_format'])


        # summary_sample_limit
        try:
            data['summary_sample_limit'] = conf.summary_sample_limit
        except AttributeError:
            message = "Did not find summary_sample_limit in config file, a limit of 5 will be used instead"
            logging.info(message)
            print(message)
            data['summary_sample_limit'] = 5
        if data['summary_sample_limit'] > 101:
            data['summary_sample_limit'] = 101
        logging.info("summary_sample_limit: " + str(data['summary_sample_limit']))


        # cooldown_between_queries
        try:
            data['cooldown_between_queries'] = conf.cooldown_between_queries
        except AttributeError:
            message = "Did not find cooldown_between_queries in config file, assuming zero instead"
            logging.info(message)
            print(message)
            data['cooldown_between_queries'] = 0
        logging.info("cooldown_between_queries: " + str(data['cooldown_between_queries']))


        # endpoint
        try:
            data['endpoint'] = conf.endpoint
        except AttributeError:
            message = "INVALID INPUT! Did not find endpoint in config file! Compare to template file generated by running 'querPy.py -t'"
            logging.error(message)
            sys.exit(message)
        message = "endpoint: " + data['endpoint']
        logging.info(message)
        print(message)


        # queries
        len_queries = len(conf.queries)
        if len_queries == 0:
            raise AttributeError()

        data['queries'] = []
        for i in range(0, len_queries):
            # get title
            try:
                query_title = conf.queries[i]["title"]
                if query_title.isspace() or query_title == "":
                    query_title = str(i + 1)
                else:
                    query_title = str(i + 1) + ". " + query_title
            except KeyError:
                query_title = str(i + 1)

            # get description
            try:
                query_description = conf.queries[i]["description"]
            except KeyError:
                query_description = ""

            # get query
            try:
                query_text = conf.queries[i]["query"]
            except KeyError:
                message = "INVALID INPUT! Did not find queries in config file! Compare to template file generated by running 'querPy.py -t'"
                logging.error(message)
                sys.exit(message)

            logging.info("got query_title: " + query_title)
            logging.info("scrubbing.")

            query_text = scrub_query(query_text)

            data['queries'].append({
                "query_id" : "Q" + str(i+1),
                "query_title": query_title,
                "query_description": query_description,
                "query_text": query_text})

            logging.info("query_text (scrubbed): \n" + data['queries'][i]['query_text'])

        return data


    def scrub_query(query_text):
        """Scrubs the queries clean from unneccessary whitespaces and indentations"""

        if not query_text.isspace() and not query_text == "":

            # replace tabs with spaces for universal formatting
            query_lines = query_text.replace("\t", "    ").splitlines()

            # get smallest number of whitespaces in front of all lines
            spaces_in_front = float("inf")
            for j in range(0, len(query_lines)):

                if not query_lines[j].isspace() and not query_lines[j] == "":

                    spaces_in_front_tmp = len(query_lines[j]) - len(query_lines[j].lstrip(" "))
                    if spaces_in_front_tmp < spaces_in_front:
                        spaces_in_front = spaces_in_front_tmp

            # remove redundant spaces in front
            if spaces_in_front > 0:
                query_text = ""
                for line in query_lines:
                    query_text += line[spaces_in_front:] + "\n"

            # remove "" and heading and unneccessary newlines
            query_lines = query_text.splitlines()
            query_text = ""
            for line in query_lines:
                if not line.isspace() and not line == "":
                    query_text += line + "\n"

        return query_text

    return main(conf)



def execute_queries(data, output_writer):
    """Executes all the queries and calls the writer-method from the OutputWriter object to write it to the specified destinations"""

    def main(data):


        message = \
            "\n################################\n" + \
            "STARTING EXECUTION OF QUERIES" + \
            "\n################################\n"
        logging.info(message)
        print(message)

        sparql_wrapper = SPARQLWrapper(data['endpoint'])

        try:
            message = "Getting count of all triples in whole triplestore"
            logging.info(message)
            print(message)

            # get count of all triples in endpoint for statistical purposes
            sparql_wrapper.setQuery("SELECT COUNT(*) WHERE {[][][]}")
            sparql_wrapper.setReturnFormat(JSON)
            count_triples_in_endpoint = sparql_wrapper.query().convert()
            data['count_triples_in_endpoint'] = count_triples_in_endpoint["results"]["bindings"][0]["callret-0"]["value"]
            logging.info("count_triples_in_endpoint: " + data['count_triples_in_endpoint'] + "\n")
            data['header_error_message'] = None

        except Exception as ex:

            message = "EXCEPTION OCCURED! " + str(ex)
            print(message)
            logging.error(message)
            data['header_error_message'] = message

        # Write header
        output_writer.write_header_summary(data)

        # execute queries
        for i in range(0, len(data['queries'])):

            query = data['queries'][i]

            message = \
                "\n################################\n" + \
                "EXECUTE: " + query['query_title'] + "\n" + query['query_text']
            logging.info(message)
            print(message)

            startTime = time.time()
            results = None
            query['results_lines_count'] = -1

            try:
                # execute query

                sparql_wrapper.setQuery(query['query_text'])

                if data['output_format'] == "XLSX":
                    sparql_wrapper.setReturnFormat(CSV)
                else:
                    sparql_wrapper.setReturnFormat(data['output_format'])

                logging.info("query_text: \n" + query['query_text'])

                startTime = time.time()
                results = sparql_wrapper.query().convert()
                query['results_execution_duration'] = time.time() - startTime


                # get count of total results for query
                # For this, search for the first select statement, and replace it with
                # select count(*) where ... and add a '}' to the end, to make the original select
                # a sub-query
                #
                # This requires a non-standard regex module to use variable length negativ look behind
                # which are needed to detect only select statements, where there are no '#'
                # before, which would make it a comment and thus not necesseray to replace

                pattern = regex.compile(r"(?<!#.*)select", regex.IGNORECASE | regex.MULTILINE)
                query_for_count = pattern.sub(
                    "SELECT COUNT(*) WHERE { \nSELECT",
                    query['query_text'],
                    1)
                query_for_count += "\n}"

                query['query_for_count'] = query_for_count
                logging.info("#### query_title: " + query['query_title'] + "\n")
                logging.info("query_for_count: \n" + query['query_for_count'] + "\n")

                sparql_wrapper.setQuery(query_for_count)
                sparql_wrapper.setReturnFormat(JSON)
                results_lines_count = sparql_wrapper.query().convert()
                results_lines_count = results_lines_count["results"]["bindings"][0]["callret-0"]["value"]
                query['results_lines_count'] = results_lines_count
                logging.info("results_lines_count: " + query['results_lines_count'] + "\n")


            except Exception as ex:
                message = "EXCEPTION OCCURED WHEN EXECUTING QUERY: " + str(ex) + "\n Continue with execution of next query."
                print(message)
                logging.error(message)
                query['error_message'] = str(ex)
                query['results_execution_duration'] = time.time() - startTime


            query['results'] = results
            logging.info("results_execution_duration: " + str(query['results_execution_duration']) + "\n")


            # harmonize results for other uses later
            logging.info("harmonizing results")
            startTime = time.time()
            query['results_harmonized'] = get_harmonized_result(query['results'], data['output_format'])
            logging.info("Done with harmonizing results, duration: " + str(time.time() - startTime))

            message = "EXECUTION FINISHED\nElapsed time: " + str(query['results_execution_duration'])
            logging.info(message)
            print(message)

            output_writer.write_query_summary(query)
            output_writer.write_query_result(query)

            if (data['cooldown_between_queries']) > 0 and i+1 < len(data['queries']):
                print("\nSleep for " + str(data['cooldown_between_queries']) + " seconds.")
                time.sleep(data['cooldown_between_queries'])

            data['queries'][i] = query

        return data


    # convert results from any format into a two-dimensional array
    def get_harmonized_result(result, format):
        """Transforms the result data from its varying data formats into a two-dimensional list, used for writing summaries or into xlsx / google sheets files"""

        def get_harmonized_rows_from_keyed_rows(result_sample_keyed):
            """Some output formats require an intermediate step where the individual result rows are initially indexed by keys and their layout might change from row to row. Thus this methods iterate over each key-value row and transforms it into a regular two-dimensional list, where every column is identifiable by the same column-key"""

            # transform the result_sample_keyed into a regular two-dimensional list used for later inserting into xlsx or google sheets
            harmonized_rows = []
            harmonized_rows.append(result_sample_keyed[0])

            for y in range(1, len(result_sample_keyed)):

                sample_row = []
                for x in range(0, len(harmonized_rows[0])):
                    key = harmonized_rows[0][x]
                    sample_row.append(result_sample_keyed[y][key])

                harmonized_rows.append(sample_row)

            return harmonized_rows



        harmonized_result = []

        if result is None:
            return None
        else:

            # CSV, TSV, XLSX (since XLSX means CSV is internally used for querying the endpoint)
            if format == CSV or format == TSV or format == "XLSX":

                result = result.decode('utf-8').splitlines()
                harmonized_result = []
                valid_row_length = float("inf")

                if format == TSV:
                    reader = csv.reader(result, delimiter="\t")
                else:
                    reader = csv.reader(result)

                for row in reader:

                    row_harmonized = []

                    for column in row:

                        # check if value could be integer, if so change type
                        try:
                            column = int(column)
                        except ValueError:
                            column = column

                        row_harmonized.append(column)

                    harmonized_result.append(row_harmonized)

                    # check validity of results
                    current_row_length = len(row)
                    if valid_row_length != float("inf") and valid_row_length != current_row_length:
                        message = "INVALID ROW LENGTH! " + str(row) + " has length " + str(current_row_length) + ", while valid length is " + str(valid_row_length)
                        logging.error(message)
                        sys.exit(message)
                    valid_row_length = current_row_length


            # JSON
            elif format == JSON:

                # construct list of dictionaries (to preserve the key-value pairing of individual row-results)

                result_keyed = []

                # get keys and save in first row of result_keyed
                keys = []
                for key in result['results']['bindings'][0]:
                    keys.append(key)
                result_keyed.append(keys)

                # go through the json - rows and extract key-value pairs from each, insert them into result_keyed
                valid_row_length = len(keys)
                for y in range(0, len(result['results']['bindings'])):
                    dict_tmp = {}

                    row = result['results']['bindings'][y]

                    for key in row:
                        column = row[key]['value']

                        # check if value could be integer, if so change type
                        try:
                            column = int(column)
                        except ValueError:
                            column = column

                        dict_tmp[key] = column

                    # check validity of results
                    if len(row) != valid_row_length:
                        message = "INVALID ROW LENGTH! " + str(row) + " has length " + str(len(row)) + ", while valid length is " + str(valid_row_length)
                        logging.error(message)
                        sys.exit(message)

                    result_keyed.append(dict_tmp)

                harmonized_result = get_harmonized_rows_from_keyed_rows(result_keyed)


            # XML
            elif format == XML:

                # construct list of dictionaries (to preserve the key-value pairing of individual row-results)

                result_keyed = []

                # get keys and save in first row of result_keyed
                vars = result.getElementsByTagName("head")[0].getElementsByTagName("variable")
                keys = []
                for var in vars:
                    keys.append(var.getAttribute('name'))
                result_keyed.append(keys)

                # get results rows
                results = result.getElementsByTagName("result")

                # go through the xml results and extract key-value pairs from each, insert them into result_keyed
                valid_row_length = len(keys)
                for y in range(0, len(results)):

                    result = results[y]

                    dict_tmp = {}
                    for binding in result.getElementsByTagName("binding"):
                        column = binding.childNodes[0].childNodes[0].nodeValue

                        # check if value could be integer, if so change type
                        try:
                            column = int(column)
                        except ValueError:
                            column = column

                        dict_tmp[binding.getAttribute('name')] = column

                    # check validity of results
                    if len(dict_tmp) != valid_row_length:
                        message = "INVALID ROW LENGTH! " + str(dict_tmp) + " has length " + str(len(dict_tmp)) + ", while valid length is " + str(valid_row_length)
                        logging.error(message)
                        sys.exit(message)

                    result_keyed.append(dict_tmp)

                harmonized_result = get_harmonized_rows_from_keyed_rows(result_keyed)

            return harmonized_result

    return main(data)




class OutputWriter:
    """the OutputWriter Class encapsulates all technical details which vary due to the specified output destinations"""

    # general variables
    output_destination_type = None
    summary_sample_limit = None
    line_number = None

    # local folder and xlsx variables
    folder = None
    file_xlsx = None
    xlsx_workbook = None
    xlsx_worksheet_summary = None
    output_format = None
    bold_format = None
    title_2_format = None
    query_text_format = None

    # google folder and sheets variables
    google_service_sheets = None
    google_service_drive = None
    google_sheets_id = None
    google_sheets_summary_sheet_id = None

    def __init__(self, data):

        def main():

            # output_destination_type, interpret from string

            if "google.com/drive/folders" in data['output_destination']:
                self.output_destination_type = "google_folder"
                logging.info("deduced output_destination_type: " + self.output_destination_type)
                init_google_folder()

            elif "google.com/spreadsheets" in data['output_destination']:
                self.output_destination_type = "google_sheets"
                logging.info("deduced output_destination_type: " + self.output_destination_type)
                init_google_sheets()

            elif data['output_format'] == "XLSX" :
                self.output_destination_type = "local_xlsx"
                logging.info("deduced output_destination_type: " + self.output_destination_type)
                init_local_xlsx()

            else:
                self.output_destination_type = "local_folder"
                logging.info("deduced output_destination_type: " + self.output_destination_type)
                init_local_folder()


        def init_local_xlsx():
            """Creates a xlsx file in the respective folder"""

            # if locally saved, then "/" needs to be replaced with "-", since otherwise "/" would be interpreted as subfolder
            data['title'] = data['title'].replace("/", "-")
            for query in data['queries']:
                query['query_title'] = query['query_title'].replace("/", "-")

            # get or create folder for xlsx file
            self.folder = Path(str(data['output_destination']))
            self.folder.mkdir(parents=True, exist_ok=True)

            # create xlsx file
            self.file_xlsx = Path(
                self.folder / str( data['timestamp_start'] + " - " + data['title'] + ".xlsx" ) )
            self.xlsx_workbook = xlsxwriter.Workbook(self.file_xlsx.open('wb'))
            self.xlsx_worksheet_summary = self.xlsx_workbook.add_worksheet("0. Summary")

            message = "Created local file: " + str(self.folder)
            logging.info(message)
            print(message)


        def init_local_folder():
            """Creates a folder (for the raw ouput) and a xlsx file (for the summary) in the respective folder"""

            # if locally saved, then "/" needs to be replaced with "-", since otherwise "/" would be interpreted as subfolder
            data['title'] = data['title'].replace("/", "-")
            for query in data['queries']:
                query['query_title'] = query['query_title'].replace("/", "-")

            # create folder for queries and summary
            self.folder = Path(str(
                data['output_destination'] + "/" +
                data['timestamp_start'] + " - " +
                data['title']))
            self.folder.mkdir(parents=True, exist_ok=False)
            self.output_format = data['output_format']

            # Create xlsx file for summary
            self.file_xlsx = Path(self.folder / "0. Summary.xlsx")
            self.xlsx_workbook = xlsxwriter.Workbook(self.file_xlsx.open('wb'))
            self.xlsx_worksheet_summary = self.xlsx_workbook.add_worksheet("0. Summary")

            message = "Created local folder: " + str(self.folder)
            logging.info(message)
            print(message)


        def init_google_services():
            """Instantiates all necessary services for writing results to a specified google folder / sheets-file"""

            SCOPES = "https://www.googleapis.com/auth/drive"
            creds_hardcoded = None


            # Hardwired credentials
            #
            # !!! CAUTION !!!
            #
            # POSSIBILITY OF GRANTING FULL ACCESS TO YOUR PRIVATE GOOGLE DRIVE
            #
            # !!! CAUTION !!!
            #
            # For ease of usage on your local machine, you can hardwire your credentials here
            # BUT ONLY DO THIS IF YOU NEVER SHARE THIS MODIFIED SCRIPT
            #
            # NEVER INSERT YOUR CREDENTIALS IF YOU WILL SHARE THIS SCRIPT!!
            #
            # creds_hardcoded = json.loads("""
            #   UNCOMMENT AND INSERT CONTENT OF CREDENTIALS.JSON FILE HERE
            # """)

            # use credentials file if available
            if data['credentials_path']:
                creds = client.GoogleCredentials.from_json(open(data['credentials_path']).read())

            # if no credentials file is available, then create one using client_secret
            elif data['client_secret_path']:
                store = file.Storage('credentials.json')
                flow = client.flow_from_clientsecrets(data['client_secret_path'], SCOPES)
                creds = tools.run_flow(flow, store, tools.argparser.parse_args(args=[]))
                # note: adding 'tools.argparser.parse_args(args=[])' here is important, otherwise
                # oauth2client.tools would parse the main command line arguments


            elif creds_hardcoded:

                creds = GoogleCredentials(
                    creds_hardcoded['access_token'],
                    creds_hardcoded['client_id'],
                    creds_hardcoded['client_secret'],
                    creds_hardcoded['refresh_token'],
                    creds_hardcoded['token_expiry'],
                    creds_hardcoded['token_uri'],
                    creds_hardcoded['user_agent'],
                    creds_hardcoded['revoke_uri']
                )

            # if neither is available, abort
            else:
                message = "ERROR: No client_secret.json or credentials.json provided nor found in local folder!."
                logging.error(message)
                sys.exit(message)

            # create services to be used by write functions
            if not creds.invalid:
                self.google_service_drive = discovery.build('drive', 'v3', http=creds.authorize(Http()))
                self.google_service_sheets = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
            else:
                message = "ERROR: Invalid credentials!"
                logging.error(message)
                sys.exit(message)


        def init_google_sheets():
            """Formats the give google sheets file, deletes old content and creates a summary-sheet"""

            init_google_services()

            # get id of google sheets file by extracting it from the url
            self.google_sheets_id = data['output_destination']\
                .split("docs.google.com/spreadsheets/d/",1)[1]\
                .split("/",1)[0]
            logging.info("ID of google sheets : " + str(self.google_sheets_id))

            # get list of existing sheets in sheets file
            google_sheets_metadata = self.google_service_sheets.spreadsheets().get(
                spreadsheetId=self.google_sheets_id).execute()
            all_sheet = google_sheets_metadata['sheets']

            ## create new sheet reserved for summary

            # max_row_count = header + number of queries * (maximum sample lines + query-header)
            max_row_count = 6 + len(data['queries']) * (data['summary_sample_limit'] + 11 )

            body_create_summary_page = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "gridProperties": {
                                    "rowCount": max_row_count,
                                    "columnCount": 26
                                }
                            }
                        }
                    }
                ]
            }
            result = self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id, body=body_create_summary_page).execute()
            self.google_sheets_summary_sheet_id = result['replies'][0]['addSheet']['properties']['sheetId']

            # delete all sheets except summary
            body_sheet_to_delete = { 'requests' : [] }
            for sheet in all_sheet:
                tmp = {
                    "deleteSheet": {
                        "sheetId": sheet['properties']['sheetId']
                    }
                }
                body_sheet_to_delete['requests'].append(tmp)

            self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id, body=body_sheet_to_delete).execute()

            # rename summary sheet to '0. Summary'
            body_to_rename = {
                "requests" : [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": self.google_sheets_summary_sheet_id,
                                "title": "0. Summary",
                            },
                            "fields": "title",
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": self.google_sheets_summary_sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex":26
                            },
                            "properties": {
                                "pixelSize": 350
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id, body=body_to_rename).execute()


        # google folder
        def init_google_folder():
            """Creates a new google sheets file inside the specified google folder"""

            init_google_services()

            # get id of google folder by extracting it from the url
            self.google_folder_id = data['output_destination']\
                .split("drive.google.com/drive/folders/",1)[1]\
                .split("?",1)[0]
            logging.info("ID of google folder : " + str(self.google_folder_id))

            # Create google sheets file in folder
            body_spreadsheet = {
                'name': data['timestamp_start'] + " - " + data['title'],
                'mimeType': 'application/vnd.google-apps.spreadsheet',
                'parents': [self.google_folder_id]
            }
            sheets =  self.google_service_drive.files().create(body=body_spreadsheet).execute()
            self.google_sheets_id = sheets['id']
            self.google_sheets_summary_sheet_id = 0

            # Sets name of first sheet to summary, sets up column width
            body_to_rename = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": self.google_sheets_summary_sheet_id,
                                "title": "0. Summary",
                            },
                            "fields": "title",
                        }
                    },
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": self.google_sheets_summary_sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex":26
                            },
                            "properties": {
                                "pixelSize": 300
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id, body=body_to_rename).execute()

            message = "Created google sheets at: \n" + "docs.google.com/spreadsheets/d/" + self.google_sheets_id
            logging.info(message)
            print(message)

        main()


    def write_header_summary(self, data):
        """Writes the initial header to the summary sheet"""

        def main(data):

            self.summary_sample_limit = data['summary_sample_limit']

            if self.output_destination_type == 'local_folder' or self.output_destination_type == 'local_xlsx':
                write_header_summary_xlsx_file(data)

            elif self.output_destination_type == 'google_folder' or self.output_destination_type == 'google_sheets':
                write_header_summary_google_sheet(data)


        def write_header_summary_xlsx_file(data):
            """Writes header to xlsx file"""

            message = "Writing header to summary in local xslx"
            logging.info(message)
            print(message)

            # setup and formats
            self.xlsx_worksheet_summary.set_column('A:Z', 70)
            self.title_format = self.xlsx_workbook.add_format({'bold': True})
            self.title_format.set_font_size(16)
            self.title_2_format = self.xlsx_workbook.add_format({'bold': True})
            self.title_2_format.set_font_size(12)
            self.query_text_format = self.xlsx_workbook.add_format({'text_wrap': True})
            self.bold_format = self.xlsx_workbook.add_format({'bold': True})

            # Write header to xlsx
            self.xlsx_worksheet_summary.set_row(0, 20)
            self.line_number = 0
            self.xlsx_worksheet_summary.write(self.line_number, 0, data['title'], self.title_format)
            if data['description'] != "":
                self.xlsx_worksheet_summary.write(self.line_number + 1, 0, data['description'])
                self.line_number += 1
            self.line_number += 2
            self.xlsx_worksheet_summary.write(self.line_number, 0, "Execution timestamp of script: " + data['timestamp_start'])
            self.line_number += 1
            if data['header_error_message'] is None:
                self.xlsx_worksheet_summary.write(self.line_number, 0, "Endpoint: " + data['endpoint'])
                self.line_number += 1
                self.xlsx_worksheet_summary.write(self.line_number, 0, "Total count of triples in endpoint: " + data[
                    'count_triples_in_endpoint'])
            else:
                self.xlsx_worksheet_summary.write(self.line_number, 0, data['header_error_message'])
            self.line_number += 4



        def write_header_summary_google_sheet(data):
            """Writes header to google sheets file"""

            message = "Writing header to summary in google sheets"
            logging.info(message)
            print(message)

            # create header info
            self.line_number = 0
            header = []
            header.append([data['title']])
            if data['description'] != "":
                header.append([data['description']])
            header.append([])
            header.append(
                ["Execution timestamp of script: " +
                 data['timestamp_start']])
            if data['header_error_message'] is None:
                header.append(["endpoint: " + data['endpoint']])
                header.append(
                    ["Total count of triples in endpoint: " +
                     data['count_triples_in_endpoint']])
            else:
                header.append([data['header_error_message']])



            # get range for header
            range = self.get_range_from_matrix(self.line_number, 0, header)
            range = "0. Summary!" + range
            self.line_number += len(header) + 3

            # write header to sheet
            self.google_service_sheets.spreadsheets().values().update(
                    spreadsheetId=self.google_sheets_id, range=range,
                    valueInputOption="RAW", body= { 'values': header } ).execute()

        main(data)


    def write_query_result(self, query):
        """Writes results of query to the respective output destination"""

        def main(query):

            if not query['results'] is None:
                message = "Writing results to output_destination"
                logging.info(message)
                print(message)

                if self.output_destination_type == 'local_xlsx':
                    write_query_result_to_xlsx_file(query)

                elif self.output_destination_type == 'local_folder':
                    write_query_result_to_local_folder(query)

                elif self.output_destination_type == 'google_sheets' or self.output_destination_type == 'google_folder':
                    write_query_result_to_google_sheets(query)


        def write_query_result_to_xlsx_file(query):
            """Writes results as harmonized two-dimensional list into a separate sheet in the xlsl file"""

            # create new worksheet and write into it
            sanitized_query_title = query['query_title']
            if len(sanitized_query_title) > 30:
                sanitized_query_title = sanitized_query_title[:29]

            worksheet = self.xlsx_workbook.add_worksheet( sanitized_query_title )
            for y in range(0, len(query['results_harmonized'])):
                for x in range(0, len(query['results_harmonized'][y])):
                    column = query['results_harmonized'][y][x]
                    if len(str(column)) > 255:
                        column = str(column)[:255]
                    worksheet.write(y, x, column)


        def write_query_result_to_local_folder(query):
            """Writes raw output using the respective data format into the specified local folder"""

            # create file for query result
            file_name = query['query_title'] + "." + self.output_format
            local_file = Path(self.folder / file_name)

            ## differentiate between different result-types which require different write-methods

            # csv and tsv files need to be written as bytes
            if self.output_format == CSV or self.output_format == TSV:
                with local_file.open('wb') as fw:
                    fw.write(query['results'])

            # xml document is passed a writer object
            elif self.output_format == XML:
                with local_file.open('w') as fw:
                    query['results'].writexml(fw)

            # json needs json.dump() method
            elif self.output_format == JSON:
                with local_file.open('w') as fw:
                    json.dump(query['results'], fw)


        def write_query_result_to_google_sheets(query):
            """Writes results as harmonized two-dimensional list into a separate sheet in the google sheets file"""

            sanitized_query_title = query['query_title']
            if len(sanitized_query_title) > 100:
                sanitized_query_title = sanitized_query_title[:99]

            # create sheet
            body_new_sheet = {
                'requests' : [
                    {
                        'addSheet': {
                            'properties': {
                                'title': sanitized_query_title,
                                'gridProperties': {
                                    'rowCount': len(query['results_harmonized']),
                                    'columnCount': len(query['results_harmonized'][0])
                                }
                            }
                        }
                    }
                ]
            }
            result = self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id,
                body=body_new_sheet
            ).execute()
            google_sheet_id = result['replies'][0]['addSheet']['properties']['sheetId']
            body_change_columns = {
                'requests': [
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": google_sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex": 26
                            },
                            "properties": {
                                "pixelSize": 300
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            self.google_service_sheets.spreadsheets().batchUpdate(
                spreadsheetId=self.google_sheets_id,
                body=body_change_columns
            ).execute()

            # get range of harmonized results
            google_sheet_range = \
                sanitized_query_title + "!" + \
                self.get_range_from_matrix(0, 0, query['results_harmonized'])

            # write into sheet
            self.google_service_sheets.spreadsheets().values().update(
                spreadsheetId=self.google_sheets_id,
                range=google_sheet_range,
                valueInputOption="RAW",
                body={ 'values': query['results_harmonized']}
            ).execute()

        main(query)


    def write_query_summary(self, query):
        """Writes the gist of the results of an executed query to a summary sheet"""

        def main(query):

            message = "Writing to summary"
            logging.info(message)
            print(message)

            if self.output_destination_type == 'local_xlsx' or self.output_destination_type == 'local_folder' :
                write_query_summary_xlsx_file(query)

            elif self.output_destination_type == 'google_sheets' or self.output_destination_type == 'google_folder':
                write_query_summary_google_sheets(query)


        def write_query_summary_xlsx_file(query):
            """Writes the gist of the results of an executed query to the summary sheet in the xlsx file"""

            # query_title
            self.xlsx_worksheet_summary.write(self.line_number, 0, query['query_title'], self.title_2_format)
            self.line_number += 1

            # query description
            if not ( query['query_description'].isspace() or query['query_description'] == "" ) :
                self.xlsx_worksheet_summary.write(self.line_number, 0, query['query_description'])
                self.line_number += 1

            # query_text
            size_of_query_text_row = 15 * (query['query_text'].count("\n") + 2)
            self.xlsx_worksheet_summary.set_row(self.line_number, size_of_query_text_row)
            self.xlsx_worksheet_summary.write(self.line_number, 0, query['query_text'], self.query_text_format)
            self.line_number += 1

            # results_execution_duration
            self.xlsx_worksheet_summary.write(self.line_number, 0, "Duration of execution in seconds: " + str(query['results_execution_duration']))
            self.line_number += 1

            if query['results'] is None:
                self.xlsx_worksheet_summary.write(self.line_number, 0, "NO RESULTS DUE TO ERROR: " + query['error_message'])
                self.line_number += 1

            else:
                # results_lines_count
                self.xlsx_worksheet_summary.write(self.line_number, 0, "Total count of lines in results: " + str(query['results_lines_count']))
                self.line_number += 2

                # results
                limit = self.summary_sample_limit
                if limit != 0:

                    self.xlsx_worksheet_summary.write(self.line_number, 0, "Sample results: ", self.bold_format)
                    self.line_number += 1
                    harmonized_rows = query['results_harmonized']

                    limit += 1
                    if len(harmonized_rows) < limit:
                        limit = len(harmonized_rows)

                    y = 0
                    for y in range(0, limit):
                        for x in range(0, len(harmonized_rows[y])):

                            column = harmonized_rows[y][x]

                            if len(str(column)) > 255:
                                column = str(column)[:255]
                            self.xlsx_worksheet_summary.write(y + self.line_number, x, column)

                    self.line_number += 1

                self.line_number += limit

            self.line_number += 2


        def write_query_summary_google_sheets(query):
            """Writes the gist of the results of an executed query to the summary sheet in the google sheets file"""

            # creating header
            query_stats = []
            query_stats.append([query['query_title']])
            if not ( query['query_description'].isspace() or query['query_description'] == "" ) :
                query_stats.append([query['query_description']])
            query_stats.append([query['query_text']])
            query_stats.append(
                ["Duration of execution in seconds: " +
                 str(query['results_execution_duration'])])


            if query['results'] is None:
                query_stats.append(["NO RESULTS DUE TO ERROR: " + query['error_message']])

            else:
                query_stats.append(
                    ["Total count of lines in results: " +
                     str(query['results_lines_count'])])

                # get sample results
                limit = self.summary_sample_limit
                if limit != 0:

                    query_stats.append([])
                    query_stats.append(["Sample results: "])
                    harmonized_rows = query['results_harmonized']

                    # set limit as defined, readjust if results should be less than it or if it exceeds gsheets-capacities
                    limit += 1
                    if len(harmonized_rows) < limit:
                        limit = len(harmonized_rows)

                    for y in range(0, limit):
                        query_stats.append(harmonized_rows[y])

            # write header and sample results to sheet
            google_sheet_range = self.get_range_from_matrix(self.line_number, 0, query_stats)
            google_sheet_range = "0. Summary!" + google_sheet_range
            self.line_number += len(query_stats) + 3

            self.google_service_sheets.spreadsheets().values().update(
                spreadsheetId=self.google_sheets_id,
                range=google_sheet_range,
                valueInputOption="RAW",
                body= { 'values': query_stats }
            ).execute()

        main(query)


    def get_range_from_matrix(self, start_y, start_x, matrix):
        """Input: starting y- and x-coordinates and a matrix.
        Output: Coordinates of the matrix (left upper cell and lower right cell) in A1-notation for updating google sheets"""

        max_len_x = 0
        for row in matrix:
            if len(row) > max_len_x:
                max_len_x = len(row)

        max_len_y = len(matrix)

        range_start = chr(64 + start_x + 1) + str(start_y + 1)

        range_end = chr(64 + start_x + max_len_x) + str(start_y + max_len_y)

        return range_start + ":" + range_end


    def close(self):
        """Closes the xlsx writer object"""

        if self.output_destination_type == "local_xlsx" or self.output_destination_type == 'local_folder' :
            logging.info("close writer")
            self.xlsx_workbook.close()




def create_template():
    """Creates a template for the config file in the relative folder, where the script is executed"""

    template = """


# title
# defines the title of the whole set of queries
# OPTIONAL, if not set, timestamp will be used
title = \"TEST QUERIES\"


# description
# defines the textual and human-intended description of the purpose of these queries
# OPTIONAL, if not set, nothing will be used or displayed
description = \"This set of queries is used as a template for showcasing a valid configuration.\"


# output_destination
# defines where to save the results, input can be: 
# * a local path to a folder 
# * a URL for a google sheets document  
# * a URL for a google folder
# NOTE: On windows, folders in a path use backslashes, in such a case it is mandatory to attach a 'r' in front of the quotes, e.g. r\"C:\\Users\\sresch\\..\"
# In the other cases the 'r' is simply ignored; thus best would be to always leave it there.
# OPTIONAL, if not set, folder of executed script will be used
output_destination = r\".\"


# output_format
# defines the format in which the result data shall be saved (currently available: csv, tsv, xml, json, xlsx)
# OPTIONAL, if not set, csv will be used
output_format = \"csv\"


# summary_sample_limit
# defines how many rows shall be displayed in the summary
# OPTIONAL, if not set, 5 will be used
summary_sample_limit = 3


# cooldown_between_queries
# defines how many seconds should be waited between execution of individual queries in order to prevent exhaustion of Google API due to too many writes per time-interval
# OPTIONAL, if not set, 0 will be used
cooldown_between_queries = 0


# endpoint
# defines the SPARQL endpoint against which all the queries are run
# MANDATORY
endpoint = \"http://dbpedia.org/sparql\"


# queries
# defines the set of queries to be run. 
# MANDATAORY
queries = [


    {
        # title
        # OPTIONAL, if not set, timestamp will be used
        \"title\" : \"Optional title of first query\" ,

        # description
        # OPTIONAL, if not set, nothing will be used or displayed
        \"description\" : \"Optional description of first query, used to describe the purpose of the query.\" ,

        # query
        # the sparql query itself
        # NOTE: best practise is to attach a 'r' before the string so that python would not interpret some characters as metacharacters, e.g. \"\\n\"
        # MANDATORY
        \"query\" : r\"\"\"
            SELECT * WHERE {
                ?s ?p ?o
            }
        \"\"\"
    }, 
    {    
        \"query\" : r\"\"\"
            SELECT COUNT (?s) AS ?count_of_subjects_with_type WHERE {
                ?s <http://www.w3.org/1999/02/22-rdf-syntax-ns#type> ?o
            }
        \"\"\"
    },  
    {    
        \"title\" : \"Last query\" , 
        \"description\" : \"This query counts the occurences of distinct predicates\" , 
        \"query\" : r\"\"\"
            SELECT * WHERE {
                ?s <http://www.w3.org/2000/01/rdf-schema#label> ?o
            }
        \"\"\"
    },
]

# Notes on syntax of queries-set:
# * the set of queries is enclosed by '[' and ']'
# * individual queries are enclosed by '{' and '},'
# * All elements of a query (title, description, query) need to be defined using quotes as well as their contents, and both need to be separated by ':'
# * All elements of a query (title, description, query) need to be separated from each other using quotes ','
# * The content of a query needs to be defined using triple quotes, e.g. \"\"\" SELECT * WHERE .... \"\"\"
# * Any indentation (tabs or spaces) do not influence the queries-syntax, they are merely syntactic sugar.


"""
    with open('template.py', 'w') as f:
        f.write(template)





main()