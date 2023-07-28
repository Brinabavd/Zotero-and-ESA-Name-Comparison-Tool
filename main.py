## This code was created on July 28th, 2023 by Emmett Kliger and Arin Jaff
## This code implements the Zotero API and the Crossref REST API.

## This code serves a few purposes:

##First, it can fully extract zotero creator information from a zotero library 
## and export citation information to a .csv file.  Subcollections information is included.
## This output comes in zotero_creators.csv.

## Second, it can extract information from a spreadsheet and export it to a .csv.
## We used this to convert a large manually entered dataset from excel into a more easily
## readable form, while also eliminating duplicates/mispelled words with find_mispellings().
## (You can lower/raise the sensitivity of this spellcheck by altering the levenshtein global threshold) 
## You should check that this code can handle the formatting of your spreadsheet by looking 
## at the esa_name_combiner() function and the names of the spreadsheet column titles it reads in.
## The easiest solution is to change the name of the columns on your spreadsheet, but edit the
## code as needed for your spreadsheet. Check ~line 250 for this information.
## Output is lablelled esa_names.csv

## Third, this code can comb through the excel spreadsheet's authors from the resulting .csv, 
## and can extract the DOI information from them. This is useful for citing papers we did not
## already have in our zotero library that were in our excel database. Note that this uses the 
## Crossref REST API, and if you are combing through presentations like we did, there will need to
## be a great similarity between the name referenced and the logged name of the presentation.
## To account for some mispellings, we used another levenshtein global threshold which can be 
## changed, but as you raise it and lower the sensitivtiy of checking you may allow more incorrect
## DOIs to be collected with similar keywords. This also occasionally runs into retry errors, which
## is accounted for in make_request_with_retry().
## output is labelled dois_names.csv


## Finally, this can compare the two sources' information to check for overlap between the two.
## We used this to find overlap in authors that we had cited in our large spreadsheet
## With authors that we had in our zotero library. The output is a .csv file labelled 
## citation_checker.csv

## -- IMPORTANT CHANGES TO MAKE FOR YOUR RUNNING --

## You may need to check the directories referenced to make sure they are valid with your names,
## and use an absolute path rather than relative if you run into trouble.

## Request your own Zotero API Key from https://www.zotero.org/settings/keys, and adjust
## the settings to your library as needed. If you run into errors, the permissions may be
## the reason why.

## You can find your zotero library/sublibrary IDs by checking the web URL when you open your library

## Doi Scraping is currently commented out at the bottom of this code. This is because it takes an 
## EXTREMELY long time to run, and we recommend you only do this once when you are fully ready and 
## are sure that your data works with the other steps we have in our code.

## For DOI scraping, replace your email with the placeholder youremail@youremail.com. This will place
## you on Crossref's polite user list and you will have faster/better/consistent access to their api.

import regex as re
import pandas as pd
from pyzotero import zotero
import Levenshtein as lv
import requests
from fuzzywuzzy import fuzz
import time

## Key Zotero Variables for identifying library ##
ZOTERO_API_KEY = 'hzhh8Foo27nRP1yDcpi6Twaa' 
LIBRARY_TYPE = 'group'
LIBRARY_ID = '4907635'
SUBLIBRARY_ID = 'D7X5DJBX'

## Key Levenshtein Variables ##
DOI_SIMILARITY_THRESHOLD = 10
MISPELLINGS_SIMILARITY_HRESHOLD = 3

## Key global reference variables ##
authors_collections = {}
full_title_list = []

## Directories ##
SHEET_DIRECTORY = 'ESA Conferences.xlsx'



## Similarity check ##
def calculate_similarity(name1, name2):
    return lv.distance(name1.lower(), name2.lower())


## Finds duplicates with mispellings
## and cleans them out of some list "names"
def find_misspellings(names):
    cleaned_names = []
    misspellings = []

    initials_last_names = set()
    for name in names:
        try:
            first_name, last_name = name.split(' ', 1)
        except Exception:
            print("check this entry manually:", name) # A few names (12) are formatted strangely, we may need to check them manually
            cleaned_names.append(name)
            continue
        initials_last_name = (first_name[0], last_name) #[0] first initial, [1], last name
        is_unique = True

        for existing_name in initials_last_names:
            if (initials_last_name[0].lower() == existing_name[0].lower()) and (initials_last_name[1].lower() == existing_name[1].lower()): #if initial and last name
                similarity_score = calculate_similarity(name, existing_name[0] + ' ' + existing_name[1]) # calculate similarity score
                if similarity_score <= 3:  # AND if similar
                    #print(name, "is similar to", existing_name[0] + ' ' + existing_name[1]) ## FOR CHECKING IN TERMINAL
                    misspellings.append(name) # goes into misspellings column (nothing happens with this set right now)
                    is_unique = False # not logged in cleaned names
                    break

        if is_unique:
            #print(name, "is a unique name:", initials_last_name) ## FOR CHECKING IN TERMINAL
            cleaned_names.append(name)
            initials_last_names.add(initials_last_name)

    return cleaned_names, misspellings





## Zotero Scraper ##

def zotero_name_scraper():
    # Create a Zotero API client
    zot = zotero.Zotero(LIBRARY_ID, LIBRARY_TYPE, API_KEY)
    #items = zot.everything(zot.items())

    allCollections = (zot.all_collections(collid=SUBLIBRARY_ID))
    allCollections = [zot.collection_items_top(c['key']) for c in allCollections],
    # allCollections is a list of lists, each being a subcollection
    # of the collid 'D7X5DJBX'

    # Create empty lists to store data
    first_names = []
    last_names = []
    full_names = []


    # Iterate over items and extract author info

    for collectionItems in allCollections:
        if(len(collectionItems) == 0):
                continue
        for item in collectionItems:
            if(len(item) == 0):
                continue
            for citation in item:
                print("\n ITEM:",item)
                creators = citation['data'].get('creators', [])
                for creator in creators:
                    first_name = creator.get('firstName', '')
                    last_name = creator.get('lastName', '')
                    full_name = f"{first_name} {last_name}".strip()
                    first_names.append(first_name)
                    last_names.append(last_name)
                    full_names.append(full_name)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame({'First Name': first_names, 'Last Name': last_names, 'Full Name': full_names})

    # Drop duplicate authors to keep each author only once
    df.drop_duplicates(inplace=True)

    # check for mispellings

    # Save the DataFrame to a CSV file
    csv_file = 'zotero_creators.csv'
    df.to_csv(csv_file, index=False)

    # Print a message indicating the code has finished
    print("Zotero Scraping completed. Data saved to CSV file.")

## Sheeet Scraper ##

def esa_name_combiner():

    initial_sheet = SHEET_DIRECTORY
    print("Scraping ESA spreadsheet using", initial_sheet, "as filepath ...")

    # List of sheet names
    sheet_names = [str(year) for year in range(2022, 2005, -1)]

    # Create empty lists to store data
    full_names = []
    first_names = []
    last_names = []
    global titles
    titles = []

    # Loop through each sheet
    for sheet_name in sheet_names:
        # Read the sheet into a DataFrame
        try:
            dfex = pd.read_excel(initial_sheet, sheet_name=sheet_name)
        except:
            dfex = pd.read_csv(initial_sheet, sheet_name=sheet_name)

        # Remove leading/trailing white spaces from column names
        dfex.columns = dfex.columns.str.strip()

        # Verify that the 'Professor' column exists in the DataFrame
        if 'Professor' not in dfex.columns:
            print(f"Error: 'Professor' column not found in sheet '{sheet_name}'. Skipping...")
            continue

        # Extract the names and titles columns
        names = dfex['Professor']

        # Remove variations of "CANCELLED" or "CANCELED" from names
        names = names.str.replace(r'CANCELLED|CANCELED', '', regex=True)

        # Remove "PRESENTATION" from names
        names = names.str.replace('PRESENTATION', '')

        # Remove special characters before the first letter
        names = names.str.replace(r'^[^A-Za-z]+', '', regex=True)
        # Remove anything within parentheses and the parentheses themselves
        names = names.str.replace(r'\([^)]*\)', '', regex=True)

        # Split names by comma or "&" and create separate rows for each name
        names_split = names.str.split('[,&]', expand=False)
        names_list, misspellings = find_misspellings([name.strip() for sublist in names_split if isinstance(sublist, list) for name in sublist])

        # Store the names in the respective lists
        full_names.extend(names_list)
        first_names.extend([name.split(' ', 1)[0] for name in names_list])
        last_names.extend([name.split(' ', 1)[1] if len(name.split(' ', 1)) > 1 else '' for name in names_list])

        # Handle different column names for 'Title of Presentation'
        if 'Title of Presentation' in dfex.columns:
            titles.extend(dfex['Title of Presentation'])
        elif 'Title of presentation' in dfex.columns:
            titles.extend(dfex['Title of presentation'])

        # Handle empty cells in titles
        for _ in range(len(names_list) - len(titles)):
            titles.append('')

    # Ensure the arrays have the same length
    min_length = min(len(full_names), len(first_names), len(last_names), len(titles))
    full_names = full_names[:min_length]
    first_names = first_names[:min_length]
    last_names = last_names[:min_length]
    titles = titles[:min_length]


    # Create a DataFrame from the extracted data
    data = {'Full Name': full_names, 'First Name': first_names, 'Last Name': last_names, 'Title of Presentation': titles}
    df = pd.DataFrame(data)

    if df.empty:
        print("Error: No data found. Please verify the column name and ensure the data exists.")

    # Remove duplicates from the DataFrame
    df.drop_duplicates(subset=['Full Name'], inplace=True)

    # Set column names for the DataFrame
    df.columns = ['Full Name', 'First Name', 'Last Name', 'Title of Presentation']

    # Write the data to a CSV file
    df.to_csv('esa_names.csv', index=False)

    # Print a message indicating the code has finished
    print("ESA Name Combination completed. Data saved to CSV file.")


## Comparing Sheet + Zotero ##

def compare(spreadsheet1, spreadsheet2):


    print("Comparing ESA Spreadsheet with Zotero Citations ...")

    # Read the two spreadsheets
    try:
        df_a = pd.read_excel(spreadsheet1, engine='openpyxl')
        df_b = pd.read_excel(spreadsheet2, engine='openpyxl')
    except:
        df_a = pd.read_csv(spreadsheet1)
        df_b = pd.read_csv(spreadsheet2)

    # Add a new column 'Appeared' to df_a with default value 'No'
    df_a['Appeared'] = 'No'

    # Convert first names and last names to lowercase for case-insensitive comparison
    df_a['First Name'] = df_a['First Name'].str.lower()
    df_a['Last Name'] = df_a['Last Name'].str.lower()
    df_b['First Name'] = df_b['First Name'].str.lower()
    df_b['Last Name'] = df_b['Last Name'].str.lower()

    # Iterate over each row in df_a
    for index_a, row_a in df_a.iterrows():
        first_name_a = row_a['First Name']
        last_name_a = row_a['Last Name']
        appeared = False
        associated_value = None

        # Check if first name or last name is NaN in df_a
        if pd.isna(first_name_a) or pd.isna(last_name_a):
            continue

        # Iterate over each row in df_b
        for index_b, row_b in df_b.iterrows():
            first_name_b = row_b['First Name']
            last_name_b = row_b['Last Name']

            # Check if first name or last name is NaN in df_b
            if pd.isna(first_name_b) or pd.isna(last_name_b):
                continue

            # Check if the first letter of first name is the same and last name is entirely the same
            if first_name_a[0] == first_name_b[0] and last_name_a == last_name_b:
                appeared = True
                associated_value = authors_collections.get(f"{first_name_a} {last_name_a}")
                break

        # Update the 'Appeared' column to 'Yes' if the name is found
        if appeared:
            df_a.at[index_a, 'Appeared'] = 'Yes'
            df_a.at[index_a, 'Subcollections'] = str(associated_value)

    # Save the updated df_a with the indicator column
    df_a.to_csv('citation_checker.csv', index=False)

    print("Finished comparing, data saved to CSV file.")
    
"""
## Finds best match in DOI search ##

def find_best_match(title, items):
    best_match = None
    min_distance = float('inf')

    for item in items:
        # Extract the individual titles from the list
        item_titles = item.get("title", [])
        for item_title in item_titles:
            if not isinstance(title, str) or not isinstance(item_title, str):
                continue

            distance = lv.distance(title.lower(), item_title.lower())
            if distance < min_distance:
                min_distance = distance
                best_match = item
                best_match["title"] = item_title  # Store the best-matched title

    if min_distance > MISPELLINGS_SIMILARITY_HRESHOLD:
        print("couldn't find a DOI for",title)
        return None

    return best_match

## Exponential backoff + retries for DOI searches on Crossref API ##

def make_request_with_retry(url, params=None, headers=None, max_retries=3, retry_delay=1):
    retries = 0

    while retries < max_retries:
        try:
            response = requests.get(url, params=params, headers=headers, verify=True)
            response.raise_for_status()  # Raise an exception for unsuccessful response status codes
            return response.json()
        except (requests.RequestException, ValueError) as e:
            print(f"Request failed: {e}")
            retries += 1
            if retries < max_retries:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponential backoff: Increase the delay for each retry
            else:
                print("Max retries exceeded. Request failed.")
                raise 




def get_dois_from_titles():
    print("Extracting DOIs from ESA titles...")
    print(len(titles))
    base_url = "https://api.crossref.org/works"
    headers = {"Accept": "application/json"}
    params = {"filter": "has-full-text:true", "mailto": "YOUREMAIL@YOUREMAIL.COM"}
    dois = {}
    i = 1
    for title in titles:
        params["query.bibliographic"] = title
        try:
            response_data = make_request_with_retry(base_url, params=params, headers=headers)
            if "message" in response_data and "items" in response_data["message"]:
                best_match = find_best_match(title, response_data["message"]["items"])
                if best_match:
                    found_title = best_match.get("title", "")
                    doi = best_match.get("DOI", "")
                    print("For", title, "found", doi, "-- entry #", i)
                    i += 1
                    if found_title and doi:
                        dois[found_title] = doi
                    elif found_title.lower() == title.lower():
                        # Perform case-insensitive match for exact title match
                        dois[found_title] = doi
        except requests.exceptions.RequestException as e:
            print(f"Request failed for title '{title}': {e}")


    print("FINAL SET OF DOIS: ", dois)
    

    # Safety
    dfSafety = pd.DataFrame(list(dois.items()), columns=['Presentation Names', 'DOI Numbers'])

    # Export the DataFrame to a CSV file
    dfSafety.to_csv('dois_names.csv', index=False)
    print("Collected DOIs in csv file.")
"""

esa_name_combiner()
zotero_name_scraper()
compare("esa_names.csv", "zotero_creators.csv" )

## FOR DOI COLLECTION ##
#get_dois_from_titles()
