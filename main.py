import regex as re
import pandas as pd
from pyzotero import zotero
import Levenshtein

## Key Zotero Variables ##
API_KEY = 'hzhh8Foo27nRP1yDcpi6Twaa'
LIBRARY_TYPE = 'group'
LIBRARY_ID = '4907635'
SUBLIBRARY_ID = 'D7X5DJBX'



## Name Checker ##


def calculate_similarity(name1, name2):
    return Levenshtein.distance(name1.lower(), name2.lower())


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

    # Fetch all subcollections under the specified collid
    allCollections = zot.all_collections(collid=SUBLIBRARY_ID)

    # Create empty lists and a dictionary to store data
    authors_collections = {}  
    # Iterate over collections and subcollections
    for collection in allCollections:
        collectionItems = zot.collection_items_top(collection['key'])
        if len(collectionItems) == 0:
            continue

        # Filter out any non-dictionary items
        collectionItems = [item for item in collectionItems if isinstance(item, dict)]

        for item in collectionItems:
            for citation in item.get('data', {}).get('creators', []):
                first_name = citation.get('firstName', '')
                last_name = citation.get('lastName', '')
                full_name = f"{first_name} {last_name}".strip()
                subcollection_name = collection['data'].get('name', '')  # Add subcollection name

                # Add author and associated collection/subcollection to the dictionary
                author_key = f"{first_name} {last_name}"
                if author_key not in authors_collections:
                    authors_collections[author_key] = {'First Name': first_name, 'Last Name': last_name, 'Full Name': full_name, 'Subcollections': []}
                if subcollection_name not in authors_collections[author_key]['Subcollections']:
                    authors_collections[author_key]['Subcollections'].append(subcollection_name)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(authors_collections.values())

    # Convert 'Subcollections' column to comma-separated string
    df['Subcollections'] = df['Subcollections'].apply(lambda x: ', '.join(x))

    # Drop duplicate authors to keep each author only once
    df.drop_duplicates(inplace=True)

    # Save the DataFrame to a CSV file
    csv_file = 'zotero_creators.csv'
    df.to_csv(csv_file, index=False)
##


def esa_name_combiner():
    import pandas as pd

    initial_sheet = input("Please copy and paste your file pathname for ESA spreadsheet: ")

    # List of sheet names
    sheet_names = [str(year) for year in range(2022, 2005, -1)]

    # Create empty lists to store data
    full_names = []
    first_names = []
    last_names = []
    titles = []

    # Loop through each sheet
    for sheet_name in sheet_names:
        # Read the sheet into a DataFrame
        try:
            df = pd.read_excel(initial_sheet, sheet_name=sheet_name)
        except:
            df = pd.read_csv(initial_sheet, sheet_name=sheet_name)

        # Remove leading/trailing white spaces from column names
        df.columns = df.columns.str.strip()

        # Verify that the 'Professor' column exists in the DataFrame
        if 'Professor' not in df.columns:
            print(f"Error: 'Professor' column not found in sheet '{sheet_name}'. Skipping...")
            continue

        # Extract the names and titles columns
        names = df['Professor']

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
        if 'Title of Presentation' in df.columns:
            titles.extend(df['Title of Presentation'])
        elif 'Title of presentation' in df.columns:
            titles.extend(df['Title of presentation'])

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




def compare(spreadsheet1, spreadsheet2):

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
                break

        # Update the 'Appeared' column to 'Yes' if the name is found
        if appeared:
            df_a.at[index_a, 'Appeared'] = 'Yes'

    # Save the updated df_a with the indicator column
    df_a.to_csv('citation_checker2.csv', index=False)


zotero_name_scraper()
esa_name_combiner()
compare("esa_names.csv", "zotero_creators.csv" )
# /Users/emmettkliger/Downloads/Zotero-and-ESA-Name-Comparison-Tool-main_spellcheck_specific_collection 2/ESA conferences.xlsx
