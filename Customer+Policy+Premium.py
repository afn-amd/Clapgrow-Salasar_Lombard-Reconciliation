import pandas as pd
import re
from fuzzywuzzy import fuzz
from collections import defaultdict

# Load data
data1 = pd.read_excel("Unmatched_Data.xlsx", sheet_name="Saiba_Dump")
data2 = pd.read_excel("Unmatched_Data.xlsx", sheet_name='Lombard_Statement')
mdata1 = pd.read_excel("Matched_Data.xlsx", sheet_name="Saiba_Dump")
mdata2 = pd.read_excel("Matched_Data.xlsx", sheet_name='Lombard_Statement')

# Preprocess name function
def preprocess_name(name):
    if pd.isna(name):
        return ""
    name = str(name)
    words_to_omit = [
        'industry', 'industries', 'corp', 'corporation', 'inc', 'incorporated', 'foundation', 
        'company', 'co', 'limited', 'ltd', 'pvt', 'llc', 'llp', 'and', 'pvtltd', '&', 'm/s', 'ms'
    ]
    name = re.sub(r'\b(Mr|Ms|Ltd|LLP|Pvt|Private|Limited|LLC|LTD|Llp|ltd|lp|PLLP|Pllp|P.L.C.|ms|m/s|pvtltd)\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\b(and|AND|And|&)\b', 'and', name, flags=re.IGNORECASE)
    name = re.sub(r'[.,]', '', name)
    for word in words_to_omit:
        name = re.sub(r'\b' + word + r'\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s+', '', name).lower()
    return name

# Compute similarity function
def compute_similarity(data1, data2, threshold=71):
    names1 = pd.Series(data1['CustName'].unique()).astype(str).apply(preprocess_name)
    names2 = pd.Series(data2['INSURED_CUSTOMER_NAME'].unique()).astype(str).apply(preprocess_name)
    results = {}
    index_dict_1 = defaultdict(list)
    index_dict_2 = defaultdict(list)
    
    for name1 in names1:
        for name2 in names2:
            similarity = fuzz.ratio(name1, name2)
            if similarity >= threshold:
                if similarity not in results:
                    results[similarity] = []
                results[similarity].append((name1, name2))
                index_dict_1[name1].extend(data1[data1['CustName'].apply(preprocess_name) == name1]['Index'].tolist())
                index_dict_2[name2].extend(data2[data2['INSURED_CUSTOMER_NAME'].apply(preprocess_name) == name2]['Index'].tolist())
    
    return results, index_dict_1, index_dict_2

# Compute similarity with a threshold of 71%
similarity_dict, index_dict_1, index_dict_2 = compute_similarity(data1, data2, threshold=71)

# Sort the similarity_dict by keys in descending order
sorted_similarity_dict = dict(sorted(similarity_dict.items(), key=lambda item: item[0], reverse=True))

indexPairs = []
for similarity, pairs in sorted_similarity_dict.items():
    for pair in pairs:
        for i in index_dict_1[pair[0]]:
            for j in index_dict_2[pair[1]]:
                result_dict = {i: j}
                if result_dict not in indexPairs:
                    indexPairs.append(result_dict)

# Sort index pairs by the numeric part of the keys
def sort_dicts_by_numeric_key(lst):
    def extract_numeric_key(d):
        key = list(d.keys())[0]
        return int(key[1:])
    return sorted(lst, key=extract_numeric_key)

sorted_list = sort_dicts_by_numeric_key(indexPairs)

# Filter data based on matched indices
index_list_1 = [list(pair.keys())[0] for pair in sorted_list]
index_list_2 = [list(pair.values())[0] for pair in sorted_list]
filtered_data1 = data1[data1['Index'].isin(index_list_1)]
filtered_data2 = data2[data2['Index'].isin(index_list_2)]


# #### Checking Policy Types


from difflib import SequenceMatcher

def preprocess_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower().strip()

def acronym(word):
    # Extract acronym from a phrase
    return ''.join([char for char in word if char.isupper()])

def clean_text(text):
    # Remove punctuation and make lowercase
    return re.sub(r'[^A-Za-z0-9\s]', '', text).lower()

def similarity(a, b):
    # Compute similarity between two strings
    return SequenceMatcher(None, a, b).ratio()

def find_similar_elements(list1, list2, threshold=0.75):
    results = []
    acronyms_list1 = [acronym(str(item)) for item in list1]
    acronyms_list2 = [acronym(str(item)) for item in list2]

    # Check for acronym matches
    for i, acr1 in enumerate(acronyms_list1):
        for j, acr2 in enumerate(acronyms_list2):
            if acr1 == acr2 and acr1:  # Only consider non-empty acronyms
                results.append((list1[i], list2[j]))

    # Check for similarity in remaining data
    for item1 in list1:
        for item2 in list2:
            cleaned_item1 = clean_text(str(item1))
            cleaned_item2 = clean_text(str(item2))
            if similarity(cleaned_item1, cleaned_item2) >= threshold:
                results.append((item1, item2))

    return results

def check_similarity_for_sorted_list(df1, df2, col1, col2, sorted_list, threshold=0.75):
    # Ensure we are working with copies to avoid SettingWithCopyWarning
    df1 = df1.copy()
    df2 = df2.copy()

    # Preprocess the policy names
    df1.loc[:, col1] = df1[col1].apply(preprocess_text)
    df2.loc[:, col2] = df2[col2].apply(preprocess_text)

    matched_sorted_list = []

    # Iterate through the sorted_list to check similarity for each pair
    for pair in sorted_list:
        index1 = list(pair.keys())[0]
        index2 = list(pair.values())[0]

        # Get the policy names corresponding to the indices
        policy_name_1 = df1.loc[df1['Index'] == index1, col1].values[0]
        policy_name_2 = df2.loc[df2['Index'] == index2, col2].values[0]

        # Combine the policy names for similarity checking
        combined_policies_1 = [policy_name_1]
        combined_policies_2 = [policy_name_2]

        # Check for similarity using the new logic
        similar_elements = find_similar_elements(combined_policies_1, combined_policies_2, threshold)

        # If any similar elements are found, keep the pair
        if similar_elements:
            matched_sorted_list.append(pair)

    return matched_sorted_list

# Compute updated sorted list based on policy name similarity
updated_sorted_list = check_similarity_for_sorted_list(filtered_data1, filtered_data2, 'Policy Type', 'PRODUCT_NAME', sorted_list, threshold=0.75)


# #### Checking Premium Amount


# Premium Amount similarity check function
def is_within_2_percent(val1, val2):
    return abs(val1 - val2) <= 0.02 * val1

# Check Premium Amount similarity for updated sorted list
def check_premium_similarity(df1, df2, col1, col2, sorted_list):
    similar_pairs = []
    for pair in sorted_list:
        index1 = list(pair.keys())[0]
        index2 = list(pair.values())[0]
        
        premium1 = df1.loc[df1['Index'] == index1, col1].values[0]
        premium2 = df2.loc[df2['Index'] == index2, col2].values[0]
        
        if is_within_2_percent(premium1, premium2):
            similar_pairs.append(pair)
    
    return similar_pairs

# Compute premium similarity
premium_similarity_list = check_premium_similarity(filtered_data1, filtered_data2, 'OD Premium', 'APPLICABLE_PREMIUM_AMOUNT', updated_sorted_list)

# Filter data based on matched premium similarity indices
keys = [list(pair.keys())[0] for pair in premium_similarity_list]
values = [list(pair.values())[0] for pair in premium_similarity_list]

Saiba_Dump = filtered_data1[filtered_data1['Index'].isin(keys)]
Lombard_Statement = filtered_data2[filtered_data2['Index'].isin(values)]

# Function to create comma-separated matching indices
def create_comma_separated_indices(df, col1, matching_col, pairs):
    matching_index_dict = defaultdict(list)
    for pair in pairs:
        index1 = list(pair.keys())[0]
        index2 = list(pair.values())[0]
        matching_index_dict[index1].append(index2)
        matching_index_dict[index2].append(index1)

    # Use .loc to set values in the DataFrame
    df.loc[:, matching_col] = df[col1].apply(lambda x: ', '.join(matching_index_dict[x]))
    df.loc[:, 'Matching_Attribute'] = 'Customer+Policy+Premium'
    return df

# Add Matching_Index and Matching_Attribute columns
Saiba_Dump = create_comma_separated_indices(Saiba_Dump.copy(), 'Index', 'Matching_Index', premium_similarity_list)
Lombard_Statement = create_comma_separated_indices(Lombard_Statement.copy(), 'Index', 'Matching_Index', premium_similarity_list)

# Reorder columns to place Matching_Index and Matching_Attribute at 2nd and 3rd positions
Saiba_Dump = Saiba_Dump[['Index', 'Matching_Index', 'Matching_Attribute'] + [col for col in Saiba_Dump.columns if col not in ['Index', 'Matching_Index', 'Matching_Attribute']]]
Lombard_Statement = Lombard_Statement[['Index', 'Matching_Index', 'Matching_Attribute'] + [col for col in Lombard_Statement.columns if col not in ['Index', 'Matching_Index', 'Matching_Attribute']]]

# Eliminate data from "Unmatched_Data.xlsx" which is present in the Pandas DataFrames
data1 = data1[~data1['Index'].isin(Saiba_Dump['Index'])]
data2 = data2[~data2['Index'].isin(Lombard_Statement['Index'])]

# Function to extract numeric part from the 'Index' column for sorting
def extract_numeric_part(index_series):
    return index_series.str.extract(r'(\d+)$').astype(int)[0]

# Add values to "Matched_Data.xlsx" that are present in the Pandas DataFrames
# Merge the dataframes on 'Index' column and concatenate to match the order
mdata1 = pd.concat([mdata1, Saiba_Dump]).drop_duplicates(subset='Index').sort_values(by='Index', key=extract_numeric_part)
mdata2 = pd.concat([mdata2, Lombard_Statement]).drop_duplicates(subset='Index').sort_values(by='Index', key=extract_numeric_part)

# Save the updated data back to Excel files
with pd.ExcelWriter('Unmatched_Data.xlsx') as writer:
    data1.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    data2.to_excel(writer, sheet_name='Lombard_Statement', index=False)

with pd.ExcelWriter('Matched_Data.xlsx') as writer:
    mdata1.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    mdata2.to_excel(writer, sheet_name='Lombard_Statement', index=False)