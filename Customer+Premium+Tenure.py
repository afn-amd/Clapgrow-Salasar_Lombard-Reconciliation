import pandas as pd
import re
from fuzzywuzzy import fuzz
from collections import defaultdict

# Load data
data1 = pd.read_excel("Unmatched_Data.xlsx", sheet_name="Saiba_Dump")
data2 = pd.read_excel("Unmatched_Data.xlsx", sheet_name='Lombard_Statement')
mdata1 = pd.read_excel("Matched_Data.xlsx", sheet_name="Saiba_Dump")
mdata2 = pd.read_excel("Matched_Data.xlsx", sheet_name='Lombard_Statement')


# #### Checking Customer Names


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


# #### Checking Premium Amount


# Premium Amount similarity check function
def is_within_2_percent(val1, val2):
    return abs(val1 - val2) <= 0.02 * val1

# Check Premium Amount similarity for the sorted list
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
premium_similarity_list = check_premium_similarity(filtered_data1, filtered_data2, 'OD Premium', 'APPLICABLE_PREMIUM_AMOUNT', sorted_list)


# #### Checking Start+End Dates


def check_tenure_similarity(df1, df2, similarity_list):
    # Lists to store indices of matched policies
    Index1 = []
    Index2 = []

    # Iterate through each dictionary in the similarity list
    for pair in similarity_list:
        index1 = list(pair.keys())[0]
        index2 = list(pair.values())[0]

        # Retrieve the start and end dates from both dataframes
        start_date1 = df1.loc[df1['Index'] == index1, 'Policy_StartDate'].values[0]
        end_date1 = df1.loc[df1['Index'] == index1, 'Exp. Date'].values[0]
        start_date2 = df2.loc[df2['Index'] == index2, 'POLICY_START_DATE'].values[0]
        end_date2 = df2.loc[df2['Index'] == index2, 'POLICY_END_DATE'].values[0]

        # Check if both the start and end dates match
        if start_date1 == start_date2 and end_date1 == end_date2:
            Index1.append(index1)
            Index2.append(index2)
    
    return Index1, Index2

Index1, Index2 = check_tenure_similarity(filtered_data1, filtered_data2, premium_similarity_list)

final_filtered_data1 = filtered_data1[filtered_data1['Index'].isin(Index1)]
final_filtered_data2 = filtered_data2[filtered_data2['Index'].isin(Index2)]

# Function to add new columns and reorder them
def add_matching_columns(df1, df2, index1, index2):
    # Create copies to avoid SettingWithCopyWarning
    df1 = df1.copy()
    df2 = df2.copy()

    # Initialize the new columns with empty strings
    df1.loc[:, 'Matching_Index'] = ""
    df1.loc[:, 'Matching_Attribute'] = "Customer+Premium+Tenure"
    df2.loc[:, 'Matching_Index'] = ""
    df2.loc[:, 'Matching_Attribute'] = "Customer+Premium+Tenure"
    
    # Create a dictionary to map indices
    match_dict_1_to_2 = defaultdict(list)
    match_dict_2_to_1 = defaultdict(list)
    
    for i, j in zip(index1, index2):
        match_dict_1_to_2[i].append(j)
        match_dict_2_to_1[j].append(i)
    
    # Fill in the Matching_Index columns
    for i in match_dict_1_to_2:
        df1.loc[df1['Index'] == i, 'Matching_Index'] = ', '.join(match_dict_1_to_2[i])
    for j in match_dict_2_to_1:
        df2.loc[df2['Index'] == j, 'Matching_Index'] = ', '.join(match_dict_2_to_1[j])
    
    # Reorder columns to place 'Matching_Index' and 'Matching_Attribute' as columns 2 and 3
    columns1 = list(df1.columns)
    columns2 = list(df2.columns)
    
    # Reorder df1
    columns1.remove('Matching_Index')
    columns1.remove('Matching_Attribute')
    new_order1 = columns1[:1] + ['Matching_Index', 'Matching_Attribute'] + columns1[1:]
    df1 = df1[new_order1]
    
    # Reorder df2
    columns2.remove('Matching_Index')
    columns2.remove('Matching_Attribute')
    new_order2 = columns2[:1] + ['Matching_Index', 'Matching_Attribute'] + columns2[1:]
    df2 = df2[new_order2]
    
    return df1, df2

final_filtered_data1, final_filtered_data2 = add_matching_columns(final_filtered_data1, final_filtered_data2, Index1, Index2)

# Eliminate data from "Unmatched_Data.xlsx" which is present in the Pandas DataFrames
data1 = data1[~data1['Index'].isin(final_filtered_data1['Index'])]
data2 = data2[~data2['Index'].isin(final_filtered_data2['Index'])]

# Function to extract numeric part from the 'Index' column for sorting
def extract_numeric_part(index_series):
    return index_series.str.extract(r'(\d+)$').astype(int)[0]

# Add values to "Matched_Data.xlsx" that are present in the Pandas DataFrames
# Merge the dataframes on 'Index' column and concatenate to match the order
mdata1 = pd.concat([mdata1, final_filtered_data1]).drop_duplicates(subset='Index').sort_values(by='Index', key=extract_numeric_part)
mdata2 = pd.concat([mdata2, final_filtered_data2]).drop_duplicates(subset='Index').sort_values(by='Index', key=extract_numeric_part)

# Save the updated data back to Excel files
with pd.ExcelWriter('Unmatched_Data.xlsx') as writer:
    data1.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    data2.to_excel(writer, sheet_name='Lombard_Statement', index=False)

with pd.ExcelWriter('Matched_Data.xlsx') as writer:
    mdata1.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    mdata2.to_excel(writer, sheet_name='Lombard_Statement', index=False)