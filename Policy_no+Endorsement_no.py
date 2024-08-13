import pandas as pd

# Load the datasets
broker_data = pd.read_excel('Saiba_Dump.xls', engine='xlrd')
company_data = pd.read_excel('Lombard_Statement.xlsx', sheet_name='RAW STATEMENT')

# Add an indexing column to each dataset
broker_data['Index'] = ['S' + str(i + 1) for i in range(len(broker_data))]
company_data['Index'] = ['L' + str(i + 1) for i in range(len(company_data))]

# Reorder columns to make 'Index' the first column
broker_data = broker_data[['Index'] + [col for col in broker_data.columns if col != 'Index']]
company_data = company_data[['Index'] + [col for col in company_data.columns if col != 'Index']]

# Save to a new Excel file with two sheets
with pd.ExcelWriter('Combined_Data.xlsx') as writer:
    broker_data.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    company_data.to_excel(writer, sheet_name='Lombard_Statement', index=False)

data1 = pd.read_excel('Combined_Data.xlsx', sheet_name='Saiba_Dump')
data2 = pd.read_excel('Combined_Data.xlsx', sheet_name='Lombard_Statement')

# Initialize dictionaries to store matching indices and attributes
broker_matches = {index: [] for index in data1['Index']}
company_matches = {index: [] for index in data2['Index']}
broker_attributes = {index: [] for index in data1['Index']}
company_attributes = {index: [] for index in data2['Index']}


# #### Checking Policy Numbers


# Find common values between 'PolicyNo' in data1 and 'POL_NUM_TXT' in data2
common_values_policy = set(data1['PolicyNo']).intersection(set(data2['POL_NUM_TXT']))

# Get the index pairs for the common values and store matching attribute
for value in common_values_policy:
    index1_list = data1[data1['PolicyNo'] == value]['Index'].values
    index2_list = data2[data2['POL_NUM_TXT'] == value]['Index'].values
    for i1 in index1_list:
        for i2 in index2_list:
            broker_matches[i1].append(i2)
            company_matches[i2].append(i1)
            broker_attributes[i1].append('POL_NUM_TXT')
            company_attributes[i2].append('PolicyNo')


# #### Checking Endoresement Numbers


# Find common values between 'EndoNo' in data1 and 'POL_NUM_TXT' in data2
common_values_endo = set(data1['EndoNo']).intersection(set(data2['POL_NUM_TXT']))

# Get the index pairs for the common values and store the matching attribute
for value in common_values_endo:
    index1_list = data1[data1['EndoNo'] == value]['Index'].values
    index2_list = data2[data2['POL_NUM_TXT'] == value]['Index'].values
    for i1 in index1_list:
        for i2 in index2_list:
            broker_matches[i1].append(i2)
            company_matches[i2].append(i1)
            broker_attributes[i1].append('POL_NUM_TXT')
            company_attributes[i2].append('EndoNo')

# Convert lists to comma-separated strings
data1['Matching_Index'] = data1['Index'].map(lambda x: ', '.join(broker_matches[x]))
data2['Matching_Index'] = data2['Index'].map(lambda x: ', '.join(company_matches[x]))
data1['Matching_Attribute'] = data1['Index'].map(lambda x: ', '.join(broker_attributes[x]))
data2['Matching_Attribute'] = data2['Index'].map(lambda x: ', '.join(company_attributes[x]))

# Reorder columns to make 'Matching_Index' and 'Matching_Attribute' the second and third columns
data1 = data1[['Index', 'Matching_Index', 'Matching_Attribute'] + [col for col in data1.columns if col not in ['Index', 'Matching_Index', 'Matching_Attribute']]]
data2 = data2[['Index', 'Matching_Index', 'Matching_Attribute'] + [col for col in data2.columns if col not in ['Index', 'Matching_Index', 'Matching_Attribute']]]

# Create dataframes for matched and unmatched data
matched_broker_data = data1[data1['PolicyNo'].isin(common_values_policy) | data1['EndoNo'].isin(common_values_endo)]
unmatched_broker_data = data1[~data1['PolicyNo'].isin(common_values_policy) & ~data1['EndoNo'].isin(common_values_endo)]
matched_company_data = data2[data2['POL_NUM_TXT'].isin(common_values_policy.union(common_values_endo))]
unmatched_company_data = data2[~data2['POL_NUM_TXT'].isin(common_values_policy.union(common_values_endo))]

# Save matched data to a new Excel file with two sheets
with pd.ExcelWriter('Matched_Data.xlsx') as writer:
    matched_broker_data.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    matched_company_data.to_excel(writer, sheet_name='Lombard_Statement', index=False)

# Drop the columns "Matching_Index", "Matching_Attribute" from the unmatched dataframes
unmatched_broker_data = unmatched_broker_data.drop(columns=['Matching_Index', 'Matching_Attribute'])
unmatched_company_data = unmatched_company_data.drop(columns=['Matching_Index', 'Matching_Attribute'])

# Save unmatched data to a new Excel file with two sheets
with pd.ExcelWriter('Unmatched_Data.xlsx') as writer:
    unmatched_broker_data.to_excel(writer, sheet_name='Saiba_Dump', index=False)
    unmatched_company_data.to_excel(writer, sheet_name='Lombard_Statement', index=False)