import pandas as pd

# Read Excel File
data = pd.read_excel('policy_filter (1).xlsx')

# Delete Draft contract
del_draft = data[data['Վիճակ'] != 'draft']

# Get Operating contract
df = del_draft[del_draft['Վիճակ'] == 'operating']

# Keep columns needed
columns_to_keep = ['Ապահովադիր', 'Անձնագրի համար', 'Սկիզբ', 'Ակտիվ էˋ մինչև', 'Ապ. գումար', 'Ապ. վճար', 'Արժույթ']
df = df[columns_to_keep]

# Write DataFrame to Excel
df.to_excel('new_excel.xlsx', index=False)

# Print DataFrame
print(df)
