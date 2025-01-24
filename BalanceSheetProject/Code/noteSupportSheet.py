import pandas as pd

# Load the input Excel file
input_file = "C:\\Users\\jadit\\OneDrive\\Desktop\\BalanceSheetProject\\ExcelFiles\\InputForScheduleTemplate\\InputForASchedule - Copy (4).xlsx"
df = pd.read_excel(input_file)

# Create a new DataFrame with the desired columns
output_df = pd.DataFrame(columns=['Note', 'Major Head', 'Sub Head', 'Sub to sub head', 'Nature'])

# Initialize a counter for assigning integers
counter = 1

# Iterate over the rows and create the 'Note' column
for index, row in df.iterrows():
    major_head = row['Major Head']
    sub_head = row['Sub Head ']
    sub_to_sub_head = row['Sub to sub Head']
    nature = row['Nature']
    
    # Create the 'Note' value with the integer under the 'Note' heading
    note = f"Note {counter}"
    
    # Create a new row in the output DataFrame
    output_df = output_df.append({'Note': note, 'Major Head': major_head, 'Sub Head': sub_head, 'Sub to Sub Head': sub_to_sub_head, 'Nature': nature}, ignore_index=True)
    
    # Increment the counter
    counter += 1

# Save the output DataFrame to a new Excel file
output_file = "C:\\Users\\jadit\\OneDrive\\Desktop\\BalanceSheetProject\\GeneratedOutput\\NotesupportSheet"
output_df.to_excel(output_file, index=False)