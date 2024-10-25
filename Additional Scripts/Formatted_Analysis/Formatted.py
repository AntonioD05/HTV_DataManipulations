import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# defining paths: change to your own files but make sure input file has the same format as example input file and make sure consent file corresponds with cohort
input_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted.xlsx"
consent_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_Consent_SleepMatch.xlsx"
output_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted(Cleaned_Analysis).xlsx"
cohort = "SP24" # or "FA23" or "SP23" to specify which cohort

# this sees which consent file you are using because each of them have different formats
if cohort == "SP24":
    # loading consent data for SP24
    consent_df = pd.read_excel(consent_file_path)
    consented_ids = consent_df[consent_df.iloc[:, 6].astype(str).str.upper() != 'FALSE'].iloc[:, 0].astype(str)
    tier_info = consent_df[consent_df.iloc[:, 6] != 'FALSE'].iloc[:, [0, 3]]
    tier_info.columns = ['Participant ID', 'Intervention Group']
    tier_info['Participant ID'] = tier_info['Participant ID'].astype(str)
elif cohort == "FA23":
    # loading consent data for FA23
    consent_df = pd.read_excel(consent_file_path)
    consented_ids = consent_df[consent_df.iloc[:, 1].astype(str).str.upper() != 'FALSE'].iloc[:, 0].astype(str)
    tier_info = consent_df[consent_df.iloc[:, 1] != 'FALSE'].iloc[:, [0, 2]]
    tier_info.columns = ['Participant ID', 'Intervention Group']
    tier_info['Participant ID'] = tier_info['Participant ID'].astype(str)
elif(cohort == "SP23"):
    # loading consent data for SP23
    consent_df = pd.read_excel(consent_file_path)
    consented_ids = consent_df[consent_df.iloc[:, 2].astype(str).str.upper() != 'FALSE'].iloc[:, 0].astype(str)
    tier_info = consent_df[consent_df.iloc[:, 2] != 'FALSE'].iloc[:, [0, 4]]
    tier_info.columns = ['Participant ID', 'Intervention Group']
    tier_info['Participant ID'] = tier_info['Participant ID'].astype(str)


# function to find first and last consecutive days with valid data
def find_consecutive_days(data, num_days=5, first=True):
    valid_indices = data.dropna().where(data != 0).dropna().index.tolist()
    if len(valid_indices) < num_days:
        return None, None
    if first: # first consecutive days
        for i in range(len(valid_indices) - num_days + 1):
            if all(data.index[j] - data.index[i] == j - i for j in range(i, i + num_days)):
                return data.iloc[i:i + num_days], (valid_indices[i], valid_indices[i + num_days - 1])
    else: # last consecutive days
        for i in range(len(valid_indices) - num_days, -1, -1):
            if all(data.index[j] - data.index[i] == j - i for j in range(i, i + num_days)):
                return data.iloc[i:i + num_days], (valid_indices[i], valid_indices[i + num_days - 1])
    return None, None

# initialize ExcelWriter
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # load Excel file
    xlsx = pd.ExcelFile(input_file_path)
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet_name)
        #df = df.drop(index=range(42, 47))  # uncomment only for SP24 (Spring Break data doesnt count)
        df = df.iloc[:, 2:]  # adjusting data frame
        df.columns = df.columns.astype(str)  
        df = df.filter(items=consented_ids.tolist())  # filter by consented IDs

        # prepare results dictionary for this sheet
        results = {'Participant ID': df.columns}
        days_distance = []
        for participant_id in df.columns:
            column_data = df[participant_id].dropna()
            recorded_data = column_data[column_data != 0]
            
            # valid days used to find mean/std and indices used to find position
            first_valid_days, first_indices = find_consecutive_days(recorded_data, first=True)
            last_valid_days, last_indices = find_consecutive_days(recorded_data, first=False)

            if first_indices and last_indices and (last_indices[0] - first_indices[1] >= 0):
                days_distance.append(last_indices[0] - first_indices[1])
                results.setdefault('#DaysOfRecordedData', []).append(len(recorded_data))
                results.setdefault('First 5 Consecutive Days Position', []).append(first_indices)
                results.setdefault('First 5 Consecutive Days Mean', []).append(round(first_valid_days.mean(), 2))
                results.setdefault('First 5 Consecutive Days Std Dev', []).append(round(first_valid_days.std(), 2))
                results.setdefault('Last 5 Consecutive Days Position', []).append(last_indices)
                results.setdefault('Last 5 Consecutive Days Mean', []).append(round(last_valid_days.mean(), 2))
                results.setdefault('Last 5 Consecutive Days Std Dev', []).append(round(last_valid_days.std(), 2))
            else: 
                days_distance.append("Not calculable")
                results.setdefault('#DaysOfRecordedData', []).append("Not calculable")
                results.setdefault('First 5 Consecutive Days Position', []).append("Not calculable")
                results.setdefault('First 5 Consecutive Days Mean', []).append("Not sufficient")
                results.setdefault('First 5 Consecutive Days Std Dev', []).append("Not sufficient")
                results.setdefault('Last 5 Consecutive Days Position', []).append("Not calculable")
                results.setdefault('Last 5 Consecutive Days Mean', []).append("Not sufficient")
                results.setdefault('Last 5 Consecutive Days Std Dev', []).append("Not sufficient")

        # convert results dictionary to a DataFrame
        new_df = pd.DataFrame(results)
        new_df['Distance Between First and Last 5 Days'] = days_distance
        # remove rows with insufficient data
        new_df = new_df[(new_df['#DaysOfRecordedData'] != "Not calculable")]
        new_df = new_df.merge(tier_info, on='Participant ID', how='left')
        # move 'Intervention Group' column right after 'Participant ID'
        intervention_group = new_df.pop('Intervention Group')
        new_df.insert(1, 'Intervention Group', intervention_group)
        new_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # this is for adjusting column widths
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

print("New formatted analysis file created", output_file_path)
