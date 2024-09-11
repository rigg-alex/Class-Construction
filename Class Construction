import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side

# Configurable parameters
block_classes = {'Block A': 4, 'Block B': 4, 'Block C': 4}  # Number of classes in each block
max_ls_per_class = 5  # Maximum number of LS students per class
max_gat_per_class = 12  # Maximum number of GAT students per class
max_students_per_class = 30  # Maximum total number of students per class
block_column = 'Block'  # Column name for timetable constraints
factor_columns = ['Factor1', 'Factor2', 'Factor3']  # Example column names, replace with actual names

# Load your Excel file into a pandas DataFrame
# Define the new file path and sheet name
file_path = r'.2025\Class Construction 2025.xlsx'
sheet_name = 'XX_XXXX'

# Load the Excel file with the specified sheet into a pandas DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

def debug_count_students(df):
    print("\nDebug: Student Counts")
    print(f"Total students: {len(df)}")
    print(f"Unique values in {block_column} column: {df[block_column].unique()}")
    for block in df[block_column].unique():
        block_df = df[df[block_column] == block]
        ls_count = block_df['LS'].sum()
        gat_count = block_df['GAT'].sum()
        total_count = len(block_df)
        print(f"{block}: LS students = {ls_count}, GAT students = {gat_count}, Total students = {total_count}")

def assign_students(block_df, block, num_classes):
    assigned_df = block_df.copy()
    assigned_df['AssignedClass'] = np.nan

    # Assign LS students from last class to first
    ls_students = assigned_df[assigned_df['LS'] == 1].copy()
    for class_num in range(num_classes, 0, -1):
        class_name = f"{block}{class_num}"
        available_spots = max_ls_per_class
        students_to_assign = ls_students[ls_students['AssignedClass'].isna()].head(available_spots)
        assigned_df.loc[students_to_assign.index, 'AssignedClass'] = class_name
        ls_students.loc[students_to_assign.index, 'AssignedClass'] = class_name
        if ls_students['AssignedClass'].isna().sum() == 0:
            break

    # Assign GAT students from first class to last
    gat_students = assigned_df[assigned_df['GAT'] == 1].copy()
    for class_num in range(1, num_classes + 1):
        class_name = f"{block}{class_num}"
        current_class_size = assigned_df['AssignedClass'].value_counts().get(class_name, 0)
        available_spots = min(max_gat_per_class, max_students_per_class - current_class_size)
        students_to_assign = gat_students[gat_students['AssignedClass'].isna()].head(available_spots)
        assigned_df.loc[students_to_assign.index, 'AssignedClass'] = class_name
        gat_students.loc[students_to_assign.index, 'AssignedClass'] = class_name
        if gat_students['AssignedClass'].isna().sum() == 0:
            break

    return assigned_df

def balance_remaining_students(block_df, block, num_classes):
    assigned_df = block_df.copy()
    class_counts = assigned_df['AssignedClass'].value_counts()
    unassigned_students = assigned_df[assigned_df['AssignedClass'].isna()].copy()
    
    for _, student in unassigned_students.iterrows():
        valid_classes = [f"{block}{i}" for i in range(1, num_classes + 1)]
        target_class = min(valid_classes, key=lambda c: class_counts.get(c, 0) if class_counts.get(c, 0) < max_students_per_class else float('inf'))
        
        if class_counts.get(target_class, 0) < max_students_per_class:
            assigned_df.loc[student.name, 'AssignedClass'] = target_class
            class_counts[target_class] = class_counts.get(target_class, 0) + 1
        else:
            print(f"Warning: Unable to assign student {student.name} to a class in {block}. All classes are full.")
    
    return assigned_df

def process_block(block_df, block, num_classes):
    print(f"\nProcessing {block}")
    
    if num_classes == 0:
        print(f"Skipping {block} as it has no classes assigned.")
        return block_df
    
    assigned_df = assign_students(block_df, block, num_classes)
    assigned_df = balance_remaining_students(assigned_df, block, num_classes)
    
    return assigned_df

def create_class_table(df, class_name):
    class_df = df[df['AssignedClass'] == class_name].copy()
    class_df = class_df[['Student Code', 'Family Name', 'First Name', 'LS', 'GAT'] + factor_columns]
    
    # Calculate averages
    averages = class_df[factor_columns + ['LS', 'GAT']].mean()
    
    # Create a summary row
    summary = pd.DataFrame({
        'Student Code': ['Average'],
        'Family Name': [''],
        'First Name': [''],
        'LS': [averages['LS']],
        'GAT': [averages['GAT']]
    })
    for factor in factor_columns:
        summary[factor] = [averages[factor]]
    
    # Combine the class data with the summary
    result = pd.concat([class_df, summary], ignore_index=True)
    
    return result

def create_block_tables(df, block):
    block_df = df[df[block_column] == block].copy()
    classes = sorted(block_df['AssignedClass'].unique())
    
    block_tables = []
    for class_name in classes:
        class_table = create_class_table(block_df, class_name)
        block_tables.append(class_table)
    
    return block_tables

def write_block_to_sheet(writer, sheet_name, block_tables):
    workbook = writer.book
    sheet = workbook.create_sheet(sheet_name)
    
    current_row = 1
    for i, table in enumerate(block_tables):
        # Write class name
        sheet.cell(row=current_row, column=1, value=f"Class {i+1}")
        sheet.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1
        
        # Write table headers
        headers = table.columns.tolist()
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=current_row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(bottom=Side(style='thin'))
        current_row += 1
        
        # Write table data
        for _, row in table.iterrows():
            for col, value in enumerate(row, 1):
                sheet.cell(row=current_row, column=col, value=value)
            current_row += 1
        
        # Add an empty row between tables
        current_row += 1
    
    # Auto-adjust column widths
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width

def create_class_combination(df, combination_number):
    print(f"\nCreating Combination {combination_number}")
    debug_count_students(df)
    
    processed_blocks = []
    for block, num_classes in block_classes.items():
        block_df = df[df[block_column] == block]
        processed_block = process_block(block_df, block, num_classes)
        processed_blocks.append(processed_block)
    
    combined_df = pd.concat(processed_blocks, ignore_index=True)
    
    print("\nCalculating class metrics")
    class_metrics = combined_df.groupby('AssignedClass').agg(
        LS_count=('LS', 'sum'),
        GAT_count=('GAT', 'sum'),
        Total_count=('AssignedClass', 'count'),
        Factor1_mean=(factor_columns[0], 'mean'),
        Factor2_mean=(factor_columns[1], 'mean'),
        Factor3_mean=(factor_columns[2], 'mean')
    )
    
    print("\nClass Metrics:")
    print(class_metrics)
    
    unassigned = combined_df[combined_df['AssignedClass'].isna()]
    if not unassigned.empty:
        print(f"\nWarning: {len(unassigned)} students could not be assigned to a class.")
        print(unassigned)
    
    # Create block tables
    block_tables = {block: create_block_tables(combined_df, block) 
                    for block in block_classes.keys() if block_classes[block] > 0}
    
    return combined_df, class_metrics, block_tables

# Create two different combinations
try:
    combination1, metrics1, tables1 = create_class_combination(df.copy(), 1)
    combination2, metrics2, tables2 = create_class_combination(df.copy(), 2)

    # Save the results to an Excel file
    output_file_path = r'.\Class_Lists_Final.xlsx'
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        combination1.to_excel(writer, sheet_name='Combination1_Assignments', index=False)
        metrics1.to_excel(writer, sheet_name='Combination1_Metrics')
        combination2.to_excel(writer, sheet_name='Combination2_Assignments', index=False)
        metrics2.to_excel(writer, sheet_name='Combination2_Metrics')
        
        # Write consolidated block tables for each combination
        for i, tables in enumerate([tables1, tables2], 1):
            for block, block_tables in tables.items():
                sheet_name = f'Comb{i}_{block}'
                write_block_to_sheet(writer, sheet_name, block_tables)

    print("Class lists, summaries, and consolidated block tables have been saved to different sheets in the Excel file.")
except Exception as e:
    print(f"An error occurred: {str(e)}")
    import traceback
    traceback.print_exc()
    print("\nDataFrame information:")
    print(df.info())
    print("\nFirst few rows of the DataFrame:")
    print(df.head())
    print("\nUnique values in Block column:")
    print(df[block_column].unique())
