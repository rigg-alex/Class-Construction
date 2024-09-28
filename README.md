# Class Construction Script

This Python script automates the process of assigning students to classes, considering various factors such as learning support (LS) and gifted and talented (GAT) designations.

## Features

- Assigns students to classes based on predefined criteria
- Balances class composition considering LS and GAT students
- Creates multiple class combinations for comparison
- Generates detailed metrics and summaries for each class and combination
- Outputs results to an Excel file with multiple sheets for easy analysis

## Requirements

- Python 3.x
- pandas
- numpy
- openpyxl

You can install the required packages using pip:

```
pip install pandas numpy openpyxl
```

## Usage

1. Ensure your input Excel file is properly formatted with the following columns:
   - Student Code
   - Family Name
   - First Name
   - LS (Learning Support)
   - GAT (Gifted and Talented)
   - Block (Timetable constraints)
   - Factor1, Factor2, Factor3 (Additional factors for consideration)

2. Update the following variables in the script:
   - `file_path`: Path to your input Excel file
   - `sheet_name`: Name of the sheet containing student data
   - `block_classes`: Dictionary defining the number of classes in each block
   - `max_ls_per_class`: Maximum number of LS students per class
   - `max_gat_per_class`: Maximum number of GAT students per class
   - `max_students_per_class`: Maximum total number of students per class
   - `block_column`: Column name for timetable constraints
   - `factor_columns`: List of column names for additional factors

3. Run the script:

```
python hsc_class_assignment.py
```

4. The script will generate an output Excel file named `Class_Lists_Final.xlsx` containing:
   - Assignments for two different combinations
   - Metrics for each combination
   - Consolidated block tables for each combination

## Key Components

1. `assign_students()`: Assigns LS and GAT students to classes
2. `balance_remaining_students()`: Distributes remaining students across classes
3. `process_block()`: Processes student assignments for each block
4. `create_class_table()`: Generates a summary table for each class
5. `write_block_to_sheet()`: Writes block tables to the Excel output
6. `create_class_combination()`: Creates a complete class assignment combination

## Output

The script generates an Excel file with the following sheets:
- Combination1_Assignments and Combination2_Assignments: Full student assignments
- Combination1_Metrics and Combination2_Metrics: Class metrics for each combination
- Comb1_[Block] and Comb2_[Block]: Consolidated tables for each block in both combinations

## Troubleshooting

If you encounter any errors, the script will print detailed information about the DataFrame and any exceptions that occur. Use this information to debug issues with data formatting or script configuration.

## Customization

You can modify the script to add additional criteria for student assignment or adjust the balancing algorithms to suit your specific needs.

