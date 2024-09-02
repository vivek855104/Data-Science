import pandas as pd
from tkinter import Tk, Label, Button, Checkbutton, IntVar, Toplevel, filedialog
from itertools import combinations

def select_columns_and_aggregation(df):
    def on_submit():
        selected_grouping_columns = [columns[i] for i, var in enumerate(group_var_list) if var.get() == 1]
        selected_aggregation_columns = [columns[i] for i, var in enumerate(agg_var_list) if var.get() == 1]
        
        if not selected_grouping_columns:
            print("No columns selected for grouping. Exiting.")
            top.destroy()
            return
        
        if not selected_aggregation_columns:
            print("No columns selected for aggregation. Exiting.")
            top.destroy()
            return
        
        top.destroy()
        generate_reports(df, selected_grouping_columns, selected_aggregation_columns)

    top = Toplevel()
    top.title("Select Columns to Group By and Aggregate")

    Label(top, text="Select columns to group by:").pack()
    
    # Grouping columns selection
    group_var_list = []
    for col in columns:
        var = IntVar()
        Checkbutton(top, text=col, variable=var).pack(anchor='w')
        group_var_list.append(var)
    
    Label(top, text="Select columns to aggregate:").pack()
    
    # Aggregation columns selection
    agg_var_list = []
    for col in columns:
        var = IntVar()
        Checkbutton(top, text=col, variable=var).pack(anchor='w')
        agg_var_list.append(var)

    Button(top, text="Submit", command=on_submit).pack()

    top.mainloop()

def generate_reports(df, grouping_columns, aggregation_columns):
    # Convert relevant columns to numeric, handling errors and missing data
    for col in aggregation_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Define the aggregation dictionary for summing up the columns
    agg_dict = {col: 'sum' for col in aggregation_columns}

    # Open save dialog for the output file
    output_file = filedialog.asksaveasfilename(title="Save the output Excel file",
                                               defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx *.xls")])
    if not output_file:
        print("No save location selected. Exiting.")
        return

    # Create a Pandas Excel writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Generate reports for different combinations of grouping columns
        for i in range(1, len(grouping_columns) + 1):
            for combo in combinations(grouping_columns, i):
                grouped_df = df.groupby(list(combo), as_index=False).agg(agg_dict)
                
                # Calculate additional metrics based on available columns
                if 'spend' in aggregation_columns and 'impressions' in aggregation_columns:
                    grouped_df['CPM'] = (grouped_df['spend'] / grouped_df['impressions']) * 1000
                if 'spend' in aggregation_columns and 'results' in aggregation_columns:
                    grouped_df['CPR'] = grouped_df['spend'] / grouped_df['results']
                if 'spend' in aggregation_columns and 'clicks_link' in aggregation_columns:
                    grouped_df['CPC_link'] = grouped_df['spend'] / grouped_df['clicks_link']
                if 'spend' in aggregation_columns and 'clicks_all' in aggregation_columns:
                    grouped_df['CPC_all'] = grouped_df['spend'] / grouped_df['clicks_all']
                if 'clicks_link' in aggregation_columns and 'impressions' in aggregation_columns:
                    grouped_df['CTR'] = (grouped_df['clicks_link'] / grouped_df['impressions']) * 100

                # Write each report to a separate sheet
                sheet_name = '_'.join(combo) if combo else 'Aggregated'
                grouped_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    
    print(f'Reports generated and saved to {output_file}')

def generate_report():
    # Initialize Tkinter and hide the root window
    root = Tk()
    root.withdraw()

    # Open file dialog to select the input file
    input_file = filedialog.askopenfilename(title="Select the input Excel file",
                                            filetypes=[("Excel files", "*.xlsx *.xls")])
    if not input_file:
        print("No file selected. Exiting.")
        return

    # Load the Excel file
    df = pd.read_excel(input_file)

    global columns
    columns = df.columns.tolist()

    # Call the column selection window
    select_columns_and_aggregation(df)

# Call the function to generate the report
generate_report()
