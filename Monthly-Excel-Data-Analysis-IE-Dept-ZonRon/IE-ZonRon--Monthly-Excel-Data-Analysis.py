import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter
import re
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings("ignore")

# ==========================================
# 1. HELPER FUNCTIONS FOR PLOTTING
# ==========================================

def add_value_labels(ax, spacing=5, format_str="{:.0f}"):
    """Add labels to the end of each bar or on top of each dot."""
    # For Line Plots
    for line in ax.lines:
        x_data = line.get_xdata()
        y_data = line.get_ydata()
        for i, (x, y) in enumerate(zip(x_data, y_data)):
            if pd.notna(y) and y > 0:
                try:
                    label = format_str.format(y)
                    ax.annotate(label, 
                                (x, y), 
                                textcoords="offset points", 
                                xytext=(0, spacing), 
                                ha='center', 
                                fontsize=8, 
                                fontweight='bold')
                except:
                    pass

    # For Bar Plots
    for container in ax.containers:
        try:
            ax.bar_label(container, fmt=format_str, padding=3, fontsize=7, rotation=90)
        except:
            pass

def add_horizontal_bar_labels(ax):
    """Specific function for horizontal bar charts."""
    for container in ax.containers:
        ax.bar_label(container, fmt='%.0f', padding=3, fontsize=8, fontweight='bold')

# ==========================================
# 2. DATA CLEANING FUNCTIONS
# ==========================================

def find_and_clean_detailed_sheet(xls):
    """
    Dynamically finds the detailed monthly sheet (e.g., 'November', 'December')
    by looking for specific columns like 'Style'.
    """
    try:
        found_sheet_name = None
        df_raw = pd.DataFrame()

        # Strategy 1: check the first sheet (User requirement)
        first_sheet = xls.sheet_names[0]
        df_test = pd.read_excel(xls, first_sheet, header=None, nrows=20)
        
        # Check if it looks like the detailed sheet (contains 'Style')
        if df_test.astype(str).apply(lambda x: x.str.contains('Style', case=False)).any().any():
            found_sheet_name = first_sheet
        else:
            # Strategy 2: Search other sheets if first one isn't it
            for sheet in xls.sheet_names:
                df_test = pd.read_excel(xls, sheet, header=None, nrows=20)
                if df_test.astype(str).apply(lambda x: x.str.contains('Style', case=False)).any().any():
                    found_sheet_name = sheet
                    break
        
        if not found_sheet_name:
            print("Could not identify the Detailed/Monthly sheet (missing 'Style' column).")
            return pd.DataFrame()

        print(f"  > Identified Detailed Sheet: '{found_sheet_name}'")
        df_raw = pd.read_excel(xls, found_sheet_name, header=None)
        
        # Now Parse it dynamically
        date_row_idx = -1
        metric_row_idx = -1
        style_col_idx = -1
        
        # Locate rows
        for i, row in df_raw.head(20).iterrows():
            row_str = row.astype(str)
            
            # Find Date Row (looks for YYYY-MM-DD)
            if row_str.str.contains(r'202\d-', regex=True).any():
                date_row_idx = i
            
            # Find Metric Row (looks for Production/Output and Target)
            # The metric row is usually below the date row
            if i > date_row_idx and date_row_idx != -1:
                 if row_str.str.contains('Output', case=False).any() or row_str.str.contains('Prod', case=False).any():
                     metric_row_idx = i
            
            # Find Style Column
            for col_idx, val in enumerate(row):
                if 'Style' in str(val):
                    style_col_idx = col_idx

        if date_row_idx == -1 or metric_row_idx == -1 or style_col_idx == -1:
            # Fallback defaults if detection fails but sheet was identified
            if date_row_idx == -1: date_row_idx = 4
            if metric_row_idx == -1: metric_row_idx = 5
            if style_col_idx == -1: style_col_idx = 4 # Common index

        # Extract Static Data (Style, Buyer, etc.)
        # We assume Style is the key. We'll grab the column at style_col_idx
        data_start_row = metric_row_idx + 1
        
        # Get Style Column
        styles = df_raw.iloc[data_start_row:, style_col_idx].astype(str)
        # Filter out junk
        valid_indices = styles[~styles.str.contains("Total", case=False, na=False) & (styles != 'nan')].index
        
        static_df = df_raw.iloc[valid_indices].copy()
        
        # Extract Date and Production Data
        date_row = df_raw.iloc[date_row_idx]
        metric_row = df_raw.iloc[metric_row_idx]
        
        long_format_data = []
        current_date = None
        
        # Iterate through columns to find 'Production' or 'Output' metrics associated with a Date
        for col_idx in range(df_raw.shape[1]):
            # Update current date if this column has a date header
            val_date = date_row.iloc[col_idx]
            if pd.notna(val_date) and str(val_date).strip() not in ['nan', '']:
                current_date = val_date
            
            val_metric = str(metric_row.iloc[col_idx]).lower()
            
            # Check for Production/Output keyword
            # We want 'output', 'prod', 'pcs' -- but NOT 'target' or 'eff'
            is_prod_col = ('output' in val_metric or 'prod' in val_metric or 'pcs' in val_metric)
            is_not_target = 'target' not in val_metric
            
            if current_date and is_prod_col and is_not_target:
                # Extract values for the valid rows
                prod_values = df_raw.iloc[valid_indices, col_idx]
                
                temp_df = pd.DataFrame({
                    'Style': df_raw.iloc[valid_indices, style_col_idx],
                    'Date': current_date,
                    'Production': pd.to_numeric(prod_values, errors='coerce').fillna(0)
                })
                long_format_data.append(temp_df)
                
            # Also extract Efficiency for other charts if needed (optional, keeping your existing logic)
            if current_date and 'eff' in val_metric and '%' in val_metric:
                eff_values = df_raw.iloc[valid_indices, col_idx]
                temp_df = pd.DataFrame({
                    'Style': df_raw.iloc[valid_indices, style_col_idx],
                    'Date': current_date,
                    'Efficiency': pd.to_numeric(eff_values, errors='coerce') * 100
                })
                long_format_data.append(temp_df)

        if not long_format_data:
            return pd.DataFrame()
            
        final_df = pd.concat(long_format_data, ignore_index=True)
        
        # Separate Production and Efficiency if mixed in long format, or just keep as is with NaNs
        # Better to have specific columns. Let's pivot or split.
        # Simple approach: Return the long df. It will have cols 'Style', 'Date', 'Production', 'Efficiency'
        # The concat might result in separate rows for Prod and Eff. Let's merge them.
        
        # Group by Style and Date to merge metrics
        final_df = final_df.groupby(['Style', 'Date'], as_index=False).first()
        
        print(f"  > Detailed Sheet Processed: {len(final_df)} records.")
        return final_df

    except Exception as e:
        print(f"Error processing Detailed Sheet: {e}")
        return pd.DataFrame()

def clean_date_wise(df):
    """Cleans the Date Wise Monthly Summary DataFrame."""
    try:
        header_row_idx = -1
        for i, row in df.iterrows():
            if row.notna().sum() > 2:
                row_str = row.astype(str).str.lower()
                if row_str.str.contains('date').any() and (row_str.str.contains('output').any() or row_str.str.contains('target').any()):
                    header_row_idx = i
                    break

        if header_row_idx != -1:
            df.columns = df.iloc[header_row_idx]
            df = df.iloc[header_row_idx+1:]
        
        df.columns = df.columns.astype(str).str.strip()
        
        rename_map = {}
        for col in df.columns:
            c_low = col.lower()
            if 'date' in c_low: rename_map[col] = 'Date'
            elif 'man' in c_low and 'power' in c_low: rename_map[col] = 'Man Power'
            elif 'output' in c_low: rename_map[col] = 'Output'
            elif 'working' in c_low and 'minutes' in c_low: rename_map[col] = 'Working Minutes'
            elif 'achieve' in c_low and 'minutes' in c_low: rename_map[col] = 'Achieve Minutes'
            elif 'target' in c_low: rename_map[col] = 'Target'
            elif 'eff' in c_low and '%' in c_low: rename_map[col] = 'Eff(%)'
        
        df = df.rename(columns=rename_map)
        df = df.dropna(subset=['Date'])
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        
        numeric_cols = ['Target', 'Output', 'Eff(%)', 'Man Power', 'Working Minutes', 'Achieve Minutes']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        if 'Man Power' in df.columns and 'Working Minutes' in df.columns and 'Output' in df.columns:
            df['Man_Hours'] = df['Man Power'] * (df['Working Minutes'] / 60)
            df['Productivity_PCS_MH'] = df['Output'] / df['Man_Hours']

        print(f"  > Date Wise Summary Loaded: {len(df)} rows.")
        return df
    except Exception as e:
        print(f"Error cleaning Date Wise Summary: {e}")
        return pd.DataFrame()

def clean_supervisor_summary(df):
    """Cleans the Summary Print or Supervisor Wise Summary DataFrame."""
    try:
        header_found = False
        for i, row in df.iterrows():
            if row.notna().sum() > 3:
                row_str = row.astype(str).str.lower()
                if row_str.str.contains('line').any() and row_str.str.contains('supervisor').any():
                    cols = row.astype(str).tolist()
                    seen = {}
                    new_cols = []
                    for c in cols:
                        c_clean = c.strip()
                        if c_clean in seen:
                            seen[c_clean] += 1
                            new_cols.append(f"{c_clean}.{seen[c_clean]}")
                        else:
                            seen[c_clean] = 0
                            new_cols.append(c_clean)

                    df.columns = new_cols
                    df = df.iloc[i+1:]
                    header_found = True
                    break

        if not header_found and 'Line no' in df.columns:
            header_found = True

        sup_col = next((c for c in df.columns if 'Supervisor' in str(c)), None)
        if sup_col: df = df.rename(columns={sup_col: 'Name of Supervisor'})

        line_col = next((c for c in df.columns if 'Line' in str(c) and 'no' in str(c).lower()), None)
        if line_col: df = df.rename(columns={line_col: 'Line no'})
        
        if 'Line no' in df.columns:
             df = df.dropna(subset=['Line no'])
             df = df[df['Line no'].astype(str).str.lower() != 'total']
             df['Line no'] = df['Line no'].astype(str).str.replace('Line ', 'Line-', regex=False)
             df['Line no'] = df['Line no'].astype(str).str.replace('Line', 'Line-', regex=False)
             df['Line no'] = df['Line no'].astype(str).str.replace('--', '-', regex=False)

        eff_col = next((c for c in df.columns if 'Eff' in str(c) and '%' in str(c)), None)
        if eff_col:
            df['Efficiency_Clean'] = pd.to_numeric(df[eff_col], errors='coerce') * 100

        print(f"  > Supervisor Summary Loaded: {len(df)} rows.")
        return df
    except Exception as e:
        print(f"Error cleaning Supervisor Summary: {e}")
        return pd.DataFrame()

def process_daily_supervisor_trend(df_raw):
    """Parses 'Line & Supervisor Wise Summary' for DAILY Efficiency."""
    try:
        date_row_idx = -1
        metric_row_idx = -1

        for i, row in df_raw.head(30).iterrows():
            row_str = row.astype(str)
            date_matches = row_str.str.contains(r'202\d-\d{2}-\d{2}', regex=True).sum()
            if date_matches > 1:
                date_row_idx = i
                break
        
        if date_row_idx == -1: return pd.DataFrame()

        metric_row_idx = date_row_idx + 1
        date_row_vals = df_raw.iloc[date_row_idx].values
        metric_row_vals = df_raw.iloc[metric_row_idx].astype(str).str.strip()
        
        col_date_map = {}
        current_date = None
        
        for col_idx, val in enumerate(date_row_vals):
            val_str = str(val)
            if '202' in val_str and '-' in val_str:
                 current_date = val
            if current_date is not None:
                col_date_map[col_idx] = current_date
                
        sup_col_idx = -1
        search_limit = 10
        for r_idx in range(max(0, date_row_idx - 2), date_row_idx + 3):
            row_vals = df_raw.iloc[r_idx, :search_limit].astype(str).str.lower()
            for c_idx, val in enumerate(row_vals):
                if 'supervisor' in val or 'name' in val:
                    sup_col_idx = c_idx
                    break
            if sup_col_idx != -1: break
            
        if sup_col_idx == -1: sup_col_idx = 1
        
        data_start_row = metric_row_idx + 1
        daily_sup_data = []
        supervisors = df_raw.iloc[data_start_row:, sup_col_idx]
        
        for col_idx in range(df_raw.shape[1]):
            if col_idx < len(metric_row_vals):
                metric_val = metric_row_vals.iloc[col_idx]
                if 'eff' in metric_val.lower() and '%' in metric_val:
                    if col_idx in col_date_map:
                        date_val = col_date_map[col_idx]
                        col_values = df_raw.iloc[data_start_row:, col_idx]
                        temp_df = pd.DataFrame({
                            'Supervisor': supervisors,
                            'Date': date_val,
                            'Efficiency': pd.to_numeric(col_values, errors='coerce') * 100
                        })
                        daily_sup_data.append(temp_df)
        
        if not daily_sup_data: return pd.DataFrame()
        
        final_df = pd.concat(daily_sup_data, ignore_index=True)
        final_df = final_df.dropna(subset=['Efficiency'])
        final_df = final_df[final_df['Supervisor'].notna()]
        final_df = final_df[~final_df['Supervisor'].astype(str).str.lower().isin(['total', 'nan', ''])]
        
        print(f"  > Daily Supervisor Trend Data Loaded: {len(final_df)} rows.")
        return final_df

    except Exception as e:
        print(f"Error processing Daily Supervisor Trend: {e}")
        return pd.DataFrame()

def clean_daily_summary(df):
    try:
        if 'Remarks' not in df.columns:
            for i, row in df.iterrows():
                if row.notna().sum() > 2:
                    if row.astype(str).str.contains('Remarks').any():
                        df.columns = df.iloc[i]
                        df = df.iloc[i+1:]
                        break
        print(f"  > Daily Summary Loaded: {len(df)} rows.")
        return df
    except Exception as e:
        print(f"Error cleaning Daily Summary: {e}")
        return pd.DataFrame()

# ==========================================
# 3. MAIN EXECUTION
# ==========================================

def main():
    root = tk.Tk()
    root.withdraw()

    print("Opening file dialog...")
    file_path = filedialog.askopenfilename(
        title="Select IE Department Excel File",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )

    if not file_path:
        print("No file selected. Exiting.")
        return

    print(f"Processing: {file_path}")

    # 2. Initialize DataFrames
    df_date = pd.DataFrame()
    df_sup = pd.DataFrame()
    df_detail = pd.DataFrame()
    df_daily_raw = pd.DataFrame()
    df_daily_sup_trend = pd.DataFrame()

    # 3. Read Excel
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        print(f"Sheets found: {sheet_names}")
        
        # --- 1. Load Detailed Sheet (The one with changing names like Nov, Dec) ---
        # Uses dynamic logic to find sheet with 'Style' and 'Production'
        df_detail = find_and_clean_detailed_sheet(xls)

        # --- 2. Load Other Sheets (Summary sheets with fixed keywords) ---
        date_sheet = next((s for s in sheet_names if "Date" in s and "Wise" in s), None)
        if date_sheet:
            df_date = clean_date_wise(pd.read_excel(xls, date_sheet, header=None))
            
        sup_sheet = next((s for s in sheet_names if "Summary Print" in s or "Supervisor" in s or "Sup." in s), None)
        if sup_sheet:
             df_sup = clean_supervisor_summary(pd.read_excel(xls, sup_sheet, header=None))
             
        line_sup_sheet = next((s for s in sheet_names if "Line" in s and "Supervisor" in s and "Summary" in s), None)
        if line_sup_sheet:
            df_daily_sup_trend = process_daily_supervisor_trend(pd.read_excel(xls, line_sup_sheet, header=None))
            
        daily_sheet = next((s for s in sheet_names if "Daily Summary" in s), None)
        if daily_sheet:
            df_daily_raw = clean_daily_summary(pd.read_excel(xls, daily_sheet, header=None))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")
        return

    # 4. Generate Visualizations
    image_buffer = []
    sns.set_theme(style="whitegrid")

    def save_plot_to_buffer(fig, title, source, category, description):
        img_data = io.BytesIO()
        fig.savefig(img_data, format='png', bbox_inches='tight', dpi=100)
        img_data.seek(0)
        image_buffer.append({
            'img': img_data, 'title': title, 'source': source, 'category': category, 'desc': description
        })
        plt.close(fig)

    print("Generating Graphs...")

    # =========================================================================
    # NEW GRAPH: Daily Production by Style (Bar Chart)
    # =========================================================================
    if not df_detail.empty and 'Production' in df_detail.columns:
        # Filter for valid production
        prod_data = df_detail[df_detail['Production'] > 0]
        
        if not prod_data.empty:
            # Sort by Date
            prod_data['Date_Str'] = prod_data['Date'].dt.strftime('%Y-%m-%d')
            prod_data = prod_data.sort_values('Date')
            
            # Since this chart can get very wide with many dates + styles, 
            # we make it wide and group by Date, hue by Style.
            
            # Calculate dynamic width: Number of unique dates * 1.5 inches
            n_dates = prod_data['Date_Str'].nunique()
            fig_width = max(15, n_dates * 0.8)
            
            fig, ax = plt.subplots(figsize=(fig_width, 8))
            
            sns.barplot(
                data=prod_data, 
                x='Date_Str', 
                y='Production', 
                hue='Style', 
                ax=ax, 
                palette='bright',
                edgecolor='black'
            )
            
            # Add labels
            add_value_labels(ax, format_str="{:.0f}")
            
            ax.set_title('Daily Production by Style (Comparison)', fontsize=16, fontweight='bold')
            ax.set_ylabel('Production (Pcs)', fontsize=12)
            ax.set_xlabel('Date', fontsize=12)
            plt.xticks(rotation=90)
            plt.legend(bbox_to_anchor=(1.01, 1), loc='upper left', title='Style')
            
            save_plot_to_buffer(fig, "Daily Production by Style", "Detailed Monthly Sheet", "Category 1: Production Analysis", "Bar chart showing Production Count for each Style on each Date.")
    
    # =========================================================================
    # EXISTING GRAPHS
    # =========================================================================

    if not df_date.empty:
        fig, ax = plt.subplots(figsize=(12, 6))
        sns.lineplot(data=df_date, x='Date', y='Eff(%)', marker='o', linewidth=2.5, color='green', ax=ax)
        add_value_labels(ax, spacing=10, format_str="{:.2f}%")
        ax.set_xticks(df_date['Date'])
        ax.set_xticklabels(df_date['Date'].dt.strftime('%Y-%m-%d'), rotation=90)
        ax.set_title('Monthly Efficiency Trend', fontsize=14, fontweight='bold')
        save_plot_to_buffer(fig, "Monthly Efficiency Trend", "Date Wise Monthly Summary", "Category 1: Executive Dashboard", "Shows the overall department performance pulse.")

    if not df_date.empty and 'Target' in df_date.columns and 'Output' in df_date.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        df_melt = df_date.melt(id_vars=['Date'], value_vars=['Target', 'Output'], var_name='Metric', value_name='Pieces')
        sns.barplot(data=df_melt, x='Date', y='Pieces', hue='Metric', palette={'Target': 'gray', 'Output': 'blue'}, ax=ax)
        add_value_labels(ax, format_str="{:.0f}")
        ax.set_xticks(range(len(df_date)))
        ax.set_xticklabels(df_date['Date'].dt.strftime('%Y-%m-%d'), rotation=90)
        ax.set_title('Daily Production vs Target', fontsize=14, fontweight='bold')
        save_plot_to_buffer(fig, "Daily Production vs Target", "Date Wise Monthly Summary", "Category 1: Executive Dashboard", "Highlights missed targets.")

    if not df_date.empty and 'Productivity_PCS_MH' in df_date.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        sns.lineplot(data=df_date, x='Date', y='Productivity_PCS_MH', marker='s', color='purple', linestyle='--', ax=ax)
        add_value_labels(ax, format_str="{:.2f}")
        ax.set_xticks(df_date['Date'])
        ax.set_xticklabels(df_date['Date'].dt.strftime('%Y-%m-%d'), rotation=90)
        ax.set_title('Labor Productivity Trend (Pcs per Man-Hour)', fontsize=14, fontweight='bold')
        save_plot_to_buffer(fig, "Labor Productivity Trend", "Date Wise Monthly Summary", "Category 1: Executive Dashboard", "Tracks labor cost-effectiveness.")

    req_cols = ['Man Power', 'Output', 'Working Minutes', 'Achieve Minutes']
    if not df_date.empty and all(col in df_date.columns for col in req_cols):
        fig_height = max(10, len(df_date) * 0.8)
        fig, ax = plt.subplots(figsize=(15, fig_height))
        df_plot = df_date.copy()
        df_plot['Date_Str'] = df_plot['Date'].dt.strftime('%Y-%m-%d')
        df_plot = df_plot.set_index('Date_Str')[req_cols]
        df_plot.plot(kind='barh', width=0.8, ax=ax, colormap='viridis', log=True)
        add_horizontal_bar_labels(ax)
        ax.grid(axis='x', linestyle='--', linewidth=0.7, alpha=0.7)
        ax.set_title('Daily Metrics Comparison (Log Scale)', fontsize=16, fontweight='bold')
        ax.set_xlabel('Value (Logarithmic Scale)', fontsize=12)
        ax.set_ylabel('Date', fontsize=12)
        ax.legend(title='Metrics', bbox_to_anchor=(1.0, 1), loc='upper left')
        save_plot_to_buffer(fig, "Daily Metrics Comparison", "Date Wise Monthly Summary", "Category 1: Executive Dashboard", "Comparison of Man Power, Output, and Minutes with Values.")

    if not df_date.empty and 'Man Power' in df_date.columns:
        fig, ax1 = plt.subplots(figsize=(12, 6))
        ax2 = ax1.twinx()
        df_date['Date_Str'] = df_date['Date'].dt.strftime('%Y-%m-%d')
        sns.barplot(data=df_date, x='Date_Str', y='Man Power', color='lightblue', alpha=0.6, ax=ax1, label='Manpower')
        sns.lineplot(data=df_date, x='Date_Str', y='Eff(%)', color='red', marker='o', ax=ax2, label='Efficiency %', linewidth=2, sort=False)
        ax1.set_ylabel('Manpower', color='blue')
        ax2.set_ylabel('Efficiency %', color='red')
        ax1.set_xticks(range(len(df_date)))
        ax1.set_xticklabels(df_date['Date_Str'], rotation=90)
        ax1.set_title('Manpower vs Efficiency Correlation', fontsize=14, fontweight='bold')
        save_plot_to_buffer(fig, "Manpower vs Efficiency", "Date Wise Monthly Summary", "Category 1: Executive Dashboard", "Checks if adding manpower correlates with higher efficiency.")

    if not df_detail.empty and 'Efficiency' in df_detail.columns:
        style_data = df_detail[df_detail['Style'].astype(str).str.len() > 1]
        style_eff = style_data.groupby('Style')['Efficiency'].mean().sort_values(ascending=False).reset_index()
        if not style_eff.empty:
            dynamic_height = max(6, len(style_eff) * 0.4)
            fig, ax = plt.subplots(figsize=(10, dynamic_height))
            sns.barplot(data=style_eff, y='Style', x='Efficiency', hue='Style', palette='magma', legend=False, ax=ax)
            add_value_labels(ax, format_str="{:.1f}%")
            ax.set_title('Efficiency by Style (All Styles)', fontsize=14, fontweight='bold')
            save_plot_to_buffer(fig, "Efficiency by Style", "Detailed Monthly Sheet", "Category 4: Strategic Analysis", "Identifies winning styles.")

    if not df_daily_sup_trend.empty:
        fig, ax = plt.subplots(figsize=(14, 7))
        sns.lineplot(data=df_daily_sup_trend, x='Date', y='Efficiency', hue='Supervisor', marker='o', linewidth=2, palette='tab10', ax=ax)
        unique_dates = sorted(df_daily_sup_trend['Date'].unique())
        formatted_dates = [pd.to_datetime(d).strftime('%Y-%m-%d') for d in unique_dates]
        ax.set_xticks(unique_dates)
        ax.set_xticklabels(formatted_dates, rotation=90)
        ax.set_title('Daily Efficiency Trend by Supervisor', fontsize=14, fontweight='bold')
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        save_plot_to_buffer(fig, "Daily Efficiency by Supervisor", "Line & Supervisor Wise Summary", "Category 2: Performance", "Compares how each supervisor performs day-by-day.")

    # 5. Save to Excel
    output_dir = os.path.dirname(file_path)
    output_filename = os.path.join(output_dir, 'IE_Department_Analysis_Report_Local.xlsx')
    
    print(f"Saving report to {output_filename}...")
    
    try:
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        workbook = writer.book

        if not df_date.empty: df_date.to_excel(writer, sheet_name='Cleaned_Date_Summary', index=False)
        if not df_sup.empty: df_sup.to_excel(writer, sheet_name='Cleaned_Sup_Summary', index=False)
        if not df_daily_sup_trend.empty: df_daily_sup_trend.to_excel(writer, sheet_name='Daily_Sup_Trend', index=False)
        if not df_detail.empty: df_detail.to_excel(writer, sheet_name='Cleaned_Detail_Summary', index=False)

        worksheet = workbook.add_worksheet('Graph and Chart Analysis')
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 50)
        worksheet.set_column('C:C', 30)

        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        cell_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        bold_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'font_size': 12})

        worksheet.write('A1', 'Category', header_fmt)
        worksheet.write('B1', 'Description & Insight', header_fmt)
        worksheet.write('C1', 'Data Source', header_fmt)
        worksheet.write('D1', 'Visualization', header_fmt)

        current_row = 1
        image_buffer.sort(key=lambda x: x['category'])

        for item in image_buffer:
            row_h = 400
            if "Style" in item['title'] or "Metrics Comparison" in item['title']:
                row_h = 600
            
            worksheet.set_row(current_row, row_h)
            
            worksheet.write(current_row, 0, item['category'], cell_fmt)
            worksheet.write(current_row, 1, f"{item['title']}\n\n{item['desc']}", bold_fmt)
            worksheet.write(current_row, 2, item['source'], cell_fmt)
            
            y_scale = 0.8
            if "Style" in item['title'] or "Metrics Comparison" in item['title']:
                y_scale = 0.6
                
            worksheet.insert_image(current_row, 3, item['title'], {'image_data': item['img'], 'x_scale': 0.8, 'y_scale': y_scale})
            current_row += 1

        writer.close()
        messagebox.showinfo("Success", f"Report generated successfully!\n\nSaved at: {output_filename}")

    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save Excel file:\n{str(e)}\n\nClose the file if it is open.")

if __name__ == "__main__":
    main()