import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
import os
import sys
import json
import requests
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import xlsxwriter
import warnings
import numpy as np

# Suppress warnings for cleaner output
warnings.filterwarnings("ignore")

# ==========================================
# 1. HELPER FUNCTIONS FOR PLOTTING
# ==========================================

def add_value_labels(ax, spacing=5, format_str="{:.0f}", fontsize=9, color='black'):
    """Add labels to the end of each bar or on top of each dot."""
    # For Line Plots
    for line in ax.lines:
        x_data = line.get_xdata()
        y_data = line.get_ydata()
        for i, (x, y) in enumerate(zip(x_data, y_data)):
            if pd.notna(y) and (isinstance(y, (int, float)) and y > 0):
                try:
                    label = format_str.format(y)
                    ax.annotate(label, 
                                (x, y), 
                                textcoords="offset points", 
                                xytext=(0, spacing), 
                                ha='center', 
                                fontsize=fontsize, 
                                fontweight='bold',
                                color=color)
                except:
                    pass

    # For Bar Plots
    for container in ax.containers:
        try:
            ax.bar_label(container, fmt=format_str, padding=3, fontsize=fontsize, fontweight='bold', color=color)
        except:
            pass

def add_horizontal_bar_labels(ax):
    """Specific function for horizontal bar charts."""
    for container in ax.containers:
        try:
            ax.bar_label(container, fmt='%.0f', padding=4, fontsize=9, fontweight='bold')
        except:
            pass

# ==========================================
# 2. DATA PROCESSING FUNCTIONS
# ==========================================

def find_and_clean_detailed_sheet(xls):
    """Finds the monthly detail sheet (e.g. November, December) by scanning for 'Style'."""
    try:
        found_sheet_name = None
        # Scan sheets
        for sheet in xls.sheet_names:
            try:
                df_test = pd.read_excel(xls, sheet, header=None, nrows=25)
                # Convert to string and search for "Style"
                if df_test.astype(str).apply(lambda x: x.str.contains('Style', case=False)).any().any():
                    found_sheet_name = sheet
                    break
            except:
                continue
        
        if not found_sheet_name:
            return pd.DataFrame()

        df_raw = pd.read_excel(xls, found_sheet_name, header=None)
        
        date_row_idx = -1
        metric_row_idx = -1
        style_col_idx = -1
        
        # Locate headers
        for i, row in df_raw.head(30).iterrows():
            row_str = row.astype(str)
            # Find Date Row
            if row_str.str.contains(r'202\d-', regex=True).any():
                date_row_idx = i
            
            # Find Metric Row (Output/Prod)
            if i > date_row_idx and date_row_idx != -1:
                 if row_str.str.contains('Output', case=False).any() or \
                    row_str.str.contains('Prod', case=False).any() or \
                    row_str.str.contains('Qty', case=False).any():
                     metric_row_idx = i
            
            # Find Style Column
            for col_idx, val in enumerate(row):
                if 'Style' in str(val):
                    style_col_idx = col_idx

        # Fallbacks
        if date_row_idx == -1: date_row_idx = 4
        if metric_row_idx == -1: metric_row_idx = 5
        if style_col_idx == -1: style_col_idx = 4

        data_start_row = metric_row_idx + 1
        styles = df_raw.iloc[data_start_row:, style_col_idx].astype(str)
        valid_indices = styles[~styles.str.contains("Total", case=False, na=False) & (styles != 'nan')].index
        
        date_row = df_raw.iloc[date_row_idx]
        metric_row = df_raw.iloc[metric_row_idx]
        
        long_format_data = []
        
        for col_idx in range(df_raw.shape[1]):
            val_date = date_row.iloc[col_idx]
            # Check for valid date
            is_valid_date = False
            try:
                if pd.notna(val_date) and str(val_date).strip() != '':
                    pd.to_datetime(val_date)
                    is_valid_date = True
            except:
                pass
            
            if is_valid_date:
                val_metric = str(metric_row.iloc[col_idx]).lower()
                
                # Logic to capture PRODUCTION columns
                is_prod = ('output' in val_metric or 'prod' in val_metric or 'pcs' in val_metric)
                is_not_target = 'target' not in val_metric
                
                if is_prod and is_not_target:
                    prod_values = df_raw.iloc[valid_indices, col_idx]
                    temp_df = pd.DataFrame({
                        'Style': df_raw.iloc[valid_indices, style_col_idx],
                        'Date': val_date,
                        'Production': pd.to_numeric(prod_values, errors='coerce').fillna(0)
                    })
                    long_format_data.append(temp_df)
                
                # Logic to capture EFFICIENCY columns
                if 'eff' in val_metric and '%' in val_metric:
                    eff_values = df_raw.iloc[valid_indices, col_idx]
                    temp_df = pd.DataFrame({
                        'Style': df_raw.iloc[valid_indices, style_col_idx],
                        'Date': val_date,
                        'Efficiency': pd.to_numeric(eff_values, errors='coerce') * 100
                    })
                    long_format_data.append(temp_df)

        if not long_format_data:
            return pd.DataFrame()
            
        final_df = pd.concat(long_format_data, ignore_index=True)
        # Group by Style+Date to merge separate rows for Eff and Prod if they exist
        final_df = final_df.groupby(['Style', 'Date'], as_index=False).first()
        return final_df

    except Exception as e:
        print(f"Error parsing detail sheet: {e}")
        return pd.DataFrame()

def clean_date_wise(df):
    """Cleans the Date Wise Summary."""
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
        
        # Mapping for safety
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
        
        # Calculate Productivity
        if 'Man Power' in df.columns and 'Working Minutes' in df.columns and 'Output' in df.columns:
            df['Man_Hours'] = df['Man Power'] * (df['Working Minutes'] / 60)
            df['Productivity_PCS_MH'] = df.apply(lambda x: x['Output'] / x['Man_Hours'] if x['Man_Hours'] > 0 else 0, axis=1)

        return df
    except:
        return pd.DataFrame()

def clean_supervisor_summary(df):
    """Cleans Supervisor Summary."""
    try:
        header_found = False
        for i, row in df.iterrows():
            if row.notna().sum() > 3:
                row_str = row.astype(str).str.lower()
                if 'line' in str(row_str) and 'supervisor' in str(row_str):
                    df.columns = row.astype(str)
                    df = df.iloc[i+1:]
                    header_found = True
                    break
        if not header_found: return pd.DataFrame()

        # Rename
        for c in df.columns:
            if 'Supervisor' in str(c): df = df.rename(columns={c: 'Name of Supervisor'})
        return df
    except:
        return pd.DataFrame()

def process_daily_supervisor_trend(df_raw):
    """Gets Daily Efficiency per Supervisor."""
    try:
        date_row_idx = -1
        for i, row in df_raw.head(30).iterrows():
            if row.astype(str).str.contains(r'202\d-\d{2}-\d{2}', regex=True).sum() > 1:
                date_row_idx = i
                break
        if date_row_idx == -1: return pd.DataFrame()

        metric_row_idx = date_row_idx + 1
        date_row_vals = df_raw.iloc[date_row_idx].values
        metric_row_vals = df_raw.iloc[metric_row_idx].astype(str).str.strip()
        
        col_date_map = {}
        current_date = None
        for col_idx, val in enumerate(date_row_vals):
            if '202' in str(val) and '-' in str(val): current_date = val
            if current_date: col_date_map[col_idx] = current_date
        
        # Find Supervisor Col
        sup_col_idx = 1
        for r_idx in range(max(0, date_row_idx-5), date_row_idx+5):
            row_vals = df_raw.iloc[r_idx].astype(str).str.lower().values
            for c, val in enumerate(row_vals):
                if 'supervisor' in val: 
                    sup_col_idx = c
                    break

        data_start_row = metric_row_idx + 1
        daily_sup_data = []
        supervisors = df_raw.iloc[data_start_row:, sup_col_idx]
        
        for col_idx in range(df_raw.shape[1]):
            if col_idx < len(metric_row_vals):
                metric_val = metric_row_vals.iloc[col_idx]
                if 'eff' in metric_val.lower() and '%' in metric_val:
                    if col_idx in col_date_map:
                        col_values = df_raw.iloc[data_start_row:, col_idx]
                        daily_sup_data.append(pd.DataFrame({
                            'Supervisor': supervisors,
                            'Date': col_date_map[col_idx],
                            'Efficiency': pd.to_numeric(col_values, errors='coerce') * 100
                        }))
        
        if not daily_sup_data: return pd.DataFrame()
        final_df = pd.concat(daily_sup_data, ignore_index=True)
        return final_df.dropna(subset=['Efficiency'])
    except:
        return pd.DataFrame()

# ==========================================
# 3. GUI DASHBOARD
# ==========================================

class IEDashboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IE Department Analytics Dashboard | Sonia & Sweaters Ltd")
        self.root.geometry("850x700")
        self.root.configure(bg="#f0f2f5")

        # 1. User Logic
        self.username = self.load_user()
        if not self.username:
            self.username = simpledialog.askstring("Setup", "Welcome! Please enter your name:")
            if not self.username: self.username = "User"
            self.save_user(self.username)

        # 2. Layout
        self.create_ui()
        self.update_weather()
        self.update_time()

    def load_user(self):
        try:
            with open("config.json", "r") as f:
                return json.load(f).get("username", "")
        except:
            return ""

    def save_user(self, name):
        with open("config.json", "w") as f:
            json.dump({"username": name}, f)

    def get_weather(self):
        # Open-Meteo for Dhaka
        url = "https://api.open-meteo.com/v1/forecast?latitude=23.8103&longitude=90.4125&current_weather=true"
        try:
            r = requests.get(url, timeout=3)
            if r.status_code == 200:
                data = r.json()['current_weather']
                temp = data['temperature']
                code = data['weathercode']
                cond = "Sunny"
                if code > 3: cond = "Cloudy"
                if code > 50: cond = "Rainy"
                return f"{cond}, {temp}°C"
        except:
            return "N/A"
        return "Loading..."

    def update_weather(self):
        w = self.get_weather()
        self.weather_lbl.config(text=f"📍 Dhaka, BD: {w}")

    def update_time(self):
        now = datetime.datetime.now().strftime("%A, %d %B %Y | %I:%M:%S %p")
        self.date_lbl.config(text=now)
        self.root.after(1000, self.update_time)

    def create_ui(self):
        # Header
        header = tk.Frame(self.root, bg="#2c3e50", height=120)
        header.pack(fill=tk.X)
        
        tk.Label(header, text=f"Welcome, {self.username}", font=("Segoe UI", 24, "bold"), fg="white", bg="#2c3e50").pack(pady=(20, 5))
        self.date_lbl = tk.Label(header, text="", font=("Segoe UI", 12), fg="#bdc3c7", bg="#2c3e50")
        self.date_lbl.pack(pady=(0, 20))

        # Weather Bar
        self.weather_lbl = tk.Label(self.root, text="...", font=("Segoe UI", 11, "bold"), fg="#2c3e50", bg="#f0f2f5")
        self.weather_lbl.pack(pady=10)

        # Main Button Area
        card = tk.Frame(self.root, bg="white", relief=tk.RAISED, bd=1)
        card.pack(pady=20, padx=40, fill=tk.X, ipady=40)

        tk.Label(card, text="IE Monthly Report Generator", font=("Segoe UI", 18, "bold"), bg="white", fg="#2c3e50").pack(pady=15)
        
        self.btn = tk.Button(card, text="📂 Upload Excel & Generate Report", font=("Segoe UI", 12, "bold"), 
                             bg="#27ae60", fg="white", cursor="hand2", command=self.process_file, relief=tk.FLAT)
        self.btn.pack(pady=15, ipadx=20, ipady=10)
        
        tk.Label(card, text="Supports: Date Wise, Summary Print, Line Supervisor sheets", font=("Segoe UI", 9), fg="gray", bg="white").pack()

        # Professional Footer
        footer = tk.Frame(self.root, bg="#ecf0f1", height=140)
        footer.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Label(footer, text="For any sort of inconvenience with this application, please contact:", 
                 font=("Segoe UI", 10, "italic"), fg="#7f8c8d", bg="#ecf0f1").pack(pady=(15, 5))

        contact_info = (
            "Prottoy Saha\n"
            "Automation Engineer\n"
            "Sonia and Sweaters Limited\n"
            "+8801745547578  |  prottoy.saha@soniagroup.com"
        )
        tk.Label(footer, text=contact_info, font=("Segoe UI", 10, "bold"), fg="#2c3e50", bg="#ecf0f1", justify=tk.CENTER).pack(pady=(0, 15))

    def process_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path: return
        
        self.btn.config(text="⏳ Analyzing Data...", state=tk.DISABLED, bg="#95a5a6")
        self.root.update()
        
        try:
            self.run_analysis(path)
            messagebox.showinfo("Success", "Report Generated Successfully!\n\nSaved as 'IE_Report_Output.xlsx'")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            print(e)
        
        self.btn.config(text="📂 Upload Excel & Generate Report", state=tk.NORMAL, bg="#27ae60")

    def run_analysis(self, path):
        xls = pd.ExcelFile(path)
        
        # Load Data
        df_detail = find_and_clean_detailed_sheet(xls)
        df_date = pd.DataFrame()
        ds = next((s for s in xls.sheet_names if "Date" in s and "Wise" in s), None)
        if ds: df_date = clean_date_wise(pd.read_excel(xls, ds, header=None))
        
        df_sup_daily = pd.DataFrame()
        ls = next((s for s in xls.sheet_names if "Line" in s and "Sup" in s), None)
        if ls: df_sup_daily = process_daily_supervisor_trend(pd.read_excel(xls, ls, header=None))

        # Generate Visuals
        image_buffer = []
        # Use high contrast aesthetic
        sns.set_theme(style="whitegrid", palette="deep")

        def save_plot(fig, title, cat, desc):
            buf = io.BytesIO()
            fig.savefig(buf, format='png', bbox_inches='tight', dpi=150)
            buf.seek(0)
            image_buffer.append({'img': buf, 'title': title, 'cat': cat, 'desc': desc})
            plt.close(fig)

        # --- CHART 1: Monthly Efficiency Trend ---
        if not df_date.empty:
            fig, ax = plt.subplots(figsize=(12, 6))
            sns.lineplot(data=df_date, x='Date', y='Eff(%)', marker='o', markersize=8, color='#006400', linewidth=3, ax=ax)
            add_value_labels(ax, format_str="{:.2f}%")
            ax.set_title('Monthly Efficiency Trend', fontsize=16, fontweight='bold')
            ax.set_xticks(df_date['Date'])
            ax.set_xticklabels(df_date['Date'].dt.strftime('%d-%b'), rotation=90)
            save_plot(fig, "Monthly Efficiency Trend", "Executive Dashboard", "Overall department performance pulse.")

        # --- CHART 2: Daily Production vs Target ---
        if not df_date.empty and 'Target' in df_date.columns:
            fig, ax = plt.subplots(figsize=(14, 7))
            df_melt = df_date.melt(id_vars=['Date'], value_vars=['Target', 'Output'], var_name='Metric', value_name='Qty')
            sns.barplot(data=df_melt, x='Date', y='Qty', hue='Metric', 
                        palette={'Target': '#95a5a6', 'Output': '#2980b9'}, 
                        edgecolor='black', linewidth=1, ax=ax)
            add_value_labels(ax)
            ax.set_xticks(range(len(df_date)))
            ax.set_xticklabels(df_date['Date'].dt.strftime('%d-%b'), rotation=90)
            ax.set_title('Daily Production vs Target', fontsize=16, fontweight='bold')
            save_plot(fig, "Daily Production vs Target", "Executive Dashboard", "Highlights missed targets.")

        # --- CHART 3: Labor Productivity Trend ---
        if not df_date.empty and 'Productivity_PCS_MH' in df_date.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            sns.lineplot(data=df_date, x='Date', y='Productivity_PCS_MH', marker='s', markersize=9,
                         color='#8e44ad', linewidth=3, linestyle='--', ax=ax)
            add_value_labels(ax, format_str="{:.2f}")
            ax.set_title('Labor Productivity (Pcs per Man-Hour)', fontsize=16, fontweight='bold')
            ax.set_xticks(df_date['Date'])
            ax.set_xticklabels(df_date['Date'].dt.strftime('%d-%b'), rotation=90)
            save_plot(fig, "Labor Productivity Trend", "Executive Dashboard", "Tracks labor cost-effectiveness.")

        # --- CHART 4: Daily Metrics Comparison (IMPROVED COLORS & VISIBILITY) ---
        req = ['Man Power', 'Output', 'Working Minutes']
        if not df_date.empty and all(c in df_date.columns for c in req):
            h = max(10, len(df_date) * 0.8)
            fig, ax = plt.subplots(figsize=(16, h))
            df_plot = df_date.copy()
            df_plot['Date_Str'] = df_plot['Date'].dt.strftime('%d-%b')
            df_plot = df_plot.set_index('Date_Str')[req]
            
            # Use 'Set2' or custom colors for clear distinction
            custom_colors = ['#e74c3c', '#3498db', '#2ecc71'] # Red, Blue, Green
            df_plot.plot(kind='barh', width=0.85, ax=ax, color=custom_colors, edgecolor='black', log=True)
            
            add_horizontal_bar_labels(ax)
            ax.set_title('Daily Metrics Comparison (Log Scale)', fontsize=16, fontweight='bold')
            ax.set_xlabel('Value (Log Scale)', fontsize=12)
            ax.grid(axis='x', which='both', linestyle='--', alpha=0.5)
            save_plot(fig, "Daily Metrics Comparison", "Executive Dashboard", "Comparison of Man Power, Output, and Minutes with Values.")

        # --- CHART 5: Manpower vs Efficiency ---
        if not df_date.empty and 'Man Power' in df_date.columns:
            fig, ax1 = plt.subplots(figsize=(12, 6))
            ax2 = ax1.twinx()
            df_date['Date_Str'] = df_date['Date'].dt.strftime('%d-%b')
            
            sns.barplot(data=df_date, x='Date_Str', y='Man Power', color='#34495e', alpha=0.6, ax=ax1, edgecolor='black')
            sns.lineplot(data=df_date, x='Date_Str', y='Eff(%)', color='#e74c3c', marker='o', linewidth=3, ax=ax2)
            
            ax1.set_ylabel('Manpower', color='#34495e', fontsize=12, fontweight='bold')
            ax2.set_ylabel('Efficiency %', color='#e74c3c', fontsize=12, fontweight='bold')
            ax1.set_xticklabels(df_date['Date_Str'], rotation=90)
            ax1.set_title('Manpower vs Efficiency Correlation', fontsize=16, fontweight='bold')
            save_plot(fig, "Manpower vs Efficiency", "Executive Dashboard", "Checks if adding manpower correlates with higher efficiency.")

        # --- CHART 6: Daily Production by Style (High Contrast) ---
        if not df_detail.empty and 'Production' in df_detail.columns:
            df_p = df_detail[df_detail['Production'] > 0].copy()
            if not df_p.empty:
                df_p['Date_Str'] = df_p['Date'].dt.strftime('%Y-%m-%d')
                n_dates = df_p['Date_Str'].nunique()
                width = max(14, n_dates * 1.2)
                
                fig, ax = plt.subplots(figsize=(width, 8))
                # Use tab20 for max distinction between styles
                sns.barplot(data=df_p, x='Date_Str', y='Production', hue='Style', 
                            palette='tab20', edgecolor='black', linewidth=1, ax=ax)
                add_value_labels(ax)
                
                ax.set_title('Daily Production by Style', fontsize=16, fontweight='bold', pad=15)
                plt.xticks(rotation=45, ha='right')
                # Move legend out
                plt.legend(bbox_to_anchor=(1.01, 1), loc='upper left', title='Style Key')
                save_plot(fig, "Daily Production by Style", "Production Analysis", "Bar chart showing Production Count for each Style on each Date.")

        # --- CHART 7: Efficiency by Style ---
        if not df_detail.empty and 'Efficiency' in df_detail.columns:
            df_style_eff = df_detail.groupby('Style')['Efficiency'].mean().sort_values(ascending=False).reset_index()
            if not df_style_eff.empty:
                h = max(6, len(df_style_eff) * 0.5)
                fig, ax = plt.subplots(figsize=(10, h))
                sns.barplot(data=df_style_eff, y='Style', x='Efficiency', palette='viridis', edgecolor='black', ax=ax)
                add_horizontal_bar_labels(ax)
                ax.set_title('Efficiency by Style (Avg)', fontsize=16, fontweight='bold')
                save_plot(fig, "Efficiency by Style", "Strategic Analysis", "Identifies winning styles.")

        # Save Excel
        out_file = os.path.join(os.path.dirname(path), "IE_Report_Output.xlsx")
        writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
        wb = writer.book
        
        if not df_date.empty: df_date.to_excel(writer, sheet_name='Date_Summary', index=False)
        if not df_detail.empty: df_detail.to_excel(writer, sheet_name='Style_Data', index=False)

        ws = wb.add_worksheet('Graph and Chart Analysis') # Named as requested
        ws.set_column('A:A', 30)
        ws.set_column('B:B', 60)
        
        bold = wb.add_format({'bold': True, 'bg_color': '#2c3e50', 'font_color': 'white', 'border': 1})
        wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
        
        ws.write('A1', 'Category', bold)
        ws.write('B1', 'Description & Insight', bold)
        ws.write('C1', 'Visualization', bold)
        
        r = 1
        for img in image_buffer:
            h = 450
            if "Metrics" in img['title']: h = 600
            ws.set_row(r, h)
            ws.write(r, 0, img['cat'], wrap)
            ws.write(r, 1, f"{img['title']}\n\n{img['desc']}", wrap)
            
            y_s = 0.7
            if "Metrics" in img['title']: y_s = 0.55
            ws.insert_image(r, 2, img['title'], {'image_data': img['img'], 'x_scale': 0.7, 'y_scale': y_s})
            r += 1
            
        writer.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = IEDashboardApp(root)
    root.mainloop()