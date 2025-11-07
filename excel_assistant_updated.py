"""
excel_assistant_updated.py

Merged Excel Analysis Assistant
- GUI (Tkinter) retained from your working version
- Expanded analysis & chatbot logic merged from your other file
- Chatbot-style commands (type naturally in the input area)
- 100% offline
"""

import sys
IN_NOTEBOOK = 'ipykernel' in sys.modules

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')  # Use Tkinter backend for matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.stats import zscore
from pathlib import Path
import threading
import warnings
from datetime import datetime
import io
import os
import tempfile

warnings.filterwarnings('ignore')
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (10, 6)


# ============================================================================
# EXCEL FORMULA GENERATOR (small helper)
# ============================================================================
class ExcelFormulaGenerator:
    """Generates Excel formulas (simple helpers)"""

    def __init__(self):
        pass

    def average_formula(self, column_name, start_row=2, end_row=100):
        return f"=AVERAGE({column_name}{start_row}:{column_name}{end_row})"


# ============================================================================
# SIMPLE EXPORT ENGINE
# ============================================================================
class SimpleExportEngine:
    """Export data to various formats"""

    def __init__(self):
        self.export_history = []

    def export_to_html(self, content, filename="report.html", title="Analysis Report"):
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>{title}</title>
<style>
body {{font-family: Arial; max-width: 900px; margin: 40px auto; padding: 20px; background: #f5f5f5;}}
.header {{background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; text-align: center;}}
.content {{background: white; padding: 30px; margin-top: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);}}
</style>
</head>
<body>
<div class="header"><h1>{title}</h1><p>Generated: {timestamp}</p></div>
<div class="content">{content.replace(chr(10), '<br>')}</div>
</body>
</html>"""

            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html)

            self.export_history.append({'type': 'HTML', 'filename': filename, 'timestamp': timestamp})
            return f"✅ HTML saved: {filename}"
        except Exception as e:
            return f"❌ Export error: {str(e)}"

    def export_to_excel(self, df, filename="export.xlsx"):
        try:
            df.to_excel(filename, index=False)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.export_history.append({'type': 'Excel', 'filename': filename, 'timestamp': timestamp})
            return f"✅ Excel saved: {filename}"
        except Exception as e:
            return f"❌ Export error: {str(e)}"

    def export_to_csv(self, df, filename="export.csv"):
        try:
            df.to_csv(filename, index=False)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.export_history.append({'type': 'CSV', 'filename': filename, 'timestamp': timestamp})
            return f"✅ CSV saved: {filename}"
        except Exception as e:
            return f"❌ Export error: {str(e)}"


# ============================================================================
# BASIC EXCEL ANALYZER (with extra methods)
# ============================================================================
class BasicExcelAnalyzer:
    """Basic Excel analysis engine"""

    def __init__(self):
        self.df = None
        self.df_filtered = None
        self.file_name = None
        self.formula_gen = ExcelFormulaGenerator()

    def load_excel(self, file_path):
        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path)
            else:
                # allow bytes-like objects or file paths
                self.df = pd.read_excel(file_path)

            self.df_filtered = self.df.copy()
            self.file_name = Path(file_path).name

            response = f"✅ Loaded: {self.file_name}\n\n"
            response += f"📊 Rows: {len(self.df)} | Columns: {len(self.df.columns)}\n"
            response += f"📋 Columns: {', '.join(self.df.columns.astype(str).tolist())}\n\n"
            response += "Ready to analyze! Try 'Show summary statistics'"
            return response
        except Exception as e:
            return f"❌ Error: {str(e)}"

    def get_summary_statistics(self):
        if self.df is None:
            return "❌ Load a file first"

        response_lines = ["📊 SUMMARY STATISTICS\n"]
        numeric_cols = self.df_filtered.select_dtypes(include=[np.number]).columns.tolist()

        if not numeric_cols:
            return "❌ No numeric columns found."

        for col in numeric_cols:
            series = self.df_filtered[col].dropna()
            response_lines.append(f"📈 {col}:")
            try:
                response_lines.append(f"  Mean: {series.mean():.4f}")
                response_lines.append(f"  Median: {series.median():.4f}")
                response_lines.append(f"  Std: {series.std():.4f}")
                response_lines.append(f"  Min: {series.min():.4f}")
                response_lines.append(f"  Max: {series.max():.4f}")
                response_lines.append(f"  Excel formula example: {self.formula_gen.average_formula(col)}")
            except Exception:
                response_lines.append("  (Could not compute numeric stats)")

            response_lines.append("")

        return "\n".join(response_lines)

    def calculate_correlations(self):
        if self.df is None:
            return "❌ Load a file first"

        numeric_cols = self.df_filtered.select_dtypes(include=[np.number]).columns.tolist()

        if len(numeric_cols) < 2:
            return "❌ Need at least 2 numeric columns"

        corr = self.df_filtered[numeric_cols].corr()

        response = "🔗 CORRELATIONS\n\n"
        response += corr.to_string() + "\n"

        return response

    def preview_data(self, rows=10):
        if self.df is None:
            return "❌ Load a file first"

        response = f"👀 DATA PREVIEW ({rows} rows)\n\n"
        response += self.df_filtered.head(rows).to_string(index=False)

        return response

    def check_missing_values(self):
        if self.df is None:
            return "❌ Load a file first"
        missing = self.df.isnull().sum()
        total = self.df.shape[0]
        lines = ["🔎 MISSING VALUES\n"]
        for col, m in missing.items():
            if m > 0:
                lines.append(f"  {col}: {m} missing ({m/total:.2%})")
        if len(lines) == 1:
            return "✅ No missing values detected."
        return "\n".join(lines)

    def percentiles(self, column, percentiles=[25, 50, 75]):
        if self.df is None:
            return "❌ Load a file first"
        if column not in self.df.columns:
            return f"❌ Column '{column}' not found"
        s = self.df[column].dropna()
        if s.empty:
            return f"❌ Column '{column}' has no data"
        p = np.percentile(s, percentiles)
        lines = [f"📊 Percentiles for {column}:"]
        for perc, val in zip(percentiles, p):
            lines.append(f"  {perc}th: {val}")
        return "\n".join(lines)

    def find_outliers_zscore(self, column, threshold=3.0):
        if self.df is None:
            return "❌ Load a file first"
        if column not in self.df.columns:
            return f"❌ Column '{column}' not found"
        s = pd.to_numeric(self.df[column], errors='coerce').dropna()
        if s.empty:
            return f"❌ Column '{column}' has no numeric data"
        zs = zscore(s)
        outlier_indices = np.where(np.abs(zs) > threshold)[0]
        outliers = s.iloc[outlier_indices]
        if outliers.empty:
            return f"✅ No outliers found in {column} using z-score > {threshold}"
        lines = [f"⚠️ Outliers in {column} (z>|{threshold}|):"]
        for idx, val in zip(outliers.index, outliers.values):
            lines.append(f"  Index {idx}: {val}")
        return "\n".join(lines)

    def calculate_z_scores(self, column):
        if self.df is None:
            return "❌ Load a file first"
        if column not in self.df.columns:
            return f"❌ Column '{column}' not found"
        s = pd.to_numeric(self.df[column], errors='coerce')
        zs = (s - s.mean()) / s.std()
        return zs

    def create_histogram(self, column, bins=20, show=True):
        if self.df is None:
            return "❌ Load a file first"
        if column not in self.df.columns:
            return f"❌ Column '{column}' not found"
        series = pd.to_numeric(self.df[column], errors='coerce').dropna()
        if series.empty:
            return f"❌ Column '{column}' has no numeric data"
        plt.figure()
        plt.hist(series, bins=bins)
        plt.title(f"Histogram: {column}")
        plt.xlabel(column)
        plt.ylabel("Frequency")
        if show:
            plt.show()
        return f"✅ Histogram for {column} displayed."

    def create_bar_chart(self, column, top_n=20, show=True):
        if self.df is None:
            return "❌ Load a file first"
        if column not in self.df.columns:
            return f"❌ Column '{column}' not found"
        counts = self.df[column].value_counts().nlargest(top_n)
        plt.figure()
        counts.plot(kind='bar')
        plt.title(f"Bar chart: {column}")
        plt.xlabel(column)
        plt.ylabel("Count")
        if show:
            plt.show()
        return f"✅ Bar chart for {column} displayed."

    def create_scatter(self, x_col, y_col, show=True):
        if self.df is None:
            return "❌ Load a file first"
        if x_col not in self.df.columns or y_col not in self.df.columns:
            return f"❌ One of the columns '{x_col}' or '{y_col}' not found"
        x = pd.to_numeric(self.df[x_col], errors='coerce')
        y = pd.to_numeric(self.df[y_col], errors='coerce')
        df2 = pd.concat([x, y], axis=1).dropna()
        if df2.empty:
            return f"❌ No overlapping numeric data between {x_col} and {y_col}"
        plt.figure()
        plt.scatter(df2.iloc[:, 0], df2.iloc[:, 1], alpha=0.6)
        plt.title(f"Scatter: {x_col} vs {y_col}")
        plt.xlabel(x_col)
        plt.ylabel(y_col)
        if show:
            plt.show()
        return f"✅ Scatter plot for {x_col} vs {y_col} displayed."


# ============================================================================
# SIMPLE CHATBOT (connects commands to analyzer & export engine)
# ============================================================================
class SimpleChatBot:
    """Simple chatbot for Excel analysis"""

    def __init__(self):
        self.analyzer = BasicExcelAnalyzer()
        self.export_engine = SimpleExportEngine()
        self.last_response = ""
        self.last_df_snapshot = None

    def process_message(self, user_message):
        msg = user_message.strip()
        msg_lower = msg.lower()

        # Help
        if 'help' in msg_lower:
            return self.get_help()

        # Load file command (if user typed a path)
        if msg_lower.startswith('load '):
            path = msg[5:].strip().strip('"').strip("'")
            return self.analyzer.load_excel(path)

        # Preview
        if 'preview' in msg_lower or 'show data' in msg_lower:
            response = self.analyzer.preview_data()
            self.last_response = response
            return response

        # Summary / statistics
        if 'summary' in msg_lower or 'statistics' in msg_lower:
            response = self.analyzer.get_summary_statistics()
            self.last_response = response
            return response

        # Correlation
        if 'correl' in msg_lower:
            response = self.analyzer.calculate_correlations()
            self.last_response = response
            return response

        # Missing values
        if 'missing' in msg_lower:
            response = self.analyzer.check_missing_values()
            self.last_response = response
            return response

        # Percentiles - expect "percentiles for <column>"
        if 'percentile' in msg_lower or 'percentiles' in msg_lower:
            # parse column name
            parts = msg_lower.replace('percentiles for', '').replace('percentiles', '').replace('percentile for', '').strip()
            col = parts.strip()
            if not col:
                return "Usage: 'Percentiles for <column_name>'"
            # use original column name casing if possible
            col_real = self._find_column_name(col)
            return self.analyzer.percentiles(col_real)

        # Outliers
        if 'outlier' in msg_lower:
            # "find outliers in <column>"
            col = self._extract_column_after(msg_lower, ['outliers in', 'outliers for', 'find outliers in', 'find outliers for'])
            if not col:
                return "Usage: 'Find outliers in <column>'"
            col_real = self._find_column_name(col)
            return self.analyzer.find_outliers_zscore(col_real)

        # Z-scores
        if 'z-score' in msg_lower or 'zscore' in msg_lower or 'z scores' in msg_lower:
            col = self._extract_column_after(msg_lower, ['z-score for', 'z scores for', 'zscore for', 'z score for'])
            if not col:
                return "Usage: 'Calculate z-scores for <column>'"
            col_real = self._find_column_name(col)
            zs = self.analyzer.calculate_z_scores(col_real)
            # show a short preview
            return f"Z-scores for {col_real} (first 10):\n{zs.head(10).to_string()}"

        # Charts: histogram, bar chart, scatter
        if 'histogram' in msg_lower or 'hist' in msg_lower:
            col = self._extract_column_after(msg_lower, ['histogram of', 'histogram for', 'histogram', 'hist of', 'show histogram of'])
            if not col:
                return "Usage: 'Show histogram of <column>'"
            col_real = self._find_column_name(col)
            return self.analyzer.create_histogram(col_real)

        if 'bar chart' in msg_lower or 'bar' in msg_lower:
            col = self._extract_column_after(msg_lower, ['bar chart for', 'create bar chart for', 'bar chart', 'bar of'])
            if not col:
                return "Usage: 'Create bar chart for <column>'"
            col_real = self._find_column_name(col)
            return self.analyzer.create_bar_chart(col_real)

        if 'scatter' in msg_lower:
            # expect "scatter plot x vs y"
            # find "scatter plot" then parse x and y
            if 'vs' in msg_lower:
                # e.g., "scatter plot age vs salary" or "scatter plot age vs salary"
                parts = msg_lower.split('scatter')[-1]
                if 'vs' in parts:
                    left, right = parts.split('vs', 1)
                    x = left.replace('plot', '').replace('plot of', '').strip()
                    y = right.strip()
                    if not x or not y:
                        return "Usage: 'Scatter plot <x> vs <y>'"
                    x_real = self._find_column_name(x)
                    y_real = self._find_column_name(y)
                    return self.analyzer.create_scatter(x_real, y_real)
            return "Usage: 'Scatter plot <x> vs <y>'"

        # Export commands
        if 'export' in msg_lower:
            if 'html' in msg_lower:
                return self.export_engine.export_to_html(self.last_response or "No content", filename=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
            if 'excel' in msg_lower or 'xlsx' in msg_lower:
                if self.analyzer.df_filtered is None:
                    return "❌ No data to export"
                return self.export_engine.export_to_excel(self.analyzer.df_filtered, filename=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            if 'csv' in msg_lower:
                if self.analyzer.df_filtered is None:
                    return "❌ No data to export"
                return self.export_engine.export_to_csv(self.analyzer.df_filtered, filename=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            return "Usage: 'Export to HTML/Excel/CSV'"

        # Preview top rows for a column: "show top 10 of <col>"
        if 'top' in msg_lower and 'of' in msg_lower:
            try:
                # e.g., "top 10 of Salary"
                import re
                m = re.search(r'top\s+(\d+)\s+of\s+(.+)', msg_lower)
                if m:
                    n = int(m.group(1))
                    col = m.group(2).strip()
                    col_real = self._find_column_name(col)
                    if col_real not in self.analyzer.df.columns:
                        return f"❌ Column {col_real} not found"
                    preview = self.analyzer.df[col_real].head(n).to_string()
                    return f"Top {n} of {col_real}:\n{preview}"
            except Exception:
                pass

        # Fallback suggestions
        return "🤔 Try: 'Show summary statistics', 'Calculate correlations', 'Preview data', 'Check for missing values', 'Percentiles for <col>', 'Find outliers in <col>', 'Show histogram of <col>', 'Export to Excel/CSV/HTML', or 'Help'"

    def get_help(self):
        help_text = """
📚 AVAILABLE COMMANDS

📊 Analysis:
  • Show summary statistics
  • Calculate correlations
  • Preview the data

🔎 Data checks:
  • Check for missing values
  • Percentiles for <column>
  • Find outliers in <column>
  • Calculate z-scores for <column>

📈 Charts:
  • Create bar chart for <column>
  • Show histogram of <column>
  • Scatter plot <x> vs <y>

💾 Export:
  • Export to HTML
  • Export to Excel
  • Export to CSV

💡 Just type naturally!
"""
        return help_text

    # Helper parse utilities
    def _find_column_name(self, partial):
        # match case-insensitive partial column name
        if self.analyzer.df is None:
            return partial
        cols = list(self.analyzer.df.columns)
        p = partial.strip().lower()
        # exact match
        for c in cols:
            if c.lower() == p:
                return c
        # substring match
        matches = [c for c in cols if p in c.lower()]
        if matches:
            return matches[0]
        # fallback use original partial
        return partial

    def _extract_column_after(self, text, prefixes):
        for pref in prefixes:
            if pref in text:
                return text.split(pref, 1)[1].strip()
        return None


# ============================================================================
# AUTOCOMPLETE WIDGET (same as original but slightly robust)
# ============================================================================
class AutocompleteEntry(tk.Frame):
    """Text widget with dropdown suggestions"""

    def __init__(self, parent, suggestions, **kwargs):
        super().__init__(parent)

        self.suggestions = suggestions
        self.filtered_suggestions = []

        self.text = tk.Text(self, height=3, wrap=tk.WORD, **kwargs)
        self.text.pack(fill=tk.BOTH, expand=True)

        self.dropdown = tk.Listbox(
            self, height=8, relief=tk.FLAT,
            bg="white", font=("Arial", 9)
        )
        self.dropdown.pack_forget()

        self.text.bind("<KeyRelease>", self.on_keyrelease)
        self.text.bind("<Down>", self.on_down_arrow)
        self.text.bind("<Up>", self.on_up_arrow)
        self.text.bind("<Return>", self.on_return)
        self.text.bind("<Escape>", self.hide_dropdown)
        self.dropdown.bind("<Double-Button-1>", self.on_select)

    def on_keyrelease(self, event):
        if event.keysym in ['Down', 'Up', 'Return', 'Escape', 'Shift_L', 'Shift_R']:
            return

        current_text = self.text.get("1.0", tk.END).strip().lower()

        if not current_text:
            self.hide_dropdown()
            return

        self.filtered_suggestions = [
            s for s in self.suggestions
            if current_text in s.lower()
        ][:10]

        if self.filtered_suggestions:
            self.show_dropdown()
        else:
            self.hide_dropdown()

    def show_dropdown(self):
        self.dropdown.delete(0, tk.END)
        for suggestion in self.filtered_suggestions:
            self.dropdown.insert(tk.END, suggestion)
        if self.dropdown.size() > 0:
            self.dropdown.selection_set(0)
            self.dropdown.pack(fill=tk.X, pady=(5, 0))

    def hide_dropdown(self, event=None):
        self.dropdown.pack_forget()

    def on_down_arrow(self, event):
        if self.dropdown.winfo_ismapped():
            current = self.dropdown.curselection()
            if current:
                next_index = min(current[0] + 1, self.dropdown.size() - 1)
                self.dropdown.selection_clear(0, tk.END)
                self.dropdown.selection_set(next_index)
                self.dropdown.see(next_index)
            return "break"

    def on_up_arrow(self, event):
        if self.dropdown.winfo_ismapped():
            current = self.dropdown.curselection()
            if current:
                prev_index = max(current[0] - 1, 0)
                self.dropdown.selection_clear(0, tk.END)
                self.dropdown.selection_set(prev_index)
                self.dropdown.see(prev_index)
            return "break"

    def on_return(self, event):
        if self.dropdown.winfo_ismapped():
            self.on_select(event)
            return "break"

    def on_select(self, event):
        selection = self.dropdown.curselection()
        if selection:
            selected_text = self.dropdown.get(selection[0])
            self.text.delete("1.0", tk.END)
            self.text.insert("1.0", selected_text)
            self.hide_dropdown()

    def get_text(self):
        return self.text.get("1.0", tk.END).strip()

    def clear_text(self):
        self.text.delete("1.0", tk.END)
        self.hide_dropdown()


# ============================================================================
# HELP PANEL (searchable help)
# ============================================================================
class HelpPanel(tk.Frame):
    """Built-in help panel with searchable commands"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.setup_help_ui()

    def setup_help_ui(self):
        title = tk.Label(self, text="📚 Quick Guide", font=("Arial", 14, "bold"), bg="#f0f0f0")
        title.pack(pady=10)

        self.help_text = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Arial", 9), bg="white", relief=tk.FLAT, padx=10, pady=10)
        self.help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        help_content = """
📊 BASIC COMMANDS:

- Show summary statistics
  Get mean, median, std, min, max for all numeric columns

- Calculate correlations
  See relationships between numeric columns

- Preview the data
  View first 10 rows of your data

- Check for missing values
  Detect incomplete data

- Percentiles for <column>
  Show percentile distribution

- Find outliers in <column>
  Detect unusual values using z-score

- Calculate z-scores for <column>
  Standardized scores

- Charts:
  Create bar chart for <column>
  Show histogram of <column>
  Scatter plot <x> vs <y>

💾 EXPORT:

- Export to HTML
  Create a beautiful report (open in browser)

- Export to Excel
  Save data with formatting

- Export to CSV
  Universal data format

💡 TIPS:

- Just type naturally!
- Start typing to see suggestions
- Press Enter to select
- Click 'Load File' to begin

🔒 100% OFFLINE - Your data stays private!
"""
        self.help_text.insert("1.0", help_content)
        self.help_text.config(state=tk.DISABLED)


# ============================================================================
# MAIN GUI (kept visually as your working one)
# ============================================================================
class ExcelAssistantGUI:
    """Main application GUI"""

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Analysis Assistant")
        self.root.geometry("1200x700")

        self.command_suggestions = [
            "Show summary statistics",
            "Calculate correlations",
            "Preview the data",
            "Check for missing values",
            "Percentiles for Salary",
            "Find outliers in Salary",
            "Calculate z-scores for Salary",
            "Create bar chart for Department",
            "Show histogram of Salary",
            "Scatter plot Age vs Salary",
            "Export to HTML",
            "Export to Excel",
            "Export to CSV",
            "Help",
        ]

        self.bot = SimpleChatBot()
        self.file_loaded = False
        self.show_help_panel = False

        self.setup_ui()

    def setup_ui(self):
        bg_color = "#f0f0f0"
        header_color = "#667eea"

        self.root.configure(bg=bg_color)

        # HEADER
        header_frame = tk.Frame(self.root, bg=header_color, height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)

        tk.Label(header_frame, text="📊 Excel Analysis Assistant", font=("Arial", 18, "bold"), bg=header_color, fg="white").pack(side=tk.LEFT, padx=20, pady=20)
        tk.Label(header_frame, text="v1.0 | 100% Offline", font=("Arial", 9), bg=header_color, fg="white").pack(side=tk.RIGHT, padx=20)

        # TOOLBAR
        toolbar_frame = tk.Frame(self.root, bg=bg_color, height=60)
        toolbar_frame.pack(fill=tk.X, side=tk.TOP)
        toolbar_frame.pack_propagate(False)

        tk.Button(toolbar_frame, text="📂 Load File", command=self.load_file, font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", relief=tk.FLAT, padx=20, pady=8, cursor="hand2").pack(side=tk.LEFT, padx=10, pady=10)

        self.file_label = tk.Label(toolbar_frame, text="No file loaded", font=("Arial", 9), bg=bg_color, fg="#666")
        self.file_label.pack(side=tk.LEFT, padx=15)

        self.help_toggle_btn = tk.Button(toolbar_frame, text="📚 Show Help", command=self.toggle_help, font=("Arial", 10), bg="#2196F3", fg="white", relief=tk.FLAT, padx=15, pady=8, cursor="hand2")
        self.help_toggle_btn.pack(side=tk.RIGHT, padx=10, pady=10)

        # MAIN CONTENT
        main_container = tk.Frame(self.root, bg=bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.chat_frame = tk.Frame(main_container, bg=bg_color)
        self.chat_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.chat_display = scrolledtext.ScrolledText(self.chat_frame, wrap=tk.WORD, font=("Arial", 10), bg="white", relief=tk.FLAT, padx=15, pady=15)
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.config(state=tk.DISABLED)

        self.chat_display.tag_config("user", foreground="#1976D2", font=("Arial", 10, "bold"))
        self.chat_display.tag_config("bot", foreground="#388E3C", font=("Arial", 10, "bold"))

        self.help_panel = HelpPanel(main_container, bg="#f0f0f0", width=350)

        # INPUT
        input_container = tk.Frame(self.root, bg=bg_color, height=100)
        input_container.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=15)
        input_container.pack_propagate(False)

        tk.Label(input_container, text="💡 Start typing for suggestions | ↓ to navigate | Enter to send", font=("Arial", 8), bg=bg_color, fg="#666").pack(anchor=tk.W, pady=(0, 5))

        input_frame = tk.Frame(input_container, bg=bg_color)
        input_frame.pack(fill=tk.BOTH, expand=True)

        self.input_text = AutocompleteEntry(input_frame, suggestions=self.command_suggestions, font=("Arial", 10), relief=tk.FLAT, bg="white", padx=10, pady=10)
        self.input_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        self.input_text.text.bind("<Return>", self.on_enter)

        tk.Button(input_frame, text="Send ➤", command=self.send_message, font=("Arial", 12, "bold"), bg=header_color, fg="white", relief=tk.FLAT, width=10, cursor="hand2").pack(side=tk.RIGHT, fill=tk.Y)

        self.add_bot_message("👋 Welcome!\n\nClick 'Load File' to start analyzing!")

    def toggle_help(self):
        self.show_help_panel = not self.show_help_panel
        if self.show_help_panel:
            self.help_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
            self.help_toggle_btn.config(text="📚 Hide Help")
        else:
            self.help_panel.pack_forget()
            self.help_toggle_btn.config(text="📚 Show Help")

    def load_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")])

        if file_path:
            response = self.bot.analyzer.load_excel(file_path)
            self.file_loaded = True
            self.file_label.config(text=f"📁 {Path(file_path).name}", fg="#4CAF50")
            self.add_bot_message(response)

    def send_message(self):
        user_input = self.input_text.get_text()
        if not user_input:
            return

        self.input_text.clear_text()

        if not self.file_loaded and user_input.lower() not in ['help']:
            self.add_bot_message("❌ Load a file first!")
            return

        self.add_user_message(user_input)
        threading.Thread(target=self.process, args=(user_input,), daemon=True).start()

    def process(self, user_input):
        response = self.bot.process_message(user_input)
        # Save last response to analyzer object (for exports)
        try:
            self.bot.last_response = response
        except Exception:
            pass
        self.root.after(0, self.add_bot_message, response)

    def on_enter(self, event):
        if self.input_text.dropdown.winfo_ismapped():
            return
        if event.state & 0x1:
            return
        self.send_message()
        return "break"

    def add_user_message(self, message):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\n👤 You:\n", "user")
        self.chat_display.insert(tk.END, f"{message}\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)

    def add_bot_message(self, message):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\n🤖 Assistant:\n", "bot")
        self.chat_display.insert(tk.END, f"{message}\n")
        self.chat_display.insert(tk.END, "\n" + "─"*70 + "\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)


# ============================================================================
# RUN
# ============================================================================
def main():
    root = tk.Tk()
    app = ExcelAssistantGUI(root)
    root.mainloop()


if __name__ == "__main__":
    if not IN_NOTEBOOK:
        main()
    else:
        print("💡 Running inside Jupyter Notebook — GUI disabled.")
        print("👉 You can still import and test your logic functions here.")
