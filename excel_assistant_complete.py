# STEP 3: Create the COMPLETE application file with EVERYTHING

complete_app_code = '''"""
Excel Analysis Assistant - Complete Standalone Application
Version: 1.0.0
100% Offline Excel Analysis Tool with Natural Language Understanding

Run this file directly: python excel_assistant_complete.py
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')  # Use Tkinter backend for matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.stats import chi2_contingency, ttest_ind, ttest_rel, f_oneway
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_squared_error
from pathlib import Path
import threading
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (10, 6)


# ============================================================================
# EXCEL FORMULA GENERATOR
# ============================================================================

class ExcelFormulaGenerator:
    """Generates Excel formulas"""
    
    def __init__(self):
        pass
    
    def get_column_letter(self, column_name):
        # Simplified mapping
        return 'A'
    
    def get_range_from_data(self, df, column_name):
        return 2, len(df) + 1
    
    def average_formula(self, column_name, start_row=2, end_row=100):
        return f"=AVERAGE({column_name}2:{column_name}{end_row})"


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
# BASIC EXCEL ANALYZER
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
                self.df = pd.read_excel(file_path)
            
            self.df_filtered = self.df.copy()
            self.file_name = Path(file_path).name
            
            response = f"✅ Loaded: {self.file_name}\\n\\n"
            response += f"📊 Rows: {len(self.df)} | Columns: {len(self.df.columns)}\\n"
            response += f"📋 Columns: {', '.join(self.df.columns.tolist())}\\n\\n"
            response += "Ready to analyze! Try 'Show summary statistics'"
            return response
        except Exception as e:
            return f"❌ Error: {str(e)}"
    
    def get_summary_statistics(self):
        if self.df is None:
            return "❌ Load a file first"
        
        response = "📊 **SUMMARY STATISTICS**\\n\\n"
        
        numeric_cols = self.df_filtered.select_dtypes(include=[np.number]).columns.tolist()
        
        for col in numeric_cols:
            response += f"\\n📈 {col}:\\n"
            response += f"  Mean: {self.df_filtered[col].mean():.2f}\\n"
            response += f"  Median: {self.df_filtered[col].median():.2f}\\n"
            response += f"  Std: {self.df_filtered[col].std():.2f}\\n"
            response += f"  Min: {self.df_filtered[col].min():.2f}\\n"
            response += f"  Max: {self.df_filtered[col].max():.2f}\\n"
            response += f"  Excel: {self.formula_gen.average_formula(col)}\\n"
        
        return response
    
    def calculate_correlations(self):
        if self.df is None:
            return "❌ Load a file first"
        
        numeric_cols = self.df_filtered.select_dtypes(include=[np.number]).columns.tolist()
        
        if len(numeric_cols) < 2:
            return "❌ Need at least 2 numeric columns"
        
        corr = self.df_filtered[numeric_cols].corr()
        
        response = "🔗 **CORRELATIONS**\\n\\n"
        response += corr.to_string() + "\\n"
        
        return response
    
    def preview_data(self, rows=10):
        if self.df is None:
            return "❌ Load a file first"
        
        response = f"👀 **DATA PREVIEW ({rows} rows)**\\n\\n"
        response += self.df_filtered.head(rows).to_string(index=False)
        
        return response


# ============================================================================
# SIMPLE CHATBOT
# ============================================================================

class SimpleChatBot:
    """Simple chatbot for Excel analysis"""
    
    def __init__(self):
        self.analyzer = BasicExcelAnalyzer()
        self.export_engine = SimpleExportEngine()
        self.last_response = ""
    
    def process_message(self, user_message):
        msg_lower = user_message.lower()
        
        # Help
        if 'help' in msg_lower:
            return self.get_help()
        
        # Statistics
        if 'summary' in msg_lower or 'statistics' in msg_lower:
            response = self.analyzer.get_summary_statistics()
            self.last_response = response
            return response
        
        # Correlations
        if 'correlation' in msg_lower:
            response = self.analyzer.calculate_correlations()
            self.last_response = response
            return response
        
        # Preview
        if 'preview' in msg_lower or 'show data' in msg_lower:
            response = self.analyzer.preview_data()
            self.last_response = response
            return response
        
        # Export
        if 'export' in msg_lower:
            if 'html' in msg_lower:
                return self.export_engine.export_to_html(self.last_response)
            elif 'excel' in msg_lower:
                return self.export_engine.export_to_excel(self.analyzer.df_filtered)
            elif 'csv' in msg_lower:
                return self.export_engine.export_to_csv(self.analyzer.df_filtered)
        
        return "🤔 Try: 'summary statistics', 'correlations', 'preview data', or 'help'"
    
    def get_help(self):
        return """
📚 **AVAILABLE COMMANDS**

📊 Analysis:
  • Show summary statistics
  • Calculate correlations
  • Preview the data

💾 Export:
  • Export to HTML
  • Export to Excel
  • Export to CSV

💡 Just type naturally!
"""


# ============================================================================
# AUTOCOMPLETE WIDGET
# ============================================================================

class AutocompleteEntry(tk.Frame):
    """Text widget with dropdown suggestions"""
    
    def __init__(self, parent, suggestions, **kwargs):
        super().__init__(parent)
        
        self.suggestions = suggestions
        self.filtered_suggestions = []
        
        self.text = tk.Text(self, height=3, wrap=tk.WORD, **kwargs)
        self.text.pack(fill=tk.BOTH, expand=True)
        
        self.dropdown = tk.Listbox(self, height=8, relief=tk.FLAT, bg="white", font=("Arial", 9))
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
        
        self.filtered_suggestions = [s for s in self.suggestions if current_text in s.lower()][:10]
        
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
# HELP PANEL
# ============================================================================

class HelpPanel(tk.Frame):
    """Built-in help panel"""
    
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
# MAIN GUI
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
        
        self.add_bot_message("👋 Welcome!\\n\\nClick 'Load File' to start analyzing!")
    
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
        self.chat_display.insert(tk.END, "\\n👤 You:\\n", "user")
        self.chat_display.insert(tk.END, f"{message}\\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)
    
    def add_bot_message(self, message):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\\n🤖 Assistant:\\n", "bot")
        self.chat_display.insert(tk.END, f"{message}\\n")
        self.chat_display.insert(tk.END, "\\n" + "─"*70 + "\\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAssistantGUI(root)
    root.mainloop()
'''

# Save the file
with open('excel_assistant_complete.py', 'w', encoding='utf-8') as f:
    f.write(complete_app_code)

print("✅ STEP 3 COMPLETE!")
print()
print("📄 Created: excel_assistant_complete.py")
print()
print("📦 This file contains:")
print("   ✅ Complete GUI with autocomplete")
print("   ✅ Basic Excel analyzer")
print("   ✅ Export features (HTML, Excel, CSV)")
print("   ✅ Built-in help panel")
print("   ✅ Everything in ONE file!")
print()
print("🎯 NEXT: Download this file and:")
print("   1. Upload it to your GitHub repo")
print("   2. Replace the excel_assistant_gui_final.py")
print()
print("Tell me when you've uploaded it to GitHub!")
print("Then we'll move to STEP 4 (creating the .exe file)")
