"""
Data Analyser Bloop - Version 1.0.2
-----------------------------------
AI-powered offline Excel analysis assistant
Developed by Soumili Panja (2025)
Built with Tkinter, Pandas, Matplotlib, Seaborn, Scikit-learn
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from pathlib import Path
import threading
import warnings

warnings.filterwarnings("ignore")
sns.set_style("whitegrid")
plt.rcParams["figure.figsize"] = (10, 6)


# =====================================================================
# Excel Analysis Engine
# =====================================================================
class DataAnalyser:
    def __init__(self):
        self.df = None
        self.file_name = None

    def load_excel(self, path):
        try:
            if path.endswith(".csv"):
                self.df = pd.read_csv(path)
            else:
                self.df = pd.read_excel(path)

            self.file_name = Path(path).name
            return f"✅ Loaded {self.file_name}\n📊 Rows: {len(self.df)} | Columns: {len(self.df.columns)}\n📋 Columns: {', '.join(self.df.columns)}"
        except Exception as e:
            return f"❌ Error loading file: {str(e)}"

    def summary(self):
        if self.df is None:
            return "❌ Load a file first."
        desc = self.df.describe(include="all").transpose().fillna("-")
        return f"📊 SUMMARY STATISTICS\n\n{desc.to_string()}"

    def correlations(self):
        if self.df is None:
            return "❌ Load a file first."
        num_cols = self.df.select_dtypes(include=np.number)
        if num_cols.empty:
            return "❌ No numeric columns found."
        corr = num_cols.corr()
        return f"🔗 CORRELATION MATRIX\n\n{corr.to_string()}"

    def preview(self, rows=10):
        if self.df is None:
            return "❌ Load a file first."
        return f"👀 DATA PREVIEW (first {rows} rows)\n\n{self.df.head(rows).to_string(index=False)}"

    def missing_values_report(self):
        if self.df is None:
            return "❌ Load a file first."

        missing = self.df.isnull().sum()
        total_missing = missing.sum()
        if total_missing == 0:
            return "✅ No missing values detected!"

        report = "🔎 MISSING VALUE ANALYSIS\n\n"
        for col, val in missing.items():
            if val > 0:
                pct = (val / len(self.df)) * 100
                suggestion = self._suggest_imputation(col)
                report += f"• {col}: {val} missing ({pct:.2f}%) → Suggest: {suggestion}\n"
        report += "\n💡 Type 'Impute missing values' to see example fix code."
        return report

    def _suggest_imputation(self, col):
        dtype = self.df[col].dtype
        if np.issubdtype(dtype, np.number):
            return "Fill with column median"
        elif dtype == 'object':
            return "Fill with mode (most frequent value)"
        else:
            return "Fill using forward-fill (ffill)"

    def simulate_imputation(self):
        if self.df is None:
            return "❌ Load a file first."
        missing = self.df.isnull().sum()
        if missing.sum() == 0:
            return "✅ No missing values found. Nothing to impute!"

        code_snippets = ["# Suggested imputation code examples:\n"]
        for col, val in missing.items():
            if val > 0:
                dtype = self.df[col].dtype
                if np.issubdtype(dtype, np.number):
                    code_snippets.append(f"df['{col}'].fillna(df['{col}'].median(), inplace=True)  # numeric")
                elif dtype == 'object':
                    code_snippets.append(f"df['{col}'].fillna(df['{col}'].mode()[0], inplace=True)  # categorical")
                else:
                    code_snippets.append(f"df['{col}'].fillna(method='ffill', inplace=True)  # other type")
        code_snippets.append("\n# ⚠️ This is simulation only – no data was changed.")
        return "\n".join(code_snippets)

    def show_imputation_code(self):
        return """📘 SAMPLE PYTHON IMPUTATION CODE:

import pandas as pd
df = pd.read_excel("data.xlsx")

# Fill numeric columns with median
df['Salary'].fillna(df['Salary'].median(), inplace=True)

# Fill categorical columns with mode
df['Department'].fillna(df['Department'].mode()[0], inplace=True)

# Forward fill for sequential data
df['Date'].fillna(method='ffill', inplace=True)
"""

    def create_histogram(self, col):
        if self.df is None:
            return "❌ Load a file first."
        if col not in self.df.columns:
            return f"❌ Column '{col}' not found."
        plt.figure()
        sns.histplot(self.df[col].dropna(), kde=True)
        plt.title(f"Histogram of {col}")
        plt.show()
        return f"✅ Displayed histogram for {col}"

    def scatter_plot(self, x, y):
        if self.df is None:
            return "❌ Load a file first."
        if x not in self.df.columns or y not in self.df.columns:
            return f"❌ One of the columns '{x}' or '{y}' not found."
        plt.figure()
        sns.scatterplot(x=self.df[x], y=self.df[y])
        plt.title(f"Scatter plot: {x} vs {y}")
        plt.show()
        return f"✅ Displayed scatter plot for {x} vs {y}"


# =====================================================================
# Chatbot Engine
# =====================================================================
class ChatBot:
    def __init__(self):
        self.analyser = DataAnalyser()

    def process(self, text):
        msg = text.lower().strip()

        if "help" in msg:
            return self.help_text()

        if "summary" in msg:
            return self.analyser.summary()

        if "correlation" in msg:
            return self.analyser.correlations()

        if "preview" in msg:
            return self.analyser.preview()

        if "missing" in msg:
            return self.analyser.missing_values_report()

        if "impute" in msg:
            return self.analyser.simulate_imputation()

        if "sample code" in msg or "example" in msg:
            return self.analyser.show_imputation_code()

        if "histogram" in msg:
            parts = text.split()
            col = parts[-1]
            return self.analyser.create_histogram(col)

        if "scatter" in msg and "vs" in msg:
            try:
                parts = msg.replace("scatter plot", "").split("vs")
                x, y = parts[0].strip(), parts[1].strip()
                return self.analyser.scatter_plot(x, y)
            except Exception:
                return "⚠️ Try: Scatter plot <col1> vs <col2>"

        if "export" in msg:
            return "💾 Export feature coming in next update."

        return "🤖 Try: 'Show summary statistics', 'Check missing values', 'Impute missing values', 'Show sample code', 'Scatter plot Age vs Salary', or 'Help'"

    def help_text(self):
        return """📚 COMPLETE COMMAND GUIDE — Data Analyser Bloop v1.0.2

📊 BASIC ANALYSIS
• Show summary statistics → display mean, std, min, max
• Calculate correlations → show correlation matrix
• Preview the data → view top rows of the dataset

🔍 DATA QUALITY
• Check missing values → detect nulls and give suggestions
• Impute missing values → show simulated imputation code (no real change)
• Show sample code for imputation → educational code snippet

📈 VISUAL ANALYSIS
• Histogram of <column> → show distribution of numeric column
• Scatter plot <x> vs <y> → show relationship between two columns

🧠 SMART SUGGESTIONS
• Type “How to handle missing data” → get LLM-style recommendations
• Type any column name to get quick hints

📦 FILE HANDLING
• Load Excel/CSV → via Load File button
• Export (coming soon)

💡 GENERAL HELP
• Help → show this guide
• Credits → show developer info
"""


# =====================================================================
# Autocomplete Entry (unchanged)
# =====================================================================
class AutoCompleteEntry(tk.Frame):
    def __init__(self, parent, suggestions, **kwargs):
        super().__init__(parent)
        self.suggestions = suggestions
        self.filtered = []
        self.text = tk.Text(self, height=3, wrap=tk.WORD, **kwargs)
        self.text.pack(fill=tk.BOTH, expand=True)
        self.listbox = tk.Listbox(self, height=6, bg="white", relief=tk.FLAT, font=("Arial", 9))
        self.listbox.pack_forget()

        self.text.bind("<KeyRelease>", self.update_list)
        self.text.bind("<Down>", self.down)
        self.text.bind("<Up>", self.up)
        self.text.bind("<Return>", self.enter)
        self.listbox.bind("<Double-Button-1>", self.select)

    def set_suggestions(self, suggestions):
        self.suggestions = suggestions or []

    def update_list(self, event):
        key = self.text.get("1.0", tk.END).strip().lower()
        if not key:
            self.listbox.pack_forget()
            return
        self.filtered = [s for s in self.suggestions if key in s.lower()]
        self.listbox.delete(0, tk.END)
        for s in self.filtered[:8]:
            self.listbox.insert(tk.END, s)
        if self.filtered:
            self.listbox.pack(fill=tk.X)
        else:
            self.listbox.pack_forget()

    def down(self, event):
        try:
            idx = self.listbox.curselection()[0]
            if idx < self.listbox.size() - 1:
                self.listbox.selection_clear(idx)
                self.listbox.selection_set(idx + 1)
        except IndexError:
            self.listbox.selection_set(0)

    def up(self, event):
        try:
            idx = self.listbox.curselection()[0]
            if idx > 0:
                self.listbox.selection_clear(idx)
                self.listbox.selection_set(idx - 1)
        except IndexError:
            pass

    def enter(self, event):
        if self.listbox.winfo_ismapped():
            self.select(None)
            return "break"

    def select(self, event):
        try:
            val = self.listbox.get(self.listbox.curselection()[0])
            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, val)
            self.listbox.pack_forget()
        except Exception:
            pass

    def get_text(self):
        return self.text.get("1.0", tk.END).strip()

    def clear(self):
        self.text.delete("1.0", tk.END)
        self.listbox.pack_forget()


# =====================================================================
# Help Panel (updated guide text will match ChatBot.help_text)
# =====================================================================
class HelpPanel(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.setup_help_ui()

    def setup_help_ui(self):
        title = tk.Label(self, text="📚 Command Guide", font=("Arial", 14, "bold"), bg="#f0f0f0")
        title.pack(pady=10)
        self.help_text = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Arial", 9), bg="white", relief=tk.FLAT, padx=10, pady=10)
        self.help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.help_text.insert("1.0", ChatBot().help_text())
        self.help_text.config(state=tk.DISABLED)


# =====================================================================
# GUI
# =====================================================================
class DataAnalyserBloopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analyser Bloop v1.0.2")
        self.root.geometry("1200x700")
        self.bot = ChatBot()
        self.help_panel_visible = False
        self.setup_ui()

    def setup_ui(self):
        header = tk.Frame(self.root, bg="#667eea", height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="📊 Data Analyser Bloop", fg="white", bg="#667eea", font=("Arial", 18, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(header, text="Version 1.0.2 | 100% Offline", fg="white", bg="#667eea").pack(side=tk.RIGHT, padx=20)

        toolbar = tk.Frame(self.root, bg="#f0f0f0")
        toolbar.pack(fill=tk.X)
        tk.Button(toolbar, text="📂 Load File", command=self.load_file, bg="#4CAF50", fg="white", relief=tk.FLAT, padx=20, pady=5).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(toolbar, text="📚 Show Help", command=self.toggle_help, bg="#2196F3", fg="white", relief=tk.FLAT, padx=12, pady=5).pack(side=tk.RIGHT, padx=10, pady=10)
        tk.Button(toolbar, text="ℹ️ Credits", command=self.show_credits, bg="#9C27B0", fg="white", relief=tk.FLAT, padx=12, pady=5).pack(side=tk.RIGHT, padx=10, pady=10)

        self.file_label = tk.Label(toolbar, text="No file loaded", bg="#f0f0f0", fg="#666")
        self.file_label.pack(side=tk.LEFT, padx=10)

        content_frame = tk.Frame(self.root, bg="#f7f7f7")
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        left_frame = tk.Frame(content_frame, bg="#f7f7f7")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.text_area = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD, bg="white", font=("Arial", 10))
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=(0, 10))
        self.text_area.insert(tk.END, "👋 Welcome to Data Analyser Bloop!\nLoad a file to get started.\n\n")

        self.help_panel = HelpPanel(content_frame, bg="#f0f0f0", width=360)

        bottom_frame = tk.Frame(self.root, bg="#f0f0f0")
        bottom_frame.pack(fill=tk.X, padx=10, pady=(0, 15))

        self.suggestions = [
            "Show summary statistics",
            "Calculate correlations",
            "Preview the data",
            "Check missing values",
            "Impute missing values",
            "Show sample code for imputation",
            "Histogram of Salary",
            "Scatter plot Age vs Salary",
            "Help"
        ]

        self.input_bar = AutoCompleteEntry(bottom_frame, self.suggestions, bg="white")
        self.input_bar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        send_btn = tk.Button(bottom_frame, text="Send ➤", bg="#667eea", fg="white", font=("Arial", 12, "bold"), relief=tk.FLAT, command=self.send_message)
        send_btn.pack(side=tk.RIGHT)

    def load_file(self):
        path = filedialog.askopenfilename(title="Select Excel/CSV File", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if path:
            response = self.bot.analyser.load_excel(path)
            self.file_label.config(text=f"📁 {Path(path).name}", fg="#4CAF50")
            try:
                cols = list(self.bot.analyser.df.columns.astype(str))
            except Exception:
                cols = []
            new_suggestions = self.suggestions + cols
            self.input_bar.set_suggestions(new_suggestions)
            self.display("🤖 Assistant", response)

    def show_credits(self):
        credits = """🪪 CREDITS & VERSION INFO
Data Analyser Bloop v1.0.2
Developed by Soumili Panja (2025)
Built with Tkinter, Pandas, Matplotlib, Seaborn, Scikit-learn
100% Offline"""
        self.display("ℹ️ Info", credits)

    def toggle_help(self):
        if not self.help_panel_visible:
            self.help_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(0, 0))
            self.help_panel_visible = True
        else:
            self.help_panel.pack_forget()
            self.help_panel_visible = False

    def send_message(self):
        msg = self.input_bar.get_text()
        if not msg:
            return
        self.display("👤 You", msg)
        self.input_bar.clear()
        threading.Thread(target=self.respond, args=(msg,), daemon=True).start()

    def respond(self, msg):
        response = self.bot.process(msg)
        self.display("🤖 Assistant", response)

    def display(self, sender, text):
        self.text_area.insert(tk.END, f"\n{sender}:\n{text}\n" + "-" * 70 + "\n")
        self.text_area.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyserBloopApp(root)
    root.mainloop()

# =====================================================================
