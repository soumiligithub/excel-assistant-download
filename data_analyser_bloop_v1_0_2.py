"""
Data Analyser Bloop - Version 1.0.2-b2
-------------------------------------
Complete offline Excel analysis assistant (GUI + Chatbot)
Includes descriptive stats, viz, missing-data recommendations, simulated imputation,
and inferential tests: z-score, t-tests, ANOVA, chi-square.

Developer: Soumili Panja (2025)
"""

import sys
IN_NOTEBOOK = 'ipykernel' in sys.modules

import tkinter as tk
from tkinter import scrolledtext, filedialog
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import zscore, ttest_ind, ttest_rel, f_oneway, chi2_contingency
from pathlib import Path
from datetime import datetime
import threading
import warnings

warnings.filterwarnings("ignore")
sns.set_style("whitegrid")
plt.rcParams["figure.figsize"] = (10, 6)


# -------------------------
# Helpers
# -------------------------
def human_pct(x):
    try:
        return f"{x:.2%}"
    except Exception:
        return str(x)


def col_lists(df):
    """Return numeric and categorical column lists"""
    if df is None:
        return [], []
    numeric = df.select_dtypes(include=[np.number]).columns.tolist()
    categorical = df.select_dtypes(include=['object', 'category']).columns.tolist()
    return numeric, categorical


# -------------------------
# Analysis engine
# -------------------------
class DataAnalyser:
    def __init__(self):
        self.df = None
        self.file_name = None

    def load(self, path_or_buffer):
        try:
            if isinstance(path_or_buffer, (str, Path)):
                p = str(path_or_buffer)
                if p.lower().endswith('.csv'):
                    df = pd.read_csv(p)
                else:
                    df = pd.read_excel(p)
            else:
                # file-like
                try:
                    df = pd.read_excel(path_or_buffer)
                except Exception:
                    path_or_buffer.seek(0)
                    df = pd.read_csv(path_or_buffer)
            self.df = df
            self.file_name = getattr(path_or_buffer, 'name', Path(path_or_buffer).name if isinstance(path_or_buffer, (str, Path)) else "uploaded_file")
            return f"✅ Loaded {self.file_name}\nRows: {len(self.df)} | Columns: {len(self.df.columns)}\nColumns: {', '.join(map(str, self.df.columns))}"
        except Exception as e:
            return f"❌ Error loading file: {e}"

    # Descriptive
    def summary(self):
        if self.df is None:
            return "❌ Load a file first."
        desc = self.df.describe(include='all').transpose().fillna('-')
        return "📊 SUMMARY STATISTICS\n\n" + desc.to_string()

    def preview(self, n=10):
        if self.df is None:
            return "❌ Load a file first."
        return f"👀 PREVIEW ({n} rows)\n\n" + self.df.head(n).to_string(index=False)

    def columns(self):
        if self.df is None:
            return "❌ Load a file first."
        return "📋 COLUMNS:\n" + "\n".join([f"• {c}" for c in self.df.columns])

    # Missing
    def missing_report(self):
        if self.df is None:
            return "❌ Load a file first."
        missing = self.df.isnull().sum()
        total = len(self.df)
        if missing.sum() == 0:
            return "✅ No missing values detected."
        lines = ["🔎 MISSING VALUE REPORT\n"]
        for c, m in missing.items():
            if m > 0:
                lines.append(f"• {c}: {m} missing ({human_pct(m/total)}) — Suggest: {self._suggest_impute(c)}")
        lines.append("\nType 'Impute missing values' to see code examples (simulation only).")
        return "\n".join(lines)

    def _suggest_impute(self, col):
        if self.df is None:
            return "No data"
        dtype = self.df[col].dtype
        if np.issubdtype(dtype, np.number):
            return "median (numeric)"
        elif dtype.name in ('object', 'category'):
            return "mode (categorical)"
        else:
            return "ffill/backfill"

    def simulate_imputation(self):
        if self.df is None:
            return "❌ Load a file first."
        missing = self.df.isnull().sum()
        if missing.sum() == 0:
            return "✅ No missing values to impute."
        snippets = ["# Simulation: suggested imputation code (does NOT change data)\n"]
        for c, m in missing.items():
            if m == 0:
                continue
            dtype = self.df[c].dtype
            if np.issubdtype(dtype, np.number):
                snippets.append(f"df['{c}'].fillna(df['{c}'].median(), inplace=True)  # numeric")
            elif dtype.name in ('object', 'category'):
                snippets.append(f"df['{c}'].fillna(df['{c}'].mode()[0], inplace=True)  # categorical")
            else:
                snippets.append(f"df['{c}'].fillna(method='ffill', inplace=True)  # other")
        snippets.append("\n# Note: This is preview-only in Data Analyser Bloop.")
        return "\n".join(snippets)

    def sample_imputation_code(self):
        return """# Example imputation code
import pandas as pd
df = pd.read_excel("yourfile.xlsx")

# numeric median
df['Salary'].fillna(df['Salary'].median(), inplace=True)

# categorical mode
df['Department'].fillna(df['Department'].mode()[0], inplace=True)

# forward-fill
df['Date'].fillna(method='ffill', inplace=True)
"""

    # Visuals
    def histogram(self, col, bins=20):
        if self.df is None:
            return "❌ Load a file first."
        if col not in self.df.columns:
            return f"❌ Column '{col}' not found."
        series = pd.to_numeric(self.df[col], errors='coerce').dropna()
        if series.empty:
            return f"❌ Column '{col}' has no numeric data."
        plt.figure()
        sns.histplot(series, kde=True, bins=bins)
        plt.title(f"Histogram: {col}")
        plt.xlabel(col)
        plt.ylabel("Frequency")
        plt.show()
        return f"✅ Histogram displayed for {col}."

    def scatter(self, x, y):
        if self.df is None:
            return "❌ Load a file first."
        if x not in self.df.columns or y not in self.df.columns:
            return f"❌ One of '{x}' or '{y}' not found."
        xnum = pd.to_numeric(self.df[x], errors='coerce')
        ynum = pd.to_numeric(self.df[y], errors='coerce')
        df2 = pd.concat([xnum, ynum], axis=1).dropna()
        if df2.empty:
            return "❌ No overlapping numeric data for scatter plot."
        plt.figure()
        plt.scatter(df2.iloc[:, 0], df2.iloc[:, 1], alpha=0.6)
        plt.xlabel(x)
        plt.ylabel(y)
        plt.title(f"Scatter: {x} vs {y}")
        plt.show()
        return f"✅ Scatter plotted for {x} vs {y}."

    # Inferential stats
    def z_scores(self, col):
        if self.df is None:
            return "❌ Load a file first."
        if col not in self.df.columns:
            return f"❌ Column '{col}' not found."
        series = pd.to_numeric(self.df[col], errors='coerce').dropna()
        if series.empty:
            return f"❌ Column '{col}' has no numeric data."
        zs = (series - series.mean()) / series.std(ddof=0)
        header = f"📘 Z-scores for {col}\nMean={series.mean():.4f} Std={series.std(ddof=0):.4f}\n\nFirst 10 z-scores:\n{zs.head(10).to_string()}\n"
        code = f"\nPython code used:\nfrom scipy.stats import zscore\nzs = (df['{col}'] - df['{col}'].mean()) / df['{col}'].std(ddof=0)\n"
        return header + code

    def independent_ttest(self, col1, col2):
        if self.df is None:
            return "❌ Load a file first."
        if col1 not in self.df.columns or col2 not in self.df.columns:
            return f"❌ Columns not found: {col1}, {col2}"
        s1 = pd.to_numeric(self.df[col1], errors='coerce').dropna()
        s2 = pd.to_numeric(self.df[col2], errors='coerce').dropna()
        if s1.empty or s2.empty:
            return "❌ One or both columns have no numeric data."
        t_stat, p = ttest_ind(s1, s2, nan_policy='omit')
        header = f"📘 Independent t-test: {col1} vs {col2}\n"
        header += f"n1={len(s1)} mean1={s1.mean():.4f} sd1={s1.std(ddof=1):.4f}\n"
        header += f"n2={len(s2)} mean2={s2.mean():.4f} sd2={s2.std(ddof=1):.4f}\n"
        header += f"t = {t_stat:.4f}, p = {p:.6f}\n"
        interpretation = "Result: " + ("Significant difference (p < 0.05)" if p < 0.05 else "No significant difference (p >= 0.05)")
        code = f"\nPython code used:\nfrom scipy.stats import ttest_ind\nt, p = ttest_ind(df['{col1}'], df['{col2}'], nan_policy='omit')\n"
        return header + interpretation + "\n" + code

    def paired_ttest(self, col_before, col_after):
        if self.df is None:
            return "❌ Load a file first."
        if col_before not in self.df.columns or col_after not in self.df.columns:
            return f"❌ Columns not found: {col_before}, {col_after}"
        s1 = pd.to_numeric(self.df[col_before], errors='coerce')
        s2 = pd.to_numeric(self.df[col_after], errors='coerce')
        df2 = pd.concat([s1, s2], axis=1).dropna()
        if df2.empty:
            return "❌ No paired numeric observations found."
        t_stat, p = ttest_rel(df2.iloc[:, 0], df2.iloc[:, 1], nan_policy='omit')
        header = f"📘 Paired t-test: {col_before} vs {col_after}\n"
        header += f"n={len(df2)} mean_diff={(df2.iloc[:, 0] - df2.iloc[:, 1]).mean():.4f} sd_diff={(df2.iloc[:, 0] - df2.iloc[:, 1]).std(ddof=1):.4f}\n"
        header += f"t = {t_stat:.4f}, p = {p:.6f}\n"
        interpretation = "Result: " + ("Significant difference (p < 0.05)" if p < 0.05 else "No significant difference (p >= 0.05)")
        code = f"\nPython code used:\nfrom scipy.stats import ttest_rel\nt, p = ttest_rel(df['{col_before}'], df['{col_after}'], nan_policy='omit')\n"
        return header + interpretation + "\n" + code

    def anova(self, dependent, group_col):
        if self.df is None:
            return "❌ Load a file first."
        if dependent not in self.df.columns or group_col not in self.df.columns:
            return f"❌ Columns not found: {dependent}, {group_col}"
        groups = []
        for g, grp in self.df.groupby(group_col):
            series = pd.to_numeric(grp[dependent], errors='coerce').dropna()
            if not series.empty:
                groups.append(series)
        if len(groups) < 2:
            return "❌ Need at least two groups with numeric data for ANOVA."
        F, p = f_oneway(*groups)
        header = f"📘 One-way ANOVA: {dependent} by {group_col}\n"
        header += f"Groups: {len(groups)} | F = {F:.4f}, p = {p:.6f}\n"
        interpretation = "Result: " + ("Significant differences exist between group means (p < 0.05)" if p < 0.05 else "No significant differences (p >= 0.05)")
        code = f"\nPython code used:\nfrom scipy.stats import f_oneway\ngroups = [grp['{dependent}'].dropna() for name, grp in df.groupby('{group_col}')]\nF, p = f_oneway(*groups)\n"
        return header + interpretation + "\n" + code

    def chi_square(self, col1, col2):
        if self.df is None:
            return "❌ Load a file first."
        if col1 not in self.df.columns or col2 not in self.df.columns:
            return f"❌ Columns not found: {col1}, {col2}"
        table = pd.crosstab(self.df[col1], self.df[col2])
        if table.size == 0:
            return "❌ Not enough data to build contingency table."
        chi2, p, dof, ex = chi2_contingency(table)
        header = f"📘 Chi-square test: {col1} vs {col2}\n"
        header += f"Contingency table:\n{table.to_string()}\n\n"
        header += f"χ² = {chi2:.4f}, p = {p:.6f}, dof = {dof}\n"
        interpretation = "Result: " + ("Variables are associated (p < 0.05)" if p < 0.05 else "No evidence of association (p >= 0.05)")
        code = f"\nPython code used:\nfrom scipy.stats import chi2_contingency\ntable = pd.crosstab(df['{col1}'], df['{col2}'])\nchi2, p, dof, ex = chi2_contingency(table)\n"
        return header + interpretation + "\n" + code

    # Utility: suggest valid pairs for tests
    def suggest_pairs(self, test_type):
        if self.df is None:
            return []
        numeric, categorical = col_lists(self.df)
        suggestions = []
        if test_type == 't-test':
            # suggest pairs of numeric columns
            for i in range(len(numeric)):
                for j in range(i+1, len(numeric)):
                    suggestions.append((numeric[i], numeric[j]))
        elif test_type == 'paired':
            # paired: also pairs but must be same length rows; suggest numeric pairs
            for i in range(len(numeric)):
                for j in range(i+1, len(numeric)):
                    suggestions.append((numeric[i], numeric[j]))
        elif test_type == 'anova':
            # dependent numeric vs group categorical
            for num in numeric:
                for cat in categorical:
                    suggestions.append((num, cat))
        elif test_type == 'chi2':
            # categorical vs categorical
            for i in range(len(categorical)):
                for j in range(i+1, len(categorical)):
                    suggestions.append((categorical[i], categorical[j]))
        return suggestions


# -------------------------
# Chatbot logic (natural language parsing)
# -------------------------
class ChatBot:
    def __init__(self):
        self.analyser = DataAnalyser()

    def process(self, msg):
        text = msg.strip()
        low = text.lower()

        # HELP or GUIDE
        if 'help' in low:
            return self.help_text()

        if low in ('list columns', 'columns', 'show columns'):
            return self.analyser.columns()

        # LOAD file (path)
        if low.startswith('load '):
            arg = text[5:].strip().strip('"').strip("'")
            return self.analyser.load(arg)

        # PREVIEW/ SUMMARY
        if 'preview' in low or 'show data' in low:
            return self.analyser.preview()

        if 'summary' in low or 'statistics' in low:
            return self.analyser.summary()

        # COLUMNS / MISSING
        if 'column' in low and ('list' in low or 'show' in low):
            return self.analyser.columns()

        if 'missing' in low:
            if 'recommend' in low or 'handle' in low or 'how' in low:
                return self.analyser.missing_report()
            if 'impute' in low or 'simulate' in low:
                return self.analyser.simulate_imputation()
            return self.analyser.missing_report()

        # IMPUTATION SAMPLE CODE
        if 'sample code' in low or 'example' in low or 'imputation code' in low:
            return self.analyser.sample_imputation_code()

        # VISUALS
        if 'histogram' in low or low.startswith('hist '):
            # parse last token as column
            parts = text.split()
            # Try to find column by matching substrings
            col = self._extract_column_name(text)
            if col:
                return self.analyser.histogram(col)
            else:
                return "⚠️ Could not identify column for histogram. Try: 'Histogram of Salary'"

        if 'scatter' in low and 'vs' in low:
            # parse "scatter plot x vs y"
            if 'vs' in low:
                left, right = low.split('vs', 1)
                # find last word before vs and first after vs
                # attempt to match to column names
                colx = self._extract_column_name(left)
                coly = self._extract_column_name(right)
                if colx and coly:
                    return self.analyser.scatter(colx, coly)
                else:
                    return "⚠️ Could not parse scatter columns. Use: 'Scatter plot Age vs Salary'"

        # INFERENTIAL TESTS
        if 'z-score' in low or 'z score' in low or low.startswith('zscore') or 'compute z' in low:
            col = self._extract_column_name(text)
            if col:
                return self.analyser.z_scores(col)
            else:
                # suggest numeric columns
                nums, _ = col_lists(self.analyser.df)
                if not nums:
                    return "❌ No numeric columns available for z-score."
                return "⚠️ Column not specified. Suggested numeric columns:\n" + ", ".join(nums)

        # T-tests
        if 't-test' in low or 'ttest' in low or ('t test' in low and 'paired' not in low):
            # try to extract two columns
            cols = self._extract_two_columns(text)
            if cols:
                return self.analyser.independent_ttest(cols[0], cols[1])
            else:
                # suggest pairs
                pairs = self.analyser.suggest_pairs('t-test')
                if pairs:
                    sample = pairs[:6]
                    li = "\n".join([f"{a} vs {b}" for a, b in sample])
                    return f"⚠️ Columns not detected. Suggested numeric pairs for independent t-test:\n{li}\nType: 'Run t-test <col1> vs <col2>'"
                return "❌ Not enough numeric columns for t-test."

        # Paired t-test
        if ('paired' in low and 't' in low) or ('paired t-test' in low) or ('paired t test' in low):
            cols = self._extract_two_columns(text)
            if cols:
                return self.analyser.paired_ttest(cols[0], cols[1])
            else:
                pairs = self.analyser.suggest_pairs('paired')
                if pairs:
                    sample = pairs[:6]
                    li = "\n".join([f"{a} vs {b}" for a, b in sample])
                    return f"⚠️ Columns not detected. Suggested numeric pairs for paired t-test:\n{li}\nType: 'Paired t-test <col_before> vs <col_after>'"
                return "❌ Not enough numeric columns for paired t-test."

        # ANOVA
        if 'anova' in low:
            # check for "by" or "vs" to separate dependent and group
            if ' by ' in low:
                parts = low.split(' by ')
                dep = self._extract_column_name(parts[0])
                grp = self._extract_column_name(parts[1])
                if dep and grp:
                    return self.analyser.anova(dep, grp)
            # else suggest combos
            pairs = self.analyser.suggest_pairs('anova')
            if pairs:
                sample = pairs[:6]
                li = "\n".join([f"{dep} by {grp}" for dep, grp in sample])
                return f"⚠️ Columns not detected. Suggested combos for ANOVA (dependent by group):\n{li}\nType: 'Run ANOVA <dependent> by <group>'"
            return "❌ Not enough columns for ANOVA."

        # Chi-square
        if 'chi' in low or 'chi-square' in low:
            cols = self._extract_two_columns(text)
            if cols:
                return self.analyser.chi_square(cols[0], cols[1])
            else:
                pairs = self.analyser.suggest_pairs('chi2')
                if pairs:
                    sample = pairs[:6]
                    li = "\n".join([f"{a} vs {b}" for a, b in sample])
                    return f"⚠️ Columns not detected. Suggested categorical pairs for chi-square:\n{li}\nType: 'Chi-square <col1> vs <col2>'"
                return "❌ Not enough categorical columns for chi-square."

        # EXPORT / OTHERS
        if 'export' in low:
            return "ℹ️ Export features are planned in a future update."

        # FALLBACK
        return ("🤖 I didn't understand exactly. Try commands such as:\n"
                "• Show summary statistics\n• Preview data\n• Check missing values\n• Impute missing values\n"
                "• Z-score <column>\n• Run t-test <col1> vs <col2>\n• Paired t-test <before> vs <after>\n"
                "• Run ANOVA <dependent> by <group>\n• Chi-square <cat1> vs <cat2>\n• Histogram of <column>\n• Scatter plot <x> vs <y>\n"
                "Type 'Help' to see the full command list.")

    # Helpers to detect column names from free text
    def _extract_column_name(self, text):
        """Try to find a column name from the current dataframe by substring match."""
        if self.analyser.df is None:
            return None
        cols = list(map(str, self.analyser.df.columns))
        low = text.lower()
        # try exact-token match first
        tokens = [t.strip(",:;()") for t in low.replace('-', ' ').split()]
        # check multi-word column names by scanning
        # check longest-first to avoid partial matches
        sorted_cols = sorted(cols, key=lambda s: -len(s))
        for c in sorted_cols:
            if c.lower() in low:
                return c
            # try token matching for multi-word columns
            if any(token == c.lower() for token in tokens):
                return c
        # no match
        return None

    def _extract_two_columns(self, text):
        """Attempt to find two column names from text (col1 vs col2 patterns)."""
        if self.analyser.df is None:
            return None
        low = text.lower()
        # common separators: ' vs ', ' v ', ' vs. ', ' versus '
        sep_candidates = [' vs ', ' v ', ' versus ', ' vs. ']
        for sep in sep_candidates:
            if sep in low:
                left, right = low.split(sep, 1)
                c1 = self._extract_column_name(left)
                c2 = self._extract_column_name(right)
                if c1 and c2:
                    return (c1, c2)
        # also try "col1 and col2" patterns
        if ' and ' in low:
            parts = low.split(' and ')
            if len(parts) >= 2:
                c1 = self._extract_column_name(parts[0])
                c2 = self._extract_column_name(parts[1])
                if c1 and c2:
                    return (c1, c2)
        # fallback to None
        return None

    def help_text(self):
        return (
            "📚 FULL COMMANDS GUIDE — Data Analyser Bloop v1.0.2-b2\n\n"
            "BASIC:\n"
            "• Load <path>  — load a file by typing path (or use Load File button)\n"
            "• Preview / Show data — display top rows\n"
            "• List columns / Columns — list column names\n"
            "• Show summary statistics — descriptive stats\n\n"
            "MISSING DATA:\n"
            "• Check missing values — counts and % missing per column\n"
            "• Impute missing values — shows suggested code to impute (simulation only)\n"
            "• Show sample code for imputation — educational snippets\n\n"
            "VISUALS:\n"
            "• Histogram of <column>\n"
            "• Scatter plot <x> vs <y>\n\n"
            "INFERENTIAL TESTS (educational outputs + code):\n"
            "• Z-score <column> — compute z-scores\n"
            "• Run t-test <col1> vs <col2> — independent t-test\n"
            "• Paired t-test <before> vs <after> — paired t-test\n"
            "• Run ANOVA <dependent> by <group> — one-way ANOVA\n"
            "• Chi-square <cat1> vs <cat2> — chi-square test of independence\n\n"
            "TIPS:\n"
            "• If a command is ambiguous, the assistant will suggest valid column combinations.\n"
            "• The imputation commands are preview-only (won't modify your data).\n"
            "• For any test, you can type 'Show sample code for <test>' to see the Python code used.\n"
        )


# -------------------------
# UI Components (Autocomplete entry & Help panel)
# -------------------------
class AutoCompleteEntry(tk.Frame):
    def __init__(self, parent, suggestions=None, **kwargs):
        super().__init__(parent)
        self.suggestions = suggestions or []
        self.filtered = []
        self.text = tk.Text(self, height=3, wrap=tk.WORD, **kwargs)
        self.text.pack(fill=tk.BOTH, expand=True)
        self.listbox = tk.Listbox(self, height=8, bg='white', relief=tk.FLAT, font=("Arial", 9))
        self.listbox.pack_forget()

        self.text.bind("<KeyRelease>", self._on_key)
        self.text.bind("<Down>", self._on_down)
        self.text.bind("<Up>", self._on_up)
        self.text.bind("<Return>", self._on_return)
        self.listbox.bind("<Double-Button-1>", self._on_select)

    def set_suggestions(self, suggestions):
        self.suggestions = suggestions or []

    def _on_key(self, event):
        if event.keysym in ('Down','Up','Return','Escape','Shift_L','Shift_R','Control_L','Control_R'):
            return
        key = self.text.get("1.0", tk.END).strip().lower()
        if not key:
            self.listbox.pack_forget()
            return
        self.filtered = [s for s in self.suggestions if key in s.lower()]
        self.listbox.delete(0, tk.END)
        for itm in self.filtered[:10]:
            self.listbox.insert(tk.END, itm)
        if self.filtered:
            self.listbox.pack(fill=tk.X)
        else:
            self.listbox.pack_forget()

    def _on_down(self, event):
        if self.listbox.winfo_ismapped():
            try:
                i = self.listbox.curselection()[0]
                if i < self.listbox.size() - 1:
                    self.listbox.selection_clear(i)
                    self.listbox.selection_set(i+1)
            except IndexError:
                self.listbox.selection_set(0)
        return "break"

    def _on_up(self, event):
        if self.listbox.winfo_ismapped():
            try:
                i = self.listbox.curselection()[0]
                if i > 0:
                    self.listbox.selection_clear(i)
                    self.listbox.selection_set(i-1)
            except IndexError:
                pass
        return "break"

    def _on_return(self, event):
        if self.listbox.winfo_ismapped():
            self._on_select(event)
            return "break"

    def _on_select(self, event):
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


class HelpPanel(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.text = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Arial", 9), bg="white")
        self.text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.text.insert("1.0", ChatBot().help_text())
        self.text.config(state=tk.DISABLED)


# -------------------------
# Main GUI
# -------------------------
class DataAnalyserBloopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analyser Bloop v1.0.2-b2")
        self.root.geometry("1200x720")
        self.bot = ChatBot()
        self.help_visible = False
        self.setup_ui()

    def setup_ui(self):
        header_color = "#2b7de9"
        header = tk.Frame(self.root, bg=header_color, height=68)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="📊 Data Analyser Bloop", bg=header_color, fg="white", font=("Arial", 18, "bold")).pack(side=tk.LEFT, padx=18)
        tk.Label(header, text="v1.0.2-b2 | 100% Offline", bg=header_color, fg="white").pack(side=tk.RIGHT, padx=18)

        toolbar = tk.Frame(self.root, bg="#f0f0f0")
        toolbar.pack(fill=tk.X, padx=6, pady=6)
        tk.Button(toolbar, text="📂 Load File", bg="#4CAF50", fg="white", relief=tk.FLAT, padx=12, command=self.load_file).pack(side=tk.LEFT, padx=(6, 8))
        self.help_btn = tk.Button(toolbar, text="📚 Show Help", bg="#2196F3", fg="white", relief=tk.FLAT, padx=12, command=self.toggle_help)
        self.help_btn.pack(side=tk.RIGHT, padx=(8, 6))
        tk.Button(toolbar, text="ℹ️ Credits", bg="#9C27B0", fg="white", relief=tk.FLAT, padx=12, command=self.show_credits).pack(side=tk.RIGHT, padx=(8, 6))

        self.file_label = tk.Label(toolbar, text="No file loaded", bg="#f0f0f0", fg="#666")
        self.file_label.pack(side=tk.LEFT, padx=(8, 12))

        # main area
        main = tk.Frame(self.root, bg="#fafafa")
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))

        # left: chat/output
        left = tk.Frame(main, bg="#fafafa")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.output = scrolledtext.ScrolledText(left, wrap=tk.WORD, bg="white", font=("Arial", 10))
        self.output.pack(fill=tk.BOTH, expand=True, padx=(0,8), pady=6)
        self.output.insert(tk.END, "👋 Welcome to Data Analyser Bloop v1.0.2-b2\nLoad a file to begin. Type 'Help' for commands.\n\n")

        # right: help panel (initially hidden)
        self.help_panel = HelpPanel(main, width=360, bg="#f7f7f7")

        # bottom input
        bottom = tk.Frame(self.root, bg="#f0f0f0")
        bottom.pack(fill=tk.X, padx=8, pady=8)
        self.default_suggestions = [
            "Show summary statistics",
            "Preview the data",
            "List columns",
            "Check missing values",
            "Impute missing values",
            "Show sample code for imputation",
            "Histogram of Salary",
            "Scatter plot Age vs Salary",
            "Z-score <column>",
            "Run t-test <col1> vs <col2>",
            "Run ANOVA <dependent> by <group>",
            "Chi-square <cat1> vs <cat2>",
            "Help"
        ]
        self.input_bar = AutoCompleteEntry(bottom, suggestions=self.default_suggestions, bg="white")
        self.input_bar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,8))
        send = tk.Button(bottom, text="Send ➤", bg=header_color, fg="white", font=("Arial", 12, "bold"), relief=tk.FLAT, command=self.send)
        send.pack(side=tk.RIGHT)

    def load_file(self):
        path = filedialog.askopenfilename(title="Select Excel/CSV", filetypes=[("Excel or CSV","*.xlsx *.xls *.csv"), ("All files","*.*")])
        if not path:
            return
        msg = self.bot.analyser.load(path)
        self.file_label.config(text=f"📁 {Path(path).name}", fg="#4CAF50")
        # update suggestions with column names
        try:
            cols = list(map(str, self.bot.analyser.df.columns))
        except Exception:
            cols = []
        suggestions = self.default_suggestions + cols
        self.input_bar.set_suggestions(suggestions)
        self._display("🤖 Assistant", msg)

    def show_credits(self):
        credits = ("🪪 CREDITS & VERSION\n\n"
                   "Data Analyser Bloop v1.0.2-b2\n"
                   "Developer: Soumili Panja (2025)\n"
                   "Built with: Tkinter, pandas, numpy, matplotlib, seaborn, scipy\n"
                   "Mode: 100% Offline")
        self._display("ℹ️ Credits", credits)

    def toggle_help(self):
        if not self.help_visible:
            self.help_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(0,0))
            self.help_visible = True
            self.help_btn.config(text="📚 Hide Help")
        else:
            self.help_panel.pack_forget()
            self.help_visible = False
            self.help_btn.config(text="📚 Show Help")

    def send(self):
        text = self.input_bar.get_text()
        if not text.strip():
            return
        self._display("👤 You", text)
        self.input_bar.clear()
        threading.Thread(target=self._respond, args=(text,), daemon=True).start()

    def _respond(self, text):
        resp = self.bot.process(text)
        # If ambiguous suggestion list returned in string, still show to user
        self._display("🤖 Assistant", resp)

    def _display(self, sender, text):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.output.insert(tk.END, f"\n{sender} [{timestamp}]:\n{text}\n" + "-"*72 + "\n")
        self.output.see(tk.END)


# -------------------------
# Run
# -------------------------
if __name__ == "__main__":
    if IN_NOTEBOOK:
        print("Running in notebook: GUI suppressed. Import DataAnalyser and ChatBot for testing.")
    else:
        root = tk.Tk()
        app = DataAnalyserBloopApp(root)
        root.mainloop()
