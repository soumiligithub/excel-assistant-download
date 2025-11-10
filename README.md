# 📊 Data Analyser Bloop v1.0.2

**Data Analyser Bloop** is a 100% offline, AI-assisted Excel/CSV data analysis application.  
It combines a **chatbot-style interface** with smart analytics, missing-value detection, and visualization — all in one desktop GUI built using **Tkinter**.

---

## ✨ Features

### 🧠 Smart Analysis Engine
- Load `.xlsx`, `.xls`, or `.csv` files instantly  
- Chatbot-style natural commands (e.g., *“Show summary statistics”*, *“Check missing values”*)  
- AI-like missing data recommendations  
- Simulated imputation (preview code suggestions, no data change)  
- LLM-style reasoning about column quality

### 📈 Visualization
- Create **Histograms** for numeric columns  
- Generate **Scatter Plots** for column relationships  
- View quick **data previews** and summaries

### 💬 Chatbot Commands
| Category | Example Command | Description |
|-----------|----------------|--------------|
| Basic Analysis | `Show summary statistics` | Summary of numeric data |
|  | `Calculate correlations` | Correlation matrix |
|  | `Preview the data` | Show top 10 rows |
| Missing Data | `Check missing values` | Detect nulls and get recommendations |
|  | `Impute missing values` | Simulated code suggestions |
|  | `Show sample code for imputation` | Educational Python code |
| Visuals | `Histogram of Salary` | Show column distribution |
|  | `Scatter plot Age vs Salary` | Compare two variables |
| Help | `Help` | Show all commands |
| Info | `Credits` | View app details |

---

## 🖥️ Interface Overview
- **Top bar:** Load file, open Help, view Credits  
- **Left side:** Chat window for conversation  
- **Bottom:** Smart autocomplete command box  
- **Right side:** Built-in quick help tab  

---

## 📦 Installation

### Requirements
Make sure you have **Python 3.8+** installed.

Install dependencies:
```bash
pip install -r requirements.txt
python data_analyser_bloop_v1_0_2.py

