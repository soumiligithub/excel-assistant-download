# ENHANCED GUI WITH AUTOCOMPLETE AND BUILT-IN HELP

enhanced_gui = '''
"""
Excel Analysis Assistant - Enhanced GUI with Autocomplete
Features: Built-in help, command suggestions, auto-complete
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import pandas as pd
from pathlib import Path
import threading
import re

class AutocompleteEntry(tk.Frame):
    """
    Custom entry widget with autocomplete dropdown
    Shows suggestions as user types (like Google search)
    """

    def __init__(self, parent, suggestions, **kwargs):
        super().__init__(parent)

        self.suggestions = suggestions
        self.filtered_suggestions = []

        # Create text widget
        self.text = tk.Text(self, height=3, wrap=tk.WORD, **kwargs)
        self.text.pack(fill=tk.BOTH, expand=True)

        # Create dropdown listbox
        self.dropdown = tk.Listbox(self, height=8, relief=tk.FLAT,
                                   bg="white", selectmode=tk.SINGLE,
                                   font=("Arial", 9))
        self.dropdown.pack_forget()  # Hidden initially

        # Bind events
        self.text.bind("<KeyRelease>", self.on_keyrelease)
        self.text.bind("<Down>", self.on_down_arrow)
        self.text.bind("<Up>", self.on_up_arrow)
        self.text.bind("<Return>", self.on_return)
        self.text.bind("<Escape>", self.hide_dropdown)
        self.dropdown.bind("<Double-Button-1>", self.on_select)
        self.dropdown.bind("<Return>", self.on_select)

    def on_keyrelease(self, event):
        """Handle key release to show suggestions"""
        # Ignore special keys
        if event.keysym in ['Down', 'Up', 'Return', 'Escape', 'Shift_L', 'Shift_R',
                           'Control_L', 'Control_R', 'Alt_L', 'Alt_R']:
            return

        # Get current text
        current_text = self.get_text().strip().lower()

        if not current_text:
            self.hide_dropdown()
            return

        # Filter suggestions
        self.filtered_suggestions = [
            s for s in self.suggestions
            if current_text in s.lower()
        ][:15]  # Limit to 15 suggestions

        if self.filtered_suggestions:
            self.show_dropdown()
        else:
            self.hide_dropdown()

    def show_dropdown(self):
        """Show dropdown with suggestions"""
        self.dropdown.delete(0, tk.END)

        for suggestion in self.filtered_suggestions:
            self.dropdown.insert(tk.END, suggestion)

        if self.dropdown.size() > 0:
            self.dropdown.selection_set(0)
            self.dropdown.pack(fill=tk.X, pady=(5, 0))

    def hide_dropdown(self, event=None):
        """Hide dropdown"""
        self.dropdown.pack_forget()

    def on_down_arrow(self, event):
        """Navigate down in dropdown"""
        if self.dropdown.winfo_ismapped():
            current = self.dropdown.curselection()
            if current:
                next_index = min(current[0] + 1, self.dropdown.size() - 1)
                self.dropdown.selection_clear(0, tk.END)
                self.dropdown.selection_set(next_index)
                self.dropdown.see(next_index)
            return "break"

    def on_up_arrow(self, event):
        """Navigate up in dropdown"""
        if self.dropdown.winfo_ismapped():
            current = self.dropdown.curselection()
            if current:
                prev_index = max(current[0] - 1, 0)
                self.dropdown.selection_clear(0, tk.END)
                self.dropdown.selection_set(prev_index)
                self.dropdown.see(prev_index)
            return "break"

    def on_return(self, event):
        """Handle Enter key"""
        if self.dropdown.winfo_ismapped():
            # Select from dropdown
            self.on_select(event)
            return "break"
        # If dropdown not visible, let parent handle it
        return None

    def on_select(self, event):
        """Select suggestion from dropdown"""
        selection = self.dropdown.curselection()
        if selection:
            selected_text = self.dropdown.get(selection[0])
            self.text.delete("1.0", tk.END)
            self.text.insert("1.0", selected_text)
            self.hide_dropdown()

    def get_text(self):
        """Get text from widget"""
        return self.text.get("1.0", tk.END).strip()

    def clear_text(self):
        """Clear text"""
        self.text.delete("1.0", tk.END)
        self.hide_dropdown()

    def focus_set(self):
        """Set focus to text widget"""
        self.text.focus_set()


class HelpPanel(tk.Frame):
    """
    Built-in help panel with searchable commands
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.setup_help_ui()

    def setup_help_ui(self):
        """Setup help panel UI"""

        # Title
        title = tk.Label(self, text="📚 Command Guide",
                        font=("Arial", 14, "bold"), bg="#f0f0f0")
        title.pack(pady=10)

        # Search box
        search_frame = tk.Frame(self, bg="#f0f0f0")
        search_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(search_frame, text="🔍", font=("Arial", 12),
                bg="#f0f0f0").pack(side=tk.LEFT, padx=5)

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.on_search)

        search_entry = tk.Entry(search_frame, textvariable=self.search_var,
                               font=("Arial", 10), relief=tk.FLAT, bg="white")
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Help content
        self.help_text = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, font=("Arial", 9), bg="white",
            relief=tk.FLAT, padx=10, pady=10
        )
        self.help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configure tags for formatting
        self.help_text.tag_config("category", font=("Arial", 11, "bold"),
                                 foreground="#667eea")
        self.help_text.tag_config("command", font=("Courier", 9),
                                 foreground="#2c3e50")
        self.help_text.tag_config("description", font=("Arial", 9),
                                 foreground="#555")

        # Load help content
        self.load_help_content()

    def load_help_content(self):
        """Load help content"""
        self.help_text.config(state=tk.NORMAL)
        self.help_text.delete("1.0", tk.END)

        help_data = {
            "📊 Basic Statistics": [
                ("Show summary statistics", "Complete statistical overview"),
                ("Calculate correlations", "Find relationships between columns"),
                ("Check for missing values", "Detect incomplete data"),
                ("Preview the data", "View first 10 rows"),
            ],
            "🔍 Filtering Data": [
                ("Filter [column] equals [value]", "Example: Filter Department equals Sales"),
                ("Show where [column] > [value]", "Example: Show where Salary > 50000"),
                ("Filter [column] contains [text]", "Example: Filter Name contains John"),
                ("Reset filters", "Clear all active filters"),
                ("Show active filters", "View current filters"),
            ],
            "📈 Advanced Analysis": [
                ("Percentiles for [column]", "Show percentile distribution"),
                ("Frequency distribution of [column]", "Show frequency table"),
                ("Calculate z-scores for [column]", "Find standardized scores"),
                ("Find outliers in [column]", "Detect unusual values"),
                ("Confidence interval for [column]", "Calculate confidence interval"),
            ],
            "🔬 Hypothesis Testing": [
                ("Chi-square test between [col1] and [col2]", "Test independence"),
                ("ANOVA test [value_col] by [group_col]", "Compare group means"),
                ("Regression [x_col] vs [y_col]", "Linear regression analysis"),
            ],
            "📊 Grouping & Pivot": [
                ("Group by [column]", "Count by group"),
                ("Average [numeric_col] by [group_col]", "Group aggregation"),
                ("Pivot table [rows] by [columns]", "Create pivot table"),
            ],
            "📉 Visualizations": [
                ("Create bar chart for [column]", "Bar chart visualization"),
                ("Show histogram of [column]", "Distribution histogram"),
                ("Scatter plot [x_col] vs [y_col]", "Relationship plot"),
                ("Pie chart of [column]", "Proportion chart"),
                ("Box plot for [column]", "Box and whisker plot"),
            ],
            "🔧 Excel Formulas": [
                ("Build VLOOKUP", "VLOOKUP formula guide"),
                ("Build INDEX MATCH", "INDEX-MATCH guide"),
                ("Build SUMIFS", "Conditional sum guide"),
                ("Build IF formula", "IF statement guide"),
                ("Conditional formatting guide", "Formatting rules"),
            ],
            "💾 Export": [
                ("Export to HTML", "Beautiful formatted report"),
                ("Export to Excel", "Formatted spreadsheet"),
                ("Export to CSV", "Universal data format"),
                ("Show export history", "View all exports"),
            ],
            "💡 Natural Language": [
                ("What's the average [column]?", "Ask questions naturally"),
                ("How many [items] in [category]?", "Count queries"),
                ("Show me [description]", "Describe what you want"),
            ],
        }

        for category, commands in help_data.items():
            self.help_text.insert(tk.END, f"\\n{category}\\n", "category")
            self.help_text.insert(tk.END, "─" * 50 + "\\n", "description")

            for command, description in commands:
                self.help_text.insert(tk.END, f"  • ", "description")
                self.help_text.insert(tk.END, f"{command}\\n", "command")
                self.help_text.insert(tk.END, f"    {description}\\n\\n", "description")

        self.help_text.config(state=tk.DISABLED)
        self.original_content = self.help_text.get("1.0", tk.END)

    def on_search(self, *args):
        """Filter help content based on search"""
        search_term = self.search_var.get().lower()

        if not search_term:
            self.help_text.config(state=tk.NORMAL)
            self.help_text.delete("1.0", tk.END)
            self.help_text.insert("1.0", self.original_content)
            self.help_text.config(state=tk.DISABLED)
            return

        # Simple search implementation
        self.help_text.config(state=tk.NORMAL)
        lines = self.original_content.split("\\n")

        self.help_text.delete("1.0", tk.END)
        self.help_text.insert(tk.END, f"🔍 Search results for: '{search_term}'\\n\\n", "category")

        for line in lines:
            if search_term in line.lower():
                self.help_text.insert(tk.END, line + "\\n")

        self.help_text.config(state=tk.DISABLED)


class EnhancedExcelAssistantGUI:
    """
    Enhanced Excel Analysis Assistant GUI
    Features: Autocomplete, Built-in Help, Professional UI
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Analysis Assistant - Professional Edition")
        self.root.geometry("1400x800")

        # Command suggestions (all possible commands)
        self.command_suggestions = [
            # Basic stats
            "Show summary statistics",
            "Show comprehensive statistics",
            "Calculate correlations",
            "Check for missing values",
            "Preview the data",
            "Show data preview",

            # Filtering
            "Filter Department equals Sales",
            "Filter Department equals IT",
            "Filter Salary > 50000",
            "Filter Age < 30",
            "Show employees where Salary > 60000",
            "Reset filters",
            "Show active filters",

            # Advanced analysis
            "Percentiles for Salary",
            "Percentiles for Age",
            "Frequency distribution of Department",
            "Frequency distribution of Age",
            "Calculate z-scores for Salary",
            "Find outliers in Salary",
            "Find outliers in Age",
            "Confidence interval for Salary",

            # Hypothesis testing
            "Chi-square test between Department and Gender",
            "ANOVA test Salary by Department",
            "Regression Age vs Salary",
            "Regression Years_Experience vs Salary",
            "Multiple regression Age Years_Experience vs Salary",

            # Grouping
            "Group by Department",
            "Group by City",
            "Average Salary by Department",
            "Average Age by Department",
            "Pivot table Department by Gender",
            "Pivot table Department by City with Salary",

            # Charts
            "Create bar chart for Department",
            "Create bar chart for Gender",
            "Show histogram of Salary",
            "Show histogram of Age",
            "Scatter plot Age vs Salary",
            "Pie chart of Department",
            "Box plot for Salary",
            "Box plot for Performance_Rating",

            # Excel formulas
            "Build VLOOKUP",
            "Build INDEX MATCH",
            "Build XLOOKUP",
            "Build IF formula",
            "Build SUMIFS",
            "Conditional formatting guide",
            "Array formula guide",

            # Export
            "Export to HTML",
            "Export to Excel",
            "Export to CSV",
            "Show export history",

            # Natural language
            "What's the average salary?",
            "How many employees in Sales?",
            "Show me IT employees",
            "Help",
        ]

        # Initialize components
        # self.bot = FullFeaturedExcelChatBot()  # Your chatbot instance
        self.file_loaded = False
        self.current_file = None
        self.show_help_panel = False

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""

        # Colors
        bg_color = "#f0f0f0"
        header_color = "#667eea"

        self.root.configure(bg=bg_color)

        # ===== HEADER =====
        header_frame = tk.Frame(self.root, bg=header_color, height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="📊 Excel Analysis Assistant - Professional Edition",
            font=("Arial", 18, "bold"),
            bg=header_color,
            fg="white"
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=20)

        # Version label
        version_label = tk.Label(
            header_frame,
            text="v1.0.0 | 100% Offline",
            font=("Arial", 9),
            bg=header_color,
            fg="white"
        )
        version_label.pack(side=tk.RIGHT, padx=20)

        # ===== TOOLBAR =====
        toolbar_frame = tk.Frame(self.root, bg=bg_color, height=60)
        toolbar_frame.pack(fill=tk.X, side=tk.TOP)
        toolbar_frame.pack_propagate(False)

        # Left side buttons
        left_buttons = tk.Frame(toolbar_frame, bg=bg_color)
        left_buttons.pack(side=tk.LEFT, padx=10)

        self.load_btn = tk.Button(
            left_buttons,
            text="📂 Load File",
            command=self.load_file,
            font=("Arial", 10, "bold"),
            bg="#4CAF50",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.load_btn.pack(side=tk.LEFT, padx=5, pady=10)

        self.file_label = tk.Label(
            left_buttons,
            text="No file loaded",
            font=("Arial", 9),
            bg=bg_color,
            fg="#666"
        )
        self.file_label.pack(side=tk.LEFT, padx=15)

        # Right side buttons
        right_buttons = tk.Frame(toolbar_frame, bg=bg_color)
        right_buttons.pack(side=tk.RIGHT, padx=10)

        self.help_toggle_btn = tk.Button(
            right_buttons,
            text="📚 Show Help",
            command=self.toggle_help_panel,
            font=("Arial", 10),
            bg="#2196F3",
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        self.help_toggle_btn.pack(side=tk.RIGHT, padx=5, pady=10)

        export_btn = tk.Button(
            right_buttons,
            text="💾 Export",
            command=self.export_menu,
            font=("Arial", 10),
            bg="#FF9800",
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        export_btn.pack(side=tk.RIGHT, padx=5, pady=10)

        # ===== MAIN CONTENT AREA =====
        main_container = tk.Frame(self.root, bg=bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Chat area (left side)
        self.chat_frame = tk.Frame(main_container, bg=bg_color)
        self.chat_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Chat Display
        self.chat_display = scrolledtext.ScrolledText(
            self.chat_frame,
            wrap=tk.WORD,
            font=("Arial", 10),
            bg="white",
            relief=tk.FLAT,
            padx=15,
            pady=15
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.config(state=tk.DISABLED)

        # Configure text tags
        self.chat_display.tag_config("user", foreground="#1976D2",
                                    font=("Arial", 10, "bold"))
        self.chat_display.tag_config("bot", foreground="#388E3C",
                                    font=("Arial", 10, "bold"))
        self.chat_display.tag_config("error", foreground="#D32F2F")
        self.chat_display.tag_config("formula", background="#f0f0f0",
                                    font=("Courier", 9))

        # Help panel (right side - initially hidden)
        self.help_panel = HelpPanel(main_container, bg="#f0f0f0", width=400)
        # Don't pack it yet - toggle with button

        # ===== INPUT AREA =====
        input_container = tk.Frame(self.root, bg=bg_color, height=120)
        input_container.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=(5, 15))
        input_container.pack_propagate(False)

        # Tips label
        tips_label = tk.Label(
            input_container,
            text="💡 Start typing to see suggestions | Press ↓ to navigate | Enter to select",
            font=("Arial", 8),
            bg=bg_color,
            fg="#666"
        )
        tips_label.pack(anchor=tk.W, pady=(0, 5))

        # Input frame with autocomplete
        input_frame = tk.Frame(input_container, bg=bg_color)
        input_frame.pack(fill=tk.BOTH, expand=True)

        # Autocomplete entry
        self.input_text = AutocompleteEntry(
            input_frame,
            suggestions=self.command_suggestions,
            font=("Arial", 10),
            relief=tk.FLAT,
            bg="white",
            padx=10,
            pady=10
        )
        self.input_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Bind Return for autocomplete entry
        self.input_text.text.bind("<Return>", self.on_enter_pressed)

        # Send button
        self.send_btn = tk.Button(
            input_frame,
            text="Send ➤",
            command=self.send_message,
            font=("Arial", 12, "bold"),
            bg=header_color,
            fg="white",
            relief=tk.FLAT,
            width=12,
            cursor="hand2"
        )
        self.send_btn.pack(side=tk.RIGHT, fill=tk.Y)

        # Welcome message
        self.add_bot_message(
            "👋 Welcome to Excel Analysis Assistant - Professional Edition!\\n\\n"
            "✨ Features:\\n"
            "  • 100+ Statistical Functions\\n"
            "  • Natural Language Understanding\\n"
            "  • Auto-complete Suggestions\\n"
            "  • Built-in Help Guide\\n"
            "  • 100% Offline & Secure\\n\\n"
            "📂 Click 'Load File' to get started!\\n"
            "📚 Click 'Show Help' to see all available commands\\n"
            "💡 Start typing to see command suggestions!"
        )

    def toggle_help_panel(self):
        """Toggle help panel visibility"""
        self.show_help_panel = not self.show_help_panel

        if self.show_help_panel:
            self.help_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
            self.help_toggle_btn.config(text="📚 Hide Help")
        else:
            self.help_panel.pack_forget()
            self.help_toggle_btn.config(text="📚 Show Help")

    def load_file(self):
        """Load Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            try:
                # Load file
                # response = self.bot.analyzer.load_excel(file_path)

                self.current_file = Path(file_path).name
                self.file_loaded = True
                self.file_label.config(text=f"📁 {self.current_file}", fg="#4CAF50")

                self.add_bot_message(
                    f"✅ Successfully loaded: {self.current_file}\\n\\n"
                    f"🎯 Ready to analyze! Try:\\n"
                    f"  • 'Show summary statistics'\\n"
                    f"  • 'Calculate correlations'\\n"
                    f"  • 'Create bar chart for Department'\\n\\n"
                    f"💡 Start typing to see more suggestions!"
                )
            except Exception as e:
                self.add_bot_message(f"❌ Error loading file: {str(e)}", tag="error")

    def send_message(self):
        """Send user message"""
        user_input = self.input_text.get_text()

        if not user_input:
            return

        # Clear input
        self.input_text.clear_text()

        # Check if file is loaded
        if not self.file_loaded and user_input.lower() not in ['help', 'hi', 'hello']:
            self.add_bot_message("❌ Please load an Excel file first!", tag="error")
            return

        # Display user message
        self.add_user_message(user_input)

        # Process in separate thread
        thread = threading.Thread(target=self.process_message, args=(user_input,))
        thread.daemon = True
        thread.start()

    def process_message(self, user_input):
        """Process user message"""
        try:
            # Process with bot
            # response = self.bot.process_message_with_export(user_input)

            # Placeholder
            response = f"This is where the response for '{user_input}' would appear.\\n\\nIntegrate your chatbot here!"

            # Update UI
            self.root.after(0, self.add_bot_message, response)
        except Exception as e:
            self.root.after(0, self.add_bot_message, f"❌ Error: {str(e)}", "error")

    def on_enter_pressed(self, event):
        """Handle Enter key"""
        # Check if dropdown is visible
        if self.input_text.dropdown.winfo_ismapped():
            return  # Let autocomplete handle it

        # Check if Shift is held
        if event.state & 0x1:
            return  # Allow newline

        # Send message
        self.send_message()
        return "break"

    def add_user_message(self, message):
        """Add user message to chat"""
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\\n👤 You:\\n", "user")
        self.chat_display.insert(tk.END, f"{message}\\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)

    def add_bot_message(self, message, tag="bot"):
        """Add bot message to chat"""
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\\n🤖 Assistant:\\n", tag)

        # Highlight Excel formulas in response
        lines = message.split("\\n")
        for line in lines:
            if line.strip().startswith("="):
                self.chat_display.insert(tk.END, f"{line}\\n", "formula")
            else:
                self.chat_display.insert(tk.END, f"{line}\\n")

        self.chat_display.insert(tk.END, "\\n" + "─"*80 + "\\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.see(tk.END)

    def export_menu(self):
        """Show export options"""
        if not self.file_loaded:
            messagebox.showwarning("No File", "Please load a file first!")
            return

        export_window = tk.Toplevel(self.root)
        export_window.title("Export Options")
        export_window.geometry("350x300")
        export_window.transient(self.root)
        export_window.configure(bg="#f0f0f0")

        tk.Label(
            export_window,
            text="Choose Export Format",
            font=("Arial", 14, "bold"),
            bg="#f0f0f0"
        ).pack(pady=20)

        buttons = [
            ("📄 Export to HTML", "html", "#2196F3"),
            ("📊 Export to Excel", "excel", "#4CAF50"),
            ("📋 Export to CSV", "csv", "#FF9800"),
            ("📝 Export to Text", "text", "#9C27B0"),
        ]

        for text, format_type, color in buttons:
            tk.Button(
                export_window,
                text=text,
                command=lambda f=format_type: self.export_action(f, export_window),
                width=25,
                font=("Arial", 10),
                bg=color,
                fg="white",
                relief=tk.FLAT,
                pady=10,
                cursor="hand2"
            ).pack(pady=5)

    def export_action(self, format_type, window):
        """Perform export"""
        window.destroy()
        self.input_text.text.delete("1.0", tk.END)
        self.input_text.text.insert("1.0", f"export to {format_type}")
        self.send_message()


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = EnhancedExcelAssistantGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
'''

# Save enhanced GUI
with open('excel_assistant_enhanced_gui.py', 'w', encoding='utf-8') as f:
    f.write(enhanced_gui)

print("✅ Enhanced GUI with Autocomplete & Built-in Help created!")
print("   File: excel_assistant_enhanced_gui.py")
print()
print("🎯 New Features:")
print("   ✅ Auto-complete dropdown (like Google Search)")
print("   ✅ Command suggestions as you type")
print("   ✅ Built-in help panel (toggleable)")
print("   ✅ Searchable command guide")
print("   ✅ Professional UI with better UX")
print("   ✅ Excel formula highlighting")
