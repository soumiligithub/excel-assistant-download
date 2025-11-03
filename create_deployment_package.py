"""
Deployment Package Creator for Excel Assistant
Author: Soumili Panja
"""

import os
import zipfile
from datetime import datetime

def create_deployment_package():
    """Creates a complete deployment package with all necessary files"""
    
    print("📦 Creating Excel Assistant Deployment Package...")
    print("=" * 60)
    
    # Create project directory
    project_name = "ExcelAssistant_v1.0.0"
    if not os.path.exists(project_name):
        os.makedirs(project_name)
    
    # Copyright header
    copyright_header = '''"""
Excel Assistant - Local Excel File Management Tool
==================================================

Copyright (c) 2025 Soumili Panja
All Rights Reserved

Author: Soumili Panja
Email: Soumili.panja.074@gmail.com
Institution: Bengal Institute of Business Studies, Kolkata
Date: November 2025

License: Educational and Personal Use
Version: 1.0.0
"""

'''
    
    # Create README.md
    readme_content = f"""# Excel Assistant v1.0.0

**A Smart Local Excel File Management Tool**

Developed by: Soumili Panja  
Institution: Bengal Institute of Business Studies, Kolkata  
Email: Soumili.panja.074@gmail.com  
Date: {datetime.now().strftime('%B %Y')}

## 🎯 Overview

Excel Assistant is a powerful desktop application that provides a natural language interface for managing Excel files locally. No internet required, no cloud dependencies - your data stays private and secure on your computer.

## ✨ Features

- **Smart Auto-Complete**: Start typing and get intelligent command suggestions
- **Natural Language Processing**: Use simple, conversational commands
- **Arrow Key Navigation**: Easily navigate through command suggestions
- **Built-in Help System**: Searchable command reference at your fingertips
- **100% Local**: Works completely offline, ensuring data privacy

## 🚀 Quick Start

### Installation

1. **Requirements**: Python 3.7 or higher
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application**:
   ```bash
   python excel_assistant_enhanced_gui.py
   ```

### First Steps

1. Launch the application
2. Type "show help" to see available commands
3. Try "show files" to list Excel files in current directory
4. Start managing your Excel files with ease!

## 📚 Example Commands

- `show files` - List all Excel files
- `create new file sales_data` - Create a new Excel file
- `open customers.xlsx` - Open an existing file
- `add sheet Q1_Report` - Add a new sheet to current file
- `show help` - Display all available commands

## 🛠️ System Requirements

- **OS**: Windows 7+, macOS 10.12+, or Linux
- **Python**: 3.7 or higher
- **RAM**: 2GB minimum (4GB recommended)
- **Disk Space**: 50MB

## 📄 Copyright & License

© 2025 Soumili Panja. All Rights Reserved.

This software is provided for **educational and personal use only**.  
Commercial use requires explicit written permission from the author.

**Disclaimer**: This software is provided "as is" without warranty of any kind.

## 📧 Contact & Support

**Email**: Soumili.panja.074@gmail.com  
**Developer**: Soumili Panja  
**Institution**: Bengal Institute of Business Studies, Kolkata

For bug reports, feature requests, or general inquiries, please contact via email.

---

*Built with ❤️ by Soumili Panja*
"""
    
    # Create requirements.txt
    requirements_content = """# Excel Assistant Requirements
# Author: Soumili Panja

pandas>=1.3.0
openpyxl>=3.0.0

# Optional but recommended:
# xlrd>=2.0.1  # For reading older .xls files
"""
    
    # Create LICENSE.txt
    license_content = """Excel Assistant - Copyright Notice
=====================================

Copyright (c) 2025 Soumili Panja
All Rights Reserved

Author: Soumili Panja
Email: Soumili.panja.074@gmail.com
Institution: Bengal Institute of Business Studies, Kolkata
Date: November 2025

LICENSE TERMS
=============

1. PERMITTED USES
   - Educational use in academic settings
   - Personal, non-commercial use
   - Learning and skill development
   - Portfolio demonstration (with attribution)

2. RESTRICTED USES
   - Commercial use without written permission
   - Redistribution without attribution
   - Removal of copyright notices
   - Claiming authorship

3. COMMERCIAL LICENSE
   For commercial use, please contact:
   Soumili.panja.074@gmail.com

4. ATTRIBUTION
   When sharing or demonstrating this software:
   - Credit must be given to Soumili Panja
   - Copyright notices must remain intact
   - Include link to original author if possible

5. WARRANTY DISCLAIMER
   This software is provided "AS IS" without warranty of any kind,
   express or implied, including but not limited to the warranties
   of merchantability, fitness for a particular purpose, and
   noninfringement.

6. LIABILITY LIMITATION
   In no event shall the author be liable for any claim, damages,
   or other liability arising from the use of this software.

7. DATA PRIVACY
   This software processes data locally only. No data is collected,
   transmitted, or stored by the author. Users are responsible for
   their own data management and backups.

=====================================
For questions about licensing:
Soumili.panja.074@gmail.com
"""
    
    # Create sample_commands.txt
    sample_commands = """Excel Assistant - Sample Commands
====================================

Author: Soumili Panja
Email: Soumili.panja.074@gmail.com

GETTING STARTED
---------------
show help          - Display all available commands
show files         - List all Excel files in current directory

FILE OPERATIONS
---------------
create new file sales_2024         - Create new Excel file
open report.xlsx                   - Open existing file
close                              - Close current file
save                               - Save current file

SHEET OPERATIONS
----------------
show sheets                        - List all sheets in current file
add sheet Q1_Data                  - Add new sheet
delete sheet OldData               - Remove a sheet
rename sheet Sheet1 to Summary     - Rename a sheet

DATA OPERATIONS
---------------
show data                          - Display current sheet data
add row                            - Add a new row
delete row 5                       - Delete specific row
clear sheet                        - Clear all data in current sheet

TIPS
----
- Use arrow keys (↑↓) to navigate auto-complete suggestions
- Press Enter to select a suggestion
- Press Escape to close suggestion list
- Commands are case-insensitive
- Type partial commands to see suggestions

KEYBOARD SHORTCUTS
------------------
Ctrl+H             - Show/Hide help panel
Tab                - Accept first suggestion
Escape             - Cancel/Clear

Need more help? Type "show help" in the application!
"""
    
    # Write all files
    files = {
        'README.md': readme_content,
        'requirements.txt': requirements_content,
        'LICENSE.txt': license_content,
        'SAMPLE_COMMANDS.txt': sample_commands
    }
    
    for filename, content in files.items():
        filepath = os.path.join(project_name, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✓ Created {filename}")
    
    # Create a template for the main Python file
    main_file_note = os.path.join(project_name, 'ADD_YOUR_PYTHON_FILE_HERE.txt')
    with open(main_file_note, 'w') as f:
        f.write("""IMPORTANT: Add Your Python File Here
=====================================

Please add your main Python file:
- excel_assistant_enhanced_gui.py

Make sure to add the copyright header at the top!

Then delete this file and create the ZIP package.
""")
    print(f"✓ Created placeholder note")
    
    print("\n" + "=" * 60)
    print("✅ Deployment package structure created successfully!")
    print(f"\n📁 Location: ./{project_name}/")
    print("\n📝 Next Steps:")
    print("1. Add your excel_assistant_enhanced_gui.py to the folder")
    print("2. Make sure the copyright header is at the top of your .py file")
    print("3. Delete the ADD_YOUR_PYTHON_FILE_HERE.txt file")
    print("4. Create a ZIP file of the entire folder")
    print("5. Upload to GitHub or Google Drive")
    print("6. Update the download link in your HTML page")
    print("\n💡 Your package will include:")
    print("   • Main Python application")
    print("   • README with instructions")
    print("   • requirements.txt for dependencies")
    print("   • LICENSE with copyright info")
    print("   • Sample commands guide")
    
    return project_name

if __name__ == "__main__":
    try:
        package_name = create_deployment_package()
        print("\n🎉 Package creation completed!")
        print(f"\n📦 Package name: {package_name}")
        print("\n🌐 Ready for deployment!")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Please check your permissions and try again.")
