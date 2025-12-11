# MCPAgent: Excel Analysis & Keyword Categorization Tool

A comprehensive Python-based MCP (Model Context Protocol) server for analyzing Excel files, extracting prerequisites, and categorizing keywords with an interactive Gradio UI.

## 🎯 Features

### Core Functionality
- **Excel File Reading**: Parse and analyze Excel workbooks with multiple sheets
- **Prerequisite Analysis**: Extract and analyze prerequisites from structured and descriptive formats
- **Keyword Extraction**: Intelligent keyword identification with common word filtering
- **Categorization**: Organize keywords into meaningful business categories
- **CSV Export**: Generate detailed categorization reports

### Categories
- 👤 User/Account - Login, email, password, account management
- 📦 Product/Catalog - Products, items, cart, inventory
- 🛒 Order/Payment - Orders, checkout, payment, invoices
- 📍 Shipping/Delivery - Delivery, address, shipping details
- ✅ Status - Order status, processing, delivery status
- 🔄 Refund/Return - Refunds, returns, exchanges
- 💰 Pricing/Discount - Prices, discounts, coupons, taxes
- 🔍 Search/Filter - Search, filtering, sorting

### User Interface
- Interactive Gradio web interface at `http://127.0.0.1:7860`
- Multiple analysis tabs with formatted HTML table outputs
- Custom column selection for analysis
- Real-time keyword categorization
- CSV export functionality

## 📋 Project Structure

```
MCPAgent/
├── analyze_ecommerce_keywords.py      # eCOMMERCE_1.xlsx analysis
├── analyze_book1_attributes.py         # Book1.xlsx keyword analysis
├── book1_keyword_categorization_table.py  # Book1 categorization
├── analyze_prerequisites.py            # Prerequisites extraction
├── analyze_attribute_values.py         # Attribute analysis
├── extract_unique_attributes.py        # Comparative attribute analysis
├── gradio_ecommerce_ui.py             # Web UI for analysis
├── excel_analyzer.py                  # Core MCP server
├── requirements.txt                   # Python dependencies
├── .env                               # Environment variables
├── .gitignore                         # Git ignore rules
├── ecommerce_keyword_summary.csv      # eCOMMERCE categorization summary
├── ecommerce_keyword_detailed.csv     # Detailed keyword breakdown
└── README.md                          # This file
```

## 🚀 Getting Started

### Prerequisites
- Python 3.8+
- pip package manager
- Git

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/MKritika28/AIRTeam1Demo.git
   cd MCPAgent
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment**
   ```bash
   # Create .env file with required variables
   echo "ANALYSIS_COLUMN=Prerequisites" > .env
   ```

### Usage

#### Option 1: Web UI (Recommended)
```bash
python gradio_ecommerce_ui.py
```
Then open your browser to: **http://127.0.0.1:7860**

#### Option 2: Command Line Analysis
```bash
# Analyze eCOMMERCE_1.xlsx
python analyze_ecommerce_keywords.py

# Analyze Book1.xlsx
python analyze_book1_attributes.py

# Extract prerequisites from any file
python analyze_prerequisites.py "path/to/file.xlsx"
```

#### Option 3: MCP Server
```bash
python excel_analyzer.py
```

## 📊 Analysis Example

### Input (Prerequisites Column)
```
User must be logged in with valid email_addr and password
User cart should contain at least 1 product
Order checkout and payment processing complete
Refund initiated and processing status available
```

### Output (Keyword Categorization)
```
Category          | Keywords        | Count | Percentage
User/Account      | user, email     | 2     | 10.2%
Product/Catalog   | product, cart   | 2     | 10.2%
Order/Payment     | order, payment  | 2     | 10.2%
Status            | status          | 1     | 5.1%
Refund/Return     | refund          | 1     | 5.1%
Other             | [various]       | 14    | 71.4%
```

## 📈 Statistics (eCOMMERCE_1.xlsx)

- **Total Prerequisites**: 24
- **Unique Keywords**: 37
- **Total Keyword Occurrences**: 88
- **Average Keywords per Category**: 6.2
- **Average Occurrences per Keyword**: 2.4

### Top Keywords
1. `product` (9x) - Product/Catalog
2. `order` (7x) - Order/Payment
3. `cart` (6x) - Product/Catalog
4. `contains` (5x) - Other
5. `user` (4x) - User/Account

## 🔧 Configuration

### Gradio UI Settings
```python
# In gradio_ecommerce_ui.py
interface = gr.Interface(
    fn=analyze_column,
    inputs=[
        gr.Textbox(label="File Path"),
        gr.Textbox(label="Analysis Column", value="Prerequisites"),
        gr.Textbox(label="Sheet Name", value="Sheet1")
    ],
    outputs=gr.HTML(),
    title="Excel Keyword Analyzer",
    description="Extract and categorize keywords from Excel files"
)
```

### Keyword Categories Configuration
Edit the `categories` dictionary in analysis scripts:
```python
categories = {
    'User/Account': ['user', 'account', 'email', 'pwd', ...],
    'Product/Catalog': ['product', 'cart', 'item', ...],
    # ... more categories
}
```

## 📦 Dependencies

- **pandas** >= 2.0.0 - Data manipulation
- **openpyxl** >= 3.10.0 - Excel file handling
- **gradio** >= 4.0.0 - Web interface
- **python-dotenv** >= 1.0.0 - Environment configuration
- **airefinery-sdk** >= 1.18.1 - MCP support

See `requirements.txt` for complete list.

## 📝 Outputs

### CSV Files Generated
1. **ecommerce_keyword_summary.csv** - Category summary with totals
   - Columns: Category, Keywords, Total Count, Percentage
   - 6 rows (one per category)

2. **ecommerce_keyword_detailed.csv** - Detailed keyword breakdown
   - Columns: Category, Keyword, Count, Percentage
   - 37 rows (one per keyword)

3. **book1_keyword_categories.csv** - Book1 category summary
4. **book1_keyword_detailed_categorization.csv** - Book1 detailed breakdown

## 🌐 Web Interface

### Tabs Available
1. **Read File** - Display Excel file contents as table
2. **Analyze Sheet** - Get file metadata and statistics
3. **Summarize** - Generate data summary
4. **Prerequisites Analysis** - Extract and categorize keywords

### Features
- ✅ HTML table formatting
- ✅ Custom column selection
- ✅ Sheet name input
- ✅ CSV export
- ✅ JSON data export

## 🐛 Troubleshooting

### Import Errors
```bash
# Reinstall dependencies
pip install -r requirements.txt --force-reinstall
```

### File Not Found
- Ensure file path is absolute or relative to workspace
- Check file permissions
- Verify Excel file is not open in another application

### Gradio Server Issues
```bash
# Kill existing process and restart
# On Windows: taskkill /F /IM python.exe
# Then: python gradio_ecommerce_ui.py
```

### Git Push Issues
```bash
# If authentication needed, use personal access token
git remote set-url origin https://<token>@github.com/MKritika28/AIRTeam1Demo.git
git push origin main
```

## 📚 Code Examples

### Extract Keywords from File
```python
import pandas as pd
from collections import Counter
import re

file_path = "path/to/excel.xlsx"
df = pd.read_excel(file_path)
prerequisites = df['Prerequisites'].dropna()

# Extract and count keywords
all_text = ' '.join([str(v) for v in prerequisites])
words = re.findall(r'\b[a-zA-Z_]+\b', all_text.lower())
word_freq = Counter(words)

# Filter common words
common_words = {'the', 'and', 'or', 'a', 'in', 'is', ...}
filtered = {w: c for w, c in word_freq.items() 
           if w not in common_words and len(w) > 2}
```

### Categorize Keywords
```python
def categorize(keyword, categories):
    for category, patterns in categories.items():
        if any(pattern in keyword for pattern in patterns):
            return category
    return 'Other'
```

## 🔗 GitHub Repository

**URL**: https://github.com/MKritika28/AIRTeam1Demo

### Initial Commit
- 10 files uploaded
- 671 insertions
- Master branch (renamed to main)
- Full project history tracked

## 👤 Author
- **Kritika Maheshwari**
- Email: kritika.maheshwari@accenture.com

## 📄 License
[Add your license information]

## 🤝 Contributing
Contributions are welcome! Please follow the existing code style and add tests for new features.

## 🎓 Learning Resources

### Related Concepts
- [Pandas Documentation](https://pandas.pydata.org/)
- [Gradio Documentation](https://www.gradio.app/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Python Regular Expressions](https://docs.python.org/3/library/re.html)

## 📞 Support
For issues or questions:
1. Check the **Troubleshooting** section above
2. Review existing code comments
3. Open an issue on GitHub

---

**Last Updated**: December 11, 2025  
**Status**: Active Development
