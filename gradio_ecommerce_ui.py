#!/usr/bin/env python3
"""
Enhanced Gradio UI for eCOMMERCE Excel Analysis with Custom Column Support
"""

import gradio as gr
import pandas as pd
import re
from collections import Counter, defaultdict
from pathlib import Path
import os

def read_excel_file(file_path):
    """Read and display Excel file contents."""
    try:
        if not file_path or not os.path.exists(file_path):
            return "❌ File not found. Please provide a valid file path.", None
        
        excel_file = pd.ExcelFile(file_path)
        sheet_name = excel_file.sheet_names[0]
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Create HTML table
        html_table = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h3>📁 File: {Path(file_path).name}</h3>
            <p><strong>Sheet:</strong> {sheet_name}</p>
            <p><strong>Size:</strong> {df.shape[0]} rows × {df.shape[1]} columns</p>
            <p><strong>Columns:</strong> {', '.join(df.columns)}</p>
            <hr>
            <h4>Preview Data:</h4>
            {df.head(10).to_html(index=False, border=1)}
        </div>
        """
        
        return html_table, df.to_csv(index=False)
    except Exception as e:
        return f"❌ Error reading file: {str(e)}", None


def analyze_keywords(file_path, column_name):
    """Analyze keywords from specified column."""
    try:
        if not file_path or not os.path.exists(file_path):
            return "❌ File not found. Please provide a valid file path."
        
        if not column_name or column_name.strip() == "":
            return "❌ Please specify the column name to analyze."
        
        # Read Excel file
        excel_file = pd.ExcelFile(file_path)
        sheet_name = excel_file.sheet_names[0]
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Find column (case-insensitive)
        target_col = None
        for col in df.columns:
            if col.lower() == column_name.lower():
                target_col = col
                break
        
        if not target_col:
            available = ', '.join(df.columns)
            return f"❌ Column '{column_name}' not found.\n\nAvailable columns: {available}"
        
        # Extract data
        text_data = df[target_col].dropna()
        
        # Define common words to filter out
        common_words = {'the', 'and', 'or', 'a', 'an', 'in', 'is', 'has', 'with', 'for', 'to', 'of', 'at', 'by', 'from', 'on', 'that', 'this', 'as', 'be', 'are', 'was', 'were', 'been', 'being', 'have', 'should', 'must', 'can', 'will', 'may', 'could', 'would', 'it', 'its', 'if', 'exists', 'exist', 'under', 'also', 'only', 'not', 'just', 'some', 'more', 'no', 'up', 'out', 'so', 'do', 'does', 'did', 'then', 'there', 'here', 'each', 'all', 'which', 'when', 'what', 'where', 'who', 'why', 'how'}
        
        all_text = ' '.join([str(v) for v in text_data])
        words = re.findall(r'\b[a-zA-Z_]+\b', all_text.lower())
        word_freq = Counter([w for w in words if len(w) > 2 and w not in common_words])
        
        # Define categories
        categories = {
            'User/Account': ['user', 'account', 'email', 'pwd', 'login', 'email_addr', 'emailid', 'username', 'password'],
            'Product/Catalog': ['product', 'cart', 'item', 'catalog', 'category', 'sku', 'inventory'],
            'Order/Payment': ['order', 'checkout', 'payment', 'invoice', 'bill', 'transaction', 'purchase'],
            'Shipping/Delivery': ['shipping', 'delivery', 'address', 'zip', 'postal', 'package', 'warehouse'],
            'Status': ['status', 'processing', 'delivered', 'completed', 'pending', 'confirmed'],
            'Refund/Return': ['refund', 'return', 'exchange', 'cancel', 'reverse'],
            'Pricing/Discount': ['price', 'discount', 'coupon', 'promo', 'tax', 'fee'],
            'Search/Filter': ['search', 'filter', 'sort', 'browse', 'find'],
        }
        
        # Categorize keywords
        categorized = defaultdict(lambda: {'keywords': [], 'counts': []})
        uncategorized = []
        
        for word in word_freq.keys():
            found = False
            for category, keywords in categories.items():
                if any(kw in word for kw in keywords):
                    categorized[category]['keywords'].append(word)
                    categorized[category]['counts'].append(word_freq[word])
                    found = True
                    break
            if not found:
                uncategorized.append((word, word_freq[word]))
        
        if uncategorized:
            categorized['Other'] = {
                'keywords': [w[0] for w in uncategorized],
                'counts': [w[1] for w in uncategorized]
            }
        
        # Build HTML table
        html_output = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f9f9f9;">
            <h3>📊 Keyword Analysis Results</h3>
            <p><strong>Column Analyzed:</strong> {target_col}</p>
            <p><strong>Total Records:</strong> {len(text_data)}</p>
            <p><strong>Unique Keywords:</strong> {len(word_freq)}</p>
            <p><strong>Total Occurrences:</strong> {len(words)}</p>
            <hr>
            
            <h4>Category Summary</h4>
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr style="background-color: #4CAF50; color: white;">
                    <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Category</th>
                    <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Keywords</th>
                    <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Total Count</th>
                    <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">% of Total</th>
                </tr>
        """
        
        # Add category rows
        for category in sorted(categorized.keys()):
            num_keywords = len(categorized[category]['keywords'])
            total_count = sum(categorized[category]['counts'])
            percentage = f"{(total_count/len(words)*100):.1f}%"
            
            html_output += f"""
                <tr style="background-color: {'#f0f0f0' if category == 'Other' else '#ffffff'};">
                    <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold;">{category}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; text-align: center;">{num_keywords}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; text-align: center;">{total_count}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; text-align: center;">{percentage}</td>
                </tr>
            """
        
        html_output += """
            </table>
            
            <h4>Detailed Keyword Breakdown</h4>
            <table style="border-collapse: collapse; width: 100%;">
                <tr style="background-color: #2196F3; color: white;">
                    <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Category</th>
                    <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Keyword</th>
                    <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Count</th>
                    <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">% of Total</th>
                </tr>
        """
        
        # Add detailed rows
        for category in sorted(categorized.keys()):
            keywords_list = categorized[category]['keywords']
            counts_list = categorized[category]['counts']
            
            for keyword, count in sorted(zip(keywords_list, counts_list), key=lambda x: x[1], reverse=True):
                percentage = f"{(count/len(words)*100):.1f}%"
                html_output += f"""
                <tr style="background-color: {'#ffffff' if category != 'Other' else '#f9f9f9'};">
                    <td style="padding: 12px; border: 1px solid #ddd;">{category}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; font-family: monospace;">{keyword}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; text-align: center;">{count}</td>
                    <td style="padding: 12px; border: 1px solid #ddd; text-align: center;">{percentage}</td>
                </tr>
                """
        
        html_output += "</table></div>"
        
        return html_output
    
    except Exception as e:
        return f"❌ Error during analysis: {str(e)}"


def quick_ecommerce():
    return analyze_keywords(
        "C:\\Users\\kritika.maheshwari\\Documents\\VSCode\\eCOMMERCE_1.xlsx",
        "Prerequisites"
    )


def quick_book1():
    return analyze_keywords(
        "C:\\Users\\kritika.maheshwari\\OneDrive - Accenture\\Book1.xlsx",
        "Prerequisites"
    )


if __name__ == "__main__":
    with gr.Blocks() as demo:
        gr.Markdown("""
        # 📊 eCOMMERCE Excel Analysis Platform
        
        **Build advanced Excel analysis with keyword categorization and custom column support**
        
        ---
        """)
        
        with gr.Tabs():
            # Tab 1: File Reader
            with gr.Tab("📁 Read Excel File"):
                gr.Markdown("### Load and preview your Excel file")
                
                with gr.Row():
                    file_input = gr.Textbox(
                        label="📂 Excel File Path",
                        placeholder="e.g., C:\\Users\\...\\eCOMMERCE_1.xlsx",
                        lines=1
                    )
                    read_btn = gr.Button("📖 Read File", variant="primary")
                
                file_output = gr.HTML(label="File Preview")
                csv_output = gr.Textbox(label="Raw CSV Data", lines=10, interactive=False)
                
                read_btn.click(
                    read_excel_file,
                    inputs=[file_input],
                    outputs=[file_output, csv_output]
                )
            
            # Tab 2: Keyword Analysis
            with gr.Tab("🔍 Keyword Analysis"):
                gr.Markdown("### Analyze keywords from any column in your Excel file")
                
                with gr.Row():
                    with gr.Column(scale=3):
                        analysis_file = gr.Textbox(
                            label="📂 Excel File Path",
                            placeholder="e.g., C:\\Users\\...\\eCOMMERCE_1.xlsx",
                            lines=1
                        )
                    with gr.Column(scale=2):
                        column_name = gr.Textbox(
                            label="📋 Column Name",
                            placeholder="e.g., Prerequisites, Description, Features"
                        )
                
                analyze_btn = gr.Button("🚀 Analyze Keywords", variant="primary")
                
                analysis_output = gr.HTML(label="Analysis Results")
                
                # Example files info
                gr.Markdown("""
                **💡 Quick Examples:**
                - `C:\\Users\\kritika.maheshwari\\Documents\\VSCode\\eCOMMERCE_1.xlsx` → Column: `Prerequisites`
                - `C:\\Users\\kritika.maheshwari\\OneDrive - Accenture\\Book1.xlsx` → Column: `Prerequisites`
                """)
                
                analyze_btn.click(
                    analyze_keywords,
                    inputs=[analysis_file, column_name],
                    outputs=[analysis_output]
                )
            
            # Tab 3: Pre-configured Analysis
            with gr.Tab("⚡ Quick Analysis"):
                gr.Markdown("### Run pre-configured analysis on known files")
                
                with gr.Row():
                    ecom_btn = gr.Button("📦 eCOMMERCE_1.xlsx Analysis", variant="primary")
                    book1_btn = gr.Button("📕 Book1.xlsx Analysis", variant="primary")
                
                quick_output = gr.HTML(label="Quick Analysis Results")
                
                ecom_btn.click(quick_ecommerce, outputs=[quick_output])
                book1_btn.click(quick_book1, outputs=[quick_output])
            
            # Tab 4: Help & Documentation
            with gr.Tab("📚 Help & Documentation"):
                gr.Markdown("""
                ## How to Use This Application
                
                ### 1️⃣ Read Excel File Tab
                - Enter the full path to your Excel file
                - Click "Read File" to preview the contents
                - View column names and sample data
                
                ### 2️⃣ Keyword Analysis Tab
                - Provide the Excel file path
                - Specify the column name to analyze (case-insensitive)
                - The system will:
                  - Extract all keywords from the specified column
                  - Categorize them (User/Account, Product, Order, etc.)
                  - Calculate occurrence frequencies
                  - Generate a comprehensive summary table
                
                ### 3️⃣ Pre-configured Analysis
                - Click buttons to quickly analyze known files
                - No file path entry needed
                
                ### 📊 Output Features
                - **Category Summary**: Shows keyword count and percentage per category
                - **Detailed Breakdown**: Lists all keywords with individual counts
                - **Filtering**: Common words automatically filtered out
                - **Categorization**: 8 predefined categories + "Other" for uncategorized keywords
                
                ### 📁 Example File Paths
                ```
                C:\\Users\\kritika.maheshwari\\Documents\\VSCode\\eCOMMERCE_1.xlsx
                C:\\Users\\kritika.maheshwari\\OneDrive - Accenture\\Book1.xlsx
                C:\\Users\\kritika.maheshwari\\Downloads\\Testcases_01.xlsx
                ```
                
                ### ✨ Keyword Categories
                - **User/Account**: user, account, email, login, password
                - **Product/Catalog**: product, cart, item, catalog, sku
                - **Order/Payment**: order, checkout, payment, invoice
                - **Shipping/Delivery**: shipping, delivery, address, package
                - **Status**: status, processing, delivered, completed
                - **Refund/Return**: refund, return, exchange, cancel
                - **Pricing/Discount**: price, discount, coupon, tax
                - **Search/Filter**: search, filter, sort, browse
                - **Other**: Uncategorized domain-specific terms
                
                ### 🔧 Tips
                - Column names are case-insensitive
                - Minimum keyword length: 3 characters
                - Common English words are automatically filtered
                - All occurrences are counted and percentages calculated
                """)
        
        gr.Markdown("""
        ---
        **Version 2.0** | eCOMMERCE Excel Analysis Platform | Powered by Gradio
        """)
    
    print("=" * 80)
    print("🚀 LAUNCHING GRADIO UI - eCOMMERCE EXCEL ANALYSIS")
    print("=" * 80)
    print()
    
    # Launch with public sharing enabled
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        show_error=True
    )
