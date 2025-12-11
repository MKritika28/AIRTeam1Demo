#!/usr/bin/env python3
"""
Analyze eCOMMERCE_1.xlsx and generate keyword categorization summary
"""

import pandas as pd
import re
from collections import Counter, defaultdict
from pathlib import Path

def analyze_ecommerce_file():
    """Analyze eCOMMERCE_1.xlsx file."""
    
    file_path = Path("C:\\Users\\kritika.maheshwari\\Documents\\VSCode\\eCOMMERCE_1.xlsx")
    
    if not file_path.exists():
        print(f"❌ Error: File not found at {file_path}")
        return
    
    print("=" * 160)
    print("eCOMMERCE_1.xlsx - KEYWORD CATEGORIZATION ANALYSIS")
    print("=" * 160)
    print()
    
    # Read Excel file
    excel_file = pd.ExcelFile(file_path)
    sheet_name = excel_file.sheet_names[0]
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"📁 File: {file_path.name}")
    print(f"📖 Sheet: {sheet_name}")
    print(f"📐 Size: {df.shape[0]} rows × {df.shape[1]} columns")
    print(f"📋 Columns: {list(df.columns)}")
    print()
    
    # Find Prerequisites column
    prereq_col = None
    for col in df.columns:
        if 'prerequisite' in col.lower() or 'pre' in col.lower():
            prereq_col = col
            break
    
    if not prereq_col:
        print(f"❌ Error: Prerequisites column not found!")
        print(f"Available columns: {list(df.columns)}")
        return
    
    prerequisites = df[prereq_col].dropna()
    
    print(f"✅ Found Prerequisites column: '{prereq_col}'")
    print(f"📊 Total prerequisites: {len(prerequisites)}")
    print()
    print("=" * 160)
    print()
    
    # Extract keywords
    common_words = {'the', 'and', 'or', 'a', 'an', 'in', 'is', 'has', 'with', 'for', 'to', 'of', 'at', 'by', 'from', 'on', 'that', 'this', 'as', 'be', 'are', 'was', 'were', 'been', 'being', 'have', 'should', 'must', 'can', 'will', 'may', 'could', 'would', 'it', 'its', 'if', 'exists', 'exist', 'under', 'also', 'only', 'not', 'just', 'some', 'more', 'no', 'up', 'out', 'so', 'do', 'does', 'did', 'then', 'there', 'here', 'each', 'all', 'which', 'when', 'what', 'where', 'who', 'why', 'how'}
    
    all_text = ' '.join([str(v) for v in prerequisites])
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
    
    # Print detailed table
    print("DETAILED KEYWORD CATEGORIZATION TABLE")
    print("-" * 160)
    print(f"{'Category':<25} | {'Keyword':<30} | {'Count':>8} | {'Percentage':>12} | {'% of Category':>15}")
    print("-" * 160)
    
    for category in sorted(categorized.keys()):
        if category == 'Other':
            continue
        keywords_list = categorized[category]['keywords']
        counts_list = categorized[category]['counts']
        category_total = sum(counts_list)
        
        for i, (keyword, count) in enumerate(sorted(zip(keywords_list, counts_list), key=lambda x: x[1], reverse=True)):
            percentage = f"{(count/len(words)*100):.1f}%"
            cat_percentage = f"{(count/category_total*100):.1f}%"
            if i == 0:
                print(f"{category:<25} | {keyword:<30} | {count:>8} | {percentage:>12} | {cat_percentage:>15}")
            else:
                print(f"{'':25} | {keyword:<30} | {count:>8} | {percentage:>12} | {cat_percentage:>15}")
        print("-" * 160)
    
    # Print Other category
    if 'Other' in categorized:
        keywords_list = categorized['Other']['keywords']
        counts_list = categorized['Other']['counts']
        category_total = sum(counts_list)
        
        for i, (keyword, count) in enumerate(sorted(zip(keywords_list, counts_list), key=lambda x: x[1], reverse=True)):
            percentage = f"{(count/len(words)*100):.1f}%"
            cat_percentage = f"{(count/category_total*100):.1f}%"
            if i == 0:
                print(f"{'Other':<25} | {keyword:<30} | {count:>8} | {percentage:>12} | {cat_percentage:>15}")
            else:
                print(f"{'':25} | {keyword:<30} | {count:>8} | {percentage:>12} | {cat_percentage:>15}")
        print("-" * 160)
    
    print()
    print("=" * 160)
    print()
    
    # Category summary
    print("CATEGORY SUMMARY TABLE")
    print("-" * 160)
    print(f"{'Category':<25} | {'Keywords':>10} | {'Total Count':>15} | {'% of Total':>12}")
    print("-" * 160)
    
    summary_data = []
    for category in sorted(categorized.keys()):
        num_keywords = len(categorized[category]['keywords'])
        total_count = sum(categorized[category]['counts'])
        percentage = f"{(total_count/len(words)*100):.1f}%"
        
        print(f"{category:<25} | {num_keywords:>10} | {total_count:>15} | {percentage:>12}")
        summary_data.append({
            'Category': category,
            'Keywords': num_keywords,
            'Total Count': total_count,
            'Percentage': percentage
        })
    
    print("-" * 160)
    print()
    
    # Save to CSV files
    print("=" * 160)
    print()
    
    # Save detailed categorization
    detail_data = []
    for category in sorted(categorized.keys()):
        for keyword, count in zip(categorized[category]['keywords'], categorized[category]['counts']):
            detail_data.append({
                'Category': category,
                'Keyword': keyword,
                'Count': count,
                'Percentage': f"{(count/len(words)*100):.1f}%"
            })
    
    detail_df = pd.DataFrame(detail_data)
    detail_df = detail_df.sort_values(['Category', 'Count'], ascending=[True, False])
    
    detail_csv = 'ecommerce_keyword_detailed.csv'
    detail_df.to_csv(detail_csv, index=False)
    print(f"✅ Detailed categorization saved to: {detail_csv}")
    
    # Save summary
    summary_df = pd.DataFrame(summary_data)
    summary_csv = 'ecommerce_keyword_summary.csv'
    summary_df.to_csv(summary_csv, index=False)
    print(f"✅ Category summary saved to: {summary_csv}")
    
    print()
    print("=" * 160)
    print()
    
    # Statistics
    print("STATISTICS")
    print("-" * 160)
    print(f"Total unique keywords: {len(word_freq)}")
    print(f"Total keyword occurrences: {len(words)}")
    print(f"Total categories: {len(categorized)}")
    print(f"Average keywords per category: {len(word_freq) / len(categorized):.1f}")
    print(f"Average occurrences per keyword: {len(words) / len(word_freq):.1f}")
    print()
    print("=" * 160)


if __name__ == "__main__":
    analyze_ecommerce_file()
