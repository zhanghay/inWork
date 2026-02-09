import pandas as pd
import re
import os

def clean_filename(text):
    """Remove illegal characters from filename"""
    if pd.isna(text) or text is None:
        return "unknown"
    return re.sub(r'[\\/*?:"<>|]', '', str(text)).strip()[:50]

def clean_line_breaks_and_spaces(text):
    """Replace line breaks with space and normalize whitespace"""
    if pd.isna(text) or text is None:
        return ""
    text = str(text).replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def remove_dot_number_colon(text):
    """
    Remove patterns like ".1:", ".2:", ".12:" etc. (dot + digits + colon)
    Example: "问题.1:描述" -> "问题:描述"
    """
    if not text or not isinstance(text, str):
        return text
    # Replace ".数字:" with ":"
    after_str = re.sub(r'\.\d+：', '：', text)
    after_str = re.sub(r'null', '', after_str)
    return after_str

def split_numbered_items(text):
    """
    Split text containing numbered items (e.g., "1. ... 2. ..." or "1、... 2、...")
    Returns a list of cleaned items without numbering prefixes.
    """
    if not text or not isinstance(text, str):
        return []
    
    # Pattern to match numbered items: digit followed by ., 、, ), or ） 
    # Must be followed by non-digit to avoid matching "S11" etc.
    pattern = r'(?:^|[\s；;。])(\d+[\.、\)\）]\s*)(?=[^\d])'
    
    # Find all split positions
    matches = list(re.finditer(pattern, text))
    
    # If no clear numbering pattern found, return as single item
    if len(matches) < 2:
        return [text.strip()]
    
    # Split text at numbering positions
    items = []
    last_end = 0
    for i, match in enumerate(matches):
        if i > 0:  # Skip the first match (it's the start of first item)
            start = match.start()
            item = text[last_end:start].strip()
            # Remove leading numbering from the item if present at very start
            item = re.sub(r'^\d+[\.、\)\）]\s*', '', item)
            if item and not re.match(r'^\d+$', item):  # Skip empty or pure numbers
                items.append(item)
            last_end = match.start()
    
    # Add the last item
    last_item = text[last_end:].strip()
    last_item = re.sub(r'^\d+[\.、\)\）]\s*', '', last_item)
    if last_item and not re.match(r'^\d+$', last_item):
        items.append(last_item)
    
    # Fallback: if splitting produced no valid items, return original text
    return items if items else [text.strip()]

def clean_and_split_problems(cell_value):
    """
    Clean cell value (remove line breaks) and split into multiple problems if numbered items exist.
    Returns a list of cleaned problem descriptions.
    """
    if pd.isna(cell_value) or str(cell_value).strip().lower() == "null" or not cell_value:
        return []
    
    # Clean line breaks and normalize whitespace
    text = clean_line_breaks_and_spaces(cell_value)
    
    if not text:
        return []
    
    # Split numbered items
    items = split_numbered_items(text)
    
    # Further clean each item
    cleaned_items = []
    for item in items:
        item = item.strip()
        # Remove leading/trailing bullets or numbering remnants
        item = re.sub(r'^[\d\s\.\、\)\）\s]+', '', item)
        item = re.sub(r'[\s\.\。]+$', '', item)
        if item and len(item) > 2:  # Skip very short fragments
            cleaned_items.append(item)
    
    return cleaned_items if cleaned_items else [text]

def process_excel(file_path, output_dir="input"):
    """
    Process Excel file and generate rectification documents for each user
    
    Args:
        file_path: Path to Excel file
        output_dir: Output directory name
    """
    os.makedirs(output_dir, exist_ok=True)
    df = pd.read_excel(file_path, header=0)
    columns = df.columns.tolist()
    generated_count = 0

    for idx, row in df.iterrows():
        application_id = row[columns[0]] if not pd.isna(row[columns[0]]) else "未知编号"
        account_number = row[columns[1]] if not pd.isna(row[columns[1]]) else "未知户号"
        account_name = row[columns[2]] if not pd.isna(row[columns[2]]) else "未知户名"
        
        # Collect all problems (split multi-problem cells)
        problems = []
        for col_idx in range(4, len(columns)):
            process_name = columns[col_idx]
            cell_value = row[col_idx]
            
            # Skip null/empty values
            if pd.isna(cell_value) or str(cell_value).strip().lower() == "null":
                continue
            
            # Clean and split problems
            problem_items = clean_and_split_problems(cell_value)
            for item in problem_items:
                problems.append(f"{process_name}：{item}")
        
        # Generate file only if problems exist
        if problems:
            clean_account_name = clean_filename(account_name)
            clean_application_id = clean_filename(application_id)
            filename = f"整改单-{clean_application_id}-{clean_account_name}.txt"
            filepath = os.path.join(output_dir, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                # First line: ApplicationID-AccountNumber-AccountName
                f.write(f"{application_id}-{account_number}-{account_name}\n")
                f.write("\n")
                # Write each problem on separate line (after cleaning .数字: patterns)
                for problem in problems:
                    cleaned_problem = remove_dot_number_colon(problem)
                    f.write(cleaned_problem + "\n")
            
            print(f"Generated: {filename} ({len(problems)} issues)")
            generated_count += 1
        else:
            print(f"Skipped: {application_id}-{account_name} (no issues found)")
    
    print(f"\nProcessing complete! Generated {generated_count} rectification documents in '{output_dir}' folder")

# Usage
if __name__ == "__main__":
    process_excel("2024.xlsx")