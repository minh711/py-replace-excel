import pandas as pd
import logging

# Set up logging
logging.basicConfig(filename='text_replacement.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Function to read the Excel file and create a replacement dictionary
def create_replacement_dict(excel_file, switch_columns=False):
    try:
        df = pd.read_excel(excel_file, header=None)
        df.columns = ['A', 'B']  # Manually set the column names since there are no headers
    except Exception as e:
        logging.error(f"Error reading the Excel file: {e}")
        raise

    # Count frequency of items in columns A and B
    item_counts_A = len(df['A'])
    item_counts_B = len(df['B'])
    
    # Log the counts
    logging.info(f"Counts for column A: {item_counts_A}")
    logging.info(f"Counts for column B: {item_counts_B}")

    if switch_columns:
        return dict(zip(df['B'], df['A']))
    else:
        return dict(zip(df['A'], df['B']))

# Function to replace text in the file based on the replacement dictionary
def replace_text_in_file(text_file, replacements):
    try:
        with open(text_file, 'r', encoding='utf-8') as file:
            content = file.read()
    except Exception as e:
        logging.error(f"Error reading the text file: {e}")
        raise

    # Initialize counts for each replacement
    replacement_counts = {old: 0 for old in replacements.keys()}
    
    for old, new in replacements.items():
        count = content.count(old)  # Count occurrences of old text
        content = content.replace(old, new)
        logging.info(f"Replaced '{old}' with '{new}' ({count} occurrences)")
        replacement_counts[old] = count
    
    # Calculate total sum of replaced items
    total_replacements = sum(replacement_counts.values())

    try:
        with open(text_file, 'w', encoding='utf-8') as file:
            file.write(content)
    except Exception as e:
        logging.error(f"Error writing to the text file: {e}")
        raise

    logging.info(f"Finished writing replacements to {text_file}")
    logging.info(f"Replacement counts: {replacement_counts}")
    logging.info(f"Total replacements: {total_replacements}")

# Main function to perform the task
def main(excel_file, text_file, switch_columns=False):
    logging.info(f"Starting text replacement from {excel_file} in {text_file}")
    replacements = create_replacement_dict(excel_file, switch_columns)
    replace_text_in_file(text_file, replacements)
    logging.info("Text replacement process completed")

if __name__ == "__main__":
    # Example usage
    excel_file = 'data.xlsx'  # Replace with your Excel file path
    text_file = 'text.txt'    # Replace with your text file path
    switch_columns = True    # Set to True to switch columns A and B
    
    main(excel_file, text_file, switch_columns)
