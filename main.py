from docx import Document
from openpyxl import load_workbook
import traceback

# =====================================================================
# FEATURE SCOPE & LIMITATIONS:
# 
# [SUPPORTED]
# - Key-Value extraction from Excel (Column A = Key, Column B = Value).
# - Placeholder replacement ({{key}}) in Word paragraphs and table cells.
# - Basic data preservation (handles standard numbers, dates, and leading zeros).
# 
# [UNSUPPORTED]
# - Complex multi-page table layout reconstruction or visual styling from scratch.
# - Bulk row-by-row data population into tables (designed for key-value variables).
# - Direct formatting/style transfer (fonts, colors) from Excel cells to Word.
# =====================================================================


def load_data_from_excel(excel_path: str, sheet_name: str = None) -> dict:
    """Reads data from an Excel file.
    Assumes a 'Key-Value' format (Column A = Key, Column B = Value).
    """
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb[sheet_name] if sheet_name else wb.active
    data = {}
    for row in sheet.iter_rows(min_row=1, values_only=True):
        if row[0] is not None:
            key = str(row[0]).strip()
            value = row[1] if len(row) > 1 and row[1] is not None else ''
            data[key] = value
    return data


def generate_word_document(
    template_path: str, output_path: str, context: dict
):
    """Fills a Word template (.docx) with data from a dictionary.
    Searches text and tables for placeholders like {{key}} and replaces them with values.
    """
    doc = Document(template_path)
    
    def replace_placeholders_in_paragraphs(paragraphs):
        for p in paragraphs:
            for key, value in context.items():
                placeholder = f'{{{{{key}}}}}'
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, str(value))
                    
    replace_placeholders_in_paragraphs(doc.paragraphs)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders_in_paragraphs(cell.paragraphs)
                
    doc.save(output_path)
    print(f'Work is done and document is saved as: {output_path}')


if __name__ == '__main__':
    try:
        print('--- Generator files is active ---')
        excel_file = input('Type Excel file name (e.g. data.xlsx): ').strip()
        template_file = input('Type name of word file (e.g. template.docx): ').strip()
        output_file = input('Type output file name (e.g. output.docx): ').strip()
        
        data = load_data_from_excel(excel_file)
        generate_word_document(template_file, output_file, data)

        print('\n[SUCCESS] Process completed successfully.')

    except Exception:
        print('\n[ERROR] An error occurred:')
        traceback.print_exc()

# How to Run
# Clone or download this repository.
# Fill data.xlsx with your data.
# Update template.docx with your tags.
# Run the script: python main.py

# Customization & Enterprise Services
# Need this script tailored to your specific document workflows, integrated into a larger system, or looking for custom automation solutions for your business?
# Let's talk. Reach out directly via Telegram: @Myhamed91]











