heres your code:
from docx import Document
from openpyxl import load_workbook
import traceback
def load_data_from_excel(excel_path: str, sheet_name: str = None) -> dict:
  """Читает данные из Excel-файла.
  Предполагаем формат «Ключ — Значение» (Колонка A — ключ, Колонка B — значение).
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
  """Заполняет шаблон Word (.docx) данными из словаря.
  Ищет в тексте и таблицах метки вида {{ключ}} и заменяет их на значения.
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
    excel_file = input('Type Excel file name (np. data.xlsx): ').strip()
    template_file = input(
        'Type name of word file (np. ???.docx): '
    ).strip()
    output_file = input(
    ).strip()
    data = load_data_from_excel(excel_file)
    generate_word_document(template_file, output_file, data)

    print('\n')

  except Exception:
    print('\n')
    traceback.print_exc()

  finally:
    input('\n')
