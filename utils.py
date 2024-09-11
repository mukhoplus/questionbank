from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT

def set_row_height_table(table):
  def set_row_height(row, height):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(height))
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)

  i = 0
  for row in table.rows:
    height = Pt(0.06) if i < 2 else Pt(0.09)
    set_row_height(row, height)
    i += 1

def set_font(cell, font_name, font_size, bold=False):
  for paragraph in cell.paragraphs:
    try:
      run = paragraph.runs[0]
      run.font.name = font_name
      run.font.size = Pt(font_size)
      run.font.bold = bold
      paragraph.style.font.name = font_name
      paragraph.style.font.size = Pt(font_size)
      paragraph.style.font.bold = bold
    except:
      pass

def apply_table_styles(table):
  header_cells = table.rows[0].cells
  for cell in header_cells:
    set_font(cell, 'NS regular', 13, bold=True)

  for row in table.rows[1:]:
    for cell in row.cells:
      set_font(cell, 'NS regular', 11)

def table_sort_center(document, row_index):
  document.tables[row_index].rows[0].cells[0].paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
