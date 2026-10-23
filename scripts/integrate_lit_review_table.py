"""
Script to cleanly integrate the 26-row Literature Review table from 'lit-reveiew-final.docx'
into Chapter 2 of 'Final BE Project Stage I Report_Updated.docx'.
"""
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def integrate_lit_review():
    # 1. Read source table from lit-reveiew-final.docx
    src_doc = docx.Document("lit-reveiew-final.docx")
    src_table = src_doc.tables[0]
    print(f"Source table extracted: {len(src_table.rows)} rows, {len(src_table.columns)} columns.")

    table_data = []
    for r in src_table.rows:
        row_vals = [c.text.strip() for c in r.cells]
        table_data.append(row_vals)

    # 2. Re-run builder to place this table directly inside Chapter 2
    # We will modify scripts/execute_sprint_2.py or build a dedicated script that builds the complete report
    # with this table cleanly positioned in Chapter 2.
    print("Table data ready for insertion.")
    return table_data

if __name__ == '__main__':
    integrate_lit_review()
