with open("course_report.py", "r", encoding="utf-8") as f:
    text = f.read()

import re

# 1. Replace build_course_report_word_identity_blocks
old_func_pattern = r"def build_course_report_word_identity_blocks\(.*?return \[header_table, word_paragraph\(''\), accent_table, word_paragraph\(''\)\]"

new_func = """def build_course_report_word_identity_blocks(course_info=None, branding=None, logo_rel_id=''):
    language = get_export_report_language() if has_request_context() else 'en'
    branding = apply_university_identity_colors(branding or get_report_branding())
    primary = docx_hex_color(branding.get('primary_color'))
    secondary = docx_optional_hex_color(branding.get('secondary_color'), branding.get('primary_color')) or primary

    blocks = []
    if logo_rel_id:
        blocks.append(word_image_paragraph(logo_rel_id, alignment='center'))
        blocks.append(word_paragraph(''))

    accent_table = word_element('tbl')
    accent_properties = word_element('tblPr')
    accent_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    if language == 'ar':
        accent_properties.append(word_element('bidiVisual'))
    accent_table.append(accent_properties)
    accent_table.append(word_row([''], fill=secondary, size='40', width='9600'))

    blocks.append(accent_table)
    blocks.append(word_paragraph(''))
    return blocks"""

text = re.sub(old_func_pattern, new_func, text, flags=re.DOTALL)

with open("course_report.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Patched identity blocks")
