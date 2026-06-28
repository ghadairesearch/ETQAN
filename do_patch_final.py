import re

with open("course_report.py", "r", encoding="utf-8") as f:
    text = f.read()

# 1. Fix word_paragraph element order (bidi must come before jc)
pattern_wp = re.compile(r"def word_paragraph\(text='', bold=False, color='', size='', alignment=''\):.*?\n    return paragraph", re.DOTALL)
new_wp = r"""def word_paragraph(text='', bold=False, color='', size='', alignment=''):
    is_arabic = False
    s_text = str(text) if text is not None else ''
    
    if s_text:
        is_arabic = any('\u0600' <= c <= '\u06FF' or '\u0750' <= c <= '\u077F' or '\u08A0' <= c <= '\u08FF' for c in s_text)
        if not alignment:
            if is_arabic:
                alignment = 'right'
            elif s_text.strip():
                alignment = 'left'

    paragraph = word_element('p')
    paragraph_properties = word_element('pPr')
    
    # OOXML Spec: bidi comes before jc
    if is_arabic:
        paragraph_properties.append(word_element('bidi', {word_tag('val'): '1'}))
    else:
        paragraph_properties.append(word_element('bidi', {word_tag('val'): '0'}))

    if alignment:
        paragraph_properties.append(word_element('jc', {word_tag('val'): alignment}))
        
    paragraph.append(paragraph_properties)

    run = word_element('r')
    if bold or color or size or is_arabic:
        run_properties = word_element('rPr')
        if bold:
            run_properties.append(word_element('b'))
        if color:
            run_properties.append(word_element('color', {word_tag('val'): str(color).lstrip('#')}))
        if size:
            run_properties.append(word_element('sz', {word_tag('val'): str(size)}))
        if is_arabic:
            run_properties.append(word_element('rtl', {word_tag('val'): '1'}))
        run.append(run_properties)

    text_element = word_element('t')
    text_element.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    text_element.text = clean_xml_text(s_text)
    run.append(text_element)
    paragraph.append(run)
    return paragraph"""

if pattern_wp.search(text):
    text = pattern_wp.sub(new_wp, text)
else:
    print("FAILED WP")

# 2. Remove accent_table (gold ribbon) from build_course_report_word_identity_blocks
pattern_id = re.compile(r"def build_course_report_word_identity_blocks.*?return blocks", re.DOTALL)
new_id = r"""def build_course_report_word_identity_blocks(course_info=None, branding=None, logo_rel_id=''):
    language = get_export_report_language() if has_request_context() else 'en'
    branding = apply_university_identity_colors(branding or get_report_branding())
    primary = docx_hex_color(branding.get('primary_color'))
    secondary = docx_optional_hex_color(branding.get('secondary_color'), branding.get('primary_color')) or primary

    blocks = []
    if logo_rel_id:
        blocks.append(word_image_paragraph(logo_rel_id, alignment='center'))
        blocks.append(word_paragraph(''))

    return blocks"""

if pattern_id.search(text):
    text = pattern_id.sub(new_id, text)
else:
    print("FAILED ID")

with open("course_report.py", "w", encoding="utf-8") as f:
    f.write(text)
print("SUCCESS")
