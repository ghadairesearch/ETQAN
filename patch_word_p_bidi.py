with open("course_report.py", "r", encoding="utf-8") as f:
    text = f.read()

old_func = """def word_paragraph(text='', bold=False, color='', size='', alignment=''):
    if not alignment and text:
        s_text = str(text)
        if any('\u0600' <= c <= '\u06FF' or '\u0750' <= c <= '\u077F' or '\u08A0' <= c <= '\u08FF' for c in s_text):
            alignment = 'right'
        elif any(c.isalpha() for c in s_text):
            alignment = 'left'

    paragraph = word_element('p')
    if alignment:
        paragraph_properties = word_element('pPr')
        paragraph_properties.append(word_element('jc', {word_tag('val'): alignment}))
        paragraph.append(paragraph_properties)"""

new_func = """def word_paragraph(text='', bold=False, color='', size='', alignment=''):
    is_arabic = False
    if text:
        s_text = str(text)
        is_arabic = any('\u0600' <= c <= '\u06FF' or '\u0750' <= c <= '\u077F' or '\u08A0' <= c <= '\u08FF' for c in s_text)
        if not alignment:
            if is_arabic:
                alignment = 'right'
            elif any(c.isalpha() for c in s_text):
                alignment = 'left'

    paragraph = word_element('p')
    
    # We always need pPr if there is alignment OR if it's arabic (to set bidi)
    if alignment or is_arabic:
        paragraph_properties = word_element('pPr')
        if alignment:
            paragraph_properties.append(word_element('jc', {word_tag('val'): alignment}))
        if is_arabic:
            paragraph_properties.append(word_element('bidi'))
        paragraph.append(paragraph_properties)"""

if old_func in text:
    text = text.replace(old_func, new_func)
    with open("course_report.py", "w", encoding="utf-8") as f:
        f.write(text)
    print("Patched successfully")
else:
    print("Could not find function")
