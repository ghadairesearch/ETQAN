import re

with open("course_report.py", "r", encoding="utf-8") as f:
    text = f.read()

# Define the new build_clo_results_docx function
new_func = """def build_clo_results_docx(stats, total_students=0, course_info=None, student_achievement_matrix=None):
    language = get_export_report_language() if has_request_context() else 'en'
    course_info = course_info or {}
    
    root = word_element('document')
    body = word_element('body')
    root.append(body)

    branding = apply_university_identity_colors(get_report_branding()) if has_request_context() else {}
    logo_path = resolve_branding_logo_path(branding, report_ready=True) if branding else ''
    logo_bytes = b''
    logo_ext = ''
    if logo_path and os.path.exists(logo_path):
        logo_ext = os.path.splitext(logo_path)[1].lower()
        if logo_ext in {'.jpg', '.jpeg', '.png'}:
            with open(logo_path, 'rb') as f:
                logo_bytes = f.read()

    logo_rel_id = 'rId2' if logo_bytes else ''
    insert_course_report_word_identity(body, course_info, branding, logo_rel_id)

    title = '\u062a\u0642\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645' if language == 'ar' else 'CLO Attainment Report'
    course_label = '\u0627\u0644\u0645\u0642\u0631\u0631' if language == 'ar' else 'Course'
    total_label = '\u0625\u062c\u0645\u0627\u0644\u064a \u0627\u0644\u0637\u0644\u0627\u0628' if language == 'ar' else 'Total Students'
    course_name = course_info.get('course_name') or course_info.get('raw_name') or ''
    course_id = course_info.get('course_id') or course_info.get('course_code') or ''
    course_text = f"{course_name} ({course_id})" if course_name and course_id else (course_name or course_id or '-')
    
    body.append(word_paragraph(title, bold=True))
    body.append(word_paragraph(f"{course_label}: {course_text}"))
    body.append(word_paragraph(f"{total_label}: {total_students or 0}"))
    body.append(word_paragraph(''))

    body.append(word_paragraph('????? ????? ??????' if language == 'ar' else 'CLO Definitions', bold=True))
    clo_definitions = build_clo_definitions(stats.keys())
    
    primary = docx_hex_color(branding.get('primary_color'))

    def_table = word_element('tbl')
    def_table_props = word_element('tblPr')
    def_table_props.append(word_element('tblW', {word_tag('w'): '5000', word_tag('type'): 'pct'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {word_tag('val'): 'single', word_tag('sz'): '6', word_tag('space'): '0', word_tag('color'): '808080'}))
    def_table_props.append(borders)
    def_table.append(def_table_props)
    
    def_headers = ['??????', '?????', '????'] if language == 'ar' else ['Domain', 'CLO', 'Wording']
    def_table.append(word_row(def_headers, header=True, fill=primary, color='FFFFFF'))
    for item in clo_definitions:
        domain_text = localized_clo_domain(item['domain'], language)
        def_table.append(word_row([domain_text, item['number'], item['wording']]))
    body.append(def_table)
    body.append(word_paragraph(''))

    body.append(word_paragraph('???? ???? ????? ??????' if language == 'ar' else 'CLO Achievement Summary', bold=True))
    body.append(build_clo_assessment_word_table(stats or {}, course_info, language))
    
    if student_achievement_matrix and student_achievement_matrix.get('students'):
        body.append(word_paragraph(''))
        body.append(word_paragraph('????? ?????? ?? ????? ??????' if language == 'ar' else 'Student CLO Achievement', bold=True))
        matrix_table = word_element('tbl')
        matrix_table_props = word_element('tblPr')
        matrix_table_props.append(word_element('tblW', {word_tag('w'): '5000', word_tag('type'): 'pct'}))
        mborders = word_element('tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            mborders.append(word_element(border_name, {word_tag('val'): 'single', word_tag('sz'): '6', word_tag('space'): '0', word_tag('color'): '808080'}))
        matrix_table_props.append(mborders)
        matrix_table.append(matrix_table_props)
        
        clos = student_achievement_matrix.get('clos') or []
        matrix_headers = ['????? ???????' if language == 'ar' else 'Student ID'] + [clo_number(c) for c in clos]
        matrix_table.append(word_row(matrix_headers, header=True, fill=primary, color='FFFFFF'))
        
        cells = student_achievement_matrix.get('cells') or {}
        for student_id in student_achievement_matrix.get('students', []):
            row_data = [display_student_id(student_id)]
            for clo in clos:
                cell = cells.get(student_id, {}).get(clo)
                if cell:
                    status = '?????' if cell.get('achieved') and language == 'ar' else '??? ?????' if language == 'ar' else 'Met' if cell.get('achieved') else 'Not met'
                    # FIX: Cast score to float before formatting to avoid ValueError on strings
                    score = 0.0
                    try:
                        score = float(cell.get('score', 0))
                    except (ValueError, TypeError):
                        pass
                    row_data.append(f"{score:.2f} ({status})")
                else:
                    row_data.append('-')
            matrix_table.append(word_row(row_data))
        body.append(matrix_table)

    body.append(word_element('sectPr'))
    document_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
        content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
        content_types.append('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        content_types.append('  <Default Extension="xml" ContentType="application/xml"/>')
        
        if logo_bytes:
            ext_clean = logo_ext.lstrip('.')
            ctype = 'image/png' if ext_clean == 'png' else 'image/jpeg'
            content_types.append(f'  <Default Extension="{ext_clean}" ContentType="{ctype}"/>')
            docx.writestr(f'word/media/logo{logo_ext}', logo_bytes)
            
        content_types.append('  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>')
        content_types.append('</Types>')
        docx.writestr('[Content_Types].xml', '\\n'.join(content_types))

        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        rels.append('  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>')
        rels.append('</Relationships>')
        docx.writestr('_rels/.rels', '\\n'.join(rels))

        if logo_bytes:
            doc_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            doc_rels.append(f'  <Relationship Id="{logo_rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo{logo_ext}"/>')
            doc_rels.append('</Relationships>')
            docx.writestr('word/_rels/document.xml.rels', '\\n'.join(doc_rels))
        
        docx.writestr('word/document.xml', document_xml)

    output.seek(0)
    return output.getvalue()"""

# Replace the old function in text
start_idx = text.find("def build_clo_results_docx(stats, total_students=0, course_info=None, student_achievement_matrix=None):")
if start_idx != -1:
    end_idx = text.find("def build_exam_mapping_docx", start_idx)
    if end_idx != -1:
        text = text[:start_idx] + new_func + "\n\n" + text[end_idx:]
        with open("course_report.py", "w", encoding="utf-8") as f:
            f.write(text)
        print("Patched successfully")
    else:
        print("Could not find end of function")
else:
    print("Could not find start of function")
