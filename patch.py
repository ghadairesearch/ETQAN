import re

with open('course_report.py', 'r', encoding='utf-8') as f:
    text = f.read()

target = r"""        selected_clos = request.form.getlist(f'question_clo_{question}')\s*resolved = resolve_detected_clos_to_course_list(selected_clos, clos)\s*if resolved:\s*metrics\['detected_clo_mappings'\]\[question\] = resolved"""

replacement = """        selected_clos = request.form.getlist(f'question_clo_{question}')
        resolved = resolve_detected_clos_to_course_list(selected_clos, clos)
        if not resolved and selected_clos:
            resolved = [str(c).strip() for c in selected_clos if str(c).strip()]
        if not selected_clos:
            for k in request.form.keys():
                if k.strip().endswith(f'_{question}') and 'question_clo' in k:
                    alt = request.form.getlist(k)
                    resolved = resolve_detected_clos_to_course_list(alt, clos)
                    if not resolved:
                        resolved = [str(c).strip() for c in alt if str(c).strip()]
                    break
        if resolved:
            metrics['detected_clo_mappings'][question] = resolved"""

# Escape parens in target regex
target = target.replace("(", r"\(").replace(")", r"\)")

new_text = re.sub(target, replacement, text)
if text != new_text:
    with open('course_report.py', 'w', encoding='utf-8') as f:
        f.write(new_text)
    print('Replaced successfully')
else:
    print('Target not found')
