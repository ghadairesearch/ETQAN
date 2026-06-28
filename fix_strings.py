with open("course_report.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace("'????? ????? ??????'", "'\\u062a\\u0639\\u0631\\u064a\\u0641 \\u0646\\u0648\\u0627\\u062a\\u062c \\u0627\\u0644\\u062a\\u0639\\u0644\\u0645'")
text = text.replace("['??????', '?????', '????']", "['\\u0627\\u0644\\u0645\\u062c\\u0627\\u0644', '\\u0627\\u0644\\u0631\\u0645\\u0632', '\\u0627\\u0644\\u0646\\u0635']")
text = text.replace("'???? ???? ????? ??????'", "'\\u0645\\u0644\\u062e\\u0635 \\u062a\\u062d\\u0642\\u0642 \\u0646\\u0648\\u0627\\u062a\\u062c \\u0627\\u0644\\u062a\\u0639\\u0644\\u0645'")
text = text.replace("'????? ?????? ?? ????? ??????'", "'\\u0625\\u0646\\u062c\\u0627\\u0632 \\u0627\\u0644\\u0637\\u0644\\u0627\\u0628 \\u0641\\u064a \\u0646\\u0648\\u0627\\u062a\\u062c \\u0627\\u0644\\u062a\\u0639\\u0644\\u0645'")
text = text.replace("['????? ???????'", "['\\u0627\\u0644\\u0631\\u0642\\u0645 \\u0627\\u0644\\u062c\\u0627\\u0645\\u0639\\u064a'")
text = text.replace("status = '?????' if cell.get('achieved') and language == 'ar' else '??? ?????' if language == 'ar'",
                    "status = '\\u0645\\u062a\\u062d\\u0642\\u0642' if cell.get('achieved') and language == 'ar' else '\\u063a\\u064a\\u0631 \\u0645\\u062a\\u062d\\u0642\\u0642' if language == 'ar'")

with open("course_report.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Strings fixed.")
