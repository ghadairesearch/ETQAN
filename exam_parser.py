import csv
import json
import os
import re

try:
    import fitz
except ImportError:
    fitz = None

try:
    import pandas as pd
except ImportError:
    pd = None


AR_QUESTION = r'(?:ال)?سؤال|سؤال|س'
AR_ORDINALS = (
    r'(?:ال)?(?:أ|ا)?ول|(?:ال)?ثاني|(?:ال)?ثالث|(?:ال)?رابع|(?:ال)?خامس|'
    r'(?:ال)?سادس|(?:ال)?سابع|(?:ال)?ثامن|(?:ال)?تاسع|(?:ال)?عاشر'
)
AR_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')


class ExamParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.file_ext = os.path.splitext(filepath)[1].lower()
        self.raw_text = ''
        self.pages = []
        self.metadata = {
            'course_code': None,
            'exam_type': None,
            'total_marks': None,
            'course_name': None,
        }
        self.parsed_questions = []
        self.validation_warnings = []

    def parse(self):
        self.extract_text()
        self.extract_metadata()
        self.parsed_questions = []
        for item in self.detect_question_boundaries():
            q_type = self.classify_question_type(item['text'])
            self.parsed_questions.append({
                'question_id': f"Q{item['number']}",
                'question_text': item['text'].strip(),
                'question_type': q_type,
                'options': self.extract_mcq_options(item['text']) if q_type == 'MCQ' else {},
                'marks': self.extract_marks(item['text']),
                'mapped_clo': None,
                'page_number': item['page'],
            })
        self.validate_questions()
        return self.parsed_questions

    def extract_text(self):
        if self.file_ext == '.pdf':
            if fitz is None:
                raise ImportError('PyMuPDF is required for PDF parsing.')
            doc = fitz.open(self.filepath)
            for page_num in range(len(doc)):
                text = doc.load_page(page_num).get_text('text')
                self.pages.append({'page': page_num + 1, 'text': text})
                self.raw_text += text + '\n\n'
            return

        if self.file_ext == '.docx':
            try:
                import docx2txt
            except ImportError:
                raise ImportError('docx2txt is required for DOCX parsing.')
            text = docx2txt.process(self.filepath)
            self.pages.append({'page': 1, 'text': text})
            self.raw_text = text
            return

        if self.file_ext == '.txt':
            with open(self.filepath, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
            self.pages.append({'page': 1, 'text': text})
            self.raw_text = text
            return

        raise ValueError('Unsupported file format. Please use PDF, DOCX, or TXT.')

    def question_start_pattern(self):
        number = r'([0-9٠-٩۰-۹]{1,3})'
        return re.compile(
            rf'(?im)^\s*(?:'
            rf'(?:Q|Qu|Que|Ques|Question)\s*(?:No\.?|#|[-.:])?\s*{number}'
            rf'|{AR_QUESTION}\s*(?:رقم|#|[-.:])?\s*(?:{number}|({AR_ORDINALS}))'
            rf'|{number}\s*[\).:\-\u2013\u2014]'
            rf')',
            re.IGNORECASE,
        )

    def normalize_question_number(self, value, fallback):
        if value is None:
            return fallback
        text = str(value).strip()
        ordinals = {
            'الأول': 1, 'الاول': 1, 'أول': 1, 'اول': 1,
            'الثاني': 2, 'ثاني': 2,
            'الثالث': 3, 'ثالث': 3,
            'الرابع': 4, 'رابع': 4,
            'الخامس': 5, 'خامس': 5,
            'السادس': 6, 'سادس': 6,
            'السابع': 7, 'سابع': 7,
            'الثامن': 8, 'ثامن': 8,
            'التاسع': 9, 'تاسع': 9,
            'العاشر': 10, 'عاشر': 10,
        }
        if text in ordinals:
            return ordinals[text]
        try:
            return int(text.translate(AR_DIGITS))
        except ValueError:
            return fallback

    def exam_body_text(self):
        text = self.raw_text or ''
        stop_patterns = [
            r'\banswer\s+key\b',
            r'\banswer\s+sheet\b',
            r'\banswers?\s+table\b',
            r'\bmodel\s+answer\b',
            r'\bmarking\s+scheme\b',
            r'\bsolution\s+key\b',
            r'\bclo\s+coverage\s+summary\b',
            r'\bcoverage\s+summary\b',
            r'مفتاح\s+الإجاب',
            r'مفتاح\s+الاجاب',
            r'نموذج\s+الإجاب',
            r'نموذج\s+الاجاب',
            r'جدول\s+الإجاب',
            r'جدول\s+الاجاب',
            r'ملخص\s+.*نواتج',
            r'تغطية\s+.*نواتج',
        ]
        first_question = None
        match = self.question_start_pattern().search(text)
        if match:
            first_question = match.start()

        cut_at = None
        for pattern in stop_patterns:
            for stop in re.finditer(pattern, text, flags=re.IGNORECASE):
                if first_question is None or stop.start() > first_question:
                    cut_at = stop.start() if cut_at is None else min(cut_at, stop.start())
                    break
        return text[:cut_at].strip() if cut_at else text

    def clean_question_block(self, block):
        lines = [line.rstrip() for line in str(block or '').splitlines()]
        while lines and not lines[0].strip():
            lines.pop(0)
        if lines:
            lines[0] = re.sub(r'^\s*(?:[0-9٠-٩۰-۹]{1,3}\s*)?[\).:\-\u2013\u2014]\s*', '', lines[0])
        text = '\n'.join(lines).strip()
        return re.sub(r'\n{3,}', '\n\n', text)

    def looks_like_answer_key_row(self, block):
        text = re.sub(r'\s+', ' ', str(block or '')).strip()
        return bool(re.fullmatch(r'[A-Da-dأبجدهـ]\s*(?:CLO\s*\d+(?:\.\d+)*)?', text))

    def detect_question_boundaries(self):
        text = self.exam_body_text()
        matches = []
        for match in self.question_start_pattern().finditer(text):
            number_text = next((group for group in match.groups() if group), None)
            number = self.normalize_question_number(number_text, len(matches) + 1)
            matches.append((number, match.start(), match.end()))

        if not matches:
            return self.heuristic_question_detection()

        questions = []
        seen_numbers = set()
        for index, (number, _start, end) in enumerate(matches):
            next_start = matches[index + 1][1] if index + 1 < len(matches) else len(text)
            block = self.clean_question_block(text[end:next_start])
            if not block or self.looks_like_answer_key_row(block):
                continue
            if number in seen_numbers and number <= max(seen_numbers):
                continue
            seen_numbers.add(number)

            page_num = 1
            for page in self.pages:
                if block[:50] and block[:50] in page['text']:
                    page_num = page['page']
                    break
            questions.append({'number': number, 'text': block, 'page': page_num})

        return questions

    def heuristic_question_detection(self):
        lines = [line.strip() for line in self.exam_body_text().splitlines() if line.strip()]
        questions = []
        current = []
        q_num = 1
        header_words = {
            'student name', 'student id', 'section', 'quiz', 'midterm', 'final',
            'duration', 'department', 'college', 'university', 'instructions',
            'اسم الطالب', 'الرقم الجامعي', 'القسم', 'الكلية', 'الجامعة', 'تعليمات',
        }

        def is_header(text):
            lowered = text.lower()
            return len(text) < 180 and any(word in lowered for word in header_words) and '?' not in text and '؟' not in text

        for line in lines:
            current.append(line)
            joined = '\n'.join(current)
            if len(current) >= 5 and self.classify_question_type(joined) == 'MCQ':
                if not is_header(joined):
                    questions.append({'number': q_num, 'text': joined, 'page': 1})
                    q_num += 1
                    current = []
                continue
            if len(current) >= 2 and re.search(r'\b(?:true|false|صح|خطأ|صواب|غلط)\b', joined, flags=re.IGNORECASE):
                if not is_header(joined):
                    questions.append({'number': q_num, 'text': joined, 'page': 1})
                    q_num += 1
                    current = []

        if current and not is_header('\n'.join(current)):
            questions.append({'number': q_num, 'text': '\n'.join(current), 'page': 1})
        return questions

    def classify_question_type(self, text):
        lowered = str(text or '').lower()
        if re.search(r'\b(?:true|false|صواب|خطأ|صح|غلط)\b', lowered):
            return 'T/F'
        if self.extract_mcq_options(text):
            return 'MCQ'
        if '___' in lowered or 'complete' in lowered or 'أكمل' in lowered:
            return 'Fill in the blank'
        return 'Essay'

    def extract_mcq_options(self, text):
        options = {}
        pattern = re.compile(
            r'(?ms)^\s*([A-Da-dأبجدهـ])[\).:-]\s*(.*?)(?=^\s*[A-Da-dأبجدهـ][\).:-]\s*|\Z)'
        )
        for match in pattern.finditer(str(text or '')):
            key = match.group(1).upper()
            value = match.group(2).strip()
            if value:
                options[key] = value
        return options

    def extract_marks(self, text):
        patterns = [
            r'[\[\(]\s*(\d+(?:\.\d+)?)\s*(?:marks?|points?|درجات?|درجة)\s*[\]\)]',
            r'(?:marks?|points?|درجات?|درجة)\s*[:=]\s*(\d+(?:\.\d+)?)',
        ]
        for pattern in patterns:
            match = re.search(pattern, str(text or ''), re.IGNORECASE)
            if match:
                return float(match.group(1))
        return 1.0

    def extract_metadata(self):
        code = re.search(r'\b([A-Za-z]{2,5}\s?\d{3,4})\b', self.raw_text or '')
        if code:
            self.metadata['course_code'] = code.group(1).upper().replace(' ', '')

        lowered = (self.raw_text or '').lower()
        exam_types = {
            'midterm': 'Midterm',
            'final': 'Final',
            'quiz': 'Quiz',
            'assignment': 'Assignment',
            'نهائي': 'Final',
            'نصفي': 'Midterm',
            'اختبار قصير': 'Quiz',
            'قصير': 'Quiz',
        }
        for cue, label in exam_types.items():
            if cue in lowered:
                self.metadata['exam_type'] = label
                break

        total = re.search(r'(?:total\s+marks?|score|الدرجة\s+الكلية|المجموع)\s*[:=]?\s*(\d+(?:\.\d+)?)', self.raw_text or '', re.IGNORECASE)
        if total:
            self.metadata['total_marks'] = float(total.group(1))

    def validate_questions(self):
        warnings = []
        numbers = []
        for question in self.parsed_questions:
            match = re.search(r'\d+', question.get('question_id', ''))
            if match:
                numbers.append(int(match.group(0)))
        if numbers and numbers != list(range(min(numbers), max(numbers) + 1)):
            warnings.append('Warning: Question numbers are not strictly sequential.')
        if self.metadata['total_marks']:
            total = sum(question.get('marks', 0) for question in self.parsed_questions)
            if abs(total - self.metadata['total_marks']) > 0.5:
                warnings.append(
                    f"Warning: Sum of individual marks ({total}) does not equal total declared marks ({self.metadata['total_marks']})."
                )
        self.validation_warnings = warnings

    def export(self, export_format='json', output_path=None):
        data = {
            'exam_id': f"{self.metadata['course_code']}_{self.metadata['exam_type']}" if self.metadata['course_code'] else 'Unknown_Exam',
            'metadata': self.metadata,
            'warnings': self.validation_warnings,
            'questions': self.parsed_questions,
        }
        if export_format == 'json':
            result = json.dumps(data, indent=2, ensure_ascii=False)
            if output_path:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(result)
            return result
        if export_format == 'csv':
            rows = self.export_rows()
            if output_path:
                with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
                    writer.writeheader()
                    writer.writerows(rows)
            return rows
        if export_format == 'excel':
            if pd is None:
                raise ImportError('pandas is required for Excel export.')
            df = pd.DataFrame(self.export_rows())
            if output_path:
                df.to_excel(output_path, index=False)
            return df
        return data

    def export_rows(self):
        return [
            {
                'Question ID': q['question_id'],
                'Type': q['question_type'],
                'Text': q['question_text'],
                'Marks': q['marks'],
                'Options': json.dumps(q['options'], ensure_ascii=False) if q['options'] else '',
                'Page': q['page_number'],
            }
            for q in self.parsed_questions
        ]


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) > 1:
        parser = ExamParser(sys.argv[1])
        parser.parse()
        print(parser.export(export_format='json'))
        if parser.validation_warnings:
            print('\nValidation Warnings:')
            for warning in parser.validation_warnings:
                print('-', warning)
    else:
        print('Usage: python exam_parser.py <path_to_pdf_or_docx_or_txt>')
