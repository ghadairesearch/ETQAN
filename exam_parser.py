import re
import os
import json
import csv
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None

import pandas as pd


class ExamParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.file_ext = os.path.splitext(filepath)[1].lower()
        self.raw_text = ""
        self.pages = []
        self.metadata = {
            "course_code": None,
            "exam_type": None,
            "total_marks": None,
            "course_name": None
        }
        self.parsed_questions = []

    def parse(self):
        """Executes the full pipeline"""
        self.extract_text()
        self.extract_metadata()
        questions = self.detect_question_boundaries()
        for q in questions:
            q_type = self.classify_question_type(q['text'])
            options = self.extract_mcq_options(q['text']) if q_type == 'MCQ' else {}
            marks = self.extract_marks(q['text'])
            
            q_text = q['text'].strip()
            if q_type == 'MCQ' and options:
                lines = [l.strip() for l in q_text.split('\n') if l.strip()]
                filtered = []
                for line in lines:
                    is_opt = False
                    for k, v in options.items():
                        if line == v or line.startswith(f"{k})") or line.startswith(f"{k}."):
                            is_opt = True
                            break
                    if not is_opt:
                        filtered.append(line)
                q_text = "\n".join(filtered)
            elif q_type == 'True/False':
                lines = [l.strip() for l in q_text.split('\n') if l.strip()]
                filtered = [l for l in lines if l.lower() not in ['true', 'false']]
                q_text = "\n".join(filtered)
                
            self.parsed_questions.append({
                "question_id": f"Q{q['number']}",
                "question_text": q_text,
                "question_type": q_type,
                "options": options,
                "marks": marks,
                "mapped_clo": None,
                "page_number": q['page']
            })
        self.validate_questions()
        return self.parsed_questions

    def extract_text(self):
        """Step 1 & 2: Input Handling and Text Extraction"""
        if self.file_ext == '.pdf':
            if fitz is None:
                raise ImportError("PyMuPDF is required for PDF parsing. Please install it.")
            doc = fitz.open(self.filepath)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text = page.get_text("text")
                self.pages.append({"page": page_num + 1, "text": text})
                self.raw_text += text + "\n\n"
        elif self.file_ext == '.docx':
            try:
                import docx2txt
            except ImportError:
                raise ImportError("docx2txt is required for DOCX parsing. Please install it.")
            text = docx2txt.process(self.filepath)
            self.pages.append({"page": 1, "text": text})
            self.raw_text = text
        else:
            raise ValueError("Unsupported file format. Please use .pdf or .docx")

    def detect_question_boundaries(self):
        """Step 3: Question Boundary Detection"""
        # Matches patterns like '1.', 'Q1:', 'Question 1:', 'س1:', 'السؤال الأول:'
        pattern = re.compile(
            r'(?m)^\s*(?:Q(?:uestion)?\s*[-#.:]?\s*(\d{1,3})|س\s*[-#.:]?\s*(\d{1,3})|السؤال\s+(?:ال)?([^\s:]+)|(\d{1,3})[\).:])',
            re.IGNORECASE
        )
        
        matches = []
        for match in pattern.finditer(self.raw_text):
            number_str = match.group(1) or match.group(2) or match.group(4)
            if number_str:
                number = int(number_str)
            else:
                number = match.group(3) # For Arabic text numbers (e.g. الأول), we'll keep as string or try to parse
                
            matches.append((number, match.start(), match.end()))
        
        if not matches:
            # Fallback: Heuristic extraction for docx files where numbers were stripped
            return self._heuristic_question_detection()

        questions = []
        for i, (num, start, end) in enumerate(matches):
            next_start = matches[i + 1][1] if i + 1 < len(matches) else len(self.raw_text)
            block = self.raw_text[end:next_start].strip()
            
            # Simple page mapping (first occurrence of block in pages)
            page_num = 1
            for p in self.pages:
                if block[:50] in p['text']:
                    page_num = p['page']
                    break
                    
            questions.append({
                "number": num,
                "text": block,
                "page": page_num
            })
            
        return questions

    def _heuristic_question_detection(self):
        """Fallback for unnumbered or stripped text"""
        lines = [line.strip() for line in self.raw_text.split('\n') if line.strip()]
        questions = []
        current_block = []
        q_num = 1
        
        # Keywords to skip (headers)
        skip_words = ['student name', 'section', 'quiz', 'midterm', 'final', 'duration:', 'pg.']
        
        def is_header(block_text):
            lower = block_text.lower()
            return any(w in lower for w in skip_words) and len(block_text) < 150 and '?' not in block_text

        def clean_block(block, is_mcq):
            num_opts = 4 if is_mcq else 2
            opts = block[-num_opts:]
            q_lines = block[:-num_opts]
            
            if not q_lines:
                return block
                
            valid_q_lines = [q_lines[-1]]
            header_words = ['student name', 'student id', 'section:', 'duration', 'notes:', 'quiz', 'midterm', 'final', 'department', 'college', 'university']
            
            for i in range(len(q_lines)-2, -1, -1):
                line = q_lines[i]
                lower = line.lower()
                if any(w in lower for w in header_words):
                    break
                if len(line) <= 2: # Skip weird short lines like page numbers or random characters in headers
                    break
                valid_q_lines.insert(0, line)
                
            return valid_q_lines + opts

        for i, line in enumerate(lines):
            current_block.append(line)
            
            # Check for True/False question
            if len(current_block) >= 3 and current_block[-2] == 'True' and current_block[-1] == 'False':
                cleaned = clean_block(current_block, is_mcq=False)
                text = "\n".join(cleaned)
                if not is_header(text):
                    questions.append({"number": q_num, "text": text, "page": 1})
                    q_num += 1
                    current_block = []
                continue
                
            # Check for MCQ question (Question + 4 options)
            if len(current_block) >= 5:
                # If the last 4 lines are short, and the 5th line back ends with ?, :, or .
                if all(len(opt) < 80 for opt in current_block[-4:]):
                    prompt_line = current_block[-5].strip()
                    if prompt_line.endswith('?') or prompt_line.endswith(':') or prompt_line.endswith('.') or not prompt_line[-1].isalnum():
                        cleaned = clean_block(current_block, is_mcq=True)
                        text = "\n".join(cleaned)
                        if not is_header(text):
                            questions.append({"number": q_num, "text": text, "page": 1})
                            q_num += 1
                            current_block = []
                        continue

        return questions

    def classify_question_type(self, text):
        """Step 4: Question Type Detection"""
        text_lower = text.lower()
        
        # True/False
        if '\ntrue\nfalse' in text_lower or re.search(r'\b(?:true|false|صواب|خطأ|صح|غلط)\b', text_lower):
            return "True/False"
            
        # MCQ rules
        lines = text.strip().split('\n')
        if len(lines) >= 5 and all(len(l) < 80 for l in lines[-4:]):
            return "MCQ"
        if re.search(r'(?m)^\s*(?:A\)|a\)|A\.|a\.|أ\)|أ\.)', text) or re.search(r'\b(?:A|B|C|D)\b', text):
            return "MCQ"
            
        # Fill blank
        if "___" in text or "complete" in text_lower or "أكمل" in text_lower:
            return "Fill in the blank"
            
        # Default to Essay / Short answer
        return "Essay"

    def extract_mcq_options(self, text):
        """Step 5: Option Extraction for MCQs"""
        options = {}
        pattern = re.compile(r'(?m)^\s*([a-dA-Dأبجد])[\).]\s+(.*?)(?=(?:^\s*[a-dA-Dأبجد][\).])|\Z)', re.IGNORECASE | re.DOTALL)
        
        for match in pattern.finditer(text):
            key = match.group(1).upper()
            val = match.group(2).strip()
            options[key] = val
            
        # Fallback if docx2txt stripped the A/B/C/D markers
        if not options:
            lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
            if len(lines) >= 5:
                # Assign the last 4 lines as A, B, C, D
                labels = ['A', 'B', 'C', 'D']
                for i in range(4):
                    options[labels[i]] = lines[-(4-i)]
                    
        return options

    def extract_marks(self, text):
        """Extracts marks/points from question text"""
        match = re.search(r'\[(\d+(?:\.\d+)?)\s*(?:marks?|points?|درجات|درجة)\]', text, re.IGNORECASE)
        if match:
            return float(match.group(1))
        
        match = re.search(r'\((\d+(?:\.\d+)?)\s*(?:marks?|points?|درجات|درجة)\)', text, re.IGNORECASE)
        if match:
            return float(match.group(1))
            
        return 1.0 # Default fallback

    def extract_metadata(self):
        """Step 6: Metadata Extraction"""
        # Course Code (e.g. CS101, SWE 312, عال ٣١٢)
        match_code = re.search(r'\b([A-Za-z]{2,4}\s?\d{3,4})\b', self.raw_text)
        if match_code:
            self.metadata['course_code'] = match_code.group(1).upper()
            
        # Exam Type
        types = ['midterm', 'final', 'quiz', 'assignment', 'نصفي', 'نهائي', 'قصير']
        for t in types:
            if t in self.raw_text.lower():
                self.metadata['exam_type'] = t.capitalize()
                break
                
        # Total Marks
        match_marks = re.search(r'(?:total marks?|score|الدرجة الكلية)[\s:]*(\d+)', self.raw_text, re.IGNORECASE)
        if match_marks:
            self.metadata['total_marks'] = float(match_marks.group(1))

    def validate_questions(self):
        """Step 7: Validation Layer"""
        warnings = []
        
        # Check sequential numbering
        numbers = []
        for q in self.parsed_questions:
            try:
                num = int(str(q['question_id']).replace('Q', ''))
                numbers.append(num)
            except ValueError:
                pass
                
        if numbers and numbers != list(range(min(numbers), max(numbers) + 1)):
            warnings.append("Warning: Question numbers are not strictly sequential.")
            
        # Check MCQs missing options
        for q in self.parsed_questions:
            if q['question_type'] == 'MCQ' and not q['options']:
                warnings.append(f"Warning: {q['question_id']} classified as MCQ but no options extracted.")
                
        # Check total marks
        if self.metadata['total_marks']:
            sum_marks = sum(q.get('marks', 0) for q in self.parsed_questions)
            if abs(sum_marks - self.metadata['total_marks']) > 0.5:
                warnings.append(f"Warning: Sum of individual marks ({sum_marks}) does not equal total declared marks ({self.metadata['total_marks']}).")
                
        self.validation_warnings = warnings

    def export(self, export_format="json", output_path=None):
        """Step 8: Output Format"""
        data = {
            "exam_id": f"{self.metadata['course_code']}_{self.metadata['exam_type']}" if self.metadata['course_code'] else "Unknown_Exam",
            "metadata": self.metadata,
            "warnings": getattr(self, 'validation_warnings', []),
            "questions": self.parsed_questions
        }
        
        if export_format == "json":
            res = json.dumps(data, indent=2, ensure_ascii=False)
            if output_path:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(res)
            return res
            
        elif export_format == "csv" or export_format == "excel":
            df_rows = []
            for q in self.parsed_questions:
                row = {
                    "Question ID": q['question_id'],
                    "Type": q['question_type'],
                    "Text": q['question_text'],
                    "Marks": q['marks'],
                    "Options": json.dumps(q['options'], ensure_ascii=False) if q['options'] else "",
                    "Page": q['page_number']
                }
                df_rows.append(row)
            df = pd.DataFrame(df_rows)
            
            if output_path:
                if export_format == "csv":
                    df.to_csv(output_path, index=False, encoding='utf-8')
                else:
                    df.to_excel(output_path, index=False)
            return df
            
        return data

if __name__ == '__main__':
    # Simple CLI for testing
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        parser = ExamParser(filepath)
        parser.parse()
        print(parser.export(export_format="json"))
        if parser.validation_warnings:
            print("\nValidation Warnings:")
            for w in parser.validation_warnings:
                print("-", w)
    else:
        print("Usage: python exam_parser.py <path_to_pdf_or_docx>")
