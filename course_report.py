import os
import json
import re
import csv
import io
import uuid
import zlib
import tempfile
import zipfile
import base64
import hashlib
import shutil
import subprocess
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import secrets
import difflib
import unicodedata
import smtplib
import ssl
import time
from email.message import EmailMessage
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, Response, send_from_directory, has_request_context, abort
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

try:
    import psycopg2
    import psycopg2.extras
    import psycopg2.errors
except ImportError:
    psycopg2 = None


def load_local_env_file():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if not os.path.exists(env_path):
        return
    with open(env_path, 'r', encoding='utf-8') as env_file:
        for line in env_file:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            key, value = line.split('=', 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


load_local_env_file()


def int_env(name, default):
    try:
        return int(os.environ.get(name) or default)
    except (TypeError, ValueError):
        return default


app = Flask(__name__, template_folder='course_report_templates', static_folder='public')
app.secret_key = os.environ.get('SECRET_KEY') or 'super_secret_key_for_course_report'
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)
app.config['PREFERRED_URL_SCHEME'] = 'https'
UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER') or os.path.join(tempfile.gettempdir(), 'clo_attainment_uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
APP_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_TEXT_CACHE = {}
PDF_TEXT_CACHE_MAX_ITEMS = 64
GEMINI_SPEC_CACHE = {}
GEMINI_SPEC_CACHE_MAX_ITEMS = 32
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
APP_PUBLIC_URL = os.environ.get('APP_PUBLIC_URL', '').strip().rstrip('/')
ADMIN_EMAILS = {
    email.strip().lower()
    for email in os.environ.get('ADMIN_EMAILS', '').split(',')
    if email.strip()
}
SMTP_HOST = os.environ.get('SMTP_HOST', '').strip()
SMTP_PORT = int(os.environ.get('SMTP_PORT') or 587)
SMTP_USERNAME = os.environ.get('SMTP_USERNAME', '').strip()
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', '')
SMTP_FROM_EMAIL = os.environ.get('SMTP_FROM_EMAIL', '').strip() or SMTP_USERNAME
SMTP_FROM_NAME = os.environ.get('SMTP_FROM_NAME', 'ETQAN').strip() or 'ETQAN'
SMTP_USE_SSL = os.environ.get('SMTP_USE_SSL', '').strip().lower() in {'1', 'true', 'yes', 'on'}
SMTP_USE_TLS = os.environ.get('SMTP_USE_TLS', 'true').strip().lower() not in {'0', 'false', 'no', 'off'}
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '').strip()
GEMINI_MODEL = os.environ.get('GEMINI_MODEL', 'gemini-3.5-flash').strip() or 'gemini-3.5-flash'
GEMINI_MAX_INLINE_BYTES = int_env('GEMINI_MAX_INLINE_BYTES', 18 * 1024 * 1024)
GROQ_KEY = os.environ.get('GROQ_KEY', '').strip()
GROQ_MODEL = os.environ.get('GROQ_MODEL', 'qwen/qwen3-32b').strip() or 'qwen/qwen3-32b'
GROQ_API_URL = os.environ.get('GROQ_API_URL', 'https://api.groq.com/openai/v1/chat/completions').strip()
SMTP_TIMEOUT_SECONDS = int(os.environ.get('SMTP_TIMEOUT_SECONDS') or 10)
RESEND_API_KEY = os.environ.get('RESEND_API_KEY', '').strip()
RESEND_FROM_EMAIL = os.environ.get('RESEND_FROM_EMAIL', '').strip() or SMTP_FROM_EMAIL
CONTACT_TO_EMAIL = os.environ.get('CONTACT_TO_EMAIL', '').strip()
LOGIN_FAILURE_WINDOW_MINUTES = 15
LOGIN_EMAIL_FAILURE_LIMIT = 5
LOGIN_IP_FAILURE_LIMIT = 20
LOGIN_DISTINCT_EMAIL_LIMIT = 10
LOGIN_LOCK_MINUTES = 15
LOGIN_ESCALATED_LOCK_MINUTES = 60
COURSE_REPORT_TEMPLATE_PATH_EN = os.path.join(APP_BASE_DIR, 'course_report_templates', 'Course Report Template EN.docx')
COURSE_REPORT_TEMPLATE_PATH_AR = os.path.join(APP_BASE_DIR, 'course_report_templates', 'Course Report Template AR.docx')
COURSE_REPORT_TEMPLATE_PATH = COURSE_REPORT_TEMPLATE_PATH_EN
COURSE_IMPROVEMENT_RECOMMENDATIONS = [
    'Improve student performance in CLOs',
    'Enhance practical/laboratory activities',
    'Update course content',
    'Improve assessment methods',
    'Strengthen industry alignment',
    'Improve prerequisite knowledge',
    'Increase project-based learning',
    'Improve course materials and resources',
    'Enhance critical thinking skills',
    'Enhance teamwork and communication skills',
    'Improve alignment between CLOs and assessments',
    'Increase use of real-world case studies',
    'Reduce content overload',
]
COURSE_IMPROVEMENT_RECOMMENDATION_GROUPS = [
    {
        'heading': 'CLO Attainment and Assessment',
        'items': [
            'Improve student performance in CLOs',
            'Improve assessment methods',
            'Improve alignment between CLOs and assessments',
        ],
    },
    {
        'heading': 'Teaching and Learning Strategies',
        'items': [
            'Increase project-based learning',
            'Increase use of real-world case studies',
        ],
    },
    {
        'heading': 'Practical and Applied Learning',
        'items': [
            'Enhance practical/laboratory activities',
            'Strengthen industry alignment',
        ],
    },
    {
        'heading': 'Course Content and Prerequisites',
        'items': [
            'Update course content',
            'Improve prerequisite knowledge',
            'Reduce content overload',
        ],
    },
    {
        'heading': 'Learning Resources and Skills',
        'items': [
            'Improve course materials and resources',
            'Enhance critical thinking skills',
            'Enhance teamwork and communication skills',
        ],
    },
]
COURSE_REPORT_AR_LABELS = {
    'CLO Attainment and Assessment': 'تحقق نواتج التعلم والتقييم',
    'Teaching and Learning Strategies': 'استراتيجيات التعليم والتعلم',
    'Curriculum and Course Content': 'المنهج ومحتوى المقرر',
    'Learning Resources and Skills': 'مصادر التعلم والمهارات',
    'Improve student performance in CLOs': 'تحسين أداء الطلاب في نواتج التعلم',
    'Enhance practical/laboratory activities': 'تعزيز الأنشطة العملية أو المعملية',
    'Update course content': 'تحديث محتوى المقرر',
    'Improve assessment methods': 'تحسين أساليب التقييم',
    'Strengthen industry alignment': 'تعزيز مواءمة المقرر مع سوق العمل',
    'Improve prerequisite knowledge': 'تحسين المعرفة السابقة المطلوبة',
    'Increase project-based learning': 'زيادة التعلم القائم على المشاريع',
    'Improve course materials and resources': 'تحسين مواد ومصادر المقرر',
    'Enhance critical thinking skills': 'تعزيز مهارات التفكير الناقد',
    'Enhance teamwork and communication skills': 'تعزيز مهارات العمل الجماعي والتواصل',
    'Improve alignment between CLOs and assessments': 'تحسين مواءمة نواتج التعلم مع التقييمات',
    'Increase use of real-world case studies': 'زيادة استخدام دراسات حالة واقعية',
    'Reduce content overload': 'تقليل كثافة محتوى المقرر',
    'Curriculum committee approval': 'موافقة لجنة المناهج',
    'No additional support required': 'لا يتطلب دعمًا إضافيًا',
    'Department approval': 'موافقة القسم',
    'Technical support': 'دعم فني',
    'Revise teaching strategies': 'مراجعة استراتيجيات التدريس',
    'Provide supplementary learning resources': 'توفير مصادر تعلم إضافية',
    'Conduct remedial/support sessions': 'تنفيذ جلسات علاجية أو داعمة',
    'Revise assessment methods': 'مراجعة أساليب التقييم',
    'Update course materials': 'تحديث مواد المقرر',
    'Review CLO-assessment alignment': 'مراجعة مواءمة نواتج التعلم مع التقييم',
    'Increase practical/application activities': 'زيادة الأنشطة العملية أو التطبيقية',
    'Coordinate with the department/course team': 'التنسيق مع القسم أو فريق المقرر',
    'Monitor implementation in the next offering': 'متابعة التنفيذ في الطرح القادم',
    'Time constraints': 'ضيق الوقت',
    'Student absence / low attendance': 'غياب الطلاب أو انخفاض الحضور',
    'Topic covered in another course': 'الموضوع مغطى في مقرر آخر',
    'Course content overload': 'كثافة محتوى المقرر',
    'Equipment unavailability': 'عدم توفر التجهيزات',
    'Topic replaced with a more current topic': 'استبدال الموضوع بموضوع أحدث',
    'Official holidays/suspension of classes': 'إجازات رسمية أو تعليق الدراسة',
    'Technical/laboratory limitations': 'قيود تقنية أو معملية',
    'Prerequisite knowledge gap': 'فجوة في المعرفة السابقة',
    'Provide supplementary learning materials.': 'توفير مواد تعلم إضافية.',
    'Share recorded lectures or demonstrations.': 'مشاركة محاضرات أو عروض مسجلة.',
    'Cover essential concepts through assignments or projects.': 'تغطية المفاهيم الأساسية من خلال تكليفات أو مشاريع.',
    'Schedule additional review/support sessions.': 'جدولة جلسات مراجعة أو دعم إضافية.',
    'Revise topic sequencing in future offerings.': 'مراجعة ترتيب الموضوعات في الطروحات القادمة.',
    'Share lecture recordings and learning resources.': 'مشاركة تسجيلات المحاضرات ومصادر التعلم.',
    'Provide make-up activities or assignments.': 'توفير أنشطة أو تكليفات بديلة.',
    'Offer office hours or support sessions.': 'تقديم ساعات مكتبية أو جلسات دعم.',
    'Implement attendance improvement measures in future offerings.': 'تطبيق إجراءات لتحسين الحضور في الطروحات القادمة.',
    'Coordinate with the related course instructor.': 'التنسيق مع أستاذ المقرر المرتبط.',
    'Recommend redistribution of content across courses.': 'اقتراح إعادة توزيع المحتوى بين المقررات.',
    'Prioritize content directly linked to CLOs.': 'إعطاء الأولوية للمحتوى المرتبط مباشرة بنواتج التعلم.',
    'Move lower-priority content to self-study activities.': 'نقل المحتوى الأقل أولوية إلى أنشطة تعلم ذاتي.',
    'Provide supplementary reading materials.': 'توفير قراءات إضافية.',
    'Use simulations or virtual laboratories.': 'استخدام المحاكاة أو المعامل الافتراضية.',
    'Provide instructor demonstrations.': 'تقديم عروض توضيحية من أستاذ المقرر.',
    'Use alternative equipment where available.': 'استخدام تجهيزات بديلة عند توفرها.',
    'Share recorded practical sessions.': 'مشاركة جلسات عملية مسجلة.',
    'Schedule the activity when resources become available.': 'جدولة النشاط عند توفر الموارد.',
    'Verify alignment with CLOs.': 'التحقق من المواءمة مع نواتج التعلم.',
    'Update course materials and references.': 'تحديث مواد ومراجع المقرر.',
    'Propose formal curriculum updates if the change is permanent.': 'اقتراح تحديث رسمي للمنهج إذا كان التغيير دائمًا.',
    'Provide asynchronous learning materials.': 'توفير مواد تعلم غير متزامنة.',
    'Share recorded lectures.': 'مشاركة محاضرات مسجلة.',
    'Conduct make-up sessions if feasible.': 'تنفيذ جلسات تعويضية إذا أمكن.',
    'Assign guided self-study activities.': 'تكليف الطلاب بأنشطة تعلم ذاتي موجهة.',
    'Adjust future course schedules to account for lost teaching time.': 'تعديل جداول الطروحات القادمة لتعويض وقت التدريس المفقود.',
    'Use virtual labs or software simulations.': 'استخدام معامل افتراضية أو محاكاة برمجية.',
    'Replace practical work with equivalent activities.': 'استبدال العمل العملي بأنشطة مكافئة.',
    'Provide recorded demonstrations.': 'توفير عروض توضيحية مسجلة.',
    'Address infrastructure issues before the next offering.': 'معالجة مشكلات البنية التحتية قبل الطرح القادم.',
    'Deliver remedial sessions.': 'تنفيذ جلسات علاجية.',
    'Provide prerequisite learning resources.': 'توفير مصادر للمعرفة السابقة المطلوبة.',
    'Offer additional tutorials or workshops.': 'تقديم دروس أو ورش إضافية.',
    'Increase formative assessments to monitor readiness.': 'زيادة التقييمات التكوينية لمتابعة الجاهزية.',
    'Review prerequisite requirements and course sequencing.': 'مراجعة المتطلبات السابقة وتسلسل المقررات.',
    'Other': 'أخرى',
}
COURSE_IMPROVEMENT_SUPPORT_OPTIONS = [
    'Curriculum committee approval',
    'No additional support required',
    'Department approval',
    'Technical support',
    'Other',
]
COURSE_IMPROVEMENT_ACTION_OPTIONS = [
    'Revise teaching strategies',
    'Provide supplementary learning resources',
    'Conduct remedial/support sessions',
    'Revise assessment methods',
    'Update course materials',
    'Review CLO-assessment alignment',
    'Increase practical/application activities',
    'Coordinate with the department/course team',
    'Monitor implementation in the next offering',
    'Other',
]
UNCOVERED_TOPIC_REASON_ACTIONS = {
    'Time constraints': [
        'Provide supplementary learning materials.',
        'Share recorded lectures or demonstrations.',
        'Cover essential concepts through assignments or projects.',
        'Schedule additional review/support sessions.',
        'Revise topic sequencing in future offerings.',
    ],
    'Student absence / low attendance': [
        'Share lecture recordings and learning resources.',
        'Provide make-up activities or assignments.',
        'Offer office hours or support sessions.',
        'Implement attendance improvement measures in future offerings.',
    ],
    'Topic covered in another course': [
        'Coordinate with the related course instructor.',
        'Recommend redistribution of content across courses.',
    ],
    'Course content overload': [
        'Prioritize content directly linked to CLOs.',
        'Move lower-priority content to self-study activities.',
        'Provide supplementary reading materials.',
        'Recommend redistribution of content across courses.',
    ],
    'Equipment unavailability': [
        'Use simulations or virtual laboratories.',
        'Provide instructor demonstrations.',
        'Use alternative equipment where available.',
        'Share recorded practical sessions.',
        'Schedule the activity when resources become available.',
    ],
    'Topic replaced with a more current topic': [
        'Verify alignment with CLOs.',
        'Update course materials and references.',
        'Propose formal curriculum updates if the change is permanent.',
    ],
    'Official holidays/suspension of classes': [
        'Provide asynchronous learning materials.',
        'Share recorded lectures.',
        'Conduct make-up sessions if feasible.',
        'Assign guided self-study activities.',
        'Adjust future course schedules to account for lost teaching time.',
    ],
    'Technical/laboratory limitations': [
        'Use virtual labs or software simulations.',
        'Replace practical work with equivalent activities.',
        'Provide recorded demonstrations.',
        'Address infrastructure issues before the next offering.',
    ],
    'Prerequisite knowledge gap': [
        'Deliver remedial sessions.',
        'Provide prerequisite learning resources.',
        'Offer additional tutorials or workshops.',
        'Increase formative assessments to monitor readiness.',
        'Review prerequisite requirements and course sequencing.',
    ],
    'Other': [
        'Other',
    ],
}

def grouped_course_improvement_recommendations():
    groups = []
    for group in COURSE_IMPROVEMENT_RECOMMENDATION_GROUPS:
        items = []
        for recommendation in group['items']:
            if recommendation in COURSE_IMPROVEMENT_RECOMMENDATIONS:
                items.append({
                    'text': recommendation,
                    'index': COURSE_IMPROVEMENT_RECOMMENDATIONS.index(recommendation),
                })
        if items:
            groups.append({'heading': group['heading'], 'items': items})
    return groups

def course_report_label(value):
    value = str(value or '')
    if has_request_context() and get_language() == 'ar':
        return COURSE_REPORT_AR_LABELS.get(value, value)
    return value

def course_report_label_for_language(value, language=None):
    value = str(value or '')
    language = language or (get_export_report_language() if has_request_context() else 'en')
    if language == 'ar':
        return COURSE_REPORT_AR_LABELS.get(value, value)
    return value

def localized_uncovered_reason_actions():
    localized = {}
    for reason, actions in UNCOVERED_TOPIC_REASON_ACTIONS.items():
        localized[reason] = [
            {'value': action, 'label': course_report_label(action)}
            for action in actions
        ]
    return localized
UNIVERSITY_IDENTITY_PATH = os.path.join(APP_BASE_DIR, 'university_identity.json')
UNIVERSITY_LOGO_FOLDER = os.path.join(APP_BASE_DIR, 'public', 'university_logos')
ORG_LOGO_FOLDER = os.environ.get('ORG_LOGO_FOLDER') or os.path.join(UPLOAD_FOLDER, 'organization_logos')
os.makedirs(UNIVERSITY_LOGO_FOLDER, exist_ok=True)
os.makedirs(ORG_LOGO_FOLDER, exist_ok=True)
FREE_REPORT_LIMIT = 1
ACADEMIC_REPORT_LIMIT_PER_YEAR = 12
ACADEMIC_COURSE_LIMIT = 10
PAY_AS_YOU_GO_PRICE_SAR = 19
ACADEMIC_SUBSCRIPTION_PRICE_SAR = 99
PROFESSIONAL_SUBSCRIPTION_PRICE_SAR = 199

UNIVERSITY_CHOICES = [
    'Umm Al-Qura University',
    'Islamic University of Madinah',
    'Imam Mohammad Ibn Saud Islamic University',
    'King Saud University',
    'King Abdulaziz University',
    'King Faisal University',
    'King Khalid University',
    'Qassim University',
    'Taibah University',
    'Taif University',
    'University of Hail',
    'Jazan University',
    'Al Jouf University',
    'Al Baha University',
    'Tabuk University',
    'Najran University',
    'Northern Border University',
    'Princess Nourah bint Abdulrahman University',
    'King Saud bin Abdulaziz University for Health Sciences',
    'Imam Abdulrahman Bin Faisal University',
    'Prince Sattam bin Abdulaziz University',
    'Shaqra University',
    'Majmaah University',
    'Saudi Electronic University',
    'University of Jeddah',
    'University of Bisha',
    'King Abdullah University of Science and Technology',
    'King Fahd University of Petroleum & Minerals',
    'Prince Sultan University',
    'Effat University',
    'Dar Al-Hekma University',
    'Alfaisal University',
    'Arab Open University',
]

UNIVERSITY_ARABIC_NAMES = {
    'Umm Al-Qura University': 'جامعة أم القرى',
    'Islamic University of Madinah': 'الجامعة الإسلامية بالمدينة المنورة',
    'Imam Mohammad Ibn Saud Islamic University': 'جامعة الإمام محمد بن سعود الإسلامية',
    'King Saud University': 'جامعة الملك سعود',
    'King Abdulaziz University': 'جامعة الملك عبدالعزيز',
    'King Faisal University': 'جامعة الملك فيصل',
    'King Khalid University': 'جامعة الملك خالد',
    'Qassim University': 'جامعة القصيم',
    'Taibah University': 'جامعة طيبة',
    'Taif University': 'جامعة الطائف',
    'University of Hail': 'جامعة حائل',
    'Jazan University': 'جامعة جازان',
    'Al Jouf University': 'جامعة الجوف',
    'Al Baha University': 'جامعة الباحة',
    'Tabuk University': 'جامعة تبوك',
    'Najran University': 'جامعة نجران',
    'Northern Border University': 'جامعة الحدود الشمالية',
    'Princess Nourah bint Abdulrahman University': 'جامعة الأميرة نورة بنت عبدالرحمن',
    'King Saud bin Abdulaziz University for Health Sciences': 'جامعة الملك سعود بن عبدالعزيز للعلوم الصحية',
    'Imam Abdulrahman Bin Faisal University': 'جامعة الإمام عبدالرحمن بن فيصل',
    'Prince Sattam bin Abdulaziz University': 'جامعة الأمير سطام بن عبدالعزيز',
    'Shaqra University': 'جامعة شقراء',
    'Majmaah University': 'جامعة المجمعة',
    'Saudi Electronic University': 'الجامعة السعودية الإلكترونية',
    'University of Jeddah': 'جامعة جدة',
    'University of Bisha': 'جامعة بيشة',
    'King Abdullah University of Science and Technology': 'جامعة الملك عبدالله للعلوم والتقنية',
    'King Fahd University of Petroleum & Minerals': 'جامعة الملك فهد للبترول والمعادن',
    'Prince Sultan University': 'جامعة الأمير سلطان',
    'Effat University': 'جامعة عفت',
    'Dar Al-Hekma University': 'جامعة دار الحكمة',
    'Alfaisal University': 'جامعة الفيصل',
    'Arab Open University': 'الجامعة العربية المفتوحة',
}

UNIVERSITY_ENGLISH_BY_ARABIC = {
    arabic_name: english_name
    for english_name, arabic_name in UNIVERSITY_ARABIC_NAMES.items()
}

UNIVERSITY_COLOR_PRESETS = {
    'Umm Al-Qura University': '#006871',
    'Islamic University of Madinah': '#005A85',
    'Imam Mohammad Ibn Saud Islamic University': '#045D84',
    'King Saud University': '#008DC3',
    'King Abdulaziz University': '#006B54',
    'King Faisal University': '#28685B',
    'King Khalid University': '#808080',
    'Qassim University': '#00519b',
    'Taibah University': '#0A8E6E',
    'Taif University': '#D2AA4D',
    'University of Hail': '#114677',
    'Jazan University': '#2B5872',
    'Al Jouf University': '#12403C',
    'Al Baha University': '#204080',
    'Tabuk University': '#009050',
    'Najran University': '#305890',
    'Northern Border University': '#236432',
    'Princess Nourah bint Abdulrahman University': '#007580',
    'King Saud bin Abdulaziz University for Health Sciences': '#006840',
    'Imam Abdulrahman Bin Faisal University': '#304060',
    'Prince Sattam bin Abdulaziz University': '#3c7974',
    'Shaqra University': '#006048',
    'Majmaah University': '#51692C',
    'Saudi Electronic University': '#111E4D',
    'University of Jeddah': '#004F6E',
    'University of Bisha': '#204078',
    'King Abdullah University of Science and Technology': '#786858',
    'King Fahd University of Petroleum & Minerals': '#337A42',
    'Prince Sultan University': '#104070',
    'Effat University': '#181848',
    'Dar Al-Hekma University': '#5F1100',
    'Alfaisal University': '#17365D',
    'Arab Open University': '#183058',
}

LEGACY_UNIVERSITY_COLOR_DEFAULTS = {
    'Islamic University of Madinah': {'#0F6B4F', '#1C8354'},
    'Imam Mohammad Ibn Saud Islamic University': {'#1F4E79', '#337AB7'},
    'King Saud University': {'#00558C', '#063E27'},
    'King Abdulaziz University': {'#006B54'},
    'King Faisal University': {'#0B6F4A', '#0D6EFD'},
    'King Khalid University': {'#005B7F', '#1B8354'},
    'Northern Border University': {'#234E70', '#9DB73B', '#195F28'},
}

EN_TRANSLATIONS = {
    'app.title': 'ETQAN',
    'app.subtitle': 'Educational Transformation & Quality ANalytics',
    'nav.services': 'Home',
    'nav.account': 'Account Info',
    'nav.my_reports': 'My Reports',
    'nav.my_courses': 'My Courses',
    'nav.my_exams': 'My Exams',
    'nav.home': 'Home',
    'nav.faq': 'FAQ',
    'nav.contact': 'Contact Us',
    'nav.organization': 'Organization Identity',
    'nav.settings': 'Settings',
    'nav.billing': 'Subscriptions',
    'nav.logout': 'Logout',
    'nav.login': 'Login',
    'nav.create_account': 'Create Account',
    'nav.language': 'Language',
    'nav.english': 'English',
    'nav.arabic': 'العربية',
    'home.title': 'Select a Service',
    'home.description': 'Choose the type of academic quality report or analysis you want to create.',
    'home.course_level': 'Course Level Services',
    'home.course_level_description': 'Workflows for course CLO analysis, attainment evidence, and course reports.',
    'home.program_level': 'Program Level Services',
    'home.program_level_description': 'Workflows that combine course evidence into program learning outcome analysis.',
    'home.reviewer_level': 'Reviewer Services',
    'home.reviewer_course_report_title': 'Course Report Review',
    'home.reviewer_course_report_description': 'Review course reports against NCAAA requirements and flag missing or inconsistent sections.',
    'home.reviewer_course_spec_title': 'Course Specification Review',
    'home.reviewer_course_spec_description': 'Review course specifications for CLOs, topics, assessment alignment, and NCAAA completeness.',
    'home.reviewer_program_spec_title': 'Program Specification Review',
    'home.reviewer_program_spec_description': 'Review program specifications for PLOs, curriculum structure, assessment methods, and NCAAA completeness.',
    'home.reviewer_clo_mapping_title': 'CLO Mapping Review',
    'home.reviewer_clo_mapping_description': 'Check the alignment between assessment questions, CLOs, targets, and reported attainment results.',
    'home.reviewer_evidence_title': 'Evidence Review',
    'home.reviewer_evidence_description': 'Inspect supporting files and evidence packages before submission or accreditation review.',
    'home.add_program_title': 'Add Program',
    'home.add_program_description': 'Add a program to generate reports.',
    'home.requires_program': 'Requires at least one program.',
    'home.inactive': 'Inactive',
    'home.clo_title': 'CLO Attainment Analysis',
    'home.clo_description': 'Measure CLO attainment from assessment data, identify performance gaps, and generate evidence-based reports.',
    'home.add_course_title': 'Add Course',
    'home.add_course_description': 'Create a course to generate reports.',
    'home.requires_course': 'Requires at least one course.',
    'home.question_mapping_title': 'Question CLO Mapping',
    'home.question_mapping_description': 'Upload an exam paper to map each question to the related CLOs.',
    'home.question_mapping_extract': 'Extract Questions',
    'question_mapping.review_title': 'Review Extracted Questions',
    'question_mapping.review_description': 'Edit the extracted questions, add missing questions if needed, then map them to CLOs.',
    'question_mapping.add_question': 'Add Question',
    'question_mapping.map_to_clos': 'Continue',
    'question_mapping.question_text': 'Question text',
    'question_mapping.question_type': 'Question type',
    'question_mapping.paper_clo': 'CLO from exam paper',
    'question_mapping.paper_clo_help': '',
    'question_mapping.step2_title': 'Step 2: Question Extraction and Explicit Mapping',
    'question_mapping.step2_description': 'Review extracted questions, detected question types, and any CLOs explicitly mentioned in the exam paper.',
    'question_mapping.total_questions': 'Total Questions',
    'question_mapping.auto_mapped_questions': 'Mapped Questions',
    'question_mapping.ai_required_questions': 'Requires Mapping',
    'question_mapping.question_number': 'Q. No.',
    'question_mapping.explicit_clo': 'CLO',
    'question_mapping.status': 'Status',
    'question_mapping.mapped_automatically': 'Mapped Automatically',
    'question_mapping.requires_ai_mapping': 'Requires Mapping',
    'question_mapping.all_mapped_success': 'All questions were mapped successfully.',
    'question_mapping.final_review': 'Step 4: Final Review',
    'question_mapping.continue_ai_mapping': 'Next',
    'question_mapping.step3_title': 'Step 3: Map Questions to CLOs',
    'question_mapping.step3_description': 'Review AI recommendations for questions without an explicitly mentioned CLO.',
    'question_mapping.suggested_clo': 'Suggested CLO',
    'question_mapping.ai_suggestions': 'AI Suggestions',
    'question_mapping.high_conf': 'High',
    'question_mapping.medium_conf': 'Medium',
    'question_mapping.low_conf': 'Low',
    'question_mapping.alt_suggestions': 'Alternative Suggestions:',
    'question_mapping.no_ai_suggestions': 'No AI suggestions available.',
    'question_mapping.ai_diagnostics_title': 'AI Mapping Diagnostics',
    'question_mapping.ai_diagnostics_help': 'This shows whether Gemini, Qwen, or local matching produced the mapping suggestions.',
    'question_mapping.no_suggestion_reason': 'No suggestion was produced for this question. See the diagnostics above for the provider status.',
    'question_mapping.final_selection': 'Final CLO Selection',
    'question_mapping.add_clo': 'Add CLO',
    'question_mapping.confidence_score': 'Confidence Score',
    'question_mapping.select_clo': 'Select CLO',
    'question_mapping.choose_clo': 'Choose a CLO...',
    'question_mapping.all_course_clos': 'All Course CLOs',
    'question_mapping.link_title': 'Map Questions to CLOs',
    'question_mapping.link_description': 'Review each question number and the related CLO suggested by ETQAN.',
    'question_mapping.related_clo': 'Related CLO',
    'question_mapping.save_mapping': 'Save',
    'question_mapping.back': 'Back',
    'question_mapping.review_saved': 'Saved successfully.',
    'question_mapping.no_questions': 'No questions were extracted. Add at least one question before mapping.',
    'question_mapping.select_at_least_one': 'Please select at least one CLO for:',
    'exams.title': 'My Exams',
    'exams.empty': 'No saved exams yet.',
    'exams.saved': 'Report saved successfully.',
    'exams.exam': 'Exam',
    'exams.course': 'Course',
    'exams.questions': 'Questions',
    'exams.created': 'Created',
    'home.assessment_balance_title': 'Assessment Balance Check',
    'home.assessment_balance_description': 'Review assessment coverage, score distribution, and balance across learning outcomes before reporting.',
    'home.plo_title': 'PLO Attainment Analysis',
    'home.plo_description': 'Measure PLO attainment by aggregating course-level evidence and generate PLO performance reports.',
    'home.course_report_title': 'Course Report',
    'home.course_report_description': 'Generate an NCAAA-aligned course report using CLO attainment results and course performance data.',
    'home.no_courses_message': 'Welcome to ETQAN. Start by adding your first course.',
    'home.add_course': 'Add Course',
    'course_report.select_description': 'Select a course to generate its course report.',
    'course_report.need_clo_report': 'To create a course report, you need a CLO attainment report first.',
    'course_report.create_clo_prompt': 'Would you like to create one now?',
    'course_report.create_clo_report': 'Create CLO Report',
    'course_report.associated_reports': 'Associated CLO Attainment Reports',
    'course_report.associated_reports_help': 'Select one or more CLO attainment reports to use for this course report.',
    'course_report.no_associated_reports': 'There is no associated CLO attainment report for this course. You have to create one before creating the course report.',
    'course_report.use_report': 'Use this report',
    'course_report.create_report': 'Next',
    'course_report.preview_title': 'Course Report Preview',
    'course_report.student_results_comment_edit': 'Comment on Student Grades',
    'course_report.preview_description': 'Review the completed course report information, then export the Word report.',
    'course_report.saved_automatically': 'تم الحفظ تلقائياً.',
    'course_report.next': 'Next',
    'course_report.export_word': 'Export Word',
    'course_report.export_pdf': 'Export PDF',
    'course_report.report_details': 'Report Details',
    'course_report.grade_distribution': 'Final Grade Distribution',
    'course_report.clo_summary': 'CLO Assessment Results',
    'course_report.recommendations': 'Recommendations',
    'course_report.select_one_report': 'Select at least one CLO attainment report.',
    'course_report.selected_reports': 'Selected CLO Attainment Reports',
    'course_report.continue': 'Next',
    'home.open_service': 'Start',
    'service.coming_soon': 'This service page is ready. Full workflow tools will be added here.',
    'service.back_home': 'Back to Services',
    'footer.copyright': 'ETQAN © 2026. All rights reserved.',
    'footer.privacy': 'Privacy Policy',
    'footer.faq': 'FAQ',
    'footer.contact': 'Contact Us',
    'auth.email': 'Email',
    'auth.password': 'Password',
    'auth.university': 'University',
    'auth.college_optional': 'College',
    'auth.department_optional': 'Department',
    'auth.select_university': 'Select a university...',
    'auth.other_university': 'Other',
    'auth.university_placeholder': 'Enter your university name',
    'auth.need_account': 'Need an account?',
    'auth.create_one': 'Create one',
    'auth.have_account': 'Already have an account?',
    'auth.login': 'Login',
    'auth.invalid_login': 'Invalid email or password.',
    'auth.register_title': 'Create Account',
    'auth.login_title': 'Login',
    'auth.register_button': 'Create Account',
    'auth.login_button': 'Login',
    'auth.forgot_password': 'Forgot password?',
    'auth.forgot_title': 'Reset Password',
    'auth.send_reset': 'Send Reset Email',
    'auth.reset_sent': 'If the email exists, a password reset link has been sent.',
    'auth.reset_email_unconfigured': 'Email delivery is not configured yet. Please set SMTP settings before sending reset emails.',
    'auth.reset_email_failed': 'We could not send the reset email right now. Please try again later.',
    'auth.login_locked_email': 'Too many failed login attempts for this email. Please try again later.',
    'auth.login_locked_ip': 'Too many failed login attempts from this network. Please try again later.',
    'auth.new_password': 'New Password',
    'auth.confirm_password': 'Confirm Password',
    'auth.reset_button': 'Update Password',
    'auth.back_login': 'Back to Login',
    'org.title': 'Organization Identity',
    'org.description': 'This data is used as the visual identity for exported reports.',
    'org.university': 'University',
    'org.other_university': 'Other',
    'org.university_placeholder': 'Enter your university name',
    'org.department': 'Department',
    'org.department_placeholder': '',
    'org.logo': 'To replace the current logo, please attach the new logo.',
    'org.official_website': 'Official website',
    'org.used_logo': 'Logo used',
    'org.logo_alt': 'Organization logo',
    'org.no_logo': 'No logo selected',
    'org.primary_color': 'Primary Color',
    'org.secondary_color': 'Secondary Color',
    'org.tertiary_color': 'Third Color',
    'org.current_logo': 'Current logo:',
    'org.save': 'Save',
    'org.saved': 'تم الحفظ بنجاح.',
    'account.danger_title': 'Delete Account',
    'account.danger_text': 'Deleting your account will permanently remove saved reports, organization identity, and subscription data from ETQAN. You will also lose any remaining report credits in your account, and they cannot be restored after deletion is completed.',
    'account.confirm_label': 'Type DELETE to confirm',
    'account.delete_button': 'Delete My Account',
    'account.delete_confirm_js': 'This will permanently delete your account and saved reports. Continue?',
    'account.delete_invalid': 'Account deletion was not confirmed.',
    'account.deleted': 'Your account has been deleted.',
    'account.delete_help': 'This action cannot be undone.',
    'account.title': 'Account Info',
    'account.description': 'Manage your login account and account-level actions.',
    'account.email': 'Email',
    'account.university': 'Institution',
    'account.college': 'College',
    'account.department': 'Department',
    'account.created_at': 'Created',
    'account.plan': 'Plan',
    'account.edit': 'Edit',
    'account.cancel': 'Cancel',
    'account.save': 'Save',
    'account.saved': 'تم الحفظ بنجاح.',
    'settings.title': 'Settings',
    'settings.description': '',
    'settings.account_title': 'Account Info',
    'settings.account_description': 'Account actions such as editing information and deletion.',
    'settings.organization_title': 'Organization Identity',
    'settings.organization_description': 'Manage the identity used in exported reports.',
    'settings.report_title': 'Report Settings',
    'settings.report_description': 'Set exported report preferences.',
    'settings.open': 'Open',
    'report_settings.title': 'Report Settings',
    'report_settings.description': '',
    'report_settings.language': 'Exported Report Language',
    'report_settings.language_help': '',
    'report_settings.same_as_interface': 'Same as interface language',
    'report_settings.english': 'English',
    'report_settings.arabic': 'Arabic',
    'report_settings.save': 'Save',
    'report_settings.saved': 'تم الحفظ بنجاح.',
    'courses.title': 'My Courses',
    'courses.description': 'Save course information once, then reuse it when creating reports.',
    'courses.add_new': 'Add Course',
    'courses.back_to_courses': 'Back to My Courses',
    'courses.add_title': 'Add Course',
    'courses.edit_title': 'Edit Course',
    'courses.edit_help': 'Update the course information or upload a new specification.',
    'courses.edit': 'Edit',
    'courses.method_help': 'Upload a course specification file, and ETQAN will extract the course information.',
    'courses.upload_method': 'Enter Course Specification',
    'courses.upload_help': 'ETQAN will extract the course name, code, and CLOs from the PDF when readable.',
    'courses.manual_prompt': 'Or would you like to enter course information manually?',
    'courses.manual_method': 'Course Information',
    'courses.manual_help': 'Use this when you do not have a course specification file or want to edit the information extracted from the specification.',
    'courses.course_name': 'Course Name',
    'courses.course_code': 'Course Code',
    'courses.college': 'College',
    'courses.program': 'Program',
    'courses.department': 'Department',
    'courses.clos': 'CLOs',
    'courses.clo': 'CLO',
    'courses.associated_plo': 'Associated PLO',
    'courses.actions': 'Actions',
    'courses.add_clo_row': 'Add CLO',
    'courses.remove_clo': 'Remove',
    'courses.topics': 'Course Topics',
    'courses.topics_placeholder': 'Enter course topics, one per line',
    'courses.optional_override': 'Optional override',
    'courses.spec_file': 'Course Specification',
    'courses.spec_file_help': 'PDF or Word files are supported.',
    'courses.extract': 'Extract Information',
    'courses.extracting': 'Extracting...',
    'courses.extract_missing': 'Upload a course specification first.',
    'courses.extracted': 'Course information extracted. Review it, then click Save Course.',
    'courses.extraction_method_prefix': 'Extraction method:',
    'courses.extraction_method_gemini': 'Gemini Flash',
    'courses.extraction_method_qwen': 'Qwen via Groq',
    'courses.extraction_method_llama': 'Llama via Groq',
    'courses.extraction_method_groq': 'Groq AI',
    'courses.extraction_method_local': 'Local text/OCR',
    'courses.save': 'Save Course',
    'courses.empty': 'No saved courses yet. Add a course from the home page.',
    'courses.delete': 'Delete',
    'courses.delete_confirm': 'Delete this saved course?',
    'courses.saved': 'تم الحفظ بنجاح.',
    'courses.deleted': 'Course deleted.',
    'programs.login_required': 'Please login to manage your programs.',
    'programs.add_title': 'Add Program',
    'programs.program_name': 'Program Name',
    'programs.program_code': 'Program Code',
    'programs.college': 'College',
    'programs.department': 'Department',
    'programs.plos': 'PLOs',
    'programs.plos_help': 'Paste one PLO per line, for example: K1 Demonstrate knowledge, S1 Apply skills, V1 Show values.',
    'programs.save': 'Save Program',
    'programs.invalid': 'Enter a program name and at least one PLO.',
    'courses.invalid': 'Enter a course name, course topics, and at least one CLO, or upload a readable course specification file.',
    'courses.limit': 'Your plan does not allow saving more courses.',
    'courses.login_required': 'Please login to manage saved courses.',
    'index.title': 'Course Information',
    'index.course': 'Course',
    'index.select_course': 'Select a course...',
    'index.course_not_found': 'Course not found?',
    'index.course_not_found_my_courses': 'Course not found? Add it through Add Course service.',
    'index.course_spec_title': 'Upload course specification PDF',
    'index.inline_course_title': 'Add Course Information',
    'index.inline_course_help': 'If the course is not in the list, enter its name and upload a course specification PDF or paste the CLOs below.',
    'index.inline_course_name': 'Course Name',
    'index.inline_course_name_placeholder': 'e.g. Data Structures (DS2206)',
    'index.inline_spec_file': 'Course Specification',
    'index.inline_clos': 'Paste CLOs',
    'index.inline_clos_placeholder': 'Paste one CLO per line, for example:\n1.1 Identify basic concepts\n2.1 Apply appropriate methods\n3.1 Demonstrate professional values',
    'index.target_label': 'Target Level per CLO',
    'index.target_help': "Set the minimum percentage a student needs to achieve for each learning outcome.",
    'index.clo': 'Course Learning Outcome (CLO)',
    'index.target_level': 'Target Level (%)',
    'domain.knowledge': '1.0 Knowledge',
    'domain.skills': '2.0 Skills',
    'domain.values': '3.0 Values',
    'domain.other': 'Other',
    'index.select_course_populate': 'Select a course to populate...',
    'index.no_clos_category': 'No CLOs detected for this category.',
    'index.assessment_files': 'Assessment Files',
    'index.assessment_help': 'Upload at least one grades file.',
    'index.type': 'Type',
    'index.grades_file': 'Grades File',
    'index.exam_paper': 'Exam Paper',
    'assessment.quiz': 'Quiz',
    'assessment.assignment': 'Assignment',
    'assessment.midterm': 'Midterm',
    'assessment.final': 'Final',
    'assessment.project': 'Project',
    'assessment.other': 'Other',
    'index.remove': 'Remove',
    'index.add_file': 'Add File',
    'index.next_mapping': 'Next',
    'index.error_course': 'Please select a course or add course information before continuing.',
    'index.error_file': 'Upload at least one grades file.',
    'spec.title': 'Upload Course Specification',
    'spec.back': 'Back',
    'spec.description': 'Upload a course specification PDF to extract the course name, course number, and CLOs. Arabic text-based and OCR-supported scanned PDFs are supported.',
    'spec.file': 'Course Specification',
    'spec.extract': 'Extract Course Information',
    'spec.preview': 'Extracted Preview',
    'spec.list_name': 'Course List Name',
    'spec.course_name': 'Course Name',
    'spec.course_number': 'Course Number',
    'spec.clos': 'CLOs',
    'spec.no_knowledge': 'No Knowledge CLOs detected.',
    'spec.no_skills': 'No Skills CLOs detected.',
    'spec.no_values': 'No Values CLOs detected.',
    'spec.add_course': 'Add Course to List',
    'spec.upload_another': 'Upload Another PDF',
    'spec.cannot_add': 'The course cannot be added until the course name and CLO list are detected.',
    'mapping.title': 'Map Questions to CLOs',
    'mapping.course': 'Course:',
    'mapping.detected': 'Detected:',
    'mapping.description': 'Assign one or more CLOs to each question.',
    'mapping.method_title': 'Choose Mapping Method',
    'mapping.method_description': 'Select how you want to map assessment questions to CLOs.',
    'mapping.method_choice': 'Mapping method',
    'mapping.method_manual_title': 'Manual mapping',
    'mapping.method_manual_description': 'Select CLOs for each question manually.',
    'mapping.method_ai_title': 'AI-assisted mapping',
    'mapping.method_ai_description': 'Select a previously generated mapping report to reuse its question-to-CLO mappings.',
    'mapping.method_exam_paper_help': 'Upload the exam paper for the assessment questions in the grade files.',
    'mapping.method_continue': 'Next',
    'mapping.exam_paper': 'Exam paper:',
    'mapping.questions': 'questions',
    'mapping.students': 'students',
    'mapping.question': 'Q. No.',
    'mapping.question_prefix': 'Q',
    'mapping.max_score': 'Max Score',
    'mapping.multi_help': 'Hold Ctrl to select multiple CLOs.',
    'mapping.selected_clo': 'Selected CLO ID',
    'mapping.no_questions': 'No questions were detected. Please go back and upload a valid grades file.',
    'mapping.back': 'Back',
    'mapping.calculate': 'Next',
    'detected.title': 'Detected Course Report Structure',
    'detected.source': 'Source:',
    'detected.questions': 'Questions',
    'detected.students': 'Students',
    'detected.detection': 'Detection',
    'detected.save': 'Save CLO Selection',
    'detected.back_upload': 'Back to Upload',
    'detected.no_table': 'No question table could be detected from this file.',
    'detected.text_sample': 'View extracted text sample',
    'detected.suggestion_source': 'Suggestion source:',
    'detected.source_gemini': 'Gemini Flash',
    'detected.source_local': 'Local semantic matching',
    'detected.mapping_used_gemini': 'Question-CLO mapping used Gemini Flash.',
    'detected.mapping_used_qwen': 'Question-CLO mapping used Qwen via Groq.',
    'detected.mapping_used_local': 'Question-CLO mapping used local semantic matching.',
    'detected.question_text': 'Question text',
    'results.title': 'CLO Attainment Report',
    'results.course': 'Course:',
    'results.total_students': 'Total Students Evaluated:',
    'results.export_csv': 'Export CSV',
    'results.export_pdf': 'Export PDF',
    'results.export_course_report': 'Export Course Report DOCX',
    'results.save_course_report': 'Save Course Report',
    'results.export_word': 'Export Word',
    'results.course_report_inputs': 'Course Report Inputs',
    'results.course_report_optional_details': 'Optional Course Report Details',
    'results.course_instructor': 'Course Instructor',
    'results.course_coordinator': 'Course Coordinator',
    'results.course_location': 'Course Location',
    'results.main_campus': 'Main campus',
    'results.branch': 'Branch',
    'results.branch_name': 'Branch name',
    'results.number_of_sections': 'Number of Sections',
    'results.students_started': 'Students who started the course',
    'results.students_completed': 'Students who completed the course',
    'results.topics_covered': 'Have you covered all course topics?',
    'results.topics_covered_yes': 'Yes, all topics were covered',
    'results.topics_covered_no': 'No, some topics were not covered',
    'results.uncovered_topics': 'Topics not covered',
    'results.uncovered_topics_help': 'Select the topics from the course specification that were not covered.',
    'results.no_extracted_topics': 'No extracted course topics are available for this course.',
    'results.uncovered_reason': 'Reason for Not Covering/discrepancies',
    'results.uncovered_reason_select': 'Select a reason',
    'results.reason_time_constraints': 'Time constraints',
    'results.reason_absence': 'Student absence / low attendance',
    'results.reason_other_course': 'Topic covered in another course',
    'results.reason_overload': 'Course content overload',
    'results.reason_equipment': 'Equipment unavailability',
    'results.reason_replaced': 'Topic replaced with a more current topic',
    'results.reason_holidays': 'Official holidays/suspension of classes',
    'results.reason_lab_limitations': 'Technical/laboratory limitations',
    'results.reason_prerequisite_gap': 'Prerequisite knowledge gap',
    'results.reason_other': 'Other',
    'results.reason_other_explanation': 'Other explanation',
    'results.please_explain': 'Please explain',
    'results.uncovered_impact': 'Extent of their Impact on Learning Outcomes',
    'results.impact_none': 'None',
    'results.impact_low': 'Low',
    'results.impact_medium': 'Medium',
    'results.impact_high': 'High',
    'results.uncovered_action': 'Compensating Action',
    'results.uncovered_action_select': 'Select an action',
        'results.all_topics_covered': 'All topics were successfully covered.',
    'results.action_supplementary_resources': 'Supplementary learning resources provided',
    'results.action_none_required': 'No action required',
    'results.action_other': 'Other',
    'results.action_other_explanation': 'Other action explanation',
    'results.final_grades_file': 'Final Grades File',
    'results.final_grades_help': 'Upload CSV, Excel, or PDF with final grades or final numeric scores. ETQAN will count A+, A, B+, B, and so on.',
    'results.course_improvement_plan': 'Course Improvement Plan',
    'results.course_improvement_help': 'Select recommendations to include under the Course Improvement Plan.',
    'results.recommendation': 'Recommendation',
    'results.actions_needed': 'Actions Needed',
    'results.needed_support': 'Needed Support',
    'results.support_other_explanation': 'Other support explanation',
    'results.complete_org': 'Complete Organization Profile',
    'results.login_export': 'Login to Export',
    'results.create_account': 'Create Account',
    'results.org_identity': 'Organization Visual Identity',
    'results.export_requires_account': "Exporting reports requires an account so the report can use your organization's visual identity.",
    'results.clo_definitions': 'CLO Definitions',
    'results.domain': 'Domain',
    'results.clo': 'Code',
    'results.wording': 'CLOs',
    'results.summary': 'CLO Achievement Summary',
    'results.mapped_questions': 'Mapped Questions',
    'results.max_possible': 'Max Possible Score',
    'results.target': 'Target',
    'results.students_achieved': 'No of Students Achieved',
    'results.code': 'Code',
    'results.max_score': 'Total Score',
    'results.achievement': 'Achievement',
    'results.achievement_pct': 'Achievement %',
    'results.mapped_question_count': 'mapped question',
    'results.student_achievement': 'Student CLO Achievement',
    'results.student_id': 'Student ID',
    'results.achieved': 'Achieved',
    'results.not_achieved': 'Not Achieved',
    'results.back_mapping': 'Back',
    'results.start_over': 'Start Over',
    'results.back_reports': 'Back to My Reports',
    'results.create_new': 'Create New Report',
    'results.no_mappings': 'No mappings were provided. Please ensure you mapped at least one question to a CLO.',
    'history.title': 'My Reports',
    'history.report': 'Report Title',
    'history.course': 'Course',
    'history.type': 'Report Type',
    'history.created': 'Created At',
    'history.open': 'Open',
    'history.rename': 'Rename',
    'history.rename_placeholder': 'Report name',
    'history.rename_save': 'Save',
    'history.renamed': 'Report renamed.',
    'history.rename_invalid': 'Enter a report name.',
    'history.rename_duplicate': 'A report with this name already exists for this course.',
    'history.delete': 'Delete',
    'history.delete_confirm': 'Are you sure you want to delete this report? This action cannot be undone.',
    'history.deleted': 'Report deleted.',
    'history.empty': 'No saved reports yet. Create a report from the home page.',
    'billing.title': 'Subscriptions',
    'billing.description': '',
    'billing.free_status': 'Free reports used',
    'billing.credits': 'Available report credits',
    'billing.plan': 'Current plan',
    'billing.yearly_active': 'Annual',
    'billing.academic_active': 'Academic',
    'billing.professional_active': 'Professional',
    'billing.free': 'Free',
    'billing.payg_title': 'Pay as You Go',
    'billing.payg_price': '19 SAR per report',
    'billing.payg_unit': 'SAR per report',
    'billing.payg_description': 'Perfect for occasional reporting needs. Each purchase adds one report credit to your account.',
    'billing.payg_button': 'Choose Plan',
    'billing.academic_title': 'Academic',
    'billing.academic_price': '99 SAR/year',
    'billing.academic_description': 'Designed for individual faculty members who prepare reports throughout the academic year. Includes up to 12 reports annually.',
    'billing.academic_button': 'Choose Plan',
    'billing.professional_title': 'Professional',
    'billing.professional_price': '199 SAR/year',
    'billing.professional_description': 'Ideal for power users who need unlimited reports, dashboards, accreditation evidence management, and advanced reporting tools.',
    'billing.professional_button': 'Choose Plan',
    'billing.yearly_unit': 'SAR/year',
    'billing.university_title': 'University Subscription',
    'billing.university_price': 'Contact us',
    'billing.university_description': 'Built for departments, colleges, and universities requiring multiple user accounts, centralized reporting, and institution-wide quality management.',
    'billing.university_button': 'Contact us',
    'billing.local_notice': '',
    'billing.back': 'Back',
    'billing.limit_message': 'You have used your free report. Please choose a billing option to create another report.',
    'billing.payg_added': 'One pay-as-you-go report credit was added to your account.',
    'billing.academic_added': 'Academic subscription was activated for your account.',
    'billing.professional_added': 'Professional subscription was activated for your account.',
    'billing.feature': 'Feature',
    'billing.free_col': 'Free',
    'billing.academic_col': 'Academic (99 SAR/year)',
    'billing.professional_col': 'Professional (199 SAR/year)',
    'billing.feature_clo_reports': 'Reports',
    'billing.feature_pdf': 'PDF Export',
    'billing.feature_csv': 'CSV Export',
    'billing.feature_spec': 'Upload Course Specification',
    'billing.feature_multi': 'Multiple Assessment Files',
    'billing.feature_watermark': 'Watermark',
    'billing.feature_courses': 'Save Courses for Reuse',
    'billing.feature_history': 'Saved Reports History',
    'billing.feature_priority': 'Priority Processing',
    'billing.feature_dashboard': 'Dashboard',
    'billing.feature_batch': 'Batch Report Generation',
    'billing.feature_evidence': 'Accreditation Evidence Folder',
    'billing.feature_early': 'Early Access to New Features',
    'billing.unlimited': 'Unlimited',
    'billing.one_free_report': '1 free report',
    'billing.ten_reports_year': '12 reports/year',
    'billing.one_saved_report': '–',
    'billing.ten_courses': '10 courses',
    'billing.last_12_months': 'Last 12 months',
    'billing.with_watermark': 'With watermark',
    'billing.no_watermark': 'No watermark',
    'contact.title': 'Contact Us',
    'contact.name': 'Name',
    'contact.email': 'Email',
    'contact.organization': 'Organization',
    'contact.college': 'College',
    'contact.department': 'Department',
    'contact.message': 'Message',
    'contact.attachment': 'Attach file',
    'contact.attachment_help': 'CSV, PDF, and image files only.',
    'contact.submit': 'Submit Enquiry',
    'contact.sent': 'Your enquiry has been submitted successfully. We will contact you soon.',
    'contact.invalid_attachment': 'Contact attachment must be a CSV, PDF, or image file.',
    'contact.university_subscription_subject': 'University Subscription Enquiry',
    'contact.enquiry_type': 'Type of Enquiry',
    'contact.enquiry_sub': 'University Subscription',
    'contact.enquiry_issue': 'Technical Support',
    'contact.enquiry_unsupported': 'File Format Support Request',
    'contact.enquiry_suggestion': 'Suggestion',
    'contact.enquiry_other': 'Other',
    'contact.enquiry_select': '-- Select Enquiry Type --',
    'privacy.title': 'Privacy Policy',
    'privacy.about_title': 'About the ETQAN Platform',
    'privacy.about_text': 'The ETQAN platform is a digital platform aimed at supporting faculty members in assessment and academic quality tasks, including data analysis and generating relevant reports. The platform is managed as an independent initiative developed to facilitate these tasks and improve their efficiency, and it does not represent an official system affiliated with any university or accreditation body.',
    'privacy.sharing': 'Data Sharing',
    'privacy.sharing_text': 'Uploaded data is not sold or shared with third parties for commercial purposes. Some data may be processed through service providers that assist in operating the services, and the platform is committed to appropriate measures to protect data and limit the sharing of personal information.',
    'privacy.account_deletion': 'Account Deletion',
    'privacy.account_deletion_text': 'Users can request the deletion of their account at any time through the account settings or by contacting the support team. Upon confirming the deletion request, all account data and associated data will be deleted, and relevant backups will be removed within a period not exceeding 60 days from the date of deletion.',
    'privacy.data_improvement': 'Data Usage for Platform Improvement',
    'privacy.data_improvement_text': 'The ETQAN platform may use uploaded data or data extracted from it, after taking appropriate privacy measures to remove or mask personal information. This is to improve platform performance, develop new features and services, and conduct general statistical analyses that contribute to enhancing the user experience and service quality.',
    'privacy.disclaimer': 'Disclaimer',
    'privacy.user_responsibility': 'User Responsibility',
    'privacy.user_responsibility_text': 'The ETQAN platform relies on automated procedures for data verification and result calculation to provide accurate and consistent reports. However, the final review and approval of the reports remains the user\'s responsibility in accordance with the policies and procedures established by their institution.',
    'privacy.updates_title': 'Privacy Policy Updates',
    'privacy.updates_text': 'The ETQAN platform reserves the right to update the privacy policy from time to time in accordance with the development of the platform and its services. Continued use of the platform after updates are published constitutes acceptance of the modified policy.',
    'privacy.back': 'Back to Tool',
    'faq.title': 'Frequently Asked Questions',
    'faq.home': 'Home',
    'faq.q_supported_files': 'What grade files does ETQAN support?',
    'faq.a_supported_files': 'ETQAN supports a wide range of grade files generated by automated grading machines used in universities, as well as files exported from ZipGrade. The tool can analyze multiple formats and automatically recognize student data, questions, grades, learning outcomes, and related fields for report generation.',
    'faq.q_specs': 'What course specification files does ETQAN support?',
    'faq.a_specs': 'ETQAN supports course specification templates approved by the Education and Training Evaluation Commission (NCAAA). It can automatically extract course data and learning outcomes from them to reduce manual entry.',
    'faq.q_unsupported': 'What if my file is not supported?',
    'faq.a_unsupported': 'If your file is not supported, contact us and send a sample of the file. We will review its structure and work on supporting it as soon as possible. We continuously strive to expand our supported file range to meet the diverse needs of faculty members.',
    'faq.q_save_courses': 'Why do I need to save courses?',
    'faq.a_save_courses': 'Saving a course stores its details and learning outcomes in your profile, saving you the time and effort of re-entering them each time you generate a new report.',
    'faq.q_visual_identity': 'Can I change the visual identity used in exported reports?',
    'faq.a_visual_identity': 'Yes. You can keep the default visual identity or override it by uploading a JPEG logo and choosing colors from the visual identity page.',
    'faq.q_language': 'Does ETQAN support Arabic?',
    'faq.a_language': 'Yes. ETQAN supports both Arabic and English, and it can process grade files and course specifications in both languages.',
    'checkout.title': 'Checkout',
    'checkout.selected_plan': 'Selected Plan',
    'checkout.quantity': 'Number of reports',
    'checkout.total': 'Total Cost',
    'checkout.billing_info': 'Billing Information',
    'checkout.card_name': 'Name on Card',
    'checkout.card_number': 'Card Number',
    'checkout.expiry': 'Expiry Date',
    'checkout.cvv': 'CVV',
    'checkout.pay': 'Pay Now',
    'validation.complete_required': 'Please complete the required fields.',
    'validation.select_item': 'Please select an item in the list.',
    'validation.fill_field': 'Please fill out this field.',
    'validation.select_file': 'Please select a file.',
    'validation.valid_email': 'Please enter a valid email address.',
    'validation.match_format': 'Please match the requested format.',
    'validation.number': 'Please enter a number.',
    'validation.range_underflow': 'The value is too low.',
    'validation.range_overflow': 'The value is too high.',
    'validation.too_short': 'The value is too short.',
    'validation.too_long': 'The value is too long.',
    'validation.step_mismatch': 'Please enter a valid value.',
    'validation.duplicate_course_report': 'This course report is already added.',
}

TRANSLATIONS = {
    'ar': {
        'app.title': 'إتقان',
        'app.subtitle': 'Educational Transformation & Quality ANalytics',
        'nav.services': 'الرئيسية',
        'nav.account': 'معلومات الحساب',
        'nav.my_reports': 'تقاريري',
        'nav.my_courses': 'مقرراتي',
        'nav.home': 'الرئيسية',
        'nav.faq': 'الأسئلة الشائعة',
        'nav.contact': 'تواصل معنا',
        'nav.organization': 'الهوية البصرية',
        'nav.settings': 'الإعدادات',
        'nav.billing': 'الاشتراكات',
        'nav.logout': 'تسجيل الخروج',
        'nav.login': 'تسجيل الدخول',
        'nav.create_account': 'إنشاء حساب',
        'nav.language': 'اللغة',
        'nav.english': 'English',
        'nav.arabic': 'العربية',
        'home.title': 'اختر الخدمة',
        'home.description': 'اختر نوع تقرير أو تحليل الجودة الأكاديمية الذي تريد إنشاءه.',
        'home.course_level': 'خدمات المقررات',
        'home.course_level_description': '\u0633\u064a\u0631 \u0639\u0645\u0644 \u062e\u0627\u0635 \u0628\u062a\u062d\u0644\u064a\u0644 \u0645\u062e\u0631\u062c\u0627\u062a \u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0642\u0631\u0631 \u0648\u0623\u062f\u0644\u0629 \u0627\u0644\u062a\u062d\u0642\u0642 \u0648\u062a\u0642\u0627\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'home.program_level': 'خدمات البرامج',
        'home.program_level_description': '\u0633\u064a\u0631 \u0639\u0645\u0644 \u064a\u062c\u0645\u0649 \u0623\u062f\u0644\u0629 \u0627\u0644\u0645\u0642\u0631\u0637\u0627\u062a \u0644\u062a\u062d\u0644\u064a\u0644 \u0646\u0648\u0627\u062a\u062c \u062a\u0639\u0644\u0645 \u0627\u0644\u0628\u0631\u0646\u0627\u0645\u062c.',
        'home.reviewer_level': 'خدمات المراجعين',
        'home.reviewer_course_report_title': 'مراجعة تقرير المقرر',
        'home.reviewer_course_report_description': 'فحص تقرير المقرر مقابل متطلبات NCAAA وتحديد النواقص أو التعارضات.',
        'home.reviewer_course_spec_title': 'مراجعة توصيف المقرر',
        'home.reviewer_course_spec_description': 'فحص توصيف المقرر من حيث نواتج التعلم والموضوعات ومواءمة التقويم واكتمال متطلبات NCAAA.',
        'home.reviewer_program_spec_title': 'مراجعة توصيف البرنامج',
        'home.reviewer_program_spec_description': 'فحص توصيف البرنامج من حيث نواتج تعلم البرنامج وبنية الخطة وطرق التقويم واكتمال متطلبات NCAAA.',
        'home.reviewer_clo_mapping_title': 'مراجعة ربط نواتج التعلم',
        'home.reviewer_clo_mapping_description': 'التحقق من مواءمة الأسئلة ونواتج التعلم والمستهدفات ونتائج التحقق.',
        'home.reviewer_evidence_title': 'مراجعة أدلة الاعتماد',
        'home.reviewer_evidence_description': 'فحص الملفات وحزم الأدلة الداعمة قبل التسليم أو زيارات المراجعة.',
        'home.add_program_title': 'إضافة برنامج',
        'home.add_program_description': 'أضف برنامجًا لتتمكن من إنشاء التقارير.',
        'home.requires_program': 'يتطلب إضافة برنامج واحد على الأقل.',
        'home.inactive': '\u063a\u064a\u0631 \u0645\u0641\u0639\u0644',
        'home.clo_title': 'تحليل تحقق نواتج تعلم المقرر',
        'home.clo_description': 'قياس تحقق نواتج تعلم المقرر من بيانات التقييمات، وتحديد فجوات الأداء، وإنشاء تقارير قائمة على الأدلة.',
        'home.add_course_title': 'إضافة مقرر',
        'home.add_course_description': 'أضف مقرراً لتتمكن من إنشاء التقارير.',
        'home.requires_course': 'يتطلب إضافة مقرر واحد على الأقل.',
        'home.question_mapping_title': 'ربط أسئلة الاختبار بنواتج التعلم',
        'home.question_mapping_description': 'ارفع ورقة الاختبار لربط الأسئلة بنواتج تعلم المقرر.',
        'home.question_mapping_extract': 'استخراج الأسئلة',
        'question_mapping.review_title': 'مراجعة الأسئلة المستخرجة',
        'question_mapping.review_description': 'عدّل الأسئلة المستخرجة، وأضف أي سؤال ناقص عند الحاجة، ثم اربطها بنواتج التعلم.',
        'question_mapping.add_question': 'إضافة سؤال',
        'question_mapping.map_to_clos': 'متابعة',
        'question_mapping.question_text': 'نص السؤال',
        'question_mapping.paper_clo': 'ناتج التعلم من الورقة',
        'question_mapping.paper_clo_help': '',
        'question_mapping.link_title': 'ربط الأسئلة بالنواتج',
        'question_mapping.link_description': 'راجع رقم كل سؤال وناتج التعلم المرتبط الذي يقترحه إتقان.',
        'question_mapping.related_clo': 'الناتج المرتبط',
        'question_mapping.no_questions': 'لم يتم استخراج أي أسئلة. أضف سؤالًا واحدًا على الأقل قبل الربط.',
        'question_mapping.select_at_least_one': 'يرجى تحديد ناتج تعلم واحد على الأقل لـ:',
        'home.assessment_balance_title': 'فحص توازن التقييم',
        'home.assessment_balance_description': 'راجع تغطية التقييم وتوزيع الدرجات وتوازنها عبر نواتج التعلم قبل إعداد التقرير.',
        'home.plo_title': 'تحليل تحقق نواتج تعلم البرنامج',
        'home.plo_description': 'قياس تحقق نواتج تعلم البرنامج من خلال تحليل أدلة المقررات وإنشاء تقرير أداء نواتج تعلم البرنامج.',
        'home.course_report_title': 'تقرير المقرر',
        'home.course_report_description': 'إنشاء تقرير مقرر متوافق مع نموذج NCAAA باستخدام نتائج تحليل تحقق نواتج التعلم وبيانات أداء المقرر.',
        'home.no_courses_message': 'مرحباً بك في إتقان. ابدأ بإضافة مقررك الأول.',
        'home.add_course': 'إضافة مقرر',
        'course_report.select_description': '\u0627\u062e\u062a\u0631 \u0645\u0642\u0631\u0631\u0627\u064b \u0644\u0625\u0646\u0634\u0627\u0621 \u062a\u0642\u0631\u064a\u0631\u0647.',
        'course_report.need_clo_report': '\u0644\u0625\u0646\u0634\u0627\u0621 \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631\u060c \u062a\u062d\u062a\u0627\u062c \u0623\u0648\u0644\u0627\u064b \u0625\u0644\u0649 \u062a\u0642\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0645\u062e\u0631\u062c\u0627\u062a \u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'course_report.create_clo_prompt': '\u0647\u0644 \u062a\u0631\u063a\u0628 \u0628\u0625\u0646\u0634\u0627\u0621 \u0648\u0627\u062d\u062f \u0627\u0644\u0622\u0646\u061f',
        'course_report.create_clo_report': '\u0625\u0646\u0634\u0627\u0621 \u062a\u0642\u0631\u064a\u0631 CLO',
        'course_report.associated_reports': '\u062a\u0642\u0627\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0645\u062e\u0631\u062c\u0627\u062a \u0627\u0644\u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0631\u062a\u0628\u0637\u0629',
        'course_report.associated_reports_help': '\u062d\u062f\u062f \u062a\u0642\u0631\u064a\u0631\u0627\u064b \u0648\u0627\u062d\u062f\u0627\u064b \u0623\u0648 \u0623\u0643\u062b\u0631 \u0645\u0646 \u062a\u0642\u0627\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0645\u062e\u0631\u062c\u0627\u062a \u0627\u0644\u062a\u0639\u0644\u0645 \u0644\u0627\u0633\u062a\u062d\u062f\u0627\u0645\u0647\u0627 \u0641\u064a \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'course_report.no_associated_reports': '\u0644\u0627 \u064a\u0648\u062c\u062f \u062a\u0642\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0645\u062e\u0631\u062c\u0627\u062a \u0627\u0644\u062a\u0639\u0644\u0645 \u0645\u0631\u062a\u0628\u0637 \u0628\u0647\u0630\u0627 \u0627\u0644\u0645\u0642\u0631\u0631. \u064a\u062c\u0628 \u0625\u0646\u0634\u0627\u0621 \u0648\u0627\u062d\u062f \u0642\u0628\u0644 \u0625\u0646\u0634\u0627\u0621 \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'course_report.use_report': '\u0627\u0633\u062a\u062e\u062f\u0627\u0645 \u0647\u0630\u0627 \u0627\u0644\u062a\u0642\u0631\u064a\u0631',
        'course_report.create_report': '\u0627\u0644\u062a\u0627\u0644\u064a',
        'course_report.preview_title': '\u0645\u0639\u0627\u064a\u0646\u0629 \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631',
        'course_report.student_results_comment_edit': 'التعليق على نتائج الطلاب',
        'course_report.preview_description': '\u0631\u0627\u062c\u0639 \u0628\u064a\u0627\u0646\u0627\u062a \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631 \u0627\u0644\u0645\u0643\u062a\u0645\u0644\u0629\u060c \u062b\u0645 \u0635\u062f\u0651\u0631 \u0627\u0644\u062a\u0642\u0631\u064a\u0631 \u0628\u0635\u064a\u063a\u0629 Word.',
        'course_report.saved_automatically': 'تم الحفظ تلقائياً.',
        'course_report.next': '\u0627\u0644\u062a\u0627\u0644\u064a',
        'course_report.export_word': '\u062a\u0635\u062f\u064a\u0631 Word',
        'course_report.export_pdf': '\u062a\u0635\u062f\u064a\u0631 PDF',
        'course_report.report_details': '\u0628\u064a\u0627\u0646\u0627\u062a \u0627\u0644\u062a\u0642\u0631\u064a\u0631',
        'course_report.grade_distribution': '\u062a\u0648\u0632\u064a\u0639 \u0627\u0644\u062f\u0631\u062c\u0627\u062a \u0627\u0644\u0646\u0647\u0627\u0626\u064a\u0629',
        'course_report.clo_summary': '\u0646\u062a\u0627\u0626\u062c \u062a\u0642\u0648\u064a\u0645 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
        'course_report.recommendations': '\u0627\u0644\u062a\u0648\u0635\u064a\u0627\u062a',
        'course_report.select_one_report': '\u062d\u062f\u062f \u062a\u0642\u0631\u064a\u0631 CLO \u0648\u0627\u062d\u062f\u0627\u064b \u0639\u0644\u0649 \u0627\u0644\u0623\u0642\u0644.',
        'course_report.selected_reports': '\u062a\u0642\u0627\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 CLO \u0627\u0644\u0645\u062d\u062f\u062f\u0629',
        'course_report.continue': '\u0627\u0644\u062a\u0627\u0644\u064a',
        'home.open_service': 'ابدأ',
        'service.coming_soon': 'صفحة هذه الخدمة جاهزة. ستتم إضافة أدوات سير العمل الكاملة هنا.',
        'service.back_home': 'عودة',
        'footer.copyright': 'إتقان © 2026. جميع الحقوق محفوظة.',
        'footer.privacy': 'سياسة الخصوصية',
        'footer.faq': 'الأسئلة الشائعة',
        'footer.contact': 'تواصل معنا',
        'auth.email': 'البريد الإلكتروني',
        'auth.password': 'كلمة المرور',
        'auth.university': 'الجامعة',
        'auth.college_optional': 'الكلية',
        'auth.department_optional': 'القسم',
        'auth.select_university': 'اختر الجامعة...',
        'auth.other_university': 'أخرى',
        'auth.university_placeholder': 'أدخل اسم الجامعة',
        'auth.need_account': 'تحتاج إلى حساب؟',
        'auth.create_one': 'أنشئ حساباً',
        'auth.have_account': 'لديك حساب؟',
        'auth.login': 'تسجيل الدخول',
        'auth.invalid_login': 'البريد الإلكتروني أو كلمة المرور غير صحيحة.',
        'auth.register_title': 'إنشاء حساب',
        'auth.login_title': 'تسجيل الدخول',
        'auth.register_button': 'إنشاء حساب',
        'auth.login_button': 'تسجيل الدخول',
        'auth.forgot_password': 'نسيت كلمة المرور؟',
        'auth.forgot_title': 'إعادة تعيين كلمة المرور',
        'auth.send_reset': 'إعادة التعيين',
        'auth.reset_sent': 'إذا كان البريد مسجلاً، فقد تم إرسال رابط إعادة تعيين كلمة المرور.',
        'auth.reset_email_unconfigured': 'إرسال البريد غير مفعّل حالياً. يرجى إعداد SMTP قبل إرسال روابط إعادة التعيين.',
        'auth.reset_email_failed': 'تعذر إرسال بريد إعادة التعيين حالياً. يرجى المحاولة لاحقاً.',
        'auth.login_locked_email': 'حدثت محاولات تسجيل دخول فاشلة كثيرة لهذا البريد. يرجى المحاولة لاحقاً.',
        'auth.login_locked_ip': 'حدثت محاولات تسجيل دخول فاشلة كثيرة من هذه الشبكة. يرجى المحاولة لاحقاً.',
        'auth.new_password': 'كلمة المرور الجديدة',
        'auth.confirm_password': 'تأكيد كلمة المرور',
        'auth.reset_button': 'تحديث كلمة المرور',
        'auth.back_login': 'عودة',
        'org.title': 'الهوية البصرية',
        'org.description': 'تُستخدم هذه البيانات كهوية بصرية في التقارير المصدّرة.',
        'org.university': 'الجامعة',
        'org.other_university': 'أخرى',
        'org.university_placeholder': 'أدخل اسم الجامعة',
        'org.department': 'القسم',
        'org.department_placeholder': '',
        'org.logo': 'لاستبدال الشعار الحالي، يرجى إرفاق الشعار الجديد.',
        'org.official_website': 'الموقع الرسمي',
        'org.used_logo': 'الشعار المستخدم',
        'org.logo_alt': 'شعار الجهة',
        'org.no_logo': 'لم يتم اختيار شعار',
        'org.primary_color': 'اللون الأساسي',
        'org.secondary_color': 'اللون الثانوي',
        'org.tertiary_color': 'اللون الثالث',
        'org.current_logo': 'الشعار الحالي:',
        'org.save': 'حفظ',
        'org.saved': 'تم الحفظ بنجاح.',
        'account.danger_title': 'حذف الحساب',
        'account.danger_text': 'سيؤدي حذف الحساب إلى إزالة التقارير المحفوظة والهوية البصرية وبيانات الاشتراك من منصة إتقان بشكل نهائي. كما ستفقد أي أرصدة تقارير متبقية في حسابك، ولا يمكن استعادتها بعد إتمام عملية الحذف.',
        'account.confirm_label': 'اكتب DELETE للتأكيد',
        'account.delete_button': 'حذف حسابي',
        'account.delete_confirm_js': 'سيتم حذف حسابك والتقارير المحفوظة نهائياً. هل تريد المتابعة؟',
        'account.delete_invalid': 'لم يتم تأكيد حذف الحساب.',
        'account.deleted': 'تم حذف حسابك.',
        'account.delete_help': 'لا يمكن التراجع عن هذا الإجراء.',
        'account.title': 'معلومات الحساب',
        'account.description': 'إدارة حساب الدخول والإجراءات المرتبطة بالحساب.',
        'account.email': 'البريد الإلكتروني',
        'account.university': 'المؤسسة',
        'account.college': 'الكلية',
        'account.department': 'القسم',
        'account.created_at': 'تاريخ الإنشاء',
        'account.plan': 'الخطة',
        'account.edit': 'تعديل',
        'account.cancel': 'إلغاء',
        'account.save': 'حفظ',
        'account.saved': 'تم الحفظ بنجاح.',
        'settings.title': 'الإعدادات',
        'settings.description': 'إدارة معلومات الحساب والهوية البصرية وتفضيلات التقارير المصدّرة.',
        'settings.account_title': 'معلومات الحساب',
        'settings.account_description': 'الإجراءات المرتبطة بالحساب كتعديل البيانات والحذف.',
        'settings.organization_title': 'الهوية البصرية',
        'settings.organization_description': 'إدارة الهوية المستخدمة في التقارير المصدّرة.',
        'settings.report_title': 'إعدادات التقارير',
        'settings.report_description': 'تحديد تفضيلات التقارير المصدّرة.',
        'settings.open': 'فتح',
        'report_settings.title': 'إعدادات التقارير',
        'report_settings.description': 'اختر طريقة إنشاء التقارير المصدّرة.',
        'report_settings.language': 'لغة التقارير المصدرة',
        'report_settings.language_help': '',
        'report_settings.same_as_interface': 'نفس لغة الواجهة',
        'report_settings.english': 'English',
        'report_settings.arabic': 'العربية',
        'report_settings.save': 'حفظ',
        'report_settings.saved': 'تم الحفظ بنجاح.',
        'courses.title': 'مقرراتي',
        'courses.description': 'احفظ معلومات المقرر مرة واحدة ثم أعد استخدامها عند إنشاء التقارير.',
        'courses.add_new': 'إضافة مقرر',
        'courses.back_to_courses': 'عودة',
        'courses.add_title': 'إضافة مقرر',
        'courses.edit_title': 'تعديل المقرر',
        'courses.edit_help': 'تحديث معلومات المقرر أو رفع توصيف جديد.',
        'courses.edit': 'تعديل',
        'courses.method_help': 'ارفع ملف توصيف المقرر، وسيستخرج إتقان بيانات المقرر.',
        'courses.upload_method': 'إدخال توصيف المقرر',
        'courses.upload_help': 'ستستخرج إتقان اسم المقرر ورمزه ونواتج التعلم من ملف PDF إذا كان قابلاً للقراءة.',
        'courses.manual_prompt': 'أو هل ترغب في إدخال معلومات المقرر يدوياً؟',
        'courses.manual_method': 'معلومات المقرر',
        'courses.manual_help': 'استخدم هذه الطريقة إذا لم يكن لديك ملف توصيف أو أردت تعديل المعلومات المستخرجة من التوصيف.',
        'courses.course_name': 'اسم المقرر',
        'courses.course_code': 'رمز المقرر',
        'courses.college': 'الكلية',
        'courses.program': 'البرنامج',
        'courses.department': 'القسم',
        'courses.optional_override': 'تعديل اختياري',
        'courses.spec_file': 'توصيف المقرر',
        'courses.spec_file_help': 'يمكن رفع ملف PDF أو Word.',
        'courses.extract': 'استخراج المعلومات',
        'courses.extracting': 'جاري الاستخراج...',
        'courses.extract_missing': 'يرجى رفع ملف توصيف المقرر أولاً.',
        'courses.extracted': 'تم استخراج بيانات المقرر. راجعها ثم اضغط حفظ المقرر.',
        'courses.extraction_method_prefix': 'طريقة الاستخراج:',
        'courses.extraction_method_gemini': 'Gemini Flash',
        'courses.extraction_method_qwen': 'Qwen via Groq',
        'courses.extraction_method_llama': 'Llama عبر Groq',
        'courses.extraction_method_groq': 'Groq AI',
        'courses.extraction_method_local': 'النص المحلي / OCR',
        'courses.clos': 'نواتج التعلم للمقرر',
        'courses.clo': '\u0646\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
        'courses.associated_plo': 'ناتج التعلم المرتبط من البرنامج',
        'courses.actions': '\u0627\u0644\u0625\u062c\u0631\u0627\u0621\u0627\u062a',
        'courses.add_clo_row': '\u0625\u0636\u0627\u0641\u0629 \u0646\u0627\u062a\u062c \u062a\u0639\u0644\u0645',
        'courses.remove_clo': '\u062d\u0630\u0641',
        'courses.topics': '\u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0627\u0644\u0645\u0642\u0631\u0631',
        'courses.topics_placeholder': '\u0623\u062f\u062e\u0644 \u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0627\u0644\u0645\u0642\u0631\u0631\u060c \u0643\u0644 \u0645\u0648\u0636\u0648\u0639 \u0641\u064a \u0633\u0637\u0631',
        'courses.save': 'حفظ المقرر',
        'courses.empty': 'لا توجد مقررات محفوظة بعد. أضف مقررًا من الصفحة الرئيسية.',
        'courses.delete': 'حذف',
        'courses.delete_confirm': 'حذف هذا المقرر المحفوظ؟',
        'courses.saved': 'تم الحفظ بنجاح.',
        'courses.deleted': 'تم حذف المقرر.',
        'programs.login_required': 'يرجى تسجيل الدخول لإدارة البرامج.',
        'programs.add_title': 'إضافة برنامج',
        'programs.program_name': 'اسم البرنامج',
        'programs.program_code': 'رمز البرنامج',
        'programs.college': 'الكلية',
        'programs.department': 'القسم',
        'programs.plos': 'نواتج تعلم البرنامج',
        'programs.plos_help': 'اكتب كل ناتج تعلم في سطر مستقل، مثل: K1 المعرفة، S1 المهارات، V1 القيم.',
        'programs.save': 'حفظ البرنامج',
        'programs.invalid': 'أدخل اسم البرنامج وناتج تعلم واحداً على الأقل.',
        'courses.invalid': 'أدخل اسم المقرر وموضوعات المقرر وناتج تعلم واحد على الأقل، أو ارفع ملف توصيف مقرر قابل للقراءة.',
        'courses.limit': 'خطتك لا تسمح بحفظ المزيد من المقررات.',
        'courses.login_required': 'يرجى تسجيل الدخول لإدارة المقررات المحفوظة.',
        'index.title': 'معلومات المقرر',
        'index.course': 'المقرر',
        'index.select_course': 'اختر مقرراً...',
        'index.course_not_found': 'المقرر غير موجود؟',
        'index.course_not_found_my_courses': 'المقرر غير موجود؟ أضفه من خلال خدمة إضافة مقرر.',
        'index.course_spec_title': 'رفع توصيف المقرر PDF',
        'index.inline_course_title': 'إضافة معلومات المقرر',
        'index.inline_course_help': 'إذا لم يكن المقرر موجوداً في القائمة، أدخل اسمه وارفع توصيف المقرر PDF أو الصق نواتج التعلم أدناه.',
        'index.inline_course_name': 'اسم المقرر',
        'index.inline_course_name_placeholder': 'مثال: هياكل البيانات (DS2206)',
        'index.inline_spec_file': 'توصيف المقرر',
        'index.inline_clos': 'لصق نواتج التعلم',
        'index.inline_clos_placeholder': 'الصق مخرجاً واحداً في كل سطر، مثال:\n1.1 يعرّف المفاهيم الأساسية\n2.1 يطبق الأساليب المناسبة\n3.1 يلتزم بالقيم المهنية',
        'index.target_label': 'المستوى المستهدف لكل ناتج تعلم',
        'index.target_help': 'حدد الحد الأدنى للنسبة التي يحتاج الطالب إلى تحقيقها لكل ناتج تعلم.',
        'index.clo': 'ناتج تعلم المقرر (CLO)',
        'index.target_level': 'المستوى المستهدف (%)',
        'domain.knowledge': '1.0 المعرفة والفهم',
        'domain.skills': '2.0 المهارات',
        'domain.values': '3.0 القيم',
        'domain.other': 'أخرى',
        'index.select_course_populate': 'اختر مقرراً لعرض النواتج...',
        'index.no_clos_category': 'لا توجد نواتج تعلم في هذه الفئة.',
        'index.assessment_files': 'ملفات التقييم',
        'index.assessment_help': 'ارفع ملف درجات واحداً على الأقل.',
        'index.type': 'النوع',
        'index.grades_file': 'ملف الدرجات',
        'index.exam_paper': 'ورقة الاختبار',
        'assessment.quiz': 'اختبار قصير',
        'assessment.assignment': 'تكليف',
        'assessment.midterm': 'اختبار نصفي',
        'assessment.final': 'اختبار نهائي',
        'assessment.project': 'مشروع',
        'assessment.other': 'أخرى',
        'index.remove': 'حذف',
        'index.add_file': 'إضافة ملف',
        'index.next_mapping': '\u0627\u0644\u062a\u0627\u0644\u064a',
        'index.error_course': 'يرجى اختيار مقرر أو إضافة معلومات المقرر قبل المتابعة.',
        'index.error_file': 'ارفع ملف درجات واحداً على الأقل.',
        'spec.title': 'رفع توصيف المقرر',
        'spec.back': 'عودة',
        'spec.description': 'ارفع ملف PDF لتوصيف المقرر لاستخراج اسم المقرر ورقمه ونواتج التعلم. تُدعم ملفات PDF النصية العربية والملفات الممسوحة عند توفر OCR.',
        'spec.file': 'ملف توصيف المقرر',
        'spec.extract': 'استخراج معلومات المقرر',
        'spec.preview': 'معاينة المعلومات المستخرجة',
        'spec.list_name': 'اسم المقرر في القائمة',
        'spec.course_name': 'اسم المقرر',
        'spec.course_number': 'رقم المقرر',
        'spec.clos': 'نواتج التعلم',
        'spec.no_knowledge': 'لم يتم اكتشاف نواتج معرفة.',
        'spec.no_skills': 'لم يتم اكتشاف نواتج مهارات.',
        'spec.no_values': 'لم يتم اكتشاف نواتج قيم.',
        'spec.add_course': 'إضافة المقرر إلى القائمة',
        'spec.upload_another': 'رفع ملف PDF آخر',
        'spec.cannot_add': 'لا يمكن إضافة المقرر حتى يتم اكتشاف اسم المقرر وقائمة نواتج التعلم.',
        'mapping.title': 'ربط الأسئلة بنواتج التعلم',
        'mapping.course': 'المقرر:',
        'mapping.detected': 'تم اكتشاف:',
        'mapping.description': 'أربط كل سؤال بناتج تعلم واحد أو أكثر',
        'mapping.exam_paper': 'ورقة الاختبار:',
        'mapping.questions': 'أسئلة',
        'mapping.students': 'طلاب',
        'mapping.question': 'السؤال',
        'mapping.max_score': 'الدرجة القصوى',
        'mapping.multi_help': 'اضغط Ctrl لاختيار أكثر من ناتج تعلم.',
        'mapping.selected_clo': 'معرّف ناتج التعلم المحدد',
        'mapping.no_questions': 'لم يتم اكتشاف أي أسئلة. يرجى الرجوع ورفع ملف درجات صالح.',
        'mapping.back': 'عودة',
        'mapping.calculate': 'التالي',
        'detected.title': 'هيكل تقرير المقرر المكتشف',
        'detected.source': 'المصدر:',
        'detected.questions': 'الأسئلة',
        'detected.students': 'الطلاب',
        'detected.detection': 'الاكتشاف',
        'detected.save': 'حفظ اختيار النواتج',
        'detected.back_upload': 'عودة',
        'detected.no_table': 'لم يتم اكتشاف جدول أسئلة من هذا الملف.',
        'detected.text_sample': 'عرض عينة النص المستخرج',
        'detected.suggestion_source': 'مصدر الاقتراح:',
        'detected.source_gemini': 'Gemini Flash',
        'detected.source_local': 'المطابقة الدلالية المحلية',
        'detected.mapping_used_gemini': 'تم ربط الأسئلة بالنواتج باستخدام Gemini Flash.',
        'detected.mapping_used_qwen': 'تم ربط الأسئلة بالنواتج باستخدام Qwen عبر Groq.',
        'detected.mapping_used_local': 'تم ربط الأسئلة بالنواتج باستخدام المطابقة الدلالية المحلية.',
        'detected.question_text': 'نص السؤال',
        'results.title': 'تقرير تحقق نواتج التعلم',
        'results.course': 'المقرر:',
        'results.total_students': 'إجمالي الطلاب:',
        'results.export_csv': 'تصدير CSV',
        'results.export_pdf': 'تصدير PDF',
        'results.export_course_report': 'تصدير تقرير المقرر DOCX',
        'results.save_course_report': 'حفظ تقرير المقرر',
        'results.export_word': 'تصدير Word',
        'results.course_report_inputs': '\u0645\u062f\u062e\u0644\u0627\u062a \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631',
        'results.course_report_optional_details': 'البيانات الأساسية',
        'results.course_instructor': 'أستاذ المقرر',
        'results.course_coordinator': 'منسق المقرر',
        'results.course_location': 'مكان تقديم المقرر',
        'results.main_campus': 'المقر الرئيس',
        'results.branch': 'فرع',
        'results.branch_name': 'اسم الفرع',
        'results.number_of_sections': 'عدد الشعب',
        'results.students_started': 'عدد الطلاب الذين بدأوا المقرر',
        'results.students_completed': 'عدد الطلاب الذين أنهوا المقرر',
        'results.topics_covered': '\u0647\u0644 \u062a\u0645\u062a \u062a\u063a\u0637\u064a\u0629 \u062c\u0645\u064a\u0639 \u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0627\u0644\u0645\u0642\u0631\u0631\u061f',
        'results.topics_covered_yes': '\u0646\u0639\u0645\u060c \u062a\u0645\u062a \u062a\u063a\u0637\u064a\u0629 \u062c\u0645\u064a\u0639 \u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a',
        'results.topics_covered_no': '\u0644\u0627\u060c \u0644\u0645 \u062a\u062a\u0645 \u062a\u063a\u0637\u064a\u0629 \u0628\u0639\u0636 \u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a',
        'results.uncovered_topics': '\u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a \u063a\u064a\u0631 \u0627\u0644\u0645\u063a\u0637\u0627\u0629',
        'results.uncovered_topics_help': '\u062d\u062f\u062f \u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0645\u0646 \u062a\u0648\u0635\u064a\u0641 \u0627\u0644\u0645\u0642\u0631\u0631 \u0627\u0644\u062a\u064a \u0644\u0645 \u062a\u062a\u0645 \u062a\u063a\u0637\u064a\u062a\u0647\u0627.',
        'results.no_extracted_topics': '\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0645\u0633\u062a\u062e\u0631\u062c\u0629 \u0644\u0647\u0630\u0627 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'results.uncovered_reason': '\u0633\u0628\u0628 \u0639\u062f\u0645 \u0627\u0644\u062a\u063a\u0637\u064a\u0629 / \u0627\u0644\u0641\u0631\u0648\u0642',
        'results.uncovered_reason_select': '\u062d\u062f\u062f \u0627\u0644\u0633\u0628\u0628',
        'results.reason_time_constraints': '\u0636\u064a\u0642 \u0627\u0644\u0648\u0642\u062a',
        'results.reason_absence': '\u063a\u064a\u0627\u0628 \u0627\u0644\u0637\u0644\u0627\u0628 / \u0627\u0646\u062e\u0641\u0627\u0636 \u0627\u0644\u062d\u0636\u0648\u0631',
        'results.reason_other_course': '\u062a\u0645\u062a \u062a\u063a\u0637\u064a\u0629 \u0627\u0644\u0645\u0648\u0636\u0648\u0639 \u0641\u064a \u0645\u0642\u0631\u0631 \u0622\u062e\u0631',
        'results.reason_overload': '\u0643\u062b\u0627\u0641\u0629 \u0645\u062d\u062a\u0648\u0649 \u0627\u0644\u0645\u0642\u0631\u0631',
        'results.reason_equipment': '\u0639\u062f\u0645 \u062a\u0648\u0641\u0631 \u0627\u0644\u0623\u062c\u0647\u0632\u0629',
        'results.reason_replaced': '\u0627\u0633\u062a\u0628\u062f\u0627\u0644 \u0627\u0644\u0645\u0648\u0636\u0648\u0639 \u0628\u0645\u0648\u0636\u0648\u0639 \u0623\u062d\u062f\u062b',
        'results.reason_holidays': '\u0627\u0644\u0625\u062c\u0627\u0632\u0627\u062a \u0627\u0644\u0631\u0633\u0645\u064a\u0629 / \u062a\u0639\u0644\u064a\u0642 \u0627\u0644\u062f\u0631\u0627\u0633\u0629',
        'results.reason_lab_limitations': '\u0642\u064a\u0648\u062f \u062a\u0642\u0646\u064a\u0629 \u0623\u0648 \u0645\u0639\u0645\u0644\u064a\u0629',
        'results.reason_prerequisite_gap': '\u0641\u062c\u0648\u0629 \u0641\u064a \u0627\u0644\u0645\u0639\u0631\u0641\u0629 \u0627\u0644\u0633\u0627\u0628\u0642\u0629',
        'results.reason_other': '\u0623\u062e\u0631\u0649',
        'results.reason_other_explanation': '\u062a\u0648\u0636\u064a\u062d \u0623\u062e\u0631\u0649',
        'results.please_explain': '\u064a\u0631\u062c\u0649 \u0627\u0644\u062a\u0648\u0636\u064a\u062d',
        'results.uncovered_impact': '\u0645\u062f\u0649 \u0627\u0644\u0623\u062b\u0631 \u0639\u0644\u0649 \u0645\u062e\u0631\u062c\u0627\u062a \u0627\u0644\u062a\u0639\u0644\u0645',
        'results.impact_none': '\u0644\u0627 \u064a\u0648\u062c\u062f',
        'results.impact_low': '\u0645\u0646\u062e\u0641\u0636',
        'results.impact_medium': '\u0645\u062a\u0648\u0633\u0637',
        'results.impact_high': '\u0639\u0627\u0644\u064a',
        'results.uncovered_action': '\u0627\u0644\u0625\u062c\u0631\u0627\u0621 \u0627\u0644\u062a\u0639\u0648\u064a\u0636\u064a',
        'results.uncovered_action_select': '\u062d\u062f\u062f \u0627\u0644\u0625\u062c\u0631\u0627\u0621',
    'results.all_topics_covered': '\u062a\u0645\u062a \u062a\u063a\u0637\u064a\u0629 \u062c\u0645\u064a\u0639 \u0627\u0644\u0645\u0648\u0627\u0636\u064a\u0639 \u0628\u0646\u062c\u0627\u062d.',
        'results.action_supplementary_resources': '\u062a\u0642\u062f\u064a\u0645 \u0645\u0635\u0627\u062f\u0631 \u062a\u0639\u0644\u0645 \u0625\u0636\u0627\u0641\u064a\u0629',
        'results.action_none_required': '\u0644\u0627 \u064a\u0648\u062c\u062f \u0625\u062c\u0631\u0627\u0621 \u0645\u0637\u0644\u0648\u0628',
        'results.action_other': '\u0623\u062e\u0631\u0649',
        'results.action_other_explanation': '\u062a\u0648\u0636\u064a\u062d \u0627\u0644\u0625\u062c\u0631\u0627\u0621 \u0627\u0644\u0622\u062e\u0631',
        'results.final_grades_file': '\u0645\u0644\u0641 \u0627\u0644\u062f\u0631\u062c\u0627\u062a \u0627\u0644\u0646\u0647\u0627\u0626\u064a\u0629',
        'results.final_grades_help': '\u0627\u0631\u0641\u0639 CSV \u0623\u0648 Excel \u0623\u0648 PDF \u064a\u062d\u062a\u0648\u064a \u0639\u0644\u0649 \u0627\u0644\u062a\u0642\u062f\u064a\u0631\u0627\u062a \u0623\u0648 \u0627\u0644\u062f\u0631\u062c\u0627\u062a \u0627\u0644\u0631\u0642\u0645\u064a\u0629 \u0627\u0644\u0646\u0647\u0627\u0626\u064a\u0629. \u0633\u064a\u062d\u0633\u0628 \u0625\u062a\u0642\u0627\u0646 \u0639\u062f\u062f A+ \u0648A \u0648B+ \u0648B \u0648\u063a\u064a\u0631\u0647\u0627.',
        'results.course_improvement_plan': '\u062e\u0637\u0629 \u062a\u062d\u0633\u064a\u0646 \u0627\u0644\u0645\u0642\u0631\u0631',
        'results.course_improvement_help': '\u0627\u062e\u062a\u0631 \u0627\u0644\u062a\u0648\u0635\u064a\u0627\u062a \u0627\u0644\u062a\u064a \u0633\u062a\u0638\u0647\u0631 \u0636\u0645\u0646 \u062e\u0637\u0629 \u062a\u062d\u0633\u064a\u0646 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'results.recommendation': '\u0627\u0644\u062a\u0648\u0635\u064a\u0629',
        'results.actions_needed': '\u0627\u0644\u0625\u062c\u0631\u0627\u0621\u0627\u062a \u0627\u0644\u0645\u0637\u0644\u0648\u0628\u0629',
        'results.needed_support': '\u0627\u0644\u062f\u0639\u0645 \u0627\u0644\u0645\u0637\u0644\u0648\u0628',
        'results.support_other_explanation': '\u062a\u0648\u0636\u064a\u062d \u0627\u0644\u062f\u0639\u0645 \u0627\u0644\u0622\u062e\u0631',
        'results.complete_org': 'إكمال ملف الجهة',
        'results.login_export': 'تسجيل الدخول للتصدير',
        'results.create_account': 'إنشاء حساب',
        'results.org_identity': 'الهوية البصرية للجهة',
        'results.export_requires_account': 'يتطلب تصدير التقارير حساباً حتى يستخدم التقرير الهوية البصرية لجهتك.',
        'results.clo_definitions': 'تعريف نواتج التعلم',
        'results.domain': 'المجال',
        'results.clo': 'الرمز',
        'results.wording': 'نواتج التعلم',
        'results.summary': 'ملخص تحقق نواتج التعلم',
        'results.mapped_questions': 'الأسئلة المرتبطة',
        'results.max_possible': 'الدرجة الكلية',
        'results.target': 'المستهدف',
        'results.students_achieved': 'عدد الطلاب المحققين',
        'results.code': '\u0627\u0644\u0631\u0645\u0632',
        'results.max_score': '\u0627\u0644\u062f\u0631\u062c\u0629 \u0627\u0644\u0643\u0644\u064a\u0629',
        'results.achievement': '\u0646\u0633\u0628\u0629 \u0627\u0644\u062a\u062d\u0642\u0642',
        'results.achievement_pct': 'نسبة التحقق',
        'results.mapped_question_count': 'سؤال مرتبط',
        'results.student_achievement': 'تحقق نواتج التعلم لكل طالب',
        'results.student_id': 'معرّف الطالب',
        'results.achieved': 'محقق',
        'results.not_achieved': 'غير محقق',
        'results.back_mapping': 'عودة',
        'results.start_over': 'البدء من جديد',
        'results.back_reports': 'عودة',
        'results.create_new': 'إنشاء تقرير جديد',
        'results.no_mappings': 'لم يتم تقديم أي ربط. يرجى التأكد من ربط سؤال واحد على الأقل بناتج تعلم.',
        'history.title': 'تقاريري',
        'history.report': 'التقرير',
        'history.course': 'المقرر',
        'history.created': 'تاريخ الإنشاء',
        'history.open': 'فتح',
        'history.rename': '\u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629',
        'history.rename_placeholder': '\u0627\u0633\u0645 \u0627\u0644\u062a\u0642\u0631\u064a\u0631',
        'history.rename_save': '\u062d\u0641\u0638',
        'history.renamed': '\u062a\u0645\u062a \u0625\u0639\u0627\u062f\u0629 \u062a\u0633\u0645\u064a\u0629 \u0627\u0644\u062a\u0642\u0631\u064a\u0631.',
        'history.rename_invalid': '\u0623\u062f\u062e\u0644 \u0627\u0633\u0645\u0627\u064b \u0644\u0644\u062a\u0642\u0631\u064a\u0631.',
        'history.rename_duplicate': '\u064a\u0648\u062c\u062f \u062a\u0642\u0631\u064a\u0631 \u0628\u0647\u0630\u0627 \u0627\u0644\u0627\u0633\u0645 \u0644\u0647\u0630\u0627 \u0627\u0644\u0645\u0642\u0631\u0631.',
        'history.delete': 'حذف',
        'history.delete_confirm': 'هل أنت متأكد من حذف هذا الملف؟ لا يمكن التراجع عن هذا الإجراء.',
        'history.deleted': 'تم حذف التقرير.',
        'history.empty': 'لا توجد تقارير محفوظة بعد. أنشئ تقرير من الصفحة الرئيسية.',
        'billing.title': 'الاشتراكات',
        'billing.description': '',
        'billing.free_status': 'التقارير المجانية المستخدمة',
        'billing.credits': 'رصيد التقارير المتاج',
        'billing.plan': 'الخطة الحالية',
        'billing.yearly_active': 'سنوي',
        'billing.academic_active': 'أكاديمي',
        'billing.professional_active': 'احترافي',
        'billing.free': 'مجانية',
        'billing.payg_title': 'الدفع حسب الاستخدام',
        'billing.payg_price': '19 ريال لكل تقرير',
        'billing.payg_unit': 'ريال لكل تقرير',
        'billing.payg_description': 'مثالي للاستخدام غير المتكرر. تضيف كل عملية شراء رصيد تقرير واحد إلى حسابك.',
        'billing.payg_button': 'اختيار الباقة',
        'billing.academic_title': 'أكاديمي',
        'billing.academic_price': '99 ريال/سنة',
        'billing.academic_description': 'مصمم لأعضاء هيئة التدريس ويشمل حتى 12 تقريراً سنوياً.',
        'billing.academic_button': 'اختيار الباقة',
        'billing.professional_title': 'احترافي',
        'billing.professional_price': '199 ريال/سنة',
        'billing.professional_description': 'مثالي للمستخدمين المتقدمين الذين يحتاجون إلى تقارير غير محدودة ولوحات متابعة وإدارة أدلة الاعتماد وأدوات تقارير متقدمة.',
        'billing.professional_button': 'اختيار الباقة',
        'billing.yearly_unit': 'ريال/سنة',
        'billing.university_title': 'اشتراك جامعة',
        'billing.university_price': 'تواصل معنا',
        'billing.university_description': 'مصمم للأقسام والكليات والجامعات التي تحتاج إلى حسابات مستخدمين متعددة وتقارير مركزية وإدارة جودة على مستوى المؤسسة.',
        'billing.university_button': 'تواصل معنا',
        'billing.local_notice': '',
        'billing.back': 'عودة',
        'billing.limit_message': 'لقد استخدمت التقرير المجاني. يرجى اختيار خيار فوترة لإنشاء تقرير آخر.',
        'billing.payg_added': 'تمت إضافة رصيد تقرير واحد بنظام الدفع حسب الاستخدام إلى حسابك.',
        'billing.academic_added': 'تم تفعيل اشتراك أكاديمي لحسابك.',
        'billing.professional_added': 'تم تفعيل اشتراك احترافي لحسابك.',
        'billing.feature': 'الميزة',
        'billing.free_col': 'مجاني',
        'billing.academic_col': 'أكاديمي (99 ريال/سنة)',
        'billing.professional_col': 'احترافي (199 ريال/سنة)',
        'billing.feature_clo_reports': 'التقارير',
        'billing.feature_pdf': 'تصدير PDF',
        'billing.feature_csv': 'تصدير CSV',
        'billing.feature_spec': 'رفع توصيف المقرر',
        'billing.feature_multi': 'ملفات تقييم متعددة',
        'billing.feature_watermark': 'العلامة المائية',
        'billing.feature_courses': 'حفظ المقررات لإعادة الاستخدام',
        'billing.feature_history': 'سجل التقارير المحفوظة',
        'billing.feature_priority': 'أولوية المعالجة',
        'billing.feature_dashboard': 'لوحة المتابعة',
        'billing.feature_batch': 'توليد تقارير دفعي',
        'billing.feature_evidence': 'مجلد أدلة الاعتماد',
        'billing.feature_early': 'وصول مبكر للميزات الجديدة',
        'billing.unlimited': 'غير محدود',
        'billing.one_free_report': 'تقرير واحد',
        'billing.ten_reports_year': '12 تقرير/سنة',
        'billing.one_saved_report': '–',
        'billing.ten_courses': '10 مقررات',
        'billing.last_12_months': 'آخر 12 شهراً',
        'billing.with_watermark': 'بعلامة مائية',
        'billing.no_watermark': 'بدون علامة مائية',
        'contact.title': 'تواصل معنا',
        'contact.name': 'الاسم',
        'contact.email': 'البريد الإلكتروني',
        'contact.organization': 'الجهة',
        'contact.college': 'الكلية',
        'contact.department': 'القسم',
        'contact.message': 'الرسالة',
        'contact.attachment': 'إرفاق ملف',
        'contact.attachment_help': 'ملفات CSV وPDF والصور فقط.',
        'contact.submit': 'إرسال',
        'contact.sent': 'تم إرسال طلبك بنجاح. سنتواصل معك قريباً.',
        'contact.invalid_attachment': 'يجب أن يكون مرفق التواصل ملف CSV أو PDF أو صورة.',
        'contact.university_subscription_subject': 'استفسار عن اشتراك جامعة',
        'contact.enquiry_type': 'نوع الاستفسار',
        'contact.enquiry_sub': 'اشتراك جامعة',
        'contact.enquiry_issue': 'الدعم الفني',
        'contact.enquiry_unsupported': 'طلب دعم تنسيق ملف',
        'contact.enquiry_suggestion': 'اقتراح',
        'contact.enquiry_other': 'أخرى',
        'contact.enquiry_select': '-- اختر نوع الاستفسار --',
        'privacy.title': 'سياسة الخصوصية',
        'privacy.about_title': 'حول منصة إتقان',
        'privacy.about_text': 'منصة إتقان هي منصة رقمية تهدف إلى دعم أعضاء هيئة التدريس في مهام التقييم والجودة الأكاديمية، بما في ذلك تحليل البيانات وإعداد التقارير ذات الصلة. وتُدار المنصة كمبادرة مستقلة طُورت لتسهيل هذه المهام وتحسين كفاءتها، ولا تمثل نظامًا رسميًا تابعًا لأي جامعة أو جهة اعتماد.',
        'privacy.sharing': 'مشاركة البيانات',
        'privacy.sharing_text': 'لا يتم بيع البيانات المرفوعة أو مشاركتها مع أطراف ثالثة لأغراض تجارية. وقد تتم معالجة بعض البيانات من خلال مزودي الخدمات الذين يساعدون في تشغيل الخدمات، وتلتزم المنصة بالإجراءات المناسبة لحماية البيانات والحد من مشاركة المعلومات الشخصية.',
        'privacy.account_deletion': 'حذف الحساب',
        'privacy.account_deletion_text': 'يمكن للمستخدم طلب حذف حسابه في أي وقت من خلال إعدادات الحساب أو عبر التواصل مع فريق الدعم. وعند تأكيد طلب الحذف، يتم حذف جميع بيانات الحساب والبيانات المرتبطة به، كما تُزال النسخ الاحتياطية ذات الصلة خلال مدة لا تتجاوز 60 يومًا من تاريخ الحذف.',
        'privacy.data_improvement': 'استخدام البيانات لتحسين المنصة',
        'privacy.data_improvement_text': 'قد تستخدم منصة إتقان البيانات المرفوعة أو البيانات المستخرجة منها، بعد اتخاذ الإجراءات المناسبة لحماية الخصوصية وإزالة أو إخفاء المعلومات الشخصية، بهدف تحسين أداء المنصة، وتطوير الميزات والخدمات الجديدة، بالإضافة إلى إجراء تحليلات إحصائية عامة تسهم في تحسين تجربة المستخدم وجودة الخدمة.',
        'privacy.disclaimer': 'إخلاء مسؤولية',
        'privacy.user_responsibility': 'مسؤولية المستخدم',
        'privacy.user_responsibility_text': 'تعتمد منصة إتقان على إجراءات آلية للتحقق من البيانات وحساب النتائج بهدف توفير تقارير دقيقة ومتسقة. ومع ذلك، تبقى المراجعة النهائية للتقارير واعتمادها ضمن مسؤولية المستخدم وفق السياسات والإجراءات المعمول بها في مؤسسته.',
        'privacy.updates_title': 'تحديثات سياسة الخصوصية',
        'privacy.updates_text': 'تحتفظ منصة إتقان بالحق في تحديث سياسة الخصوصية من وقت لآخر بما يتوافق مع تطور المنصة وخدماتها. ويُعد استمرار استخدام المنصة بعد نشر التحديثات موافقة على السياسة المعدلة.',
        'privacy.back': 'عودة',
        'faq.title': 'الأسئلة الشائعة',
        'faq.home': 'الرئيسية',
        'faq.q_supported_files': 'ما ملفات الدرجات التي يدعمها إتقان؟',
        'faq.a_supported_files': 'يدعم إتقان مجموعة واسعة من ملفات الدرجات الناتجة عن أجهزة التصحيح الآلي المستخدمة في الجامعات، كما يدعم الملفات المصدّرة من ZipGrade. وتستطيع الأداة تحليل تنسيقات متعددة والتعرّف تلقائيًا على بيانات الطلاب والأسئلة والدرجات ونواتج التعلم والحقول ذات الصلة لإعداد التقارير.',
        'faq.q_specs': 'ما ملفات توصيف المقررات التي يدعمها إتقان؟',
        'faq.a_specs': 'يدعم إتقان قوالب توصيف المقررات المعتمدة من هيئة تقويم التعليم والتدريب NCAAA، ويستطيع استخراج بيانات المقرر ونواتج التعلم منها تلقائيًا لتقليل الإدخال اليدوي.',
        'faq.q_unsupported': 'ماذا أفعل إذا لم يكن ملفي مدعوماً؟',
        'faq.a_unsupported': 'إذا لم يكن ملفك مدعوماً، تواصل معنا وأرسل عينة من الملف. سنراجع بنيته ونعمل على دعمه في أقرب وقت. نسعى باستمرار إلى توسيع نطاق الملفات المدعومة لتلبية احتياجات أعضاء هيئة التدريس المختلفة.',
        'faq.q_save_courses': 'لماذا أحتاج إلى حفظ المقررات؟',
        'faq.a_save_courses': 'يتيح لك حفظ المقرر تخزين تفاصيله ونواتج التعلم في ملفك الشخصي، مما يوفر عليك وقت وجهد إعادة إدخالها في كل مرة تقوم فيها بإنشاء تقرير جديد.',
        'faq.q_visual_identity': 'هل يمكنني تعديل الهوية البصرية المستخدمة في التقارير المصدّرة؟',
        'faq.a_visual_identity': 'يمكنك استخدام هذه الهوية الافتراضية أو تعديلها برفع شعار بصيغة JPEG واختيار الألوان.',
        'faq.q_language': 'هل يدعم إتقان اللغة العربية؟',
        'faq.a_language': 'نعم. يدعم إتقان اللغتين العربية والإنجليزية، ويمكنه معالجة ملفات الدرجات وتوصيفات المقررات باللغتين.',
        'checkout.title': 'إتمام الدفع',
        'checkout.selected_plan': 'الخطة المحددة',
        'checkout.quantity': 'عدد التقارير',
        'checkout.total': 'التكلفة الإجمالية',
        'checkout.billing_info': 'معلومات الدفع',
        'checkout.card_name': 'الاسم على البطاقة',
        'checkout.card_number': 'رقم البطاقة',
        'checkout.expiry': 'تاريخ الانتهاء',
        'checkout.cvv': 'رمز الأمان',
        'checkout.pay': 'ادفع الآن',
    }
}

TRANSLATIONS['ar'].update({
    'nav.my_exams': '\u0627\u062e\u062a\u0628\u0627\u0631\u0627\u062a\u064a',
    'question_mapping.question_type': '\u0646\u0648\u0639 \u0627\u0644\u0633\u0624\u0627\u0644',
        'question_mapping.paper_clo_help': '',
    'question_mapping.step2_title': '\u0627\u0644\u062e\u0637\u0648\u0629 2: \u0627\u0633\u062a\u062e\u0631\u0627\u062c \u0627\u0644\u0623\u0633\u0626\u0644\u0629 \u0648\u0627\u0644\u0631\u0628\u0637 \u0627\u0644\u0635\u0631\u064a\u062d',
    'question_mapping.step2_description': '\u0631\u0627\u062c\u0639 \u0627\u0644\u0623\u0633\u0626\u0644\u0629 \u0627\u0644\u0645\u0633\u062a\u062e\u0631\u062c\u0629\u060c \u0648\u0623\u0646\u0648\u0627\u0639\u0647\u0627\u060c \u0648\u0646\u0648\u0627\u062a\u062d \u0627\u0644\u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0630\u0643\u0648\u0631\u0629 \u0635\u0631\u0627\u062d\u0629 \u0641\u064a \u0648\u0631\u0642\u0629 \u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631.',
    'question_mapping.total_questions': '\u0625\u062c\u0645\u0627\u0644\u064a \u0627\u0644\u0623\u0633\u0626\u0644\u0629',
    'question_mapping.auto_mapped_questions': '\u0623\u0633\u0626\u0644\u0629 \u0645\u0631\u0628\u0648\u0637\u0629',
    'question_mapping.ai_required_questions': '\u062a\u062d\u062a\u0627\u062c \u0631\u0628\u0637',
    'question_mapping.question_number': '\u0631\u0642\u0645 \u0627\u0644\u0633\u0624\u0627\u0644',
    'question_mapping.explicit_clo': '\u0646\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'question_mapping.status': '\u0627\u0644\u062d\u0627\u0644\u0629',
    'question_mapping.mapped_automatically': '\u0631\u064f\u0628\u0637 \u062a\u0644\u0642\u0627\u0626\u064a\u0627\u064b',
    'question_mapping.requires_ai_mapping': '\u064a\u062d\u062a\u0627\u062c \u0631\u0628\u0637',
    'question_mapping.all_mapped_success': '\u062a\u0645 \u0631\u0628\u0637 \u062c\u0645\u064a\u0639 \u0627\u0644\u0623\u0633\u0626\u0644\u0629 \u0628\u0646\u062c\u0627\u062d.',
    'question_mapping.final_review': '\u0627\u0644\u062e\u0637\u0648\u0629 4: \u0627\u0644\u0645\u0631\u0627\u062c\u0639\u0629 \u0627\u0644\u0646\u0647\u0627\u0626\u064a\u0629',
    'question_mapping.continue_ai_mapping': '\u0627\u0644\u062a\u0627\u0644\u064a',
    'question_mapping.step3_title': 'الخطوة 3: ربط الأسئلة بنواتج التعلم',
    'question_mapping.step3_description': '\u0631\u0627\u062c\u0639 \u0627\u0642\u062a\u0631\u0627\u062d\u0627\u062a \u0627\u0644\u0630\u0643\u0627\u0621 \u0627\u0644\u0627\u0635\u0637\u0646\u0627\u0639\u064a \u0644\u0644\u0623\u0633\u0626\u0644\u0629 \u0627\u0644\u062a\u064a \u0644\u0645 \u064a\u0630\u0643\u0631 \u0641\u064a\u0647\u0627 \u0645\u062e\u0631\u062c \u062a\u0639\u0644\u0645 \u0628\u0634\u0643\u0644 \u0635\u0631\u064a\u062d.',
    'question_mapping.suggested_clo': '\u0645\u062e\u0631\u062c \u0627\u0644\u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0642\u062a\u0631\u062d',
    'question_mapping.ai_suggestions': 'اقتراحات الذكاء الاصطناعي',
    'question_mapping.high_conf': 'عالي',
    'question_mapping.medium_conf': 'متوسط',
    'question_mapping.low_conf': 'منخفض',
    'question_mapping.alt_suggestions': 'اقتراحات بديلة:',
    'question_mapping.no_ai_suggestions': 'لا توجد اقتراحات متاحة.',
    'question_mapping.ai_diagnostics_title': '\u062a\u0634\u062e\u064a\u0635 \u0627\u0642\u062a\u0631\u0627\u062d\u0627\u062a \u0627\u0644\u0631\u0628\u0637',
    'question_mapping.ai_diagnostics_help': '\u064a\u0648\u0636\u062d \u0647\u0630\u0627 \u0647\u0644 \u0623\u0646\u062a\u062c Gemini \u0623\u0648 Qwen \u0623\u0648 \u0627\u0644\u0631\u0628\u0637 \u0627\u0644\u0645\u062d\u0644\u064a \u0627\u0644\u0627\u0642\u062a\u0631\u0627\u062d\u0627\u062a.',
    'question_mapping.no_suggestion_reason': '\u0644\u0645 \u064a\u0646\u062a\u062c \u0627\u0642\u062a\u0631\u0627\u062d \u0644\u0647\u0630\u0627 \u0627\u0644\u0633\u0624\u0627\u0644. \u0631\u0627\u062c\u0639 \u062a\u0634\u062e\u064a\u0635 \u0627\u0644\u0645\u0632\u0648\u062f\u0627\u062a \u0623\u0639\u0644\u0627\u0647.',
    'question_mapping.final_selection': 'النواتج المختارة',
    'question_mapping.add_clo': 'إضافة ناتج تعلم',
    'question_mapping.confidence_score': '\u062f\u0631\u062c\u0629 \u0627\u0644\u062b\u0642\u0629',
    'question_mapping.select_clo': '\u0627\u062e\u062a\u0631 \u0645\u062e\u0631\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'question_mapping.choose_clo': '\u0627\u062e\u062a\u0631 \u0645\u062e\u0631\u062c\u0627\u064b...',
    'question_mapping.all_course_clos': '\u062c\u0645\u064a\u0639 \u0645\u062e\u0631\u062c\u0627\u062a \u0627\u0644\u0645\u0642\u0631\u0631',
    'question_mapping.save_mapping': '\u062d\u0641\u0638',
    'question_mapping.back': '\u0639\u0648\u062f\u0629',
    'question_mapping.review_saved': 'تم الحفظ بنجاح.',
    'exams.title': '\u0627\u062e\u062a\u0628\u0627\u0631\u0627\u062a\u064a',
    'exams.empty': '\u0644\u0627 \u062a\u0648\u062c\u062f \u0627\u062e\u062a\u0628\u0627\u0631\u0627\u062a \u0645\u062d\u0641\u0648\u0638\u0629 \u0628\u0639\u062f.',
    'exams.saved': 'تم حفظ التقرير بنجاح.',
    'exams.exam': '\u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631',
    'exams.course': '\u0627\u0644\u0645\u0642\u0631\u0631',
    'exams.questions': '\u0627\u0644\u0623\u0633\u0626\u0644\u0629',
    'exams.created': '\u062a\u0627\u0631\u064a\u062e \u0627\u0644\u0625\u0646\u0634\u0627\u0621',
})

TRANSLATIONS['ar'].update({
    'mapping.method_title': '\u0627\u062e\u062a\u064a\u0627\u0631 \u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u0631\u0628\u0637',
    'mapping.method_description': '\u062d\u062f\u062f \u0643\u064a\u0641 \u062a\u0631\u064a\u062f \u0631\u0628\u0637 \u0623\u0633\u0626\u0644\u0629 \u0627\u0644\u062a\u0642\u064a\u064a\u0645 \u0628\u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645.',
    'mapping.method_choice': '\u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u0631\u0628\u0637',
    'mapping.method_manual_title': '\u0631\u0628\u0637 \u064a\u062f\u0648\u064a',
    'mapping.method_manual_description': 'حدد نواتج التعلم لكل سؤال يدويا.',
    'mapping.method_ai_title': '\u0631\u0628\u0637 \u0628\u0627\u0633\u062a\u062e\u062f\u0627\u0645 \u0627\u0644\u0630\u0643\u0627\u0621 \u0627\u0644\u0627\u0635\u0637\u0646\u0627\u0639\u064a',
    'mapping.method_ai_description': 'اختر تقرير ربط تم إنشاؤه مسبقًا لإعادة استخدام ربط الأسئلة بنواتج التعلم.',
    'mapping.method_exam_paper_help': '\u0627\u0631\u0641\u0639 \u0648\u0631\u0642\u0629 \u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631 \u0627\u0644\u0645\u0631\u062a\u0628\u0637\u0629 \u0628\u0623\u0633\u0626\u0644\u0629 \u0645\u0644\u0641\u0627\u062a \u0627\u0644\u062f\u0631\u062c\u0627\u062a.',
    'mapping.method_continue': '\u0627\u0644\u062a\u0627\u0644\u064a',
})

TRANSLATIONS['ar'].update({
    'validation.complete_required': '\u064a\u0631\u062c\u0649 \u0625\u0643\u0645\u0627\u0644 \u0627\u0644\u062d\u0642\u0648\u0644 \u0627\u0644\u0645\u0637\u0644\u0648\u0628\u0629.',
    'validation.select_item': '\u064a\u0631\u062c\u0649 \u0627\u062e\u062a\u064a\u0627\u0631 \u0639\u0646\u0635\u0631 \u0645\u0646 \u0627\u0644\u0642\u0627\u0626\u0645\u0629.',
    'validation.fill_field': '\u064a\u0631\u062c\u0649 \u062a\u0639\u0628\u0626\u0629 \u0647\u0630\u0627 \u0627\u0644\u062d\u0642\u0644.',
    'validation.select_file': '\u064a\u0631\u062c\u0649 \u0627\u062e\u062a\u064a\u0627\u0631 \u0645\u0644\u0641.',
    'validation.valid_email': '\u064a\u0631\u062c\u0649 \u0625\u062f\u062e\u0627\u0644 \u0628\u0631\u064a\u062f \u0625\u0644\u0643\u062a\u0631\u0648\u0646\u064a \u0635\u062d\u064a\u062d.',
    'validation.match_format': '\u064a\u0631\u062c\u0649 \u0625\u062f\u062e\u0627\u0644 \u0642\u064a\u0645\u0629 \u0628\u0627\u0644\u0635\u064a\u063a\u0629 \u0627\u0644\u0645\u0637\u0644\u0648\u0628\u0629.',
    'validation.number': '\u064a\u0631\u062c\u0649 \u0625\u062f\u062e\u0627\u0644 \u0631\u0642\u0645 \u0635\u062d\u064a\u062d.',
    'validation.range_underflow': '\u0627\u0644\u0642\u064a\u0645\u0629 \u0623\u0642\u0644 \u0645\u0646 \u0627\u0644\u062d\u062f \u0627\u0644\u0645\u0633\u0645\u0648\u062d.',
    'validation.range_overflow': '\u0627\u0644\u0642\u064a\u0645\u0629 \u0623\u0639\u0644\u0649 \u0645\u0646 \u0627\u0644\u062d\u062f \u0627\u0644\u0645\u0633\u0645\u0648\u062d.',
    'validation.too_short': '\u0627\u0644\u0642\u064a\u0645\u0629 \u0642\u0635\u064a\u0631\u0629 \u062c\u062f\u0627\u064b.',
    'validation.too_long': '\u0627\u0644\u0642\u064a\u0645\u0629 \u0637\u0648\u064a\u0644\u0629 \u062c\u062f\u0627\u064b.',
    'validation.step_mismatch': '\u064a\u0631\u062c\u0649 \u0625\u062f\u062e\u0627\u0644 \u0642\u064a\u0645\u0629 \u0635\u062d\u064a\u062d\u0629.',
    'validation.duplicate_course_report': '\u062a\u0645\u062a \u0625\u0636\u0627\u0641\u0629 \u062a\u0642\u0631\u064a\u0631 \u0647\u0630\u0627 \u0627\u0644\u0645\u0642\u0631\u0631 \u0645\u0633\u0628\u0642\u0627\u064b.',
})

SUPPORTED_LANGUAGES = {'en', 'ar'}

def get_language():
    lang = session.get('language', 'en')
    return lang if lang in SUPPORTED_LANGUAGES else 'en'

def translate(key):
    lang = get_language()
    if lang == 'en':
        return EN_TRANSLATIONS.get(key, key)
    return TRANSLATIONS.get(lang, {}).get(key, key)

AR_FLASH_EXACT_TRANSLATIONS = {
    'Exam deleted successfully.': 'تم حذف الاختبار بنجاح.',
    'No course report file uploaded': '\u0644\u0645 \u064a\u062a\u0645 \u0631\u0641\u0639 \u0645\u0644\u0641 \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
    'No selected course report file': '\u0644\u0645 \u064a\u062a\u0645 \u0627\u062e\u062a\u064a\u0627\u0631 \u0645\u0644\u0641 \u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
    'Please upload a PDF, CSV, or Excel course report.': '\u064a\u0631\u062c\u0649 \u0631\u0641\u0639 \u062a\u0642\u0631\u064a\u0631 \u0645\u0642\u0631\u0631 \u0628\u0635\u064a\u063a\u0629 PDF \u0623\u0648 CSV \u0623\u0648 Excel.',
    'Please enter an assessment name.': '\u064a\u0631\u062c\u0649 \u0625\u062f\u062e\u0627\u0644 \u0627\u0633\u0645 \u0627\u0644\u062a\u0642\u064a\u064a\u0645.',
    'Please upload an assessment file.': '\u064a\u0631\u062c\u0649 \u0631\u0641\u0639 \u0645\u0644\u0641 \u0627\u0644\u062a\u0642\u064a\u064a\u0645.',
    'Invalid file format. Please upload PDF, DOCX, TXT, CSV, or Excel.': '\u0635\u064a\u063a\u0629 \u0627\u0644\u0645\u0644\u0641 \u063a\u064a\u0631 \u0635\u062d\u064a\u062d\u0629. \u064a\u0631\u062c\u0649 \u0631\u0641\u0639 PDF \u0623\u0648 DOCX \u0623\u0648 TXT \u0623\u0648 CSV \u0623\u0648 Excel.',
    'Course specification must be uploaded as a PDF or Word document (.docx).': '\u064a\u062c\u0628 \u0631\u0641\u0639 \u062a\u0648\u0635\u064a\u0641 \u0627\u0644\u0645\u0642\u0631\u0631 \u0628\u0635\u064a\u063a\u0629 PDF \u0623\u0648 Word (.docx).',
    'No saved CLO attainment report was selected.': '\u0644\u0645 \u064a\u062a\u0645 \u062a\u062d\u062f\u064a\u062f \u0623\u064a \u062a\u0642\u0631\u064a\u0631 \u0645\u062d\u0641\u0648\u0638 \u0644\u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645.',
    'Invalid question mapping draft.': '\u0645\u0633\u0648\u062f\u0629 \u0631\u0628\u0637 \u0627\u0644\u0623\u0633\u0626\u0644\u0629 \u063a\u064a\u0631 \u0635\u062d\u064a\u062d\u0629.',
    'Question mapping draft was not found.': '\u0644\u0645 \u064a\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0645\u0633\u0648\u062f\u0629 \u0631\u0628\u0637 \u0627\u0644\u0623\u0633\u0626\u0644\u0629.',
    'No CLOs were found for the selected course. Add or update the course through My Courses.': '\u0644\u0645 \u064a\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0646\u0648\u0627\u062a\u062c \u062a\u0639\u0644\u0645 \u0644\u0644\u0645\u0642\u0631\u0631 \u0627\u0644\u0645\u062d\u062f\u062f. \u0623\u0636\u0641 \u0627\u0644\u0645\u0642\u0631\u0631 \u0623\u0648 \u062d\u062f\u062b\u0647 \u0645\u0646 \u0635\u0641\u062d\u0629 \u0645\u0642\u0631\u0631\u0627\u062a\u064a.',
}

AR_FLASH_PREFIX_TRANSLATIONS = {
    'Could not read course specification:': '\u062a\u0639\u0630\u0631 \u0642\u0631\u0627\u0621\u0629 \u062a\u0648\u0635\u064a\u0641 \u0627\u0644\u0645\u0642\u0631\u0631:',
    'Could not read course specification PDF:': '\u062a\u0639\u0630\u0631 \u0642\u0631\u0627\u0621\u0629 \u0645\u0644\u0641 PDF \u0644\u062a\u0648\u0635\u064a\u0641 \u0627\u0644\u0645\u0642\u0631\u0631:',
    'Error reading exam paper:': '\u062d\u062f\u062b \u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0642\u0631\u0627\u0621\u0629 \u0648\u0631\u0642\u0629 \u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631:',
    'Error reading file:': '\u062d\u062f\u062b \u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0645\u0644\u0641:',
    'A required file for the course report could not be found.': '\u062a\u0639\u0630\u0631 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0645\u0644\u0641 \u0645\u0637\u0644\u0648\u0628 \u0644\u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u0642\u0631\u0631.',
    'The saved report is missing required field:': '\u0627\u0644\u062a\u0642\u0631\u064a\u0631 \u0627\u0644\u0645\u062d\u0641\u0648\u0638 \u064a\u0641\u062a\u0642\u062f \u062d\u0642\u0644\u0627\u064b \u0645\u0637\u0644\u0648\u0628\u0627\u064b:',
}

def localized_flash_message(message):
    message = str(message or '')
    if not has_request_context() or get_language() != 'ar':
        return message
    if message in AR_FLASH_EXACT_TRANSLATIONS:
        return AR_FLASH_EXACT_TRANSLATIONS[message]
    for prefix, translated_prefix in AR_FLASH_PREFIX_TRANSLATIONS.items():
        if message.startswith(prefix):
            return translated_prefix + message[len(prefix):]
    return message

def count_unit(kind, count):
    try:
        count = int(count)
    except (TypeError, ValueError):
        count = 0

    language = get_language() if has_request_context() else 'en'
    if language == 'ar':
        if kind == 'question':
            return 'أسئلة' if 2 <= count <= 10 else 'سؤال'
        if kind == 'student':
            return 'طلاب' if 2 <= count <= 10 else 'طالب'
    if kind == 'question':
        return 'question' if count == 1 else 'questions'
    if kind == 'student':
        return 'student' if count == 1 else 'students'
    return ''

def get_export_report_language(user=None):
    user = user if user is not None else (current_user() if has_request_context() else None)
    setting = ''
    if user and 'report_language' in user.keys():
        setting = str(user['report_language'] or '').strip()
    if setting in {'en', 'ar'}:
        return setting
    return 'en'

PDF_REPORT_LABELS = {
    'en': {
        'title': 'CLO Attainment Report',
        'university': 'University',
        'department': 'Department',
        'course_name': 'Course Name',
        'course_id': 'Course ID',
        'report_date': 'Report Date',
        'total_students': 'Total Students Evaluated',
        'mapped_clos': 'Mapped CLOs',
        'clo_definitions': 'CLO Definitions',
        'domain': 'Domain',
        'clo': 'CLO',
        'wording': 'Wording',
        'summary': 'CLO Achievement Summary',
        'questions': 'Questions',
        'max': 'Max',
        'target': 'Target',
        'achieved': 'No of Students Achieved',
        'achieved_status': 'Achieved',
        'student_achieved_status': 'Met',
        'achievement': 'Achievement',
        'student_achievement': 'Student CLO Achievement',
        'student_id': 'Student ID',
        'not_achieved': 'Not Achieved',
        'student_not_achieved_status': 'Not met',
        'na': 'N/A',
    },
    'ar': {
        'title': 'تقرير تحقق نواتج التعلم',
        'university': 'الجامعة',
        'department': 'القسم',
        'course_name': 'اسم المقرر',
        'course_id': 'رمز المقرر',
        'report_date': 'تاريخ التقرير',
        'total_students': 'إجمالي الطلاب',
        'mapped_clos': 'نواتج التعلم المرتبطة',
        'clo_definitions': 'نواتج التعلم',
        'domain': 'المجال',
        'clo': 'الرمز',
        'wording': 'نواتج التعلم',
        'summary': 'ملخص تحقق نواتج التعلم',
        'questions': 'الأسئلة',
        'max': 'الدرجة الكلية',
        'target': 'المستوى المستهدف',
        'achieved': 'عدد الطلاب المحققين',
        'achieved_status': 'محقق',
        'student_achieved_status': 'محقق',
        'achievement': 'نسبة التحقق',
        'student_achievement': 'تحقق نواتج التعلم لكل طالب',
        'student_id': 'معرّف الطالب',
        'not_achieved': 'غير محقق',
        'student_not_achieved_status': 'غير محقق',
        'na': 'غير متوفر',
    }
}

PDF_REPORT_LABELS['ar'] = {
    'title': '\u062a\u0642\u0631\u064a\u0631 \u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'university': '\u0627\u0644\u062c\u0627\u0645\u0639\u0629',
    'department': '\u0627\u0644\u0642\u0633\u0645',
    'course_name': '\u0627\u0633\u0645 \u0627\u0644\u0645\u0642\u0631\u0631',
    'course_id': '\u0631\u0645\u0632 \u0627\u0644\u0645\u0642\u0631\u0631',
    'report_date': '\u062a\u0627\u0631\u064a\u062e \u0627\u0644\u062a\u0642\u0631\u064a\u0631',
    'total_students': '\u0625\u062c\u0645\u0627\u0644\u064a \u0627\u0644\u0637\u0644\u0627\u0628',
    'mapped_clos': '\u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645 \u0627\u0644\u0645\u0631\u062a\u0628\u0637\u0629',
    'clo_definitions': '\u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'domain': '\u0627\u0644\u0645\u062c\u0627\u0644',
    'clo': '\u0627\u0644\u0631\u0645\u0632',
    'wording': '\u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'summary': '\u0645\u0644\u062e\u0635 \u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645',
    'questions': '\u0627\u0644\u0623\u0633\u0626\u0644\u0629',
    'max': '\u0627\u0644\u062f\u0631\u062c\u0629 \u0627\u0644\u0643\u0644\u064a\u0629',
    'target': '\u0627\u0644\u0645\u0633\u062a\u0648\u0649 \u0627\u0644\u0645\u0633\u062a\u0647\u062f\u0641',
    'achieved': '\u0639\u062f\u062f \u0627\u0644\u0637\u0644\u0627\u0628 \u0627\u0644\u0645\u062d\u0642\u0642\u064a\u0646',
    'achieved_status': '\u0645\u062d\u0642\u0642',
    'student_achieved_status': '\u0645\u062d\u0642\u0642',
    'achievement': '\u0646\u0633\u0628\u0629 \u0627\u0644\u062a\u062d\u0642\u0642',
    'student_achievement': '\u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645 \u0644\u0643\u0644 \u0637\u0627\u0644\u0628',
    'student_id': '\u0645\u0639\u0631\u0651\u0641 \u0627\u0644\u0637\u0627\u0644\u0628',
    'not_achieved': '\u063a\u064a\u0631 \u0645\u062d\u0642\u0642',
    'student_not_achieved_status': '\u063a\u064a\u0631 \u0645\u062d\u0642\u0642',
    'na': '\u063a\u064a\u0631 \u0645\u062a\u0648\u0641\u0631',
}

AR_CLO_DOMAIN_LABELS = {
    'Knowledge': '1.0 \u0627\u0644\u0645\u0639\u0631\u0641\u0629 \u0648\u0627\u0644\u0641\u0647\u0645',
    'Skills': '2.0 \u0627\u0644\u0645\u0647\u0627\u0631\u0627\u062a',
    'Values': '3.0 \u0627\u0644\u0642\u064a\u0645',
    'Other': '\u0623\u062e\u0631\u0649',
}

def pdf_report_labels(language=None):
    language = language if language in {'en', 'ar'} else get_export_report_language()
    return PDF_REPORT_LABELS.get(language, PDF_REPORT_LABELS['en'])

def get_localized_university_options():
    language = get_language()
    return [
        {
            'value': university,
            'label': UNIVERSITY_ARABIC_NAMES.get(university, university) if language == 'ar' else university
        }
        for university in UNIVERSITY_CHOICES
    ]

def localized_university_name(value, language=None):
    value = str(value or '').strip()
    if not value:
        return ''
    language = language or get_language()
    canonical = canonical_university_name(value)
    if language == 'ar':
        return UNIVERSITY_ARABIC_NAMES.get(canonical, value)
    return canonical or value

def localized_clo_domain(domain, language=None):
    domain = str(domain or '').strip()
    language = language or get_language()
    if language == 'ar':
        return AR_CLO_DOMAIN_LABELS.get(domain, domain)
    if language == 'ar':
        return {
            'Knowledge': '1.0 المعرفة والفهم',
            'Skills': '2.0 المهارات',
            'Values': '3.0 القيم',
            'Other': 'أخرى',
        }.get(domain, domain)
    return {
        'Knowledge': '1.0 Knowledge',
        'Skills': '2.0 Skills',
        'Values': '3.0 Values',
        'Other': 'Other',
    }.get(domain, domain)

def canonical_university_name(value):
    value = str(value or '').strip()
    if value in UNIVERSITY_CHOICES:
        return value
    return UNIVERSITY_ENGLISH_BY_ARABIC.get(value, value)

def get_registration_university_name():
    choice = (request.form.get('university_choice') or '').strip()
    if choice == '__other__':
        return canonical_university_name(request.form.get('other_university_name') or '')
    if choice in UNIVERSITY_CHOICES:
        return choice
    return ''

def external_url_for(endpoint, **values):
    if APP_PUBLIC_URL:
        return f"{APP_PUBLIC_URL}{url_for(endpoint, **values)}"
    return url_for(endpoint, _external=True, _scheme='https', **values)

def get_profile_university_name():
    choice = (request.form.get('university_choice') or '').strip()
    if choice == '__other__':
        return canonical_university_name(request.form.get('other_university_name') or '')
    if choice in UNIVERSITY_CHOICES:
        return choice
    return canonical_university_name(request.form.get('university_name') or '')

def is_email_configured():
    return bool((RESEND_API_KEY and RESEND_FROM_EMAIL) or (SMTP_HOST and SMTP_FROM_EMAIL))

def send_email(recipient_email, subject, text_body, html_body=''):
    if RESEND_API_KEY and RESEND_FROM_EMAIL:
        payload = {
            'from': f"{SMTP_FROM_NAME} <{RESEND_FROM_EMAIL}>",
            'to': [recipient_email],
            'subject': subject,
            'text': text_body,
        }
        if html_body:
            payload['html'] = html_body
        request_data = json.dumps(payload).encode('utf-8')
        request = urllib.request.Request(
            'https://api.resend.com/emails',
            data=request_data,
            headers={
                'Authorization': f"Bearer {RESEND_API_KEY}",
                'Content-Type': 'application/json',
            },
            method='POST'
        )
        try:
            with urllib.request.urlopen(request, timeout=SMTP_TIMEOUT_SECONDS) as response:
                status = response.getcode()
                if 200 <= status < 300:
                    return True, ''
                return False, f'resend_status_{status}'
        except urllib.error.HTTPError as exc:
            app.logger.error("Resend email failed: HTTP %s %s", exc.code, exc.read().decode('utf-8', errors='replace'))
            return False, 'send_failed'
        except Exception:
            app.logger.exception("Resend email failed")
            return False, 'send_failed'

    if not (SMTP_HOST and SMTP_FROM_EMAIL):
        return False, 'missing_config'
    message = EmailMessage()
    message['Subject'] = subject
    message['From'] = f"{SMTP_FROM_NAME} <{SMTP_FROM_EMAIL}>"
    message['To'] = recipient_email
    message.set_content(text_body)
    if html_body:
        message.add_alternative(html_body, subtype='html')
    try:
        if SMTP_USE_SSL:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT_SECONDS, context=ssl.create_default_context()) as server:
                if SMTP_USERNAME and SMTP_PASSWORD:
                    server.login(SMTP_USERNAME, SMTP_PASSWORD)
                server.send_message(message)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT_SECONDS) as server:
                if SMTP_USE_TLS:
                    server.starttls(context=ssl.create_default_context())
                if SMTP_USERNAME and SMTP_PASSWORD:
                    server.login(SMTP_USERNAME, SMTP_PASSWORD)
                server.send_message(message)
    except Exception:
        app.logger.exception("SMTP email failed")
        return False, 'send_failed'
    return True, ''

def send_password_reset_email(recipient_email, reset_url):
    if not is_email_configured():
        return False, 'missing_config'

    text_body = (
        "Hello,\n\n"
        "We received a request to reset your ETQAN password.\n\n"
        f"Use this link within 1 hour:\n{reset_url}\n\n"
        "If you did not request this, you can ignore this email."
    )
    html_body = f"""
    <p>Hello,</p>
    <p>We received a request to reset your ETQAN password.</p>
    <p><a href="{reset_url}">Reset your password</a></p>
    <p>This link expires in 1 hour.</p>
    <p>If you did not request this, you can ignore this email.</p>
    """
    return send_email(recipient_email, 'Reset your ETQAN password', text_body, html_body)

def get_upload_path(filename):
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    return os.path.join(app.config['UPLOAD_FOLDER'], filename)

def find_command(command, fallback_paths=None):
    found = shutil.which(command)
    if found:
        return found
    for path in fallback_paths or []:
        if path and os.path.exists(path):
            return path
    return ''

class PostgresConnection:
    def __init__(self):
        if not DATABASE_URL:
            raise RuntimeError("DATABASE_URL is required. ETQAN now uses PostgreSQL only.")
        if psycopg2 is None:
            raise RuntimeError("psycopg2 is not installed. Install psycopg2-binary before starting ETQAN.")
        self.conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.DictCursor)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type:
            self.conn.rollback()
        else:
            self.conn.commit()
        self.conn.close()

    def execute(self, sql, params=None):
        sql = self._translate_sql(sql)
        cursor = self.conn.cursor()
        cursor.execute(sql, params or ())
        return cursor

    @staticmethod
    def _translate_sql(sql):
        translated = sql.replace('?', '%s')
        translated = translated.replace('ORDER BY display_name COLLATE NOCASE', 'ORDER BY LOWER(display_name)')
        return translated


def get_db():
    return PostgresConnection()


def get_table_columns(conn, table_name):
    validate_sql_identifier(table_name)
    rows = conn.execute(
        """
        SELECT column_name
          FROM information_schema.columns
         WHERE table_schema = 'public'
           AND table_name = ?
        """,
        (table_name,)
    ).fetchall()
    return {row['column_name'] for row in rows}


def validate_sql_identifier(value):
    value = str(value or '')
    if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', value):
        raise ValueError("Invalid SQL identifier.")
    return value


def add_missing_columns(conn, table_name, columns):
    table_name = validate_sql_identifier(table_name)
    existing_columns = get_table_columns(conn, table_name)
    for column, definition in columns.items():
        column = validate_sql_identifier(column)
        if column not in existing_columns:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column} {definition}")


def init_postgres_db(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL,
            university_name TEXT DEFAULT '',
            college TEXT DEFAULT '',
            department TEXT DEFAULT '',
            org_primary_color TEXT DEFAULT '#26365f',
            org_secondary_color TEXT DEFAULT '#9d6b16',
            org_tertiary_color TEXT DEFAULT '',
            org_logo_stored_name TEXT DEFAULT '',
            org_logo_original_name TEXT DEFAULT '',
            report_language TEXT DEFAULT 'interface',
            billing_plan TEXT DEFAULT 'free',
            report_credits INTEGER DEFAULT 0,
            subscription_started_at TEXT DEFAULT '',
            reset_token TEXT DEFAULT '',
            reset_token_expires_at TEXT DEFAULT ''
        )
    """)
    add_missing_columns(conn, 'users', {
        'university_name': "TEXT DEFAULT ''",
        'college': "TEXT DEFAULT ''",
        'department': "TEXT DEFAULT ''",
        'org_primary_color': "TEXT DEFAULT '#26365f'",
        'org_secondary_color': "TEXT DEFAULT '#9d6b16'",
        'org_tertiary_color': "TEXT DEFAULT ''",
        'org_logo_stored_name': "TEXT DEFAULT ''",
        'org_logo_original_name': "TEXT DEFAULT ''",
        'report_language': "TEXT DEFAULT 'interface'",
        'billing_plan': "TEXT DEFAULT 'free'",
        'report_credits': "INTEGER DEFAULT 0",
        'subscription_started_at': "TEXT DEFAULT ''",
        'reset_token': "TEXT DEFAULT ''",
        'reset_token_expires_at': "TEXT DEFAULT ''"
    })
    conn.execute("""
        CREATE TABLE IF NOT EXISTS saved_reports (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            title TEXT NOT NULL,
            course_name TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            report_hash TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(user_id, report_hash)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS saved_exams (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            title TEXT NOT NULL,
            course_name TEXT NOT NULL,
            filename TEXT DEFAULT '',
            payload_json TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_courses (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            display_name TEXT NOT NULL,
            course_name TEXT NOT NULL,
            course_code TEXT DEFAULT '',
            college TEXT DEFAULT '',
            program TEXT DEFAULT '',
            department TEXT DEFAULT '',
            clos_json TEXT NOT NULL,
            target_percentages_json TEXT NOT NULL,
            topics_json TEXT DEFAULT '[]',
            clo_plos_json TEXT DEFAULT '{}',
            extraction_metadata_json TEXT DEFAULT '{}',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(user_id, display_name)
        )
    """)
    add_missing_columns(conn, 'user_courses', {
        'college': "TEXT DEFAULT ''",
        'program': "TEXT DEFAULT ''",
        'topics_json': "TEXT DEFAULT '[]'",
        'clo_plos_json': "TEXT DEFAULT '{}'",
        'extraction_metadata_json': "TEXT DEFAULT '{}'"
    })
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_programs (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            display_name TEXT NOT NULL,
            program_name TEXT NOT NULL,
            program_code TEXT DEFAULT '',
            college TEXT DEFAULT '',
            department TEXT DEFAULT '',
            plos_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(user_id, display_name)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS contact_requests (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            name TEXT DEFAULT '',
            email TEXT DEFAULT '',
            organization TEXT DEFAULT '',
            college TEXT DEFAULT '',
            department TEXT DEFAULT '',
            enquiry_type TEXT DEFAULT '',
            message TEXT DEFAULT '',
            attachment_stored_name TEXT DEFAULT '',
            attachment_original_name TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS login_failures (
            id SERIAL PRIMARY KEY,
            email TEXT DEFAULT '',
            ip_address TEXT DEFAULT '',
            attempted_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS login_locks (
            id SERIAL PRIMARY KEY,
            lock_type TEXT NOT NULL,
            lock_key TEXT NOT NULL,
            locked_until TIMESTAMP NOT NULL,
            lock_count INTEGER DEFAULT 1,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(lock_type, lock_key)
        )
    """)


def init_db():
    with get_db() as conn:
        init_postgres_db(conn)

init_db()

def is_unique_violation(exc):
    return psycopg2 is not None and isinstance(exc, psycopg2.IntegrityError)

def ensure_default_professional_account(user):
    if not user:
        return user
    started_at = str(user['subscription_started_at'] or '').strip()
    professional_active = False
    if user['billing_plan'] == 'professional' and started_at:
        try:
            professional_active = datetime.now() < datetime.strptime(started_at, "%Y-%m-%d") + timedelta(days=365)
        except ValueError:
            professional_active = False
    if professional_active:
        return user
    with get_db() as conn:
        conn.execute(
            """
            UPDATE users
               SET billing_plan = 'professional',
                   subscription_started_at = ?
             WHERE id = ?
            """,
            (datetime.now().strftime("%Y-%m-%d"), user['id'])
        )
        return conn.execute("SELECT * FROM users WHERE id = ?", (user['id'],)).fetchone()


def current_user():
    user_id = session.get('user_id')
    if not user_id:
        return None
    with get_db() as conn:
        user = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return ensure_default_professional_account(user)

def is_admin_user(user=None):
    user = user or current_user()
    if not user:
        return False
    return str(user['email'] or '').strip().lower() in ADMIN_EMAILS

def get_client_ip():
    forwarded = request.headers.get('X-Forwarded-For', '')
    if forwarded:
        return forwarded.split(',')[0].strip()
    return request.remote_addr or ''

def login_window_start():
    return datetime.utcnow() - timedelta(minutes=LOGIN_FAILURE_WINDOW_MINUTES)

def get_active_login_lock(conn, lock_type, lock_key):
    if not lock_key:
        return None
    return conn.execute(
        """
        SELECT *
          FROM login_locks
         WHERE lock_type = ?
           AND lock_key = ?
           AND locked_until > CURRENT_TIMESTAMP
        """,
        (lock_type, lock_key)
    ).fetchone()

def upsert_login_lock(conn, lock_type, lock_key, minutes):
    if not lock_key:
        return
    locked_until = datetime.utcnow() + timedelta(minutes=minutes)
    now = datetime.utcnow()
    conn.execute(
        """
        INSERT INTO login_locks (lock_type, lock_key, locked_until, lock_count, created_at, updated_at)
        VALUES (?, ?, ?, 1, ?, ?)
        ON CONFLICT (lock_type, lock_key)
        DO UPDATE SET
            locked_until = EXCLUDED.locked_until,
            lock_count = login_locks.lock_count + 1,
            updated_at = EXCLUDED.updated_at
        """,
        (lock_type, lock_key, locked_until, now, now)
    )

def login_lock_minutes_for_email(conn, email):
    row = conn.execute(
        "SELECT lock_count FROM login_locks WHERE lock_type = 'email' AND lock_key = ?",
        (email,)
    ).fetchone()
    if row and int(row['lock_count'] or 0) >= 1:
        return LOGIN_ESCALATED_LOCK_MINUTES
    return LOGIN_LOCK_MINUTES

def evaluate_login_lock(conn, email, ip_address):
    email_lock = get_active_login_lock(conn, 'email', email)
    if email_lock:
        return 'email', email_lock
    ip_lock = get_active_login_lock(conn, 'ip', ip_address)
    if ip_lock:
        return 'ip', ip_lock
    return '', None

def record_failed_login(conn, email, ip_address):
    attempted_at = datetime.utcnow()
    conn.execute(
        "INSERT INTO login_failures (email, ip_address, attempted_at) VALUES (?, ?, ?)",
        (email, ip_address, attempted_at)
    )
    window_start = login_window_start()
    email_failures = conn.execute(
        "SELECT COUNT(*) AS count FROM login_failures WHERE email = ? AND attempted_at >= ?",
        (email, window_start)
    ).fetchone()['count'] if email else 0
    ip_failures = conn.execute(
        "SELECT COUNT(*) AS count FROM login_failures WHERE ip_address = ? AND attempted_at >= ?",
        (ip_address, window_start)
    ).fetchone()['count'] if ip_address else 0
    distinct_email_failures = conn.execute(
        "SELECT COUNT(DISTINCT email) AS count FROM login_failures WHERE ip_address = ? AND attempted_at >= ? AND email <> ''",
        (ip_address, window_start)
    ).fetchone()['count'] if ip_address else 0

    if email and email_failures >= LOGIN_EMAIL_FAILURE_LIMIT:
        upsert_login_lock(conn, 'email', email, login_lock_minutes_for_email(conn, email))
    if ip_address and (ip_failures >= LOGIN_IP_FAILURE_LIMIT or distinct_email_failures >= LOGIN_DISTINCT_EMAIL_LIMIT):
        upsert_login_lock(conn, 'ip', ip_address, LOGIN_LOCK_MINUTES)

def clear_login_failures(conn, email, ip_address):
    window_start = login_window_start()
    conn.execute(
        "DELETE FROM login_failures WHERE (email = ? OR ip_address = ?) AND attempted_at >= ?",
        (email, ip_address, window_start)
    )


@app.context_processor
def inject_global_template_data():
    language = get_language()
    user = current_user()
    return {
        'current_user': user,
        'is_admin': is_admin_user(user),
        'report_branding': get_report_branding(),
        'university_choices': UNIVERSITY_CHOICES,
        'university_options': get_localized_university_options(),
        'canonical_university_name': canonical_university_name,
        'language': language,
        'page_dir': 'rtl' if language == 'ar' else 'ltr',
        '_': translate,
        'count_unit': count_unit,
        'resolve_detected_clos_to_course_list': resolve_detected_clos_to_course_list,
        'display_student_id': display_student_id,
        'text_direction': text_direction,
        'course_report_label': course_report_label,
        'localized_flash_message': localized_flash_message,
    }

def is_valid_hex_color(value):
    return bool(re.match(r'^#[0-9A-Fa-f]{6}$', str(value or '').strip()))

def normalize_brand_color(value):
    value = str(value or '').strip()
    return value if is_valid_hex_color(value) else '#26365f'

def normalize_secondary_color(value):
    value = str(value or '').strip()
    return value if is_valid_hex_color(value) else '#9d6b16'

def normalize_optional_brand_color(value):
    value = str(value or '').strip()
    return value if is_valid_hex_color(value) else ''

def distinct_brand_colors(*values):
    colors = []
    seen = set()
    for value in values:
        color = normalize_optional_brand_color(value)
        key = color.upper()
        if color and key not in seen:
            colors.append(color)
            seen.add(key)
    return colors

def normalize_palette_colors(values):
    if not isinstance(values, list):
        return []
    normalized = []
    seen = set()
    for value in values:
        color = normalize_optional_brand_color(value)
        key = color.upper()
        if color and key not in seen:
            normalized.append(color)
            seen.add(key)
    return normalized

def display_colors_from_palette(colors):
    colors = normalize_palette_colors(colors)
    if len(colors) <= 3:
        return colors
    return [colors[0], colors[-2], colors[-1]]

def resolve_palette_display_colors(identity):
    display_colors = normalize_palette_colors(identity.get('display_palette_colors'))
    if display_colors:
        return display_colors[:3]
    return display_colors_from_palette(identity.get('palette_colors'))

def load_university_identity_presets():
    if not os.path.exists(UNIVERSITY_IDENTITY_PATH):
        return {}
    try:
        with open(UNIVERSITY_IDENTITY_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        str(name): value
        for name, value in data.items()
        if isinstance(value, dict)
    }

def get_university_identity(university_name):
    university_name = canonical_university_name(university_name)
    presets = load_university_identity_presets()
    identity = dict(presets.get(university_name) or {})
    has_identity_preset = bool(identity) or university_name in UNIVERSITY_COLOR_PRESETS
    identity.setdefault('website', '')
    identity.setdefault('resolved_website', '')
    identity.setdefault('logo_filename', '')
    identity.setdefault('logo_url', '')
    identity['has_identity_preset'] = has_identity_preset
    if has_identity_preset:
        identity['primary_color'] = normalize_brand_color(
            identity.get('primary_color') or UNIVERSITY_COLOR_PRESETS.get(university_name)
        )
        identity['secondary_color'] = normalize_optional_brand_color(identity.get('secondary_color'))
    else:
        identity['primary_color'] = normalize_optional_brand_color(identity.get('primary_color'))
        identity['secondary_color'] = normalize_optional_brand_color(identity.get('secondary_color'))
    identity['tertiary_color'] = normalize_optional_brand_color(identity.get('tertiary_color'))
    identity['palette_colors'] = normalize_palette_colors(identity.get('palette_colors'))
    identity['display_palette_colors'] = normalize_palette_colors(identity.get('display_palette_colors'))
    palette_display_colors = resolve_palette_display_colors(identity)
    if palette_display_colors:
        deduped_display_colors = distinct_brand_colors(*palette_display_colors)
        identity['primary_color'] = normalize_brand_color(deduped_display_colors[0] if deduped_display_colors else identity.get('primary_color'))
        identity['secondary_color'] = deduped_display_colors[1] if len(deduped_display_colors) > 1 else ''
        identity['tertiary_color'] = deduped_display_colors[2] if len(deduped_display_colors) > 2 else ''
    else:
        deduped_identity_colors = distinct_brand_colors(
            identity.get('primary_color'),
            identity.get('secondary_color'),
            identity.get('tertiary_color')
        )
        identity['primary_color'] = normalize_brand_color(deduped_identity_colors[0] if deduped_identity_colors else identity.get('primary_color'))
        identity['secondary_color'] = deduped_identity_colors[1] if len(deduped_identity_colors) > 1 else ''
        identity['tertiary_color'] = deduped_identity_colors[2] if len(deduped_identity_colors) > 2 else ''
    return identity

def get_university_identity_options():
    return {
        university: get_university_identity(university)
        for university in UNIVERSITY_CHOICES
    }

def get_public_logo_path(filename, report_ready=False):
    safe_filename = os.path.basename(str(filename or '').strip())
    if not safe_filename:
        return ''
    logo_path = os.path.join(UNIVERSITY_LOGO_FOLDER, safe_filename)
    if not os.path.exists(logo_path) or os.path.getsize(logo_path) <= 500:
        return ''
    if report_ready and os.path.splitext(safe_filename)[1].lower() not in {'.jpg', '.jpeg', '.png'}:
        return ''
    return logo_path

def get_organization_logo_path(filename):
    safe_filename = os.path.basename(str(filename or '').strip())
    if not safe_filename:
        return ''
    persistent_path = os.path.join(ORG_LOGO_FOLDER, safe_filename)
    if os.path.exists(persistent_path):
        return persistent_path
    legacy_path = get_upload_path(safe_filename)
    if os.path.exists(legacy_path):
        return legacy_path
    return persistent_path

def delete_organization_logo(filename):
    safe_filename = os.path.basename(str(filename or '').strip())
    if not safe_filename:
        return
    for folder in (ORG_LOGO_FOLDER, app.config['UPLOAD_FOLDER']):
        path = os.path.join(folder, safe_filename)
        if os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass

def get_branding_logo_url(branding):
    branding = branding or {}
    logo_stored_name = branding.get('logo_stored_name')
    if logo_stored_name:
        safe_filename = os.path.basename(str(logo_stored_name))
        if os.path.exists(os.path.join(ORG_LOGO_FOLDER, safe_filename)):
            return f"/organization-logos/{safe_filename}"
        if os.path.exists(get_upload_path(safe_filename)):
            return f"/uploads/{safe_filename}"
    logo_public_filename = branding.get('logo_public_filename')
    if logo_public_filename:
        return f"/university-logos/{os.path.basename(str(logo_public_filename))}"
    return ''

def resolve_branding_logo_path(branding, report_ready=False, legacy_pdf=False):
    logo_stored_name = branding.get('logo_stored_name') if branding else ''
    if logo_stored_name:
        uploaded_path = get_organization_logo_path(logo_stored_name)
        if os.path.exists(uploaded_path):
            return uploaded_path
    public_logo_path = get_public_logo_path((branding or {}).get('logo_public_filename'), report_ready=report_ready)
    if legacy_pdf and os.path.splitext(public_logo_path)[1].lower() not in {'.jpg', '.jpeg'}:
        return ''
    return public_logo_path

def get_image_data_uri(image_path):
    if not image_path or not os.path.exists(image_path):
        return ''
    ext = os.path.splitext(image_path)[1].lower()
    mime_type = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.svg': 'image/svg+xml',
    }.get(ext, 'application/octet-stream')
    with open(image_path, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('ascii')
    return f"data:{mime_type};base64,{encoded}"

def get_report_branding():
    user = current_user()
    if user:
        organization_name = user['university_name'] or ''
        organization_key = canonical_university_name(organization_name)
        identity = get_university_identity(organization_key)
        has_identity_preset = bool(identity.get('has_identity_preset'))
        stored_color = str(user['org_primary_color'] or '').strip()
        legacy_colors = LEGACY_UNIVERSITY_COLOR_DEFAULTS.get(organization_key, set())
        stored_secondary_color = str(user['org_secondary_color'] or '').strip() if 'org_secondary_color' in user.keys() else ''
        stored_tertiary_color = str(user['org_tertiary_color'] or '').strip() if 'org_tertiary_color' in user.keys() else ''
        if has_identity_preset:
            primary_color = normalize_brand_color(
                stored_color
                if stored_color and stored_color != '#26365f' and stored_color.upper() not in {color.upper() for color in legacy_colors}
                else identity.get('primary_color')
            )
            secondary_color = normalize_optional_brand_color(
                stored_secondary_color if stored_secondary_color and stored_secondary_color != '#9d6b16' else identity.get('secondary_color')
            )
        else:
            primary_color = '' if stored_color.lower() == '#26365f' else normalize_optional_brand_color(stored_color)
            secondary_color = '' if stored_secondary_color.lower() == '#9d6b16' else normalize_optional_brand_color(stored_secondary_color)
        tertiary_color = normalize_optional_brand_color(stored_tertiary_color or identity.get('tertiary_color'))
        uploaded_logo = user['org_logo_stored_name'] or ''
        payload = {
            'primary_color': primary_color,
            'secondary_color': secondary_color,
            'tertiary_color': tertiary_color,
            'logo_stored_name': uploaded_logo,
            'logo_original_name': user['org_logo_original_name'] or '',
            'logo_public_filename': '' if uploaded_logo else identity.get('logo_filename', ''),
            'logo_source_url': identity.get('logo_url', ''),
            'organization_website': identity.get('resolved_website') or identity.get('website', ''),
            'organization_name': organization_name,
            'department': user['department'] or ''
        }
        payload['logo_preview_url'] = get_branding_logo_url(payload)
        return payload
    branding = session.get('report_branding') or {}
    organization_name = branding.get('organization_name', '')
    identity = get_university_identity(canonical_university_name(organization_name))
    if identity.get('has_identity_preset'):
        primary_color = normalize_brand_color(branding.get('primary_color') or identity.get('primary_color'))
        secondary_color = normalize_optional_brand_color(branding.get('secondary_color') or identity.get('secondary_color'))
    else:
        stored_primary_color = str(branding.get('primary_color') or '').strip()
        stored_secondary_color = str(branding.get('secondary_color') or '').strip()
        primary_color = '' if stored_primary_color.lower() == '#26365f' else normalize_optional_brand_color(stored_primary_color)
        secondary_color = '' if stored_secondary_color.lower() == '#9d6b16' else normalize_optional_brand_color(stored_secondary_color)
    tertiary_color = normalize_optional_brand_color(branding.get('tertiary_color') or identity.get('tertiary_color'))
    uploaded_logo = branding.get('logo_stored_name', '')
    payload = {
        'primary_color': primary_color,
        'secondary_color': secondary_color,
        'tertiary_color': tertiary_color,
        'logo_stored_name': uploaded_logo,
        'logo_original_name': branding.get('logo_original_name', ''),
        'logo_public_filename': '' if uploaded_logo else identity.get('logo_filename', ''),
        'logo_source_url': identity.get('logo_url', ''),
        'organization_website': identity.get('resolved_website') or identity.get('website', ''),
        'organization_name': organization_name,
        'department': branding.get('department', '')
    }
    payload['logo_preview_url'] = get_branding_logo_url(payload)
    return payload

def update_report_branding_from_request():
    user = current_user()
    if not user:
        raise ValueError("Please login and complete your organization profile before exporting reports.")
    branding = get_report_branding()
    university_name = canonical_university_name(user['university_name'] or '')
    if not university_name:
        raise ValueError("Please enter your university name before exporting reports.")
    identity = get_university_identity(university_name)
    custom_university = not identity.get('has_identity_preset')
    primary_empty = request.form.get('brand_primary_color_empty') == '1'
    secondary_empty = request.form.get('brand_secondary_color_empty') == '1'
    tertiary_empty = request.form.get('brand_tertiary_color_empty') == '1'
    primary_custom = request.form.get('brand_primary_color_custom') == '1'
    secondary_custom = request.form.get('brand_secondary_color_custom') == '1'
    tertiary_custom = request.form.get('brand_tertiary_color_custom') == '1'
    if custom_university:
        branding['primary_color'] = '' if primary_empty else normalize_optional_brand_color(request.form.get('brand_primary_color'))
        branding['secondary_color'] = '' if secondary_empty else normalize_optional_brand_color(request.form.get('brand_secondary_color'))
        branding['tertiary_color'] = '' if tertiary_empty else normalize_optional_brand_color(request.form.get('brand_tertiary_color'))
    else:
        branding['primary_color'] = (
            normalize_brand_color(request.form.get('brand_primary_color'))
            if primary_custom
            else normalize_brand_color(identity.get('primary_color'))
        )
        branding['secondary_color'] = (
            ('' if secondary_empty else normalize_optional_brand_color(request.form.get('brand_secondary_color')))
            if secondary_custom
            else normalize_optional_brand_color(identity.get('secondary_color'))
        )
        branding['tertiary_color'] = (
            ('' if tertiary_empty else normalize_optional_brand_color(request.form.get('brand_tertiary_color')))
            if tertiary_custom
            else normalize_optional_brand_color(identity.get('tertiary_color'))
        )
    branding['organization_name'] = university_name
    branding['department'] = user['department'] or ''
    logo_file = request.files.get('report_logo')
    if logo_file and logo_file.filename:
        logo_ext = os.path.splitext(logo_file.filename)[1].lower()
        if logo_ext not in {'.jpg', '.jpeg'}:
            raise ValueError("Report logo must be uploaded as a JPEG file so it can be included in the PDF report.")
        old_logo = branding.get('logo_stored_name')
        logo_stored_name = f"{uuid.uuid4()}{logo_ext}"
        logo_file.save(os.path.join(ORG_LOGO_FOLDER, logo_stored_name))
        branding['logo_stored_name'] = logo_stored_name
        branding['logo_original_name'] = logo_file.filename
        if old_logo:
            delete_organization_logo(old_logo)
    with get_db() as conn:
        conn.execute(
            """
            UPDATE users
               SET university_name = ?,
                   org_primary_color = ?,
                   org_secondary_color = ?,
                   org_tertiary_color = ?,
                   org_logo_stored_name = ?,
                   org_logo_original_name = ?
             WHERE id = ?
            """,
            (
                university_name,
                branding['primary_color'],
                branding['secondary_color'],
                branding['tertiary_color'],
                branding.get('logo_stored_name', ''),
                branding.get('logo_original_name', ''),
                user['id']
            )
        )
    return branding

def organization_profile_complete(user=None):
    user = user or current_user()
    return bool(user and str(user['university_name'] or '').strip())

def require_export_profile():
    user = current_user()
    if not user:
        flash("Please login before exporting reports so the report can use your organization's visual identity.", "error")
        return False
    if not organization_profile_complete(user):
        flash("Please enter your university name before exporting reports.", "error")
        return False
    return True

def hex_to_pdf_rgb(value):
    color = normalize_brand_color(value).lstrip('#')
    return tuple(round(int(color[index:index + 2], 16) / 255, 3) for index in (0, 2, 4))

def pdf_rgb_command(value, operator='rg'):
    red, green, blue = hex_to_pdf_rgb(value)
    return f"{red} {green} {blue} {operator}"

def json_safe(value):
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if hasattr(value, 'item'):
        return value.item()
    return value

def get_display_branding_payload():
    branding = get_report_branding()
    payload = dict(branding)
    logo_path = resolve_branding_logo_path(branding)
    if logo_path:
        payload['logo_data_uri'] = get_image_data_uri(logo_path)
    return payload

def merge_with_current_branding(saved_branding=None):
    saved_branding = dict(saved_branding or {})
    current_branding = get_report_branding()
    merged = dict(current_branding)
    for key, value in saved_branding.items():
        if value not in (None, ''):
            merged[key] = value

    for color_key in ('primary_color', 'secondary_color', 'tertiary_color'):
        saved_color = str(saved_branding.get(color_key) or '').strip()
        if not saved_color or saved_color.lower() in {'#26365f', '#9d6b16'}:
            merged[color_key] = current_branding.get(color_key) or saved_color
    merged['primary_color'] = normalize_brand_color(merged.get('primary_color') or current_branding.get('primary_color'))
    merged['secondary_color'] = normalize_optional_brand_color(merged.get('secondary_color') or current_branding.get('secondary_color'))
    merged['tertiary_color'] = normalize_optional_brand_color(merged.get('tertiary_color') or current_branding.get('tertiary_color'))
    merged['logo_preview_url'] = get_branding_logo_url(merged)
    return merged

def apply_university_identity_colors(branding=None):
    branding = dict(branding or get_report_branding())
    organization_name = branding.get('organization_name') or get_profile_university_name()
    identity = get_university_identity(canonical_university_name(organization_name))
    if identity.get('has_identity_preset'):
        branding['primary_color'] = normalize_brand_color(identity.get('primary_color'))
        branding['secondary_color'] = normalize_optional_brand_color(identity.get('secondary_color'))
        branding['tertiary_color'] = normalize_optional_brand_color(identity.get('tertiary_color'))
        if not branding.get('logo_stored_name'):
            branding['logo_public_filename'] = identity.get('logo_filename', '')
            branding['logo_source_url'] = identity.get('logo_url', '')
        branding['organization_website'] = identity.get('resolved_website') or identity.get('website') or branding.get('organization_website', '')
    return branding

def compute_color_luminance(hex_color):
    hex_color = str(hex_color or '#FFFFFF').lstrip('#')
    if len(hex_color) == 3:
        hex_color = ''.join([c*2 for c in hex_color])
    try:
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    except ValueError:
        return 1.0
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0

def get_contrast_text_color(hex_color):
    return '#FFFFFF' if compute_color_luminance(hex_color) < 0.6 else '#000000'

def get_standard_table_style(primary_color_hex, is_arabic, num_rows, header_rows=1):
    from reportlab.lib import colors
    from reportlab.platypus import TableStyle
    
    header_text_color_hex = get_contrast_text_color(primary_color_hex)
    try:
        primary_color = colors.HexColor(primary_color_hex)
    except Exception:
        primary_color = colors.HexColor('#26365f')
    try:
        header_text_color = colors.HexColor(header_text_color_hex)
    except Exception:
        header_text_color = colors.white
        
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, header_rows - 1), primary_color),
        ('TEXTCOLOR', (0, 0), (-1, header_rows - 1), header_text_color),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT' if is_arabic else 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.45, colors.HexColor('#cbd5e1')),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]
    
    for row in range(header_rows, num_rows):
        bg_color = colors.HexColor('#f8fafc') if (row - header_rows) % 2 == 1 else colors.white
        style_commands.append(('BACKGROUND', (0, row), (-1, row), bg_color))
        
    return TableStyle(style_commands)

def get_standard_paragraph_styles(primary_color_hex, is_arabic, regular_font, bold_font):
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib import colors
    
    try:
        primary_color = colors.HexColor(primary_color_hex)
    except Exception:
        primary_color = colors.HexColor('#26365f')
        
    header_text_color_hex = get_contrast_text_color(primary_color_hex)
    try:
        header_text_color = colors.HexColor(header_text_color_hex)
    except Exception:
        header_text_color = colors.white

    base_styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'ReportTitle', parent=base_styles['Title'], fontName=bold_font, fontSize=17,
        leading=21, textColor=primary_color, alignment=TA_RIGHT if is_arabic else TA_LEFT,
        spaceAfter=12
    )
    meta_style = ParagraphStyle(
        'ReportMeta', parent=base_styles['Normal'], fontName=regular_font, fontSize=10,
        leading=13, textColor=colors.black, alignment=TA_RIGHT if is_arabic else TA_LEFT, spaceAfter=2
    )
    section_style = ParagraphStyle(
        'ReportSection', parent=base_styles['Heading2'], fontName=bold_font, fontSize=12,
        leading=15, textColor=primary_color, alignment=TA_RIGHT if is_arabic else TA_LEFT,
        spaceBefore=12, spaceAfter=6
    )
    table_header = ParagraphStyle(
        'TableHeader', parent=base_styles['Normal'], fontName=bold_font, fontSize=10,
        textColor=header_text_color, alignment=TA_RIGHT if is_arabic else TA_LEFT
    )
    table_text = ParagraphStyle(
        'TableText', parent=base_styles['Normal'], fontName=regular_font, fontSize=10,
        textColor=colors.black, alignment=TA_RIGHT if is_arabic else TA_LEFT
    )
    
    return {
        'title': title_style,
        'meta': meta_style,
        'section': section_style,
        'table_header': table_header,
        'table_text': table_text
    }
def get_saved_report_count(user_id):
    with get_db() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM saved_reports WHERE user_id = ?",
            (user_id,)
        ).fetchone()[0]

def get_saved_report_count_since(user_id, start_date):
    if not start_date:
        return get_saved_report_count(user_id)
    with get_db() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM saved_reports WHERE user_id = ? AND created_at >= ?",
            (user_id, start_date)
        ).fetchone()[0]

def subscription_active(user):
    if not user or user['billing_plan'] not in {'academic', 'professional'}:
        return False
    started_at = str(user['subscription_started_at'] or '').strip()
    if not started_at:
        return False
    try:
        started = datetime.strptime(started_at, "%Y-%m-%d")
    except ValueError:
        return False
    return datetime.now() < started + timedelta(days=365)

def get_billing_status(user=None):
    user = user or current_user()
    if not user:
        return {
            'saved_report_count': 0,
            'free_limit': FREE_REPORT_LIMIT,
            'free_remaining': FREE_REPORT_LIMIT,
            'report_credits': 0,
            'billing_plan': 'free',
            'yearly_active': False,
            'academic_report_limit': ACADEMIC_REPORT_LIMIT_PER_YEAR,
            'academic_reports_used': 0,
            'academic_reports_remaining': ACADEMIC_REPORT_LIMIT_PER_YEAR,
        }
    saved_count = get_saved_report_count(user['id'])
    active_plan = user['billing_plan'] if subscription_active(user) else 'free'
    academic_reports_used = 0
    if active_plan == 'academic':
        academic_reports_used = get_saved_report_count_since(user['id'], user['subscription_started_at'])
    return {
        'saved_report_count': saved_count,
        'free_limit': FREE_REPORT_LIMIT,
        'free_remaining': max(FREE_REPORT_LIMIT - saved_count, 0),
        'report_credits': int(user['report_credits'] or 0),
        'billing_plan': user['billing_plan'] or 'free',
        'active_plan': active_plan,
        'yearly_active': active_plan in {'academic', 'professional'},
        'academic_report_limit': ACADEMIC_REPORT_LIMIT_PER_YEAR,
        'academic_reports_used': academic_reports_used,
        'academic_reports_remaining': max(ACADEMIC_REPORT_LIMIT_PER_YEAR - academic_reports_used, 0),
    }

def load_saved_report_payload(report_id, user_id):
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM saved_reports WHERE id = ? AND user_id = ?",
            (report_id, user_id)
        ).fetchone()
    if not row:
        return None, None
    return row, normalize_saved_report_payload(safe_json_loads(row_get(row, 'payload_json'), {}))

def normalize_saved_report_payload(payload):
    if not isinstance(payload, dict):
        return {}
    payload = dict(payload)
    stats = payload.get('stats') or {}
    if isinstance(stats, dict) and isinstance(stats.get('clo_overall'), dict):
        stats = stats.get('clo_overall') or {}
    if not isinstance(stats, dict):
        stats = {}
    payload['stats'] = stats
    payload['course_info'] = payload.get('course_info') if isinstance(payload.get('course_info'), dict) else {}
    payload['total_students'] = safe_int_value(payload.get('total_students'))
    return payload

def row_get(row, key, default=''):
    try:
        return row[key]
    except (KeyError, IndexError, TypeError):
        return default

def safe_json_loads(value, default=None):
    try:
        return json.loads(value or '')
    except (TypeError, ValueError, json.JSONDecodeError):
        return default

def safe_int_value(value, default=0):
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return default

def safe_float_value(value, default=0.0):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return default

def load_selected_report_payloads(report_ids, user_id):
    cleaned_ids = []
    for report_id in report_ids or []:
        try:
            cleaned_ids.append(int(report_id))
        except (TypeError, ValueError):
            continue
    cleaned_ids = list(dict.fromkeys(cleaned_ids))
    if not cleaned_ids:
        return []

    reports = []
    with get_db() as conn:
        placeholders = ','.join('?' for _ in cleaned_ids)
        rows = conn.execute(
            f"SELECT * FROM saved_reports WHERE user_id = ? AND id IN ({placeholders}) ORDER BY id DESC",
            [user_id] + cleaned_ids
        ).fetchall()
    rows_by_id = {safe_int_value(row_get(row, 'id')): row for row in rows}
    for report_id in cleaned_ids:
        row = rows_by_id.get(report_id)
        if not row:
            continue
        reports.append({'row': row, 'payload': normalize_saved_report_payload(safe_json_loads(row_get(row, 'payload_json'), {}))})
    return reports

def aggregate_course_report_payloads(report_payloads, selected_course_name=''):
    if not report_payloads:
        return {}, {}, 0, []

    combined_stats = {}
    total_students = 0
    selected_reports = []
    course_info = {}
    first_report_course_name = ''

    for item in report_payloads:
        row = (item or {}).get('row') or {}
        payload = normalize_saved_report_payload((item or {}).get('payload') or {})
        if not first_report_course_name:
            first_report_course_name = row_get(row, 'course_name')
        report_title = display_saved_report_title(row)
        report_total_students = safe_int_value(payload.get('total_students'))
        total_students += report_total_students
        selected_reports.append({
            'id': row_get(row, 'id'),
            'display_title': report_title,
            'created_at': row_get(row, 'created_at'),
            'total_students': report_total_students,
        })
        if not course_info:
            row_course_name = row_get(row, 'course_name')
            course_info = payload.get('course_info') or {'course_name': row_course_name, 'raw_name': row_course_name}

        for clo, data in (payload.get('stats') or {}).items():
            if not isinstance(data, dict):
                continue
            target = combined_stats.setdefault(clo, {
                'questions': [],
                'students_achieved': 0,
                'total_possible_score': 0.0,
                'target_score': 0.0,
                'target_pct_total': 0.0,
                'target_pct_weight': 0,
                '_total_students': 0,
            })
            questions = data.get('questions') or []
            target['questions'].extend([f"{report_title}: {question}" for question in questions])
            target['students_achieved'] += safe_int_value(data.get('students_achieved'))
            target['total_possible_score'] += safe_float_value(data.get('total_possible_score'))
            target['target_score'] += safe_float_value(data.get('target_score'))
            target['target_pct_total'] += safe_float_value(data.get('target_pct')) * max(report_total_students, 1)
            target['target_pct_weight'] += max(report_total_students, 1)
            target['_total_students'] += report_total_students

    for data in combined_stats.values():
        clo_total_students = int(data.pop('_total_students', 0) or 0)
        target_pct_weight = int(data.pop('target_pct_weight', 0) or 0)
        target_pct_total = float(data.pop('target_pct_total', 0) or 0)
        data['target_pct'] = round(target_pct_total / target_pct_weight, 2) if target_pct_weight else 0
        data['achievement_percentage'] = round((data['students_achieved'] / clo_total_students) * 100, 2) if clo_total_students else 0
        data['total_possible_score'] = round(float(data.get('total_possible_score') or 0), 2)
        data['target_score'] = round(float(data.get('target_score') or 0), 2)

    course_info = enrich_course_info_from_course(
        course_info,
        selected_course_name or first_report_course_name
    )

    return dict(sorted_clo_items(combined_stats)), course_info, total_students, selected_reports

def normalize_clo_summary_item(data):
    if not isinstance(data, dict):
        return None
    questions = data.get('questions') or data.get('mapped_questions') or []
    if isinstance(questions, str):
        questions = [questions]
    if not isinstance(questions, list):
        questions = []
    return {
        'questions': [str(question or '').strip() for question in questions if str(question or '').strip()],
        'students_achieved': safe_int_value(data.get('students_achieved', data.get('achieved_students', 0))),
        'total_possible_score': safe_float_value(data.get('total_possible_score', data.get('max_score', 0))),
        'target_score': safe_float_value(data.get('target_score', data.get('target', 0))),
        'target_pct': safe_float_value(data.get('target_pct', data.get('target_percentage', 0))),
        'achievement_percentage': safe_float_value(data.get('achievement_percentage', data.get('attainment_percentage', 0))),
    }

def extract_saved_report_stats(payload):
    if not isinstance(payload, dict):
        return {}
    candidates = [
        payload.get('stats'),
        (payload.get('stats') or {}).get('clo_overall') if isinstance(payload.get('stats'), dict) else None,
        payload.get('clo_overall'),
        payload.get('results'),
    ]
    for candidate in candidates:
        if isinstance(candidate, dict):
            if isinstance(candidate.get('clo_overall'), dict):
                candidate = candidate.get('clo_overall') or {}
            usable = {
                str(clo): value
                for clo, value in candidate.items()
                if isinstance(value, dict)
            }
            if usable:
                return usable
    return {}

def course_report_template_context_defaults(export_action, selected_course_name=''):
    return {
        'export_action': export_action,
        'selected_report_ids': [],
        'selected_course_name': selected_course_name or '',
        'selected_reports': [],
        'course_info': enrich_course_info_from_course(
            {'course_name': selected_course_name or '', 'raw_name': selected_course_name or ''},
            selected_course_name or ''
        ),
        'course_topics': [],
        'course_improvement_recommendations': COURSE_IMPROVEMENT_RECOMMENDATIONS,
        'course_improvement_recommendation_groups': grouped_course_improvement_recommendations(),
        'course_improvement_action_options': COURSE_IMPROVEMENT_ACTION_OPTIONS,
        'course_improvement_support_options': COURSE_IMPROVEMENT_SUPPORT_OPTIONS,
        'uncovered_reason_actions': UNCOVERED_TOPIC_REASON_ACTIONS,
        'uncovered_reason_actions_json': json.dumps(localized_uncovered_reason_actions(), ensure_ascii=False),
        'total_students': 0,
        'stats_items': [],
        'course_report_warnings': [],
    }

def course_report_input_error_message(exc, report_ids=None, selected_course_name=''):
    report_ids_text = ', '.join(str(report_id) for report_id in (report_ids or []) if str(report_id).strip()) or 'none'
    course_text = str(selected_course_name or '').strip() or 'not provided'
    if isinstance(exc, ValueError):
        return f"{exc} Course: {course_text}. Report ID(s): {report_ids_text}."
    if isinstance(exc, FileNotFoundError):
        return f"A required file for the course report could not be found. Course: {course_text}. Report ID(s): {report_ids_text}."
    if isinstance(exc, KeyError):
        return f"The saved report is missing required field: {exc}. Course: {course_text}. Report ID(s): {report_ids_text}."
    return f"Course report input error ({type(exc).__name__}): {exc}. Course: {course_text}. Report ID(s): {report_ids_text}."

def load_course_report_records(report_ids, user_id):
    cleaned_ids = []
    for report_id in report_ids or []:
        try:
            cleaned_ids.append(int(report_id))
        except (TypeError, ValueError):
            continue
    cleaned_ids = list(dict.fromkeys(cleaned_ids))
    if not cleaned_ids:
        return []

    placeholders = ','.join('?' for _ in cleaned_ids)
    with get_db() as conn:
        rows = conn.execute(
            f"""
            SELECT id, title, course_name, payload_json, created_at
              FROM saved_reports
             WHERE user_id = ?
               AND id IN ({placeholders})
             ORDER BY id DESC
            """,
            [user_id] + cleaned_ids
        ).fetchall()

    rows_by_id = {safe_int_value(row_get(row, 'id')): row for row in rows}
    records = []
    for report_id in cleaned_ids:
        row = rows_by_id.get(report_id)
        if not row:
            continue
        payload = normalize_saved_report_payload(safe_json_loads(row_get(row, 'payload_json'), {}))
        records.append({
            'id': safe_int_value(row_get(row, 'id')),
            'title': row_get(row, 'title') or 'CLO Attainment Report',
            'display_title': display_saved_report_title(row),
            'course_name': row_get(row, 'course_name'),
            'created_at': row_get(row, 'created_at'),
            'payload': payload,
        })
    return records



def is_report_saved(stats, total_students, student_achievement_matrix, course_info):
    user = current_user()
    if not user:
        return False
    payload = {
        'stats': json_safe(stats),
        'total_students': total_students,
        'student_achievement_matrix': json_safe(student_achievement_matrix),
        'course_info': json_safe(course_info),
        'branding': get_display_branding_payload()
    }
    report_hash = hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
    with get_db() as conn:
        existing = conn.execute(
            "SELECT id FROM saved_reports WHERE user_id = ? AND report_hash = ?",
            (user['id'], report_hash)
        ).fetchone()
        return bool(existing)

def build_course_report_context_from_records(records, export_action, selected_course_name=''):
    context = course_report_template_context_defaults(export_action, selected_course_name)
    combined_stats = {}
    selected_reports = []
    total_students_candidates = []
    course_info = {}
    first_course_name = selected_course_name or ''
    selected_report_ids = []

    for record in records or []:
        record = record or {}
        payload = record.get('payload') or {}
        if not isinstance(payload, dict):
            payload = {}
        report_id = safe_int_value(record.get('id'))
        if report_id:
            selected_report_ids.append(report_id)
        report_title = record.get('display_title') or record.get('title') or 'CLO Attainment Report'
        report_course_name = record.get('course_name') or ''
        if not first_course_name:
            first_course_name = report_course_name

        report_students = safe_int_value(
            payload.get('total_students')
            or payload.get('students')
            or payload.get('total_students_evaluated')
        )
        if report_students:
            total_students_candidates.append(report_students)

        selected_reports.append({
            'id': report_id,
            'display_title': report_title,
            'created_at': record.get('created_at') or '',
            'total_students': report_students,
        })

        if not course_info:
            payload_course_info = payload.get('course_info') if isinstance(payload.get('course_info'), dict) else {}
            course_info = dict(payload_course_info or {})
            if not course_info:
                course_info = {'course_name': report_course_name, 'raw_name': report_course_name}

        report_stats = extract_saved_report_stats(payload)
        for clo, raw_data in report_stats.items():
            data = normalize_clo_summary_item(raw_data)
            if not data:
                continue
            target = combined_stats.setdefault(str(clo), {
                'questions': [],
                'students_achieved': 0,
                'total_possible_score': 0.0,
                'target_score': 0.0,
                'target_pct_total': 0.0,
                'target_pct_weight': 0,
                '_total_students': 0,
            })
            target['questions'].extend([f"{report_title}: {question}" for question in data['questions']])
            target['students_achieved'] += data['students_achieved']
            target['total_possible_score'] += data['total_possible_score']
            target['target_score'] += data['target_score']
            student_weight = report_students or safe_int_value(raw_data.get('total_students') if isinstance(raw_data, dict) else 0) or 1
            target['target_pct_total'] += data['target_pct'] * student_weight
            target['target_pct_weight'] += student_weight
            target['_total_students'] += student_weight

    for data in combined_stats.values():
        clo_total_students = int(data.pop('_total_students', 0) or 0)
        target_pct_weight = int(data.pop('target_pct_weight', 0) or 0)
        target_pct_total = float(data.pop('target_pct_total', 0) or 0)
        data['target_pct'] = round(target_pct_total / target_pct_weight, 2) if target_pct_weight else 0
        data['achievement_percentage'] = round((data['students_achieved'] / clo_total_students) * 100, 2) if clo_total_students else 0
        data['total_possible_score'] = round(float(data.get('total_possible_score') or 0), 2)
        data['target_score'] = round(float(data.get('target_score') or 0), 2)

    selected_course_name = selected_course_name or first_course_name
    if not course_info:
        course_info = {'course_name': selected_course_name, 'raw_name': selected_course_name}
    course_info = enrich_course_info_from_course(course_info, selected_course_name)
    raw_course_name = course_info.get('raw_name') or course_info.get('course_name') or selected_course_name
    try:
        course_topics = get_course_topics(raw_course_name)
    except Exception:
        app.logger.exception("Failed to load course topics for course report inputs")
        course_topics = []
    warnings = []
    if selected_reports and not combined_stats:
        warnings.append("The selected report was loaded, but no CLO summary rows were found in its saved data.")
    if not course_info.get('course_name') and not course_info.get('raw_name'):
        warnings.append("Course information could not be resolved from the selected report.")

    context.update({
        'selected_report_ids': selected_report_ids,
        'selected_course_name': selected_course_name or raw_course_name,
        'selected_reports': selected_reports,
        'course_info': course_info,
        'course_topics': course_topics,
        'course_report_warnings': warnings,
        'total_students': max(total_students_candidates) if total_students_candidates else 0,
        'stats_items': sorted_clo_items(combined_stats),
    })
    return context

def render_course_report_inputs_from_records(records, export_action, selected_course_name='', grade_distribution_provided=False):
    context = build_course_report_context_from_records(records, export_action, selected_course_name)
    if not context.get('selected_report_ids'):
        raise ValueError("No saved CLO attainment report was selected.")
    if not context.get('stats_items'):
        raise ValueError(
            "The selected CLO attainment report does not contain readable CLO summary data. "
            "Create a new CLO Attainment Analysis report, then select it for the course report."
        )
    return render_template('course_report_inputs.html', grade_distribution_provided=grade_distribution_provided, **context)

def display_saved_report_title(row):
    title = str(row_get(row, 'title') or '').strip()
    course_name = str(row_get(row, 'course_name') or '').strip()
    if title and course_name and title.endswith(f" - {course_name}"):
        return title[:-(len(course_name) + 3)].strip() or title
    return title or "CLO Attainment Report"

def normalize_report_title(value):
    return re.sub(r'\s+', ' ', str(value or '').strip())[:160]

def default_saved_report_title(course_name=''):
    return "CLO Attainment Report"

def report_title_exists(conn, user_id, course_name, title, exclude_report_id=None):
    query = """
        SELECT id FROM saved_reports
         WHERE user_id = ? AND course_name = ? AND lower(title) = lower(?)
    """
    params = [user_id, course_name, title]
    if exclude_report_id:
        query += " AND id <> ?"
        params.append(exclude_report_id)
    return conn.execute(query, params).fetchone() is not None

def unique_saved_report_title(conn, user_id, course_name, base_title):
    base_title = normalize_report_title(base_title) or "CLO Attainment Report"
    if not report_title_exists(conn, user_id, course_name, base_title):
        return base_title
    index = 1
    while True:
        suffix = f" ({index})"
        candidate = f"{base_title[:160 - len(suffix)]}{suffix}"
        if not report_title_exists(conn, user_id, course_name, candidate):
            return candidate
        index += 1

def student_rows_from_matrix(student_achievement_matrix):
    rows = []
    matrix = student_achievement_matrix or {}
    cells = matrix.get('cells') or {}
    for student_id in matrix.get('students') or []:
        for clo in matrix.get('clos') or []:
            cell = (cells.get(student_id) or {}).get(clo)
            if not cell:
                continue
            rows.append({
                'student_id': student_id,
                'clo': clo,
                'score': cell.get('score', 0),
                'target_score': cell.get('target_score', 0),
                'target_pct': cell.get('target_pct', 0),
                'achieved': cell.get('achieved', False),
                'status': cell.get('status', 'Achieved' if cell.get('achieved') else 'Not Achieved')
            })
    return rows

def build_clo_csv_response(stats, total_students, course_info, student_achievement_matrix, branding=None):
    branding = branding or get_report_branding()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["CLO Attainment Report"])
    writer.writerow(["University", branding.get('organization_name') or "N/A"])
    writer.writerow(["Course Name", course_info.get('course_name', '')])
    writer.writerow(["Course ID", course_info.get('course_id') or "N/A"])
    writer.writerow(["Report Date", datetime.now().strftime("%Y-%m-%d")])
    writer.writerow(["Total Students Evaluated", total_students])
    writer.writerow([])
    writer.writerow(["CLOs"])
    writer.writerow(["Domain", "Code", "CLOs"])
    for item in build_clo_definitions((stats or {}).keys()):
        writer.writerow([item['domain'], item['number'], item['wording']])

    writer.writerow([])
    writer.writerow(["Code", "Mapped Questions", "Max Possible Score", "Target Score", "Target %", "Students Achieved", "Achievement %"])
    for clo, data in sorted_clo_items(stats):
        writer.writerow([
            clo_number(clo),
            format_mapped_questions_for_report(data.get('questions', []), []),
            f"{float(data.get('total_possible_score', 0)):.2f}",
            f"{float(data.get('target_score', 0)):.2f}",
            f"{float(data.get('target_pct', 0)):.2f}",
            data.get('students_achieved', 0),
            f"{float(data.get('achievement_percentage', 0)):.2f}"
        ])

    writer.writerow([])
    writer.writerow(["Student CLO Achievement"])
    writer.writerow(["Student ID"] + [clo_number(clo) for clo in student_achievement_matrix.get('clos', [])])
    for student_id in student_achievement_matrix.get('students', []):
        row_values = [display_student_id(student_id)]
        for clo in student_achievement_matrix.get('clos', []):
            cell = (student_achievement_matrix.get('cells', {}).get(student_id) or {}).get(clo)
            if cell:
                row_values.append(f"{float(cell.get('score', 0)):.2f} - {cell.get('status', '')}")
            else:
                row_values.append("")
        writer.writerow(row_values)

    response = Response("\ufeff" + output.getvalue(), mimetype="text/csv; charset=utf-8")
    response.headers["Content-Disposition"] = 'attachment; filename="clo_achievement_report.csv"'
    return response

def report_creation_entitlement(user, saved_report_count):
    if not user:
        return 'anonymous'
    if subscription_active(user):
        if user['billing_plan'] == 'professional':
            return 'professional'
        if user['billing_plan'] == 'academic':
            academic_used = get_saved_report_count_since(user['id'], user['subscription_started_at'])
            if academic_used < ACADEMIC_REPORT_LIMIT_PER_YEAR:
                return 'academic'
    if saved_report_count < FREE_REPORT_LIMIT:
        return 'free'
    if int(user['report_credits'] or 0) > 0:
        return 'credit'
    return ''

def save_report_snapshot(stats, total_students, student_achievement_matrix, course_info):
    user = current_user()
    if not user:
        return {'allowed': True, 'saved': False, 'reason': 'anonymous'}

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    payload = {
        'stats': json_safe(stats),
        'total_students': total_students,
        'student_achievement_matrix': json_safe(student_achievement_matrix),
        'course_info': json_safe(course_info),
        'branding': get_display_branding_payload(),
        'created_at': created_at
    }
    hash_payload = dict(payload)
    hash_payload.pop('created_at', None)
    payload_json = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    report_hash = hashlib.sha256(json.dumps(hash_payload, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
    course_name = course_info.get('raw_name') or course_info.get('course_name') or ''
    with get_db() as conn:
        existing = conn.execute(
            "SELECT id FROM saved_reports WHERE user_id = ? AND report_hash = ?",
            (user['id'], report_hash)
        ).fetchone()
        if existing:
            return {'allowed': True, 'saved': False, 'reason': 'existing'}

        saved_count = conn.execute(
            "SELECT COUNT(*) FROM saved_reports WHERE user_id = ?",
            (user['id'],)
        ).fetchone()[0]
        entitlement = report_creation_entitlement(user, saved_count)
        if not entitlement:
            return {'allowed': False, 'saved': False, 'reason': 'billing_required'}

        title = unique_saved_report_title(
            conn,
            user['id'],
            course_name,
            default_saved_report_title(course_name)
        )
        conn.execute(
            """
            INSERT INTO saved_reports
                (user_id, title, course_name, payload_json, report_hash, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                user['id'],
                title,
                course_name,
                payload_json,
                report_hash,
                created_at
            )
        )
        if entitlement == 'credit':
            conn.execute(
                "UPDATE users SET report_credits = CASE WHEN report_credits > 0 THEN report_credits - 1 ELSE 0 END WHERE id = ?",
                (user['id'],)
            )
        return {'allowed': True, 'saved': True, 'reason': entitlement}

def load_courses():
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'courses_config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f).get('courses', [])
    return []

def build_course_display_name(course_name, course_code=''):
    course_name = re.sub(r'\s+', ' ', str(course_name or '').strip())
    course_code = re.sub(r'\s+', '', str(course_code or '').strip())
    if course_name and course_code and course_code not in course_name:
        return f"{course_name} ({course_code})"
    return course_name or course_code

def build_program_display_name(program_name, program_code=''):
    program_name = re.sub(r'\s+', ' ', str(program_name or '').strip())
    program_code = re.sub(r'\s+', '', str(program_code or '').strip())
    if program_name and program_code and program_code not in program_name:
        return f"{program_name} ({program_code})"
    return program_name or program_code

def build_course_select_label(course, language='ar'):
    course_name = re.sub(r'\s+', ' ', str((course or {}).get('course_name') or (course or {}).get('name') or '').strip())
    course_code = re.sub(r'\s+', '', str((course or {}).get('course_code') or '').strip())
    if not course_code:
        match = re.search(r'\(([A-Za-z]{1,8}\s*\d{2,6}[A-Za-z]?)\)', course_name)
        if match:
            course_code = re.sub(r'\s+', '', match.group(1)).upper()
            course_name = compact_text((course_name[:match.start()] + course_name[match.end():]).strip())
    if not course_code:
        match = re.search(r'\b([A-Za-z]{1,8}\s*\d{2,6}[A-Za-z]?)\b', course_name)
        if match:
            course_code = re.sub(r'\s+', '', match.group(1)).upper()
            course_name = compact_text((course_name[:match.start()] + course_name[match.end():]).strip())
    if course_code and course_name:
        has_arabic_name = bool(re.search(r'[\u0600-\u06FF]', course_name))
        if language == 'ar':
            return f"{course_code} - {course_name}" if has_arabic_name else f"{course_name} - {course_code}"
        return f"{course_name} - {course_code}" if has_arabic_name else f"{course_code} - {course_name}"
    return course_code or course_name

def set_course_display_labels(course):
    course['display_label_ar'] = build_course_select_label(course, 'ar')
    course['display_label_en'] = build_course_select_label(course, 'en')
    language = get_language() if has_request_context() else 'ar'
    course['display_label'] = course['display_label_ar'] if language == 'ar' else course['display_label_en']
    return course

def get_user_course_limit(user=None):
    user = user or current_user()
    if not user:
        return 0
    if subscription_active(user):
        if user['billing_plan'] == 'professional':
            return None
        if user['billing_plan'] == 'academic':
            return ACADEMIC_COURSE_LIMIT
    return 0

def get_user_course_count(user_id):
    with get_db() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM user_courses WHERE user_id = ?",
            (user_id,)
        ).fetchone()[0]

def row_to_course(row):
    try:
        clos = json.loads(row['clos_json'] or '[]')
    except json.JSONDecodeError:
        clos = []
    try:
        target_percentages = json.loads(row['target_percentages_json'] or '{}')
    except json.JSONDecodeError:
        target_percentages = {}
    try:
        topics = json.loads(row['topics_json'] or '[]')
    except json.JSONDecodeError:
        topics = []
    try:
        clo_plos = json.loads(row['clo_plos_json'] or '{}') if 'clo_plos_json' in row.keys() else {}
    except (json.JSONDecodeError, KeyError):
        clo_plos = {}
    if not isinstance(clo_plos, dict):
        clo_plos = {}
    course = {
        'id': row['id'],
        'name': row['display_name'],
        'course_name': row['course_name'],
        'course_code': row['course_code'] or '',
        'college': row['college'] or '',
        'program': row['program'] or '',
        'department': row['department'] or '',
        'clos': clos,
        'topics': topics,
        'clo_plos': clo_plos,
        'target_percentages': target_percentages,
        'grouped_clos': group_clos_by_domain(clos),
        'source': 'saved'
    }
    return set_course_display_labels(course)

def get_user_courses(user_id=None):
    user = current_user()
    if user_id is None:
        if not user:
            return []
        user_id = user['id']
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM user_courses WHERE user_id = ? ORDER BY display_name COLLATE NOCASE",
            (user_id,)
        ).fetchall()
    return [row_to_course(row) for row in rows]

def get_available_courses():
    courses = load_courses()
    for course in courses:
        course.setdefault('target_percentages', {})
        course.setdefault('topics', [])
        course.setdefault('clo_plos', {})
        course.setdefault('source', 'built_in')
        set_course_display_labels(course)
    user_courses = get_user_courses()
    for course in user_courses:
        set_course_display_labels(course)
    existing_names = {course.get('name') for course in courses}
    for course in user_courses:
        if course.get('name') and course.get('name') not in existing_names:
            courses.append(course)
            existing_names.add(course.get('name'))
    if has_request_context():
        for course in session.get('custom_courses', []) or []:
            if course.get('name') and course.get('name') not in existing_names:
                course.setdefault('target_percentages', {})
                course.setdefault('topics', [])
                course.setdefault('clo_plos', {})
                course.setdefault('source', 'session')
                set_course_display_labels(course)
                courses.append(course)
                existing_names.add(course.get('name'))
    return courses

def safe_available_courses():
    try:
        return get_available_courses()
    except Exception:
        app.logger.exception("Failed to load available courses")
        return []

def get_course_clos(course_name):
    try:
        courses = get_available_courses()
    except Exception:
        app.logger.exception("Failed to load course CLOs")
        courses = []
    return next((c.get('clos', []) for c in courses if c.get('name') == course_name), [])

def get_course_topics(course_name):
    try:
        courses = get_available_courses()
    except Exception:
        app.logger.exception("Failed to load course topics")
        courses = []
    return next((c.get('topics', []) for c in courses if c.get('name') == course_name), [])

def get_course_clo_plos(course_name):
    try:
        courses = get_available_courses()
    except Exception:
        app.logger.exception("Failed to load course CLO-PLO mappings")
        courses = []
    return next((c.get('clo_plos', {}) for c in courses if c.get('name') == course_name), {})

def get_course_by_name(course_name):
    course_name = str(course_name or '').strip()
    if not course_name:
        return {}
    try:
        courses = get_available_courses()
    except Exception:
        app.logger.exception("Failed to load course by name")
        courses = []
    for course in courses:
        names = {
            str(course.get('name') or '').strip(),
            str(course.get('course_name') or '').strip(),
            f"{str(course.get('course_name') or '').strip()} ({str(course.get('course_code') or '').strip()})".strip(),
        }
        if course_name in names:
            return course
    return {}

def enrich_course_info_from_course(course_info, selected_course_name=''):
    info = dict(course_info or {})
    candidate_names = [
        selected_course_name,
        info.get('raw_name'),
        info.get('course_name'),
        info.get('name'),
    ]
    course = {}
    for candidate in candidate_names:
        course = get_course_by_name(candidate)
        if course:
            break
    if not course:
        return info

    course_name = course.get('course_name') or course.get('name') or ''
    course_code = course.get('course_code') or ''
    if course_name:
        info['course_name'] = course_name
    if course_code:
        info['course_id'] = course_code
    info['raw_name'] = course.get('name') or info.get('raw_name') or selected_course_name or course_name
    for field in ('college', 'department', 'program'):
        value = course.get(field) or ''
        if value:
            info[field] = value
    if course.get('clo_plos'):
        info['clo_plos'] = course.get('clo_plos') or {}
    return info

def group_clos_by_domain(clos):
    grouped = {
        'knowledge': [],
        'skills': [],
        'values': [],
        'other': []
    }
    for clo in clos or []:
        clo_text = str(clo).strip()
        if clo_text.startswith('1.'):
            grouped['knowledge'].append(clo_text)
        elif clo_text.startswith('2.'):
            grouped['skills'].append(clo_text)
        elif clo_text.startswith('3.'):
            grouped['values'].append(clo_text)
        elif clo_text:
            grouped['other'].append(clo_text)
    return grouped

def parse_pasted_clos(value):
    clos = []
    for line in str(value or '').splitlines():
        cleaned = re.sub(r'\s+', ' ', line).strip()
        if not cleaned:
            continue
        if re.match(r'^(?:\d+(?:\.\d+)+|CLO\s*\d+(?:\.\d+)*)\b', cleaned, flags=re.I):
            clos.append(cleaned)
    return clos

def parse_course_topic_lines(value):
    topics = []
    for line in str(value or '').splitlines():
        topic = re.sub(r'^\s*[-\u2022\u2023\u25e6\u2043\u2219]\s*', '', line)
        topic = compact_text(topic)
        if topic:
            topics.append(topic)
    return topics

def format_course_topics_text(topics):
    return "\n".join(f"- {compact_text(topic)}" for topic in topics or [] if compact_text(topic))

def clo_plo_from_mapping(clo, clo_plos):
    if not isinstance(clo_plos, dict):
        return ''
    normalized_lookup = {
        re.sub(r'\s+', '', str(key or '').strip()).upper(): value
        for key, value in clo_plos.items()
    }
    for candidate in (clo_number_key(clo), clo_number(clo), str(clo or '').strip()):
        lookup_key = re.sub(r'\s+', '', str(candidate or '').strip()).upper()
        if lookup_key in normalized_lookup:
            return compact_text(normalized_lookup[lookup_key])
    return ''

def build_course_clo_rows(clos, clo_plos=None):
    rows = []
    for clo in clos or []:
        clo_text = compact_text(clo)
        if not clo_text:
            continue
        rows.append({
            'clo': clo_text,
            'plo': clo_plo_from_mapping(clo_text, clo_plos or {})
        })
    return rows

def parse_course_clo_rows(form):
    clos = []
    clo_plos = {}
    for index in form.getlist('clo_row_indexes'):
        index = str(index)
        clo_text = compact_text(form.get(f'clo_text_{index}') or '')
        if not clo_text:
            continue
        if not re.match(r'^(?:\d+(?:\.\d+)+|CLO\s*\d+(?:\.\d+)*)\b', clo_text, flags=re.I):
            continue
        clos.append(clo_text)
        plo_text = compact_text(form.get(f'clo_plo_{index}') or '')
        if plo_text:
            codes = extract_plo_codes(plo_text)
            key = clo_number_key(clo_text) or clo_number(clo_text)
            clo_plos[key] = ', '.join(codes) if codes else plo_text
    return clos, clo_plos

def normalize_plo_code(value):
    digit_map = globals().get('ARABIC_DIGITS') or str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
    raw_code = normalize_course_spec_text(value).translate(digit_map)
    raw_code = raw_code.replace('ـ', '')
    code = re.sub(r'[\s:_\-]+', '', raw_code.upper())
    code = re.sub(r'^PLO(?=[KSV])', '', code)
    reversed_arabic = re.match(r'^(\d+)([عمق])$', code)
    if reversed_arabic:
        code = reversed_arabic.group(2) + reversed_arabic.group(1)
    code = re.sub(r'^PLO(?=[عمقك])', '', code)
    code = code.replace('ك', 'ع')
    arabic_code_match = re.match(r'^([عمق])(\d{2,})$', code)
    if arabic_code_match and arabic_code_match.group(1) in {'م', 'ق'} and arabic_code_match.group(2).startswith('5'):
        code = arabic_code_match.group(1) + arabic_code_match.group(2)[1:]
    return code

def extract_plo_codes(value):
    text = normalize_course_spec_text(value)
    codes = []
    pattern = r'\b(?:PLO\s*[-:_]?\s*)?[KSV]\s*\d+(?:\.\d+)?\b|\bPLO\s*[-:_]?\s*\d+(?:\.\d+)?\b'
    for match in re.finditer(pattern, text, flags=re.I):
        code = normalize_plo_code(match.group(0))
        if code and code not in codes:
            codes.append(code)
    arabic_pattern = (
        r'(?<![\u0600-\u06FFA-Za-z0-9])'
        r'(?:PLO\s*[-:_]?\s*)?'
        r'([عمقك]\s*[\d٠-٩۰-۹]{1,2}(?:\.[\d٠-٩۰-۹]+)?)'
        r'(?![\u0600-\u06FFA-Za-z0-9])'
    )
    for match in re.finditer(arabic_pattern, text, flags=re.I):
        code = normalize_plo_code(match.group(1))
        if code and code not in codes:
            codes.append(code)
    reversed_arabic_pattern = r'(?<![\u0600-\u06FFA-Za-z0-9])([\dÙ -Ù©Û°-Û¹]{1,2}\s*[عمق])(?![\u0600-\u06FFA-Za-z0-9])'
    for match in re.finditer(reversed_arabic_pattern, text, flags=re.I):
        code = normalize_plo_code(match.group(1))
        if code and code not in codes:
            codes.append(code)
    return codes

def parse_clo_plos_json(value):
    try:
        data = json.loads(value or '{}')
    except (TypeError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    cleaned = {}
    for key, val in data.items():
        key = re.sub(r'\s+', '', str(key or '').strip())
        if isinstance(val, (list, tuple)):
            codes = [normalize_plo_code(item) for item in val]
            val = ', '.join(code for code in codes if code)
        else:
            codes = extract_plo_codes(val)
            val = ', '.join(codes) if codes else compact_text(val)
        if key and val:
            cleaned[key] = val
    return cleaned

def arabic_plo_letter_for_clo(clo):
    number = clo_number(clo) or ''
    domain = number.split('.', 1)[0] if '.' in number else ''
    if domain == '1':
        return 'ع'
    if domain == '2':
        return 'م'
    if domain == '3':
        return 'ق'
    body = re.sub(r'^\s*(?:[123]\.\d+|CLO\s*\d+(?:\.\d+)*)\s+', '', str(clo or ''), flags=re.I)
    domain = arabic_starter_domain(body)
    return {'1': 'ع', '2': 'م', '3': 'ق'}.get(domain, '')

def extract_noisy_arabic_plo_code(line, clo):
    codes = extract_plo_codes(line)
    if codes:
        return codes[0]

    letter = arabic_plo_letter_for_clo(clo)
    if not letter:
        return ''

    digit_map = globals().get('ARABIC_DIGITS') or str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
    normalized = normalize_course_spec_text(line).translate(digit_map)
    normalized = re.sub(r'^[^\u0600-\u06FFA-Za-z0-9%#@¢£€]+', '', normalized)
    if re.search(r'\b(?:المحاضرة|المناقشة|الواجبات|الاختبار|البحوث|التقارير|استبانة|تقويم)\b', normalized):
        normalized = normalized.split(re.search(r'\b(?:المحاضرة|المناقشة|الواجبات|الاختبار|البحوث|التقارير|استبانة|تقويم)\b', normalized).group(0), 1)[0]

    starter = re.search(r'(?:يبين|يوضح|يشرح|يصف|يعرف|يذكر|يستعرض|تعرض|يؤصل|يطبق|يحلل|يقارن|يناقش|يجيد|يستخدم|يوظف|يستنبط|يلتزم|يقدر|يتحلى|يتعاون|يبادر|يتحمل|يشارك)', normalized)
    prefix = normalized[:starter.start()] if starter else normalized[:45]
    prefix = compact_text(prefix)
    if not prefix:
        return ''

    if letter == 'ع':
        match = re.search(r'^\s*([1-9])\s*[-–—]', prefix) or re.search(r'(?<!\d)([1-9])\s*[-–—]?\s*$', prefix)
        if match:
            return f'{letter}{match.group(1)}'
    if letter == 'م':
        match = re.search(r'^\s*(?:[مmM]|5)\s*([1-9])(?=\s|[^\d]|$)', prefix) or re.search(r'(?:[مmM]|5)\s*([1-9])\s*$', prefix)
        if match:
            return f'{letter}{match.group(1)}'
    if letter == 'ق':
        match = re.search(r'^\s*(?:[قqQgG%#¢]|2)\s*([1-9])(?=\s|[^\d]|$)', prefix) or re.search(r'(?:[قqQgG%#¢]|2)\s*([1-9])\s*$', prefix)
        if match:
            return f'{letter}{match.group(1)}'
    return ''

def extract_arabic_clo_plo_mapping_by_order(text, clos):
    if not clos or not contains_arabic(text):
        return {}
    lines = [compact_text(line) for line in normalize_course_spec_text(text).splitlines()]
    lines = [line for line in lines if line]
    start_indexes = [
        index for index, line in enumerate(lines)
        if (
            line_has_arabic_label(line, 'ناتج التعلم المرتبط بالبرنامج')
            or (
                line_has_arabic_label(line, 'الرمز')
                and any(line_has_arabic_label(next_line, 'ناتج التعلم') for next_line in lines[index:index + 5])
            )
        )
    ]
    if not start_indexes:
        return {}

    starter_pattern = re.compile(
        r'(يبين|يوضح|يشرح|يصف|يعرف|يذكر|يستعرض|تعرض|يؤصل|يطبق|يحلل|يقارن|يناقش|يجيد|يستخدم|يوظف|يستنبط|يلتزم|يقدر|يتحلى|يتعاون|يبادر|يتحمل|يشارك)'
    )

    sections = []
    for start_index in start_indexes:
        section_lines = []
        for line in lines[start_index + 1:start_index + 70]:
            if (
                line_has_arabic_label(line, 'موضوعات المقرر')
                or line_has_arabic_label(line, 'مصادر التعلم')
                or re.match(r'^\s*(?:نوازل|توازل|توائل)\b', line)
            ):
                break
            section_lines.append(line)
        starter_count = sum(
            1 for line in section_lines
            if starter_pattern.search(clean_arabic_outcome_line(line))
        )
        sections.append((starter_count, 1 if start_index >= 50 else 0, section_lines))
    if not sections:
        return {}
    section_lines = max(sections, key=lambda item: (item[0], item[1], len(item[2])))[2]
    starter_indexes = []
    for index, line in enumerate(section_lines):
        if starter_pattern.search(clean_arabic_outcome_line(line)):
            starter_indexes.append(index)

    outcome_windows = []
    for position, start_index in enumerate(starter_indexes):
        end_index = starter_indexes[position + 1] if position + 1 < len(starter_indexes) else min(len(section_lines), start_index + 5)
        segment_lines = section_lines[start_index:end_index]
        outcome_windows.append(segment_lines)

    mapping = {}
    for clo, segment_lines in zip(clos, outcome_windows):
        key = clo_number_key(clo) or clo_number(clo)
        code = ''
        for line in segment_lines:
            code = extract_noisy_arabic_plo_code(line, clo)
            if code:
                break
        if not code:
            code = extract_noisy_arabic_plo_code(' '.join(segment_lines), clo)
        if key and code:
            mapping[key] = code
    return mapping

def extract_clo_plo_mapping(text, clos=None):
    source = normalize_course_spec_text(text)
    if not source:
        return {}
    lines = [compact_text(line) for line in source.splitlines() if compact_text(line)]
    if not lines:
        lines = [compact_text(source)]
    mapping = {}
    clo_pattern = re.compile(r'\b((?:CLO\s*)?\d+\s*[\.\-]\s*[1-9]\d*|CLO\s*\d+)\b', flags=re.I)

    for index, line in enumerate(lines):
        for match in clo_pattern.finditer(line):
            clo_id = re.sub(r'\s+', '', match.group(1).upper()).replace('-', '.')
            if re.match(r'^[123]\.0$', clo_id):
                continue
            tail = line[match.end():]
            next_clo = clo_pattern.search(tail)
            if next_clo:
                tail = tail[:next_clo.start()]
            codes = extract_plo_codes(tail)
            if not codes:
                lookahead = []
                for next_line in lines[index + 1:index + 3]:
                    if clo_pattern.search(next_line):
                        break
                    lookahead.append(next_line)
                codes = extract_plo_codes(' '.join(lookahead))
            if codes:
                mapping[clo_id] = ', '.join(codes)

    if clos:
        normalized_mapping = {}
        for clo in clos:
            key = clo_number_key(clo) or clo_number(clo)
            if key and key in mapping:
                normalized_mapping[key] = mapping[key]
        if contains_arabic(source):
            arabic_mapping = extract_arabic_clo_plo_mapping_by_order(source, clos)
            normalized_mapping.update({key: value for key, value in arabic_mapping.items() if key and value})
        if normalized_mapping:
            return normalized_mapping
    if clos and contains_arabic(source):
        arabic_mapping = extract_arabic_clo_plo_mapping_by_order(source, clos)
        if arabic_mapping:
            return arabic_mapping
    return mapping

PLO_MATRIX_DEFAULT_CODES = ['K1', 'K2', 'K3', 'S1', 'S2', 'S3', 'S4', 'V1', 'V2', 'V3']
PLO_MATRIX_MARKS = {
    'I': 'I',
    'INTRODUCED': 'I',
    'P': 'P',
    'PRACTICED': 'P',
    'PRACTISED': 'P',
    'M': 'M',
    'MASTERED': 'M',
    'X': 'X',
    '✓': 'X',
    '✔': 'X',
    '☑': 'X',
    '1': 'I',
    '2': 'P',
    '3': 'M',
    'س': 'I',
    'ر': 'P',
    'ت': 'M',
}

def clean_matrix_cell(value):
    if value is None:
        return ''
    return compact_text(str(value).replace('\n', ' '))

def normalize_course_title_for_match(value):
    value = clean_matrix_cell(value).lower()
    value = re.sub(r'\([^)]*\)', ' ', value)
    value = re.sub(r'[^\u0600-\u06FFa-z0-9]+', ' ', value)
    value = re.sub(r'\b(?:the|of|and|for|in|to)\b', ' ', value)
    return compact_text(value)

def normalize_plo_matrix_mark(value):
    text = clean_matrix_cell(value).translate(ARABIC_DIGITS).upper()
    if not text:
        return ''
    text = text.strip(' .,:;|/\\')
    return PLO_MATRIX_MARKS.get(text, '')

def extract_program_plo_definitions(text):
    source = normalize_course_spec_text(text or '')
    lines = [compact_text(line) for line in source.splitlines()]
    definitions = {}
    current_code = ''
    current_parts = []
    for line in lines:
        if not line:
            continue
        if re.search(r'\b(?:C\.|Curriculum|Program Courses|Course code|Program learning Outcomes Mapping Matrix)\b|المنهج|مقررات|مصفوفة', line, flags=re.I):
            if current_code and current_parts:
                definitions[current_code] = compact_text(' '.join(current_parts))
            if definitions:
                break
        match = re.match(r'^\s*((?:PLO\s*)?[KSV]\s*\d+(?:\.\d+)?)\s+(.+)$', line, flags=re.I)
        if not match:
            match = re.match(r'^\s*([عمق]\s*[\dÙ -Ù©Û°-Û¹]{1,2}|[\dÙ -Ù©Û°-Û¹]{1,2}\s*[عمق])\s+(.+)$', line, flags=re.I)
        if match:
            if current_code and current_parts:
                definitions[current_code] = compact_text(' '.join(current_parts))
            current_code = normalize_plo_code(match.group(1))
            current_parts = [match.group(2)]
        elif current_code and not re.match(r'^(?:Knowledge|Understanding|Skills|Values|Autonomy|Responsibility)\b', line, flags=re.I):
            current_parts.append(line)
    if current_code and current_parts:
        definitions[current_code] = compact_text(' '.join(current_parts))
    return definitions

def filter_matrix_to_defined_plos(courses, plo_codes, definitions):
    defined_codes = {normalize_plo_code(code) for code in (definitions or {}) if normalize_plo_code(code)}
    if len(defined_codes) < 6:
        return courses, plo_codes
    filtered_courses = []
    used_codes = []
    for course in courses or []:
        filtered_plos = [
            item for item in (course.get('plos') or [])
            if normalize_plo_code(item.get('code')) in defined_codes
        ]
        if not filtered_plos:
            continue
        updated = dict(course)
        updated['plos'] = filtered_plos
        updated['plo_codes'] = [
            normalize_plo_code(item.get('code'))
            for item in filtered_plos
            if normalize_plo_code(item.get('code'))
        ]
        for code in updated['plo_codes']:
            if code and code not in used_codes:
                used_codes.append(code)
        filtered_courses.append(updated)
    filtered_codes = [normalize_plo_code(code) for code in (plo_codes or []) if normalize_plo_code(code) in defined_codes]
    for code in used_codes:
        if code not in filtered_codes:
            filtered_codes.append(code)
    return filtered_courses or courses, filtered_codes or plo_codes

def build_course_catalog_from_text(text):
    source = re.sub(r'\s+', ' ', normalize_course_spec_text(text or ' '))
    catalog = []
    pattern = re.compile(
        r'\b([A-Z]{2,6}\d{3,5})\s+(.{3,100}?)\s+(?:Required|Elective)\b',
        flags=re.I
    )
    for match in pattern.finditer(source):
        code = match.group(1).upper()
        title = clean_matrix_cell(match.group(2))
        title = re.sub(r'\b(?:Level|Required|Elective)\b.*$', '', title, flags=re.I).strip()
        if len(title) >= 3 and not any(item['course_code'] == code for item in catalog):
            catalog.append({
                'course_code': code,
                'course_title': title,
                'match_key': normalize_course_title_for_match(title)
            })
    arabic_pattern = re.compile(
        r'\b([A-Z]{2,6}\d{3,5})\s+(.{3,120}?)(?=\s+(?:ير\s*ابجإ|ي\s*رابجإ|إجباري|اختياري)\b)',
        flags=re.I
    )
    for match in arabic_pattern.finditer(source):
        code = match.group(1).upper()
        title = clean_matrix_cell(match.group(2))
        if len(title) >= 3 and not any(item['course_code'] == code for item in catalog):
            catalog.append({
                'course_code': code,
                'course_title': title,
                'match_key': normalize_course_title_for_match(title)
            })
    return catalog

def lookup_course_code_from_catalog(course_title, catalog):
    key = normalize_course_title_for_match(course_title)
    if not key or not catalog:
        return ''
    for item in catalog:
        if key == item.get('match_key'):
            return item.get('course_code', '')
    for item in catalog:
        item_key = item.get('match_key') or ''
        if key and item_key and (key in item_key or item_key in key):
            return item.get('course_code', '')
    choices = [item.get('match_key', '') for item in catalog]
    matches = difflib.get_close_matches(key, choices, n=1, cutoff=0.76)
    if matches:
        return next((item.get('course_code', '') for item in catalog if item.get('match_key') == matches[0]), '')
    return ''

def table_contains_plo_matrix_header(table_rows):
    text = ' '.join(clean_matrix_cell(cell) for row in table_rows[:8] for cell in row)
    return bool(
        (
            re.search(r'Program\s+Learning\s+Outcomes|Mapping\s+Matrix|K\s*1|K1', text, flags=re.I)
            and re.search(r'\bS\s*1\b|\bS1\b|\bV\s*1\b|\bV1\b', text, flags=re.I)
        )
        or (
            re.search(r'نواتج|ملعتلا|مصفوفة|جمان', text)
            and re.search(r'[ععممقق]\s*[\dÙ -Ù©Û°-Û¹]|[\dÙ -Ù©Û°-Û¹]\s*[ععممقق]', text)
        )
    )

def matrix_columns_for_table(table_rows, continuation=False):
    if not table_rows:
        return []
    ncols = max((len(row) for row in table_rows), default=0)
    header_text = ' '.join(clean_matrix_cell(cell) for row in table_rows[:6] for cell in row)
    if ncols >= 20 and table_contains_plo_matrix_header(table_rows) and not contains_arabic(header_text):
        return [
            (1, 'K1'), (4, 'K2'), (5, 'K3'),
            (8, 'S1'), (10, 'S2'), (12, 'S3'), (14, 'S4'),
            (16, 'V1'), (18, 'V2'), (20, 'V3')
        ]
    if continuation and 12 <= ncols <= 14:
        return [
            (1, 'K1'), (2, 'K2'), (3, 'K3'),
            (5, 'S1'), (6, 'S2'), (7, 'S3'), (8, 'S4'),
            (10, 'V1'), (11, 'V2'), (12, 'V3')
        ]

    discovered = []
    for row in table_rows[:10]:
        for col_index, cell in enumerate(row):
            cell_text = clean_matrix_cell(cell)
            codes = extract_plo_codes(cell_text)
            for code in codes:
                if re.match(r'^(?:[KSV]\d+|[عمق]\d+)', code, flags=re.I) and (col_index, code) not in discovered:
                    discovered.append((col_index, code))
    discovered = sorted(discovered, key=lambda item: item[0])
    if ncols > 24 and re.search(r'نواتج|ملعتلا|تاررق|المقررات|جمان', header_text) and len(discovered) >= 4:
        return discovered
    if 18 <= ncols <= 24 and re.search(r'نواتج|ملعتلا|تاررق|المقررات|جمان', header_text) and re.search(r'[ععممقق]\s*[\dÙ -Ù©Û°-Û¹]|[\dÙ -Ù©Û°-Û¹]\s*[ععممقق]', header_text):
        return [
            (0, 'ق5'), (1, 'ق4'), (2, 'ق3'), (3, 'ق2'), (4, 'ق1'),
            (5, 'م5'), (6, 'م4'), (7, 'م3'), (8, 'م2'), (9, 'م1'),
            (10, 'ع5'), (11, 'ع4'), (12, 'ع3'), (13, 'ع2'), (14, 'ع1')
        ]
    if len(discovered) >= 2:
        return discovered
    return []

def row_looks_like_matrix_header(row):
    text = ' '.join(clean_matrix_cell(cell) for cell in row)
    return bool(re.search(r'Course\s+code|Program\s+Learning\s+Outcomes|Knowledge|Skills|Responsibility|---|K\s*2|K2|نواتج|ملعتلا|تاررق|المقررات|مقررات|المعرفة|ةفرع|المهارات|تارا|القيم|ميقلا|المستوى|ىوتسم|ىوتسلما|لول ا|لولأا|يناثلا|ثلاثلا|عبارلا|سماخلا|سداسلا|عباسلا|نماثلا', text, flags=re.I))

def is_matrix_level_title(value):
    text = clean_matrix_cell(value)
    if not text:
        return False
    return bool(re.match(r'^(?:المستوى|ىوتسم|ىوتسلما|لول ا|لولأا|يناثلا|ثلاثلا|عبارلا|سماخلا|سداسلا|عباسلا|نماثلا)(?:\s*\d+|\s*[Ù -Ù©Û°-Û¹]+)?$', text, flags=re.I))

def matrix_course_code_from_row(row, columns):
    if not row or not columns:
        return ''
    max_col = max(col for col, _ in columns)
    searchable = ' '.join(clean_matrix_cell(cell) for cell in row[max_col + 1:])
    match = re.search(r'\b[A-Z]{2,8}\s*-?\s*\d{3,5}\b', searchable, flags=re.I)
    if match:
        return re.sub(r'\s+', '', match.group(0)).upper()
    return ''

def matrix_course_title_from_row(row, columns):
    if not row or not columns:
        return ''
    min_col = min(col for col, _ in columns)
    max_col = max(col for col, _ in columns)
    left = [clean_matrix_cell(cell) for cell in row[:min_col]]
    right = [clean_matrix_cell(cell) for cell in row[max_col + 1:]]
    code_pattern = re.compile(r'^[A-Z]{2,8}\s*-?\s*\d{3,5}$', flags=re.I)
    skip_pattern = re.compile(r'^(?:Course\s+code|اسم\s+المقرر|رمز\s+المقرر|ررقلما|ررقلام|المستوى|ىوتسم|ىوتسلما|لول ا|لولأا|يناثلا|ثلاثلا|عبارلا|سماخلا|سداسلا|عباسلا|نماثلا)$', flags=re.I)
    left = [cell for cell in left if cell and not normalize_plo_matrix_mark(cell) and not code_pattern.match(cell) and not skip_pattern.match(cell)]
    right = [cell for cell in right if cell and not normalize_plo_matrix_mark(cell) and not code_pattern.match(cell) and not skip_pattern.match(cell)]
    left_text = compact_text(' '.join(left))
    right_text = compact_text(' '.join(right))
    if contains_arabic(right_text) and len(right_text) >= len(left_text):
        return right_text
    if left_text:
        return left_text
    return right_text

def matrix_mark_for_column(row, col_index, used_indexes=None, allow_nearby=True):
    used_indexes = used_indexes or set()
    offsets = [0, -1, 1] if allow_nearby else [0]
    for offset in offsets:
        index = col_index + offset
        if index < 0 or index >= len(row) or index in used_indexes:
            continue
        mark = normalize_plo_matrix_mark(row[index])
        if mark:
            used_indexes.add(index)
            return mark
    return ''

def parse_plo_matrix_table(table_rows, catalog=None, continuation=False):
    if not table_rows:
        return [], []
    columns = matrix_columns_for_table(table_rows, continuation=continuation)
    if not columns:
        return [], []
    ncols = max((len(row) for row in table_rows), default=0)
    allow_nearby = ncols >= 20
    courses = []
    for row in table_rows:
        cells = list(row)
        if row_looks_like_matrix_header(cells):
            continue
        used_indexes = set()
        plos = []
        for col_index, code in columns:
            mark = matrix_mark_for_column(cells, col_index, used_indexes, allow_nearby=allow_nearby)
            if mark:
                plos.append({'code': normalize_plo_code(code), 'level': mark})
        course_title = matrix_course_title_from_row(cells, columns)
        if not course_title or len(course_title) < 2:
            continue
        if is_matrix_level_title(course_title) or re.match(r'^\d+$', course_title) or re.search(r'\b(?:Academic Rank|Evaluation|KPIs|No\.)\b', course_title, flags=re.I):
            continue
        if plos:
            courses.append({
                'course_title': course_title,
                'course_code': matrix_course_code_from_row(cells, columns) or lookup_course_code_from_catalog(course_title, catalog or []),
                'plos': plos,
                'plo_codes': [item['code'] for item in plos]
            })
        elif courses and contains_arabic(course_title):
            previous = courses[-1]
            combined = compact_text(f"{previous.get('course_title', '')} {course_title}")
            previous['course_title'] = combined
            if not previous.get('course_code'):
                previous['course_code'] = lookup_course_code_from_catalog(combined, catalog or [])
    return courses, [code for _, code in columns]

def docx_tables(filepath):
    tables = []
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    try:
        with zipfile.ZipFile(filepath) as docx:
            xml = docx.read('word/document.xml')
    except Exception:
        return tables
    root = ET.fromstring(xml)
    for tbl in root.findall('.//w:tbl', ns):
        rows = []
        for tr in tbl.findall('./w:tr', ns):
            row = []
            for tc in tr.findall('./w:tc', ns):
                texts = [node.text or '' for node in tc.findall('.//w:t', ns)]
                row.append(clean_matrix_cell(' '.join(texts)))
            if any(row):
                rows.append(row)
        if rows:
            tables.append(rows)
    return tables

def extract_plo_matrix_from_tables(tables, catalog=None):
    all_courses = []
    plo_codes = []
    active_matrix = False
    empty_after_active = 0
    for table in tables:
        is_matrix = table_contains_plo_matrix_header(table)
        courses, codes = parse_plo_matrix_table(table, catalog=catalog, continuation=active_matrix and not is_matrix)
        if courses:
            active_matrix = True
            empty_after_active = 0
            all_courses.extend(courses)
            for code in codes:
                code = normalize_plo_code(code)
                if code and code not in plo_codes:
                    plo_codes.append(code)
        elif is_matrix:
            active_matrix = True
            empty_after_active = 0
        elif active_matrix:
            empty_after_active += 1
            if empty_after_active >= 2:
                active_matrix = False
    cleaned = []
    seen = set()
    for course in all_courses:
        key = normalize_course_title_for_match(course.get('course_title'))
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(course)
    return cleaned, plo_codes or PLO_MATRIX_DEFAULT_CODES

def extract_plo_matrix_from_pdf(filepath, text):
    tables = []
    try:
        import fitz
        doc = fitz.open(filepath)
        for page in doc:
            if not hasattr(page, 'find_tables'):
                continue
            for table in page.find_tables().tables:
                tables.append(table.extract())
    except Exception:
        tables = []
    catalog = build_course_catalog_from_text(text)
    return extract_plo_matrix_from_tables(tables, catalog=catalog)

def extract_plo_matrix_from_excel(filepath):
    tables = []
    try:
        sheets = pd.read_excel(filepath, sheet_name=None, header=None)
    except Exception:
        return [], PLO_MATRIX_DEFAULT_CODES
    for df in sheets.values():
        table = []
        for _, row in df.iterrows():
            table.append([clean_matrix_cell(value) for value in row.tolist()])
        tables.append(table)
    return extract_plo_matrix_from_tables(tables, catalog=[])

def extract_plo_matrix_from_docx(filepath, text):
    catalog = build_course_catalog_from_text(text)
    return extract_plo_matrix_from_tables(docx_tables(filepath), catalog=catalog)

def extract_plo_matrix_from_text(text):
    source = normalize_course_spec_text(text)
    start_match = re.search(r'Program\s+learning\s+Outcomes\s+Mapping\s+Matrix|Mapping\s+Matrix', source, flags=re.I)
    if not start_match:
        return [], PLO_MATRIX_DEFAULT_CODES
    section = source[start_match.start():]
    stop_match = re.search(r'\n\s*(?:5\.|D\.|Student Admission|Teaching and learning strategies|Assessment Plan)\b', section, flags=re.I)
    if stop_match:
        section = section[:stop_match.start()]
    lines = [line.rstrip() for line in section.splitlines() if compact_text(line)]
    courses = []
    title_parts = []
    mark_line_pattern = re.compile(r'^\s*(?:[IPMX✓✔☑سرت]\s*){1,18}$', flags=re.I)
    inline_pattern = re.compile(r'^(?P<title>.+?)\s+(?P<marks>(?:[IPMXسرت]\s*){1,18})$', flags=re.I)
    for line in lines:
        clean_line = compact_text(line)
        if row_looks_like_matrix_header([clean_line]) or re.match(r'^\d+$', clean_line):
            continue
        marks_text = ''
        title_text = ''
        if mark_line_pattern.match(line):
            marks_text = clean_line
            title_text = compact_text(' '.join(title_parts))
            title_parts = []
        else:
            inline = inline_pattern.match(clean_line)
            if inline:
                title_text = compact_text(' '.join(title_parts + [inline.group('title')]))
                marks_text = inline.group('marks')
                title_parts = []
            else:
                title_parts.append(clean_line)
                continue
        marks = [normalize_plo_matrix_mark(item) for item in re.findall(r'[IPMX✓✔☑سرت]', marks_text, flags=re.I)]
        marks = [mark for mark in marks if mark]
        if title_text and marks:
            plos = [
                {'code': code, 'level': mark}
                for code, mark in zip(PLO_MATRIX_DEFAULT_CODES, marks)
            ]
            courses.append({
                'course_title': title_text,
                'course_code': '',
                'plos': plos,
                'plo_codes': [item['code'] for item in plos]
            })
    return courses, PLO_MATRIX_DEFAULT_CODES

def extract_program_plo_matrix(filepath, file_ext):
    file_ext = (file_ext or '').lower()
    text = ''
    if file_ext in {'.pdf', '.docx', '.txt'}:
        text = extract_document_text(filepath, allow_ocr=True) if file_ext != '.txt' else open(filepath, 'r', encoding='utf-8', errors='ignore').read()
    if file_ext == '.pdf':
        courses, plo_codes = extract_plo_matrix_from_pdf(filepath, text)
    elif file_ext == '.docx':
        courses, plo_codes = extract_plo_matrix_from_docx(filepath, text)
    elif file_ext in {'.xlsx', '.xls'}:
        courses, plo_codes = extract_plo_matrix_from_excel(filepath)
    else:
        courses, plo_codes = extract_plo_matrix_from_text(text)
    if not courses and text:
        courses, plo_codes = extract_plo_matrix_from_text(text)
    definitions = extract_program_plo_definitions(text) if text else {}
    courses, plo_codes = filter_matrix_to_defined_plos(courses, plo_codes, definitions)
    return {
        'filename': os.path.basename(filepath),
        'plo_codes': plo_codes,
        'plo_definitions': definitions,
        'courses': courses,
        'course_count': len(courses),
        'confidence': 'High' if courses and file_ext in {'.pdf', '.docx', '.xlsx', '.xls'} else ('Medium' if courses else 'Low')
    }

def format_matrix_course_plos(course):
    parts = []
    for item in course.get('plos') or []:
        code = item.get('code', '')
        level = item.get('level', '')
        if code:
            parts.append(f"{code} ({level})" if level else code)
    return ', '.join(parts)

def find_matrix_course_for_report(course_name, matrix):
    courses = (matrix or {}).get('courses') or []
    key = normalize_course_title_for_match(course_name)
    if not key:
        return {}
    for course in courses:
        candidate = normalize_course_title_for_match(course.get('course_title'))
        if key == candidate or key in candidate or candidate in key:
            return course
    choices = [normalize_course_title_for_match(course.get('course_title')) for course in courses]
    matches = difflib.get_close_matches(key, choices, n=1, cutoff=0.72)
    if matches:
        return next((course for course in courses if normalize_course_title_for_match(course.get('course_title')) == matches[0]), {})
    return {}

def decode_pdf_string(value, cmap=None):
    value = value.replace(r'\\', '\u0000')
    value = value.replace(r'\(', '(').replace(r'\)', ')')
    value = value.replace(r'\n', ' ').replace(r'\r', ' ').replace(r'\t', ' ')

    def replace_octal(match):
        try:
            return chr(int(match.group(1), 8))
        except ValueError:
            return ''

    value = re.sub(r'\\([0-7]{1,3})', replace_octal, value)
    value = value.replace('\u0000', '\\')
    if cmap:
        return ''.join(cmap.get(ord(char), char) for char in value)
    return value

def decode_pdf_hex_string(value, cmap=None):
    try:
        raw = bytes.fromhex(value)
    except ValueError:
        return ''

    if cmap:
        return ''.join(cmap.get(byte, chr(byte)) for byte in raw)
    if raw.startswith(b'\xfe\xff'):
        return raw[2:].decode('utf-16-be', errors='ignore')
    return raw.decode('latin-1', errors='ignore')

def decode_pdf_text_array(array_content, cmap=None):
    tokens = re.finditer(
        r'\((?:\\.|[^\\()])*\)|<([0-9A-Fa-f\s]+)>|[-+]?\d*\.?\d+',
        array_content,
        flags=re.S
    )
    text = ''
    for token_match in tokens:
        token = token_match.group(0)
        if token.startswith('('):
            text += decode_pdf_string(token[1:-1], cmap)
        elif token.startswith('<'):
            text += decode_pdf_hex_string(token[1:-1], cmap)
        else:
            try:
                spacing_adjustment = abs(float(token))
            except ValueError:
                continue
            if spacing_adjustment > 120 and text and not text.endswith(' '):
                text += ' '
    return text

def extract_pdf_block_text(block, cmap=None):
    text = ''
    for text_match in re.finditer(
        r'\[(.*?)\]\s*TJ|\((.*?)\)\s*Tj|<([0-9A-Fa-f\s]+)>\s*Tj',
        block,
        flags=re.S
    ):
        if text_match.group(1) is not None:
            text += decode_pdf_text_array(text_match.group(1), cmap)
        elif text_match.group(2) is not None:
            text += decode_pdf_string(text_match.group(2), cmap)
        elif text_match.group(3) is not None:
            text += decode_pdf_hex_string(text_match.group(3), cmap)
    return text

def estimate_pdf_text_width(text, font_size):
    width = 0.0
    for char in text:
        if char.isspace():
            width += font_size * 0.28
        elif char in 'il.,;:\'!|':
            width += font_size * 0.25
        elif char in 'mwMW@#%':
            width += font_size * 0.75
        else:
            width += font_size * 0.50
    return width

def join_pdf_text_segments(segments):
    output = []
    previous = None
    for segment in segments:
        text = segment['text']
        if not text:
            continue
        if previous is None:
            output.append(text.strip())
            previous = segment
            continue

        same_line = (
            previous.get('x') is not None
            and segment.get('x') is not None
            and abs((segment.get('y') or 0) - (previous.get('y') or 0)) < 2.0
        )
        if same_line:
            previous_text = output[-1]
            separator = '' if previous_text.endswith(' ') or text.startswith(' ') else ' '
            output[-1] = f"{previous_text}{separator}{text.strip()}"
        else:
            output.append(text.strip())
        previous = segment
    return '\n'.join(part for part in output if part)

def decode_cmap_hex(value):
    try:
        raw = bytes.fromhex(value)
    except ValueError:
        return ''
    if len(raw) % 2 == 0:
        return raw.decode('utf-16-be', errors='ignore')
    return raw.decode('latin-1', errors='ignore')

def parse_tounicode_cmap(cmap_text):
    cmap = {}
    for src, dst in re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', cmap_text):
        if len(src) <= 2:
            cmap[int(src, 16)] = decode_cmap_hex(dst)

    for src_start, src_end, dst_start in re.findall(
        r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>',
        cmap_text
    ):
        if len(src_start) > 2 or len(src_end) > 2:
            continue
        start = int(src_start, 16)
        end = int(src_end, 16)
        dst = int(dst_start, 16)
        for offset, code in enumerate(range(start, end + 1)):
            cmap[code] = chr(dst + offset)

    return cmap

def decode_pdf_stream_object(object_body):
    stream_match = re.search(rb'stream\r?\n(.*?)\r?\nendstream', object_body, flags=re.S)
    if not stream_match:
        return b''
    stream = stream_match.group(1)
    dictionary = object_body[:stream_match.start()]
    if b'FlateDecode' in dictionary:
        try:
            return zlib.decompress(stream)
        except zlib.error:
            return b''
    return stream

def build_pdf_font_cmaps(pdf_bytes):
    objects = {
        int(match.group(1)): match.group(2)
        for match in re.finditer(rb'(\d+)\s+0\s+obj(.*?)endobj', pdf_bytes, flags=re.S)
    }
    tounicode_maps = {}
    for object_id, body in objects.items():
        stream = decode_pdf_stream_object(body)
        if b'begincmap' in stream:
            cmap_text = stream.decode('latin-1', errors='ignore')
            tounicode_maps[object_id] = parse_tounicode_cmap(cmap_text)

    font_object_cmaps = {}
    for object_id, body in objects.items():
        match = re.search(rb'/ToUnicode\s+(\d+)\s+0\s+R', body)
        if match:
            cmap = tounicode_maps.get(int(match.group(1)))
            if cmap:
                font_object_cmaps[object_id] = cmap

    font_resource_cmaps = {}
    for body in objects.values():
        font_block_match = re.search(rb'/Font\s*<<(.*?)>>', body, flags=re.S)
        if not font_block_match:
            continue
        for name, object_ref in re.findall(rb'/([A-Za-z0-9]+)\s+(\d+)\s+0\s+R', font_block_match.group(1)):
            cmap = font_object_cmaps.get(int(object_ref))
            if cmap:
                font_resource_cmaps[name.decode('ascii', errors='ignore')] = cmap
    return font_resource_cmaps

def extract_text_from_pdf_streams(pdf_bytes):
    extracted = []
    font_cmaps = build_pdf_font_cmaps(pdf_bytes)
    for match in re.finditer(rb'stream\r?\n(.*?)\r?\nendstream', pdf_bytes, flags=re.S):
        stream = match.group(1)
        dictionary = pdf_bytes[max(0, match.start() - 700):match.start()]
        if b'FlateDecode' in dictionary:
            try:
                stream = zlib.decompress(stream)
            except zlib.error:
                continue

        content = stream.decode('latin-1', errors='ignore')
        segments = []
        for block_match in re.finditer(r'BT(.*?)ET', content, flags=re.S):
            block = block_match.group(1)
            font_name_match = re.findall(r'/([A-Za-z0-9]+)\s+[0-9.]+\s+Tf', block)
            font_name = font_name_match[-1] if font_name_match else ''
            cmap = font_cmaps.get(font_name)
            block_text = extract_pdf_block_text(block, cmap)
            if not block_text.strip():
                continue

            font_match = re.findall(r'/[A-Za-z0-9]+\s+([0-9.]+)\s+Tf', block)
            font_size = float(font_match[-1]) if font_match else 10.0
            matrix_match = re.findall(
                r'[-+]?[0-9.]+\s+[-+]?[0-9.]+\s+[-+]?[0-9.]+\s+[-+]?[0-9.]+\s+([-+]?[0-9.]+)\s+([-+]?[0-9.]+)\s+Tm',
                block
            )
            x = y = None
            if matrix_match:
                x = float(matrix_match[-1][0])
                y = float(matrix_match[-1][1])

            segments.append({
                'text': block_text,
                'x': x,
                'y': y,
                'font_size': font_size,
                'end_x': (x + estimate_pdf_text_width(block_text, font_size)) if x is not None else None
            })

        if segments:
            extracted.append(join_pdf_text_segments(segments))

    return '\n'.join(extracted)

def file_sha256(filepath):
    digest = hashlib.sha256()
    with open(filepath, 'rb') as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def get_pdf_text_cache(filepath, mode):
    try:
        file_hash = file_sha256(filepath)
    except Exception:
        return None, None
    return file_hash, PDF_TEXT_CACHE.get((file_hash, mode))


def set_pdf_text_cache(file_hash, mode, text):
    if not file_hash:
        return text
    if len(PDF_TEXT_CACHE) >= PDF_TEXT_CACHE_MAX_ITEMS:
        try:
            PDF_TEXT_CACHE.pop(next(iter(PDF_TEXT_CACHE)))
        except StopIteration:
            pass
    PDF_TEXT_CACHE[(file_hash, mode)] = text or ''
    return text


def extract_pdf_pages_embedded(filepath):
    for module_name in ('pypdf', 'PyPDF2'):
        try:
            module = __import__(module_name)
            reader = module.PdfReader(filepath)
            return [page.extract_text() or '' for page in reader.pages]
        except Exception:
            continue
    return []


def pdf_page_count(filepath):
    try:
        import pypdf
        return len(pypdf.PdfReader(filepath).pages)
    except Exception:
        pass
    try:
        import fitz
        with fitz.open(filepath) as doc:
            return len(doc)
    except Exception:
        return 0


def embedded_pdf_text_is_enough(text):
    normalized = compact_text(text)
    if len(normalized) >= 250:
        return True
    return bool(re.search(
        r'Course\s+(?:Learning\s+Outcomes|Name|Code)|'
        r'[\u0645\u0646]\u062e?[\u0631\u0627\u062c\u0627\u062a]*\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
        r'\u0646\u0648\u0627\u062a\u062c\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
        r'\u0627\u0633\u0645\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
        r'\u0631\u0645\u0632\s+\u0627\u0644\u0645\u0642\u0631\u0631',
        normalized,
        flags=re.I
    ))


def course_spec_heading_page_indexes(page_texts):
    headings = re.compile(
        r'Course\s+Learning\s+Outcomes|Learning\s+Outcomes|'
        r'Knowledge\s+and\s+Understanding|Skills|Values|'
        r'\u0645\u062e\u0631\u062c\u0627\u062a\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
        r'\u0646\u0648\u0627\u062a\u062c\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
        r'\u0627\u0644\u0645\u0639\u0631\u0641\u0629\s+\u0648\u0627\u0644\u0641\u0647\u0645|'
        r'\u0627\u0644\u0645\u0647\u0627\u0631\u0627\u062a|'
        r'\u0627\u0644\u0642\u064a\u0645',
        flags=re.I
    )
    indexes = set()
    for index, page_text in enumerate(page_texts or []):
        if headings.search(normalize_course_spec_text(page_text or '')):
            indexes.update(range(max(0, index - 1), min(len(page_texts), index + 3)))
    return sorted(indexes)


def run_tesseract_on_pdf_pages(filepath, page_indexes, dpi=200, psm='6'):
    tesseract_cmd = find_command('tesseract', [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    ])
    if not tesseract_cmd or not page_indexes:
        return ''
    project_tessdata_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tessdata')
    extracted_parts = []
    image_paths = []
    try:
        import fitz
        doc = fitz.open(filepath)
        matrix = fitz.Matrix(dpi / 72, dpi / 72)
        for page_index in sorted(set(index for index in page_indexes if 0 <= index < len(doc))):
            image_path = get_upload_path(f"{uuid.uuid4()}_ocr_page_{page_index + 1}.png")
            image_paths.append(image_path)
            doc[page_index].get_pixmap(matrix=matrix, alpha=False).save(image_path)
            args = [tesseract_cmd, image_path, 'stdout', '-l', 'ara+eng', '--psm', psm]
            if os.path.exists(os.path.join(project_tessdata_dir, 'ara.traineddata')):
                args.extend(['--tessdata-dir', project_tessdata_dir])
            result = subprocess.run(
                args,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60
            )
            text = result.stdout.decode('utf-8', errors='ignore')
            if compact_text(text):
                extracted_parts.append(f"\n--- OCR page {page_index + 1} ---\n{text}")
    except Exception:
        return ''
    finally:
        for image_path in image_paths:
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                except OSError:
                    pass
    return '\n'.join(extracted_parts)


def initial_ocr_page_indexes(filepath, page_texts=None):
    count = pdf_page_count(filepath)
    if not count and page_texts:
        count = len(page_texts)
    first_pages = set(range(min(count, 5)))
    heading_pages = set(course_spec_heading_page_indexes(page_texts or []))
    return sorted(first_pages | heading_pages)


def targeted_ocr_page_indexes(filepath, page_texts=None):
    count = pdf_page_count(filepath)
    if not count and page_texts:
        count = len(page_texts)
    heading_pages = set(course_spec_heading_page_indexes(page_texts or []))
    if heading_pages:
        return sorted(heading_pages)
    return list(range(min(count, 12)))


def run_pdf_ocr(filepath, page_indexes=None):
    project_tessdata_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tessdata')
    ocrmypdf_cmd = find_command('ocrmypdf')
    pdftoppm_cmd = find_command('pdftoppm', [
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'WinGet', 'Packages', 'oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe', 'poppler-25.07.0', 'Library', 'bin', 'pdftoppm.exe')
    ])
    tesseract_cmd = find_command('tesseract', [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    ])

    if page_indexes is not None:
        text = run_tesseract_on_pdf_pages(filepath, page_indexes, dpi=200, psm='6')
        if compact_text(text):
            return text

    def run_cover_table_ocr_with_fitz():
        if not tesseract_cmd:
            return ''
        cover_table_path = ''
        try:
            import fitz
            from PIL import Image
            doc = fitz.open(filepath)
            if not doc:
                return ''
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image_path = get_upload_path(f"{uuid.uuid4()}_cover_page.png")
            cover_table_path = get_upload_path(f"{uuid.uuid4()}_cover_table.png")
            pix.save(image_path)
            image = Image.open(image_path)
            width, height = image.size
            table = image.crop((
                int(width * 0.14),
                int(height * 0.46),
                int(width * 0.90),
                int(height * 0.80)
            ))
            table.save(cover_table_path)
            args = [tesseract_cmd, cover_table_path, 'stdout', '-l', 'ara+eng', '--psm', '11']
            if os.path.exists(os.path.join(project_tessdata_dir, 'ara.traineddata')):
                args.extend(['--tessdata-dir', project_tessdata_dir])
            result = subprocess.run(
                args,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60
            )
            return result.stdout.decode('utf-8', errors='ignore')
        except Exception:
            return ''
        finally:
            for path in (locals().get('image_path', ''), cover_table_path):
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError:
                        pass

    def run_outcomes_table_ocr_with_fitz():
        if not tesseract_cmd:
            return ''
        page_image_path = ''
        table_path = ''
        try:
            import fitz
            from PIL import Image
            doc = fitz.open(filepath)
            if len(doc) < 4:
                return ''
            pix = doc[3].get_pixmap(matrix=fitz.Matrix(2.8, 2.8), alpha=False)
            page_image_path = get_upload_path(f"{uuid.uuid4()}_outcomes_page.png")
            table_path = get_upload_path(f"{uuid.uuid4()}_outcomes_table.png")
            pix.save(page_image_path)
            image = Image.open(page_image_path)
            width, height = image.size
            table = image.crop((
                int(width * 0.06),
                int(height * 0.36),
                int(width * 0.94),
                int(height * 0.94)
            ))
            table.save(table_path)
            args = [tesseract_cmd, table_path, 'stdout', '-l', 'ara+eng', '--psm', '6']
            if os.path.exists(os.path.join(project_tessdata_dir, 'ara.traineddata')):
                args.extend(['--tessdata-dir', project_tessdata_dir])
            result = subprocess.run(
                args,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60
            )
            return result.stdout.decode('utf-8', errors='ignore')
        except Exception:
            return ''
        finally:
            for path in (page_image_path, table_path):
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError:
                        pass

    if ocrmypdf_cmd:
        output_path = get_upload_path(f"{uuid.uuid4()}_ocr.pdf")
        try:
            subprocess.run(
                [ocrmypdf_cmd, '--language', 'ara+eng', '--skip-text', filepath, output_path],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=180
            )
            return extract_pdf_text(output_path, allow_ocr=False)
        except Exception:
            return ''
        finally:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass

    if pdftoppm_cmd and tesseract_cmd:
        output_prefix = get_upload_path(str(uuid.uuid4()))
        extracted_parts = [
            part for part in [
                run_cover_table_ocr_with_fitz(),
                run_outcomes_table_ocr_with_fitz()
            ] if part
        ]
        image_paths = []
        try:
            subprocess.run(
                [pdftoppm_cmd, '-png', '-r', '200', '-f', '1', '-l', '12', filepath, output_prefix],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=120
            )
            upload_dir = app.config['UPLOAD_FOLDER']
            prefix_name = os.path.basename(output_prefix)
            image_paths = [
                os.path.join(upload_dir, name)
                for name in os.listdir(upload_dir)
                if name.startswith(prefix_name) and name.lower().endswith('.png')
            ]
            sorted_image_paths = sorted(image_paths)
            if sorted_image_paths:
                try:
                    from PIL import Image
                    cover_image = Image.open(sorted_image_paths[0])
                    width, height = cover_image.size
                    cover_table = cover_image.crop((
                        int(width * 0.14),
                        int(height * 0.46),
                        int(width * 0.90),
                        int(height * 0.80)
                    ))
                    cover_table_path = f"{output_prefix}_cover_table.png"
                    cover_table.save(cover_table_path)
                    image_paths.append(cover_table_path)
                    cover_args = [tesseract_cmd, cover_table_path, 'stdout', '-l', 'ara+eng', '--psm', '11']
                    if os.path.exists(os.path.join(project_tessdata_dir, 'ara.traineddata')):
                        cover_args.extend(['--tessdata-dir', project_tessdata_dir])
                    cover_result = subprocess.run(
                        cover_args,
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        timeout=60
                    )
                    extracted_parts.append(cover_result.stdout.decode('utf-8', errors='ignore'))
                except Exception:
                    pass
            for image_path in sorted_image_paths:
                tesseract_args = [tesseract_cmd, image_path, 'stdout', '-l', 'ara+eng']
                if os.path.exists(os.path.join(project_tessdata_dir, 'ara.traineddata')):
                    tesseract_args.extend(['--tessdata-dir', project_tessdata_dir])
                result = subprocess.run(
                    tesseract_args,
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=60
                )
                extracted_parts.append(result.stdout.decode('utf-8', errors='ignore'))
        except Exception:
            return ''
        finally:
            for image_path in image_paths:
                if os.path.exists(image_path):
                    try:
                        os.remove(image_path)
                    except OSError:
                        pass
        return '\n'.join(extracted_parts)

    return ''

def pdf_ocr_available():
    return bool(
        find_command('ocrmypdf')
        or (
            find_command('pdftoppm', [
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'WinGet', 'Packages', 'oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe', 'poppler-25.07.0', 'Library', 'bin', 'pdftoppm.exe')
            ])
            and find_command('tesseract', [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
            ])
        )
    )

def extract_pdf_text(filepath, allow_ocr=True, force_ocr=False, ocr_strategy='initial'):
    cache_mode = f"pdf:{'ocr' if allow_ocr else 'embedded'}:{'force' if force_ocr else 'auto'}:{ocr_strategy}"
    file_hash, cached_text = get_pdf_text_cache(filepath, cache_mode)
    if cached_text is not None:
        return cached_text

    page_texts = extract_pdf_pages_embedded(filepath)
    extracted_text = '\n'.join(page_texts)
    if not compact_text(extracted_text):
        try:
            with open(filepath, 'rb') as f:
                extracted_text = extract_text_from_pdf_streams(f.read())
        except Exception:
            extracted_text = ''

    if not force_ocr and embedded_pdf_text_is_enough(extracted_text):
        return set_pdf_text_cache(file_hash, cache_mode, extracted_text)

    if allow_ocr:
        if force_ocr or ocr_strategy == 'targeted':
            page_indexes = targeted_ocr_page_indexes(filepath, page_texts)
        else:
            page_indexes = initial_ocr_page_indexes(filepath, page_texts)
        ocr_text = run_pdf_ocr(filepath, page_indexes=page_indexes)
        combined_text = '\n'.join(part for part in [extracted_text, ocr_text] if compact_text(part))
        if compact_text(combined_text):
            return set_pdf_text_cache(file_hash, cache_mode, combined_text)

        # Last resort keeps the old behavior available for unusual scanned files.
        full_ocr_text = run_pdf_ocr(filepath)
        combined_text = '\n'.join(part for part in [combined_text, full_ocr_text] if compact_text(part))
        return set_pdf_text_cache(file_hash, cache_mode, combined_text)

    return set_pdf_text_cache(file_hash, cache_mode, extracted_text)

def extract_docx_text(filepath):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    def element_text(element):
        parts = []
        for node in element.iter():
            tag = node.tag.rsplit('}', 1)[-1] if isinstance(node.tag, str) else ''
            if tag == 't':
                parts.append(node.text or '')
            elif tag == 'tab':
                parts.append(' ')
            elif tag in {'br', 'cr'}:
                parts.append('\n')
        return compact_text(''.join(parts).replace('\xa0', ' '))

    try:
        with zipfile.ZipFile(filepath) as docx:
            xml = docx.read('word/document.xml')
        root = ET.fromstring(xml)
    except Exception:
        return ''

    body = root.find('.//w:body', ns)
    if body is None:
        body = root

    lines = []
    for child in body:
        tag = child.tag.rsplit('}', 1)[-1] if isinstance(child.tag, str) else ''
        if tag == 'p':
            text = element_text(child)
            if text:
                lines.append(text)
        elif tag == 'tbl':
            for tr in child.findall('./w:tr', ns):
                cells = []
                for tc in tr.findall('./w:tc', ns):
                    cell_parts = [element_text(p) for p in tc.findall('./w:p', ns)]
                    cell_text = compact_text(' '.join(part for part in cell_parts if part))
                    if cell_text:
                        cells.append(cell_text)
                if cells:
                    lines.append(' | '.join(cells))

    return '\n'.join(line for line in lines if compact_text(line))

def extract_document_text(filepath, allow_ocr=True):
    if str(filepath).lower().endswith('.docx'):
        return extract_docx_text(filepath)
    return extract_pdf_text(filepath, allow_ocr)


def course_spec_extraction_score(extracted):
    extracted = extracted or {}
    score = 0
    if compact_text(extracted.get('course_name')):
        score += 2
    if compact_text(extracted.get('course_code')):
        score += 2
    if compact_text(extracted.get('college')):
        score += 1
    if compact_text(extracted.get('department')):
        score += 1
    clos = extracted.get('clos') or []
    score += min(len(clos), 10) * 3
    return score


def course_spec_extraction_is_usable(extracted):
    extracted = extracted or {}
    return bool(
        compact_text(extracted.get('course_name') or extracted.get('name'))
        and compact_text(extracted.get('course_code') or extracted.get('course_number'))
        and extracted.get('clos')
    )


def get_gemini_spec_cache(filepath):
    try:
        file_hash = file_sha256(filepath)
    except Exception:
        return None, None
    return file_hash, GEMINI_SPEC_CACHE.get((file_hash, GEMINI_MODEL))


def set_gemini_spec_cache(file_hash, extracted):
    if not file_hash:
        return extracted
    if len(GEMINI_SPEC_CACHE) >= GEMINI_SPEC_CACHE_MAX_ITEMS:
        try:
            GEMINI_SPEC_CACHE.pop(next(iter(GEMINI_SPEC_CACHE)))
        except StopIteration:
            pass
    GEMINI_SPEC_CACHE[(file_hash, GEMINI_MODEL)] = extracted
    return extracted


def gemini_course_spec_prompt():
    return (
        "Extract course specification information from this document and return valid JSON only "
        "(no markdown, explanations, comments, or code fences).\n\n"
        "The document may be in Arabic or English.\n\n"
        "General Rules:\n\n"
        "* Extract information only from the document content.\n"
        "* Do not infer, guess, or generate missing values.\n"
        "* Do not infer the course code, course title, or any field from the file name.\n"
        "* If a value cannot be reliably identified, return an empty string (\"\").\n"
        "* Preserve the original wording whenever possible.\n"
        "* Preserve Arabic text exactly when it is readable.\n"
        "* If Arabic text is corrupted by PDF extraction (fragmented letters, isolated characters, "
        "excessive spaces, repeated underscores, broken RTL ordering, or unreadable OCR output), "
        "do not return the corrupted text.\n"
        "* Reconstruct Arabic text only when the intended wording is clearly recoverable from context; "
        "otherwise return an empty string.\n"
        "* Remove obvious OCR artifacts, page headers/footers, page numbers, and duplicated extraction noise.\n\n"
        "Specific Extraction Rules:\n\n"
        "* Extract Course Learning Outcomes (CLOs).\n"
        "* Extract course topics from the Course Content/List of Topics section as separate ordered items.\n"
        "* Extract all associated Program Learning Outcome (PLO) codes from the corresponding "
        "PLO/program outcome column.\n"
        "* Do not merge multiple topics into one item unless they appear as a single topic in the specification.\n\n"
        "Output Requirements:\n\n"
        "* Return JSON only.\n"
        "* Ensure the JSON is valid and parseable.\n"
        "* Do not include explanatory text before or after the JSON.\n\n"
        "Required JSON format:\n"
        '{"course_name":"","course_code":"","college":"","department":"","program":"",'
        '"clos":["1.1 outcome text","1.2 outcome text","2.1 outcome text","3.1 outcome text"],'
        '"topics":["topic 1","topic 2"],"clo_plos":{"1.1":"K1","2.1":"S2"}}'
    )


def parse_gemini_json_response(text):
    text = compact_text(text or '')
    if not text:
        return None
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.I)
    text = re.sub(r'\s*```$', '', text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    start = text.find('{')
    end = text.rfind('}')
    if start >= 0 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            return None
    return None


def call_groq_json_with_error(system_prompt, user_payload, timeout=90, model_name=None):
    if not GROQ_KEY:
        return None, "GROQ_KEY is not configured."
    try:
        payload = {
            'model': model_name or GROQ_MODEL,
            'temperature': 0,
            'response_format': {'type': 'json_object'},
            'messages': [
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_payload},
            ],
        }
        groq_request = urllib.request.Request(
            GROQ_API_URL,
            data=json.dumps(payload).encode('utf-8'),
            headers={
                'Authorization': f'Bearer {GROQ_KEY}',
                'Content-Type': 'application/json',
                'User-Agent': 'ETQAN-GroqQwenFallback/1.0',
            },
            method='POST',
        )
        with urllib.request.urlopen(groq_request, timeout=timeout) as response:
            groq_payload = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode('utf-8', errors='replace')
        app.logger.warning("Groq/Qwen request failed with HTTP %s: %s", exc.code, body[:800])
        return None, f"HTTP {exc.code}: {body[:500]}"
    except Exception as exc:
        app.logger.warning("Groq/Qwen request failed: %s", exc)
        return None, str(exc)

    content = ''
    for choice in groq_payload.get('choices') or []:
        message = choice.get('message') or {}
        content += message.get('content') or ''
    parsed = parse_gemini_json_response(content)
    if parsed is None:
        return None, "Qwen/Groq returned a response, but it was not valid JSON."
    return parsed, ''


def call_groq_json(system_prompt, user_payload, timeout=90, model_name=None):
    parsed, _error = call_groq_json_with_error(system_prompt, user_payload, timeout, model_name=model_name)
    return parsed

def groq_extraction_method_for_model(model_name=None):
    model = str(model_name or GROQ_MODEL or '').lower()
    if 'qwen' in model:
        return 'qwen'
    if 'llama' in model:
        return 'llama'
    return 'groq'


def call_gemini_json(system_prompt, user_payload, timeout=90):
    if not GEMINI_API_KEY:
        return None
    try:
        payload = {
            'contents': [
                {
                    'role': 'user',
                    'parts': [
                        {'text': f"{system_prompt}\n\n{user_payload}"},
                    ],
                }
            ],
            'generationConfig': {
                'temperature': 0.2,
                'responseMimeType': 'application/json',
            },
        }
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            + urllib.parse.quote(GEMINI_MODEL, safe='')
            + ":generateContent?key="
            + urllib.parse.quote(GEMINI_API_KEY, safe='')
        )
        gemini_request = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={
                'Content-Type': 'application/json',
                'User-Agent': 'ETQAN-CourseReportAI/1.0',
            },
            method='POST',
        )
        with urllib.request.urlopen(gemini_request, timeout=timeout) as response:
            gemini_payload = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode('utf-8', errors='replace')
        app.logger.warning("Gemini course report insight request failed with HTTP %s: %s", exc.code, body[:800])
        return None
    except Exception as exc:
        app.logger.warning("Gemini course report insight request failed: %s", exc)
        return None

    response_text = ''
    for candidate in gemini_payload.get('candidates') or []:
        content = candidate.get('content') or {}
        for part in content.get('parts') or []:
            response_text += part.get('text') or ''
    return parse_gemini_json_response(response_text)


def normalize_gemini_clos(raw_clos):
    if isinstance(raw_clos, dict):
        raw_clos = [
            {'code': key, 'text': value}
            for key, value in raw_clos.items()
        ]
    clos_by_id = {}
    for item in raw_clos or []:
        if isinstance(item, dict):
            clo_id = compact_text(
                item.get('code')
                or item.get('id')
                or item.get('clo_id')
                or item.get('outcome_code')
                or item.get('number')
            )
            body = compact_text(
                item.get('text')
                or item.get('description')
                or item.get('outcome')
                or item.get('learning_outcome')
                or item.get('value')
            )
        else:
            clo_id = ''
            body = compact_text(str(item or ''))

        match = re.match(r'^(?:CLO\s*)?([123]\s*[\.\-]\s*[1-9]\d*)\s+(.+)$', body, flags=re.I)
        if match:
            clo_id = clo_id or match.group(1)
            body = match.group(2)

        clo_id = re.sub(r'\s+', '', clo_id or '').upper().replace('-', '.')
        clo_id = re.sub(r'^CLO', '', clo_id)
        if not re.match(r'^[123]\.[1-9]\d*$', clo_id or ''):
            continue
        body = clean_clo_text(body)
        if not body or len(body) < 4:
            continue
        clos_by_id[clo_id] = f"{clo_id} {body}"

    ordered_ids = sorted(
        clos_by_id,
        key=lambda value: tuple(int(part) for part in value.split('.', 1))
    )
    return [clos_by_id[clo_id] for clo_id in ordered_ids]


def normalize_gemini_course_spec(payload):
    if not isinstance(payload, dict):
        return None

    course_name = clean_pdf_fragment(compact_text(
        payload.get('course_name')
        or payload.get('name')
        or payload.get('course_title')
        or ''
    ))
    course_code = normalize_extracted_course_code(compact_text(
        payload.get('course_code')
        or payload.get('course_number')
        or payload.get('code')
        or ''
    ))
    college = clean_pdf_fragment(compact_text(payload.get('college') or payload.get('faculty') or ''))
    department = clean_pdf_fragment(compact_text(payload.get('department') or payload.get('dept') or ''))
    program = clean_pdf_fragment(compact_text(payload.get('program') or payload.get('programme') or ''))
    clos = normalize_gemini_clos(payload.get('clos') or payload.get('clos_by_domain') or payload.get('learning_outcomes'))

    display_name = course_name
    if course_code and course_name and course_code not in course_name:
        display_name = f"{course_name} ({course_code})"
    elif course_code and not course_name:
        display_name = course_code

    clo_plos = payload.get('clo_plos') if isinstance(payload.get('clo_plos'), dict) else {}
    topics = payload.get('topics') if isinstance(payload.get('topics'), list) else []
    topics = [compact_text(topic) for topic in topics if compact_text(topic)]

    return {
        'name': display_name,
        'course_name': course_name,
        'course_code': course_code,
        'course_number': course_code,
        'college': college,
        'department': department,
        'program': program,
        'clos': clos,
        'topics': topics,
        'clo_plos': clo_plos,
        'grouped_clos': group_clos_by_domain(clos),
        'extraction_method': 'gemini'
    }


def extract_course_spec_with_gemini(filepath, filename=''):
    file_ext = os.path.splitext(filename or filepath)[1].lower()
    if not GEMINI_API_KEY or file_ext != '.pdf':
        return None, "GEMINI_API_KEY is not configured or file is not PDF."
    try:
        if os.path.getsize(filepath) > GEMINI_MAX_INLINE_BYTES:
            app.logger.info("Skipping Gemini course specification extraction because file is larger than GEMINI_MAX_INLINE_BYTES.")
            return None, "File is larger than GEMINI_MAX_INLINE_BYTES."
    except OSError as e:
        return None, str(e)

    file_hash, cached = get_gemini_spec_cache(filepath)
    if cached is not None:
        if cached is False:
            return None, "Using cached failure."
        return cached, ""

    try:
        with open(filepath, 'rb') as file:
            encoded_pdf = base64.b64encode(file.read()).decode('ascii')
        payload = {
            'contents': [
                {
                    'parts': [
                        {'text': gemini_course_spec_prompt()},
                        {
                            'inline_data': {
                                'mime_type': 'application/pdf',
                                'data': encoded_pdf
                            }
                        }
                    ]
                }
            ],
            'generationConfig': {
                'temperature': 0,
                'responseMimeType': 'application/json'
            }
        }
        endpoint = (
            'https://generativelanguage.googleapis.com/v1beta/models/'
            + urllib.parse.quote(GEMINI_MODEL, safe='')
            + ':generateContent?key='
            + urllib.parse.quote(GEMINI_API_KEY, safe='')
        )
        request_data = json.dumps(payload).encode('utf-8')
        gemini_request = urllib.request.Request(
            endpoint,
            data=request_data,
            headers={
                'Content-Type': 'application/json',
                'User-Agent': 'ETQAN-CourseSpecExtractor/1.0'
            },
            method='POST'
        )
        with urllib.request.urlopen(gemini_request, timeout=90) as response:
            gemini_payload = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode('utf-8', errors='replace')
        app.logger.warning("Gemini course specification extraction failed with HTTP %s: %s", exc.code, body[:800])
        set_gemini_spec_cache(file_hash, False)
        return None, f"HTTP {exc.code}: {body[:200]}"
    except Exception as exc:
        app.logger.warning("Gemini course specification extraction failed: %s", exc)
        set_gemini_spec_cache(file_hash, False)
        return None, str(exc)

    response_text = ''
    for candidate in gemini_payload.get('candidates') or []:
        content = candidate.get('content') or {}
        for part in content.get('parts') or []:
            response_text += part.get('text') or ''

    extracted = normalize_gemini_course_spec(parse_gemini_json_response(response_text))
    if course_spec_extraction_is_usable(extracted):
        set_gemini_spec_cache(file_hash, extracted)
        return extracted, ""
    app.logger.info("Gemini course specification extraction did not return complete course data; falling back to local parser/OCR.")
    set_gemini_spec_cache(file_hash, False)
    return None, "Gemini did not return complete course data (missing Name, Code, or CLOs)."


def extract_course_spec_with_groq_text(text, model_name=None):
    if not GROQ_KEY or not compact_text(text):
        return None, "GROQ_KEY is not configured or no text extracted."
    user_payload = (
        gemini_course_spec_prompt()
        + "\n\nCourse specification text:\n"
        + compact_text(text)[:60000]
    )
    parsed, error = call_groq_json_with_error(
        "You extract structured course specification data. Return valid JSON only.",
        user_payload,
        model_name=model_name
    )
    if error:
        return None, error
    extracted = normalize_gemini_course_spec(parsed)
    if course_spec_extraction_is_usable(extracted):
        extracted['extraction_method'] = groq_extraction_method_for_model(model_name)
        return extracted, ""
    return None, "Groq did not return complete course data."


def clean_ncaaa_pdf_layout_text(value):
    text = normalize_course_spec_text(value)
    replacements = {
        'أخصاEي': 'أخصائي',
        'معPد': 'معهد',
        'الاسUشارات': 'الاستشارات',
        'الاسwشارات': 'الاستشارات',
        'و^فسر': 'ويفسر',
        'ويفسرمفاهيم': 'ويفسر مفاهيم',
        'مفا`يم': 'مفاهيم',
        '@عمل': 'بعمل',
        'الداخjk': 'الداخلي',
        'العمjk': 'العملي',
        'الداخl†': 'الداخلي',
        ' lk ': ' في ',
        'واpqاص': 'والخاص',
        'المtام': 'المهام',
        '{عد': 'يعد',
        'تقار^ر': 'تقارير',
        'تقارcر': 'تقارير',
        'مكتو|ة': 'مكتوبة',
        'تقاريرمكتوبة': 'تقارير مكتوبة',
        'و^وصل': 'ويوصل',
        'اQqتلفة': 'المختلفة',
        'اللƒ„': 'التي',
        'المtنة': 'المهنة',
        'أدأب': 'آداب',
        'اZ[دود': 'الحدود',
        'مفPوم': 'مفهوم',
        'وأsداف': 'وأهداف',
        'ا—^اطر': 'المخاطر',
        'تحسŽن': 'تحسين',
        'بTMئة': 'بيئة',
        '›†': 'في',
        'بŽن': 'بين',
        'واZ^ار•†': 'والخارجي',
        "ا'’موع": '',
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = text.replace('ويفسرمفاهيم', 'ويفسر مفاهيم')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([:،,.؛])', r'\1', text)
    return text.strip()


def join_pdf_words_rtl(words, same_line_tolerance=4, word_gap=1.5):
    rows = []
    for word in sorted(words or [], key=lambda item: item[1]):
        if not rows or abs(rows[-1][0] - word[1]) > same_line_tolerance:
            rows.append([word[1], [word]])
        else:
            rows[-1][1].append(word)

    lines = []
    for _, row_words in rows:
        parts = []
        previous_x0 = None
        for word in sorted(row_words, key=lambda item: item[0], reverse=True):
            x0, _y0, x1, _y1, text = word[:5]
            if previous_x0 is not None and previous_x0 - x1 > word_gap:
                parts.append(' ')
            parts.append(text)
            previous_x0 = x0
        lines.append(''.join(parts))

    return clean_ncaaa_pdf_layout_text(' '.join(lines))


def value_after_pdf_colon(line):
    line = clean_ncaaa_pdf_layout_text(line)
    if ':' in line:
        return clean_ncaaa_pdf_layout_text(line.split(':', 1)[1])
    if '؛' in line:
        return clean_ncaaa_pdf_layout_text(line.split('؛', 1)[1])
    return line


def extract_ncaaa_pdf_cover_metadata(doc):
    if not doc:
        return {}
    try:
        words = doc[0].get_text('words')
    except Exception:
        return {}

    page_width = float(doc[0].rect.width or 595)
    candidate_words = [
        word for word in words
        if 0.40 * page_width <= word[0] <= 0.86 * page_width
        and 390 <= word[1] <= 590
    ]
    rows = []
    for word in sorted(candidate_words, key=lambda item: item[1]):
        if not rows or abs(rows[-1][0] - word[1]) > 8:
            rows.append([word[1], [word]])
        else:
            rows[-1][1].append(word)

    row_values = [value_after_pdf_colon(join_pdf_words_rtl(row_words)) for _, row_words in rows]
    metadata = {}
    fields = ['course_name', 'course_code', 'program', 'department', 'college']
    for field, value in zip(fields, row_values):
        if not value:
            continue
        if field == 'course_code':
            code_match = re.search(r'\b([A-Z]{2,8}[-\s]?\d{2,5}[A-Z]?)\b', value, flags=re.I)
            metadata[field] = normalize_extracted_course_code(code_match.group(1) if code_match else value)
        else:
            metadata[field] = clean_arabic_metadata_value(value, field)
    return metadata


def extract_ncaaa_pdf_layout_clos(doc):
    clos = {}
    if not doc:
        return clos

    for page_index in range(len(doc)):
        try:
            page = doc[page_index]
            words = page.get_text('words')
        except Exception:
            continue

        page_width = float(page.rect.width or 595)
        code_x_min = page_width * 0.84
        outcome_x_min = page_width * 0.60
        outcome_x_max = page_width * 0.84
        codes = []
        for word in words:
            x0, y0, _x1, _y1, text = word[:5]
            if x0 >= code_x_min and re.fullmatch(r'[123]\.\d+', str(text or '')):
                codes.append((text, y0))
        codes = sorted(codes, key=lambda item: item[1])
        if not codes:
            continue

        for index, (code, y0) in enumerate(codes):
            previous_y = codes[index - 1][1] if index else y0 - 44
            next_y = codes[index + 1][1] if index + 1 < len(codes) else y0 + 110
            top = (previous_y + y0) / 2 if index else y0 - 22
            bottom = (y0 + next_y) / 2 if index + 1 < len(codes) else y0 + 55
            if code.endswith('.0'):
                continue

            outcome_words = [
                word for word in words
                if outcome_x_min <= word[0] <= outcome_x_max
                and top - 5 <= word[1] <= bottom + 5
            ]
            body = join_pdf_words_rtl(outcome_words)
            body = clean_clo_text(clean_arabic_outcome_line(body))
            if body and len(body) >= 8:
                clos[code] = f"{code} {body}"

    return final_clean_clo_map(clos)


def extract_ncaaa_pdf_layout_clo_plos(doc, clos=None):
    mapping = {}
    if not doc:
        return mapping

    clo_lookup = {
        clo_number_key(clo) or clo_number(clo): clo
        for clo in (clos or [])
        if clo_number_key(clo) or clo_number(clo)
    }

    for page_index in range(len(doc)):
        try:
            page = doc[page_index]
            words = page.get_text('words')
        except Exception:
            continue

        page_width = float(page.rect.width or 595)
        code_x_min = page_width * 0.84
        plo_x_min = page_width * 0.42
        plo_x_max = page_width * 0.63
        codes = []
        for word in words:
            x0, y0, _x1, _y1, text = word[:5]
            if x0 >= code_x_min and re.fullmatch(r'[123]\.\d+', str(text or '')):
                codes.append((str(text), y0))
        codes = sorted(codes, key=lambda item: item[1])
        if not codes:
            continue

        for index, (code, y0) in enumerate(codes):
            if code.endswith('.0'):
                continue
            previous_y = codes[index - 1][1] if index else y0 - 44
            next_y = codes[index + 1][1] if index + 1 < len(codes) else y0 + 110
            top = (previous_y + y0) / 2 if index else y0 - 22
            bottom = (y0 + next_y) / 2 if index + 1 < len(codes) else y0 + 55

            plo_words = [
                word for word in words
                if plo_x_min <= word[0] <= plo_x_max
                and top - 5 <= word[1] <= bottom + 5
            ]
            plo_text = join_pdf_words_rtl(plo_words)
            extracted_codes = extract_plo_codes(plo_text)
            if not extracted_codes:
                extracted_code = extract_noisy_arabic_plo_code(plo_text, clo_lookup.get(code, code))
                extracted_codes = [extracted_code] if extracted_code else []
            if extracted_codes:
                mapping[code] = ', '.join(extracted_codes)

    return mapping


def extract_ncaaa_pdf_layout_topics(doc):
    if not doc:
        return []

    for page_index in range(len(doc)):
        try:
            page = doc[page_index]
            words = page.get_text('words')
        except Exception:
            continue

        page_width = float(page.rect.width or 595)
        code_x_min = page_width * 0.84
        topic_x_min = page_width * 0.30
        topic_x_max = page_width * 0.84
        codes = []
        for word in words:
            x0, y0, _x1, _y1, text = word[:5]
            if x0 >= code_x_min and y0 > 320 and re.fullmatch(r'\d{1,2}', str(text or '')):
                number = int(text)
                if 1 <= number <= 30:
                    codes.append((number, y0))
        codes = sorted(codes, key=lambda item: item[1])
        if len(codes) < 5 or codes[0][0] != 1:
            continue

        topics = []
        for index, (number, y0) in enumerate(codes):
            previous_y = codes[index - 1][1] if index else y0 - 35
            next_y = codes[index + 1][1] if index + 1 < len(codes) else y0 + 35
            top = (previous_y + y0) / 2 if index else y0 - 16
            bottom = (y0 + next_y) / 2 if index + 1 < len(codes) else y0 + 24
            topic_words = [
                word for word in words
                if topic_x_min <= word[0] <= topic_x_max
                and top - 3 <= word[1] <= bottom + 3
            ]
            topic = clean_arabic_topic_ocr_artifacts(join_pdf_words_rtl(topic_words))
            topic = re.sub(r'^\d+\s+', '', topic).strip(' .')
            if topic and contains_arabic(topic):
                topics.append(topic)
        if topics:
            return topics

    return []


def extract_course_spec_from_pdf_layout(filepath, text=''):
    try:
        import fitz
        with fitz.open(filepath) as doc:
            metadata = extract_ncaaa_pdf_cover_metadata(doc)
            clo_map = extract_ncaaa_pdf_layout_clos(doc)
            clo_plos = extract_ncaaa_pdf_layout_clo_plos(doc, list(clo_map.values()))
            topics = extract_ncaaa_pdf_layout_topics(doc)
    except Exception:
        return None

    if not metadata and not clo_map:
        return None

    course_name = metadata.get('course_name', '')
    course_code = metadata.get('course_code', '')
    display_name = course_name
    if course_code and course_name and course_code not in course_name:
        display_name = f"{course_name} ({course_code})"
    elif course_code and not course_name:
        display_name = course_code

    clos = list(clo_map.values())
    return {
        'name': display_name,
        'course_name': course_name,
        'course_code': course_code,
        'course_number': course_code,
        'college': metadata.get('college', ''),
        'department': metadata.get('department', ''),
        'program': metadata.get('program', ''),
        'clos': clos,
        'topics': topics or (extract_course_topics(text) if compact_text(text) else []),
        'clo_plos': clo_plos or (extract_clo_plo_mapping(text, clos) if compact_text(text) else {}),
        'grouped_clos': group_clos_by_domain(clos),
        'extraction_method': 'local'
    }


def extract_course_spec_document(filepath, filename=''):
    extraction_start = time.perf_counter()
    file_ext = os.path.splitext(filename or filepath)[1].lower()
    ai_diagnostics = []
    if file_ext == '.docx':
        text = extract_docx_text(filepath)
        extracted = extract_course_spec_metadata(text, filename)
        extracted['extraction_method'] = 'local'
        extracted['extraction_metadata'] = {
            'task': 'course_specification_extraction',
            'source': 'local',
            'model': 'local-docx-parser',
            'duration_seconds': elapsed_seconds(extraction_start),
            'filename': filename or os.path.basename(filepath),
        }
        return text, extracted

    gemini_extracted, _gemini_error = extract_course_spec_with_gemini(filepath, filename)
    if _gemini_error:
        ai_diagnostics.append({
            'provider': 'Gemini',
            'status': 'skipped_or_failed',
            'message': _gemini_error,
        })
    if gemini_extracted:
        gemini_extracted.setdefault('extraction_metadata', {
            'task': 'course_specification_extraction',
            'source': 'gemini',
            'model': GEMINI_MODEL,
            'duration_seconds': elapsed_seconds(extraction_start),
            'filename': filename or os.path.basename(filepath),
        })
        return '', gemini_extracted

    text = extract_pdf_text(filepath, allow_ocr=False)
    if not compact_text(text):
        text = extract_pdf_text(filepath, allow_ocr=True)
    elif arabic_text_layer_looks_fragmented(text):
        page_texts = extract_pdf_pages_embedded(filepath)
        ocr_text = run_pdf_ocr(filepath, page_indexes=targeted_ocr_page_indexes(filepath, page_texts))
        if compact_text(ocr_text):
            ai_diagnostics.append({
                'provider': 'Local text layer',
                'status': 'replaced',
                'message': 'Embedded Arabic PDF text looked fragmented, so OCR text was used before AI extraction.',
            })
            text = ocr_text
        
    groq_extracted, _groq_error = extract_course_spec_with_groq_text(text, model_name="llama-3.1-8b-instant")
    if _groq_error:
        ai_diagnostics.append({
            'provider': 'Llama via Groq',
            'status': 'skipped_or_failed',
            'message': _groq_error,
        })
    if groq_extracted:
        if course_spec_extracted_arabic_looks_fragmented(groq_extracted):
            ai_diagnostics.append({
                'provider': 'Llama via Groq',
                'status': 'rejected',
                'message': 'Llama returned fragmented Arabic text from the PDF text layer; using local/OCR extraction instead.',
            })
            groq_extracted = None
        else:
            groq_extracted['extraction_metadata'] = {
                'task': 'course_specification_extraction',
                'source': groq_extracted.get('extraction_method') or 'groq',
                'model': 'llama-3.1-8b-instant',
                'duration_seconds': elapsed_seconds(extraction_start),
                'filename': filename or os.path.basename(filepath),
                'ai_diagnostics': ai_diagnostics,
            }
            return text, groq_extracted
    ocr_text = extract_pdf_text(filepath, allow_ocr=True)
    if len(compact_text(ocr_text)) > len(compact_text(text)):
        text = ocr_text
    extracted = extract_course_spec_metadata(text, filename)
    layout_extracted = extract_course_spec_from_pdf_layout(filepath, text)
    if course_spec_extraction_score(layout_extracted) > course_spec_extraction_score(extracted):
        extracted = layout_extracted
    extracted['extraction_method'] = 'local'
    extracted['extraction_metadata'] = {
        'task': 'course_specification_extraction',
        'source': 'local',
        'model': 'local-pdf-parser',
        'duration_seconds': elapsed_seconds(extraction_start),
        'filename': filename or os.path.basename(filepath),
        'ai_diagnostics': ai_diagnostics,
    }
    if extracted.get('course_name') and extracted.get('course_code') and extracted.get('clos'):
        return text, extracted

    targeted_text = extract_pdf_text(filepath, allow_ocr=True, force_ocr=True, ocr_strategy='targeted')
    if compact_text(targeted_text):
        targeted_extracted = extract_course_spec_metadata(targeted_text, filename)
        targeted_extracted['extraction_method'] = 'local'
        targeted_extracted['extraction_metadata'] = {
            'task': 'course_specification_extraction',
            'source': 'local',
            'model': 'local-targeted-ocr-parser',
            'duration_seconds': elapsed_seconds(extraction_start),
            'filename': filename or os.path.basename(filepath),
            'ai_diagnostics': ai_diagnostics,
        }
        targeted_layout_extracted = extract_course_spec_from_pdf_layout(filepath, targeted_text)
        if course_spec_extraction_score(targeted_layout_extracted) > course_spec_extraction_score(targeted_extracted):
            targeted_extracted = targeted_layout_extracted
            targeted_extracted['extraction_method'] = 'local'
            targeted_extracted['extraction_metadata'] = {
                'task': 'course_specification_extraction',
                'source': 'local',
                'model': 'local-targeted-layout-parser',
                'duration_seconds': elapsed_seconds(extraction_start),
                'filename': filename or os.path.basename(filepath),
                'ai_diagnostics': ai_diagnostics,
            }
        if course_spec_extraction_score(targeted_extracted) > course_spec_extraction_score(extracted):
            return targeted_text, targeted_extracted
    return text, extracted


def flash_course_spec_extraction_method(extracted):
    method = (extracted or {}).get('extraction_method')
    if method == 'gemini':
        method_label = translate('courses.extraction_method_gemini')
    elif method == 'qwen':
        method_label = translate('courses.extraction_method_qwen')
    elif method == 'llama':
        method_label = translate('courses.extraction_method_llama')
    elif method == 'groq':
        method_label = translate('courses.extraction_method_groq')
    else:
        method_label = translate('courses.extraction_method_local')
    message = f"{translate('courses.extraction_method_prefix')} {method_label}"
    diagnostics = ((extracted or {}).get('extraction_metadata') or {}).get('ai_diagnostics') or []
    if diagnostics:
        reason_parts = []
        for item in diagnostics:
            provider = compact_text(item.get('provider') or '')
            reason = compact_text(item.get('message') or '')
            if not provider or not reason:
                continue
            if len(reason) > 180:
                reason = reason[:177].rstrip() + '...'
            reason_parts.append(f"{provider}: {reason}")
        if reason_parts:
            message += " — " + ("؛ ".join(reason_parts) if get_language() == 'ar' else "; ".join(reason_parts))
    flash(message, "info")


def extract_exam_paper_text(filepath, file_ext):
    if file_ext == '.pdf':
        return extract_pdf_text(filepath)
    if file_ext == '.docx':
        return extract_docx_text(filepath)
    if file_ext == '.txt':
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    return ''

def question_number_from_label(value):
    text = str(value or '')
    patterns = [
        r'\bQ(?:uestion)?\s*[-#:.]?\s*(\d{1,3})\b',
        r'\bQuestion\s+No\.?\s*(\d{1,3})\b',
        r'\bAnswers?\s*[-_ #:.]?\s*(\d{1,3})\b',
        r'\bItems?\s*[-_ #:.]?\s*(\d{1,3})\b'
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return int(match.group(1))
    return None

def detect_clo_tags_from_text(text):
    tags = []
    for match in re.finditer(r'\b(?:CLO\s*[-_ ]?\d+(?:\.\d+)*|[123]\.\d+)\b', text or '', flags=re.I):
        tag = re.sub(r'\s+', '', match.group(0).upper())
        tag = re.sub(r'^CLO', 'CLO', tag)
        if tag not in tags:
            tags.append(tag)
    return tags

SEMANTIC_STOPWORDS = {
    'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'in', 'into', 'is', 'it', 'of',
    'on', 'or', 'that', 'the', 'to', 'with', 'within', 'without', 'using', 'use', 'used', 'given',
    'question', 'answer', 'mark', 'marks', 'score', 'points', 'student', 'students', 'following',
    'below', 'above', 'write', 'show', 'explain', 'describe', 'discuss', 'compare', 'solve',
    'في', 'من', 'على', 'الى', 'إلى', 'عن', 'مع', 'او', 'أو', 'و', 'ثم', 'ذلك', 'هذه', 'هذا',
    'التي', 'الذي', 'أن', 'ان', 'كان', 'كانت', 'يكون', 'السؤال', 'اجب', 'أجب', 'وضح', 'اشرح',
    'اذكر', 'ناقش', 'قارن', 'حل', 'اكتب', 'التالي', 'الآتية', 'الاتية', 'درجة', 'درجات'
}

SEMANTIC_SYNONYMS = {
    'classify': 'categorize',
    'classification': 'categorize',
    'identify': 'recognize',
    'recognise': 'recognize',
    'demonstrate': 'show',
    'implement': 'develop',
    'create': 'develop',
    'design': 'develop',
    'program': 'code',
    'programming': 'code',
    'algorithm': 'algorithms',
    'ethical': 'ethics',
    'regulation': 'compliance',
    'regulations': 'compliance',
    'policy': 'compliance',
    'policies': 'compliance',
    'صنف': 'تصنيف',
    'يصنف': 'تصنيف',
    'حدد': 'تعرف',
    'يتعرف': 'تعرف',
    'اشرح': 'شرح',
    'يوضح': 'شرح',
    'طبق': 'تطبيق',
    'يطبق': 'تطبيق',
    'حل': 'حل',
    'يصمم': 'تصميم',
    'صمم': 'تصميم',
    'برمج': 'برمجة',
    'قارن': 'مقارنة',
    'قيم': 'تقييم',
    'قيّم': 'تقييم',
    'اخلاقي': 'اخلاقيات',
    'الأخلاقي': 'اخلاقيات',
    'اخلاقيات': 'اخلاقيات'
}

ACTION_CUES = {
    'knowledge': {
        'define', 'describe', 'explain', 'identify', 'recognize', 'list', 'state', 'name', 'classify',
        'categorize', 'ذكر', 'اذكر', 'عرف', 'تعرف', 'شرح', 'اشرح', 'وضح', 'صنف', 'تصنيف'
    },
    'skills': {
        'apply', 'solve', 'calculate', 'compute', 'analyze', 'analyse', 'compare', 'evaluate',
        'design', 'develop', 'implement', 'code', 'construct', 'use', 'تطبيق', 'طبق', 'حل',
        'احسب', 'حلل', 'قارن', 'تقييم', 'قيم', 'تصميم', 'صمم', 'طور', 'برمجة'
    },
    'values': {
        'justify', 'ethics', 'ethical', 'responsible', 'responsibility', 'professional', 'compliance',
        'privacy', 'fair', 'transparent', 'appreciate', 'recognize', 'commit', 'تبرير', 'يبرر',
        'اخلاقيات', 'مسؤولية', 'مهني', 'خصوصية', 'امتثال', 'عدالة', 'شفافية', 'يلتزم', 'يقدر'
    }
}

def normalize_semantic_text(value):
    text = str(value or '').lower()
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = text.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ى', 'ي').replace('ة', 'ه')
    text = re.sub(r'[^\w\s\u0600-\u06ff]', ' ', text, flags=re.UNICODE)
    return re.sub(r'\s+', ' ', text).strip()

def semantic_tokens(value):
    tokens = []
    for token in normalize_semantic_text(value).split():
        if len(token) <= 1 or token in SEMANTIC_STOPWORDS:
            continue
        token = SEMANTIC_SYNONYMS.get(token, token)
        if token not in SEMANTIC_STOPWORDS:
            tokens.append(token)
    return tokens

def semantic_phrases(tokens, size):
    return {
        ' '.join(tokens[index:index + size])
        for index in range(0, max(len(tokens) - size + 1, 0))
    }

def action_categories(tokens):
    token_set = set(tokens)
    return {
        category
        for category, cues in ACTION_CUES.items()
        if token_set & cues
    }

def action_words(tokens):
    token_set = set(tokens)
    cues = set()
    for words in ACTION_CUES.values():
        cues.update(token_set & words)
    return cues

def clo_domain_from_number(clo):
    number = clo_number(clo)
    if number.startswith('1.'):
        return 'knowledge'
    if number.startswith('2.'):
        return 'skills'
    if number.startswith('3.'):
        return 'values'
    return ''

def score_question_clo_match(question_text, clo, explicit=False):
    q_tokens = semantic_tokens(question_text)
    clo_tokens = semantic_tokens(clo)
    if not q_tokens or not clo_tokens:
        return {'score': 0, 'reasons': []}

    q_set = set(q_tokens)
    c_set = set(clo_tokens)
    overlap = q_set & c_set
    union_size = max(len(q_set | c_set), 1)
    score = 0.0
    reasons = []

    if explicit:
        score += 70
        reasons.append('explicit CLO tag')

    if overlap:
        overlap_score = min(28, (len(overlap) / union_size) * 70)
        score += overlap_score
        reasons.append('shared concepts')

    q_bigrams = semantic_phrases(q_tokens, 2) | semantic_phrases(q_tokens, 3)
    c_bigrams = semantic_phrases(clo_tokens, 2) | semantic_phrases(clo_tokens, 3)
    phrase_overlap = q_bigrams & c_bigrams
    if phrase_overlap:
        score += min(18, len(phrase_overlap) * 6)
        reasons.append('matching concept phrase')

    q_actions = action_words(q_tokens)
    c_actions = action_words(clo_tokens)
    if q_actions and c_actions and (q_actions & c_actions):
        score += 16
        reasons.append('same action verb')

    q_categories = action_categories(q_tokens)
    c_categories = action_categories(clo_tokens)
    domain = clo_domain_from_number(clo)
    if q_categories & c_categories:
        score += 10
        reasons.append('same learning intent')
    if domain and domain in q_categories:
        score += 12
        reasons.append(f'{domain} domain cue')

    q_norm = normalize_semantic_text(question_text)
    c_norm = normalize_semantic_text(clo)
    similarity = difflib.SequenceMatcher(None, q_norm[:500], c_norm[:500]).ratio()
    if similarity > 0.22:
        score += min(10, similarity * 18)
        reasons.append('semantic text similarity')

    return {'score': round(min(score, 100), 2), 'reasons': reasons[:4]}

def add_question_clo_diagnostic(metrics, provider, status, message):
    metrics = dict(metrics or {})
    diagnostics = list(metrics.get('question_clo_diagnostics') or [])
    message = re.sub(r'\s+', ' ', str(message or '')).strip()
    entry = {
        'provider': str(provider or '').strip(),
        'status': str(status or '').strip(),
        'message': message[:700],
    }
    if entry['provider'] and entry['status'] and entry not in diagnostics:
        diagnostics.append(entry)
    metrics['question_clo_diagnostics'] = diagnostics
    return metrics

def elapsed_seconds(start_time):
    return round(max(time.perf_counter() - start_time, 0), 3)

def build_smart_clo_suggestions(metrics, clos, only_unmapped=False):
    metrics = metrics or {}
    clos = list(clos or [])
    question_texts = metrics.get('question_texts') or {}
    explicit_mappings = metrics.get('detected_clo_mappings') or {}
    suggestions = {}
    detected = dict(explicit_mappings)
    mapping_metadata = dict(metrics.get('question_clo_mapping_metadata') or {})
    attempted = 0

    for question in metrics.get('questions') or []:
        if only_unmapped and detected.get(question):
            continue
        attempted += 1
        question_start = time.perf_counter()
        question_text = question_texts.get(question) or question
        explicit_clos = set(resolve_detected_clos_to_course_list(explicit_mappings.get(question, []), clos))
        ranked = []
        for clo in clos:
            result = score_question_clo_match(question_text, clo, explicit=clo in explicit_clos)
            if result['score'] > 0:
                ranked.append({
                    'clo': clo,
                    'score': result['score'],
                    'reasons': result['reasons']
                })
        ranked.sort(key=lambda item: item['score'], reverse=True)
        if ranked:
            best_score = ranked[0]['score']
            chosen = [
                item['clo']
                for item in ranked[:3]
                if item['score'] >= 32 or item['score'] >= best_score - 7 or item['clo'] in explicit_clos
            ]
            if chosen:
                detected[question] = chosen
            suggestions[question] = ranked[:3]
            mapping_metadata[question] = {
                'source': 'local',
                'model': 'local-semantic',
                'duration_seconds': elapsed_seconds(question_start),
                'confidence': ranked[0].get('score'),
            }

    metrics['smart_clo_suggestions'] = suggestions
    metrics['detected_clo_mappings'] = detected
    metrics['question_clo_mapping_metadata'] = mapping_metadata
    if suggestions and not metrics.get('question_clo_suggestion_source'):
        metrics['question_clo_suggestion_source'] = 'local'
    if attempted:
        if suggestions:
            metrics = add_question_clo_diagnostic(
                metrics,
                'Local',
                'success',
                f"Local semantic matching produced suggestions for {len(suggestions)} question(s)."
            )
        else:
            metrics = add_question_clo_diagnostic(
                metrics,
                'Local',
                'failed',
                "Local semantic matching could not find enough overlap between the question text and the available CLOs."
            )
    return metrics


def gemini_exam_question_extraction_prompt():
    return (
        "Extract exam questions from the provided exam paper text. "
        "The exam may be Arabic or English. Preserve each question text as completely as possible, "
        "including sub-parts and options. IMPORTANT: For multiple choice questions, ensure each option appears on a new line within the question text (insert \\n before options if needed). "
        "Detect the question type, such as MCQ, True/False, Essay, Short Answer, Problem Solving, or Other. "
        "Detect only CLOs explicitly written in the exam paper near or inside each question, such as CLO1, CLO 1, "
        "CLO1.1, 1.1, or ناتج 1.1. Do not infer CLOs by meaning. "
        "Return JSON only with this exact schema: "
        '{"questions":[{"number":"Q1","text":"full question text with \\n separating options","type":"MCQ","explicit_clos":["1.1"]}]}'
    )

def normalize_gemini_exam_metrics(payload, source='gemini', confidence='Gemini'):
    if not isinstance(payload, dict):
        return {}
    raw_questions = payload.get('questions') or payload.get('items') or []
    if not isinstance(raw_questions, list):
        return {}
    questions = []
    question_texts = {}
    question_types = {}
    detected_clo_mappings = {}
    for index, item in enumerate(raw_questions, start=1):
        if not isinstance(item, dict):
            continue
        raw_number = compact_text(item.get('number') or item.get('id') or item.get('question') or '')
        number = question_number_from_label(raw_number) or index
        question_id = f"Q{number}"
        if question_id in questions:
            question_id = f"Q{len(questions) + 1}"
        text = compact_text(item.get('text') or item.get('question_text') or item.get('body') or '')
        if not text:
            continue
        questions.append(question_id)
        question_texts[question_id] = text[:4000]
        question_types[question_id] = compact_text(item.get('type') or item.get('question_type') or '')
        explicit = item.get('explicit_clos') or item.get('clos') or item.get('clo') or []
        if isinstance(explicit, str):
            explicit = [explicit]
        tags = detect_clo_tags_from_text(' '.join(str(value or '') for value in explicit))
        if not tags:
            tags = detect_clo_tags_from_text(text)
        if tags:
            detected_clo_mappings[question_id] = tags
    if not questions:
        return {}
    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': 0,
        'confidence': confidence,
        'text_sample': '',
        'question_texts': question_texts,
        'question_types': question_types,
        'detected_clo_mappings': detected_clo_mappings,
        'question_extraction_source': source,
    }

def extract_exam_text_for_ai(filepath):
    try:
        from exam_parser import ExamParser
        parser = ExamParser(filepath)
        parser.extract_text()
        return parser.raw_text or ''
    except Exception as exc:
        app.logger.warning("Could not extract exam text for Gemini: %s", exc)
        return ''

def parse_exam_paper_with_gemini(filepath):
    if not GEMINI_API_KEY:
        return {}
    text = extract_exam_text_for_ai(filepath)
    if not compact_text(text):
        return {}
    extraction_start = time.perf_counter()
    payload = {
        'exam_text': text[:60000],
    }
    parsed = call_gemini_json(
        gemini_exam_question_extraction_prompt(),
        json.dumps(payload, ensure_ascii=False),
        timeout=120
    )
    metrics = normalize_gemini_exam_metrics(parsed)
    if metrics:
        metrics['question_extraction_model'] = GEMINI_MODEL
        metrics['question_extraction_duration_seconds'] = elapsed_seconds(extraction_start)
        return metrics
    return {}

def parse_exam_paper_with_qwen(filepath):
    if not GROQ_KEY:
        return {}
    text = extract_exam_text_for_ai(filepath)
    if not compact_text(text):
        return {}
    extraction_start = time.perf_counter()
    parsed, error = call_groq_json_with_error(
        "You extract exam questions from exam papers. Return valid JSON only.",
        gemini_exam_question_extraction_prompt()
        + "\n\nInput JSON:\n"
        + json.dumps({'exam_text': text[:60000]}, ensure_ascii=False),
        timeout=120
    )
    if error:
        app.logger.warning("Qwen exam question extraction failed: %s", error)
        return {}
    metrics = normalize_gemini_exam_metrics(parsed, source='qwen', confidence='Qwen')
    if metrics:
        metrics['question_extraction_model'] = GROQ_MODEL
        metrics['question_extraction_duration_seconds'] = elapsed_seconds(extraction_start)
        return metrics
    return {}

def gemini_question_clo_prompt():
    return (
        "You are an academic assessment alignment assistant. "
        "Map exam questions to Course Learning Outcomes. "
        "Use question intent, required action, concepts, and CLO domain. "
        "The questions and CLOs may be Arabic or English. "
        "Return JSON only, with no markdown. Do not invent CLO codes. "
        "Use only the provided CLO codes. A question may map to more than one CLO when justified. "
        "Use this exact schema: "
        '{"mappings":[{"question":"Q1","clos":["1.1"],"confidence":0.85,"reason":"short reason"}]}'
    )


def apply_llm_question_clo_mappings(metrics, clos, parsed, source, reason_fallback, source_model=None, duration_seconds=None):
    metrics = dict(metrics or {})
    questions = list(metrics.get('questions') or [])
    mappings = parsed.get('mappings') if isinstance(parsed, dict) else []
    if not isinstance(mappings, list):
        return metrics

    detected = dict(metrics.get('detected_clo_mappings') or {})
    suggestions = dict(metrics.get('smart_clo_suggestions') or {})
    mapping_metadata = dict(metrics.get('question_clo_mapping_metadata') or {})
    mapped_count = 0
    question_lookup = {str(question): str(question) for question in questions}
    question_lookup.update({str(question).upper(): str(question) for question in questions})

    for item in mappings:
        if not isinstance(item, dict):
            continue
        question = str(item.get('question') or '').strip()
        question = question_lookup.get(question) or question_lookup.get(question.upper()) or question
        if question not in questions:
            number = question_number_from_label(question)
            if number:
                question = next((q for q in questions if question_number_from_label(q) == number), question)
        if question not in questions:
            continue
        resolved = resolve_detected_clos_to_course_list(item.get('clos') or [], clos)
        if not resolved:
            continue
        confidence = item.get('confidence', 0.85)
        try:
            confidence_score = float(confidence)
            if confidence_score <= 1:
                confidence_score *= 100
        except (TypeError, ValueError):
            confidence_score = 85.0
        reason = compact_text(item.get('reason') or reason_fallback)
        detected[question] = resolved
        suggestions[question] = [
            {
                'clo': clo,
                'score': round(max(0, min(confidence_score, 100)), 2),
                'reasons': [reason or reason_fallback]
            }
            for clo in resolved[:3]
        ]
        mapping_metadata[question] = {
            'source': source,
            'model': source_model or source,
            'duration_seconds': duration_seconds,
            'confidence': round(max(0, min(confidence_score, 100)), 2),
        }
        mapped_count += 1

    if mapped_count:
        metrics['detected_clo_mappings'] = detected
        metrics['smart_clo_suggestions'] = suggestions
        metrics['question_clo_mapping_metadata'] = mapping_metadata
        metrics['question_clo_suggestion_source'] = source
    return metrics


def question_clo_llm_payload(metrics, clos):
    questions = list((metrics or {}).get('questions') or [])
    question_texts = (metrics or {}).get('question_texts') or {}
    question_items = [
        {
            'question': question,
            'text': compact_text(question_texts.get(question) or question)[:1200]
        }
        for question in questions[:80]
    ]
    clo_items = [
        {'code': clo_number(clo), 'text': str(clo)}
        for clo in list(clos or [])
        if clo_number(clo)
    ]
    return question_items, clo_items


def build_qwen_question_clo_suggestions(metrics, clos):
    metrics = dict(metrics or {})
    if not GROQ_KEY:
        metrics = add_question_clo_diagnostic(metrics, 'Qwen', 'skipped', 'GROQ_KEY is not configured.')
        return metrics
    question_items, clo_items = question_clo_llm_payload(metrics, clos)
    if not question_items or not clo_items:
        metrics = add_question_clo_diagnostic(metrics, 'Qwen', 'skipped', 'No question text or valid CLO codes were available for Qwen.')
        return metrics
    provider_start = time.perf_counter()
    parsed, error = call_groq_json_with_error(
        "You map exam questions to Course Learning Outcomes. Return valid JSON only.",
        gemini_question_clo_prompt()
        + "\n\nInput JSON:\n"
        + json.dumps(
            {
                'course_learning_outcomes': clo_items,
                'questions': question_items
            },
            ensure_ascii=False
        )
    )
    duration = elapsed_seconds(provider_start)
    if error:
        return add_question_clo_diagnostic(metrics, 'Qwen', 'failed', f"{error} ({duration}s)")
    mapped_metrics = apply_llm_question_clo_mappings(metrics, clos, parsed, 'qwen', 'Qwen via Groq', GROQ_MODEL, duration)
    if mapped_metrics.get('question_clo_suggestion_source') == 'qwen':
        mapped_count = len(mapped_metrics.get('smart_clo_suggestions') or {})
        return add_question_clo_diagnostic(mapped_metrics, 'Qwen', 'success', f"Qwen produced valid CLO suggestions for {mapped_count} question(s) using {GROQ_MODEL} in {duration}s.")
    return add_question_clo_diagnostic(metrics, 'Qwen', 'failed', f"Qwen returned JSON, but no mappings matched the available course CLOs. ({duration}s)")


def build_gemini_question_clo_suggestions(metrics, clos):
    metrics = dict(metrics or {})
    if not GEMINI_API_KEY:
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'skipped', 'GEMINI_API_KEY is not configured.')
        return build_qwen_question_clo_suggestions(metrics, clos)
    questions = list(metrics.get('questions') or [])
    clos = list(clos or [])
    if not questions or not clos:
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'skipped', 'No questions or CLOs were available for Gemini mapping.')
        return metrics

    question_items, clo_items = question_clo_llm_payload(metrics, clos)
    if not question_items or not clo_items:
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'skipped', 'No question text or valid CLO codes were available for Gemini.')
        return metrics

    provider_start = time.perf_counter()
    try:
        payload = {
            'contents': [
                {
                    'parts': [
                        {'text': gemini_question_clo_prompt()},
                        {
                            'text': json.dumps(
                                {
                                    'course_learning_outcomes': clo_items,
                                    'questions': question_items
                                },
                                ensure_ascii=False
                            )
                        }
                    ]
                }
            ],
            'generationConfig': {
                'temperature': 0,
                'responseMimeType': 'application/json'
            }
        }
        endpoint = (
            'https://generativelanguage.googleapis.com/v1beta/models/'
            + urllib.parse.quote(GEMINI_MODEL, safe='')
            + ':generateContent?key='
            + urllib.parse.quote(GEMINI_API_KEY, safe='')
        )
        gemini_request = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode('utf-8'),
            headers={
                'Content-Type': 'application/json',
                'User-Agent': 'ETQAN-QuestionCLOMapper/1.0'
            },
            method='POST'
        )
        with urllib.request.urlopen(gemini_request, timeout=90) as response:
            gemini_payload = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        duration = elapsed_seconds(provider_start)
        body = exc.read().decode('utf-8', errors='replace')
        app.logger.warning("Gemini question-CLO mapping failed with HTTP %s: %s", exc.code, body[:800])
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'failed', f"HTTP {exc.code}: {body[:500]} ({duration}s)")
        return build_qwen_question_clo_suggestions(metrics, clos)
    except Exception as exc:
        duration = elapsed_seconds(provider_start)
        app.logger.warning("Gemini question-CLO mapping failed: %s", exc)
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'failed', f"{exc} ({duration}s)")
        return build_qwen_question_clo_suggestions(metrics, clos)
    duration = elapsed_seconds(provider_start)

    response_text = ''
    for candidate in gemini_payload.get('candidates') or []:
        content = candidate.get('content') or {}
        for part in content.get('parts') or []:
            response_text += part.get('text') or ''

    parsed = parse_gemini_json_response(response_text)
    if not isinstance(parsed, dict):
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'failed', f"Gemini returned a response, but it was not valid JSON. ({duration}s)")
        return build_qwen_question_clo_suggestions(metrics, clos)
    mappings = parsed.get('mappings')
    if not isinstance(mappings, list):
        metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'failed', f"Gemini returned a response, but it did not contain a valid mappings list. ({duration}s)")
        return build_qwen_question_clo_suggestions(metrics, clos)
    mapped_metrics = apply_llm_question_clo_mappings(metrics, clos, parsed, 'gemini', 'Gemini Flash', GEMINI_MODEL, duration)
    if mapped_metrics.get('question_clo_suggestion_source') == 'gemini':
        mapped_count = len(mapped_metrics.get('smart_clo_suggestions') or {})
        mapped_metrics = add_question_clo_diagnostic(mapped_metrics, 'Gemini', 'success', f"Gemini produced valid CLO suggestions for {mapped_count} question(s) using {GEMINI_MODEL} in {duration}s.")
        return mapped_metrics
    metrics = add_question_clo_diagnostic(metrics, 'Gemini', 'failed', f"Gemini returned JSON, but no mappings matched the available course CLOs. ({duration}s)")
    return build_qwen_question_clo_suggestions(metrics, clos)

def parse_exam_paper_with_module(filepath):
    extraction_start = time.perf_counter()
    try:
        from exam_parser import ExamParser
        parser = ExamParser(filepath)
        parsed_questions = parser.parse()
        
        questions = []
        question_texts = {}
        question_types = {}
        detected_clo_mappings = {}
        
        for q in parsed_questions:
            q_id = q['question_id']
            questions.append(q_id)
            question_types[q_id] = str(q.get('question_type') or '').strip()
            full_text = str(q.get('question_text') or '').strip()
            if q.get('marks', 1.0) != 1.0:
                full_text += f"\nMarks: {q['marks']}"
                
            question_texts[q_id] = full_text[:2500] + ('...' if len(full_text) > 2500 else '')
            tags = detect_clo_tags_from_text(q['question_text'])
            if tags:
                detected_clo_mappings[q_id] = tags
                
        return {
            'questions': questions,
            'total_questions': len(questions),
            'total_students': 0,
            'confidence': 'High' if questions else 'Low',
            'text_sample': (parser.raw_text[:150] + '...') if hasattr(parser, 'raw_text') else '',
            'question_texts': question_texts,
            'question_types': question_types,
            'detected_clo_mappings': detected_clo_mappings,
            'question_extraction_source': 'local',
            'question_extraction_model': 'ExamParser',
            'question_extraction_duration_seconds': elapsed_seconds(extraction_start)
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise e

def parse_exam_paper_metrics(filepath):
    gemini_metrics = parse_exam_paper_with_gemini(filepath)
    if gemini_metrics:
        return gemini_metrics
    qwen_metrics = parse_exam_paper_with_qwen(filepath)
    if qwen_metrics:
        return qwen_metrics
    return parse_exam_paper_with_module(filepath)

def infer_exam_paper_metrics(text):
    # Deprecated fallback
    return {'questions': [], 'total_questions': 0, 'confidence': 'Low', 'question_texts': {}, 'detected_clo_mappings': {}}

def question_mapping_draft_path(draft_id):
    safe_id = re.sub(r'[^A-Za-z0-9_-]', '', str(draft_id or ''))
    if not safe_id:
        raise ValueError("Invalid question mapping draft.")
    return get_upload_path(f"question_mapping_{safe_id}.json")

def save_question_mapping_draft(payload):
    draft_id = str(uuid.uuid4())
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False)
    return draft_id

def load_question_mapping_draft(draft_id):
    path = question_mapping_draft_path(draft_id)
    if not os.path.exists(path):
        raise FileNotFoundError("Question mapping draft was not found.")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def build_question_mapping_metrics_from_texts(question_texts, question_types=None):
    questions = []
    normalized_texts = {}
    normalized_types = {}
    detected_clo_mappings = {}
    question_types = question_types or []
    for index, text in enumerate(question_texts, start=1):
        cleaned = re.sub(r'\s+', ' ', str(text or '')).strip()
        if not cleaned:
            continue
        question_id = f"Q{len(questions) + 1}"
        questions.append(question_id)
        normalized_texts[question_id] = cleaned
        if index - 1 < len(question_types):
            normalized_types[question_id] = re.sub(r'\s+', ' ', str(question_types[index - 1] or '')).strip()
        else:
            normalized_types[question_id] = ''
        tags = detect_clo_tags_from_text(cleaned)
        if tags:
            detected_clo_mappings[question_id] = tags
    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': 0,
        'confidence': 'Edited',
        'text_sample': '',
        'question_texts': normalized_texts,
        'question_types': normalized_types,
        'detected_clo_mappings': detected_clo_mappings
    }

def build_question_review_metrics_from_form(clos):
    question_texts = request.form.getlist('question_text')
    question_types = request.form.getlist('question_type')
    metrics = build_question_mapping_metrics_from_texts(question_texts, question_types)
    paper_detected_mappings = {}
    for index, question in enumerate(metrics.get('questions') or []):
        selected_from_paper = request.form.getlist(f'paper_clo_{index}')
        resolved_from_paper = resolve_detected_clos_to_course_list(selected_from_paper, clos)
        if resolved_from_paper:
            paper_detected_mappings[question] = resolved_from_paper
    if paper_detected_mappings:
        detected = dict(metrics.get('detected_clo_mappings') or {})
        detected.update(paper_detected_mappings)
        metrics['detected_clo_mappings'] = detected
    return metrics, paper_detected_mappings

def build_question_review_summary(metrics, clos):
    metrics = metrics or {}
    explicit_by_question = {}
    needs_ai_questions = []
    for question in metrics.get('questions') or []:
        resolved = resolve_detected_clos_to_course_list(
            (metrics.get('detected_clo_mappings') or {}).get(question, []),
            clos
        )
        explicit_by_question[question] = resolved
        if not resolved:
            needs_ai_questions.append(question)

    total = len(metrics.get('questions') or [])
    needs_ai = len(needs_ai_questions)
    return {
        'total': total,
        'auto_mapped': max(total - needs_ai, 0),
        'needs_ai': needs_ai,
        'all_mapped': bool(total) and needs_ai == 0,
        'needs_ai_questions': needs_ai_questions,
        'explicit_by_question': explicit_by_question,
    }

def build_question_final_metrics_from_form(clos):
    question_ids = [str(item or '').strip() for item in request.form.getlist('question_ids')]
    question_ids = [item for item in question_ids if item]
    if not question_ids:
        return build_question_review_metrics_from_form(clos)[0]

    metrics = {
        'questions': [],
        'total_questions': 0,
        'total_students': 0,
        'confidence': 'Edited',
        'text_sample': '',
        'question_texts': {},
        'question_types': {},
        'detected_clo_mappings': {},
        'ai_suggested_clos': {},
    }
    seen = set()
    for raw_question in question_ids:
        question = raw_question
        if question in seen:
            continue
        seen.add(question)
        metrics['questions'].append(question)
        metrics['question_texts'][question] = re.sub(
            r'\s+',
            ' ',
            str(request.form.get(f'question_text_{question}') or '').strip()
        )
        metrics['question_types'][question] = re.sub(
            r'\s+',
            ' ',
            str(request.form.get(f'question_type_{question}') or '').strip()
        )
        selected_clos = request.form.getlist(f'question_clo_{question}')
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
            metrics['detected_clo_mappings'][question] = resolved
            
        ai_suggested = request.form.get(f'ai_suggested_clo_{question}')
        if ai_suggested:
            metrics['ai_suggested_clos'][question] = ai_suggested
            
    metrics['total_questions'] = len(metrics['questions'])
    return metrics

def merge_ai_default_clo_selections_from_form(metrics, clos):
    metrics = dict(metrics or {})
    mappings = dict(metrics.get('detected_clo_mappings') or {})
    removed = dict(metrics.get('ai_removed_clos') or {})
    suggested = dict(metrics.get('ai_suggested_clos') or {})
    for question in metrics.get('questions') or []:
        default_clo = (request.form.get(f'ai_default_clo_{question}') or '').strip()
        if not default_clo:
            continue
        suggested.setdefault(question, default_clo)
        if request.form.get(f'ai_removed_clo_{question}') == '1':
            removed[question] = default_clo
            continue
        resolved_default = resolve_detected_clos_to_course_list([default_clo], clos)
        if not resolved_default:
            resolved_default = [default_clo]
        current = list(mappings.get(question) or [])
        for clo in resolved_default:
            if clo not in current:
                current.append(clo)
        mappings[question] = current
    metrics['detected_clo_mappings'] = mappings
    metrics['ai_suggested_clos'] = suggested
    if removed:
        metrics['ai_removed_clos'] = removed
    return metrics

def question_mapping_values_for_key(mapping, question):
    if not isinstance(mapping, dict):
        return []
    question_text = re.sub(r'\s+', ' ', str(question or '').strip()).lower()
    question_number = question_number_from_label(question)

    candidates = []
    if question in mapping:
        candidates.append(mapping.get(question))

    for key, value in mapping.items():
        key_text = re.sub(r'\s+', ' ', str(key or '').strip()).lower()
        if key_text and key_text == question_text:
            candidates.append(value)
            continue
        if question_number and question_number_from_label(key) == question_number:
            candidates.append(value)

    values = []
    for candidate in candidates:
        if candidate is None:
            continue
        if isinstance(candidate, (list, tuple, set)):
            iterable = candidate
        else:
            iterable = [candidate]
        for item in iterable:
            item = str(item or '').strip()
            if item and item not in values:
                values.append(item)
    return values

def ensure_final_review_clo_selections(metrics, draft_metrics, clos):
    metrics = dict(metrics or {})
    draft_metrics = dict(draft_metrics or {})
    mappings = {
        question: list(values or [])
        for question, values in (metrics.get('detected_clo_mappings') or {}).items()
    }
    suggested = dict(draft_metrics.get('ai_suggested_clos') or {})
    suggested.update(metrics.get('ai_suggested_clos') or {})
    removed = dict(draft_metrics.get('ai_removed_clos') or {})
    removed.update(metrics.get('ai_removed_clos') or {})
    draft_selections = draft_metrics.get('ai_draft_clo_selections') or {}
    draft_detected = draft_metrics.get('detected_clo_mappings') or {}

    for question in metrics.get('questions') or []:
        current = list(resolve_detected_clos_to_course_list(question_mapping_values_for_key(mappings, question), clos))
        fallback_values = []
        fallback_values.extend(question_mapping_values_for_key(draft_selections, question))
        fallback_values.extend(question_mapping_values_for_key(draft_detected, question))
        removed_values = question_mapping_values_for_key(removed, question)
        suggested_values = question_mapping_values_for_key(suggested, question)
        if not removed_values:
            fallback_values.extend(suggested_values)
        resolved_fallbacks = resolve_detected_clos_to_course_list(fallback_values, clos)
        if not resolved_fallbacks:
            resolved_fallbacks = [str(c).strip() for c in fallback_values if str(c).strip()]
        for clo in resolved_fallbacks:
            if clo not in current:
                current.append(clo)
        if current:
            mappings[question] = current

    metrics['detected_clo_mappings'] = mappings
    if suggested:
        metrics['ai_suggested_clos'] = suggested
    if removed:
        metrics['ai_removed_clos'] = removed
    return metrics

def build_ai_suggestions_for_unmapped(metrics, clos, review_summary):
    metrics = dict(metrics or {})
    unmapped_questions = list((review_summary or {}).get('needs_ai_questions') or [])
    if not unmapped_questions:
        return metrics

    ai_metrics = dict(metrics)
    ai_metrics['questions'] = unmapped_questions
    ai_metrics['detected_clo_mappings'] = {}
    ai_metrics['smart_clo_suggestions'] = {}
    ai_metrics.pop('question_clo_suggestion_source', None)

    ai_metrics = build_gemini_question_clo_suggestions(ai_metrics, clos)
    ai_metrics = build_smart_clo_suggestions(
        ai_metrics,
        clos,
        only_unmapped=ai_metrics.get('question_clo_suggestion_source') == 'gemini'
    )

    merged = dict(metrics)
    merged['smart_clo_suggestions'] = dict(ai_metrics.get('smart_clo_suggestions') or {})
    merged['question_clo_mapping_metadata'] = dict(ai_metrics.get('question_clo_mapping_metadata') or {})
    merged['question_clo_diagnostics'] = list(ai_metrics.get('question_clo_diagnostics') or [])
    if ai_metrics.get('question_clo_suggestion_source'):
        merged['question_clo_suggestion_source'] = ai_metrics.get('question_clo_suggestion_source')
    elif merged['smart_clo_suggestions']:
        merged['question_clo_suggestion_source'] = 'local'
    return merged

def apply_exam_paper_mappings(report_metrics, exam_metrics):
    if not exam_metrics:
        return report_metrics
    detected = dict(report_metrics.get('detected_clo_mappings') or {})
    exam_detected = exam_metrics.get('detected_clo_mappings') or {}
    question_texts = dict(report_metrics.get('question_texts') or {})
    exam_question_texts = exam_metrics.get('question_texts') or {}
    by_number = {}
    text_by_number = {}
    for question, tags in exam_detected.items():
        question_number = question_number_from_label(question)
        if question_number:
            by_number[question_number] = tags
    for question, text in exam_question_texts.items():
        question_number = question_number_from_label(question)
        if question_number:
            text_by_number[question_number] = text

    for question in report_metrics.get('questions') or []:
        question_number = question_number_from_label(question)
        if question_number and by_number.get(question_number):
            detected.setdefault(question, by_number[question_number])
        if question_number and text_by_number.get(question_number):
            question_texts[question] = text_by_number[question_number]

    report_metrics['detected_clo_mappings'] = detected
    report_metrics['question_texts'] = question_texts
    if exam_metrics.get('questions'):
        note = f" Exam paper detected {len(exam_metrics['questions'])} question(s)"
        mapped_count = len(exam_detected)
        if mapped_count:
            note += f" and CLO tags for {mapped_count} question(s)."
        else:
            note += "."
        report_metrics['text_sample'] = f"{report_metrics.get('text_sample', '')} {note}".strip()
    return report_metrics

def compact_text(text):
    return re.sub(r'\s+', ' ', text or '').strip()

METADATA_VALUE_STOP_LABELS = [
    'Course Name', 'Course Title', 'Course Code', 'Course Number', 'Course ID', 'Course No',
    'Program', 'Programme', 'Program Name', 'Academic Program',
    'Department', 'Dept', 'Academic Department', 'Scientific Department',
    'College', 'Faculty', 'School', 'Institution', 'University',
    'Version', 'Last Revision Date', 'Revision Date', 'Date',
    'Credit hours', 'Course type', 'Level/year'
]

def trim_metadata_value(value, current_labels=None):
    value = compact_text(value)
    current = {str(label).casefold() for label in (current_labels or [])}
    stop_labels = [
        label for label in METADATA_VALUE_STOP_LABELS
        if label.casefold() not in current
    ]
    if stop_labels:
        stop_pattern = '|'.join(re.escape(label) for label in stop_labels)
        value = re.split(rf'\s+(?:{stop_pattern})\s*:', value, maxsplit=1, flags=re.I)[0]
    return clean_pdf_fragment(value).strip()

ARABIC_DIGIT_TRANSLATION = str.maketrans({
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
    '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    '٫': '.', '۔': '.', '،': ','
})

def normalize_course_spec_text(text):
    text = unicodedata.normalize('NFKC', text or '').translate(ARABIC_DIGIT_TRANSLATION)
    text = text.replace('\u0640', '')
    text = re.sub(r'[\u200e\u200f\u061c\u202a-\u202e]', '', text)
    text = text.replace('\xa0', ' ')
    return text

def clean_report_pdf_text(value):
    text = str(value or '')
    text = re.sub(r'[\u200b-\u200f\u061c\u202a-\u202e\ufeff\ufffc]', '', text)
    text = text.replace('\ufffd', '').replace('\xa0', ' ')
    return text

def contains_arabic(text):
    return bool(re.search(r'[\u0600-\u06FF]', text or ''))

def arabic_value_looks_fragmented(value):
    source = str(value or '')
    normalized = normalize_course_spec_text(source)
    if not contains_arabic(normalized):
        return False
    if re.search(r'[_\u0640]{3,}', source):
        return True
    words = re.findall(r'[\u0600-\u06FF]+', normalized)
    if len(words) < 6:
        return False
    short_ratio = sum(1 for word in words if len(word) <= 2) / len(words)
    single_ratio = sum(1 for word in words if len(word) == 1) / len(words)
    return short_ratio > 0.58 or single_ratio > 0.24

def arabic_text_layer_looks_fragmented(text):
    normalized = normalize_course_spec_text(text or '')
    if not contains_arabic(normalized):
        return False
    words = re.findall(r'[\u0600-\u06FF]+', normalized)
    if len(words) < 25:
        return arabic_value_looks_fragmented(text)
    short_ratio = sum(1 for word in words if len(word) <= 2) / len(words)
    single_ratio = sum(1 for word in words if len(word) == 1) / len(words)
    isolated_runs = len(re.findall(r'(?:^|\s)[\u0600-\u06FF](?:\s+[\u0600-\u06FF]){3,}(?=\s|$)', normalized))
    return bool(re.search(r'[_\u0640]{3,}', str(text or ''))) or short_ratio > 0.52 or single_ratio > 0.18 or isolated_runs >= 2

def course_spec_extracted_arabic_looks_fragmented(extracted):
    if not isinstance(extracted, dict):
        return False
    values = []
    values.extend(extracted.get('clos') or [])
    values.extend(extracted.get('topics') or [])
    values.extend([
        extracted.get('course_name') or '',
        extracted.get('college') or '',
        extracted.get('department') or '',
        extracted.get('program') or '',
    ])
    arabic_values = [value for value in values if contains_arabic(str(value or ''))]
    if not arabic_values:
        return False
    return any(arabic_value_looks_fragmented(value) for value in arabic_values)

def text_direction(text):
    return 'rtl' if contains_arabic(str(text or '')) else 'ltr'

def extract_first_int(patterns, text):
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                continue
    return None

def infer_course_report_metrics(text):
    normalized = compact_text(text)
    question_numbers = set()

    question_patterns = [
        r'\bQ(?:uestion)?\s*[-#:]?\s*(\d{1,3})\b',
        r'\bQuestion\s+No\.?\s*(\d{1,3})\b',
        r'\bItem\s*[-#:]?\s*(\d{1,3})\b'
    ]
    for pattern in question_patterns:
        for match in re.finditer(pattern, normalized, flags=re.I):
            number = int(match.group(1))
            if 0 < number <= 200:
                question_numbers.add(number)

    total_questions = extract_first_int([
        r'(?:number|no\.?|total)\s+of\s+questions?\D{0,30}(\d{1,3})',
        r'questions?\s*(?:count|total|number)?\s*[:=]\s*(\d{1,3})',
        r'(\d{1,3})\s+questions?\b'
    ], normalized)

    if question_numbers:
        total_questions = max(total_questions or 0, max(question_numbers))

    total_students = extract_first_int([
        r'(?:number|no\.?|total)\s+of\s+students?\D{0,30}(\d{1,4})',
        r'students?\s*(?:count|total|number)?\s*[:=]\s*(\d{1,4})',
        r'(\d{1,4})\s+students?\b'
    ], normalized)

    questions = []
    if total_questions:
        questions = [f'Q{i}' for i in range(1, total_questions + 1)]
    elif question_numbers:
        questions = [f'Q{i}' for i in sorted(question_numbers)]

    confidence = 'High' if questions and total_students else 'Medium' if questions or total_students else 'Low'
    return {
        'questions': questions,
        'total_questions': len(questions) if questions else (total_questions or 0),
        'total_students': total_students or 0,
        'confidence': confidence,
        'text_sample': normalized[:1200]
    }

def is_metadata_label_line(line):
    labels = [
        'Course Name', 'Course Title', 'Course', 'Course Code', 'Course Number',
        'Course ID', 'Course No', 'College', 'Faculty', 'School', 'Department',
        'Dept', 'Academic Department', 'Scientific Department', 'Program',
        'Programme', 'Program Name', 'Academic Program', 'Institution',
        'Version', 'Last Revision Date'
    ]
    return any(
        re.search(rf'(?i)^\s*(?:\d+[\.\)]\s*)?{re.escape(label)}\b\s*[:\-]+', line or '')
        for label in labels
    )


def collect_metadata_continuation(lines, start_index, first_value, current_labels):
    parts = [first_value]
    open_parens = str(first_value).count('(') > str(first_value).count(')')
    for offset in range(start_index + 1, min(len(lines), start_index + 5)):
        line = compact_text(lines[offset])
        if not line:
            continue
        if is_metadata_label_line(line):
            break
        if re.match(r'^\d+$', line) or re.match(r'^[A-Z]\.\s+', line):
            break
        if re.search(r'(?i)table of contents|general information|course learning outcomes', line):
            break
        if open_parens or str(parts[-1]).rstrip().lower().endswith(('/', ',', 'and', 'of')):
            parts.append(line)
            open_parens = ' '.join(parts).count('(') > ' '.join(parts).count(')')
        else:
            break
    return clean_pdf_fragment(' '.join(parts))


def value_after_label(lines, labels):
    for index, line in enumerate(lines):
        table_value = value_after_table_label_line(line, labels)
        if table_value:
            value = trim_metadata_value(table_value, labels)
            if value:
                return collect_metadata_continuation(lines, index, value, labels)
        for label in labels:
            pattern = rf'(?i)^\s*(?:\d+[\.\)]\s*)?{re.escape(label)}\b\s*[:\-]+\s*(.+)$'
            match = re.search(pattern, line)
            if match and match.group(1).strip():
                value = trim_metadata_value(match.group(1), labels)
                if len(value) > 1:
                    return collect_metadata_continuation(lines, index, value, labels)
            if re.search(rf'(?i)^\s*(?:\d+[\.\)]\s*)?{re.escape(label)}\b\s*$', line) and index + 1 < len(lines):
                next_value = trim_metadata_value(lines[index + 1], labels)
                if next_value:
                    return collect_metadata_continuation(lines, index + 1, next_value, labels)
    return ''

def value_after_label_loose(lines, labels):
    stop_label_words = [
        'اسم المقرر', 'عنوان المقرر', 'رمز المقرر', 'رقم المقرر', 'كود المقرر',
        'البرنامج', 'القسم العلمي', 'الكلية', 'المؤسسة', 'تاريخ آخر مراجعة',
        'Course Name', 'Course Title', 'Course Code', 'Course Number', 'Course ID'
    ]
    sorted_labels = sorted(labels, key=len, reverse=True)
    for index, line in enumerate(lines):
        clean_line = compact_text(line)
        for label in sorted_labels:
            label_index = clean_line.find(label)
            if label_index < 0:
                continue
            if label_index > 0 and re.match(r'[\u0600-\u06FFA-Za-z]', clean_line[label_index - 1]):
                continue
            after_raw = clean_line[label_index + len(label):]
            if after_raw and not re.match(r'^\s*[:\-–—؛،]\s*', after_raw):
                continue
            after = after_raw.strip(' :-–—؛،\t')
            after = re.sub(r'^(?:اسم|رمز|رقم|كود)?\s*[:\-–—؛،]?\s*', '', after).strip()
            if after and after != clean_line:
                if any(after.startswith(stop_label) for stop_label in stop_label_words):
                    continue
                return after
            if index + 1 < len(lines):
                next_value = compact_text(lines[index + 1])
                if next_value and not any(next_value.startswith(stop_label) for stop_label in stop_label_words):
                    return next_value
    return ''


def flexible_arabic_label_pattern(label):
    parts = [re.escape(char) for char in str(label or '').replace(' ', '')]
    return r'\s*'.join(parts)


def value_after_table_label_line(line, labels):
    if '|' not in str(line or ''):
        return ''
    cells = [compact_text(cell).strip(' :-\u061b\u060c') for cell in str(line).split('|')]
    cells = [cell for cell in cells if cell]
    if len(cells) < 2:
        return ''
    for index, cell in enumerate(cells):
        for label in labels:
            label_text = str(label or '').strip()
            if not label_text:
                continue
            english_match = re.fullmatch(rf'(?i)(?:\d+[\.\)]\s*)?{re.escape(label_text)}\s*:?', cell)
            arabic_match = bool(arabic_label_key(label_text)) and arabic_label_key(cell).startswith(arabic_label_key(label_text))
            if not english_match and not arabic_match:
                continue
            for candidate in cells[index + 1:]:
                if not candidate:
                    continue
                is_label = any(
                    re.fullmatch(rf'(?i)(?:\d+[\.\)]\s*)?{re.escape(str(other or '').strip())}\s*:?', candidate)
                    or (arabic_label_key(other) and arabic_label_key(candidate).startswith(arabic_label_key(other)))
                    for other in labels
                )
                if not is_label:
                    return candidate
    return ''


def value_after_arabic_label(lines, labels, max_lines=80):
    stop_labels = [
        'اسم المقرر', 'اسمالمقرر', 'عنوان المقرر', 'رمز المقرر', 'رمزالمقرر',
        'رقم المقرر', 'كود المقرر', 'البرنامج', 'القسم العلمي', 'القسم',
        'الكلية', 'المؤسسة', 'المحتويات', 'التعريف بالمقرر'
    ]
    for line in (lines or [])[:max_lines]:
        clean_line = compact_text(line)
        table_value = value_after_table_label_line(clean_line, labels)
        if table_value:
            return table_value
        for label in labels:
            label_pattern = flexible_arabic_label_pattern(label)
            match = re.match(
                rf'^\s*{label_pattern}\s*[:\-–—؛،]?\s*(.+?)\s*$',
                clean_line
            )
            if not match:
                continue
            value = match.group(1).strip(' :-–—؛،\t')
            if not value:
                continue
            if any(arabic_label_key(value).startswith(arabic_label_key(stop)) for stop in stop_labels):
                continue
            return value
    return ''


def clean_arabic_metadata_value(value, field=''):
    value = clean_pdf_fragment(value)
    value = re.sub(r'\s+', ' ', value or '').strip(' :-–—؛،')
    value = value.strip(' |')
    if field == 'course_name':
        value = re.sub(r'^\s*توصيف\s+مقرر\s+', '', value).strip()
        value = re.sub(r'^\s*مقرر\s+', '', value).strip()
    if field == 'program' and value in {'الطلبة', 'الطالب', 'الطلاب'}:
        return ''
    return value

def normalize_extracted_course_code(value):
    code = re.sub(r'[\s\]\[(){}:؛،]+', '', str(value or '')).upper()
    code = code.replace('|', '')
    return code

def prefer_alphanumeric_course_code(text, current_code=''):
    current = normalize_extracted_course_code(current_code)
    if current and re.search(r'[A-Z]', current):
        return current
    for match in re.finditer(r'\b([A-Z]{2,6}[A-Z0-9]?\d{3,4}[A-Z]?)\b', text or '', flags=re.I):
        candidate = normalize_extracted_course_code(match.group(1))
        if re.search(r'[A-Z]', candidate) and re.search(r'\d', candidate):
            return candidate
    return current

def clean_clo_text(value):
    value = re.sub(r'\s+', ' ', value or '').strip()
    value = re.sub(r'\s+(Teaching\s+Strategies|Assessment\s+Methods|Code\s+of\s+PLOs|Domain)\b.*$', '', value, flags=re.I).strip()
    value = value.replace('م لك ومناقشة', 'ومناقشة')
    value = value.replace('الي توصل م إلها', 'التي توصل إليها')
    value = value.replace('إلها', 'إليها')
    value = re.sub(r'نوازل\s+التفكير\s+الناقد\.?\s+العبادات', 'نوازل العبادات', value)
    value = re.sub(r'\bdat\s+a\b|\bda\s+ta\b', 'data', value, flags=re.I)
    value = re.sub(r'\bsucg\b', 'such', value, flags=re.I)
    value = re.sub(r'\bevalute\b', 'evaluate', value, flags=re.I)
    value = re.sub(r'\s+-\s+', '-', value)
    value = re.sub(r'\band\s+along\s+with\b', 'along with', value, flags=re.I)
    value = re.sub(r'\bdesign\s+create\b', 'Design/Create', value, flags=re.I)
    value = re.sub(r'\bai\b', 'AI', value, flags=re.I)
    value = re.sub(r'\s+[عقم]\s*\d+\s*$', '', value)
    value = re.sub(
        r'\s+(?:المحاضرات|المناقشات|الاختبارات|التحريرية|الشفوية|التعلم\s+التعاوني|'
        r'التعلم\s+الذاتي|تعلم\s+الأقران|العصف\s+الذهني|الطريقة\s+الاستقرائية|'
        r'الملاحظة|التقييم|التقويم|تكليف|البحوث|الواجبات)\b.*$',
        '',
        value
    ).strip()
    value = re.sub(r'\s+([,.;:])', r'\1', value)
    if value:
        value = value[0].upper() + value[1:]
    return value


def extract_legacy_course_goal_clos(raw_text):
    normalized = compact_text(str(raw_text or '').replace('|', ' '))
    goal_label = r'(?:هدف\s+المقرر|أهداف\s+المقرر|الهدف\s+الرئيس\s+للمقرر|الهدف\s+الرئيسي\s+للمقرر)'
    stop_label = (
        r'(?:موضوعات\s+المقرر|قائمة\s+الموضوعات|محتوى\s+المقرر|نواتج\s+التعلم|'
        r'المراجع|مصادر\s+التعلم|Course\s+Learning\s+Outcomes|Course\s+Content|References)'
    )
    match = re.search(rf'{goal_label}\s*[:\-–—؛،]?\s*(.+?)(?=\s+{stop_label}\b|$)', normalized, flags=re.I | re.S)
    if not match:
        return {}
    goal_text = clean_clo_text(clean_pdf_fragment(match.group(1)))
    if len(goal_text) < 12:
        return {}

    numbered = []
    for item in re.finditer(r'(?:^|\s)(\d{1,2})\s+(.+?)(?=\s+\d{1,2}\s+|$)', goal_text, flags=re.S):
        body = clean_clo_text(item.group(2))
        if len(body) >= 10:
            numbered.append(body)
    if len(numbered) >= 2 and re.match(r'^\s*1\s+', goal_text):
        return {f'1.{index}': f'1.{index} {body}' for index, body in enumerate(numbered, start=1)}

    return {'1.1': f'1.1 {goal_text}'}


def flexible_label(label):
    parts = []
    for char in label:
        if char.isspace():
            parts.append(r'\s+')
        else:
            parts.append(re.escape(char) + r'\s*')
    return ''.join(parts)

COURSE_SPEC_WORDS = {
    'a', 'an', 'and', 'application', 'apply', 'appropriate', 'algorithms', 'as', 'assignment',
    'assignments', 'basic', 'big', 'brainstorming', 'class', 'collaborate', 'course', 'data', 'deletion',
    'develop', 'different', 'discussion', 'ended', 'evaluate', 'exam', 'exams', 'given', 'homework',
    'identify', 'implementation', 'in', 'insertion', 'knowledge', 'labs', 'learning', 'lectures',
    'of', 'on', 'open', 'outcomes', 'problem', 'problems', 'program', 'programming', 'quizzes',
    'relation', 'rely', 'require', 'requires', 'responsibility', 'searching', 'skills', 'solve',
    'solving', 'sorting', 'strategies', 'strengths', 'structures', 'such', 'teams', 'that', 'the',
    'to', 'types', 'understanding', 'values', 'weaknesses', 'with',
    'abilities', 'about', 'addressing', 'AI', 'ai', 'analyse', 'analytical', 'applications',
    'assess', 'assessments', 'capability', 'complex', 'computational', 'computer',
    'considerations', 'critical', 'datasets', 'decision', 'developments', 'discuss', 'driven',
    'ethical', 'evaluating', 'examine', 'expertise', 'fields', 'for', 'frameworks', 'from', 'ideas',
    'improve', 'including', 'influence', 'issues', 'machine', 'making', 'manage', 'management', 'methods',
    'multiple', 'new', 'practical', 'predictive', 'privacy', 'procedures', 'purposes',
    'reasoning', 'recent', 'recognize', 'robust', 'science', 'sophisticated', 'special',
    'specialized', 'statistical', 'statistics', 'tackle', 'tech', 'techniques', 'through',
    'tools', 'topics', 'trends', 'understand', 'utilize',
    'align', 'alignment', 'along', 'another', 'choice', 'classify', 'compare', 'concepts',
    'create', 'demonstrate', 'design', 'effectively', 'explain', 'justify', 'levels',
    'managing', 'maturity', 'metadata', 'one', 'open', 'practices', 'proficiency',
    'over', 'quality', 'reuse', 'sharing', 'solution', 'solutions', 'specify', 'standards', 'strategy', 'variety'
}

COURSE_SPEC_ALIASES = {
    'evalute': 'evaluate',
    'sucg': 'such'
}

def segment_compact_words(value):
    compact = re.sub(r'[^A-Za-z]', '', value or '').lower()
    if not compact:
        return ''

    max_word_length = 18
    dp = [None] * (len(compact) + 1)
    dp[0] = (0, [])
    for start in range(len(compact)):
        if dp[start] is None:
            continue
        for end in range(start + 1, min(len(compact), start + max_word_length) + 1):
            raw_word = compact[start:end]
            word = COURSE_SPEC_ALIASES.get(raw_word, raw_word)
            if word not in COURSE_SPEC_WORDS:
                continue
            score = dp[start][0] + len(raw_word) ** 2
            if dp[end] is None or score > dp[end][0]:
                dp[end] = (score, dp[start][1] + [word])

    if dp[-1] is None:
        return ''
    return ' '.join(dp[-1][1])

def title_case_course_name(value):
    lowered_words = {'and', 'in', 'of', 'for', 'to', 'with'}
    titled = []
    for index, word in enumerate(str(value or '').split()):
        lower = word.lower()
        if index > 0 and lower in lowered_words:
            titled.append(lower)
        else:
            titled.append(lower[:1].upper() + lower[1:])
    return ' '.join(titled)

def repair_remaining_pdf_fragments(value):
    def replace_match(match):
        fragment = match.group(0)
        segmented = segment_compact_words(fragment)
        return segmented if segmented else fragment

    value = re.sub(r'\b(?:[A-Za-z]{1,4}\s+){2,}[A-Za-z]{1,6}\b', replace_match, value)
    value = re.sub(r'\b([A-Za-z]+men)\s+t\b', lambda m: m.group(1) + 't', value, flags=re.I)
    value = re.sub(r'\b([A-Za-z]+ica)\s+l\b', lambda m: m.group(1) + 'l', value, flags=re.I)
    value = re.sub(
        r'\b(algorithm|application|concept|dataset|format|level|method|practice|procedure|solution|structure|system)\s+s\b',
        lambda m: m.group(1) + 's',
        value,
        flags=re.I
    )
    value = re.sub(r'\b([A-Za-z]{3,})\s+([bcdefghjklmnopqrstuvwxyz])\b(?=\s*(?:[.,;:]|$))', lambda m: m.group(1) + m.group(2), value)
    value = re.sub(
        r'\b([A-Za-z]{2,})\s+(ated|ence|hms|rstanding|edictive|soning|ment|tion|sion|tions|sions)\b',
        lambda m: m.group(1) + m.group(2),
        value,
        flags=re.I
    )
    value = re.sub(r'\b([B-HJ-Zb-hj-km-ru-z])\s+([a-z]{3,})\b', lambda m: m.group(1) + m.group(2), value)
    value = re.sub(r'\bdat\s+a\b|\bda\s+ta\b', 'data', value, flags=re.I)
    value = re.sub(r'\ba\s+lgorithm(s?)\b', r'algorithm\1', value, flags=re.I)
    value = re.sub(r'\balgorithm\s+s\b', 'algorithms', value, flags=re.I)
    value = re.sub(r'\bs\s+tandards\b', 'standards', value, flags=re.I)
    return value

def clean_pdf_fragment(value, preserve_colon=False, preserve_topic_punctuation=False):
    value = re.sub(r'[\x00-\x1f]+', ' ', value or '')
    value = re.sub(r'\\', ' ', value)
    value = re.sub(r'\s*&\s*', ' and ', value)
    if preserve_topic_punctuation:
        value = re.sub(r'/+', ' ', value)
        value = re.sub(r'\s*;\s*', '; ', value)
    else:
        value = re.sub(r'[/;]+', ' ', value)
    if preserve_colon:
        value = re.sub(r'\s*:\s*', ': ', value)
    else:
        value = re.sub(r':+', ' ', value)
    value = re.sub(r'\s{3,}', '  ', value).strip()
    groups = re.split(r'\s{2,}', value)
    cleaned_groups = []
    stopwords = {'of', 'in', 'to', 'on', 'as', 'is', 'be', 'or', 'and', 'the', 'for', 'with', 'that'}

    for group in groups:
        tokens = [token for token in group.split() if token]
        small_token_ratio = sum(1 for token in tokens if len(token.strip('.,;:-')) <= 3) / len(tokens) if tokens else 0
        segmented = segment_compact_words(group) if len(tokens) >= 2 and small_token_ratio > 0.4 else ''
        if segmented:
            suffix = ''
            if re.search(r'\.\s*$', group):
                suffix = '.'
            elif re.search(r',\s*$', group):
                suffix = ','
            cleaned_groups.append(segmented + suffix)
        else:
            cleaned_groups.append(group)

    cleaned = re.sub(r'\s+', ' ', ' '.join(cleaned_groups)).strip()
    return repair_remaining_pdf_fragments(cleaned)

def extract_course_spec_section(raw_text, start_label, end_label):
    start_matches = list(re.finditer(flexible_label(start_label), raw_text, flags=re.I | re.S))
    end_matches = list(re.finditer(flexible_label(end_label), raw_text, flags=re.I | re.S))
    candidates = []
    for start_match in start_matches:
        next_end = next((end_match for end_match in end_matches if end_match.start() > start_match.end()), None)
        if next_end:
            section = raw_text[start_match.end():next_end.start()]
            clo_ids = len(re.findall(r'\b[123]\s*\.\s*[1-9]\d*\b', section))
            candidates.append((clo_ids, len(section), section))
    if not candidates:
        return ''
    return max(candidates, key=lambda item: (item[0], item[1]))[2]

TOPIC_START_PATTERN = re.compile(
    r'('
    r'\u0642\u0627\u0626\u0645\u0629\s+\u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a|'
    r'\u0645\u062d\u062a\u0648\u0649\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0645\u0648\u0636\u0648\u0639\u0627\u062a\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0627\u0644\u0645\u062d\u062a\u0648\u0649\s+\u0627\u0644\u062f\u0631\u0627\u0633\u064a|'
    r'List\s+of\s+Topics|Course\s+Content|Course\s+Topics'
    r')',
    re.I
)

TOPIC_HEADER_PATTERN = re.compile(
    r'('
    r'\u0642\u0627\u0626\u0645\u0629\s+\u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a|'
    r'\u0645\u062d\u062a\u0648\u0649\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0645\u0648\u0636\u0648\u0639\u0627\u062a\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0633\u0627\u0639\u0627\u062a\s+\u0627\u0644\u0627\u062a\u0635\u0627\u0644|'
    r'\u0627\u0644\u0633\u0627\u0639\u0627\u062a\s+\u0627\u0644\u062a\u062f\u0631\u064a\u0633\u064a\u0629|'
    r'\u0631\u0642\u0645|'
    r'List\s+of\s+Topics|Contact\s+Hours|No\.'
    r')',
    re.I
)

TOPIC_END_PATTERN = re.compile(
    r'(^|\s)('
    r'\u0627\u0644\u0645\u062c\u0645\u0648\u0639|'
    r'\u0627\u0644\u062a\u0642\u064a\u064a\u0645|'
    r'\u0645\u0635\u0627\u062f\u0631\s*\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u0627\u0644\u0645\u0631\u0627\u0641\u0642|'
    r'\u0627\u0644\u062a\u0639\u0644\u064a\u0645\s+\u0648\u0627\u0644\u062a\u0642\u064a\u064a\u0645|'
    r'\u0637\u0631\u0642\s+\u0627\u0644\u062a\u062f\u0631\u064a\u0633|'
    r'\u0623\u0646\u0634\u0637\u0629\s+\u0627\u0644\u062a\u0642\u064a\u064a\u0645|'
    r'\u0627\u0639\u062a\u0645\u0627\u062f\s+\u0627\u0644\u062a\u0648\u0635\u064a\u0641|'
    r'Total|D\.\s*Students\s+Assessment|Students\s+Assessment|'
    r'E\.\s*Learning\s+Resources|Learning\s+Resources|Assessment|Specification\s+Approval'
    r')',
    re.I
)

ARABIC_TOPIC_ANCHOR_PATTERN = re.compile(
    r'[\u0600-\u06FF][^:\uff1a]{2,90}[:\uff1a]\s*[\u0600-\u06FFA-Za-z]'
)

TOPIC_NOISE_PATTERN = re.compile(
    r'('
    r'Education\s+(?:and|&)?\s*Training\s+Evaluation|ETEC|GOV\.SA|'
    r'\u0647\u064a\u0626\u0629\s+\u062a\u0642\u0648\u064a\u0645|'
    r'\u062a\u0648\u0635\u064a\u0641\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0646\u0648\u0627\u062a\u062c\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u0627\u0633\u062a\u0631\u0627\u062a\u064a\u062c\u064a\u0627\u062a\s+\u0627\u0644\u062a\u062f|'
    r'\u0645\u0635\u0627\u062f\u0631\s*\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u062a\u0642\u0648\u064a\u0645\s+\u062c\u0648\u062f\u0629|'
    r'\u0627\u0639\u062a\u0645\u0627\u062f\s+\u0627\u0644\u062a\u0648\u0635\u064a\u0641|'
    r'\u0623\u0646\u0634\u0637\u0629\s+\u0627\u0644\u062a\u0642\u064a\u064a\u0645|'
    r'\u0627\u0644\u0645\u062d\u0627\u0636\u0631\u0629|'
    r'\u0627\u0644\u0645\u0646\u0627\u0642\u0634\u0629|'
    r'\u0627\u0644\u0648\u0627\u062c\u0628\u0627\u062a|'
    r'\u0627\u0644\u062a\u0643\u0627\u0644\u064a\u0641|'
    r'\u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631|'
    r'\u0627\u0644\u0628\u062d\u0648\u062b|'
    r'\u0627\u0644\u062a\u0642\u0627\u0631\u064a\u0631|'
    r'\u0627\u0644\u062a\u0641\u0643\u064a\u0631\s+\u0627\u0644\u0646\u0627\u0642\u062f|'
    r'\u062d\u0644\s+\u0627\u0644\u0645\u0634\u0643\u0644\u0627\u062a|'
    r'\u0627\u0633\u062a\u0628\u0627\u0646\u0629|'
    r'Course\s+Learning\s+Outcomes|Teaching\s+Strategies|Assessment\s+Methods|'
    r'Course\s+Identification|Course\s+General\s+Description|Course\s+Main\s+Objective'
    r')',
    re.I
)

def clean_arabic_topic_ocr_artifacts(topic):
    if not contains_arabic(topic):
        return topic
    replacements = {
        '\u00a5': ' ',
        '\u00a2': ' ',
        '@': ' ',
        '\u00a9': ' ',
        '\u00ae': ' ',
        '\u00ab': ' ',
        '\u00bb': ' ',
        '\\': ' ',
    }
    for old, new in replacements.items():
        topic = topic.replace(old, new)
    topic = re.sub(r'\(\s*@?\s*Al\s*\)', ' ', topic, flags=re.I)
    topic = re.sub(r'\(\s*[A-Za-z]+\s*[-–—]?\s*', ' ', topic)
    topic = re.sub(r'(?<![A-Za-z])(?:Y|V|v|L|N|AR)(?![A-Za-z])', ' ', topic)
    topic = re.sub(r'\b[A-Za-z]{1,4}\b', ' ', topic)
    topic = topic.replace('\u062a\u0648\u0627\u0632\u0644', '\u0646\u0648\u0627\u0632\u0644')
    topic = topic.replace('\u062a\u0648\u0627\u0626\u0644', '\u0646\u0648\u0627\u0632\u0644')
    topic = topic.replace('\u0627\u0644\u0635\u0628\u064a\u0627\u0645', '\u0627\u0644\u0635\u064a\u0627\u0645')
    topic = topic.replace('\u062a\u0648\u0627\u0632\u0644 \u0627\u0644\u062d\u062c', '\u0646\u0648\u0627\u0632\u0644 \u0627\u0644\u062d\u062c')
    topic = topic.replace('\u062a\u0648\u0627\u0626\u0644 \u0627\u0644\u062d\u062c', '\u0646\u0648\u0627\u0632\u0644 \u0627\u0644\u062d\u062c')
    topic = re.sub(r'\s+[.]\s*"\s*$', '', topic)
    topic = re.sub(r'["“”]+', ' ', topic)
    topic = re.sub(r'\s+([،؛:,.])', r'\1', topic)
    return compact_text(topic)

def clean_topic_candidate(value):
    raw_topic = str(value or '')
    raw_topic = re.sub(r'^\s*(?:[\-\u2022\u2023\u25e6\u2043\u2219]|[0-9]{1,2}[\.\-\)]?)\s*', '', raw_topic)
    raw_topic = re.sub(r'\s+[0-9]{1,2}(?:\s*%)?\s*$', '', raw_topic).strip()
    raw_topic = re.sub(r'\s+[0-9]{1,2}\s*$', '', raw_topic).strip()
    topic = clean_pdf_fragment(raw_topic, preserve_colon=True, preserve_topic_punctuation=True)
    topic = clean_arabic_topic_ocr_artifacts(topic)
    topic = re.sub(r'^\s*(?:[\-\u2022\u2023\u25e6\u2043\u2219]|[0-9]{1,2}[\.\-\)]?)\s*', '', topic)
    topic = re.sub(r'\s+[0-9]{1,2}(?:\s*%)?\s*$', '', topic).strip()
    topic = re.sub(r'\s+[0-9]{1,2}\s*$', '', topic).strip()
    topic = clean_arabic_topic_ocr_artifacts(topic)
    topic = topic.strip(' :.-')
    if len(topic) < 5 or len(topic) > 420:
        return ''
    if TOPIC_HEADER_PATTERN.search(topic) or TOPIC_NOISE_PATTERN.search(topic):
        return ''
    if not re.search(r'[A-Za-z\u0600-\u06FF]', topic):
        return ''
    return topic

def dedupe_topics(topics):
    seen = set()
    filtered = []
    for topic in topics:
        key = re.sub(r'\s+', ' ', topic).strip().casefold()
        if key and key not in seen:
            seen.add(key)
            filtered.append(topic)
    return filtered

def extract_numbered_topics_from_text(text):
    normalized = compact_text(text)
    normalized = re.sub(r'\s+([0-9]{1,2})\s+([0-9]{1,2})\s+([A-Z][A-Za-z])', r' \1 \2. \3', normalized)
    normalized = re.sub(r'^.*?(?:No\s+List\s+of\s+Topics\s+Contact\s+Hours|Course\s+Content)', '', normalized, flags=re.I)
    normalized = re.sub(r'\b(?:D\.\s*)?Students\s+Assessment\b.*$', '', normalized, flags=re.I)
    normalized = re.sub(r'\b(?:E\.\s*)?Learning\s+Resources\b.*$', '', normalized, flags=re.I)
    normalized = re.sub(r'\bTotal\s+\d+.*$', '', normalized, flags=re.I)
    topics = []
    pattern = re.compile(
        r'(?:^|\s)([0-9]{1,2})[\.\)]\s+'
        r'(.+?)(?=\s+[0-9]{1,2}[\.\)]\s+|\s+Total\b|\s+D\.\s+Students|\s+E\.\s+Learning|$)',
        re.I | re.S
    )
    for match in pattern.finditer(normalized):
        topic = clean_topic_candidate(match.group(2))
        if topic:
            topics.append(topic)
    return dedupe_topics(topics)

def extract_line_topics(lines, start_index, end_index):
    topics = []
    for line in lines[start_index:end_index]:
        cleaned = clean_topic_candidate(line)
        if not cleaned:
            continue
        if re.match(r'^[A-Za-z]\.\s+', cleaned):
            continue
        topics.append(cleaned)
    return dedupe_topics(topics)

def score_topic_candidates(topics):
    score = len(topics) * 10
    score += sum(4 for topic in topics if re.match(r'^[A-Za-z\u0600-\u06FF]', topic))
    score -= sum(8 for topic in topics if len(topic) > 260)
    return score

ARABIC_TOPIC_ADMIN_PATTERN = re.compile(
    r'('
    r'\u0627\u0633\u0645\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0627\u0644\u062a\u0639\u0631\u064a\u0641\s+\u0628\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0627\u0644\u0633\u0627\u0639\u0627\u062a\s+\u0627\u0644\u0645\u0639\u062a\u0645\u062f\u0629|'
    r'\u0627\u0644\u0633\u0646\u0629\s*/?\s*\u0627\u0644\u0645\u0633\u062a\u0648\u0649|'
    r'\u0627\u0644\u0648\u0635\u0641\s+\u0627\u0644\u0639\u0627\u0645|'
    r'\u064a\u062a\u0646\u0627\u0648\u0644\s+\u0647\u0630\u0627\s+\u0627\u0644\u0645\u0642\u0631\u0631|'
    r'\u0627\u0644\u0645\u062a\u0637\u0644\u0628\u0627\u062a|'
    r'\u0627\u0644\u0647\u062f\u0641\s+\u0627\u0644\u0631\u0626\u064a\u0633|'
    r'\u0646\u0645\u0637\s+\u0627\u0644\u062a\u0639\u0644\u064a\u0645|'
    r'\u0627\u0644\u062a\u0639\u0644\u064a\u0645\s+\u0627\u0644(?:\u0625|\u0627)\u0644\u0643\u062a\u0631\u0648\u0646\u064a|'
    r'\u0627\u0644\u062a\u0639\u0644\u064a\u0645\s+\u0627\u0644\u0645\u062f\u0645\u062c|'
    r'\u0627\u0644\u062a\u0639\u0644\u064a\u0645\s+\u0639\u0646\s+\u0628\u0639\u062f|'
    r'\u0627\u0644\u0633\u0627\u0639\u0627\u062a\s+\u0627\u0644\u062a\u062f\u0631\u064a\u0633\u064a\u0629|'
    r'\u0645\u062d\u0627\u0636\u0631\u0627\u062a|'
    r'\u0645\u0639\u0645\u0644|'
    r'\u0645\u064a\u062f\u0627\u0646\u064a|'
    r'\u062f\u0631\u0648\u0633\s+\u0625\u0636\u0627\u0641\u064a\u0629|'
    r'\u0627\u0644\u0631\u0645\u0632|'
    r'\u0646\u0648\u0627\u062a\u062c\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u0646\u0627\u062a\u062c\s+\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u0627\u0633\u062a\u0631\u0627\u062a\u064a\u062c\u064a\u0627\u062a\s+\u0627\u0644\u062a\u062f|'
    r'\u0637\u0631\u0642\s+\u0627\u0644\u062a\u0642\u064a\u064a\u0645|'
    r'\u0645\u0635\u0627\u062f\u0631\s*\u0627\u0644\u062a\u0639\u0644\u0645|'
    r'\u0627\u0644\u0645\u0631\u0627\u0641\u0642|'
    r'\u062a\u0642\u0648\u064a\u0645\s+\u062c\u0648\u062f\u0629|'
    r'\u0627\u0639\u062a\u0645\u0627\u062f\s+\u0627\u0644\u062a\u0648\u0635\u064a\u0641|'
    r'\u0647\u064a\u0626\u0629\s+\u062a\u0642\u0648\u064a\u0645|'
    r'Education\s+&?\s*Training\s+Evaluation'
    r')',
    re.I
)

def is_probable_arabic_topic_start(line):
    cleaned = clean_topic_candidate(line)
    if not cleaned or not contains_arabic(cleaned):
        return False
    if ARABIC_TOPIC_ADMIN_PATTERN.search(cleaned):
        return False
    if not ARABIC_TOPIC_ANCHOR_PATTERN.search(cleaned):
        return False
    label = re.split(r'[:\uff1a]', cleaned, maxsplit=1)[0]
    label_words = re.findall(r'[\u0600-\u06FF]+', label)
    return 1 <= len(label_words) <= 8

def extract_probable_arabic_topic_block(lines):
    candidates = []
    for start_index, line in enumerate(lines):
        if not is_probable_arabic_topic_start(line):
            continue

        end_index = min(len(lines), start_index + 90)
        quiet_lines = 0
        for index in range(start_index + 1, min(len(lines), start_index + 120)):
            current_line = lines[index]
            if TOPIC_END_PATTERN.search(current_line) or re.match(r'^\s*(?:\u0627\u0644?\u062c\u0645\u064a\w*|\u0627\u0644?\u0645\u062c\u0645\w*|Total)\b', current_line, flags=re.I):
                end_index = index
                break
            cleaned = clean_topic_candidate(current_line)
            if cleaned and contains_arabic(cleaned) and not ARABIC_TOPIC_ADMIN_PATTERN.search(cleaned):
                quiet_lines = 0
                continue
            quiet_lines += 1
            if quiet_lines >= 4:
                end_index = max(start_index + 1, index - quiet_lines + 1)
                break

        topics = extract_line_topics(lines, start_index, end_index)
        if len(topics) >= 2:
            candidates.append((score_topic_candidates(topics), topics))

    if not candidates:
        return []
    return max(candidates, key=lambda item: item[0])[1]

def extract_course_topics(raw_text):
    text = normalize_course_spec_text(raw_text)
    lines = [compact_text(line) for line in (text or '').splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        return []

    candidates = []
    for start_index, line in enumerate(lines):
        if not TOPIC_START_PATTERN.search(line):
            continue
        if re.search(r'\.{3,}|Table\s+of\s+Contents', line, re.I):
            continue

        section_lines = []
        end_index = min(len(lines), start_index + 80)
        found_content = bool(re.search(r'\b1[\.\)]\s+', line))
        for index in range(start_index, min(len(lines), start_index + 90)):
            current_line = lines[index]
            if index > start_index and TOPIC_END_PATTERN.search(current_line) and not found_content and index - start_index <= 4:
                section_lines = []
                end_index = index
                break
            if index > start_index and TOPIC_END_PATTERN.search(current_line) and found_content:
                section_lines.append(current_line)
                end_index = index
                break
            if index > start_index and re.search(r'\b[0-9]{1,2}[\.\)]\s+', current_line):
                found_content = True
            section_lines.append(current_line)
        if not section_lines:
            continue

        section_text = ' '.join(section_lines)
        numbered_topics = extract_numbered_topics_from_text(section_text)
        line_topics = extract_line_topics(lines, start_index + 1, end_index)
        topics = numbered_topics if len(numbered_topics) >= len(line_topics) else line_topics
        if topics:
            toc_penalty = -50 if re.search(r'\.{3,}|Table\s+of\s+Contents', line, re.I) else 0
            candidates.append((score_topic_candidates(topics) + toc_penalty, topics))

    if candidates:
        return max(candidates, key=lambda item: item[0])[1]

    # Fallback for OCR/PDF text where the topic heading is missing but the table body survived.
    arabic_topics = extract_probable_arabic_topic_block(lines)
    if arabic_topics:
        return arabic_topics

    description_index = next(
        (
            index for index, line in enumerate(lines)
            if re.search(r'\u064a\u062a\u0646\u0627\u0648\u0644\s+\u0647\u0630\u0627\s+\u0627\u0644\u0645\u0642\u0631\u0631.*\u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a', line)
        ),
        -1
    )
    if description_index >= 0:
        description_end = next(
            (
                index for index in range(description_index + 1, min(len(lines), description_index + 12))
                if re.search(r'\u0627\u0644\u0645\u062a\u0637\u0644\u0628\u0627\u062a|\u0627\u0644\u0647\u062f\u0641\s+\u0627\u0644\u0631\u0626\u064a\u0633', lines[index])
            ),
            min(len(lines), description_index + 8)
        )
        description_text = ' '.join(lines[description_index + 1:description_end])
        pieces = re.split(r'[.;\u061b]\s+|[\u060c]\s+', description_text)
        return dedupe_topics([topic for topic in (clean_topic_candidate(piece) for piece in pieces) if topic])

    return []

def strip_arabic_outcome_noise(value):
    value = re.sub(r'[\u200e\u200f\u061c\u202a-\u202e]', ' ', value or '')
    value = re.sub(r'[\u064B-\u065F\u0670]', '', value)
    value = re.sub(r'\b(?:Education|Training|Evaluation|Commission|ETEC|GOV|SA)\b', ' ', value, flags=re.I)
    value = re.sub(r'\b(?:المحاضرة|المناقشة|الواجبات|الاختبار|شفوي|تحريري|استبانة|التقييم|أوراق العمل|البحوث|والتقارير|التكاليف|تقويم الأقران)\b', ' ', value)
    value = re.sub(r'[A-Za-z0-9@#%+<>()\\/\[\]{}|=_*"\']+', ' ', value)
    value = re.sub(r'\s+', ' ', value).strip(' .،؛:-')
    return value

def arabic_outcome_domain_for_text(value):
    value = value or ''
    if re.match(r'^\s*(?:يعرف|يذكر|يظهر|تعرض)', value):
        return '1'
    if re.match(r'^\s*(?:يطبق|يوظف|يستخدم|يحلل|يقارن|يناقش)', value):
        return '2'
    if re.match(r'^\s*(?:يلتزم|يتحلى|يقدر|يبادر|يتعاون|يشارك)', value):
        return '3'
    return ''

def extract_noisy_arabic_ocr_clos(raw_text):
    if not contains_arabic(raw_text):
        return {}

    lines = [compact_text(line) for line in normalize_course_spec_text(raw_text).splitlines()]
    lines = [line for line in lines if line]
    start_index = next((i for i, line in enumerate(lines) if 'نواتج التعلم للمقرر' in line), -1)
    if start_index < 0:
        return {}
    end_index = next(
        (
            i for i in range(start_index + 1, len(lines))
            if i > start_index + 10 and any(label in lines[i] for label in ['موضوعات المقرر', 'الساعات التدريسية', 'مصادر التعلم'])
        ),
        min(len(lines), start_index + 90)
    )
    section_lines = lines[start_index:end_index]
    starter_pattern = re.compile(r'(يعرف|يعرّف|يذكر|يظهر|تعرض|يطبق|يوظف|يستخدم|يلتزم|يتحلى|يقدر|يبادر|يتعاون|يشارك|يحلل|يقارن|يناقش)')
    outcomes = []
    current = ''

    for line in section_lines:
        line = strip_arabic_outcome_noise(line)
        if not line or len(line) < 4:
            continue
        match = starter_pattern.search(line)
        if match:
            if current:
                outcomes.append(current)
            current = line[match.start():]
        elif current and len(current) < 260:
            current = f"{current} {line}"
    if current:
        outcomes.append(current)

    cleaned = []
    seen = set()
    for outcome in outcomes:
        outcome = strip_arabic_outcome_noise(outcome)
        outcome = re.sub(r'\s+', ' ', outcome).strip()
        if len(outcome) < 18:
            continue
        if any(stop in outcome for stop in ['نمط التعليم', 'الساعات التدريسية', 'موضوعات المقرر']):
            continue
        key = outcome[:45]
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(outcome)

    clo_map = {}
    counts = {'1': 0, '2': 0, '3': 0}
    for outcome in cleaned:
        domain = arabic_outcome_domain_for_text(outcome)
        if not domain:
            continue
        counts[domain] += 1
        clo_id = f"{domain}.{counts[domain]}"
        clo_map[clo_id] = f"{clo_id} {clean_clo_text(outcome)}"

    if not any(key.startswith('3.') for key in clo_map):
        for index, line in enumerate(section_lines):
            candidate = strip_arabic_outcome_noise(' '.join(section_lines[index:index + 3]))
            if candidate.startswith('يلتزم'):
                counts['3'] += 1
                clo_id = f"3.{counts['3']}"
                clo_map[clo_id] = f"{clo_id} {clean_clo_text(candidate)}"
                break
    if not any(key.startswith('3.') for key in clo_map):
        value_match = re.search(r'(يلتزم\s+.{12,220}?)(?=\s+(?:الساعات|موضوعات|مصادر|حقيقة|نمط التعليم|$))', normalize_course_spec_text(raw_text), flags=re.S)
        if value_match:
            candidate = strip_arabic_outcome_noise(value_match.group(1))
            if len(candidate) > 18:
                counts['3'] += 1
                clo_id = f"3.{counts['3']}"
                clo_map[clo_id] = f"{clo_id} {clean_clo_text(candidate)}"
    return clo_map

def infer_arabic_course_name_from_ocr(raw_text):
    text = normalize_course_spec_text(raw_text)
    known_title = re.search(r'فقه\s+القضاء\s+وطرق\s+الإثبات', text)
    if known_title:
        return compact_text(known_title.group(0))
    candidates = []
    for match in re.finditer(r'([\u0600-\u06FF]{2,}(?:\s+[\u0600-\u06FF]{2,}){1,6})', text):
        phrase = compact_text(match.group(1))
        if len(phrase) < 8 or len(phrase) > 80:
            continue
        if any(skip in phrase for skip in ['هيئة تقويم', 'توصيف المقرر', 'معلومات عامة', 'نواتج التعلم', 'موضوعات المقرر']):
            continue
        score = text.count(phrase)
        if 'فقه' in phrase or 'القضاء' in phrase or 'الإثبات' in phrase:
            score += 5
        candidates.append((score, len(phrase), phrase))
    if not candidates:
        return ''
    return max(candidates, key=lambda item: (item[0], item[1]))[2]

def arabic_label_key(value):
    value = normalize_course_spec_text(value)
    value = re.sub(r'[\u064B-\u065F\u0670]', '', value)
    return re.sub(r'[^\u0600-\u06FF0-9A-Za-z]+', '', value or '')

def line_has_arabic_label(line, label):
    return arabic_label_key(label) in arabic_label_key(line)

def infer_course_name_from_filename(filename):
    base = os.path.splitext(os.path.basename(filename or ''))[0]
    base = normalize_course_spec_text(base)
    base = re.sub(r'[_\-]+', ' ', base)
    base = re.sub(r'^\s*[A-Z]{2,6}\s*\d{3,4}[A-Z]?\s+', ' ', base, flags=re.I)
    base = re.sub(r'\b\d{6,}(?:\s+\d{4,})*\b', ' ', base)
    base = re.sub(r'\b(?:specification|specifications|course|report)\b', ' ', base, flags=re.I)
    for word in [
        '\u062a\u0648\u0635\u064a\u0641',
        '\u0645\u0642\u0631\u0631',
        '\u0627\u0644\u0645\u0642\u0631\u0631',
        '\u0627\u0644\u062f\u0631\u0627\u0633\u064a',
        '\u062f\u0631\u0627\u0633\u064a',
    ]:
        base = re.sub(rf'(^|\s){word}(\s|$)', ' ', base)
    base = re.sub(r'\s+', ' ', base).strip(' .-_()[]')
    return base if len(base) >= 4 else ''

def is_bad_arabic_course_name(value):
    key = arabic_label_key(value)
    bad_keys = [
        arabic_label_key('\u0631\u0645\u0632 \u0627\u0644\u0645\u0642\u0631\u0631'),
        arabic_label_key('\u0627\u0644\u0642\u0633\u0645 \u0627\u0644\u0639\u0644\u0645\u064a'),
        arabic_label_key('\u0627\u0644\u0628\u0631\u0646\u0627\u0645\u062c'),
    ]
    return not key or any(bad in key for bad in bad_keys)

def clean_arabic_outcome_line(line):
    line = normalize_course_spec_text(line)
    line = re.sub(r'[\u064B-\u065F\u0670]', '', line)
    line = re.sub(r'\b(?:Education|Training|Evaluation|Commission|ETEC|GOV|SA)\b', ' ', line, flags=re.I)
    line = re.sub(r'^[^\u0600-\u06FF]+', ' ', line)
    line = re.sub(r'^[0-9\u0660-\u0669\u06F0-\u06F9\s\.\-:،؛]+', ' ', line)
    for marker in [
        '\u0627\u0644\u0645\u062d\u0627\u0636\u0631\u0629',
        '\u0627\u0644\u0645\u0646\u0627\u0642\u0634\u0629',
        '\u0627\u0644\u0648\u0627\u062c\u0628\u0627\u062a',
        '\u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631',
        '\u0627\u0644\u0623\u0628\u062d\u0627\u062b',
        '\u0627\u0644\u0628\u062d\u0648\u062b',
        '\u0627\u0644\u062a\u0642\u0627\u0631\u064a\u0631',
        '\u0627\u0644\u062a\u0643\u0627\u0644\u064a\u0641',
        '\u0627\u0644\u062a\u0639\u0644\u064a\u0645 \u0627\u0644\u062a\u0639\u0627\u0648\u0646\u064a',
        '\u0627\u0644\u062a\u0639\u0644\u064a\u0645 \u0627\u0644\u0630\u0627\u062a\u064a',
        '\u0627\u0633\u062a\u0628\u0627\u0646\u0629',
        '\u062a\u0642\u0648\u064a\u0645',
        '\u0627\u0644\u0642\u064a\u0645 \u0648\u0627\u0644\u0627\u0633\u062a\u0642\u0644\u0627\u0644\u064a\u0629',
    ]:
        marker_index = line.find(marker)
        if marker_index > 10:
            line = line[:marker_index]
    line = re.sub(r'[A-Za-z0-9@#%+<>()\[\]{}|=_*"\'/\\]+', ' ', line)
    line = re.sub(r'[^\u0600-\u06FF\s.,،؛:\-]', ' ', line)
    line = re.sub(r'\s+', ' ', line).strip(' .،؛:-»«')
    line = line.replace('\u0627\u0644\u062a\u0648\u0627\u0632\u0644', '\u0627\u0644\u0646\u0648\u0627\u0632\u0644')
    line = line.replace('\u062a\u0648\u0627\u0632\u0644', '\u0646\u0648\u0627\u0632\u0644')
    line = line.replace('\u0627\u0644\u0645\u0648\u062b\u0631\u0629', '\u0627\u0644\u0645\u0624\u062b\u0631\u0629')
    line = line.replace('\u062a\u0639\u0631\u0636 \u0627\u0644\u0646\u0648\u0627\u0632\u0644', '\u064a\u0633\u062a\u0639\u0631\u0636 \u0627\u0644\u0646\u0648\u0627\u0632\u0644')
    line = re.sub(r'\b\u0641\u064a\s+\u0645\s+(?=[\u0600-\u06FF])', '\u0641\u064a ', line)
    line = re.sub(r'\b\u0641\u064a\s+\u062f\u0631\u0627\b', '\u0641\u064a \u062f\u0631\u0627\u0633\u0629', line)
    line = re.sub(r'\b\u062a\u0648\u0635\u0644\s+\u0645\s+\u0648\u062f\s+\u0625\u0644\u0627\b', '\u062a\u0648\u0635\u0644 \u0625\u0644\u064a\u0647\u0627', line)
    line = re.sub(r'\s+\u0648\s*$', '', line)
    return line

def arabic_starter_domain(value):
    value = clean_arabic_outcome_line(value)
    if re.match(r'^(?:\u064a\u0628\u064a\u0646|\u064a\u0648\u0636\u062d|\u064a\u0634\u0631\u062d|\u064a\u0635\u0641|\u064a\u0639\u0631\u0641|\u064a\u0630\u0643\u0631|\u064a\u0633\u062a\u0639\u0631\u0636|\u062a\u0639\u0631\u0636)', value):
        return '1'
    if re.match(r'^(?:\u064a\u0624\u0635\u0644|\u064a\u0637\u0628\u0642|\u064a\u062d\u0644\u0644|\u064a\u0642\u0627\u0631\u0646|\u064a\u0646\u0627\u0642\u0634|\u064a\u062c\u064a\u062f|\u064a\u0633\u062a\u062e\u062f\u0645|\u064a\u0648\u0638\u0641|\u064a\u0633\u062a\u0646\u0628\u0637)', value):
        return '2'
    if re.match(r'^(?:\u064a\u0644\u062a\u0632\u0645|\u064a\u0642\u062f\u0631|\u064a\u062a\u062d\u0644\u0649|\u064a\u062a\u0639\u0627\u0648\u0646|\u064a\u0628\u0627\u062f\u0631|\u064a\u062a\u062d\u0645\u0644|\u064a\u0634\u0627\u0631\u0643)', value):
        return '3'
    return ''

def extract_arabic_table_clos(raw_text):
    if not contains_arabic(raw_text):
        return {}

    lines = [compact_text(line) for line in normalize_course_spec_text(raw_text).splitlines()]
    lines = [line for line in lines if line]
    start_indexes = [
        index for index, line in enumerate(lines)
        if (
            line_has_arabic_label(line, '\u0627\u0644\u0631\u0645\u0632')
            and (
                line_has_arabic_label(line, '\u0646\u0648\u0627\u062a\u062c')
                or any(line_has_arabic_label(next_line, '\u0646\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645') for next_line in lines[index:index + 4])
            )
        )
    ]
    if not start_indexes:
        start_indexes = [
            index for index, line in enumerate(lines)
            if line_has_arabic_label(line, '\u0646\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645')
        ]
    if not start_indexes:
        return {}

    section_lines = []
    for start_index in start_indexes[:2]:
        for line in lines[start_index + 1:start_index + 45]:
            if line_has_arabic_label(line, '\u0645\u0648\u0636\u0648\u0639\u0627\u062a \u0627\u0644\u0645\u0642\u0631\u0631') or line_has_arabic_label(line, '\u0645\u0635\u0627\u062f\u0631 \u0627\u0644\u062a\u0639\u0644\u0645'):
                break
            if re.match(r'^\s*(?:\u0646\u0648\u0627\u0632\u0644|\u0642\u0627\u0626\u0645\u0629\s+\u0627\u0644\u0645\u0648\u0636\u0648\u0639\u0627\u062a)\b', line):
                break
            section_lines.append(line)

    starter_pattern = re.compile(
        r'(\u064a\u0628\u064a\u0646|\u064a\u0648\u0636\u062d|\u064a\u0634\u0631\u062d|\u064a\u0635\u0641|\u064a\u0639\u0631\u0641|\u064a\u0630\u0643\u0631|\u064a\u0633\u062a\u0639\u0631\u0636|\u062a\u0639\u0631\u0636|'
        r'\u064a\u0624\u0635\u0644|\u064a\u0637\u0628\u0642|\u064a\u062d\u0644\u0644|\u064a\u0642\u0627\u0631\u0646|\u064a\u0646\u0627\u0642\u0634|\u064a\u062c\u064a\u062f|\u064a\u0633\u062a\u062e\u062f\u0645|\u064a\u0648\u0638\u0641|\u064a\u0633\u062a\u0646\u0628\u0637|'
        r'\u064a\u0644\u062a\u0632\u0645|\u064a\u0642\u062f\u0631|\u064a\u062a\u062d\u0644\u0649|\u064a\u062a\u0639\u0627\u0648\u0646|\u064a\u0628\u0627\u062f\u0631|\u064a\u062a\u062d\u0645\u0644|\u064a\u0634\u0627\u0631\u0643)'
    )
    outcomes = []
    current = ''
    for line in section_lines:
        if any(
            line_has_arabic_label(line, label)
            for label in [
                '\u0627\u0633\u0645 \u0627\u0644\u0645\u0642\u0631\u0631',
                '\u0631\u0645\u0632 \u0627\u0644\u0645\u0642\u0631\u0631',
                '\u0627\u0644\u0642\u0633\u0645 \u0627\u0644\u0639\u0644\u0645\u064a',
                '\u0627\u0644\u0643\u0644\u064a\u0629',
                '\u0627\u0644\u0645\u0624\u0633\u0633\u0629',
                '\u062a\u0648\u0635\u064a\u0641 \u0627\u0644\u0645\u0642\u0631\u0631'
            ]
        ):
            if current:
                outcomes.append(current)
                current = ''
            break
        if line_has_arabic_label(line, '\u0627\u0644\u0642\u064a\u0645 \u0648\u0627\u0644\u0627\u0633\u062a\u0642\u0644\u0627\u0644\u064a\u0629') or line_has_arabic_label(line, '\u0647\u064a\u0626\u0629 \u062a\u0642\u0648\u064a\u0645'):
            if current:
                outcomes.append(current)
                current = ''
            continue
        cleaned_line = clean_arabic_outcome_line(line)
        if len(cleaned_line) < 4:
            continue
        if cleaned_line in {'\u0627\u0644\u0627\u0633\u062a\u0646\u0628\u0627\u0637\u064a\u0629', '\u0627\u0644\u062a\u0641\u0643\u064a\u0631 \u0627\u0644\u0646\u0627\u0642\u062f'}:
            continue
        starter_match = starter_pattern.search(cleaned_line)
        if starter_match and (starter_match.start() < 8 or not current):
            if current:
                outcomes.append(current)
            current = cleaned_line[starter_match.start():]
        elif current and len(current) < 260:
            current = f"{current} {cleaned_line}"
    if current:
        outcomes.append(current)

    clo_map = {}
    counts = {'1': 0, '2': 0, '3': 0}
    seen = set()
    for outcome in outcomes:
        outcome = clean_arabic_outcome_line(outcome)
        outcome = re.sub(r'\s+', ' ', outcome).strip()
        if len(outcome) < 14:
            continue
        if len(outcome) > 260:
            outcome = outcome[:260].rsplit(' ', 1)[0].strip()
        if line_has_arabic_label(outcome, '\u0627\u0644\u0645\u062a\u0637\u0644\u0628\u0627\u062a') or line_has_arabic_label(outcome, '\u0627\u0644\u0647\u062f\u0641 \u0627\u0644\u0631\u0626\u064a\u0633'):
            continue
        key = arabic_label_key(outcome)[:30]
        if key in seen:
            continue
        seen.add(key)
        domain = arabic_starter_domain(outcome)
        if not domain:
            continue
        counts[domain] += 1
        clo_id = f"{domain}.{counts[domain]}"
        clo_map[clo_id] = f"{clo_id} {clean_clo_text(outcome)}"
    return clo_map


def extract_arabic_numbered_table_clos(raw_text):
    if not contains_arabic(raw_text):
        return {}

    lines = [compact_text(line) for line in normalize_course_spec_text(raw_text).splitlines()]
    lines = [line for line in lines if line]
    section_lines = []
    in_section = False
    for line in lines:
        if line_has_arabic_label(line, 'نواتج التعلم'):
            if re.search(r'\.{5,}|…{2,}', line):
                continue
            in_section = True
            continue
        if not in_section:
            continue
        if (
            line_has_arabic_label(line, 'موضوعات المقرر')
            or line_has_arabic_label(line, 'مصادر التعلم')
            or line_has_arabic_label(line, 'أنشطة تقييم الطلبة')
            or line_has_arabic_label(line, 'التدريس والتقييم')
            or line_has_arabic_label(line, 'د التدريس')
        ):
            if section_lines:
                break
        section_lines.append(line)

    if not section_lines:
        return {}

    clos = {}
    current_id = ''
    current_parts = []

    def flush_current():
        nonlocal current_id, current_parts
        if not current_id:
            return
        body = clean_clo_text(clean_arabic_outcome_line(' '.join(current_parts)))
        body = re.sub(r'\s+', ' ', body).strip()
        if len(body) >= 8:
            clos[current_id] = f"{current_id} {body}"
        current_id = ''
        current_parts = []

    for line in section_lines:
        normalized_line = compact_text(line)
        domain_header = re.match(r'^\s*([123])\s*[\.\-]\s*0\b', normalized_line)
        if domain_header:
            continue

        match = re.match(r'^\s*([123])\s*[\.\-]\s*([1-9]\d*)\s*(.*)$', normalized_line)
        if match:
            flush_current()
            current_id = f"{match.group(1)}.{match.group(2)}"
            remainder = match.group(3).strip()
            if remainder:
                current_parts.append(remainder)
            continue

        if not current_id:
            continue

        if re.match(r'^[عقم]\s*\d+\b', normalized_line):
            flush_current()
            continue
        if re.search(r'(?:المحاضرة|المناقشة|الاختبارات|الملاحظة|التعلم|الحوار|التعاوني|الذاتي|طرق\s+التقييم|استراتيجيات\s+التدريس)', normalized_line):
            flush_current()
            continue
        if re.match(r'^\d+$', normalized_line):
            continue

        cleaned = clean_arabic_outcome_line(normalized_line)
        if cleaned and len(cleaned) > 2:
            current_parts.append(cleaned)

    flush_current()
    return clos

def arabic_clo_map_is_noisy(clo_map):
    if not clo_map:
        return False
    values = list(clo_map.values())
    joined = ' '.join(values)
    noisy_labels = [
        '\u0627\u0644\u0645\u062a\u0637\u0644\u0628\u0627\u062a',
        '\u0627\u0644\u0647\u062f\u0641 \u0627\u0644\u0631\u0626\u064a\u0633',
        '\u0646\u0645\u0637 \u0627\u0644\u062a\u0639\u0644\u064a\u0645',
        '\u0627\u0644\u0633\u0627\u0639\u0627\u062a \u0627\u0644\u062a\u062f\u0631\u064a\u0633\u064a\u0629',
    ]
    return (
        len(values) <= 1
        and any(line_has_arabic_label(joined, label) for label in noisy_labels)
    ) or any(len(value) > 360 for value in values)

def final_clean_clo_map(clo_map):
    def body_from_clo(value):
        return re.sub(r'^\s*(?:[123]\.\d+|CLO\s*\d+(?:\.\d+)*)\s+', '', str(value or ''), flags=re.I).strip()

    def score_body(value):
        body = body_from_clo(value)
        score = min(len(body), 180)
        for noise in ['م لك', 'الي توصل م إلها', 'تعرض التوازل', 'الفقبية', 'الواجنا']:
            if noise in body:
                score -= 35
        if 'مستخدما التقنية الحديثة' in body:
            score += 15
        if 'والمشروعات البحثية والعلمية' in body:
            score += 15
        return score

    grouped = {}
    order = []
    for key, value in (clo_map or {}).items():
        number = clo_number(value) or str(key or '').strip()
        domain = number.split('.', 1)[0] if '.' in number else ''
        if domain not in {'1', '2', '3'}:
            domain = arabic_starter_domain(body_from_clo(value)) or domain
        body = body_from_clo(value)
        if not body:
            continue
        dedupe_key = f"{domain}:{arabic_label_key(body)[:22]}"
        if dedupe_key not in grouped:
            order.append(dedupe_key)
            grouped[dedupe_key] = {'domain': domain, 'body': body, 'score': score_body(value)}
        elif score_body(value) > grouped[dedupe_key]['score']:
            grouped[dedupe_key] = {'domain': domain, 'body': body, 'score': score_body(value)}

    counts = {'1': 0, '2': 0, '3': 0}
    cleaned = {}
    for domain in ['1', '2', '3']:
        for dedupe_key in order:
            item = grouped.get(dedupe_key)
            if not item or item['domain'] != domain:
                continue
            counts[domain] += 1
            clo_id = f"{domain}.{counts[domain]}"
            cleaned[clo_id] = f"{clo_id} {clean_clo_text(item['body'])}"
    return cleaned

def add_missing_arabic_review_clo(clo_map, raw_text):
    if any(value.startswith('1.2 ') for value in (clo_map or {}).values()):
        return clo_map
    text = normalize_course_spec_text(raw_text)
    if re.search(r'(?:يستعرض|تعرض)\s+النوازل\s+الفقهية', text) and re.search(r'العبادات', text):
        existing_knowledge = [
            value for value in (clo_map or {}).values()
            if str(value).startswith('1.')
        ]
        if len(existing_knowledge) == 1:
            clo_map = dict(clo_map or {})
            clo_map['1.2'] = '1.2 يستعرض النوازل الفقهية في العبادات والأبحاث والآراء المتعلقة بها'
    return clo_map

def extract_course_spec_metadata(text, source_filename=''):
    text = normalize_course_spec_text(text)
    lines = [compact_text(line) for line in (text or '').splitlines()]
    lines = [line for line in lines if line]
    raw_text = re.sub(r'[\x00-\x1f]+W\b', ' ', text or '')
    raw_text = re.sub(r'[\x00-\x1f]+', ' ', raw_text)
    normalized = compact_text(raw_text)

    course_name = value_after_label(lines, ['Course Name', 'Course Title', 'Course'])
    course_code = value_after_label(lines, ['Course Code', 'Course Number', 'Course ID', 'Course No'])
    college = value_after_label(lines, ['College', 'Faculty', 'School'])
    department = value_after_label(lines, ['Department', 'Dept', 'Academic Department', 'Scientific Department'])
    program = value_after_label(lines, ['Program', 'Programme', 'Program Name', 'Academic Program'])

    arabic_course_name = value_after_arabic_label(lines, [
        'اسمالمقرر',
        'اسم المقرر',
        'عنوان المقرر',
        'اسم المقرر الدراسي',
        'اسم مقرر'
    ]) or value_after_label_loose(lines, [
        'اسمالمقرر',
        'اسم المقرر',
        'عنوان المقرر',
        'اسم المقرر الدراسي',
        'اسم مقرر'
    ])
    arabic_course_code = value_after_arabic_label(lines, [
        'رمزالمقرر',
        'رمز المقرر',
        'رمزورقم المقرر',
        'رمزالمقرر',
        'رمز ورقم المقرر',
        'رقم المقرر',
        'كود المقرر',
        'رمز المقرر الدراسي'
    ]) or value_after_label_loose(lines, [
        'رمزالمقرر',
        'رمز المقرر',
        'رمز ورقم المقرر',
        'رقم المقرر',
        'كود المقرر',
        'رمز المقرر الدراسي'
    ])
    arabic_college = value_after_arabic_label(lines, ['الكلية']) or value_after_label_loose(lines, ['الكلية', 'Ø§Ù„ÙƒÙ„ÙŠØ©'])
    arabic_department = value_after_arabic_label(lines, ['القسم العلمي', 'القسم']) or value_after_label_loose(lines, ['القسم', 'القسم العلمي', 'Ø§Ù„Ù‚Ø³Ù…', 'Ø§Ù„Ù‚Ø³Ù… Ø§Ù„Ø¹Ù„Ù…ÙŠ'])
    arabic_program = value_after_arabic_label(lines, ['اسم البرنامج', 'البرنامج']) or value_after_label_loose(lines, ['البرنامج', 'اسم البرنامج', 'Ø§Ù„Ø¨Ø±Ù†Ø§Ù…Ø¬', 'Ø§Ø³Ù… Ø§Ù„Ø¨Ø±Ù†Ø§Ù…Ø¬'])
    if arabic_course_name and not course_name:
        course_name = clean_arabic_metadata_value(arabic_course_name, 'course_name')
    if arabic_course_code and not course_code:
        course_code = normalize_extracted_course_code(arabic_course_code)
    if arabic_college and not college:
        college = clean_arabic_metadata_value(arabic_college, 'college')
    if arabic_department and not department:
        department = clean_arabic_metadata_value(arabic_department, 'department')
    if arabic_program and not program:
        program = clean_arabic_metadata_value(arabic_program, 'program')

    if contains_arabic(raw_text):
        arabic_code_match = re.search(
            r'(?:\u0631\u0645\u0632|\u0631\u0642\u0645|\u0643\u0648\u062f)\s*'
            r'(?:\u0648\s*\u0631\u0642\u0645\s*)?'
            r'\u0627\u0644\s*\u0645\u0642\u0631\u0631\s*[:\-–—،؛]?\s*'
            r'([A-Za-z]{0,8}\s*\d(?:\s*\d){2,9}[A-Za-z]?)',
            normalized
        )
        if arabic_code_match:
            course_code = normalize_extracted_course_code(arabic_code_match.group(1))
        if is_bad_arabic_course_name(course_name):
            filename_course_name = infer_course_name_from_filename(source_filename)
            if filename_course_name:
                course_name = filename_course_name

    title_match = re.search(
        rf'{flexible_label("Course Title")}\s*[:\-\u061b]?\s*(.+?)\s+{flexible_label("Course Code")}',
        raw_text,
        flags=re.I | re.S
    )
    if title_match:
        course_name = clean_pdf_fragment(title_match.group(1))
        if course_name.islower():
            course_name = title_case_course_name(course_name)

    code_match = re.search(
        rf'{flexible_label("Course Code")}\s*[:\-\u061b]?\s*(?:\S\s*){{0,12}}?([A-Z](?:\s*[A-Z]){{1,5}}\s*\d(?:\s*\d){{2,3}}[A-Z]?|[A-Z]{{2,5}}\s*\d{{3,4}}[A-Z]?)',
        raw_text,
        flags=re.I | re.S
    )
    if code_match:
        course_code = normalize_extracted_course_code(code_match.group(1))

    if not course_code:
        code_match = re.search(r'\b([A-Z]{2,5}\s*\d{3,4}[A-Z]?)\b', normalized)
        if code_match:
            course_code = normalize_extracted_course_code(code_match.group(1))

    if not course_name:
        title_match = re.search(r'(?i)(?:Course\s+(?:Name|Title)\s*[:\-]?\s*)(.{4,120}?)(?=\s+(?:Course\s+(?:Code|Number|ID)|Credit|Prerequisite|$))', normalized)
        if title_match:
            course_name = title_match.group(1).strip()

    clo_map = {}
    line_text = '\n'.join(lines)
    for match in re.finditer(r'(?m)^\s*((?:[123]\.\d+|CLO\s*\d+))\s+(.+)$', line_text, flags=re.I):
        clo_id = re.sub(r'\s+', '', match.group(1).upper())
        if re.match(r'^[123]\.0$', clo_id):
            continue
        clo_body = clean_clo_text(clean_pdf_fragment(match.group(2)))
        if clo_body and len(clo_body) > 8:
            clo_map[clo_id] = f"{clo_id} {clo_body}"

    section = extract_course_spec_section(raw_text, 'Course Learning Outcomes', 'Course Content')
    if section:
        for match in re.finditer(r'\b([123]\s*\.\s*[1-9]\d*)\s+(.+?)\s+\b([KSV]\s*\d+)\b', section, flags=re.I | re.S):
            clo_id = re.sub(r'\s+', '', match.group(1))
            clo_body = clean_clo_text(clean_pdf_fragment(match.group(2)))
            if clo_body and len(clo_body) > 8:
                clo_map[clo_id] = f"{clo_id} {clo_body}"

    if not clo_map:
        for match in re.finditer(r'\b([123]\.\d+)\s+(.{12,220}?)(?=\s+[123]\.\d+\s+|\s+CLO\s*\d+\s+|$)', normalized, flags=re.I):
            clo_id = match.group(1)
            if re.match(r'^[123]\.0$', clo_id):
                continue
            clo_body = clean_clo_text(clean_pdf_fragment(match.group(2)))
            if clo_body:
                clo_map[clo_id] = f"{clo_id} {clo_body}"

    if contains_arabic(raw_text):
        arabic_stop_labels = (
            r'المعرفة|الفهم|المهارات|القيم|استراتيجيات|طرق\s+التدريس|'
            r'طرق\s+التقييم|أساليب\s+التقييم|محتوى\s+المقرر|موضوعات\s+المقرر|'
            r'قائمة\s+الموضوعات|نواتج\s+التعلم|رمز\s+المقرر|اسم\s+المقرر'
        )
        arabic_clo_pattern = re.compile(
            rf'(?:^|\s)([123]\s*[\.\-]\s*[1-9]\d*)\s+(.+?)(?=\s+[123]\s*[\.\-]\s*[1-9]\d*\s+|\s+(?:{arabic_stop_labels})\b|$)',
            flags=re.S
        )
        for match in arabic_clo_pattern.finditer(normalized):
            clo_id = re.sub(r'\s+', '', match.group(1)).replace('-', '.')
            if re.match(r'^[123]\.0$', clo_id):
                continue
            clo_body = clean_clo_text(clean_pdf_fragment(match.group(2)))
            clo_body = re.sub(r'\b(?:K|S|V)\s*\d+\b.*$', '', clo_body, flags=re.I).strip()
            if clo_body and len(clo_body) > 6:
                clo_map.setdefault(clo_id, f"{clo_id} {clo_body}")

        if not course_name:
            title_match = re.search(r'(?:اسم\s+المقرر|عنوان\s+المقرر)\s*[:\-–—،؛]?\s*(.{3,120}?)(?=\s+(?:رمز|رقم|كود|الساعات|المستوى|المتطلبات|$))', normalized)
            if title_match:
                course_name = clean_pdf_fragment(title_match.group(1))

        if not course_code:
            code_match = re.search(r'(?:رمز|رقم|كود)(?:\s+ورقم)?\s*المقرر\s*[:\-–—،؛]?\s*([A-Za-z]{0,8}\s*\d{3,10}[A-Za-z]?|\d{3,10})', normalized)
            if code_match:
                course_code = normalize_extracted_course_code(code_match.group(1))

        if not clo_map:
            clo_map.update(extract_noisy_arabic_ocr_clos(text))

        arabic_table_clos = extract_arabic_table_clos(text)
        arabic_numbered_clos = extract_arabic_numbered_table_clos(text)
        if arabic_numbered_clos and len(arabic_numbered_clos) > len(arabic_table_clos):
            arabic_table_clos = arabic_numbered_clos
        if arabic_table_clos and (arabic_clo_map_is_noisy(clo_map) or len(arabic_table_clos) > len(clo_map)):
            clo_map = arabic_table_clos
        elif arabic_table_clos:
            clo_map.update({key: value for key, value in arabic_table_clos.items() if key not in clo_map})

        if (not course_name or course_name in {'رمز المقرر', 'البرنامج', 'القسم العلمي'} or 'رمز المقرر' in course_name) and clo_map:
            inferred_course_name = infer_arabic_course_name_from_ocr(text)
            if inferred_course_name:
                course_name = inferred_course_name
        if is_bad_arabic_course_name(course_name):
            filename_course_name = infer_course_name_from_filename(source_filename)
            if filename_course_name:
                course_name = filename_course_name
        if course_code and any(label in course_code for label in ['البرنامج', 'القسم', 'الكلية']):
            course_code = ''

    course_code = prefer_alphanumeric_course_code(text, course_code)
    clo_map = add_missing_arabic_review_clo(clo_map, text)
    if not clo_map:
        clo_map.update(extract_legacy_course_goal_clos(text))
    clo_map = final_clean_clo_map(clo_map)
    clos = list(clo_map.values())
    clo_plos = extract_clo_plo_mapping(section or raw_text, clos)
    if contains_arabic(raw_text):
        arabic_clo_plos = extract_clo_plo_mapping(text, clos)
        if arabic_clo_plos:
            clo_plos.update(arabic_clo_plos)
    display_name = course_name
    if course_code and course_name and course_code not in course_name:
        display_name = f"{course_name} ({course_code})"
    elif course_code and not course_name:
        display_name = course_code

    return {
        'name': display_name,
        'course_name': course_name,
        'course_code': course_code,
        'course_number': course_code,
        'college': college,
        'department': department,
        'program': program,
        'clos': clos,
        'topics': extract_course_topics(text),
        'clo_plos': clo_plos,
        'grouped_clos': group_clos_by_domain(clos)
    }

def question_number_from_label(label):
    label = str(label or '').strip()
    patterns = [
        r'(?:السؤال|سؤال|س)\s*[-#:.]?\s*(\d{1,3})',
        r'\bQ(?:uestion)?\s*[-#:.]?\s*(\d{1,3})\b',
        r'\bQuestion\s+No\.?\s*(\d{1,3})\b',
        r'\bAnswers?\s*[-_ #:.]?\s*(\d{1,3})\b',
        r'\bItems?\s*[-_ #:.]?\s*(\d{1,3})\b'
    ]
    for pattern in patterns:
        match = re.search(pattern, label, flags=re.I)
        if match:
            return int(match.group(1))
    return None

def find_question_header_row(df):
    best = None
    for row_position, (_, row) in enumerate(df.iterrows()):
        question_cells = []
        for col_index, value in row.items():
            if pd.isna(value):
                continue
            number = question_number_from_label(value)
            if number:
                question_cells.append((col_index, number, str(value).strip()))
        if best is None or len(question_cells) > len(best['question_cells']):
            best = {'row_position': row_position, 'question_cells': question_cells}
    return best if best and len(best['question_cells']) >= 2 else None

def normalize_report_label(value):
    text = '' if pd.isna(value) else str(value).strip()
    text = re.sub(r'[\u064b-\u065f\u0670]', '', text)
    text = text.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    text = text.replace('ؤ', 'و').replace('ئ', 'ي').replace('ة', 'ه')
    return re.sub(r'\s+', ' ', text).strip().lower()

def is_answer_key_label(value):
    label = normalize_report_label(value)
    compact = re.sub(r'\s+', '', label)
    return bool(
        re.search(r'answer\s*key|key\s*answers?', label, flags=re.I)
        or ('مفتاح' in label and ('اجابه' in label or 'اجابات' in label))
        or compact in {'مفتاحالاجابه', 'مفتاحالاجابات'}
    )

def is_summary_row_label(value):
    label = normalize_report_label(value)
    if not label:
        return False
    if re.search(r'^(mean|average|median|total|subtotal|score|percent|percentage)$', label, flags=re.I):
        return True
    return any(term in label for term in [
        'متوسط',
        'المتوسط',
        'اجمالي',
        'الاجمالي',
        'نسبه ميويه',
        'النسبه الميويه',
        'الدرجه الوسيطه',
        'احصاءات'
    ])

def get_question_sheet_rows(df, header_info):
    answer_columns = [col for col, _, _ in header_info['question_cells']]
    first_answer_column = min(answer_columns) if answer_columns else 0
    answer_key = None
    student_rows = []

    for _, row in df.iloc[header_info['row_position'] + 1:].iterrows():
        first_value = '' if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        populated_answers = sum(
            0 if pd.isna(row[col]) or str(row[col]).strip() == '' else 1
            for col in answer_columns
        )
        if is_answer_key_label(first_value):
            answer_key = row
            continue
        if is_summary_row_label(first_value):
            continue

        has_student_metadata = any(
            not pd.isna(row.iloc[col]) and str(row.iloc[col]).strip() != ''
            for col in range(first_answer_column)
        )
        if populated_answers == 0 and not has_student_metadata:
            continue
        student_rows.append(row)

    return answer_key, student_rows

def count_students_from_question_sheet(df, header_info):
    _, student_rows = get_question_sheet_rows(df, header_info)
    return len(student_rows)

def summarize_question_sheet_performance(df, header_info):
    answer_key, student_rows = get_question_sheet_rows(df, header_info)
    if answer_key is None or not student_rows:
        return {}

    performance = {}
    for col_index, number, _ in header_info['question_cells']:
        question = f'Q{number}'
        correct_answer = normalize_answer(answer_key[col_index])
        if not correct_answer:
            continue

        correct_count = sum(
            1
            for row in student_rows
            if normalize_answer(row[col_index]) == correct_answer
        )
        performance[question] = {
            'students_answered': len(student_rows),
            'students_correct': correct_count,
            'correct_percentage': round((correct_count / len(student_rows)) * 100, 2)
        }
    return performance

def find_answer_key_mapping_table(df):
    clean_df = df.dropna(how='all')
    if clean_df.empty:
        return None

    for row_position, (_, row) in enumerate(clean_df.iterrows()):
        normalized_headers = {
            normalize_report_label(value): col_index
            for col_index, value in row.items()
            if not pd.isna(value) and str(value).strip()
        }
        question_col = next((col for label, col in normalized_headers.items() if label in {'question number', 'question no', 'question'}), None)
        point_col = next((col for label, col in normalized_headers.items() if label in {'point value', 'points', 'point', 'score'}), None)
        tags_col = next((col for label, col in normalized_headers.items() if label in {'tags', 'tag', 'clo', 'clos'}), None)
        response_col = next((col for label, col in normalized_headers.items() if label in {'response/mapping', 'response', 'mapping', 'answer'}), None)

        if question_col is not None and (point_col is not None or tags_col is not None):
            return {
                'row_position': row_position,
                'question_col': question_col,
                'point_col': point_col,
                'tags_col': tags_col,
                'response_col': response_col,
            }
    return None

def split_clo_tags(value):
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    tags = re.split(r'[,;/|]+', text)
    return [tag.strip() for tag in tags if tag.strip()]

def clo_number_key(value):
    text = str(value or '').strip()
    clo_match = re.search(r'\bCLO\s*[-_ ]?(\d+(?:\.\d+)*)', text, flags=re.I)
    if clo_match:
        return clo_match.group(1)
    leading_match = re.match(r'^\s*(\d+(?:\.\d+)*)\b', text)
    if leading_match:
        return leading_match.group(1)
    return ''

def resolve_detected_clos_to_course_list(detected_clos, course_clos):
    by_text = {str(clo).strip().lower(): clo for clo in course_clos or []}
    numbered_clos = []
    for clo in course_clos or []:
        key = clo_number_key(clo)
        if key:
            numbered_clos.append((key, clo))

    resolved = []
    for detected_clo in detected_clos or []:
        detected_text = str(detected_clo).strip()
        detected_key = clo_number_key(detected_text)
        exact_text_match = by_text.get(detected_text.lower())
        if exact_text_match and exact_text_match not in resolved:
            resolved.append(exact_text_match)
            continue

        if not detected_key:
            continue

        matches = [
            clo
            for key, clo in numbered_clos
            if key == detected_key or key.startswith(f"{detected_key}.")
        ]
        for matched in matches:
            if matched not in resolved:
                resolved.append(matched)
    return resolved

def get_column_by_alias(df, aliases):
    alias_set = {normalize_report_label(alias) for alias in aliases}
    compact_alias_set = {re.sub(r'\s+', '', alias) for alias in alias_set}
    for column in df.columns:
        label = normalize_report_label(column)
        compact_label = re.sub(r'\s+', '', label)
        if label in alias_set or compact_label in compact_alias_set:
            return column
    return None

def get_column_by_flexible_alias(df, aliases):
    exact = get_column_by_alias(df, aliases)
    if exact is not None:
        return exact

    alias_labels = [normalize_report_label(alias) for alias in aliases]
    compact_aliases = [re.sub(r'\s+', '', alias) for alias in alias_labels]
    for column in df.columns:
        label = normalize_report_label(column)
        compact_label = re.sub(r'\s+', '', label)
        for alias, compact_alias in zip(alias_labels, compact_aliases):
            if len(compact_alias) < 3:
                continue
            if compact_alias in compact_label:
                return column
            if re.search(rf'\b{re.escape(alias)}\b', label):
                return column
    return None

def get_tag_class_columns(df):
    columns = {
        'tag': get_column_by_alias(df, ['Tag', 'Tags', 'CLO', 'CLOs']),
        'student_id': get_column_by_alias(df, ['StudentID', 'Student ID', 'StudentExternalID', 'Student External ID']),
        'tag_type': get_column_by_alias(df, ['TagType', 'Tag Type']),
        'question_number': get_column_by_alias(df, ['QuestionNumber', 'Question Number']),
        'earned_points': get_column_by_alias(df, ['EarnedPoints', 'Earned Points']),
        'possible_points': get_column_by_alias(df, ['PossiblePoints', 'Possible Points']),
    }
    required = ['tag', 'student_id', 'tag_type', 'question_number', 'earned_points', 'possible_points']
    return columns if all(columns.get(key) is not None for key in required) else None

def get_tag_class_question_rows(df, columns):
    rows = df.copy()
    tag_type = rows[columns['tag_type']].astype(str).str.strip().str.lower()
    rows = rows[tag_type == 'question'].copy()
    rows['_question_number'] = pd.to_numeric(rows[columns['question_number']], errors='coerce')
    rows = rows.dropna(subset=['_question_number'])
    rows['_question_number'] = rows['_question_number'].astype(int)
    rows = rows[rows['_question_number'] > 0]
    return rows

def normalize_grade_item_label(value, fallback_prefix='Item'):
    text = '' if pd.isna(value) else str(value).strip()
    if not text:
        return ''
    number = question_number_from_label(text)
    if number is None:
        try:
            if re.match(r'^\d+(?:\.0+)?$', text):
                number = int(float(text))
        except ValueError:
            number = None
    if number is not None:
        return f'Q{number}'
    return re.sub(r'\s+', ' ', text)

def get_generic_long_grade_columns(df):
    columns = {
        'student_id': get_column_by_flexible_alias(df, [
            'StudentID', 'Student ID', 'StudentExternalID', 'Student External ID',
            'Student Number', 'Student No', 'Student', 'Students', 'ID', 'IDs',
            'Learner ID', 'Name', 'Student Name'
        ]),
        'clo': get_column_by_flexible_alias(df, [
            'CLO', 'CLOs', 'Course Learning Outcome', 'Course Learning Outcomes',
            'Learning Outcome', 'Learning Outcomes', 'Learning Objective',
            'Outcome', 'Outcomes', 'LO'
        ]),
        'question': get_column_by_flexible_alias(df, [
            'QuestionNumber', 'Question Number', 'Question No', 'Question',
            'Questions', 'Item Number', 'Item', 'Q'
        ]),
        'earned_points': get_column_by_flexible_alias(df, [
            'EarnedPoints', 'Earned Points', 'Score', 'Student Score', 'Grade',
            'Mark', 'Marks', 'Points Earned', 'Obtained Points', 'Earned',
            'Result', 'Points'
        ]),
        'possible_points': get_column_by_flexible_alias(df, [
            'PossiblePoints', 'Possible Points', 'Max Score', 'Maximum Score',
            'Max Points', 'Possible', 'Out Of', 'Total Points', 'Point Value',
            'Full Mark', 'Full Score', 'Total'
        ]),
    }

    if not columns['student_id'] or not columns['earned_points']:
        return None
    if not columns['question'] and not columns['clo']:
        return None
    return columns

def get_generic_long_grade_rows(df, columns):
    rows = df.copy()
    rows['_student_id'] = rows[columns['student_id']].map(normalize_student_id)
    rows['_earned_points'] = pd.to_numeric(rows[columns['earned_points']], errors='coerce')
    rows = rows[(rows['_student_id'] != '') & rows['_earned_points'].notna()].copy()
    if rows.empty:
        return rows

    if columns.get('possible_points'):
        rows['_possible_points'] = pd.to_numeric(rows[columns['possible_points']], errors='coerce')
    else:
        rows['_possible_points'] = pd.NA

    if columns.get('question'):
        rows['_question_label'] = rows[columns['question']].map(normalize_grade_item_label)
    else:
        rows['_question_label'] = rows[columns['clo']].map(lambda value: normalize_grade_item_label(value, 'CLO'))

    rows = rows[rows['_question_label'] != ''].copy()
    return rows

def infer_generic_long_grade_metrics(df):
    columns = get_generic_long_grade_columns(df)
    if not columns:
        return None

    rows = get_generic_long_grade_rows(df, columns)
    if rows.empty:
        return None

    questions = sorted(rows['_question_label'].unique(), key=lambda question: (0, int(question[1:])) if re.match(r'^Q\d+$', str(question)) else (1, str(question)))
    max_scores = {}
    detected_clo_mappings = {}
    performance = {}

    for question in questions:
        question_rows = rows[rows['_question_label'] == question]
        possible_values = pd.to_numeric(question_rows['_possible_points'], errors='coerce').dropna()
        earned_values = pd.to_numeric(question_rows['_earned_points'], errors='coerce').dropna()
        max_score = float(possible_values.max()) if not possible_values.empty else (float(earned_values.max()) if not earned_values.empty else 1.0)
        max_scores[question] = max_score or 1.0

        if columns.get('clo'):
            tags = []
            for value in question_rows[columns['clo']].dropna().unique():
                tags.extend(split_clo_tags(value))
            if tags:
                detected_clo_mappings[question] = sorted(set(tags))

        deduped = question_rows.sort_values('_earned_points').drop_duplicates('_student_id', keep='last')
        answered = len(deduped)
        achieved = int((deduped['_earned_points'] >= max_scores[question]).sum()) if answered else 0
        performance[question] = {
            'students_answered': answered,
            'students_correct': achieved,
            'correct_percentage': round((achieved / answered) * 100, 2) if answered else 0
        }

    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': len(set(rows['_student_id'])),
        'confidence': 'High',
        'text_sample': 'Detected generic grade table with student IDs, score columns, and question/CLO/learning outcome columns.',
        'max_scores': max_scores,
        'question_performance': performance,
        'detected_clo_mappings': detected_clo_mappings
    }

def infer_tag_class_metrics(df):
    columns = get_tag_class_columns(df)
    if not columns:
        return None

    question_rows = get_tag_class_question_rows(df, columns)
    if question_rows.empty:
        return None

    question_numbers = sorted(question_rows['_question_number'].unique())
    questions = [f'Q{number}' for number in question_numbers]
    max_scores = {}
    detected_clo_mappings = {}
    performance = {}

    for number in question_numbers:
        question = f'Q{number}'
        rows = question_rows[question_rows['_question_number'] == number]
        possible_values = pd.to_numeric(rows[columns['possible_points']], errors='coerce').dropna()
        max_scores[question] = float(possible_values.max()) if not possible_values.empty else 1.0

        tags = []
        for value in rows[columns['tag']].dropna().unique():
            tags.extend(split_clo_tags(value))
        detected_clo_mappings[question] = sorted(set(tags))

        deduped = rows.copy()
        deduped['_student_id'] = deduped[columns['student_id']].map(normalize_student_id)
        deduped['_earned_points'] = pd.to_numeric(deduped[columns['earned_points']], errors='coerce').fillna(0)
        deduped['_possible_points'] = pd.to_numeric(deduped[columns['possible_points']], errors='coerce').fillna(max_scores[question])
        deduped = deduped.sort_values('_earned_points').drop_duplicates('_student_id', keep='last')
        answered = len(deduped)
        achieved = int((deduped['_earned_points'] >= deduped['_possible_points']).sum()) if answered else 0
        performance[question] = {
            'students_answered': answered,
            'students_correct': achieved,
            'correct_percentage': round((achieved / answered) * 100, 2) if answered else 0
        }

    student_ids = question_rows[columns['student_id']].map(normalize_student_id)
    total_students = len({student_id for student_id in student_ids if student_id})

    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': total_students,
        'confidence': 'High',
        'text_sample': 'Detected long-form tag class data with student question scores, possible points, and CLO tags.',
        'max_scores': max_scores,
        'question_performance': performance,
        'detected_clo_mappings': detected_clo_mappings
    }

def infer_answer_key_mapping_metrics(df):
    table_info = find_answer_key_mapping_table(df)
    if not table_info:
        return None

    rows = df.dropna(how='all').iloc[table_info['row_position'] + 1:]
    questions = []
    max_scores = {}
    detected_clo_mappings = {}

    for _, row in rows.iterrows():
        question_raw = row[table_info['question_col']]
        if pd.isna(question_raw) or str(question_raw).strip() == '':
            continue

        question_number = question_number_from_label(question_raw)
        if question_number is None:
            try:
                question_number = int(float(str(question_raw).strip()))
            except ValueError:
                continue

        question = f'Q{question_number}'
        questions.append(question)

        if table_info['point_col'] is not None:
            point_value = pd.to_numeric(row[table_info['point_col']], errors='coerce')
            max_scores[question] = float(point_value) if not pd.isna(point_value) else 1.0
        else:
            max_scores[question] = 1.0

        if table_info['tags_col'] is not None:
            tags = split_clo_tags(row[table_info['tags_col']])
            if tags:
                detected_clo_mappings[question] = tags

    questions = sorted(set(questions), key=lambda question: int(question[1:]))
    if not questions:
        return None

    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': 0,
        'confidence': 'High',
        'text_sample': 'Detected answer-key mapping table with point values and CLO tags. Upload student score/response files as well to calculate student achievement.',
        'max_scores': max_scores,
        'question_performance': {},
        'detected_clo_mappings': detected_clo_mappings
    }

def infer_simple_table_metrics(df):
    clean_df = df.dropna(how='all')
    if clean_df.empty:
        return None

    tag_class_metrics = infer_tag_class_metrics(clean_df)
    if tag_class_metrics:
        return tag_class_metrics

    generic_long_metrics = infer_generic_long_grade_metrics(clean_df)
    if generic_long_metrics:
        return generic_long_metrics

    answer_key_metrics = infer_answer_key_mapping_metrics(clean_df)
    if answer_key_metrics:
        return answer_key_metrics

    header_info = find_question_header_row(clean_df)
    if header_info:
        questions = [f'Q{number}' for _, number, _ in sorted(header_info['question_cells'], key=lambda item: item[1])]
        total_students = count_students_from_question_sheet(clean_df, header_info)
        performance = summarize_question_sheet_performance(clean_df, header_info)
        return {
            'questions': questions,
            'total_questions': len(questions),
            'total_students': total_students,
            'confidence': 'High',
            'text_sample': '',
            'max_scores': {question: 1.0 for question in questions},
            'question_performance': performance
        }

    df_with_headers = clean_df.copy()
    df_with_headers.columns = [str(col).strip() for col in df_with_headers.iloc[0]]
    df_with_headers = df_with_headers.iloc[1:].dropna(how='all')
    question_columns = []
    for col_position, col in enumerate(df_with_headers.columns):
        number = question_number_from_label(col)
        if number:
            question_columns.append((col_position, col, number))

    if question_columns:
        questions = [f'Q{number}' for _, _, number in sorted(question_columns, key=lambda item: item[2])]
        max_scores = {}
        for col_position, _, number in question_columns:
            values = pd.to_numeric(df_with_headers.iloc[:, col_position], errors='coerce').dropna()
            max_scores[f'Q{number}'] = float(values.max()) if not values.empty else 1.0
        return {
            'questions': questions,
            'total_questions': len(questions),
            'total_students': len(df_with_headers),
            'confidence': 'Medium',
            'text_sample': '',
            'max_scores': max_scores
        }

    numeric_df = df_with_headers.apply(pd.to_numeric, errors='coerce')
    numeric_positions = [
        col_position
        for col_position in range(numeric_df.shape[1])
        if not numeric_df.iloc[:, col_position].dropna().empty
    ]
    if numeric_positions:
        questions = [str(df_with_headers.columns[col_position]) for col_position in numeric_positions]
        max_scores = {}
        for col_position in numeric_positions:
            question = str(df_with_headers.columns[col_position])
            values = numeric_df.iloc[:, col_position].dropna()
            max_scores[question] = float(values.max()) if not values.empty else 1.0
        return {
            'questions': questions,
            'total_questions': len(questions),
            'total_students': len(df_with_headers),
            'confidence': 'Low',
            'text_sample': '',
            'max_scores': max_scores
        }
    return None

def infer_spreadsheet_metrics(filepath, file_ext):
    if file_ext == '.pdf':
        text = extract_pdf_text(filepath)
        metrics = infer_pdf_grade_metrics(text)
        if metrics:
            return metrics

    if file_ext == '.csv':
        df_with_headers = pd.read_csv(filepath)
        tag_class_metrics = infer_tag_class_metrics(df_with_headers)
        if tag_class_metrics:
            tag_class_metrics['text_sample'] = f"Detected tag class data from CSV. Rows: {df_with_headers.shape[0]}, columns: {df_with_headers.shape[1]}."
            return tag_class_metrics
        generic_long_metrics = infer_generic_long_grade_metrics(df_with_headers)
        if generic_long_metrics:
            generic_long_metrics['text_sample'] = f"Detected generic grade data from CSV. Rows: {df_with_headers.shape[0]}, columns: {df_with_headers.shape[1]}."
            return generic_long_metrics

        df = pd.read_csv(filepath, header=None)
        metrics = infer_simple_table_metrics(df)
        if metrics:
            metrics['text_sample'] = f"Detected from CSV. Rows: {df.shape[0]}, columns: {df.shape[1]}."
            return metrics
    else:
        workbook = pd.ExcelFile(filepath)
        best_metrics = None
        best_sheet = None
        for sheet_name in workbook.sheet_names:
            df_with_headers = pd.read_excel(filepath, sheet_name=sheet_name)
            tag_class_metrics = infer_tag_class_metrics(df_with_headers)
            if tag_class_metrics and (best_metrics is None or tag_class_metrics['total_questions'] > best_metrics['total_questions']):
                best_metrics = tag_class_metrics
                best_sheet = sheet_name
                continue
            generic_long_metrics = infer_generic_long_grade_metrics(df_with_headers)
            if generic_long_metrics and (best_metrics is None or generic_long_metrics['total_questions'] > best_metrics['total_questions']):
                best_metrics = generic_long_metrics
                best_sheet = sheet_name
                continue

            df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            metrics = infer_simple_table_metrics(df)
            if metrics and (best_metrics is None or metrics['total_questions'] > best_metrics['total_questions']):
                best_metrics = metrics
                best_sheet = sheet_name
        if best_metrics:
            best_metrics['text_sample'] = f"Detected from sheet: {best_sheet}. Questions were read from answer/question columns."
            return best_metrics

    return {
        'questions': [],
        'total_questions': 0,
        'total_students': 0,
        'confidence': 'Low',
        'text_sample': 'No spreadsheet question columns were detected.',
        'max_scores': {}
    }

PDF_GRADE_ROW_PATTERN = re.compile(
    r'(?:^|\s)(\d{1,3})\s+.*?\s+(-?\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)(?=\s+\d{1,3}\s+|$)'
)

def parse_pdf_grade_report_text(text):
    rows = []
    max_scores = {}
    for line in (text or '').splitlines():
        if 'Stu Pts Poss' not in line:
            continue

        body = line.split('Stu Pts Poss', 1)[1]
        score_row = {}
        for match in PDF_GRADE_ROW_PATTERN.finditer(body):
            number = int(match.group(1))
            score = float(match.group(2))
            possible = float(match.group(3))
            if number <= 0:
                continue
            question = f'Q{number}'
            score_row[question] = score
            max_scores[question] = max(max_scores.get(question, 0.0), possible)

        if len(score_row) >= 2:
            rows.append(score_row)

    if not rows:
        rows, max_scores = parse_tokenized_pdf_grade_report_text(text)

    if not rows:
        return None, {}

    questions = sorted(max_scores.keys(), key=lambda question: int(question[1:]))
    score_df = pd.DataFrame(rows).reindex(columns=questions).fillna(0)
    score_df = apply_student_id_index(score_df, [f"__missing_student_{index + 1}" for index in range(len(score_df))])
    return score_df, max_scores


def is_ascii_integer_token(value):
    return bool(re.fullmatch(r'[0-9]{1,3}', str(value or '').strip()))


def is_ascii_number_token(value):
    return bool(re.fullmatch(r'-?[0-9]+(?:\.[0-9]+)?', str(value or '').strip()))


def parse_tokenized_pdf_grade_report_text(text):
    tokens = [line.strip() for line in (text or '').splitlines() if line.strip()]
    rows = []
    max_scores = {}
    current_row = {}
    last_question_number = 0
    in_score_table = False
    index = 0

    def append_current_row():
        nonlocal current_row, last_question_number
        if current_row:
            rows.append(current_row)
        current_row = {}
        last_question_number = 0

    while index < len(tokens):
        token = tokens[index]
        next_one = tokens[index + 1] if index + 1 < len(tokens) else ''
        next_two = tokens[index + 2] if index + 2 < len(tokens) else ''

        if token == 'Stu' and next_one == 'Pts' and next_two == 'Poss':
            in_score_table = True
            index += 3
            continue

        if not in_score_table:
            index += 1
            continue

        if (
            is_ascii_integer_token(token)
            and index + 4 < len(tokens)
            and is_ascii_number_token(tokens[index + 3])
            and is_ascii_number_token(tokens[index + 4])
        ):
            question_number = int(token)
            score = float(tokens[index + 3])
            possible = float(tokens[index + 4])
            if current_row and question_number <= last_question_number:
                append_current_row()
            question = f'Q{question_number}'
            current_row[question] = score
            max_scores[question] = max(max_scores.get(question, 0.0), possible)
            last_question_number = question_number
            index += 5
            continue

        index += 1

    append_current_row()
    return rows, max_scores

def infer_pdf_grade_metrics(text):
    score_df, max_scores = parse_pdf_grade_report_text(text)
    if score_df is None or score_df.empty:
        return None

    questions = list(score_df.columns)
    performance = {}
    total_students = len(score_df)
    for question in questions:
        possible = max_scores.get(question, 1.0) or 1.0
        values = pd.to_numeric(score_df[question], errors='coerce').fillna(0)
        correct_count = int((values >= possible).sum())
        performance[question] = {
            'students_answered': total_students,
            'students_correct': correct_count,
            'correct_percentage': round((correct_count / total_students) * 100, 2) if total_students else 0
        }

    return {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': total_students,
        'confidence': 'High',
        'text_sample': 'Detected from PDF score sections labeled Stu Pts Poss. Student IDs were not exposed in the PDF text, so anonymous per-file rows are used for this PDF.',
        'max_scores': max_scores,
        'question_performance': performance
    }

def normalize_answer(value):
    if pd.isna(value):
        return ''
    text = str(value).strip().upper()
    return re.sub(r'\s+', '', text)

def normalize_student_id(value):
    if pd.isna(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if re.match(r'^\d+\.0$', text):
        text = text[:-2]
    return re.sub(r'\s+', '', text)

def detect_student_id_column(df, question_columns):
    question_columns = set(question_columns or [])
    candidates = [col for col in df.columns if col not in question_columns]
    if not candidates:
        return None

    preferred_patterns = [
        r'\bstudent\b',
        r'\bstudent\s*id\b',
        r'\bid\b',
        r'\bnumber\b',
        r'\bname\b',
        r'الطلاب',
        r'الطالب',
        r'رقم'
    ]
    for col in candidates:
        label = str(col).strip()
        if any(re.search(pattern, label, flags=re.I) for pattern in preferred_patterns):
            return col

    return candidates[0]

def apply_student_id_index(score_df, student_ids):
    normalized_ids = []
    seen = {}
    for position, raw_id in enumerate(student_ids):
        student_id = normalize_student_id(raw_id)
        if not student_id:
            student_id = f"__missing_student_{position + 1}"
        seen[student_id] = seen.get(student_id, 0) + 1
        if seen[student_id] > 1:
            student_id = f"{student_id}__{seen[student_id]}"
        normalized_ids.append(student_id)

    score_df = score_df.reset_index(drop=True)
    score_df.index = normalized_ids[:len(score_df)]
    return score_df


def display_student_id(student_id):
    text = str(student_id or '').strip()
    match = re.search(r'__missing_student_(\d+)', text)
    if match:
        return match.group(1)
    return text

def build_scores_from_question_sheet(df, requested_questions):
    clean_df = df.dropna(how='all')
    if clean_df.empty:
        return None

    header_info = find_question_header_row(clean_df)
    if not header_info:
        return None

    requested = set(requested_questions)
    question_columns = {
        f'Q{number}': col_index
        for col_index, number, _ in header_info['question_cells']
        if f'Q{number}' in requested
    }
    if not question_columns:
        return None

    answer_key_row, student_rows = get_question_sheet_rows(clean_df, header_info)
    score_rows = []

    answer_key = None
    if answer_key_row is not None:
        answer_key = {
            question: normalize_answer(answer_key_row[col_index])
            for question, col_index in question_columns.items()
        }

    for row in student_rows:
        populated_answers = sum(
            0 if pd.isna(row[col_index]) or str(row[col_index]).strip() == '' else 1
            for col_index in question_columns.values()
        )
        if populated_answers == 0 and not answer_key:
            continue

        if answer_key:
            score_rows.append({
                question: 1.0 if normalize_answer(row[col_index]) == answer_key.get(question, '') else 0.0
                for question, col_index in question_columns.items()
            })
        else:
            score_rows.append({
                question: pd.to_numeric(row[col_index], errors='coerce')
                for question, col_index in question_columns.items()
            })

    if not score_rows:
        return None
    score_df = pd.DataFrame(score_rows)
    student_ids = [row.iloc[0] if len(row) else '' for row in student_rows]
    return apply_student_id_index(score_df, student_ids), 'binary' if answer_key else 'numeric'

def build_scores_from_tag_class_data(df, requested_questions):
    columns = get_tag_class_columns(df)
    if not columns:
        return None

    requested = set(requested_questions)
    question_rows = get_tag_class_question_rows(df, columns)
    if question_rows.empty:
        return None

    question_rows['_question_label'] = question_rows['_question_number'].map(lambda number: f'Q{int(number)}')
    question_rows = question_rows[question_rows['_question_label'].isin(requested)].copy()
    if question_rows.empty:
        return None

    question_rows['_student_id'] = question_rows[columns['student_id']].map(normalize_student_id)
    question_rows['_earned_points'] = pd.to_numeric(question_rows[columns['earned_points']], errors='coerce').fillna(0)
    question_rows = question_rows[question_rows['_student_id'] != '']
    if question_rows.empty:
        return None

    score_df = question_rows.pivot_table(
        index='_student_id',
        columns='_question_label',
        values='_earned_points',
        aggfunc='max',
        fill_value=0
    )
    available_questions = [question for question in requested_questions if question in score_df.columns]
    if not available_questions:
        return None
    return score_df[available_questions].apply(pd.to_numeric, errors='coerce').fillna(0), 'numeric'

def build_scores_from_generic_long_grade_data(df, requested_questions):
    columns = get_generic_long_grade_columns(df)
    if not columns:
        return None

    rows = get_generic_long_grade_rows(df, columns)
    if rows.empty:
        return None

    requested = set(requested_questions)
    rows = rows[rows['_question_label'].isin(requested)].copy()
    if rows.empty:
        return None

    score_df = rows.pivot_table(
        index='_student_id',
        columns='_question_label',
        values='_earned_points',
        aggfunc='max',
        fill_value=0
    )
    available_questions = [question for question in requested_questions if question in score_df.columns]
    if not available_questions:
        return None
    return score_df[available_questions].apply(pd.to_numeric, errors='coerce').fillna(0), 'numeric'

def build_score_dataframe(filepath, file_ext, requested_questions):
    requested_questions = list(requested_questions)

    if file_ext == '.pdf':
        text = extract_pdf_text(filepath)
        score_df, _ = parse_pdf_grade_report_text(text)
        if score_df is not None:
            available_questions = [question for question in requested_questions if question in score_df.columns]
            if available_questions:
                return score_df[available_questions].apply(pd.to_numeric, errors='coerce').fillna(0), 'numeric'
        return pd.DataFrame(columns=requested_questions), 'numeric'

    if file_ext == '.csv':
        df = pd.read_csv(filepath)
        tag_class_scores = build_scores_from_tag_class_data(df, requested_questions)
        if tag_class_scores:
            return tag_class_scores
        generic_long_scores = build_scores_from_generic_long_grade_data(df, requested_questions)
        if generic_long_scores:
            return generic_long_scores

        if all(question in df.columns for question in requested_questions):
            score_df = df[requested_questions].apply(pd.to_numeric, errors='coerce').fillna(0)
            id_column = detect_student_id_column(df, requested_questions)
            if id_column is not None:
                score_df = apply_student_id_index(score_df, df[id_column].tolist())
            return score_df, 'numeric'

        raw_df = pd.read_csv(filepath, header=None)
        scores = build_scores_from_question_sheet(raw_df, requested_questions)
        if scores:
            return scores
    else:
        df = pd.read_excel(filepath)
        tag_class_scores = build_scores_from_tag_class_data(df, requested_questions)
        if tag_class_scores:
            return tag_class_scores
        generic_long_scores = build_scores_from_generic_long_grade_data(df, requested_questions)
        if generic_long_scores:
            return generic_long_scores

        if all(question in df.columns for question in requested_questions):
            score_df = df[requested_questions].apply(pd.to_numeric, errors='coerce').fillna(0)
            id_column = detect_student_id_column(df, requested_questions)
            if id_column is not None:
                score_df = apply_student_id_index(score_df, df[id_column].tolist())
            return score_df, 'numeric'

        workbook = pd.ExcelFile(filepath)
        for sheet_name in workbook.sheet_names:
            raw_df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            scores = build_scores_from_question_sheet(raw_df, requested_questions)
            if scores:
                return scores

    return pd.DataFrame(columns=requested_questions), 'numeric'

def prefixed_question(assessment_name, question):
    return f"{assessment_name} {question}"

def combine_assessment_metrics(assessment_files):
    combined_questions = []
    combined_max_scores = {}
    combined_question_performance = {}
    combined_detected_clo_mappings = {}
    combined_question_texts = {}
    total_students = 0
    notes = []
    student_counts = {}
    confidence = 'Low'

    for assessment in assessment_files:
        metrics = assessment.get('metrics', {})
        questions = metrics.get('questions') or []
        student_counts[assessment['label']] = metrics.get('total_students') or 0
        if questions and metrics.get('confidence') == 'High':
            confidence = 'High'
        elif questions and confidence != 'High':
            confidence = 'Medium'

        for question in questions:
            combined = prefixed_question(assessment['label'], question)
            combined_questions.append(combined)
            combined_max_scores[combined] = metrics.get('max_scores', {}).get(question, 1.0)
            if metrics.get('question_performance', {}).get(question):
                combined_question_performance[combined] = metrics['question_performance'][question]
            if metrics.get('detected_clo_mappings', {}).get(question):
                combined_detected_clo_mappings[combined] = metrics['detected_clo_mappings'][question]
            if metrics.get('question_texts', {}).get(question):
                combined_question_texts[combined] = metrics['question_texts'][question]

        total_students = max(total_students, metrics.get('total_students') or 0)
        if metrics.get('text_sample'):
            notes.append(f"{assessment['label']}: {metrics['text_sample']}")

    nonzero_counts = {label: count for label, count in student_counts.items() if count > 0}
    student_count_warning = ''
    if len(set(nonzero_counts.values())) > 1:
        count_text = ', '.join(f"{label}: {count}" for label, count in nonzero_counts.items())
        student_count_warning = f"Student count mismatch across uploaded files ({count_text}). Results match students by ID across files and count any missing assessment score as 0."
    student_matching_warning = get_student_matching_warning(assessment_files)
    if student_matching_warning:
        student_count_warning = student_matching_warning

    return {
        'questions': combined_questions,
        'total_questions': len(combined_questions),
        'total_students': total_students,
        'confidence': confidence,
        'text_sample': ' '.join(notes),
        'max_scores': combined_max_scores,
        'question_performance': combined_question_performance,
        'detected_clo_mappings': combined_detected_clo_mappings,
        'question_texts': combined_question_texts,
        'student_counts': student_counts,
        'student_count_warning': student_count_warning
    }

def build_combined_score_dataframe(assessment_files, requested_questions):
    frames = []
    column_modes = {}
    for assessment in assessment_files:
        local_questions = []
        rename_map = {}
        for question in requested_questions:
            prefix = f"{assessment['label']} "
            if str(question).startswith(prefix):
                local_question = str(question)[len(prefix):]
                local_questions.append(local_question)
                rename_map[local_question] = question

        if not local_questions:
            continue

        filepath = get_upload_path(assessment['stored_name'])
        score_df, score_mode = build_score_dataframe(filepath, assessment['ext'], local_questions)
        if score_df.empty:
            continue
        score_df = score_df.rename(columns=rename_map)
        for local_question, combined_question in rename_map.items():
            if combined_question in score_df.columns:
                column_modes[combined_question] = score_mode
        score_df.index = [
            f"{assessment['label']}:{student_id}" if str(student_id).startswith("__missing_student_") else str(student_id)
            for student_id in score_df.index
        ]
        frames.append(score_df)

    if not frames:
        return pd.DataFrame(columns=list(requested_questions)), {}

    return pd.concat(frames, axis=1, join='outer').fillna(0), column_modes

def get_assessment_student_ids(assessment, requested_questions=None):
    questions = requested_questions or (assessment.get('metrics', {}).get('questions') or [])
    if not questions:
        return set()
    filepath = get_upload_path(assessment['stored_name'])
    score_df, _ = build_score_dataframe(filepath, assessment['ext'], questions)
    return {
        f"{assessment.get('label', 'Assessment')}:{student_id}" if str(student_id).startswith("__missing_student_") else str(student_id)
        for student_id in score_df.index
    }

def get_student_matching_warning(assessment_files):
    if len(assessment_files or []) <= 1:
        return ''

    student_ids_by_label = {}
    for assessment in assessment_files:
        label = assessment.get('label', 'Assessment')
        try:
            student_ids_by_label[label] = get_assessment_student_ids(assessment)
        except Exception:
            student_ids_by_label[label] = set()

    counts = {
        label: len(student_ids)
        for label, student_ids in student_ids_by_label.items()
        if student_ids
    }
    if len(counts) <= 1:
        return ''

    all_students = set().union(*student_ids_by_label.values()) if student_ids_by_label else set()
    missing_parts = []
    for label, student_ids in student_ids_by_label.items():
        if not student_ids:
            continue
        missing_count = len(all_students - student_ids)
        if missing_count:
            missing_parts.append(f"{label}: {missing_count} missing")

    if len(set(counts.values())) > 1 or missing_parts:
        count_text = ', '.join(f"{label}: {count}" for label, count in counts.items())
        missing_text = f" Missing by file: {', '.join(missing_parts)}." if missing_parts else ''
        return f"Student mismatch across uploaded files ({count_text}). Results match students by ID across files and count any missing assessment score as 0.{missing_text}"

    return ''

def calculate_clo_results():
    assessment_files = session.get('assessment_files') or []
    file_id = session.get('file_id')
    file_ext = session.get('file_ext')
    target_percentages = session.get('target_percentages', {"_global": 60.0})
    mapping_data = session.get('mapping', {})
    course_clos = session.get('custom_clos') or get_course_clos(session.get('course_name'))

    if not (assessment_files or file_id) or not mapping_data:
        return None, 0, [], "No mappings were provided."

    if assessment_files:
        score_df, score_mode = build_combined_score_dataframe(assessment_files, mapping_data.keys())
    else:
        filepath = get_upload_path(f"{file_id}{file_ext}")
        score_df, score_mode = build_score_dataframe(filepath, file_ext, mapping_data.keys())
    total_students = len(score_df)
    if total_students == 0:
        return None, 0, [], "Could not calculate scores from the uploaded file. Please check that the selected questions exist in the file."

    clo_stats = {}
    for col, data in mapping_data.items():
        for clo in expand_compact_clos(data.get('clos', []), course_clos):
            if clo not in clo_stats:
                clo_stats[clo] = {
                    'questions': [],
                    'students_achieved': 0,
                    'total_possible_score': 0
                }
            clo_stats[clo]['questions'].append(col)
            clo_stats[clo]['total_possible_score'] += data['max_score']

    student_achievement_rows = []
    for clo, stats in clo_stats.items():
        cols = stats['questions']
        max_possible = stats['total_possible_score']
        clo_target_pct = target_percentages.get(clo, target_percentages.get('_global', 60.0))
        target_score = max_possible * (clo_target_pct / 100.0)

        student_scores = pd.Series(0.0, index=score_df.index)
        for col in cols:
            if col not in score_df.columns:
                continue
            question_max = mapping_data.get(col, {}).get('max_score', 1.0)
            question_mode = score_mode.get(col, 'numeric') if isinstance(score_mode, dict) else score_mode
            if question_mode == 'binary':
                student_scores = student_scores + (score_df[col].fillna(0).astype(float) * question_max)
            else:
                student_scores = student_scores + score_df[col].fillna(0).astype(float)

        achieved_count = (student_scores >= target_score).sum()
        stats['students_achieved'] = int(achieved_count)
        stats['achievement_percentage'] = round((achieved_count / total_students) * 100, 2) if total_students > 0 else 0
        stats['target_score'] = round(target_score, 2)
        stats['target_pct'] = round(clo_target_pct, 2)
        stats['student_scores'] = {
            str(student_id): round(float(score), 2)
            for student_id, score in student_scores.items()
        }

        for student_id, score in student_scores.items():
            achieved = float(score) >= target_score
            student_achievement_rows.append({
                'student_id': str(student_id),
                'clo': clo,
                'score': round(float(score), 2),
                'target_score': round(target_score, 2),
                'target_pct': round(clo_target_pct, 2),
                'achieved': achieved,
                'status': 'Achieved' if achieved else 'Not Achieved'
            })

    sorted_clo_stats = dict(sorted_clo_items(clo_stats))
    student_achievement_rows.sort(key=lambda row: (row['student_id'], clo_sort_key(row['clo'])))
    return sorted_clo_stats, total_students, student_achievement_rows, None

def build_student_achievement_matrix(student_achievement_rows, clo_order=None):
    clo_order = sorted(list(clo_order or []), key=clo_sort_key)
    clo_set = set(clo_order)
    student_ids = []
    student_set = set()
    matrix = {}

    for row in student_achievement_rows or []:
        student_id = str(row.get('student_id', ''))
        clo = row.get('clo', '')
        if not student_id or not clo:
            continue
        if student_id not in student_set:
            student_ids.append(student_id)
            student_set.add(student_id)
        if clo not in clo_set:
            clo_order.append(clo)
            clo_order = sorted(clo_order, key=clo_sort_key)
            clo_set.add(clo)
        matrix.setdefault(student_id, {})[clo] = {
            'score': row.get('score', 0),
            'target_score': row.get('target_score', 0),
            'target_pct': row.get('target_pct', 0),
            'achieved': row.get('achieved', False),
            'status': row.get('status', 'Not Achieved')
        }

    return {
        'students': student_ids,
        'clos': clo_order,
        'cells': matrix
    }

def clo_number(clo):
    match = re.match(r'^\s*((?:CLO\s*)?\d+(?:\.\d+)*|[KSV]\s*\d+)\b', str(clo or ''), flags=re.I)
    if match:
        return re.sub(r'\s+', '', match.group(1)).upper()
    return str(clo or '').strip()


def compact_clo_value(clo):
    return clo_number_key(clo) or clo_number(clo)


def expand_compact_clos(selected_clos, course_clos):
    by_number = {}
    by_text = {}
    for clo in course_clos or []:
        text = str(clo or '').strip()
        number = compact_clo_value(text)
        if number and number not in by_number:
            by_number[number] = text
        if text:
            by_text[text] = text

    expanded = []
    for selected in selected_clos or []:
        selected_text = str(selected or '').strip()
        if not selected_text:
            continue
        match = by_number.get(selected_text) or by_text.get(selected_text) or selected_text
        if match not in expanded:
            expanded.append(match)
    return expanded

def clo_wording(clo):
    clo_text = str(clo or '').strip()
    number = clo_number(clo_text)
    if number and clo_text.upper().startswith(number.upper()):
        return clo_text[len(number):].strip(' .:-')
    match = re.match(r'^\s*(?:CLO\s*)?\d+(?:\.\d+)*\s+(.+)$', clo_text, flags=re.I)
    if match:
        return match.group(1).strip()
    return clo_text

def clo_domain_label(number):
    text = str(number or '').strip().upper()
    if text.startswith(('1.', 'CLO1', 'K')):
        return 'Knowledge'
    if text.startswith(('2.', 'CLO2', 'S')):
        return 'Skills'
    if text.startswith(('3.', 'CLO3', 'V')):
        return 'Values'
    return 'Other'

def clo_domain_translation_key(domain):
    return {
        'Knowledge': 'domain.knowledge',
        'Skills': 'domain.skills',
        'Values': 'domain.values',
        'Other': 'domain.other',
    }.get(domain, 'domain.other')

def clo_definition_sort_key(item):
    domain_order = {
        'Knowledge': 1,
        'Skills': 2,
        'Values': 3,
        'Other': 4,
    }
    number = str(item.get('number') or '')
    numeric_parts = [int(part) for part in re.findall(r'\d+', number)]
    return (domain_order.get(item.get('domain'), 4), numeric_parts, number)

def clo_sort_key(clo):
    number = clo_number(clo)
    domain = clo_domain_label(number)
    return clo_definition_sort_key({'number': number, 'domain': domain})

def sorted_clo_items(stats):
    return sorted((stats or {}).items(), key=lambda item: clo_sort_key(item[0]))

def build_clo_definitions(clos):
    definitions = []
    seen = set()
    for clo in clos or []:
        number = clo_number(clo)
        if not number or number in seen:
            continue
        seen.add(number)
        domain = clo_domain_label(number)
        definitions.append({
            'number': number,
            'domain': domain,
            'domain_key': clo_domain_translation_key(domain),
            'wording': clo_wording(clo),
            'full': str(clo or '').strip()
        })
    return sorted(definitions, key=clo_definition_sort_key)

def get_course_report_info():
    raw_name = session.get('course_name') or ''
    course = get_course_by_name(raw_name)
    match = re.search(r'\(([^()]*)\)\s*$', raw_name)
    course_id = (course.get('course_code') or '').strip() if course else ''
    if not course_id:
        course_id = match.group(1).strip() if match else ''
    course_name = (course.get('course_name') or '').strip() if course else ''
    if not course_name:
        course_name = raw_name[:match.start()].strip() if match else raw_name.strip()
    return {
        'course_name': course_name or raw_name,
        'course_id': course_id,
        'raw_name': raw_name,
        'college': (course.get('college') or '') if course else '',
        'department': (course.get('department') or '') if course else '',
        'program': (course.get('program') or '') if course else '',
        'clo_plos': (course.get('clo_plos') or {}) if course else {},
    }

def format_question_label(question):
    question = str(question)
    language = get_language() if has_request_context() else 'en'
    question_word = 'سؤال' if language == 'ar' else 'Question'
    match = re.match(r'^(.+?)\s+Q(\d+)$', question)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    match = re.match(r'^Q(\d+)$', question)
    if match:
        return f"{question_word} {match.group(1)}"
    match = re.match(r'^Question\s+(\d+)$', question, flags=re.IGNORECASE)
    if match:
        return f"{question_word} {match.group(1)}"
    return question

def localized_question_list_label(count, language=None):
    language = language or (get_language() if has_request_context() else 'en')
    if language == 'ar':
        return '\u0627\u0644\u0633\u0624\u0627\u0644' if count == 1 else '\u0627\u0644\u0623\u0633\u0626\u0644\u0629'
    return 'Question' if count == 1 else 'Questions'


def format_assessment_label(label, language=None):
    label_text = str(label or '').strip()
    language = language or (get_language() if has_request_context() else 'en')
    if language != 'ar':
        return label_text
    match = re.match(r'^(Final|Midterm|Quiz|Project|Assignment|Other)(?:\s+(\d+))?$', label_text, flags=re.I)
    if not match:
        return label_text
    key = match.group(1).lower()
    number = match.group(2)
    translated = TRANSLATIONS.get('ar', {}).get(f'assessment.{key}', label_text)
    return f"{translated} {number}" if number else translated


def format_question_label(question, language=None):
    question = str(question)
    language = language or (get_language() if has_request_context() else 'en')
    question_word = '\u0633\u0624\u0627\u0644' if language == 'ar' else 'Question'
    match = re.match(r'^(.+?)\s+Q(\d+)$', question)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    match = re.match(r'^Q(\d+)$', question)
    if match:
        return f"{question_word} {match.group(1)}"
    match = re.match(r'^Question\s+(\d+)$', question, flags=re.IGNORECASE)
    if match:
        return f"{question_word} {match.group(1)}"
    return question


def build_mapping_groups(columns, assessment_files):
    columns = list(columns or [])
    groups = []
    grouped_columns = set()

    if assessment_files:
        for assessment in assessment_files:
            label = assessment.get('label', 'Assessment')
            prefix = f"{label} "
            group_columns = [col for col in columns if str(col).startswith(prefix)]
            if not group_columns:
                continue
            grouped_columns.update(group_columns)
            groups.append({
                'label': label,
                'original_name': assessment.get('original_name', ''),
                'paper_original_name': assessment.get('paper_original_name', ''),
                'total_questions': len(group_columns),
                'total_students': assessment.get('metrics', {}).get('total_students') or 0,
                'columns': [
                    {
                        'value': col,
                        'display': format_question_label(str(col)[len(prefix):])
                    }
                    for col in group_columns
                ]
            })

    remaining_columns = [col for col in columns if col not in grouped_columns]
    if remaining_columns:
        groups.append({
            'label': 'Uploaded File',
            'original_name': '',
            'total_questions': len(remaining_columns),
            'total_students': 0,
            'columns': [
                {
                    'value': col,
                    'display': format_question_label(col)
                }
                for col in remaining_columns
            ]
        })

    return groups

def compact_question_list(question_keys, language=None):
    numbers = []
    for question in question_keys:
        match = re.match(r'^Q(\d+)$', str(question))
        if not match:
            return ', '.join(format_question_label(question, language=language) for question in question_keys)
        numbers.append(int(match.group(1)))

    if not numbers:
        return ''

    numbers = sorted(numbers)
    ranges = []
    start = numbers[0]
    previous = numbers[0]
    for number in numbers[1:]:
        if number == previous + 1:
            previous = number
            continue
        ranges.append((start, previous))
        start = previous = number
    ranges.append((start, previous))

    range_text = ', '.join(str(start) if start == end else f"{start}-{end}" for start, end in ranges)
    label = localized_question_list_label(len(numbers), language=language)
    return f"{label} {range_text}"

def format_missing_mapping_questions(missing_columns, assessment_files, language=None):
    missing_columns = [str(column) for column in missing_columns or []]
    if not missing_columns:
        return ''

    parts = []
    handled = set()
    for assessment in assessment_files or []:
        label = assessment.get('label', 'Assessment')
        prefix = f"{label} "
        group_columns = [column for column in missing_columns if column.startswith(prefix)]
        local_questions = [column[len(prefix):] for column in group_columns]
        if local_questions:
            handled.update(group_columns)
            parts.append(f"{format_assessment_label(label, language=language)}: {compact_question_list(local_questions, language=language)}")

    remaining = [column for column in missing_columns if column not in handled]
    if remaining:
        parts.append(compact_question_list(remaining, language=language))

    return '; '.join(parts)

def format_mapped_questions_for_report(questions, assessment_files=None, language=None):
    question_keys = [str(question) for question in questions or []]
    if not question_keys:
        return ''

    assessment_files = assessment_files if assessment_files is not None else (session.get('assessment_files') or [])
    parts = []
    handled = set()

    for assessment in assessment_files:
        label = assessment.get('label', 'Assessment')
        prefix = f"{label} "
        group_questions = [question for question in question_keys if question.startswith(prefix)]
        local_questions = [question[len(prefix):] for question in group_questions]
        if local_questions:
            parts.append(f"{format_assessment_label(label, language=language)}: {compact_question_list(local_questions, language=language)}")
            handled.update(group_questions)

    remaining = [question for question in question_keys if question not in handled]
    if remaining:
        parts.append(compact_question_list(remaining, language=language))

    return '; '.join(parts)

def pdf_escape(value):
    return str(value).replace('\\', r'\\').replace('(', r'\(').replace(')', r'\)')

def get_jpeg_size(image_bytes):
    index = 2
    while index < len(image_bytes):
        if image_bytes[index] != 0xFF:
            index += 1
            continue
        marker = image_bytes[index + 1]
        index += 2
        if marker in (0xD8, 0xD9):
            continue
        if index + 2 > len(image_bytes):
            break
        segment_length = int.from_bytes(image_bytes[index:index + 2], 'big')
        if marker in range(0xC0, 0xC4) and index + 7 < len(image_bytes):
            height = int.from_bytes(image_bytes[index + 3:index + 5], 'big')
            width = int.from_bytes(image_bytes[index + 5:index + 7], 'big')
            return width, height
        index += segment_length
    return 500, 500

def wrap_pdf_text(value, max_chars=95):
    words = str(value).split()
    lines = []
    current = ''
    for word in words:
        if len(current) + len(word) + 1 > max_chars:
            if current:
                lines.append(current)
            current = word
        else:
            current = f"{current} {word}".strip()
    if current:
        lines.append(current)
    return lines or ['']

def pdf_text(parts, x, y, text, size=10, font="F1"):
    safe_text = pdf_escape(str(text).encode('latin-1', errors='replace').decode('latin-1'))
    parts.append(f"BT /{font} {size} Tf {x} {y} Td ({safe_text}) Tj ET")

def pdf_line(parts, x1, y1, x2, y2):
    parts.append(f"{x1} {y1} m {x2} {y2} l S")

def pdf_rect(parts, x, y, width, height, fill=False):
    operator = "f" if fill else "S"
    parts.append(f"{x} {y} {width} {height} re {operator}")

def draw_pdf_lines(parts, lines, x, top_y, size=7, line_height=9, font="F1"):
    for index, line in enumerate(lines):
        pdf_text(parts, x, top_y - (index * line_height), line, size, font)

def get_report_pdf_font_paths():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    regular_candidates = [
        os.environ.get('REPORT_PDF_FONT_REGULAR'),
        os.path.join(base_dir, 'fonts', 'NotoNaskhArabic-Regular.ttf'),
        os.path.join(base_dir, 'fonts', 'NotoSansArabic-Regular.ttf'),
        os.path.join(base_dir, 'fonts', 'DejaVuSans.ttf'),
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf',
        '/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf',
        r'C:\Windows\Fonts\tahoma.ttf',
        r'C:\Windows\Fonts\arial.ttf',
        os.path.join(base_dir, 'public', 'fonts', 'RB.ttf'),
    ]
    bold_candidates = [
        os.environ.get('REPORT_PDF_FONT_BOLD'),
        os.path.join(base_dir, 'fonts', 'NotoNaskhArabic-Bold.ttf'),
        os.path.join(base_dir, 'fonts', 'NotoSansArabic-Bold.ttf'),
        os.path.join(base_dir, 'fonts', 'DejaVuSans-Bold.ttf'),
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
        '/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf',
        '/usr/share/fonts/truetype/noto/NotoSansArabic-Bold.ttf',
        r'C:\Windows\Fonts\tahomabd.ttf',
        r'C:\Windows\Fonts\arialbd.ttf',
        os.path.join(base_dir, 'public', 'fonts', 'RBTitle.ttf'),
        os.path.join(base_dir, 'public', 'fonts', 'RB.ttf'),
    ]

    regular_path = next((path for path in regular_candidates if path and os.path.exists(path)), '')
    bold_path = next((path for path in bold_candidates if path and os.path.exists(path)), regular_path)
    return regular_path, bold_path

def build_results_pdf_reportlab(stats, total_students, course_info, student_achievement_rows=None, branding=None):
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import BaseDocTemplate, Frame, Image, NextPageTemplate, PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
    except Exception:
        arabic_reshaper = None
        get_display = None

    regular_font_path, bold_font_path = get_report_pdf_font_paths()
    if not regular_font_path:
        raise RuntimeError("No Unicode PDF font found.")

    regular_font = 'CLOReportRegular'
    bold_font = 'CLOReportBold'
    if regular_font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(regular_font, regular_font_path))
    if bold_font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(bold_font, bold_font_path or regular_font_path))

    def display_text(value, reorder_arabic=True):
        text = clean_report_pdf_text(value)
        if contains_arabic(text) and arabic_reshaper:
            reshaped = arabic_reshaper.reshape(text)
            if reorder_arabic and get_display:
                return get_display(reshaped)
            return reshaped
        return text

    def paragraph(value, style, reorder_arabic=True):
        lines = str(value or '').split('\n')
        text = '<br/>'.join(
            escape(display_text(line, reorder_arabic=reorder_arabic))
            for line in lines
        )
        return Paragraph(text or '&nbsp;', style)

    def header_paragraph(value):
        text = str(value or '')
        if is_arabic_report and contains_arabic(text):
            text = '\n'.join(text.split())
        return paragraph(text, table_header)

    def student_header_paragraph(value):
        text = str(value or '')
        if is_arabic_report and contains_arabic(text):
            text = '\n'.join(text.split())
        return paragraph(text, student_table_header)

    def hex_color(value, fallback='#26365f'):
        try:
            return colors.HexColor(value or fallback)
        except Exception:
            return colors.HexColor(fallback)

    branding = apply_university_identity_colors(branding or get_report_branding())
    report_language = get_export_report_language()
    labels = pdf_report_labels(report_language)
    is_arabic_report = report_language == 'ar'
    organization_display_name = localized_university_name(branding.get('organization_name'), report_language) or labels['na']
    primary_color = hex_color(branding.get('primary_color'))
    accent_color = hex_color(branding.get('secondary_color') or branding.get('primary_color'))
    body_text_color = colors.black
    logo_path = resolve_branding_logo_path(branding, report_ready=True)
    report_date = datetime.now().strftime("%Y-%m-%d")

    buffer = io.BytesIO()
    portrait_size = letter
    landscape_size = landscape(letter)
    left_margin = 0.45 * inch
    right_margin = 0.45 * inch
    top_margin = 0.42 * inch
    bottom_margin = 0.55 * inch
    doc = BaseDocTemplate(
        buffer,
        pagesize=portrait_size,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )
    portrait_frame = Frame(
        left_margin,
        bottom_margin,
        portrait_size[0] - left_margin - right_margin,
        portrait_size[1] - top_margin - bottom_margin,
        id='portrait_frame'
    )
    landscape_frame = Frame(
        left_margin,
        bottom_margin,
        landscape_size[0] - left_margin - right_margin,
        landscape_size[1] - top_margin - bottom_margin,
        id='landscape_frame'
    )

    std_styles = get_standard_paragraph_styles(branding.get('primary_color'), is_arabic_report, regular_font, bold_font)
    title_style = std_styles['title']
    meta_style = std_styles['meta']
    section_style = std_styles['section']
    table_header = std_styles['table_header']
    
    table_text = ParagraphStyle(
        'ReportTableText',
        parent=std_styles['table_text'],
        fontSize=7.4,
        leading=9.2,
    )
    table_text_ar = ParagraphStyle(
        'ReportTableTextArabic',
        parent=table_text,
        alignment=TA_RIGHT,
    )
    student_table_text = ParagraphStyle(
        'StudentReportTableText',
        parent=table_text,
        fontSize=5.6,
        leading=6.6,
        alignment=TA_CENTER,
        splitLongWords=0,
    )
    student_table_header = ParagraphStyle(
        'StudentReportTableHeader',
        parent=table_header,
        fontSize=5.6,
        leading=6.6,
        alignment=TA_CENTER,
        splitLongWords=0,
    )

    def table_paragraph(value):
        return paragraph(value, table_text_ar if contains_arabic(str(value or '')) else table_text)

    def rtl_table(rows, col_widths):
        if not is_arabic_report:
            return rows, col_widths
        return [list(reversed(row)) for row in rows], list(reversed(col_widths))

    elements = []
    heading_cells = []
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path)
            ratio = logo.imageHeight / max(logo.imageWidth, 1)
            logo.drawWidth = 1.15 * inch
            logo.drawHeight = min(1.05 * inch, logo.drawWidth * ratio)
            heading_cells.append(logo)
        except Exception:
            heading_cells.append('')
    heading_text = [
        paragraph(labels['title'], title_style),
        paragraph(f"{labels['university']}: {organization_display_name}", meta_style),
    ]
    if branding.get('department'):
        heading_text.append(paragraph(f"{labels['department']}: {branding.get('department')}", meta_style))
    heading_text.extend([
        paragraph(f"{labels['course_name']}: {course_info.get('course_name', '')}", meta_style),
        paragraph(f"{labels['course_id']}: {course_info.get('course_id', '') or labels['na']}", meta_style),
        paragraph(f"{labels['report_date']}: {report_date}", meta_style),
    ])
    if heading_cells:
        if is_arabic_report:
            header_table = Table([[heading_text, heading_cells[0]]], colWidths=[5.45 * inch, 1.35 * inch])
        else:
            heading_cells.append(heading_text)
            header_table = Table([heading_cells], colWidths=[1.35 * inch, 5.45 * inch])
    else:
        header_table = Table([[heading_text]], colWidths=[6.8 * inch])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LINEBELOW', (0, 0), (-1, -1), 1, accent_color),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 10))

    summary_rows, summary_widths = rtl_table([[
        paragraph(labels['total_students'], table_text),
        paragraph(str(total_students), ParagraphStyle('SummaryNumber1', parent=table_text, fontName=bold_font, fontSize=13, textColor=body_text_color)),
        paragraph(labels['mapped_clos'], table_text),
        paragraph(str(len(stats)), ParagraphStyle('SummaryNumber2', parent=table_text, fontName=bold_font, fontSize=13, textColor=body_text_color)),
    ]], [1.9 * inch, 0.8 * inch, 1.4 * inch, 0.7 * inch])
    summary_table = Table(
        summary_rows,
        colWidths=summary_widths
    )
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f1f5f9')),
        ('BOX', (0, 0), (-1, -1), 0.7, colors.HexColor('#cbd5e1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 12))

    clo_definitions = build_clo_definitions(stats.keys())
    elements.append(paragraph(labels['clo_definitions'], section_style))
    definition_rows = [[header_paragraph(labels['domain']), header_paragraph(labels['clo']), header_paragraph(labels['wording'])]]
    for item in clo_definitions:
        definition_rows.append([
            paragraph(localized_clo_domain(item['domain'], report_language), table_text),
            paragraph(item['number'], table_text),
            table_paragraph(item['wording']),
        ])
    definition_rows, definition_widths = rtl_table(definition_rows, [1.15 * inch, 0.75 * inch, 5.0 * inch])
    definition_table = Table(definition_rows, colWidths=definition_widths, repeatRows=1)
    definition_table.setStyle(get_standard_table_style(branding.get('primary_color'), is_arabic_report, len(definition_rows)))
    elements.append(definition_table)
    elements.append(Spacer(1, 12))

    elements.append(paragraph(labels['summary'], section_style))
    headers = [labels['clo'], labels['questions'], labels['max'], labels['target'], labels['achieved'], labels['achievement']]
    rows = [[header_paragraph(header) for header in headers]]
    for clo, data in sorted_clo_items(stats):
        rows.append([
            paragraph(clo_number(clo), table_text),
            table_paragraph(format_mapped_questions_for_report(data['questions'], language=report_language)),
            paragraph(f"{data['total_possible_score']:.2f}", table_text),
            paragraph(f"{data['target_score']:.2f}", table_text),
            paragraph(str(data['students_achieved']), table_text),
            paragraph(f"{data['achievement_percentage']:.2f}%", table_text),
        ])
    rows, summary_result_widths = rtl_table(rows, [0.62 * inch, 2.55 * inch, 0.78 * inch, 1.03 * inch, 1.12 * inch, 0.80 * inch])
    summary_results = Table(rows, colWidths=summary_result_widths, repeatRows=1)
    summary_results.setStyle(get_standard_table_style(branding.get('primary_color'), is_arabic_report, len(rows)))
    elements.append(summary_results)

    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    if student_achievement_matrix['students']:
        elements.append(NextPageTemplate('landscape'))
        elements.append(PageBreak())
        elements.append(paragraph(labels['student_achievement'], section_style))
        all_clos = student_achievement_matrix['clos']
        student_headers = [labels['student_id']] + [clo_number(clo) for clo in all_clos]
        student_rows = [[student_header_paragraph(header) for header in student_headers]]
        for student_id in student_achievement_matrix['students']:
            row = [paragraph(display_student_id(student_id), student_table_text)]
            for clo in all_clos:
                cell = student_achievement_matrix['cells'].get(student_id, {}).get(clo)
                if cell:
                    status = labels.get('student_achieved_status', labels.get('achieved_status', labels['achieved'])) if cell.get('achieved') else labels.get('student_not_achieved_status', labels['not_achieved'])
                    status = str(status).replace(' ', '\u00a0')
                    row.append(paragraph(f"{cell.get('score', 0):.2f}\n{status}", student_table_text))
                else:
                    row.append(paragraph("-", student_table_text))
            student_rows.append(row)
        available_landscape_width = landscape_size[0] - left_margin - right_margin
        student_id_width = 1.05 * inch
        clo_width = (available_landscape_width - student_id_width) / max(len(all_clos), 1)
        student_rows, student_widths = rtl_table(student_rows, [student_id_width] + [clo_width] * len(all_clos))
        student_table = Table(
            student_rows,
            colWidths=student_widths,
            repeatRows=1
        )
        student_table.setStyle(get_standard_table_style(branding.get('primary_color'), is_arabic_report, len(student_rows)))
        elements.append(student_table)

    def footer(canvas, _doc):
        canvas.saveState()
        canvas.setFont(regular_font, 8)
        canvas.setFillColor(colors.HexColor('#64748b'))
        page_width, _ = canvas._pagesize
        canvas.drawString(0.45 * inch, 0.27 * inch, "Generated by ETQAN")
        canvas.drawRightString(page_width - 0.45 * inch, 0.27 * inch, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    doc.addPageTemplates([
        PageTemplate(id='portrait', frames=[portrait_frame], pagesize=portrait_size, onPage=footer),
        PageTemplate(id='landscape', frames=[landscape_frame], pagesize=landscape_size, onPage=footer),
    ])
    doc.build(elements)
    return buffer.getvalue()

def build_results_pdf(stats, total_students, course_info, student_achievement_rows=None, branding=None):
    def payload_contains_arabic():
        candidates = [
            course_info.get('course_name') if isinstance(course_info, dict) else '',
            course_info.get('course_id') if isinstance(course_info, dict) else '',
        ]
        for data in (stats or {}).values():
            candidates.extend(data.get('questions') or [])
            candidates.append(data.get('clo') or '')
        branding_data = branding or get_report_branding()
        candidates.extend([
            branding_data.get('organization_name') or '',
            branding_data.get('college') or '',
            branding_data.get('department') or '',
        ])
        return any(contains_arabic(value) for value in candidates)

    try:
        return build_results_pdf_reportlab(stats, total_students, course_info, student_achievement_rows, branding)
    except Exception as exc:
        app.logger.warning("Falling back to legacy PDF renderer: %s", exc)
        if payload_contains_arabic():
            app.logger.exception("Unicode PDF rendering failed for Arabic content; legacy PDF renderer cannot preserve Arabic text.")
            raise
        return build_results_pdf_legacy(stats, total_students, course_info, student_achievement_rows, branding)

GRADE_ORDER = ['A+', 'A', 'B+', 'B', 'C+', 'C', 'D+', 'D', 'F']
ARABIC_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
GRADE_HEADER_HINTS = (
    'grade', 'garde', 'letter', 'result', 'score', 'mark', 'marks', 'final', 'total',
    'degree', 'points', 'تقدير', 'درجة', 'درجات', 'نتيجة', 'النتيجة', 'المجموع',
    'نهائي', 'نهائية', 'النهائي', 'النهائية', 'علامة', 'علامات'
)
IDENTITY_HEADER_HINTS = (
    'student', 'student id', 'student number', 'student no', 'id', 'name', 'full name',
    'اسم', 'الاسم', 'الطالب', 'الطالبة', 'طلاب', 'طالبات', 'رقم', 'الرقم',
    'جامعي', 'الجامعي', 'السجل', 'القيد'
)
GRADE_SOURCE_PRIORITY_HINTS = ('final', 'total', 'grade', 'garde', 'result', 'نهائي', 'النهائي', 'المجموع', 'تقدير')

GRADE_HEADER_HINTS = GRADE_HEADER_HINTS + (
    'تقدير', 'التقدير', 'درجة', 'درجات', 'نتيجة', 'النتيجة', 'المجموع', 'مجموع',
    'الدرجة رمزاً', 'الدرجة رمزا', 'نهائي', 'النهائي', 'الدرجة النهائية', 'علامة', 'علامات',
)
IDENTITY_HEADER_HINTS = IDENTITY_HEADER_HINTS + (
    'اسم', 'الاسم', 'اسم الطالب', 'الطالب', 'الرقم الجامعي', 'رقم جامعي', 'رقم الطالب',
    'السجل', 'القيد', 'حالة الطالب',
)
GRADE_SOURCE_PRIORITY_HINTS = GRADE_SOURCE_PRIORITY_HINTS + (
    'نهائي', 'النهائي', 'المجموع', 'الدرجة رمزاً', 'الدرجة رمزا', 'تقدير',
)

def empty_grade_counts():
    return {grade: 0 for grade in GRADE_ORDER}

def normalize_grade_text(value):
    text = str(value or '').strip()
    if not text or text.upper() in {'NAN', 'NONE', 'NULL'}:
        return ''
    text = text.translate(ARABIC_DIGITS)
    text = text.replace('＋', '+').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def normalize_letter_grade(value):
    text = normalize_grade_text(value)
    if not text:
        return ''
    arabic_grade_map = [
        ('ممتاز مرتفع', 'A+'),
        ('ممتاز', 'A'),
        ('جيد جدا مرتفع', 'B+'),
        ('جيد جداً مرتفع', 'B+'),
        ('جيد جدا', 'B'),
        ('جيد جداً', 'B'),
        ('جيد مرتفع', 'C+'),
        ('جيد', 'C'),
        ('مقبول مرتفع', 'D+'),
        ('مقبول', 'D'),
        ('راسب', 'F'),
        ('ه+', 'F'),
        ('هـ+', 'F'),
        ('ه', 'F'),
        ('هـ', 'F'),
        ('ا+', 'A+'),
        ('ا', 'A'),
        ('ب+', 'B+'),
        ('ب', 'B'),
        ('ج+', 'C+'),
        ('ج', 'C'),
        ('د+', 'D+'),
        ('د', 'D'),
    ]
    compact_arabic = re.sub(r'[\s\W_]+', ' ', text, flags=re.UNICODE).strip()
    for phrase, grade in arabic_grade_map:
        if compact_arabic == phrase:
            return grade
    english = text.upper()
    english = english.replace('A PLUS', 'A+').replace('B PLUS', 'B+').replace('C PLUS', 'C+').replace('D PLUS', 'D+')
    english = english.replace('APLUS', 'A+').replace('BPLUS', 'B+').replace('CPLUS', 'C+').replace('DPLUS', 'D+')
    english = english.strip('.,;:()[]{}')
    if english in GRADE_ORDER:
        return english
    match = re.search(r'\b(A\+?|B\+?|C\+?|D\+?|F)\b', english)
    return match.group(1) if match and match.group(1) in GRADE_ORDER else ''

def parse_numeric_score_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        score = float(value)
        return score if 0 <= score <= 100 else None
    text = normalize_grade_text(value)
    if not text:
        return None
    text = text.replace('%', ' ')
    if re.fullmatch(r'\d{4,}', text):
        return None
    ratio = re.search(r'(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)', text)
    if ratio:
        score = float(ratio.group(1))
        maximum = float(ratio.group(2))
        if maximum > 0 and maximum != 100:
            score = score / maximum * 100
        return score if 0 <= score <= 100 else None
    matches = re.findall(r'(?<!\d)(\d+(?:\.\d+)?)(?!\d)', text)
    for match in reversed(matches):
        score = float(match)
        if 0 <= score <= 100:
            return score
    return None

def parse_grade_count_value(value):
    score = parse_numeric_score_value(value)
    if score is None:
        return None
    rounded = round(score)
    if abs(score - rounded) > 1e-9:
        return None
    count = int(rounded)
    return count if 0 <= count <= 1000 else None

def numeric_score_to_letter_grade(value):
    try:
        score = float(value)
    except (TypeError, ValueError):
        return ''
    if score > 100:
        return ''
    if score >= 95:
        return 'A+'
    if score >= 90:
        return 'A'
    if score >= 85:
        return 'B+'
    if score >= 80:
        return 'B'
    if score >= 75:
        return 'C+'
    if score >= 70:
        return 'C'
    if score >= 65:
        return 'D+'
    if score >= 60:
        return 'D'
    if score >= 0:
        return 'F'
    return ''

def grade_distribution_from_letters(grades):
    counts = empty_grade_counts()
    for grade in grades:
        normalized = normalize_letter_grade(grade)
        if normalized:
            counts[normalized] += 1
    return counts

def grade_distribution_from_scores(scores):
    counts = empty_grade_counts()
    for score in scores:
        grade = numeric_score_to_letter_grade(score)
        if grade:
            counts[grade] += 1
    return counts

def normalize_header(value):
    text = normalize_grade_text(value).lower()
    return re.sub(r'[\s_\-:/\\|]+', ' ', text).strip()

def header_has_hint(header, hints):
    compact = header.replace(' ', '')
    return any(hint in header or hint.replace(' ', '') in compact for hint in hints)

def is_identity_header(header):
    header = normalize_header(header)
    if header.startswith('unnamed') or re.fullmatch(r'column \d+', header):
        return False
    return header_has_hint(header, IDENTITY_HEADER_HINTS)

def is_grade_header(header):
    header = normalize_header(header)
    return header_has_hint(header, GRADE_HEADER_HINTS) and not is_identity_header(header)

def grade_header_priority(header):
    header = normalize_header(header)
    return sum(3 for hint in GRADE_SOURCE_PRIORITY_HINTS if hint in header) + (1 if is_grade_header(header) else 0)

def nonempty_series_mask(series):
    return series.apply(lambda value: bool(str(value).strip()) and str(value).strip().lower() not in {'nan', 'none', 'null'})

def dataframe_structured_grade_distribution(frame, source='Structured grade table'):
    if frame is None or frame.empty:
        return finalize_grade_distribution(empty_grade_counts())

    headers = [str(column or '').strip() for column in frame.columns]
    identity_columns = [column for column, header in zip(frame.columns, headers) if is_identity_header(header)]
    grade_columns = [column for column, header in zip(frame.columns, headers) if is_grade_header(header)]
    if not identity_columns or not grade_columns:
        return finalize_grade_distribution(empty_grade_counts())

    identity_mask = pd.Series(False, index=frame.index)
    for column in identity_columns:
        identity_mask = identity_mask | nonempty_series_mask(frame[column])
    if not bool(identity_mask.any()):
        for column in grade_columns:
            identity_mask = identity_mask | nonempty_series_mask(frame[column])

    candidates = []
    for column in grade_columns:
        header = str(column or '').strip()
        series = frame.loc[identity_mask, column].dropna()
        if series.empty:
            continue

        letters = [normalize_letter_grade(value) for value in series]
        letter_count = sum(1 for grade in letters if grade)
        if letter_count:
            candidates.append((letter_count + grade_header_priority(header), letter_count, 'letter', header, letters))

        scores = [parse_numeric_score_value(value) for value in series]
        scores = [score for score in scores if score is not None]
        if scores:
            candidates.append((len(scores) + grade_header_priority(header), len(scores), 'score', header, scores))

    if not candidates:
        return finalize_grade_distribution(empty_grade_counts())

    _, _, mode, header, values = max(candidates, key=lambda item: (item[0], item[1]))
    if mode == 'letter':
        return finalize_grade_distribution(grade_distribution_from_letters(values), f"{source}: letter grades column: {header}")
    return finalize_grade_distribution(grade_distribution_from_scores(values), f"{source}: numeric scores column: {header}")

def grade_distribution_from_summary_frame(raw_frame, source='Grade summary table'):
    if raw_frame is None or raw_frame.empty:
        return finalize_grade_distribution(empty_grade_counts())

    counts = empty_grade_counts()
    found_pairs = 0
    row_count, column_count = raw_frame.shape
    for row_index in range(row_count):
        for column_index in range(column_count):
            grade = normalize_letter_grade(raw_frame.iat[row_index, column_index])
            if not grade:
                continue

            # Student rows often put a letter grade immediately after a numeric total.
            # A summary table normally puts the grade label first, then the count nearby.
            if column_index > 0 and parse_numeric_score_value(raw_frame.iat[row_index, column_index - 1]) is not None:
                continue
            if column_index > 1 and parse_numeric_score_value(raw_frame.iat[row_index, column_index - 2]) is not None:
                continue

            for offset in range(1, 5):
                next_column = column_index + offset
                if next_column >= column_count:
                    break
                count = parse_grade_count_value(raw_frame.iat[row_index, next_column])
                if count is None:
                    continue
                counts[grade] += count
                found_pairs += 1
                break

    total = grade_distribution_total(counts)
    if total >= 3 and found_pairs >= 2:
        return finalize_grade_distribution(counts, source)
    return finalize_grade_distribution(empty_grade_counts())

def unique_headers(headers):
    seen = {}
    result = []
    for index, header in enumerate(headers):
        base = compact_text(str(header or '')).strip() or f'column_{index + 1}'
        count = seen.get(base, 0)
        seen[base] = count + 1
        result.append(base if not count else f'{base}_{count + 1}')
    return result

def grade_distribution_from_headerless_frame(raw_frame, source='Detected grade table'):
    if raw_frame is None or raw_frame.empty:
        return finalize_grade_distribution(empty_grade_counts())
    best = grade_distribution_from_summary_frame(raw_frame, f"{source}: grade summary")
    max_header_rows = min(12, len(raw_frame.index))
    for header_row in range(max_header_rows):
        headers = raw_frame.iloc[header_row].tolist()
        has_grade = any(is_grade_header(header) for header in headers)
        has_identity = any(is_identity_header(header) for header in headers)
        if not has_grade or not has_identity:
            continue
        body = raw_frame.iloc[header_row + 1:].copy()
        body.columns = unique_headers(headers)
        distribution = dataframe_structured_grade_distribution(body, f"{source}: header row {header_row + 1}")
        if distribution['total'] > best['total']:
            best = distribution
    return best

def pdf_generic_grade_values(text):
    scores = []
    letters = []
    for line in (text or '').splitlines():
        raw = normalize_grade_text(line)
        if not raw:
            continue
        lowered = raw.lower()
        if header_has_hint(normalize_header(raw), GRADE_HEADER_HINTS) and not re.search(r'\d', raw):
            continue

        has_id = bool(re.search(r'\b\d{6,15}\b', raw))
        word_count = len(re.findall(r'[A-Za-z\u0600-\u06FF]{2,}', raw))
        if not has_id and word_count < 2:
            continue

        without_ids = re.sub(r'\b\d{6,15}\b', ' ', raw)
        tokens = [token.strip('.,;:()[]{}') for token in re.split(r'\s+', without_ids) if token.strip('.,;:()[]{}')]
        tail_grade = ''
        for size in (3, 2, 1):
            if len(tokens) >= size:
                tail_grade = normalize_letter_grade(' '.join(tokens[-size:]))
                if tail_grade:
                    break
        if tail_grade:
            letters.append(tail_grade)
            continue

        trailing_score = re.search(r'(\d+(?:\.\d+)?)(?:\s*/\s*\d+(?:\.\d+)?)?\s*%?\s*$', without_ids)
        if trailing_score:
            score = parse_numeric_score_value(trailing_score.group(0))
            if score is not None:
                scores.append(score)
                continue

        if any(hint in lowered for hint in ('score', 'mark', 'final', 'total', 'degree', 'درجة', 'درجات', 'المجموع', 'نهائي', 'نهائية')):
            score = parse_numeric_score_value(without_ids)
            if score is not None:
                scores.append(score)

    if len(scores) >= 3:
        return 'score', scores
    if len(letters) >= 3:
        return 'letter', letters
    return '', []

def pdf_student_row_scores(text):
    scores = []
    for line in (text or '').splitlines():
        raw = str(line or '').strip()
        if not raw:
            continue
        match = re.search(r'(\d{10,12})(?=\D*$)', raw)
        if not match:
            continue
        suffix = raw[match.end():].strip()
        if not suffix and not re.search(r'grade|score|mark|final|degree|درجة|درجات|تقدير', raw, flags=re.I):
            continue
        number = match.group(1)
        score = 100 if number.endswith('100') else int(number[-2:])
        if 0 <= score <= 100:
            scores.append(score)
    return scores if len(scores) >= 3 else []

def grade_distribution_total(counts):
    return sum(int(counts.get(grade, 0) or 0) for grade in GRADE_ORDER)

def finalize_grade_distribution(counts, source=''):
    counts = {grade: int(counts.get(grade, 0) or 0) for grade in GRADE_ORDER}
    total = grade_distribution_total(counts)
    return {
        'counts': counts,
        'rows': [
            {
                'grade': grade,
                'count': counts[grade],
                'percentage': round((counts[grade] / total * 100), 2) if total else 0,
            }
            for grade in GRADE_ORDER
        ],
        'total': total,
        'source': source,
    }

def dataframe_grade_distribution(df):
    if df is None or df.empty:
        return finalize_grade_distribution(empty_grade_counts())

    frame = df.dropna(how='all').copy()
    if frame.empty:
        return finalize_grade_distribution(empty_grade_counts())

    structured = dataframe_structured_grade_distribution(frame)
    if structured['total']:
        return structured

    header_hints = ('grade', 'letter', 'final', 'result', 'score', 'mark', 'total', 'تقدير', 'نهائي', 'نهائية', 'المجموع', 'درجة')
    letter_candidates = []
    score_candidates = []
    for column in frame.columns:
        header = str(column or '').strip().lower()
        series = frame[column].dropna()
        if series.empty:
            continue
        letters = [normalize_letter_grade(value) for value in series]
        letter_count = sum(1 for grade in letters if grade)
        header_score = 2 if any(hint in header for hint in header_hints) else 0
        if letter_count:
            letter_candidates.append((letter_count + header_score, letter_count, column, letters))

        numeric = [parse_numeric_score_value(value) for value in series]
        numeric = [score for score in numeric if score is not None]
        if numeric:
            preferred = any(hint in header for hint in ('final', 'score', 'mark', 'total', 'result', 'نهائي', 'نهائية', 'المجموع', 'درجة'))
            score_candidates.append((len(numeric) + (3 if preferred else 0), len(numeric), column, numeric))

    if letter_candidates:
        _, _, column, letters = max(letter_candidates, key=lambda item: (item[0], item[1]))
        return finalize_grade_distribution(grade_distribution_from_letters(letters), f"Letter grades column: {column}")

    if score_candidates:
        _, _, column, scores = max(score_candidates, key=lambda item: (item[0], item[1]))
        return finalize_grade_distribution(grade_distribution_from_scores(scores), f"Numeric final scores column: {column}")

    row_letters = []
    for _, row in frame.iterrows():
        for value in row.tolist():
            grade = normalize_letter_grade(value)
            if grade:
                row_letters.append(grade)
                break
    return finalize_grade_distribution(grade_distribution_from_letters(row_letters), "Detected letter grades in rows")

def best_grade_distribution(distributions):
    best = finalize_grade_distribution(empty_grade_counts())
    best_source_score = 0
    for distribution in distributions:
        if not distribution:
            continue
        source = distribution.get('source') or ''
        source_score = 2 if 'header row' in source else 1 if 'Structured grade table' in source else 0
        if (
            distribution.get('total', 0) > best.get('total', 0)
            or (distribution.get('total', 0) == best.get('total', 0) and source_score > best_source_score)
        ):
            best = distribution
            best_source_score = source_score
    return best

def parse_csv_grade_distribution(filepath):
    distributions = []
    try:
        distributions.append(dataframe_grade_distribution(pd.read_csv(filepath)))
    except Exception:
        pass
    try:
        raw = pd.read_csv(filepath, header=None)
        distributions.append(grade_distribution_from_headerless_frame(raw, "CSV grade table"))
    except Exception:
        pass
    return best_grade_distribution(distributions)

def parse_excel_sheet_grade_distribution(filepath, sheet_name):
    distributions = []
    try:
        distributions.append(dataframe_grade_distribution(pd.read_excel(filepath, sheet_name=sheet_name)))
    except Exception:
        pass
    try:
        raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        distributions.append(grade_distribution_from_summary_frame(raw, f"Excel sheet {sheet_name}: grade summary"))
        distributions.append(grade_distribution_from_headerless_frame(raw, f"Excel sheet {sheet_name}"))
    except Exception:
        pass
    return best_grade_distribution(distributions)

def parse_final_grade_distribution(filepath, file_ext):
    file_ext = str(file_ext or '').lower()
    if file_ext == '.pdf':
        text = extract_pdf_text(filepath)
        pdf_mode, pdf_values = pdf_generic_grade_values(text)
        if pdf_mode == 'score':
            return finalize_grade_distribution(grade_distribution_from_scores(pdf_values), "PDF numeric final scores")
        if pdf_mode == 'letter':
            return finalize_grade_distribution(grade_distribution_from_letters(pdf_values), "PDF letter grades")
        scores = pdf_student_row_scores(text)
        if scores:
            return finalize_grade_distribution(grade_distribution_from_scores(scores), "PDF numeric final scores")
        grades = []
        for line in (text or '').splitlines():
            line_grades = []
            for token in re.split(r'[\s,;|]+', line):
                grade = normalize_letter_grade(token)
                if grade:
                    line_grades.append(grade)
            if line_grades and not re.search(r'\bgrade\b|\bletter\b', line, flags=re.I):
                grades.append(line_grades[-1])
        return finalize_grade_distribution(grade_distribution_from_letters(grades), "PDF letter grades")

    if file_ext == '.csv':
        return parse_csv_grade_distribution(filepath)

    workbook = pd.ExcelFile(filepath)
    best = finalize_grade_distribution(empty_grade_counts())
    best_sheet = ''
    for sheet_name in workbook.sheet_names:
        distribution = parse_excel_sheet_grade_distribution(filepath, sheet_name)
        if distribution['total'] > best['total']:
            best = distribution
            best_sheet = sheet_name
    if best_sheet:
        best['source'] = f"{best.get('source') or 'Detected grades'}; sheet: {best_sheet}"
    return best

def read_uncovered_topic_details():
    details = []
    for index in request.form.getlist('uncovered_topic_indexes'):
        index = str(index)
        topic = compact_text(request.form.get(f'uncovered_topic_{index}') or '')
        if not topic:
            continue
        reason = compact_text(request.form.get(f'uncovered_reason_{index}') or '')
        other_reason = compact_text(request.form.get(f'uncovered_reason_other_{index}') or '')
        if reason == 'Other' and other_reason:
            reason = f"Other: {other_reason}"
        action = compact_text(request.form.get(f'uncovered_action_{index}') or '')
        other_action = compact_text(request.form.get(f'uncovered_action_other_{index}') or '')
        if action == 'Other' and other_action:
            action = f"Other: {other_action}"
        details.append({
            'topic': topic,
            'reason': reason,
            'impact': compact_text(request.form.get(f'uncovered_impact_{index}') or ''),
            'action': action,
        })
    return details

def read_course_improvement_plan():
    selected = {compact_text(value) for value in request.form.getlist('course_improvement_recommendations')}
    allowed_action = set(COURSE_IMPROVEMENT_ACTION_OPTIONS)
    allowed_support = set(COURSE_IMPROVEMENT_SUPPORT_OPTIONS)
    items = []
    seen = set()
    for index, recommendation in enumerate(COURSE_IMPROVEMENT_RECOMMENDATIONS):
        if recommendation not in selected or recommendation in seen:
            continue
        seen.add(recommendation)
        action = compact_text(request.form.get(f'course_improvement_action_{index}') or '')
        if action not in allowed_action:
            action = ''
        other_action = compact_text(request.form.get(f'course_improvement_action_other_{index}') or '')
        if action == 'Other' and other_action:
            action = f"Other: {other_action}"
        support = compact_text(request.form.get(f'course_improvement_support_{index}') or '')
        if support not in allowed_support:
            support = ''
        other_support = compact_text(request.form.get(f'course_improvement_support_other_{index}') or '')
        if support == 'Other' and other_support:
            support = f"Other: {other_support}"
        items.append({
            'recommendation': recommendation,
            'actions_needed': action,
            'support': support,
        })
    if request.form.get('course_improvement_other_selected'):
        recommendation = compact_text(request.form.get('course_improvement_other_recommendation') or '')
        if recommendation:
            action = compact_text(request.form.get('course_improvement_action_other_recommendation') or '')
            if action not in allowed_action:
                action = ''
            other_action = compact_text(request.form.get('course_improvement_action_other_recommendation_text') or '')
            if action == 'Other' and other_action:
                action = f"Other: {other_action}"
            support = compact_text(request.form.get('course_improvement_support_other_recommendation') or '')
            if support not in allowed_support:
                support = ''
            other_support = compact_text(request.form.get('course_improvement_support_other_recommendation_text') or '')
            if support == 'Other' and other_support:
                support = f"Other: {other_support}"
            items.append({
                'recommendation': recommendation,
                'actions_needed': action,
                'support': support,
            })
    return items

def read_course_report_optional_details():
    location = compact_text(request.form.get('course_location') or '')
    if location not in {'main', 'branch'}:
        location = ''
    return {
        'course_instructor': compact_text(request.form.get('course_instructor') or ''),
        'course_coordinator': compact_text(request.form.get('course_coordinator') or ''),
        'location': location,
        'branch_name': compact_text(request.form.get('branch_name') or ''),
        'sections_count': compact_text(request.form.get('sections_count') or ''),
        'students_started': compact_text(request.form.get('students_started') or ''),
        'students_completed': compact_text(request.form.get('students_completed') or ''),
        'report_date': datetime.now().strftime('%Y-%m-%d'),
    }

WORD_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WORD_XML_NS = 'http://www.w3.org/XML/1998/namespace'
WORD_R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WORD_WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WORD_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WORD_PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
DOCX_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
DOCX_CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

ET.register_namespace('w', WORD_W_NS)
ET.register_namespace('r', WORD_R_NS)
ET.register_namespace('wp', WORD_WP_NS)
ET.register_namespace('a', WORD_A_NS)
ET.register_namespace('pic', WORD_PIC_NS)

def word_tag(name):
    return f'{{{WORD_W_NS}}}{name}'

def word_element(name, attributes=None):
    return ET.Element(word_tag(name), attributes or {})

def word_r_tag(name):
    return f'{{{WORD_R_NS}}}{name}'

def word_block_text(element):
    return ''.join((text_node.text or '') for text_node in element.findall(f'.//{word_tag("t")}'))

def clean_xml_text(text):
    if not text:
        return ''
    # Remove XML control characters that cause lxml to throw ValueError during tostring
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', str(text))

def word_paragraph(text='', bold=False, color='', size='', alignment=''):
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
    
    # OOXML Spec: bidi must come before jc in pPr
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
    return paragraph

def word_cell(text='', bold=False, width='2400', fill='', color='', size='', alignment=''):
    cell = word_element('tc')
    cell_properties = word_element('tcPr')
    cell_properties.append(word_element('tcW', {word_tag('w'): str(width), word_tag('type'): 'dxa'}))
    if fill:
        cell_properties.append(word_element('shd', {
            word_tag('val'): 'clear',
            word_tag('color'): 'auto',
            word_tag('fill'): str(fill).lstrip('#'),
        }))
    cell.append(cell_properties)
    parts = str(text or '').split('\n') or ['']
    for part in parts:
        cell.append(word_paragraph(part, bold=bold, color=color, size=size, alignment=alignment))
    return cell

def word_row(values, header=False, alignment='', fill='', color='', size=''):
    row = word_element('tr')
    for value in values:
        row.append(word_cell(value, bold=header, alignment=alignment, fill=fill, color=color, size=size))
    return row

def word_image_paragraph(rel_id, width_emu=1300000, height_emu=850000, alignment='right'):
    paragraph = word_element('p')
    paragraph_properties = word_element('pPr')
    paragraph_properties.append(word_element('jc', {word_tag('val'): alignment}))
    paragraph.append(paragraph_properties)
    run = word_element('r')
    drawing_xml = f'''
    <w:drawing xmlns:w="{WORD_W_NS}" xmlns:r="{WORD_R_NS}" xmlns:wp="{WORD_WP_NS}" xmlns:a="{WORD_A_NS}" xmlns:pic="{WORD_PIC_NS}">
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{int(width_emu)}" cy="{int(height_emu)}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="1" name="Organization Logo"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="Organization Logo"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{rel_id}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="{int(width_emu)}" cy="{int(height_emu)}"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
    '''
    run.append(ET.fromstring(drawing_xml))
    paragraph.append(run)
    return paragraph

def docx_hex_color(value, fallback='#26365f'):
    color = normalize_brand_color(value or fallback)
    return color.lstrip('#').upper()

def docx_optional_hex_color(value, fallback=''):
    color = normalize_optional_brand_color(value or fallback)
    return color.lstrip('#').upper() if color else ''

def build_course_report_word_identity_blocks(course_info=None, branding=None, logo_rel_id=''):
    language = get_export_report_language() if has_request_context() else 'en'
    branding = apply_university_identity_colors(branding or get_report_branding())
    primary = docx_hex_color(branding.get('primary_color'))
    secondary = docx_optional_hex_color(branding.get('secondary_color'), branding.get('primary_color')) or primary

    blocks = []
    if logo_rel_id:
        blocks.append(word_image_paragraph(logo_rel_id, alignment='center'))
        blocks.append(word_paragraph(''))

    return blocks

def insert_course_report_word_identity(body, course_info=None, branding=None, logo_rel_id=''):
    blocks = build_course_report_word_identity_blocks(course_info, branding, logo_rel_id)
    if not blocks:
        return
    insert_index = 0
    for block in reversed(blocks):
        body.insert(insert_index, block)

def get_next_docx_relationship_id(input_docx):
    rels_path = 'word/_rels/document.xml.rels'
    try:
        rels_xml = input_docx.read(rels_path)
        root = ET.fromstring(rels_xml)
    except Exception:
        return 'rId100'
    max_id = 1
    for rel in root:
        rel_id = rel.get('Id') or ''
        match = re.match(r'rId(\d+)$', rel_id)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"

def update_docx_relationships_for_image(rels_bytes, rel_id, target):
    if rels_bytes:
        root = ET.fromstring(rels_bytes)
    else:
        root = ET.Element(f'{{{DOCX_RELS_NS}}}Relationships')
    relationship = ET.Element(f'{{{DOCX_RELS_NS}}}Relationship', {
        'Id': rel_id,
        'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'Target': target,
    })
    root.append(relationship)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def update_docx_content_types_for_image(content_types_bytes, ext):
    if not content_types_bytes:
        return content_types_bytes
    ext = ext.lower().lstrip('.')
    content_type = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png'}.get(ext)
    if not content_type:
        return content_types_bytes
    root = ET.fromstring(content_types_bytes)
    default_tag = f'{{{DOCX_CONTENT_TYPES_NS}}}Default'
    if not any((item.get('Extension') or '').lower() == ext for item in root.findall(default_tag)):
        root.append(ET.Element(default_tag, {'Extension': ext, 'ContentType': content_type}))
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def format_plo_value(value):
    if isinstance(value, (list, tuple)):
        codes = [normalize_plo_code(item) for item in value]
        return ', '.join(code for code in codes if code)
    codes = extract_plo_codes(value)
    return ', '.join(codes) if codes else compact_text(value)

def resolve_clo_plo_code(clo, clo_plos):
    if not isinstance(clo_plos, dict):
        return ''
    normalized_lookup = {
        re.sub(r'\s+', '', str(key or '').strip()).upper(): value
        for key, value in clo_plos.items()
    }
    candidates = [
        clo_number_key(clo),
        clo_number(clo),
        str(clo or '').strip(),
    ]
    for candidate in candidates:
        lookup_key = re.sub(r'\s+', '', str(candidate or '').strip()).upper()
        if lookup_key in normalized_lookup:
            return format_plo_value(normalized_lookup[lookup_key])
    return ''

def build_clo_course_report_rows(stats, course_info=None, language=None):
    rows = []
    clo_plos = (course_info or {}).get('clo_plos') or {}
    for clo, data in sorted_clo_items(stats):
        number = clo_number(clo)
        wording = clo_wording(clo)
        target = f"{float(data.get('target_pct', 0)):.2f}%"
        actual = f"{float(data.get('achievement_percentage', 0)):.2f}%"
        achieved = float(data.get('achievement_percentage', 0)) >= float(data.get('target_pct', 0))
        comment = ('تم تحقيق المستوى المستهدف' if achieved else 'لم يتم تحقيق المستوى الهدف') if language == 'ar' else ('Target achieved' if achieved else 'Below target')
        rows.append([
            f"{number} {wording}".strip(),
            resolve_clo_plo_code(clo, clo_plos),
            format_mapped_questions_for_report(data.get('questions', []), session.get('assessment_files') if has_request_context() else []),
            target,
            actual,
            comment,
        ])
    return rows

def build_clo_assessment_word_table(stats, course_info=None, language=None, branding=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    branding = apply_university_identity_colors(branding or (get_report_branding() if has_request_context() else {}))
    primary = docx_hex_color(branding.get('primary_color'))

    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '5000', word_tag('type'): 'pct'}))
    if language == 'ar':
        table_properties.append(word_element('bidiVisual'))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): '808080',
        }))
    table_properties.append(borders)
    table.append(table_properties)

    headers = (
        ['نواتج تعلم المقرر', 'نواتج البرنامج المرتبطة', 'طرق التقييم', 'المستوى المستهدف', 'المستوى الفعلي', 'التعليق على نتائج القياس']
        if language == 'ar'
        else ['Course Learning Outcomes (CLOs)', 'Related PLOs Code', 'Assessment Methods', 'Targeted Level', 'Actual Level', 'Comment on Assessment Results']
    )
    table.append(word_row(headers, header=True, fill=primary, color='FFFFFF'))
    for row in build_clo_course_report_rows(stats, course_info, language):
        table.append(word_row(row))
    return table

def build_grade_distribution_word_table(distribution, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): '808080',
        }))
    table_properties.append(borders)
    table.append(table_properties)
    table.append(word_row(['التقدير', 'العدد', 'النسبة'] if language == 'ar' else ['Grade', 'Count', 'Percentage'], header=True))
    for row in distribution.get('rows') or []:
        table.append(word_row([
            row.get('grade', ''),
            row.get('count', 0),
            f"{float(row.get('percentage', 0)):.2f}%",
        ]))
    status_summary = grade_distribution_pass_fail_summary(distribution)
    table.append(word_row([
        'ناجح' if language == 'ar' else 'Pass',
        status_summary['Pass']['count'],
        f"{status_summary['Pass']['percentage']:.2f}%",
    ]))
    table.append(word_row([
        'راسب' if language == 'ar' else 'Fail',
        status_summary['Fail']['count'],
        f"{status_summary['Fail']['percentage']:.2f}%",
    ]))
    table.append(word_row(['الإجمالي' if language == 'ar' else 'Total', distribution.get('total', 0), '100.00%' if distribution.get('total') else '0.00%'], header=True))
    return table

def grade_distribution_pass_fail_summary(distribution):
    distribution = distribution or {}
    raw_counts = distribution.get('counts') or {}
    counts = {grade: int(raw_counts.get(grade, 0) or 0) for grade in GRADE_ORDER}
    pass_count = sum(counts.get(grade, 0) for grade in GRADE_ORDER if grade != 'F') + int(raw_counts.get('Pass', 0) or 0)
    fail_count = int(counts.get('F', 0) or 0) + int(raw_counts.get('Fail', 0) or 0)
    total = int(distribution.get('total') or 0)
    if raw_counts.get('Pass') or raw_counts.get('Fail'):
        total = max(total, pass_count + fail_count)
    return {
        'Pass': {
            'count': pass_count,
            'percentage': round((pass_count / total * 100), 2) if total else 0,
        },
        'Fail': {
            'count': fail_count,
            'percentage': round((fail_count / total * 100), 2) if total else 0,
        },
    }

def build_student_grade_comment(distribution, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    distribution = distribution or {}
    total = int(distribution.get('total') or 0)
    if not total:
        return ''
    raw_counts = distribution.get('counts') or {}
    counts = {grade: int(raw_counts.get(grade, 0) or 0) for grade in GRADE_ORDER}
    status_summary = grade_distribution_pass_fail_summary(distribution)
    pass_count = status_summary['Pass']['count']
    fail_count = status_summary['Fail']['count']
    pass_pct = status_summary['Pass']['percentage']
    fail_pct = status_summary['Fail']['percentage']
    top_grade = max(GRADE_ORDER, key=lambda grade: (counts.get(grade, 0), -GRADE_ORDER.index(grade)))
    top_count = counts.get(top_grade, 0)
    top_pct = round((top_count / total * 100), 2) if total else 0
    high_count = counts.get('A+', 0) + counts.get('A', 0)
    high_pct = round((high_count / total * 100), 2) if total else 0
    at_risk_count = counts.get('D+', 0) + counts.get('D', 0) + counts.get('F', 0)
    at_risk_pct = round((at_risk_count / total * 100), 2) if total else 0

    if language == 'ar':
        if pass_pct >= 90:
            overall = 'كان أداء الطلاب العام قويًا.'
        elif pass_pct >= 75:
            overall = 'كان أداء الطلاب العام مرضيًا.'
        elif pass_pct >= 60:
            overall = 'كان أداء الطلاب العام مقبولًا مع وجود حاجة لتحسينات موجهة.'
        else:
            overall = 'كان أداء الطلاب العام دون المستوى المتوقع ويتطلب إجراءات تحسين.'
        return (
            f"من أصل {total} طالبًا، نجح {pass_count} طالبًا ({pass_pct:.2f}%) "
            f"ولم ينجح {fail_count} طالبًا ({fail_pct:.2f}%). {overall} كان التقدير الأكثر تكرارًا "
            f"{top_grade} ({top_count} طالبًا، {top_pct:.2f}%). مثّلت تقديرات الإنجاز المرتفع "
            f"(A+ و A) عدد {high_count} طالبًا ({high_pct:.2f}%)، بينما مثّلت التقديرات التي تحتاج متابعة "
            f"(D+ و D و F) عدد {at_risk_count} طالبًا ({at_risk_pct:.2f}%)."
        )

    if pass_pct >= 90:
        overall = 'overall student performance was strong.'
    elif pass_pct >= 75:
        overall = 'overall student performance was satisfactory.'
    elif pass_pct >= 60:
        overall = 'overall student performance was acceptable, with room for targeted improvement.'
    else:
        overall = 'overall student performance was below the expected level and requires improvement actions.'

    return (
        f"Out of {total} students, {pass_count} ({pass_pct:.2f}%) passed and "
        f"{fail_count} ({fail_pct:.2f}%) failed; {overall} The most frequent grade was "
        f"{top_grade} ({top_count} students, {top_pct:.2f}%). High achievement grades "
        f"(A+ and A) accounted for {high_count} students ({high_pct:.2f}%), while "
        f"at-risk grades (D+, D, and F) accounted for {at_risk_count} students ({at_risk_pct:.2f}%)."
    )

def build_template_grade_distribution_word_table(distribution):
    grades_and_statuses = GRADE_ORDER + ['Denied Entry', 'In Progress', 'Incomplete', 'Pass', 'Fail', 'Withdrawn']
    counts = distribution.get('counts') or {}
    percentages = {row.get('grade'): row.get('percentage', 0) for row in distribution.get('rows') or []}
    status_summary = grade_distribution_pass_fail_summary(distribution)

    def grade_count_value(grade):
        if grade in GRADE_ORDER:
            return counts.get(grade, '')
        if grade in status_summary:
            return status_summary[grade]['count']
        return ''

    def grade_percentage_value(grade):
        if not distribution.get('total'):
            return ''
        if grade in GRADE_ORDER:
            return f"{float(percentages.get(grade, 0)):.2f}%"
        if grade in status_summary:
            return f"{status_summary[grade]['percentage']:.2f}%"
        return ''

    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): '808080',
        }))
    table_properties.append(borders)
    table.append(table_properties)
    table.append(word_row(['Grades / Status Distributions'] + grades_and_statuses, header=True))
    table.append(word_row(['Number of Students'] + [grade_count_value(grade) for grade in grades_and_statuses]))
    table.append(word_row(['Percentage'] + [grade_percentage_value(grade) for grade in grades_and_statuses]))
    return table

def course_report_stats_summary_for_ai(stats):
    rows = []
    for clo, data in sorted_clo_items(stats or {}):
        rows.append({
            'clo': clo_number(clo),
            'wording': clo_wording(clo),
            'target_percentage': round(float(data.get('target_pct', 0) or 0), 2),
            'achievement_percentage': round(float(data.get('achievement_percentage', 0) or 0), 2),
            'students_achieved': int(data.get('students_achieved', 0) or 0),
            'total_possible_score': round(float(data.get('total_possible_score', 0) or 0), 2),
        })
    return rows

def course_report_grade_summary_for_ai(distribution):
    distribution = distribution or {}
    return {
        'total_students': int(distribution.get('total') or 0),
        'grade_counts': distribution.get('counts') or {},
        'grade_percentages': {
            row.get('grade'): row.get('percentage', 0)
            for row in distribution.get('rows') or []
        },
        'pass_fail': grade_distribution_pass_fail_summary(distribution),
    }

def course_report_ai_system_prompt(language):
    if language == 'ar':
        return (
            "أنت مساعد أكاديمي يكتب نصوصًا مختصرة لتقرير مقرر. "
            "اكتب بالعربية الأكاديمية الواضحة. أعد JSON فقط دون Markdown. "
            "لا تخترع أرقامًا غير موجودة. اجعل التعليق على نتائج الطلاب فقرة واحدة، "
            "واجعل التوصيات عملية ومبنية على تحقق نواتج التعلم."
        )
    return (
        "You are an academic reporting assistant. Write concise, formal course report text. "
        "Return JSON only with no Markdown. Do not invent numbers. Provide one paragraph for "
        "student results and practical recommendations based on CLO attainment."
    )

def generate_course_report_ai_insights(stats, grade_distribution, course_info=None, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    payload = {
        'language': 'Arabic' if language == 'ar' else 'English',
        'course': {
            'name': (course_info or {}).get('course_name') or (course_info or {}).get('raw_name') or '',
            'code': (course_info or {}).get('course_id') or '',
        },
        'grade_distribution': course_report_grade_summary_for_ai(grade_distribution),
        'clo_attainment': course_report_stats_summary_for_ai(stats),
        'required_json_schema': {
            'student_results_comment': 'one formal paragraph',
            'recommendations': [
                {
                    'recommendation': 'short recommendation',
                    'actions_needed': 'short action',
                    'support': 'short support or No additional support required',
                }
            ],
        },
    }
    user_payload = json.dumps(payload, ensure_ascii=False)
    parsed = call_gemini_json(course_report_ai_system_prompt(language), user_payload)
    if not parsed:
        parsed = call_groq_json(course_report_ai_system_prompt(language), user_payload)
    if not isinstance(parsed, dict):
        return {}

    comment = compact_text(parsed.get('student_results_comment') or parsed.get('student_comment') or '')
    recommendations = []
    for item in parsed.get('recommendations') or []:
        if not isinstance(item, dict):
            continue
        recommendation = compact_text(item.get('recommendation') or item.get('text') or '')
        if not recommendation:
            continue
        recommendations.append({
            'recommendation': recommendation,
            'actions_needed': compact_text(item.get('actions_needed') or item.get('action') or ''),
            'support': compact_text(item.get('support') or item.get('needed_support') or ''),
        })
    return {
        'student_results_comment': comment,
        'recommendations': recommendations[:5],
        'source': 'gemini' if GEMINI_API_KEY else 'groq',
    }

def fallback_course_report_recommendations(stats, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    stats_items = sorted_clo_items(stats or {})
    below_target = [
        clo_number(clo)
        for clo, data in stats_items
        if float(data.get('achievement_percentage', 0) or 0) < float(data.get('target_pct', 0) or 0)
    ]
    if language == 'ar':
        if below_target:
            return [{
                'recommendation': f"تحسين أداء الطلاب في نواتج التعلم التي لم تحقق المستوى المستهدف ({', '.join(below_target)}).",
                'actions_needed': 'مراجعة استراتيجيات التدريس وتوفير أنشطة دعم وتدريب إضافية.',
                'support': 'متابعة من القسم أو لجنة البرنامج.',
            }]
        return [{
            'recommendation': 'المحافظة على مستوى تحقق نواتج التعلم وتعزيز الممارسات التعليمية الفاعلة.',
            'actions_needed': 'توظيف أنشطة إثرائية ومتابعة مؤشرات الأداء في الطرح القادم.',
            'support': 'لا يتطلب دعمًا إضافيًا.',
        }]
    if below_target:
        return [{
            'recommendation': f"Improve student performance in CLOs below the target level ({', '.join(below_target)}).",
            'actions_needed': 'Revise teaching strategies and provide additional support activities.',
            'support': 'Department or program committee follow-up.',
        }]
    return [{
        'recommendation': 'Maintain CLO attainment performance and reinforce effective teaching practices.',
        'actions_needed': 'Use enrichment activities and monitor performance indicators in the next offering.',
        'support': 'No additional support required.',
    }]

def merge_course_report_recommendations(selected_items, generated_items):
    merged = []
    seen = set()
    for item in list(selected_items or []) + list(generated_items or []):
        if not isinstance(item, dict):
            continue
        recommendation = compact_text(item.get('recommendation') or '')
        if not recommendation:
            continue
        key = re.sub(r'\s+', ' ', recommendation).strip().lower()
        if key in seen:
            continue
        seen.add(key)
        merged.append({
            'recommendation': recommendation,
            'actions_needed': compact_text(item.get('actions_needed') or ''),
            'support': compact_text(item.get('support') or ''),
        })
    return merged

def build_uncovered_topics_word_table(uncovered_details, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): '808080',
        }))
    table_properties.append(borders)
    table.append(table_properties)
    table.append(word_row(
        ['الموضوع', 'سبب عدم التغطية أو الاختلاف', 'مدى التأثير على نواتج التعلم', 'الإجراء التعويضي']
        if language == 'ar'
        else ['Topic', 'Reason for Not Covering/discrepancies', 'Extent of their Impact on Learning Outcomes', 'Compensating Action'],
        header=True
    ))
    for item in uncovered_details or []:
        table.append(word_row([
            item.get('topic', ''),
            course_report_label_for_language(item.get('reason', ''), language),
            item.get('impact', ''),
            course_report_label_for_language(item.get('action', ''), language),
        ]))
    return table

def build_course_improvement_plan_word_table(improvement_items, language=None):
    language = language or (get_export_report_language() if has_request_context() else 'en')
    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): '808080',
        }))
    table_properties.append(borders)
    table.append(table_properties)
    alignment = 'right' if language == 'ar' else ''
    table.append(word_row(
        ['التوصيات', 'الإجراءات المطلوبة', 'الدعم المطلوب'] if language == 'ar' else ['Recommendations', 'Actions Needed', 'Support'],
        header=True,
        alignment=alignment
    ))
    for item in improvement_items or []:
        table.append(word_row([
            course_report_label_for_language(item.get('recommendation', ''), language),
            course_report_label_for_language(item.get('actions_needed', ''), language),
            course_report_label_for_language(item.get('support', ''), language),
        ], alignment=alignment))
    return table

def build_course_report_input_blocks(course_report_inputs):
    if not course_report_inputs:
        return []

    language = get_export_report_language() if has_request_context() else 'en'
    blocks = []
    topics_value = course_report_inputs.get('topics_covered')
    topics_label = 'Yes' if topics_value == 'yes' else 'No' if topics_value == 'no' else 'Not specified'
    uncovered_details = course_report_inputs.get('uncovered_topic_details') or []
    blocks.append(word_paragraph('Topics Coverage', bold=True))
    blocks.append(word_paragraph(f"Have all course topics been covered? {topics_label}"))
    if topics_value == 'no' and uncovered_details:
        blocks.append(word_paragraph('Topics not covered:'))
        blocks.append(build_uncovered_topics_word_table(uncovered_details, language))
    grade_distribution = course_report_inputs.get('grade_distribution') or {}
    if grade_distribution:
        blocks.append(word_paragraph('Final Grade Distribution', bold=True))
        blocks.append(build_grade_distribution_word_table(grade_distribution, language))
        student_grade_comment = course_report_inputs.get('student_results_comment') or build_student_grade_comment(grade_distribution, language)
        if student_grade_comment:
            blocks.append(word_paragraph('Comment on Student Grades', bold=True))
            blocks.append(word_paragraph(student_grade_comment))
    improvement_items = course_report_inputs.get('course_improvement_plan') or []
    if improvement_items:
        blocks.append(word_paragraph('Course Improvement Plan', bold=True))
        blocks.append(build_course_improvement_plan_word_table(improvement_items, language))
    return blocks

def replace_word_text(root, replacements):
    for text_node in root.findall(f'.//{word_tag("t")}'):
        value = text_node.text or ''
        for old, new in replacements.items():
            if old in value:
                value = value.replace(old, str(new or ''))
        text_node.text = value
    for paragraph in root.findall(f'.//{word_tag("p")}'):
        text_nodes = paragraph.findall(f'.//{word_tag("t")}')
        if not text_nodes:
            continue
        combined = ''.join(text_node.text or '' for text_node in text_nodes)
        updated = combined
        for old, new in replacements.items():
            updated = updated.replace(old, str(new or ''))
        if updated != combined:
            text_nodes[0].set(f'{{{WORD_XML_NS}}}space', 'preserve')
            text_nodes[0].text = updated
            for text_node in text_nodes[1:]:
                text_node.text = ''

def find_body_child_index(body, needle, start=0):
    children = list(body)
    fallback_index = None
    for index in range(start, len(children)):
        text = re.sub(r'\s+', ' ', word_block_text(children[index])).strip()
        if text == needle or text.startswith(needle):
            return index
        if needle in text and fallback_index is None:
            fallback_index = index
    return fallback_index

def replace_next_table_after_heading(body, heading_text, table_element):
    children = list(body)
    heading_options = list(heading_text) if isinstance(heading_text, (list, tuple)) else [heading_text]
    heading_index = None
    matched_heading = heading_options[0] if heading_options else ''
    for heading_option in heading_options:
        heading_index = find_body_child_index(body, heading_option)
        if heading_index is not None:
            matched_heading = heading_option
            break
    if heading_index is None:
        raise ValueError(f"Could not find '{matched_heading}' in the Word template.")
    for index in range(heading_index + 1, len(children)):
        if children[index].tag.split('}')[-1] == 'tbl':
            body.remove(children[index])
            body.insert(index, table_element)
            return
    raise ValueError(f"Could not find a table after '{matched_heading}' in the Word template.")

def course_report_template_replacements(course_info=None, course_report_inputs=None, total_students=None):
    course_info = course_info or {}
    course_report_inputs = course_report_inputs or {}
    grade_distribution = (course_report_inputs or {}).get('grade_distribution') or {}
    completed_students = grade_distribution.get('total') or total_students or ''
    report_details = course_report_inputs.get('report_details') or {}
    student_grade_comment = course_report_inputs.get('student_results_comment') or build_student_grade_comment(grade_distribution)
    branding = get_report_branding() if has_request_context() else {}
    user = current_user() if has_request_context() else None
    institution = (user['university_name'] if user else '') or branding.get('organization_name') or ''
    department = course_info.get('department') or branding.get('department') or ''
    location = report_details.get('location') or ''
    branch_name = report_details.get('branch_name') or ''
    if get_export_report_language() == 'ar':
        location_text = 'مكان تقديم المقرر: ☒ المقر الرئيس ☐       فرع ...' if location == 'main' else f"مكان تقديم المقرر: ☐ المقر الرئيس ☒       فرع {branch_name}".rstrip()
    else:
        location_text = 'Location:   Main campus ☒             branch ☐' if location == 'main' else f"Location:   Main campus ☐             branch ☒ {branch_name}".rstrip()
    started_students = report_details.get('students_started') or completed_students
    completed_students = report_details.get('students_completed') or completed_students
    return {
        'Enter Course Title.': course_info.get('course_name') or course_info.get('raw_name') or '',
        'Enter Course Code.': course_info.get('course_id') or '',
        'Enter Department Name.': department,
        'Enter Program Name.': course_info.get('program') or '',
        'Enter College Name.': course_info.get('college') or '',
        'Enter Institution Name.': institution,
        'Enter Academic Year.': course_info.get('academic_year') or str(datetime.now().year),
        'Enter Course Instructor Name.': report_details.get('course_instructor') or course_info.get('instructor') or '',
        'Course Coordinator:': f"Course Coordinator: {report_details.get('course_coordinator') or ''}",
        'Location:   Main campus ☐             branch ☐': location_text,
        'Number of Section(s):': f"Number of Section(s): {report_details.get('sections_count') or ''}",
        'Enter Number of Students Starting the Course.': started_students,
        'Enter Number of Students Completed the Course.': completed_students,
        'Pick Report Date.': datetime.now().strftime('%Y-%m-%d'),
        'Including particular factors (if any) affecting the results': student_grade_comment,
        'اسم المقرر:   اكتب هنا': f"اسم المقرر:   {course_info.get('course_name') or course_info.get('raw_name') or ''}",
        'رمز المقرر:  اكتب هنا': f"رمز المقرر:  {course_info.get('course_id') or ''}",
        'أستاذ المقرر:  اكتب هنا': f"أستاذ المقرر:  {report_details.get('course_instructor') or course_info.get('instructor') or ''}",
        'منسق المقرر:  اكتب هنا': f"منسق المقرر:  {report_details.get('course_coordinator') or ''}",
        'مكان تقديم المقرر: ☐ المقر الرئيس ☐       فرع ...': location_text,
        'عدد الشعب:  اكتب هنا': f"عدد الشعب:  {report_details.get('sections_count') or ''}",
        'عدد الطلاب (الذين بدأوا المقرر):  اكتب هنا': f"عدد الطلاب (الذين بدأوا المقرر):  {started_students}",
        'عدد الطلاب (الذين أنهوا المقرر):  اكتب هنا': f"عدد الطلاب (الذين أنهوا المقرر):  {completed_students}",
        'تاريخ إعداد التقرير:  اكتب هنا': f"تاريخ إعداد التقرير:  {datetime.now().strftime('%Y-%m-%d')}",
        'متضمنًا العوامل التي أثرت على النتائج - إن وجدت-.': student_grade_comment,
    }

def fill_course_report_docx(template_bytes, stats, course_report_inputs=None, course_info=None, total_students=None):
    source = io.BytesIO(template_bytes)
    output = io.BytesIO()
    try:
        with zipfile.ZipFile(source, 'r') as input_docx:
            branding = apply_university_identity_colors(get_report_branding()) if has_request_context() else {}
            logo_path = resolve_branding_logo_path(branding, report_ready=True) if branding else ''
            logo_rel_id = ''
            logo_media_zip_path = ''
            logo_media_target = ''
            logo_bytes = b''
            if logo_path and os.path.exists(logo_path):
                logo_ext = os.path.splitext(logo_path)[1].lower()
                if logo_ext in {'.jpg', '.jpeg', '.png'}:
                    logo_rel_id = get_next_docx_relationship_id(input_docx)
                    base_media_name = f'organization_identity_logo{logo_ext}'
                    logo_media_zip_path = f'word/media/{base_media_name}'
                    existing_names = set(input_docx.namelist())
                    suffix = 1
                    while logo_media_zip_path in existing_names:
                        base_media_name = f'organization_identity_logo_{suffix}{logo_ext}'
                        logo_media_zip_path = f'word/media/{base_media_name}'
                        suffix += 1
                    logo_media_target = f'media/{base_media_name}'
                    with open(logo_path, 'rb') as logo_file:
                        logo_bytes = logo_file.read()

            document_xml = input_docx.read('word/document.xml')
            root = ET.fromstring(document_xml)
            body = root.find(word_tag('body'))
            if body is None:
                raise ValueError("The Word template does not contain a document body.")

            replace_word_text(root, course_report_template_replacements(course_info, course_report_inputs, total_students))
            insert_course_report_word_identity(body, course_info, branding, logo_rel_id)
            report_language = get_export_report_language()
            grade_distribution = (course_report_inputs or {}).get('grade_distribution') or {}
            replace_next_table_after_heading(
                body,
                ['1. Grade Distribution', '1. توزيع التقديرات'],
                build_template_grade_distribution_word_table(grade_distribution)
            )
            replace_next_table_after_heading(
                body,
                ['1. Course Learning Outcomes Assessment Results', '1. قياس نواتج التعلم للمقرر'],
                build_clo_assessment_word_table(stats, course_info, report_language)
            )
            uncovered_details = (course_report_inputs or {}).get('uncovered_topic_details') or []
            replace_next_table_after_heading(
                body,
                ['C. Topics not covered.', 'ج. الموضوعات التي لم يتم تغطيتها'],
                build_uncovered_topics_word_table(uncovered_details, report_language)
            )
            improvement_items = (course_report_inputs or {}).get('course_improvement_plan') or []
            replace_next_table_after_heading(
                body,
                ['2. Recommendations', '2. التوصيات:'],
                build_course_improvement_plan_word_table(improvement_items, report_language)
            )
            replace_next_table_after_heading(
                body,
                ['D. Course Improvement Plan', 'F. Course Improvement Plan', 'F. Course Improvement Plan (if any)', 'خطة تحسين المقرر', 'خطط التحسين'],
                build_course_improvement_plan_word_table(improvement_items, report_language)
            )
            updated_document_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)

            with zipfile.ZipFile(output, 'w') as output_docx:
                rels_written = False
                content_types_written = False
                rels_path = 'word/_rels/document.xml.rels'
                content_types_path = '[Content_Types].xml'
                for item in input_docx.infolist():
                    if item.filename == 'word/document.xml':
                        output_docx.writestr(item, updated_document_xml)
                    elif logo_rel_id and item.filename == rels_path:
                        output_docx.writestr(
                            item,
                            update_docx_relationships_for_image(input_docx.read(item.filename), logo_rel_id, logo_media_target)
                        )
                        rels_written = True
                    elif logo_rel_id and item.filename == content_types_path:
                        output_docx.writestr(
                            item,
                            update_docx_content_types_for_image(input_docx.read(item.filename), os.path.splitext(logo_media_zip_path)[1])
                        )
                        content_types_written = True
                    else:
                        output_docx.writestr(item, input_docx.read(item.filename))
                if logo_rel_id and not rels_written:
                    output_docx.writestr(
                        rels_path,
                        update_docx_relationships_for_image(b'', logo_rel_id, logo_media_target)
                    )
                if logo_rel_id and not content_types_written:
                    output_docx.writestr(
                        content_types_path,
                        update_docx_content_types_for_image(b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''', os.path.splitext(logo_media_zip_path)[1])
                    )
                if logo_rel_id and logo_bytes:
                    output_docx.writestr(logo_media_zip_path, logo_bytes)
    except zipfile.BadZipFile as exc:
        raise ValueError("Please upload a valid DOCX Word template.") from exc

    output.seek(0)
    return output.getvalue()

def build_generated_course_report_docx(stats, course_report_inputs=None, course_info=None):
    root = word_element('document')
    body = word_element('body')
    root.append(body)

    course_info = course_info or {}
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
    language = get_export_report_language() if has_request_context() else 'en'
    body.append(word_paragraph('تقرير المقرر' if language == 'ar' else 'Course Report', bold=True))
    
    report_details = (course_report_inputs or {}).get('report_details') or {}
    course_name = course_info.get('course_name') or course_info.get('raw_name') or ''
    course_id = course_info.get('course_id') or ''
    
    def add_detail(label_en, label_ar, value):
        if value:
            body.append(word_paragraph(f"{label_ar if language == 'ar' else label_en}: {value}"))

    add_detail('Course Name', 'اسم المقرر', course_name)
    add_detail('Course Code', 'رمز المقرر', course_id)
    add_detail('Course Instructor', 'أستاذ المقرر', report_details.get('course_instructor'))
    add_detail('Course Coordinator', 'منسق المقرر', report_details.get('course_coordinator'))
    
    location = report_details.get('location')
    if location == 'main':
        add_detail('Location', 'المقر', 'المقر الرئيسي' if language == 'ar' else 'Main Campus')
    elif location == 'branch':
        branch_name = report_details.get('branch_name')
        val = f"{'الفرع' if language == 'ar' else 'Branch'}: {branch_name}" if branch_name else ('الفرع' if language == 'ar' else 'Branch')
        add_detail('Location', 'المقر', val)
        
    add_detail('Number of Sections', 'عدد الشعب', report_details.get('sections_count'))
    add_detail('Students Started', 'الطلاب في بداية الفصل', report_details.get('students_started'))
    add_detail('Students Completed', 'الطلاب الذين أكملوا المقرر', report_details.get('students_completed'))

    body.append(word_paragraph('نتائج تقييم نواتج التعلم للمقرر' if language == 'ar' else 'Course Learning Outcomes Assessment Results', bold=True))
    body.append(build_clo_assessment_word_table(stats, course_info, get_export_report_language() if has_request_context() else 'en'))
    for block in build_course_report_input_blocks(course_report_inputs):
        body.append(block)

    body.append(word_element('sectPr'))
    document_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
        content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
        content_types.append('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        content_types.append('  <Default Extension="xml" ContentType="application/xml"/>')
        
        if logo_bytes:
            ext_clean = logo_ext.lstrip('.')
            ctype = 'image/png' if ext_clean == 'png' else 'image/jpeg'
            content_types.append(f'  <Default Extension="{ext_clean}" ContentType="{ctype}"/>')
            docx.writestr(f'word/media/logo{logo_ext}', logo_bytes)
            
        content_types.append('  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>')
        content_types.append('</Types>')
        docx.writestr('[Content_Types].xml', '\n'.join(content_types))

        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        rels.append('  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>')
        rels.append('</Relationships>')
        docx.writestr('_rels/.rels', '\n'.join(rels))

        if logo_bytes:
            doc_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            doc_rels.append(f'  <Relationship Id="{logo_rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo{logo_ext}"/>')
            doc_rels.append('</Relationships>')
            docx.writestr('word/_rels/document.xml.rels', '\n'.join(doc_rels))
        
        docx.writestr('word/document.xml', document_xml)

    output.seek(0)
    return output.getvalue()

def build_simple_word_document(body_elements):
    root = word_element('document')
    body = word_element('body')
    root.append(body)
    for element in body_elements:
        body.append(element)
    body.append(word_element('sectPr'))
    document_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')
        docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
        docx.writestr('word/document.xml', document_xml)
    output.seek(0)
    return output.getvalue()

def build_clo_results_docx(stats, total_students=0, course_info=None, student_achievement_matrix=None):
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

    title = 'تقرير تحقق نواتج التعلم' if language == 'ar' else 'CLO Attainment Report'
    course_label = 'المقرر' if language == 'ar' else 'Course'
    total_label = 'إجمالي الطلاب' if language == 'ar' else 'Total Students'
    course_name = course_info.get('course_name') or course_info.get('raw_name') or ''
    course_id = course_info.get('course_id') or course_info.get('course_code') or ''
    course_text = f"{course_name} ({course_id})" if course_name and course_id else (course_name or course_id or '-')
    
    body.append(word_paragraph(title, bold=True))
    body.append(word_paragraph(f"{course_label}: {course_text}"))
    body.append(word_paragraph(f"{total_label}: {total_students or 0}"))
    body.append(word_paragraph(''))

    body.append(word_paragraph('\u062a\u0639\u0631\u064a\u0641 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645' if language == 'ar' else 'CLO Definitions', bold=True))
    clo_definitions = build_clo_definitions(stats.keys())
    
    primary = docx_hex_color(branding.get('primary_color'))

    def_table = word_element('tbl')
    def_table_props = word_element('tblPr')
    def_table_props.append(word_element('tblW', {word_tag('w'): '5000', word_tag('type'): 'pct'}))
    if language == 'ar':
        def_table_props.append(word_element('bidiVisual'))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {word_tag('val'): 'single', word_tag('sz'): '6', word_tag('space'): '0', word_tag('color'): '808080'}))
    def_table_props.append(borders)
    def_table.append(def_table_props)
    
    def_headers = ['\u0627\u0644\u0645\u062c\u0627\u0644', '\u0627\u0644\u0631\u0645\u0632', '\u0627\u0644\u0646\u0635'] if language == 'ar' else ['Domain', 'CLO', 'Wording']
    def_table.append(word_row(def_headers, header=True, fill=primary, color='FFFFFF'))
    for item in clo_definitions:
        domain_text = localized_clo_domain(item['domain'], language)
        def_table.append(word_row([domain_text, item['number'], item['wording']]))
    body.append(def_table)
    body.append(word_paragraph(''))

    body.append(word_paragraph('\u0645\u0644\u062e\u0635 \u062a\u062d\u0642\u0642 \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645' if language == 'ar' else 'CLO Achievement Summary', bold=True))
    body.append(build_clo_assessment_word_table(stats or {}, course_info, language))
    
    if student_achievement_matrix and student_achievement_matrix.get('students'):
        body.append(word_paragraph(''))
        body.append(word_paragraph('\u0625\u0646\u062c\u0627\u0632 \u0627\u0644\u0637\u0644\u0627\u0628 \u0641\u064a \u0646\u0648\u0627\u062a\u062c \u0627\u0644\u062a\u0639\u0644\u0645' if language == 'ar' else 'Student CLO Achievement', bold=True))
        matrix_table = word_element('tbl')
        matrix_table_props = word_element('tblPr')
        matrix_table_props.append(word_element('tblW', {word_tag('w'): '5000', word_tag('type'): 'pct'}))
        if language == 'ar':
            matrix_table_props.append(word_element('bidiVisual'))
        mborders = word_element('tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            mborders.append(word_element(border_name, {word_tag('val'): 'single', word_tag('sz'): '6', word_tag('space'): '0', word_tag('color'): '808080'}))
        matrix_table_props.append(mborders)
        matrix_table.append(matrix_table_props)
        
        clos = student_achievement_matrix.get('clos') or []
        matrix_headers = ['\u0627\u0644\u0631\u0642\u0645 \u0627\u0644\u062c\u0627\u0645\u0639\u064a' if language == 'ar' else 'Student ID'] + [clo_number(c) for c in clos]
        matrix_table.append(word_row(matrix_headers, header=True, fill=primary, color='FFFFFF'))
        
        cells = student_achievement_matrix.get('cells') or {}
        for student_id in student_achievement_matrix.get('students', []):
            row_data = [display_student_id(student_id)]
            for clo in clos:
                cell = cells.get(student_id, {}).get(clo)
                if cell:
                    status = '\u0645\u062a\u062d\u0642\u0642' if cell.get('achieved') and language == 'ar' else '\u063a\u064a\u0631 \u0645\u062a\u062d\u0642\u0642' if language == 'ar' else 'Met' if cell.get('achieved') else 'Not met'
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
        content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
        content_types.append('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        content_types.append('  <Default Extension="xml" ContentType="application/xml"/>')
        
        if logo_bytes:
            ext_clean = logo_ext.lstrip('.')
            ctype = 'image/png' if ext_clean == 'png' else 'image/jpeg'
            content_types.append(f'  <Default Extension="{ext_clean}" ContentType="{ctype}"/>')
            docx.writestr(f'word/media/logo{logo_ext}', logo_bytes)
            
        content_types.append('  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>')
        content_types.append('</Types>')
        docx.writestr('[Content_Types].xml', '\n'.join(content_types))

        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        rels.append('  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>')
        rels.append('</Relationships>')
        docx.writestr('_rels/.rels', '\n'.join(rels))

        if logo_bytes:
            doc_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            doc_rels.append(f'  <Relationship Id="{logo_rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo{logo_ext}"/>')
            doc_rels.append('</Relationships>')
            docx.writestr('word/_rels/document.xml.rels', '\n'.join(doc_rels))
        
        docx.writestr('word/document.xml', document_xml)

    output.seek(0)
    return output.getvalue()

def build_exam_mapping_docx(payload, title='', course_name='', filename=''):
    payload = payload or {}
    language = get_export_report_language() if has_request_context() else 'en'
    report_title = 'تقرير موائمة التقييم' if language == 'ar' else 'Assessment Alignment Report'
    course_label = 'المقرر' if language == 'ar' else 'Course'
    source_label = 'المصدر' if language == 'ar' else 'Source'
    type_label = 'نوع السؤال' if language == 'ar' else 'Question Type'
    question_label = 'السؤال' if language == 'ar' else 'Question'
    text_label = 'نص السؤال' if language == 'ar' else 'Question Text'
    clo_label = 'ناتج التعلم' if language == 'ar' else 'Mapped CLOs'
    elements = []
    
    branding = apply_university_identity_colors(get_report_branding())
    labels = pdf_report_labels(language)
    organization_display_name = localized_university_name(branding.get('organization_name'), language) or labels['na']
    
    if organization_display_name and organization_display_name != labels['na']:
        elements.append(word_paragraph(f"{labels['university']}: {organization_display_name}", bold=True))
    if branding.get('department'):
        elements.append(word_paragraph(f"{labels['department']}: {branding.get('department')}"))
            
    elements.append(word_paragraph(report_title, bold=True))
    if course_name:
        elements.append(word_paragraph(f"{course_label}: {course_name}"))
    if filename:
        elements.append(word_paragraph(f"{source_label}: {filename}"))
    elements.append(word_paragraph(''))
    
    # Add CLO definitions table
    clos = get_course_clos(course_name)
    if clos:
        clo_defs = build_clo_definitions(clos)
        labels = pdf_report_labels(language)
        elements.append(word_paragraph(labels['clo_definitions'], bold=True))
        def_table = word_element('tbl')
        def_table_props = word_element('tblPr')
        def_table_props.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
        def_borders = word_element('tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            def_borders.append(word_element(border_name, {
                word_tag('val'): 'single',
                word_tag('sz'): '6',
                word_tag('space'): '0',
                word_tag('color'): 'cbd5e1',
            }))
        def_table_props.append(def_borders)
        def_table.append(def_table_props)
        def_table.append(word_row([labels['domain'], labels['clo'], labels['wording']], header=True))
        for item in clo_defs:
            def_table.append(word_row([
                localized_clo_domain(item['domain'], language),
                item['number'],
                item['wording']
            ]))
        elements.append(def_table)
        elements.append(word_paragraph(''))
        elements.append(word_paragraph('تفاصيل الموائمة' if language == 'ar' else 'Alignment Details', bold=True))

    table = word_element('tbl')
    table_properties = word_element('tblPr')
    table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
    borders = word_element('tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(word_element(border_name, {
            word_tag('val'): 'single',
            word_tag('sz'): '6',
            word_tag('space'): '0',
            word_tag('color'): 'cbd5e1',
        }))
    table_properties.append(borders)
    table.append(table_properties)
    table.append(word_row([question_label, type_label, text_label, clo_label], header=True))
    for index, item in enumerate(payload.get('questions') or [], start=1):
        clos = '\n'.join(clo_number(clo) or str(clo or '') for clo in item.get('clos') or [])
        table.append(word_row([
            f"{question_label} {index}",
            item.get('question_type') or item.get('type') or '-',
            item.get('question_text') or item.get('text') or '',
            clos or '-',
        ]))
    elements.append(table)
    
    # --- Matrix Generation ---
    matrix_data = compute_exam_alignment_matrix(payload)
    unique_clos = matrix_data['unique_clos']
    
    if unique_clos:
        elements.append(word_paragraph(''))
        matrix_title = 'مصفوفة الموائمة' if language == 'ar' else 'Alignment Matrix'
        elements.append(word_paragraph(matrix_title, bold=True))
        
        m_table = word_element('tbl')
        m_table_properties = word_element('tblPr')
        m_table_properties.append(word_element('tblW', {word_tag('w'): '0', word_tag('type'): 'auto'}))
        m_borders = word_element('tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            m_borders.append(word_element(border_name, {
                word_tag('val'): 'single', word_tag('sz'): '6', word_tag('space'): '0', word_tag('color'): 'cbd5e1'
            }))
        m_table_properties.append(m_borders)
        m_table.append(m_table_properties)
        
        m_header = [question_label] + unique_clos
        m_table.append(word_row(m_header, header=True))
        
        for row_info in matrix_data['rows']:
            m_row = [f"{question_label} {row_info['index']}"]
            for clo in unique_clos:
                m_row.append("✓" if row_info['clos'].get(clo) else "")
            m_table.append(word_row(m_row))
            
        count_label = 'الإجمالي' if language == 'ar' else 'Total Count'
        count_row = [count_label]
        for clo in unique_clos:
            count_row.append(str(matrix_data['totals'].get(clo, 0)))
        m_table.append(word_row(count_row, header=True))
        
        perc_label = 'النسبة من الإجمالي' if language == 'ar' else 'Percentage'
        perc_row = [perc_label]
        for clo in unique_clos:
            perc_row.append(f"{matrix_data['percentages'].get(clo, 0)}%")
        m_table.append(word_row(perc_row, header=True))
        
        elements.append(m_table)

    return build_simple_word_document(elements)

def build_course_report_docx(stats, course_report_inputs=None, course_info=None, total_students=None):
    template_path = COURSE_REPORT_TEMPLATE_PATH_AR if get_export_report_language() == 'ar' else COURSE_REPORT_TEMPLATE_PATH_EN
    if os.path.exists(template_path):
        with open(template_path, 'rb') as template_file:
            return fill_course_report_docx(
                template_file.read(),
                stats,
                course_report_inputs,
                course_info,
                total_students
            )
    return build_generated_course_report_docx(stats, course_report_inputs, course_info)

def read_course_report_export_inputs(redirect_url, stats=None, course_info=None, total_students=None):
    grade_distribution = None
    final_grades_file = request.files.get('final_grades_file')
    if final_grades_file and final_grades_file.filename:
        final_grades_ext = os.path.splitext(final_grades_file.filename)[1].lower()
        if final_grades_ext not in {'.csv', '.xlsx', '.xls', '.pdf'}:
            flash("Final grades file must be CSV, Excel, or PDF.", "error")
            return None, redirect(redirect_url)
        grade_filepath = get_upload_path(f"{uuid.uuid4()}{final_grades_ext}")
        final_grades_file.save(grade_filepath)
        try:
            grade_distribution = parse_final_grade_distribution(grade_filepath, final_grades_ext)
        except Exception as e:
            app.logger.warning("Could not read final grades file: %s", e)
            grade_distribution = {}
        finally:
            try:
                os.remove(grade_filepath)
            except OSError:
                pass
    else:
        grade_distribution = session.get('temp_grade_distribution')
        if grade_distribution is None:
            flash("Please upload the final grades file as CSV, Excel, or PDF.", "error")
            return None, redirect(redirect_url)

    topics_covered = (request.form.get('topics_covered') or '').strip()
    if topics_covered not in {'yes', 'no'}:
        flash("Please answer whether all course topics were covered.", "error")
        return None, redirect(redirect_url)

    if not grade_distribution.get('total'):
        grade_distribution = {}

    report_details = read_course_report_optional_details()
    if total_students:
        default_student_count = str(total_students)
        report_details['students_started'] = report_details.get('students_started') or default_student_count
        report_details['students_completed'] = report_details.get('students_completed') or default_student_count

    course_report_inputs = {
        'topics_covered': topics_covered,
        'uncovered_topic_details': read_uncovered_topic_details(),
        'grade_distribution': grade_distribution,
        'report_details': report_details,
        'course_improvement_plan': read_course_improvement_plan(),
    }

    language = get_export_report_language() if has_request_context() else 'en'
    ai_insights = generate_course_report_ai_insights(stats or {}, grade_distribution, course_info or {}, language)
    student_comment = compact_text(ai_insights.get('student_results_comment') or '')
    if not student_comment:
        student_comment = build_student_grade_comment(grade_distribution, language)
    generated_recommendations = ai_insights.get('recommendations') or fallback_course_report_recommendations(stats or {}, language)
    course_report_inputs['student_results_comment'] = student_comment
    course_report_inputs['ai_recommendations'] = generated_recommendations
    course_report_inputs['course_improvement_plan'] = merge_course_report_recommendations(
        course_report_inputs.get('course_improvement_plan') or [],
        generated_recommendations
    )
    course_report_inputs['ai_source'] = ai_insights.get('source') or 'fallback'
    return course_report_inputs, None

def docx_response(docx_bytes, filename="report.docx"):
    response = Response(
        docx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    safe_name = secure_filename(filename) or "report.docx"
    if not safe_name.lower().endswith('.docx'):
        safe_name += '.docx'
    response.headers["Content-Disposition"] = f'attachment; filename="{safe_name}"'
    return response

def course_report_docx_response(docx_bytes):
    return docx_response(docx_bytes, "course_report_filled.docx")

def build_course_report_pdf(stats, course_report_inputs=None, course_info=None, total_students=None):
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
    except Exception:
        arabic_reshaper = None
        get_display = None

    report_language = get_export_report_language() if has_request_context() else 'en'
    is_arabic_report = report_language == 'ar'
    regular_font_path, bold_font_path = get_report_pdf_font_paths()
    regular_font = 'Helvetica'
    bold_font = 'Helvetica-Bold'
    if regular_font_path:
        regular_font = 'CourseReportRegular'
        bold_font = 'CourseReportBold'
        if regular_font not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(regular_font, regular_font_path))
        if bold_font not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(bold_font, bold_font_path or regular_font_path))

    def display_text(value):
        text = clean_report_pdf_text(value)
        if is_arabic_report and contains_arabic(text) and arabic_reshaper:
            text = arabic_reshaper.reshape(text)
            if get_display:
                text = get_display(text)
        return text

    def paragraph(value, style):
        text = str(value or '')
        text = '<br/>'.join(escape(display_text(line)) for line in text.splitlines())
        return Paragraph(text or '&nbsp;', style)

    def label(key, fallback):
        translated = translate(key) if has_request_context() else key
        return translated if translated != key else fallback

    def course_label(value):
        return course_report_label_for_language(value, report_language)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=0.45 * inch,
        rightMargin=0.45 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.45 * inch
    )
    alignment = TA_RIGHT if is_arabic_report else TA_LEFT
    title_style = ParagraphStyle('CourseReportTitle', fontName=bold_font, fontSize=18, leading=22, alignment=alignment, textColor=colors.HexColor('#26365f'))
    section_style = ParagraphStyle('CourseReportSection', fontName=bold_font, fontSize=12, leading=15, alignment=alignment, textColor=colors.HexColor('#26365f'), spaceBefore=10)
    text_style = ParagraphStyle('CourseReportText', fontName=regular_font, fontSize=9, leading=12, alignment=alignment)
    header_style = ParagraphStyle('CourseReportHeader', fontName=bold_font, fontSize=8.5, leading=11, alignment=alignment, textColor=colors.white)

    def table(rows, widths):
        result = Table(rows, colWidths=widths, repeatRows=1)
        result.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#26365f')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), bold_font),
            ('FONTNAME', (0, 1), (-1, -1), regular_font),
            ('FONTSIZE', (0, 0), (-1, -1), 8.5),
            ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#d1d5db')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ]))
        return result

    course_report_inputs = course_report_inputs or {}
    course_info = course_info or {}
    report_details = course_report_inputs.get('report_details') or {}
    story = [
        paragraph(label('course_report.preview_title', 'Course Report'), title_style),
        Spacer(1, 0.16 * inch),
        paragraph(label('course_report.report_details', 'Report Details'), section_style),
    ]

    details = [
        (label('results.course', 'Course'), course_info.get('course_name') or course_info.get('raw_name') or '-'),
        (label('courses.course_code', 'Course Code'), course_info.get('course_id') or course_info.get('course_code') or '-'),
        (label('results.course_instructor', 'Course Instructor'), report_details.get('course_instructor') or '-'),
        (label('results.course_coordinator', 'Course Coordinator'), report_details.get('course_coordinator') or '-'),
        (label('results.number_of_sections', 'Number of Sections'), report_details.get('sections_count') or '-'),
        (label('results.students_started', 'Students Started'), report_details.get('students_started') or total_students or '-'),
        (label('results.students_completed', 'Students Completed'), report_details.get('students_completed') or total_students or '-'),
    ]
    detail_rows = [[paragraph(label('results.item', 'Item'), header_style), paragraph(label('results.value', 'Value'), header_style)]]
    detail_rows.extend([[paragraph(name, text_style), paragraph(value, text_style)] for name, value in details])
    story.append(table(detail_rows, [2.2 * inch, 7.3 * inch]))

    grade_distribution = course_report_inputs.get('grade_distribution') or {}
    if grade_distribution.get('total'):
        story.extend([Spacer(1, 0.12 * inch), paragraph(label('course_report.grade_distribution', 'Final Grade Distribution'), section_style)])
        grade_rows = [[paragraph(item, header_style) for item in GRADE_ORDER + [label('checkout.total', 'Total')]]]
        counts = grade_distribution.get('counts') or {}
        grade_rows.append([paragraph(counts.get(grade, 0), text_style) for grade in GRADE_ORDER] + [paragraph(grade_distribution.get('total'), text_style)])
        story.append(table(grade_rows, [0.72 * inch] * (len(GRADE_ORDER) + 1)))
        if course_report_inputs.get('student_results_comment'):
            story.extend([Spacer(1, 0.08 * inch), paragraph(course_report_inputs.get('student_results_comment'), text_style)])

    story.extend([Spacer(1, 0.12 * inch), paragraph(label('course_report.clo_summary', 'CLO Assessment Results'), section_style)])
    clo_rows = [[
        paragraph(label('results.code', 'Code'), header_style),
        paragraph(label('results.max_score', 'Max Score'), header_style),
        paragraph(label('results.target', 'Target'), header_style),
        paragraph(label('results.students_achieved', 'Students Achieved'), header_style),
        paragraph(label('results.achievement', 'Achievement'), header_style),
    ]]
    for clo, item in sorted_clo_items(stats or {}):
        target = f"{float(item.get('target_score') or 0):.2f}"
        if item.get('target_pct') is not None:
            target = f"{target} ({float(item.get('target_pct') or 0):.2f}%)"
        clo_rows.append([
            paragraph(clo_number(clo), text_style),
            paragraph(f"{float(item.get('total_possible_score') or 0):.2f}", text_style),
            paragraph(target, text_style),
            paragraph(item.get('students_achieved') or 0, text_style),
            paragraph(f"{float(item.get('achievement_percentage') or 0):.2f}%", text_style),
        ])
    story.append(table(clo_rows, [1.05 * inch, 1.35 * inch, 2.2 * inch, 1.55 * inch, 1.4 * inch]))

    improvement_items = course_report_inputs.get('course_improvement_plan') or []
    if improvement_items:
        story.extend([Spacer(1, 0.12 * inch), paragraph(label('course_report.recommendations', 'Recommendations'), section_style)])
        recommendation_rows = [[
            paragraph(label('results.recommendation', 'Recommendation'), header_style),
            paragraph(label('results.actions_needed', 'Actions Needed'), header_style),
            paragraph(label('results.needed_support', 'Needed Support'), header_style),
        ]]
        for item in improvement_items:
            recommendation_rows.append([
                paragraph(course_label(item.get('recommendation')), text_style),
                paragraph(course_label(item.get('actions_needed')), text_style),
                paragraph(course_label(item.get('support')), text_style),
            ])
        story.append(table(recommendation_rows, [3.2 * inch, 3.2 * inch, 3.1 * inch]))

    uncovered_items = course_report_inputs.get('uncovered_topic_details') or []
    if course_report_inputs.get('topics_covered') != 'yes' and uncovered_items:
        story.extend([Spacer(1, 0.12 * inch), paragraph(label('results.uncovered_topics', 'Uncovered Topics'), section_style)])
        uncovered_rows = [[
            paragraph(label('results.topic', 'Topic'), header_style),
            paragraph(label('results.uncovered_reason', 'Reason'), header_style),
            paragraph(label('results.uncovered_action', 'Action'), header_style),
        ]]
        for item in uncovered_items:
            uncovered_rows.append([
                paragraph(item.get('topic'), text_style),
                paragraph(course_label(item.get('reason')), text_style),
                paragraph(course_label(item.get('action')), text_style),
            ])
        story.append(table(uncovered_rows, [3.2 * inch, 3.2 * inch, 3.1 * inch]))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def course_report_pdf_response(pdf_bytes):
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'attachment; filename="course_report.pdf"'
    return response

def save_course_report_snapshot(stats, course_report_inputs, course_info, total_students=None, source_report_ids=None):
    user = current_user()
    if not user:
        return {'allowed': False, 'saved': False, 'reason': 'anonymous', 'id': None}

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    course_name = course_info.get('raw_name') or course_info.get('course_name') or ''
    payload = {
        'report_type': 'course_report',
        'stats': json_safe(stats or {}),
        'course_report_inputs': json_safe(course_report_inputs or {}),
        'total_students': total_students or 0,
        'course_info': json_safe(course_info or {}),
        'source_report_ids': source_report_ids or [],
        'branding': get_display_branding_payload(),
        'created_at': created_at,
    }
    payload_json = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    report_hash = hashlib.sha256(
        json.dumps({'payload': payload, 'nonce': str(uuid.uuid4())}, ensure_ascii=False, sort_keys=True).encode('utf-8')
    ).hexdigest()
    with get_db() as conn:
        saved_count = conn.execute(
            "SELECT COUNT(*) FROM saved_reports WHERE user_id = ?",
            (user['id'],)
        ).fetchone()[0]
        entitlement = report_creation_entitlement(user, saved_count)
        if not entitlement:
            return {'allowed': False, 'saved': False, 'reason': 'billing_required', 'id': None}

        title = unique_saved_report_title(
            conn,
            user['id'],
            course_name,
            f"Course Report - {course_name}".strip(' -') or 'Course Report'
        )
        row = conn.execute(
            """
            INSERT INTO saved_reports
                (user_id, title, course_name, payload_json, report_hash, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            RETURNING id
            """,
            (
                user['id'],
                title,
                course_name,
                payload_json,
                report_hash,
                created_at
            )
        ).fetchone()
        if entitlement == 'credit':
            conn.execute(
                "UPDATE users SET report_credits = CASE WHEN report_credits > 0 THEN report_credits - 1 ELSE 0 END WHERE id = ?",
                (user['id'],)
            )
        return {'allowed': True, 'saved': True, 'reason': entitlement, 'id': row['id'] if row else None}

def course_report_draft_path(draft_id):
    safe_id = re.sub(r'[^A-Za-z0-9_-]', '', str(draft_id or ''))
    if not safe_id:
        raise ValueError("Invalid course report draft.")
    return get_upload_path(f"course_report_draft_{safe_id}.json")

def save_course_report_draft(payload):
    draft_id = str(uuid.uuid4())
    with open(course_report_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False)
    return draft_id

def load_course_report_draft(draft_id):
    path = course_report_draft_path(draft_id)
    if not os.path.exists(path):
        raise FileNotFoundError("Course report draft was not found. Please try again.")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def render_course_report_preview(report_id, payload, is_draft=False):
    payload = payload or {}
    course_report_inputs = payload.get('course_report_inputs') or {}
    grade_distribution = course_report_inputs.get('grade_distribution') or {}
    return render_template(
        'course_report_preview.html',
        report_id=report_id,
        stats=payload.get('stats') or {},
        stats_items=sorted_clo_items(payload.get('stats') or {}),
        course_info=payload.get('course_info') or {},
        course_report_inputs=course_report_inputs,
        report_details=course_report_inputs.get('report_details') or {},
        grade_distribution=grade_distribution,
        grade_order=GRADE_ORDER,
        total_students=payload.get('total_students') or 0,
        clo_number=clo_number,
        is_draft=is_draft
    )

def build_results_pdf_legacy(stats, total_students, course_info, student_achievement_rows=None, branding=None):
    branding = apply_university_identity_colors(branding or get_report_branding())
    labels = pdf_report_labels(get_export_report_language())
    logo_path = resolve_branding_logo_path(branding, report_ready=True, legacy_pdf=True)
    primary_color = branding.get('primary_color') or '#26365f'
    secondary_color = branding.get('secondary_color') or primary_color
    department = (branding.get('department') or '').strip()
    logo_bytes = b''
    logo_width = 0
    logo_height = 0
    if os.path.exists(logo_path):
        with open(logo_path, 'rb') as f:
            logo_bytes = f.read()
        logo_width, logo_height = get_jpeg_size(logo_bytes)

    content_parts = []
    if logo_bytes:
        content_parts.append("q")
        content_parts.append("115 0 0 105 45 682 cm")
        content_parts.append("/Im1 Do")
        content_parts.append("Q")

    title_x = 180 if logo_bytes else 50
    report_date = datetime.now().strftime("%Y-%m-%d")
    content_parts.append(pdf_rgb_command(primary_color, "RG"))
    content_parts.append(pdf_rgb_command(primary_color, "rg"))
    pdf_text(content_parts, title_x, 750, "CLO Attainment Report", 17, "F2")
    content_parts.append(pdf_rgb_command(secondary_color, "RG"))
    pdf_line(content_parts, title_x, 735, 560, 735)

    content_parts.append("0 0 0 RG")
    content_parts.append("0 0 0 rg")
    pdf_text(content_parts, title_x, 710, f"Course Name: {course_info.get('course_name', '')}", 11, "F1")
    pdf_text(content_parts, title_x, 692, f"Course ID: {course_info.get('course_id', '') or 'N/A'}", 11, "F1")
    pdf_text(content_parts, title_x, 674, f"University: {branding.get('organization_name') or 'N/A'}", 11, "F1")
    if department:
        pdf_text(content_parts, title_x, 656, f"Department: {department}", 11, "F1")
        pdf_text(content_parts, title_x, 638, f"Report Date: {report_date}", 11, "F1")
    else:
        pdf_text(content_parts, title_x, 656, f"Report Date: {report_date}", 11, "F1")

    content_parts.append("0.965 0.973 0.980 rg")
    pdf_rect(content_parts, 50, 625, 510, 42, True)
    content_parts.append("0.835 0.855 0.890 RG")
    pdf_rect(content_parts, 50, 625, 510, 42, False)
    content_parts.append("0 0 0 rg")
    pdf_text(content_parts, 70, 650, "Total Students Evaluated", 10, "F1")
    pdf_text(content_parts, 70, 632, str(total_students), 16, "F2")
    pdf_text(content_parts, 245, 650, "Mapped CLOs", 10, "F1")
    pdf_text(content_parts, 245, 632, str(len(stats)), 16, "F2")

    page_contents = [content_parts]
    table_x = 30
    table_y = 590
    col_widths = [46, 198, 58, 78, 82, 58]
    headers = ["Code", "Questions", "Max", "Target", "Achieved", "Achievement"]

    def draw_table_header(parts, y_position):
        parts.append(pdf_rgb_command(primary_color, "rg"))
        pdf_rect(parts, table_x, y_position, sum(col_widths), 24, True)
        parts.append("1 1 1 rg")
        header_x = table_x + 5
        for header, width in zip(headers, col_widths):
            pdf_text(parts, header_x, y_position + 8, header, 8, "F2")
            header_x += width
        parts.append("0 0 0 rg")
        parts.append("0.835 0.855 0.890 RG")

    def new_continuation_page():
        parts = []
        parts.append(pdf_rgb_command(primary_color, "RG"))
        parts.append(pdf_rgb_command(primary_color, "rg"))
        pdf_text(parts, 40, 755, "CLO Attainment Report", 14, "F2")
        parts.append(pdf_rgb_command(secondary_color, "RG"))
        pdf_line(parts, 40, 740, 560, 740)
        parts.append("0 0 0 RG")
        parts.append("0 0 0 rg")
        pdf_text(parts, 40, 720, f"Course Name: {course_info.get('course_name', '')}", 9, "F1")
        if department:
            pdf_text(parts, 40, 706, f"Department: {department}", 9, "F1")
        pdf_text(parts, 330, 720, f"Report Date: {report_date}", 9, "F1")
        draw_table_header(parts, 680)
        page_contents.append(parts)
        return parts, 680

    draw_table_header(content_parts, table_y)
    y = table_y
    current_parts = content_parts
    for clo, data in sorted_clo_items(stats):
        question_text = format_mapped_questions_for_report(data['questions'], language=get_export_report_language())
        clo_lines = wrap_pdf_text(clo_number(clo), 48)
        question_lines = wrap_pdf_text(question_text, 42)
        line_count = max(len(clo_lines), len(question_lines), 2)
        row_height = max(36, 16 + (line_count * 9))

        if y - row_height < 55:
            current_parts, y = new_continuation_page()

        row_y = y - row_height
        pdf_rect(current_parts, table_x, row_y, sum(col_widths), row_height, False)
        x = table_x
        for width in col_widths[:-1]:
            x += width
            pdf_line(current_parts, x, row_y, x, row_y + row_height)

        text_top = row_y + row_height - 12
        draw_pdf_lines(current_parts, clo_lines, table_x + 5, text_top, 6.7, 9, "F1")
        draw_pdf_lines(current_parts, question_lines, table_x + 225, text_top, 6.7, 9, "F1")
        number_y = row_y + row_height - 21
        pdf_text(current_parts, table_x + 350, number_y, f"{data['total_possible_score']:.2f}", 7.5, "F1")
        pdf_text(current_parts, table_x + 400, number_y, f"{data['target_score']:.2f}", 7.5, "F1")
        pdf_text(current_parts, table_x + 455, number_y, str(data['students_achieved']), 7.5, "F1")
        pdf_text(current_parts, table_x + 500, number_y, f"{data['achievement_percentage']:.2f}%", 7.5, "F1")
        y = row_y

    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    if student_achievement_matrix['students']:
        student_table_x = 40

        def draw_student_table_header(parts, y_position, clo_chunk, col_widths):
            headers = ["Student ID"] + [clo_number(clo) for clo in clo_chunk]
            parts.append(pdf_rgb_command(primary_color, "rg"))
            pdf_rect(parts, student_table_x, y_position, sum(col_widths), 26, True)
            parts.append("1 1 1 rg")
            header_x = student_table_x + 5
            for header, width in zip(headers, col_widths):
                header_lines = wrap_pdf_text(header, max(8, int(width / 5)))
                draw_pdf_lines(parts, header_lines[:2], header_x, y_position + 17, 6.5, 8, "F2")
                header_x += width
            parts.append("0 0 0 rg")
            parts.append("0.835 0.855 0.890 RG")

        def new_student_page(clo_chunk, col_widths):
            parts = []
            parts.append(pdf_rgb_command(primary_color, "RG"))
            parts.append(pdf_rgb_command(primary_color, "rg"))
            pdf_text(parts, 40, 755, "Student CLO Achievement", 14, "F2")
            parts.append(pdf_rgb_command(secondary_color, "RG"))
            pdf_line(parts, 40, 740, 560, 740)
            parts.append("0 0 0 RG")
            parts.append("0 0 0 rg")
            pdf_text(parts, 40, 720, f"Course Name: {course_info.get('course_name', '')}", 9, "F1")
            if department:
                pdf_text(parts, 40, 706, f"Department: {department}", 9, "F1")
            pdf_text(parts, 330, 720, f"Report Date: {report_date}", 9, "F1")
            draw_student_table_header(parts, 680, clo_chunk, col_widths)
            page_contents.append(parts)
            return parts, 680

        clo_chunks = [
            student_achievement_matrix['clos'][index:index + 4]
            for index in range(0, len(student_achievement_matrix['clos']), 4)
        ]
        for clo_chunk in clo_chunks:
            col_widths = [95] + [int(445 / max(len(clo_chunk), 1))] * len(clo_chunk)
            current_parts, y = new_student_page(clo_chunk, col_widths)
            for student_id in student_achievement_matrix['students']:
                row_height = 34
                if y - row_height < 55:
                    current_parts, y = new_student_page(clo_chunk, col_widths)

                row_y = y - row_height
                pdf_rect(current_parts, student_table_x, row_y, sum(col_widths), row_height, False)
                x = student_table_x
                for width in col_widths[:-1]:
                    x += width
                    pdf_line(current_parts, x, row_y, x, row_y + row_height)

                pdf_text(current_parts, student_table_x + 5, row_y + 15, display_student_id(student_id), 7.5, "F1")
                cell_x = student_table_x + col_widths[0]
                for clo, width in zip(clo_chunk, col_widths[1:]):
                    cell = student_achievement_matrix['cells'].get(student_id, {}).get(clo)
                    if cell:
                        status = labels.get('student_achieved_status', labels.get('achieved_status', 'Achieved')) if cell.get('achieved') else labels.get('student_not_achieved_status', labels.get('not_achieved', 'Not Achieved'))
                        pdf_text(current_parts, cell_x + 5, row_y + 19, f"{cell.get('score', 0):.2f}", 7.5, "F1")
                        pdf_text(current_parts, cell_x + 5, row_y + 8, status, 6.8, "F1")
                    else:
                        pdf_text(current_parts, cell_x + 5, row_y + 15, "-", 7.5, "F1")
                    cell_x += width
                y = row_y

    for parts in page_contents:
        parts.append("0.5 0.5 0.5 rg")
        pdf_text(parts, 50, 35, "Generated by ETQAN", 8, "F1")

    page_count = len(page_contents)
    font_regular_id = 3 + page_count
    font_bold_id = font_regular_id + 1
    content_start_id = font_bold_id + 1
    image_id = content_start_id + page_count if logo_bytes else None
    image_resource = f" /XObject << /Im1 {image_id} 0 R >>" if logo_bytes else ""
    page_kids = " ".join(f"{3 + index} 0 R" for index in range(page_count))
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        f"<< /Type /Pages /Kids [{page_kids}] /Count {page_count} >>".encode('ascii')
    ]
    for page_index in range(page_count):
        content_id = content_start_id + page_index
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 {font_regular_id} 0 R /F2 {font_bold_id} 0 R >>{image_resource} >> /Contents {content_id} 0 R >>".encode('ascii')
        )
    objects.extend([
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"
    ])
    for parts in page_contents:
        stream = "\n".join(parts).encode('latin-1')
        objects.append(b"<< /Length " + str(len(stream)).encode('ascii') + b" >>\nstream\n" + stream + b"\nendstream")
    if logo_bytes:
        objects.append(
            f"<< /Type /XObject /Subtype /Image /Width {logo_width} /Height {logo_height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {len(logo_bytes)} >>\nstream\n".encode('ascii')
            + logo_bytes
            + b"\nendstream"
        )

    pdf = io.BytesIO()
    pdf.write(b"%PDF-1.4\n")
    offsets = [0]
    for idx, obj in enumerate(objects, start=1):
        offsets.append(pdf.tell())
        pdf.write(f"{idx} 0 obj\n".encode('ascii'))
        pdf.write(obj)
        pdf.write(b"\nendobj\n")
    xref_offset = pdf.tell()
    pdf.write(f"xref\n0 {len(objects) + 1}\n".encode('ascii'))
    pdf.write(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.write(f"{offset:010d} 00000 n \n".encode('ascii'))
    pdf.write(f"trailer << /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF".encode('ascii'))
    return pdf.getvalue()

@app.route('/api/courses')
def api_courses():
    return json.dumps(get_available_courses())

def delete_session_uploads():
    for assessment in session.get('assessment_files') or []:
        stored_name = assessment.get('stored_name')
        if not stored_name:
            continue
        filepath = get_upload_path(stored_name)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except OSError:
                pass

        paper_stored_name = assessment.get('paper_stored_name')
        if paper_stored_name:
            filepath = get_upload_path(paper_stored_name)
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except OSError:
                    pass

    branding = session.get('report_branding') or {}
    logo_stored_name = branding.get('logo_stored_name')
    if logo_stored_name:
        filepath = get_upload_path(logo_stored_name)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except OSError:
                pass

    file_id = session.get('file_id')
    file_ext = session.get('file_ext')
    if file_id and file_ext:
        filepath = get_upload_path(f"{file_id}{file_ext}")
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except OSError:
                pass

def reset_course_workflow_session():
    delete_session_uploads()
    for key in [
        'assessment_files',
        'file_id',
        'file_ext',
        'course_name',
        'selected_course_name',
        'target_percentages',
        'custom_clos',
        'report_metrics',
        'mapping_method',
        'mapping',
    ]:
        session.pop(key, None)

@app.route('/start-over')
def start_over():
    delete_session_uploads()
    user_id = session.get('user_id')
    session.clear()
    if user_id:
        session['user_id'] = user_id
    return redirect(url_for('clo_attainment'))

@app.route('/back-to-upload')
def back_to_upload():
    delete_session_uploads()
    upload_only_keys = [
        'assessment_files',
        'file_id',
        'file_ext',
        'mapping',
        'mapping_method',
        'report_metrics',
        'report_branding',
    ]
    for key in upload_only_keys:
        session.pop(key, None)
    session.modified = True
    return redirect(url_for('clo_attainment'))

@app.route('/privacy-policy')
def privacy_policy():
    return render_template('privacy_policy.html')

@app.route('/faq')
def faq():
    return render_template('faq.html')

@app.route('/contact-us', methods=['GET', 'POST'])
def contact_us():
    topic = request.args.get('topic') or ''
    if request.method == 'POST':
        attachment = request.files.get('attachment')
        stored_name = ''
        original_name = ''
        if attachment and attachment.filename:
            attachment_ext = os.path.splitext(attachment.filename)[1].lower()
            if attachment_ext not in {'.csv', '.pdf', '.png', '.jpg', '.jpeg', '.gif', '.webp'}:
                flash(translate('contact.invalid_attachment'), "error")
                return redirect(url_for('contact_us', topic=topic) if topic else url_for('contact_us'))
            contact_dir = get_upload_path('contact_attachments')
            os.makedirs(contact_dir, exist_ok=True)
            original_name = attachment.filename
            safe_name = secure_filename(attachment.filename) or f"attachment{attachment_ext}"
            stored_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex}_{safe_name}"
            attachment.save(os.path.join(contact_dir, stored_name))
        user = current_user()
        with get_db() as conn:
            conn.execute(
                """
                INSERT INTO contact_requests
                    (user_id, name, email, organization, college, department, enquiry_type, message,
                     attachment_stored_name, attachment_original_name, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user['id'] if user else None,
                    (request.form.get('name') or '').strip(),
                    (request.form.get('email') or '').strip(),
                    (request.form.get('organization') or '').strip(),
                    (request.form.get('college') or '').strip(),
                    (request.form.get('department') or '').strip(),
                    (request.form.get('enquiry_type') or '').strip(),
                    (request.form.get('message') or '').strip(),
                    stored_name,
                    original_name,
                    datetime.now().strftime("%Y-%m-%d %H:%M")
                )
            )
        if CONTACT_TO_EMAIL and is_email_configured():
            enquiry_type = (request.form.get('enquiry_type') or '').strip()
            sender_name = (request.form.get('name') or '').strip()
            sender_email = (request.form.get('email') or '').strip()
            organization = (request.form.get('organization') or '').strip()
            college = (request.form.get('college') or '').strip()
            department = (request.form.get('department') or '').strip()
            message_text = (request.form.get('message') or '').strip()
            attachment_note = original_name or 'No attachment'
            text_body = (
                "New ETQAN contact request\n\n"
                f"Type: {enquiry_type}\n"
                f"Name: {sender_name}\n"
                f"Email: {sender_email}\n"
                f"Organization: {organization}\n"
                f"College: {college}\n"
                f"Department: {department}\n"
                f"Attachment: {attachment_note}\n\n"
                f"Message:\n{message_text}\n"
            )
            html_body = f"""
            <h2>New ETQAN contact request</h2>
            <p><strong>Type:</strong> {enquiry_type}</p>
            <p><strong>Name:</strong> {sender_name}</p>
            <p><strong>Email:</strong> {sender_email}</p>
            <p><strong>Organization:</strong> {organization}</p>
            <p><strong>College:</strong> {college}</p>
            <p><strong>Department:</strong> {department}</p>
            <p><strong>Attachment:</strong> {attachment_note}</p>
            <p><strong>Message:</strong></p>
            <pre>{message_text}</pre>
            """
            send_email(CONTACT_TO_EMAIL, 'New ETQAN contact request', text_body, html_body)
        flash(translate('contact.sent'))
        return redirect(url_for('contact_us'))
    default_message = translate('contact.university_subscription_subject') if topic == 'university-subscription' else ''
    return render_template('contact_us.html', default_message=default_message, topic=topic)

@app.route('/set-language/<lang>')
def set_language(lang):
    if lang in SUPPORTED_LANGUAGES:
        session['language'] = lang
    next_url = request.args.get('next') or url_for('index')
    return redirect(next_url)

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    if str(filename or '').replace('\\', '/').startswith('contact_attachments/') and not is_admin_user():
        abort(403)
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/assets/etqan-logo.png')
def etqan_logo_asset():
    return send_from_directory(os.path.join(APP_BASE_DIR, 'public'), 'ETQAN.png')

@app.route('/admin')
def admin_dashboard():
    user = current_user()
    if not is_admin_user(user):
        abort(403)
    with get_db() as conn:
        totals = {
            'users': conn.execute("SELECT COUNT(*) AS count FROM users").fetchone()['count'],
            'courses': conn.execute("SELECT COUNT(*) AS count FROM user_courses").fetchone()['count'],
            'reports': conn.execute("SELECT COUNT(*) AS count FROM saved_reports").fetchone()['count'],
            'contact_requests': conn.execute("SELECT COUNT(*) AS count FROM contact_requests").fetchone()['count'],
        }
        latest_users = conn.execute(
            """
            SELECT id, email, university_name, college, department, billing_plan, created_at
              FROM users
             ORDER BY id DESC
             LIMIT 20
            """
        ).fetchall()
        latest_reports = conn.execute(
            """
            SELECT r.id, r.title, r.course_name, r.created_at, u.email
              FROM saved_reports r
              JOIN users u ON u.id = r.user_id
             ORDER BY r.id DESC
             LIMIT 20
            """
        ).fetchall()
        latest_courses = conn.execute(
            """
            SELECT c.id, c.display_name, c.course_code, c.college, c.department, c.program, c.updated_at, u.email
              FROM user_courses c
              JOIN users u ON u.id = c.user_id
             ORDER BY c.id DESC
             LIMIT 20
            """
        ).fetchall()
        latest_contacts = conn.execute(
            """
            SELECT id, name, email, organization, college, department, enquiry_type, message,
                   attachment_stored_name, attachment_original_name, created_at
              FROM contact_requests
             ORDER BY id DESC
             LIMIT 20
            """
        ).fetchall()
    return render_template(
        'admin_dashboard.html',
        totals=totals,
        latest_users=latest_users,
        latest_reports=latest_reports,
        latest_courses=latest_courses,
        latest_contacts=latest_contacts
    )

@app.route('/admin/contact-attachments/<path:filename>')
def admin_contact_attachment(filename):
    if not is_admin_user():
        abort(403)
    return send_from_directory(get_upload_path('contact_attachments'), os.path.basename(filename))

@app.route('/university-logos/<path:filename>')
def university_logo(filename):
    return send_from_directory(UNIVERSITY_LOGO_FOLDER, os.path.basename(filename))

@app.route('/organization-logos/<path:filename>')
def organization_logo(filename):
    return send_from_directory(ORG_LOGO_FOLDER, os.path.basename(filename))

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = (request.form.get('email') or '').strip().lower()
        password = request.form.get('password') or ''
        university_name = get_registration_university_name()
        college = (request.form.get('college') or '').strip()
        department = (request.form.get('department') or '').strip()
        if not email or not password:
            flash("Please enter an email and password.", "error")
            return redirect(request.url)
        if not university_name:
            flash("Please enter your university name.", "error")
            return redirect(request.url)
        if len(password) < 6:
            flash("Password must be at least 6 characters.", "error")
            return redirect(request.url)
        try:
            university_identity = get_university_identity(university_name)
            with get_db() as conn:
                cursor = conn.execute(
                    """
                    INSERT INTO users
                        (email, password_hash, created_at, university_name, college, department, org_primary_color, org_secondary_color, org_tertiary_color, billing_plan, subscription_started_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    RETURNING id
                    """,
                    (
                        email,
                        generate_password_hash(password),
                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                        university_name,
                        college,
                        department,
                        university_identity.get('primary_color', '#26365f'),
                        university_identity.get('secondary_color', ''),
                        university_identity.get('tertiary_color', ''),
                        'professional',
                        datetime.now().strftime("%Y-%m-%d")
                    )
                )
                session['user_id'] = cursor.fetchone()['id']
        except Exception as exc:
            if not is_unique_violation(exc):
                raise
            flash("An account already exists for this email.", "error")
            return redirect(url_for('login'))
        return redirect(url_for('index'))
    return render_template('auth.html', title='Create Account', button_label='Create Account', mode='register')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = (request.form.get('email') or '').strip().lower()
        password = request.form.get('password') or ''
        ip_address = get_client_ip()
        with get_db() as conn:
            lock_type, active_lock = evaluate_login_lock(conn, email, ip_address)
            if active_lock:
                return render_template(
                    'auth.html',
                    title='Login',
                    button_label='Login',
                    mode='login',
                    auth_message=translate('auth.login_locked_email' if lock_type == 'email' else 'auth.login_locked_ip'),
                    auth_message_category='error',
                    auth_email=email
                )
            user = conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
            if not user or not check_password_hash(user['password_hash'], password):
                record_failed_login(conn, email, ip_address)
                return render_template(
                    'auth.html',
                    title='Login',
                    button_label='Login',
                    mode='login',
                    auth_message=translate('auth.invalid_login'),
                    auth_message_category='error',
                    auth_email=email
                )
            clear_login_failures(conn, email, ip_address)
        session['user_id'] = user['id']
        # Removed flash("Logged in successfully.") to hide the message
        return redirect(url_for('index'))
    return render_template('auth.html', title='Login', button_label='Login', mode='login')

@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    reset_message = ''
    reset_message_category = ''
    if request.method == 'POST':
        email = (request.form.get('email') or '').strip().lower()
        reset_email_status = ''
        with get_db() as conn:
            user = conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
            if user:
                token = secrets.token_urlsafe(32)
                expires_at = (datetime.now() + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
                conn.execute(
                    "UPDATE users SET reset_token = ?, reset_token_expires_at = ? WHERE id = ?",
                    (token, expires_at, user['id'])
                )
                reset_url = external_url_for('reset_password', token=token)
                sent, reset_email_status = send_password_reset_email(user['email'], reset_url)
                if not sent:
                    conn.execute(
                        "UPDATE users SET reset_token = '', reset_token_expires_at = '' WHERE id = ?",
                        (user['id'],)
                    )
        if not user:
            reset_message = translate('auth.reset_sent')
        elif reset_email_status == 'missing_config':
            reset_message = translate('auth.reset_email_unconfigured')
            reset_message_category = 'error'
        elif reset_email_status:
            reset_message = translate('auth.reset_email_failed')
            reset_message_category = 'error'
        else:
            reset_message = translate('auth.reset_sent')
    return render_template(
        'forgot_password.html',
        reset_message=reset_message,
        reset_message_category=reset_message_category
    )

@app.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    with get_db() as conn:
        user = conn.execute("SELECT * FROM users WHERE reset_token = ?", (token,)).fetchone()
    if not user:
        flash("Invalid or expired reset link.", "error")
        return redirect(url_for('forgot_password'))
    try:
        expires_at = datetime.strptime(user['reset_token_expires_at'], "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        expires_at = datetime.min
    if datetime.now() > expires_at:
        flash("Invalid or expired reset link.", "error")
        return redirect(url_for('forgot_password'))

    if request.method == 'POST':
        password = request.form.get('password') or ''
        confirm_password = request.form.get('confirm_password') or ''
        if len(password) < 6:
            flash("Password must be at least 6 characters.", "error")
            return redirect(request.url)
        if password != confirm_password:
            flash("Passwords do not match.", "error")
            return redirect(request.url)
        with get_db() as conn:
            conn.execute(
                """
                UPDATE users
                   SET password_hash = ?,
                       reset_token = '',
                       reset_token_expires_at = ''
                 WHERE id = ?
                """,
                (generate_password_hash(password), user['id'])
            )
        flash("Password updated. Please login.")
        return redirect(url_for('login'))

    return render_template('reset_password.html')

@app.route('/logout')
def logout():
    session.pop('user_id', None)
    return redirect(url_for('index'))

@app.route('/account', methods=['GET', 'POST'])
def account_info():
    user = current_user()
    if not user:
        flash("Please login to manage your account.", "error")
        return redirect(url_for('login'))
    if request.method == 'POST':
        university_name = get_profile_university_name()
        college = (request.form.get('college') or '').strip()
        department = (request.form.get('department') or '').strip()
        if not university_name:
            flash("Please enter your institution name.", "error")
            return redirect(request.url)
        current_university_name = canonical_university_name(user['university_name'] or '')
        university_changed = canonical_university_name(university_name) != current_university_name
        with get_db() as conn:
            if university_changed:
                identity = get_university_identity(university_name)
                primary_color = identity.get('primary_color') if identity.get('has_identity_preset') else ''
                secondary_color = identity.get('secondary_color') if identity.get('has_identity_preset') else ''
                tertiary_color = identity.get('tertiary_color') if identity.get('has_identity_preset') else ''
                conn.execute(
                    """
                    UPDATE users
                       SET university_name = ?,
                           college = ?,
                           department = ?,
                           org_primary_color = ?,
                           org_secondary_color = ?,
                           org_tertiary_color = ?
                     WHERE id = ?
                    """,
                    (university_name, college, department, primary_color, secondary_color, tertiary_color, user['id'])
                )
            else:
                conn.execute(
                    """
                    UPDATE users
                       SET university_name = ?,
                           college = ?,
                           department = ?
                     WHERE id = ?
                    """,
                    (university_name, college, department, user['id'])
                )
        flash(translate('account.saved'))
        return redirect(url_for('account_info'))
    return render_template('account_info.html', billing_status=get_billing_status(user))

@app.route('/account/settings')
def account_settings():
    user = current_user()
    if not user:
        flash("Please login to manage your settings.", "error")
        return redirect(url_for('login'))
    return render_template('settings.html')

@app.route('/account/report-settings', methods=['GET', 'POST'])
def report_settings():
    user = current_user()
    if not user:
        flash("Please login to manage your report settings.", "error")
        return redirect(url_for('login'))
    if request.method == 'POST':
        report_language = (request.form.get('report_language') or 'en').strip()
        if report_language not in {'en', 'ar'}:
            report_language = 'en'
        with get_db() as conn:
            conn.execute(
                "UPDATE users SET report_language = ? WHERE id = ?",
                (report_language, user['id'])
            )
        flash(translate('report_settings.saved'))
        return redirect(url_for('report_settings'))
    report_language = user['report_language'] if 'report_language' in user.keys() else 'en'
    if report_language not in {'en', 'ar'}:
        report_language = 'en'
    return render_template('report_settings.html', report_language=report_language)

@app.route('/account/courses')
def my_courses():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    return render_template(
        'my_courses.html',
        courses=get_user_courses(user['id']),
        course_limit=get_user_course_limit(user),
        course_count=get_user_course_count(user['id'])
    )

@app.route('/account/courses/new', methods=['GET', 'POST'])
def new_course():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    draft_course = {}
    if request.method == 'POST':
        form_action = request.form.get('form_action') or 'save'
        course_name = (request.form.get('course_name') or '').strip()
        course_code = (request.form.get('course_code') or '').strip()
        college = (request.form.get('college') or '').strip()
        department = (request.form.get('department') or '').strip()
        program = (request.form.get('program') or '').strip()
        clos = parse_pasted_clos(request.form.get('clos'))
        topics = parse_course_topic_lines(request.form.get('topics'))
        clo_plos = parse_clo_plos_json(request.form.get('clo_plos_json'))
        extraction_metadata = safe_json_loads(request.form.get('extraction_metadata_json'), {}) or {}
        row_clos, row_clo_plos = parse_course_clo_rows(request.form)
        if row_clos:
            clos = row_clos
            clo_plos = row_clo_plos
        spec_file = request.files.get('course_spec_file')

        if spec_file and spec_file.filename:
            spec_ext = os.path.splitext(spec_file.filename)[1].lower()
            if spec_ext not in ['.pdf', '.docx']:
                flash("Course specification must be uploaded as a PDF or Word document (.docx).", "error")
                return redirect(request.url)
            spec_stored_name = f"{uuid.uuid4()}{spec_ext}"
            spec_filepath = get_upload_path(spec_stored_name)
            spec_file.save(spec_filepath)
            try:
                spec_text, extracted = extract_course_spec_document(spec_filepath, spec_file.filename)
                extraction_metadata = extracted.get('extraction_metadata') or {
                    'task': 'course_specification_extraction',
                    'source': extracted.get('extraction_method') or 'unknown',
                    'model': '',
                    'duration_seconds': None,
                    'filename': spec_file.filename,
                }
            except Exception as e:
                flash(f"Could not read course specification: {e}", "error")
                return redirect(request.url)
            if form_action == 'extract':
                course_name = extracted.get('course_name') or extracted.get('name') or ''
                course_code = extracted.get('course_code') or ''
                college = extracted.get('college') or ''
                department = extracted.get('department') or ''
                program = extracted.get('program') or ''
                clos = extracted.get('clos') or []
                topics = extracted.get('topics') or []
                clo_plos = extracted.get('clo_plos') or {}
            else:
                course_name = course_name or extracted.get('course_name') or extracted.get('name') or ''
                course_code = course_code or extracted.get('course_code') or ''
                college = college or extracted.get('college') or ''
                department = department or extracted.get('department') or ''
                program = program or extracted.get('program') or ''
                if not clos:
                    clos = extracted.get('clos') or []
                if not topics:
                    topics = extracted.get('topics') or []
                if not clo_plos:
                    clo_plos = extracted.get('clo_plos') or {}

        if form_action == 'extract':
            if not spec_file or not spec_file.filename:
                flash(translate('courses.extract_missing'), "error")
                course_name = ''
                course_code = ''
                college = ''
                department = ''
                program = ''
                clos = []
                topics = []
                clo_plos = {}
            else:
                flash(translate('courses.extracted'))
                flash_course_spec_extraction_method(extracted)
            draft_course = {
                'source': 'spec' if spec_file and spec_file.filename else '',
                'course_name': course_name,
                'course_code': course_code,
                'college': college,
                'department': department,
                'program': program,
                'clos_text': "\n".join(clos),
                'topics_text': format_course_topics_text(topics),
                'clo_plos_json': json.dumps(clo_plos, ensure_ascii=False),
                'extraction_metadata_json': json.dumps(extraction_metadata, ensure_ascii=False),
                'extraction_metadata': extraction_metadata,
                'clo_rows': build_course_clo_rows(clos, clo_plos)
            }
            return render_template(
                'course_new.html',
                draft_course=draft_course
            )

        display_name = build_course_display_name(course_name, course_code)
        if not display_name or not clos or not topics:
            flash(translate('courses.invalid'), "error")
            return redirect(request.url)

        target_percentages = {}

        with get_db() as conn:
            existing = conn.execute(
                "SELECT id FROM user_courses WHERE user_id = ? AND display_name = ?",
                (user['id'], display_name)
            ).fetchone()
            if not existing:
                limit = get_user_course_limit(user)
                saved_count = conn.execute(
                    "SELECT COUNT(*) FROM user_courses WHERE user_id = ?",
                    (user['id'],)
                ).fetchone()[0]
                if limit is not None and saved_count >= limit:
                    flash(translate('courses.limit'), "error")
                    return redirect(request.url)

            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            if existing:
                conn.execute(
                    """
                    UPDATE user_courses
                       SET course_name = ?,
                           course_code = ?,
                           college = ?,
                           department = ?,
                           program = ?,
                           clos_json = ?,
                           target_percentages_json = ?,
                           topics_json = ?,
                           clo_plos_json = ?,
                           extraction_metadata_json = ?,
                           updated_at = ?
                     WHERE id = ? AND user_id = ?
                    """,
                    (
                        course_name or display_name,
                        course_code,
                        college,
                        department,
                        program,
                        json.dumps(clos, ensure_ascii=False),
                        json.dumps(target_percentages, ensure_ascii=False),
                        json.dumps(topics, ensure_ascii=False),
                        json.dumps(clo_plos, ensure_ascii=False),
                        json.dumps(extraction_metadata, ensure_ascii=False),
                        now,
                        existing['id'],
                        user['id']
                    )
                )
            else:
                conn.execute(
                    """
                    INSERT INTO user_courses
                        (user_id, display_name, course_name, course_code, college, department, program, clos_json, target_percentages_json, topics_json, clo_plos_json, extraction_metadata_json, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        user['id'],
                        display_name,
                        course_name or display_name,
                        course_code,
                        college,
                        department,
                        program,
                        json.dumps(clos, ensure_ascii=False),
                        json.dumps(target_percentages, ensure_ascii=False),
                        json.dumps(topics, ensure_ascii=False),
                        json.dumps(clo_plos, ensure_ascii=False),
                        json.dumps(extraction_metadata, ensure_ascii=False),
                        now,
                        now
                    )
                )
        return redirect(url_for('my_courses'))

    return render_template(
        'course_new.html',
        draft_course=draft_course
    )

@app.route('/account/courses/<int:course_id>/edit', methods=['GET', 'POST'])
def edit_course(course_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    with get_db() as conn:
        existing = conn.execute(
            "SELECT * FROM user_courses WHERE id = ? AND user_id = ?",
            (course_id, user['id'])
        ).fetchone()

    if not existing:
        flash(translate('index.course_not_found'), "error")
        return redirect(url_for('my_courses'))

    draft_course = {
        'source': 'manual',
        'course_name': existing['course_name'] or '',
        'course_code': existing['course_code'] or '',
        'college': existing['college'] or '',
        'department': existing['department'] or '',
        'program': existing['program'] or '',
    }
    draft_course['extraction_metadata_json'] = row_get(existing, 'extraction_metadata_json') or '{}'
    draft_course['extraction_metadata'] = safe_json_loads(draft_course['extraction_metadata_json'], {})

    existing_clo_plos = {}
    try:
        existing_clo_plos = parse_clo_plos_json(existing['clo_plos_json'])
    except (KeyError, IndexError):
        existing_clo_plos = {}
    draft_course['clo_plos_json'] = json.dumps(existing_clo_plos, ensure_ascii=False)
    
    clos_text = ""
    clos_list = []
    try:
        if existing['clos_json']:
            clos_list = json.loads(existing['clos_json'])
            if clos_list:
                clos_text = "\n".join(clos_list)
    except:
        pass
    draft_course['clos_text'] = clos_text
    draft_course['clo_rows'] = build_course_clo_rows(clos_list, existing_clo_plos)
    
    topics_text = ""
    try:
        if existing['topics_json']:
            topics_list = json.loads(existing['topics_json'])
            if topics_list:
                topics_text = format_course_topics_text(topics_list)
    except:
        pass
    draft_course['topics_text'] = topics_text

    if request.method == 'POST':
        form_action = request.form.get('form_action') or 'save'
        course_name = (request.form.get('course_name') or '').strip()
        course_code = (request.form.get('course_code') or '').strip()
        college = (request.form.get('college') or '').strip()
        department = (request.form.get('department') or '').strip()
        program = (request.form.get('program') or '').strip()
        clos = parse_pasted_clos(request.form.get('clos'))
        topics = parse_course_topic_lines(request.form.get('topics'))
        clo_plos = parse_clo_plos_json(request.form.get('clo_plos_json'))
        extraction_metadata = safe_json_loads(request.form.get('extraction_metadata_json'), {}) or {}
        row_clos, row_clo_plos = parse_course_clo_rows(request.form)
        if row_clos:
            clos = row_clos
            clo_plos = row_clo_plos
        spec_file = request.files.get('course_spec_file')

        if spec_file and spec_file.filename:
            spec_ext = os.path.splitext(spec_file.filename)[1].lower()
            if spec_ext not in ['.pdf', '.docx']:
                flash("Course specification must be uploaded as a PDF or Word document (.docx).", "error")
                return redirect(request.url)
            spec_stored_name = f"{uuid.uuid4()}{spec_ext}"
            spec_filepath = get_upload_path(spec_stored_name)
            spec_file.save(spec_filepath)
            try:
                spec_text, extracted = extract_course_spec_document(spec_filepath, spec_file.filename)
                extraction_metadata = extracted.get('extraction_metadata') or {
                    'task': 'course_specification_extraction',
                    'source': extracted.get('extraction_method') or 'unknown',
                    'model': '',
                    'duration_seconds': None,
                    'filename': spec_file.filename,
                }
            except Exception as e:
                flash(f"Could not read course specification: {e}", "error")
                return redirect(request.url)
            if form_action == 'extract':
                course_name = extracted.get('course_name') or extracted.get('name') or ''
                course_code = extracted.get('course_code') or ''
                college = extracted.get('college') or ''
                department = extracted.get('department') or ''
                program = extracted.get('program') or ''
                clos = extracted.get('clos') or []
                topics = extracted.get('topics') or []
                clo_plos = extracted.get('clo_plos') or {}
            else:
                course_name = course_name or extracted.get('course_name') or extracted.get('name') or ''
                course_code = course_code or extracted.get('course_code') or ''
                college = college or extracted.get('college') or ''
                department = department or extracted.get('department') or ''
                program = program or extracted.get('program') or ''
                if not clos:
                    clos = extracted.get('clos') or []
                if not topics:
                    topics = extracted.get('topics') or []
                if not clo_plos:
                    clo_plos = extracted.get('clo_plos') or {}

        if form_action == 'extract':
            if not spec_file or not spec_file.filename:
                flash(translate('courses.extract_missing'), "error")
                course_name = ''
                course_code = ''
                college = ''
                department = ''
                program = ''
                clos = []
                topics = []
                clo_plos = {}
            else:
                flash(translate('courses.extracted'))
                flash_course_spec_extraction_method(extracted)
            draft_course = {
                'source': 'spec' if spec_file and spec_file.filename else '',
                'course_name': course_name,
                'course_code': course_code,
                'college': college,
                'department': department,
                'program': program,
                'clos_text': "\n".join(clos),
                'topics_text': format_course_topics_text(topics),
                'clo_plos_json': json.dumps(clo_plos, ensure_ascii=False),
                'extraction_metadata_json': json.dumps(extraction_metadata, ensure_ascii=False),
                'extraction_metadata': extraction_metadata,
                'clo_rows': build_course_clo_rows(clos, clo_plos)
            }
            return render_template('course_edit.html', draft_course=draft_course)

        display_name = build_course_display_name(course_name, course_code)
        if not display_name or not clos or not topics:
            flash(translate('courses.invalid'), "error")
            return redirect(request.url)

        with get_db() as conn:
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn.execute(
                """
                UPDATE user_courses
                   SET display_name = ?,
                       course_name = ?,
                       course_code = ?,
                       college = ?,
                       department = ?,
                       program = ?,
                       clos_json = ?,
                       topics_json = ?,
                       clo_plos_json = ?,
                       extraction_metadata_json = ?,
                       updated_at = ?
                 WHERE id = ? AND user_id = ?
                """,
                (
                    display_name,
                    course_name or display_name,
                    course_code,
                    college,
                    department,
                    program,
                    json.dumps(clos, ensure_ascii=False),
                    json.dumps(topics, ensure_ascii=False),
                    json.dumps(clo_plos, ensure_ascii=False),
                    json.dumps(extraction_metadata, ensure_ascii=False),
                    now,
                    course_id,
                    user['id']
                )
            )
        return redirect(url_for('my_courses'))

    return render_template(
        'course_edit.html',
        draft_course=draft_course
    )

@app.route('/account/courses/<int:course_id>/delete', methods=['POST'])
def delete_saved_course(course_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        conn.execute(
            "DELETE FROM user_courses WHERE id = ? AND user_id = ?",
            (course_id, user['id'])
        )
    flash(translate('courses.deleted'))
    return redirect(url_for('my_courses'))

@app.route('/account/programs/new', methods=['GET', 'POST'])
def new_program():
    user = current_user()
    if not user:
        flash(translate('programs.login_required'), "error")
        return redirect(url_for('login'))
    draft = {
        'program_name': '',
        'program_code': '',
        'college': '',
        'department': '',
        'plos': ''
    }
    if request.method == 'POST':
        program_name = (request.form.get('program_name') or '').strip()
        program_code = (request.form.get('program_code') or '').strip()
        college = (request.form.get('college') or '').strip()
        department = (request.form.get('department') or '').strip()
        plos_text = (request.form.get('plos') or '').strip()
        plos = [line.strip() for line in plos_text.splitlines() if line.strip()]
        draft = {
            'program_name': program_name,
            'program_code': program_code,
            'college': college,
            'department': department,
            'plos': plos_text
        }
        display_name = build_program_display_name(program_name, program_code)
        if not display_name or not plos:
            flash(translate('programs.invalid'), "error")
            return render_template('program_new.html', draft=draft)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        with get_db() as conn:
            conn.execute(
                """
                INSERT INTO user_programs
                    (user_id, display_name, program_name, program_code, college, department, plos_json, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (user_id, display_name)
                DO UPDATE SET
                    program_name = EXCLUDED.program_name,
                    program_code = EXCLUDED.program_code,
                    college = EXCLUDED.college,
                    department = EXCLUDED.department,
                    plos_json = EXCLUDED.plos_json,
                    updated_at = EXCLUDED.updated_at
                """,
                (
                    user['id'],
                    display_name,
                    program_name or display_name,
                    program_code,
                    college,
                    department,
                    json.dumps(plos, ensure_ascii=False),
                    now,
                    now
                )
            )
        return redirect(url_for('index'))
    return render_template('program_new.html', draft=draft)

@app.route('/account/organization', methods=['GET', 'POST'])
def organization_settings():
    user = current_user()
    if not user:
        flash("Please login to manage your organization identity.", "error")
        return redirect(url_for('login'))
    if request.method == 'POST':
        try:
            update_report_branding_from_request()
        except ValueError as e:
            flash(str(e), "error")
            return redirect(request.url)
        flash(translate('org.saved'))
        return redirect(url_for('organization_settings'))
    return render_template(
        'organization_settings.html',
        active_branding=get_report_branding(),
        university_choices=UNIVERSITY_CHOICES,
        university_color_presets=UNIVERSITY_COLOR_PRESETS,
        university_identity_presets=get_university_identity_options()
    )

@app.route('/account/delete', methods=['POST'])
def delete_account():
    user = current_user()
    if not user:
        flash("Please login to manage your account.", "error")
        return redirect(url_for('login'))

    confirmation = (request.form.get('delete_confirmation') or '').strip()
    if confirmation != 'DELETE':
        flash(translate('account.delete_invalid'), "error")
        return redirect(url_for('organization_settings'))

    logo_stored_name = user['org_logo_stored_name'] or ''
    with get_db() as conn:
        conn.execute("DELETE FROM saved_reports WHERE user_id = ?", (user['id'],))
        conn.execute("DELETE FROM user_courses WHERE user_id = ?", (user['id'],))
        conn.execute("DELETE FROM users WHERE id = ?", (user['id'],))
    if logo_stored_name:
        delete_organization_logo(logo_stored_name)
    session.clear()
    flash(translate('account.deleted'))
    return redirect(url_for('index'))

@app.route('/reports')
def reports():
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, title, course_name, created_at, payload_json FROM saved_reports WHERE user_id = ? ORDER BY id DESC",
            (user['id'],)
        ).fetchall()
    reports = [
        {
            'id': row['id'],
            'title': row['title'],
            'display_title': display_saved_report_title(row),
            'course_name': row['course_name'],
            'created_at': row['created_at'],
            'report_type': (safe_json_loads(row_get(row, 'payload_json'), {}) or {}).get('report_type') or 'clo_attainment',
        }
        for row in rows
    ]
    return render_template('report_history.html', reports=reports)

@app.route('/reports/<int:report_id>')
def report_detail(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("Report not found.", "error")
        return redirect(url_for('reports'))
    if payload.get('report_type') == 'course_report':
        return render_course_report_preview(report_id, payload)
    branding = merge_with_current_branding(payload.get('branding') or {})
    return render_template(
        'report_results.html',
        stats=payload.get('stats') or {},
        stats_items=sorted_clo_items(payload.get('stats') or {}),
        total_students=payload.get('total_students') or 0,
        course_info=payload.get('course_info') or {},
        student_achievement_rows=[],
        student_achievement_matrix=payload.get('student_achievement_matrix') or {},
        display_branding=branding,
        show_exports=True,
        saved_report_view=True,
        is_saved=True,
        show_course_report_export=False,
        export_word_url=url_for('export_saved_course_report_word', report_id=report_id),
        export_pdf_url=url_for('export_saved_report_pdf', report_id=report_id),
        delete_report_url=url_for('delete_saved_report', report_id=report_id),
        clo_definitions=build_clo_definitions((payload.get('stats') or {}).keys()),
        clo_number=clo_number,
        format_question_label=format_question_label,
        format_mapped_questions_for_report=format_mapped_questions_for_report,
        student_count_warning=''
    )

@app.route('/reports/<int:report_id>/export/csv')
def export_saved_report_csv(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("Report not found.", "error")
        return redirect(url_for('reports'))

    return build_clo_csv_response(
        payload.get('stats') or {},
        payload.get('total_students') or 0,
        payload.get('course_info') or {},
        payload.get('student_achievement_matrix') or {},
        merge_with_current_branding(payload.get('branding') or {})
    )

@app.route('/reports/<int:report_id>/export/pdf')
def export_saved_report_pdf(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("Report not found.", "error")
        return redirect(url_for('reports'))

    student_rows = student_rows_from_matrix(payload.get('student_achievement_matrix') or {})
    pdf_bytes = build_results_pdf(
        payload.get('stats') or {},
        payload.get('total_students') or 0,
        payload.get('course_info') or {},
        student_rows,
        merge_with_current_branding(payload.get('branding') or {})
    )
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'attachment; filename="clo_achievement_report.pdf"'
    return response

@app.route('/reports/<int:report_id>/export/docx')
def export_saved_course_report_word(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("Report not found.", "error")
        return redirect(url_for('reports'))
    if payload.get('report_type') == 'course_report':
        try:
            docx_bytes = build_course_report_docx(
                payload.get('stats') or {},
                payload.get('course_report_inputs') or {},
                payload.get('course_info') or {},
                payload.get('total_students') or None
            )
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for('reports'))
        return course_report_docx_response(docx_bytes)

    docx_bytes = build_clo_results_docx(
        payload.get('stats') or {},
        payload.get('total_students') or 0,
        payload.get('course_info') or {},
        payload.get('student_achievement_matrix') or {}
    )
    return docx_response(docx_bytes, "clo_attainment_report.docx")

@app.route('/reports/<int:report_id>/export/course-report/pdf')
def export_saved_course_report_pdf(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("Report not found.", "error")
        return redirect(url_for('reports'))
    if payload.get('report_type') != 'course_report':
        flash("This report is not a course report.", "error")
        return redirect(url_for('reports'))
    try:
        pdf_bytes = build_course_report_pdf(
            payload.get('stats') or {},
            payload.get('course_report_inputs') or {},
            payload.get('course_info') or {},
            payload.get('total_students') or None
        )
    except Exception as e:
        app.logger.exception("Failed to export course report PDF: %s", e)
        flash(f"Failed to generate PDF: {e}", "error")
        return redirect(url_for('reports'))
    return course_report_pdf_response(pdf_bytes)

@app.route('/reports/<int:report_id>/rename', methods=['POST'])
def rename_saved_report(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))

    title = normalize_report_title(request.form.get('title'))
    if not title:
        flash(translate('history.rename_invalid'), "error")
        return redirect(url_for('reports'))

    with get_db() as conn:
        report_row = conn.execute(
            "SELECT id, course_name FROM saved_reports WHERE id = ? AND user_id = ?",
            (report_id, user['id'])
        ).fetchone()
        if not report_row:
            flash("Report not found.", "error")
            return redirect(url_for('reports'))
        course_name = row_get(report_row, 'course_name')
        if report_title_exists(conn, user['id'], course_name, title, exclude_report_id=report_id):
            flash(translate('history.rename_duplicate'), "error")
            return redirect(url_for('reports'))
        result = conn.execute(
            "UPDATE saved_reports SET title = ? WHERE id = ? AND user_id = ?",
            (title, report_id, user['id'])
        )
    if result.rowcount:
        flash(translate('history.renamed'))
    else:
        flash("Report not found.", "error")
    return redirect(url_for('reports'))

@app.route('/reports/<int:report_id>/delete', methods=['POST'])
def delete_saved_report(report_id):
    user = current_user()
    if not user:
        flash("Please login to view saved reports.", "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        result = conn.execute(
            "DELETE FROM saved_reports WHERE id = ? AND user_id = ?",
            (report_id, user['id'])
        )
    if result.rowcount:
        pass # Removed flash(translate('history.deleted')) to not show message on delete
    else:
        flash("Report not found.", "error")
    return redirect(url_for('reports'))

@app.route('/billing')
def billing():
    user = current_user()
    if not user:
        flash("Please login to view billing options.", "error")
        return redirect(url_for('login'))

    return render_template(
        'billing.html',
        billing_status=get_billing_status(user),
        payg_price=PAY_AS_YOU_GO_PRICE_SAR,
        academic_price=ACADEMIC_SUBSCRIPTION_PRICE_SAR,
        professional_price=PROFESSIONAL_SUBSCRIPTION_PRICE_SAR
    )

@app.route('/checkout', methods=['GET', 'POST'])
def checkout():
    user = current_user()
    if not user:
        flash("Please login to proceed with checkout.", "error")
        return redirect(url_for('login'))

    plan = request.args.get('plan') if request.method == 'GET' else request.form.get('plan')
    
    if plan not in {'payg', 'academic', 'professional'}:
        flash("Invalid plan selected.", "error")
        return redirect(url_for('billing'))

    if request.method == 'POST':
        with get_db() as conn:
            if plan == 'payg':
                quantity = int(request.form.get('quantity', 1))
                if quantity < 1:
                    quantity = 1
                conn.execute(
                    "UPDATE users SET report_credits = report_credits + ? WHERE id = ?",
                    (quantity, user['id'])
                )
                flash(f"{quantity} report credit(s) added to your account." if get_language() == 'en' else f"تمت إضافة {quantity} رصيد تقرير إلى حسابك.")
            elif plan in {'academic', 'professional'}:
                conn.execute(
                    """
                    UPDATE users
                       SET billing_plan = ?,
                           subscription_started_at = ?
                     WHERE id = ?
                    """,
                    (plan, datetime.now().strftime("%Y-%m-%d"), user['id'])
                )
                flash(translate(f'billing.{plan}_added'))

        if session.get('mapping'):
            return redirect(url_for('results'))
        return redirect(url_for('clo_attainment'))

    base_price = 0
    if plan == 'payg':
        base_price = PAY_AS_YOU_GO_PRICE_SAR
    elif plan == 'academic':
        base_price = ACADEMIC_SUBSCRIPTION_PRICE_SAR
    elif plan == 'professional':
        base_price = PROFESSIONAL_SUBSCRIPTION_PRICE_SAR

    return render_template(
        'checkout.html',
        plan=plan,
        base_price=base_price,
        payg_price=PAY_AS_YOU_GO_PRICE_SAR,
        academic_price=ACADEMIC_SUBSCRIPTION_PRICE_SAR,
        professional_price=PROFESSIONAL_SUBSCRIPTION_PRICE_SAR
    )

@app.route('/course-specification', methods=['GET', 'POST'])
def course_specification():
    extracted = None
    if request.method == 'POST':
        if request.form.get('action') == 'add':
            course_name = request.form.get('course_name', '').strip()
            course_code = request.form.get('course_code', '').strip()
            try:
                clos = json.loads(request.form.get('clos_json', '[]'))
            except json.JSONDecodeError:
                clos = []
            try:
                topics = json.loads(request.form.get('topics_json', '[]'))
            except json.JSONDecodeError:
                topics = []
            clo_plos = parse_clo_plos_json(request.form.get('clo_plos_json'))

            clos = [clo.strip() for clo in clos if isinstance(clo, str) and clo.strip()]
            topics = [topic.strip() for topic in topics if isinstance(topic, str) and topic.strip()]
            display_name = course_name
            if course_code and course_code not in display_name:
                display_name = f"{course_name} ({course_code})" if course_name else course_code

            if not display_name or not clos:
                flash("Please extract a valid course name and CLO list before adding the course.")
                return redirect(request.url)

            custom_courses = session.get('custom_courses', [])
            custom_courses = [course for course in custom_courses if course.get('name') != display_name]
            custom_courses.append({'name': display_name, 'clos': clos, 'topics': topics, 'clo_plos': clo_plos})
            session['custom_courses'] = custom_courses
            session['selected_course_name'] = display_name
            flash(f"Added course from specification: {display_name}")
            return redirect(url_for('clo_attainment'))

        if 'course_spec_file' not in request.files:
            flash("Please upload a course specification (PDF or DOCX).")
            return redirect(request.url)

        file = request.files['course_spec_file']
        if not file or file.filename == '':
            flash("Please upload a course specification (PDF or DOCX).")
            return redirect(request.url)

        file_ext = os.path.splitext(file.filename)[1].lower()
        if file_ext not in ['.pdf', '.docx']:
            flash("Course specification must be uploaded as a PDF or DOCX.")
            return redirect(request.url)

        file_id = str(uuid.uuid4())
        filepath = get_upload_path(f"{file_id}{file_ext}")
        reset_course_workflow_session()
        file.save(filepath)

        try:
            text, extracted = extract_course_spec_document(filepath, file.filename)
            if not compact_text(text):
                if pdf_ocr_available():
                    message = "This PDF does not contain readable embedded text, and OCR did not return usable text. Please try an OCR/searchable PDF version."
                else:
                    message = "This PDF does not contain readable embedded text. It appears to be scanned or image-based, and OCR tools are not installed on this server. Please use an OCR/searchable PDF version."
                flash(message, "error")
                return render_template('course_specification.html', extracted={
                    'name': '',
                    'course_name': '',
                    'course_code': '',
                    'clos': [],
                    'grouped_clos': group_clos_by_domain([])
                })
        except Exception as e:
            flash(f"Could not read course specification PDF: {e}")
            return redirect(request.url)

        if not extracted.get('name') or not extracted.get('clos'):
            flash("Could not fully extract the course name/code and CLOs. Please check the PDF text and try again.")
            return render_template('course_specification.html', extracted=extracted)

        flash("Review the extracted course information, then add it to the course list if it is correct.")
        return render_template('course_specification.html', extracted=extracted)

    return render_template('course_specification.html', extracted=extracted)

@app.route('/analyze-report', methods=['POST'])
def analyze_report():
    course_name = request.form.get('report_course_name')
    clos = get_course_clos(course_name)

    if 'report_file' not in request.files:
        flash("No course report file uploaded")
        return redirect(url_for('clo_attainment'))

    file = request.files['report_file']
    if file.filename == '':
        flash("No selected course report file")
        return redirect(url_for('clo_attainment'))

    file_ext = os.path.splitext(file.filename)[1].lower()
    allowed_exts = {'.pdf', '.csv', '.xlsx', '.xls'}
    if file_ext not in allowed_exts:
        flash("Please upload a PDF, CSV, or Excel course report.")
        return redirect(url_for('clo_attainment'))

    file_id = str(uuid.uuid4())
    filepath = get_upload_path(f"{file_id}{file_ext}")
    file.save(filepath)

    try:
        if file_ext == '.pdf':
            text = extract_pdf_text(filepath)
            metrics = infer_course_report_metrics(text)
        else:
            metrics = infer_spreadsheet_metrics(filepath, file_ext)
    except Exception as e:
        flash(f"Error reading course report file: {e}")
        return redirect(url_for('clo_attainment'))

    if not metrics['questions']:
        flash("Could not detect question labels in the file. You can still enter the question count manually below.")
    metrics = build_smart_clo_suggestions(metrics, clos)

    return render_template(
        'report_detected.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        filename=file.filename
    )

@app.route('/manual-report', methods=['POST'])
def manual_report():
    course_name = request.form.get('manual_course_name')
    clos = get_course_clos(course_name)
    total_students = request.form.get('manual_students', type=int, default=0)
    total_questions = request.form.get('manual_questions', type=int, default=0)
    questions = [f'Q{i}' for i in range(1, max(total_questions, 0) + 1)]
    metrics = {
        'questions': questions,
        'total_questions': len(questions),
        'total_students': max(total_students, 0),
        'confidence': 'Manual',
        'text_sample': ''
    }
    return render_template(
        'report_detected.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        filename='Manual entry'
    )

@app.route('/save-question-clos', methods=['POST'])
def save_question_clos():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    course_name = (request.form.get('course_name') or '').strip()
    filename = (request.form.get('filename') or '').strip()
    question_ids = request.form.getlist('question_ids')
    if not question_ids:
        for key in request.form.keys():
            if key.startswith('question_clo_'):
                question_ids.append(key.replace('question_clo_', ''))

    cleaned_questions = []
    seen_questions = set()
    for question in question_ids:
        question = str(question or '').strip()
        if not question or question in seen_questions:
            continue
        seen_questions.add(question)
        clos = [clo for clo in request.form.getlist(f'question_clo_{question}') if clo and clo != 'IGNORE']
        cleaned_questions.append({
            'question': question,
            'text': (request.form.get(f'question_text_{question}') or '').strip(),
            'type': (request.form.get(f'question_type_{question}') or '').strip(),
            'clos': clos,
            'mapping_source': 'manual',
            'mapping_model': '',
            'mapping_duration_seconds': None,
            'mapping_confidence': None,
            'mapping_metadata': {'source': 'manual'},
        })

    created_at = datetime.now().strftime('%Y-%m-%d %H:%M')
    base_title = os.path.splitext(os.path.basename(filename or ''))[0].strip() or course_name or 'Question CLO Mapping'
    payload = {
        'course_name': course_name,
        'filename': filename,
        'questions': cleaned_questions,
        'question_extraction': {},
        'ai_mapping': {'source': 'manual', 'diagnostics': []},
        'created_at': created_at,
    }
    with get_db() as conn:
        existing_titles = {
            row['title']
            for row in conn.execute(
                "SELECT title FROM saved_exams WHERE user_id = ?",
                (user['id'],)
            ).fetchall()
        }
        title = base_title
        counter = 2
        while title in existing_titles:
            title = f"{base_title} ({counter})"
            counter += 1
        conn.execute(
            """
            INSERT INTO saved_exams (user_id, title, course_name, filename, payload_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (user['id'], title, course_name, filename, json.dumps(payload, ensure_ascii=False), created_at)
        )

    flash(translate('exams.saved'))
    return redirect(url_for('my_exams'))

@app.route('/account/exams')
def my_exams():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        rows = conn.execute(
            """
            SELECT id, title, course_name, filename, payload_json, created_at
              FROM saved_exams
             WHERE user_id = ?
             ORDER BY id DESC
            """,
            (user['id'],)
        ).fetchall()

    exams = []
    for row in rows:
        payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}
        exams.append({
            'id': row_get(row, 'id'),
            'title': row_get(row, 'title'),
            'course_name': row_get(row, 'course_name'),
            'filename': row_get(row, 'filename'),
            'question_count': len(payload.get('questions') or []),
            'created_at': row_get(row, 'created_at'),
        })
    return render_template('exam_history.html', exams=exams)

@app.route('/account/exams/<int:exam_id>')
def exam_view(exam_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, title, course_name, filename, payload_json, created_at FROM saved_exams WHERE id = ? AND user_id = ?",
            (exam_id, user['id'])
        ).fetchone()
    if not row:
        flash("Exam not found.", "error")
        return redirect(url_for('my_exams'))
    
    exam = {
        'id': row_get(row, 'id'),
        'title': row_get(row, 'title'),
        'course_name': row_get(row, 'course_name'),
        'filename': row_get(row, 'filename'),
        'created_at': row_get(row, 'created_at'),
    }
    payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}
    return render_template('exam_view.html', exam=exam, payload=payload)

@app.route('/account/exams/<int:exam_id>/delete', methods=['POST'])
def exam_delete(exam_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        conn.execute(
            "DELETE FROM saved_exams WHERE id = ? AND user_id = ?",
            (exam_id, user['id'])
        )
    flash("Exam deleted successfully.")
    return redirect(url_for('my_exams'))

def build_exam_mapping_pdf_reportlab(payload, title='', course_name='', filename='', user=None):
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import io

    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
    except Exception:
        arabic_reshaper = None
        get_display = None

    regular_font_path, bold_font_path = get_report_pdf_font_paths()
    if not regular_font_path:
        raise RuntimeError("No Unicode PDF font found.")

    regular_font = 'CLOReportRegular'
    bold_font = 'CLOReportBold'
    if regular_font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(regular_font, regular_font_path))
    if bold_font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(bold_font, bold_font_path or regular_font_path))

    language = get_export_report_language() if has_request_context() else 'en'
    is_arabic = language == 'ar'

    def display_text(value, reorder_arabic=True):
        text = clean_report_pdf_text(value)
        if contains_arabic(text) and arabic_reshaper:
            reshaped = arabic_reshaper.reshape(text)
            if reorder_arabic and get_display:
                return get_display(reshaped)
            return reshaped
        return text

    def paragraph(value, style, reorder_arabic=True):
        lines = str(value or '').split('\n')
        text = '<br/>'.join(
            escape(display_text(line, reorder_arabic=reorder_arabic))
            for line in lines
        )
        return Paragraph(text or '&nbsp;', style)

    buffer = io.BytesIO()
    doc = BaseDocTemplate(
        buffer, pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch
    )
    frame = Frame(0.5*inch, 0.5*inch, letter[0]-1*inch, letter[1]-1*inch)
    doc.addPageTemplates([PageTemplate(id='First', frames=frame)])

    branding = apply_university_identity_colors(get_report_branding())
    labels = pdf_report_labels(language)
    organization_display_name = localized_university_name(branding.get('organization_name'), language) or labels['na']

    def hex_color(value, fallback='#26365f'):
        try:
            return colors.HexColor(value or fallback)
        except Exception:
            return colors.HexColor(fallback)

    primary_color = hex_color(branding.get('primary_color'))
    accent_color = hex_color(branding.get('secondary_color') or branding.get('primary_color'))
    body_text_color = colors.black
    logo_path = resolve_branding_logo_path(branding, report_ready=True)

    std_styles = get_standard_paragraph_styles(branding.get('primary_color'), is_arabic, regular_font, bold_font)
    title_style = std_styles['title']
    meta_style = std_styles['meta']
    section_style = std_styles['section']
    table_header = std_styles['table_header']
    table_text = std_styles['table_text']

    elements = []
    
    # Organization Header
    from reportlab.platypus import Image, Spacer
    import os
    heading_cells = []
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path)
            logo._restrictSize(1.2 * inch, 0.8 * inch)
            logo.hAlign = 'LEFT' if is_arabic else 'RIGHT'
            heading_cells.append(logo)
        except Exception:
            heading_cells.append('')
    
    report_title = 'تقرير موائمة التقييم' if is_arabic else 'Assessment Alignment Report'
    heading_text = [
        paragraph(report_title, title_style),
        paragraph(f"{labels['university']}: {organization_display_name}", meta_style),
    ]
    if branding.get('department'):
        heading_text.append(paragraph(f"{labels['department']}: {branding.get('department')}", meta_style))
    heading_text.extend([
        paragraph(f"{labels['course_name']}: {course_name or '-'}", meta_style),
        paragraph(f"{'المصدر' if is_arabic else 'Source'}: {filename or '-'}", meta_style),
    ])
    
    if heading_cells:
        if is_arabic:
            header_table = Table([[heading_text, heading_cells[0]]], colWidths=[5.45 * inch, 1.35 * inch])
        else:
            heading_cells.append(heading_text)
            header_table = Table([heading_cells], colWidths=[1.35 * inch, 5.45 * inch])
    else:
        header_table = Table([[heading_text]], colWidths=[6.8 * inch])
        
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LINEBELOW', (0, 0), (-1, -1), 1, accent_color),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 10))

    # Add CLO definitions table
    clos = get_course_clos(course_name)
    if clos:
        clo_definitions = build_clo_definitions(clos)
        elements.append(paragraph(labels['clo_definitions'], section_style))
        definition_rows = [[
            paragraph(labels['domain'], table_header),
            paragraph(labels['clo'], table_header),
            paragraph(labels['wording'], table_header)
        ]]
        for item in clo_definitions:
            definition_rows.append([
                paragraph(localized_clo_domain(item['domain'], language), table_text),
                paragraph(item['number'], table_text),
                paragraph(item['wording'], table_text),
            ])
        
        if is_arabic:
            for row in definition_rows:
                row.reverse()
            def_widths = [5.0 * inch, 0.75 * inch, 1.15 * inch]
        else:
            def_widths = [1.15 * inch, 0.75 * inch, 5.0 * inch]
            
        definition_table = Table(definition_rows, colWidths=def_widths, repeatRows=1)
        definition_table.setStyle(get_standard_table_style(branding.get('primary_color'), is_arabic, len(definition_rows)))
        elements.append(definition_table)
        elements.append(Spacer(1, 12))
        elements.append(paragraph('تفاصيل الموائمة' if is_arabic else 'Alignment Details', section_style))
        
    type_label = 'نوع السؤال' if is_arabic else 'Question Type'
    question_label = 'السؤال' if is_arabic else 'Question'
    text_label = 'نص السؤال' if is_arabic else 'Question Text'
    clo_label = 'ناتج التعلم' if is_arabic else 'Mapped CLOs'

    data = [
        [paragraph(question_label, table_header), paragraph(type_label, table_header), paragraph(text_label, table_header), paragraph(clo_label, table_header)]
    ]
    
    for index, item in enumerate(payload.get('questions') or [], start=1):
        clos = '\n'.join(clo_number(clo) or str(clo or '') for clo in item.get('clos') or [])
        q_type = item.get('question_type') or item.get('type') or '-'
        q_text = item.get('question_text') or item.get('text') or ''
        data.append([
            paragraph(f"{question_label} {index}", table_text),
            paragraph(q_type, table_text),
            paragraph(q_text, table_text),
            paragraph(clos, table_text)
        ])

    if is_arabic:
        for row in data:
            row.reverse()
        colWidths = [1.5*inch, 3.5*inch, 1*inch, 1*inch]
    else:
        colWidths = [1*inch, 1*inch, 3.5*inch, 1.5*inch]

    table = Table(data, colWidths=colWidths)
    table.setStyle(get_standard_table_style(branding.get('primary_color'), is_arabic, len(data)))
    elements.append(table)
    
    # --- Matrix Generation ---
    from reportlab.platypus import Spacer
    elements.append(Spacer(1, 0.4 * inch))
    
    matrix_title = 'مصفوفة الموائمة' if is_arabic else 'Alignment Matrix'
    elements.append(paragraph(matrix_title, title_style))
    
    matrix_data = compute_exam_alignment_matrix(payload)
    unique_clos = matrix_data['unique_clos']
    
    if unique_clos:
        m_header = [paragraph(question_label, table_header)]
        for clo in unique_clos:
            m_header.append(paragraph(clo, table_header))
        m_table_data = [m_header]
        
        for row_info in matrix_data['rows']:
            m_row = [paragraph(f"{question_label} {row_info['index']}", table_text)]
            for clo in unique_clos:
                if row_info['clos'].get(clo):
                    m_row.append(paragraph("✓", table_text))
                else:
                    m_row.append(paragraph("", table_text))
            m_table_data.append(m_row)
            
        count_label = 'الإجمالي' if is_arabic else 'Total Count'
        count_row = [paragraph(count_label, table_header)]
        for clo in unique_clos:
            count_row.append(paragraph(str(matrix_data['totals'].get(clo, 0)), table_header))
        m_table_data.append(count_row)
        
        perc_label = 'النسبة من الإجمالي' if is_arabic else 'Percentage'
        perc_row = [paragraph(perc_label, table_header)]
        for clo in unique_clos:
            perc = matrix_data['percentages'].get(clo, 0)
            perc_row.append(paragraph(f"{perc}%", table_header))
        m_table_data.append(perc_row)
        
        if is_arabic:
            for row in m_table_data:
                row.reverse()
                
        # Calculate matrix column widths dynamically
        num_cols = len(unique_clos) + 1
        available_width = letter[0] - 1 * inch
        col_width = available_width / num_cols
        
        m_table = Table(m_table_data, colWidths=[col_width] * num_cols)
        m_style = get_standard_table_style(branding.get('primary_color'), is_arabic, len(m_table_data))
        m_style.add('BACKGROUND', (0, -2), (-1, -1), colors.HexColor('#e2e8f0'))
        m_style.add('ALIGN', (0, 0), (-1, -1), 'CENTER')
        m_table.setStyle(m_style)
        elements.append(m_table)

    doc.build(elements)
    return buffer.getvalue()

@app.route('/account/exams/<int:exam_id>/export/pdf')
def export_exam_pdf(exam_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, title, course_name, filename, payload_json, created_at FROM saved_exams WHERE id = ? AND user_id = ?",
            (exam_id, user['id'])
        ).fetchone()
    if not row:
        flash("Exam not found.", "error")
        return redirect(url_for('my_exams'))
    
    exam = {
        'id': row_get(row, 'id'),
        'title': row_get(row, 'title'),
        'course_name': row_get(row, 'course_name'),
        'filename': row_get(row, 'filename'),
        'created_at': row_get(row, 'created_at'),
    }
    payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}

    try:
        pdf_bytes = build_exam_mapping_pdf_reportlab(
            payload=payload,
            title=exam['title'],
            course_name=exam['course_name'],
            filename=exam['filename'],
            user=user
        )
    except Exception as exc:
        flash(f"Failed to generate PDF. Error: {exc}", "error")
        return redirect(url_for('exam_view', exam_id=exam_id))
    
    response = Response(pdf_bytes, mimetype="application/pdf")
    filename = secure_filename(f"{exam['title']}_mapping.pdf")
    response.headers["Content-Disposition"] = f"attachment; filename={filename}"
    return response

@app.route('/account/exams/<int:exam_id>/export/docx')
def export_exam_docx(exam_id):
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, title, course_name, filename, payload_json, created_at FROM saved_exams WHERE id = ? AND user_id = ?",
            (exam_id, user['id'])
        ).fetchone()
    if not row:
        flash("Exam not found.", "error")
        return redirect(url_for('my_exams'))

    payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}
    docx_bytes = build_exam_mapping_docx(
        payload,
        title=row_get(row, 'title'),
        course_name=row_get(row, 'course_name'),
        filename=row_get(row, 'filename')
    )
    return docx_response(docx_bytes, f"{row_get(row, 'title')}_mapping.docx")

SERVICE_PLACEHOLDERS = {
    'review-course-report': ('home.reviewer_course_report_title', 'home.reviewer_course_report_description'),
    'review-course-specification': ('home.reviewer_course_spec_title', 'home.reviewer_course_spec_description'),
    'review-program-specification': ('home.reviewer_program_spec_title', 'home.reviewer_program_spec_description'),
    'review-clo-mapping': ('home.reviewer_clo_mapping_title', 'home.reviewer_clo_mapping_description'),
    'review-evidence': ('home.reviewer_evidence_title', 'home.reviewer_evidence_description'),
}

@app.route('/services/<service_slug>')
def service_placeholder(service_slug):
    title_key, description_key = SERVICE_PLACEHOLDERS.get(
        service_slug,
        ('home.open_service', 'service.coming_soon')
    )
    return render_template(
        'service_placeholder.html',
        title=translate(title_key),
        description=translate(description_key)
    )

@app.route('/')
def index():
    user = current_user()
    saved_course_count = 0
    saved_program_count = 0
    if user:
        with get_db() as conn:
            saved_course_count = conn.execute(
                "SELECT COUNT(*) AS count FROM user_courses WHERE user_id = ?",
                (user['id'],)
            ).fetchone()['count']
            saved_program_count = conn.execute(
                "SELECT COUNT(*) AS count FROM user_programs WHERE user_id = ?",
                (user['id'],)
            ).fetchone()['count']
    return render_template(
        'service_home.html',
        saved_course_count=saved_course_count,
        saved_program_count=saved_program_count
    )

@app.route('/plo-analysis', methods=['GET', 'POST'])
def plo_analysis_service():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))
        
    with get_db() as conn:
        saved_program_count = conn.execute(
            "SELECT COUNT(*) AS count FROM user_programs WHERE user_id = ?",
            (user['id'],)
        ).fetchone()['count']
        if saved_program_count == 0:
            return redirect(url_for('new_program'))
        reports = conn.execute(
            "SELECT id, title, course_name, created_at, payload_json FROM saved_reports WHERE user_id = ? ORDER BY created_at DESC",
            (user['id'],)
        ).fetchall()
        
    if request.method == 'POST':
        step = request.form.get('step')
        
        if step == 'upload_matrix':
            matrix_file = request.files.get('matrix_file')
            if not matrix_file or not matrix_file.filename:
                flash("Please upload a program specification or PLO alignment matrix file.", "error")
                return redirect(url_for('plo_analysis_service'))

            file_ext = os.path.splitext(matrix_file.filename)[1].lower()
            if file_ext not in {'.pdf', '.docx', '.xlsx', '.xls', '.txt'}:
                flash("Please upload PDF, Word, Excel, or TXT files.", "error")
                return redirect(url_for('plo_analysis_service'))

            stored_name = f"{uuid.uuid4()}{file_ext}"
            filepath = get_upload_path(stored_name)
            matrix_file.save(filepath)
            try:
                matrix = extract_program_plo_matrix(filepath, file_ext)
            except Exception as e:
                flash(f"Could not extract the PLO matrix: {e}", "error")
                return redirect(url_for('plo_analysis_service'))

            matrix['original_name'] = matrix_file.filename
            matrix['stored_name'] = stored_name
            session['plo_matrix'] = json_safe(matrix)
            session.modified = True
            if not matrix.get('courses'):
                flash("No course-to-PLO matrix could be detected. You can continue, but the PLO mappings will need to be entered manually.", "error")
            else:
                flash(f"Detected {len(matrix.get('courses', []))} course(s) with associated PLOs.")
            
            return render_template(
                'plo_analysis.html',
                step='select_reports',
                reports=reports,
                matrix=matrix,
                format_matrix_course_plos=format_matrix_course_plos
            )
        
        if step == 'map_plos':
            selected_ids = request.form.getlist('report_ids')
            if not selected_ids:
                flash("Please select at least one course report.", "error")
                return redirect(url_for('plo_analysis_service'))
                
            selected_reports_data = []
            matrix = session.get('plo_matrix') or {}
            for r in reports:
                if str(r['id']) in selected_ids:
                    payload = json.loads(r['payload_json'])
                    clos = payload.get('stats', {}).get('clo_overall', {})
                    matrix_course = find_matrix_course_for_report(r['course_name'], matrix)
                    clo_list = []
                    for clo_text, data in clos.items():
                        clo_list.append({
                            'text': clo_text,
                            'attainment': data.get('attainment_percentage', 0)
                        })
                    selected_reports_data.append({
                        'id': r['id'],
                        'course_name': r['course_name'],
                        'clos': clo_list,
                        'matrix_course': matrix_course,
                        'matrix_plos': format_matrix_course_plos(matrix_course) if matrix_course else ''
                    })
                    
            if not selected_reports_data:
                flash("Could not load data for selected reports.", "error")
                return redirect(url_for('plo_analysis_service'))
                
            return render_template(
                'plo_analysis.html',
                step='map_plos',
                reports_data=selected_reports_data,
                selected_ids=",".join(selected_ids),
                matrix=matrix
            )
            
        elif step == 'results':
            selected_ids = request.form.get('selected_ids', '').split(',')
            plo_attainments = {} # format: { plo_name: { sum: 0, count: 0, mapped_clos: [] } }
            
            for r in reports:
                if str(r['id']) in selected_ids:
                    payload = json.loads(r['payload_json'])
                    clos = payload.get('stats', {}).get('clo_overall', {})
                    
                    for i, (clo_text, data) in enumerate(clos.items()):
                        plo_mapped = request.form.get(f'plo_mapping_{r["id"]}_{i}')
                        if plo_mapped and plo_mapped.strip():
                            plo_mapped = plo_mapped.strip()
                            attainment = data.get('attainment_percentage', 0)
                            
                            if plo_mapped not in plo_attainments:
                                plo_attainments[plo_mapped] = {'sum': 0, 'count': 0, 'mapped_clos': []}
                                
                            plo_attainments[plo_mapped]['sum'] += attainment
                            plo_attainments[plo_mapped]['count'] += 1
                            plo_attainments[plo_mapped]['mapped_clos'].append({
                                'course': r['course_name'],
                                'clo': clo_text,
                                'attainment': attainment
                            })
            
            results = []
            for plo, data in plo_attainments.items():
                results.append({
                    'plo': plo,
                    'average': round(data['sum'] / data['count'], 2),
                    'mapped_clos': data['mapped_clos']
                })
                
            results.sort(key=lambda x: x['plo'])
            
            return render_template(
                'plo_analysis.html',
                step='results',
                results=results
            )

    return render_template(
        'plo_analysis.html',
        step='upload_matrix',
        reports=reports,
        matrix=session.get('plo_matrix') or {},
        format_matrix_course_plos=format_matrix_course_plos
    )

@app.route('/assessment-balance-check', methods=['GET', 'POST'])
def assessment_balance_check_service():
    courses = get_available_courses()
    
    if request.method == 'POST':
        course_name = (request.form.get('course_name') or '').strip()
        assessment_name = (request.form.get('assessment_name') or '').strip()
        
        if not course_name:
            flash(translate('index.error_course'), "error")
            return redirect(request.url)
            
        if not assessment_name:
            flash("Please enter an assessment name.", "error")
            return redirect(request.url)
            
        assessment_file = request.files.get('assessment_file')
        if not assessment_file or not assessment_file.filename:
            flash("Please upload an assessment file.", "error")
            return redirect(request.url)
            
        file_ext = os.path.splitext(assessment_file.filename)[1].lower()
        if file_ext not in {'.csv', '.xlsx', '.xls', '.pdf', '.docx', '.txt'}:
            flash("Invalid file format. Please upload PDF, DOCX, TXT, CSV, or Excel.", "error")
            return redirect(request.url)
            
        file_id = str(uuid.uuid4())
        stored_name = f"{file_id}{file_ext}"
        filepath = get_upload_path(stored_name)
        assessment_file.save(filepath)
        
        try:
            if file_ext in {'.pdf', '.docx', '.txt'}:
                metrics = parse_exam_paper_metrics(filepath)
            else:
                metrics = infer_spreadsheet_metrics(filepath, file_ext)
        except Exception as e:
            flash(f"Error reading file: {e}", "error")
            return redirect(request.url)
            
        session['ab_course_name'] = course_name
        session['ab_assessment_name'] = assessment_name
        session['ab_file_id'] = file_id
        session['ab_file_ext'] = file_ext
        session['ab_metrics'] = metrics
        session.pop('ab_mapping', None)
        
        return redirect(url_for('assessment_balance_mapping'))
        
    return render_template(
        'assessment_balance_create.html',
        courses=courses
    )

@app.route('/assessment-balance-mapping', methods=['GET', 'POST'])
def assessment_balance_mapping():
    course_name = session.get('ab_course_name')
    file_id = session.get('ab_file_id')
    file_ext = session.get('ab_file_ext')
    metrics = session.get('ab_metrics') or {}
    
    if not course_name or not file_id:
        return redirect(url_for('assessment_balance_check_service'))
        
    filepath = get_upload_path(f"{file_id}{file_ext}")
    numeric_cols = []
    
    if file_ext in {'.csv', '.xlsx', '.xls'} and not metrics.get('questions'):
        try:
            if file_ext == '.csv':
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        except Exception:
            pass
            
    detected_questions = metrics.get('questions') or []
    columns = detected_questions if detected_questions else numeric_cols
    
    clos = get_course_clos(course_name)
    parsed_questions = []
    if file_ext in {'.pdf', '.docx', '.txt'}:
        try:
            from exam_parser import ExamParser
            parser = ExamParser(filepath)
            parsed_questions = parser.parse()
            
            # Auto-heal metrics questions if the file was re-parsed with new logic
            if parsed_questions:
                metrics['questions'] = [q['question_id'] for q in parsed_questions]
                if 'question_texts' not in metrics:
                    metrics['question_texts'] = {}
                for q in parsed_questions:
                    q_id = q['question_id']
                    if q_id not in metrics['question_texts']:
                        full_text = f"[{q['question_type']}] {q['question_text']}"
                        if q.get('options'):
                            opts = " | ".join([f"{k}: {v}" for k, v in q['options'].items()])
                            full_text += f"\nOptions: {opts}"
                        metrics['question_texts'][q_id] = full_text[:120]
        except Exception:
            pass

    metrics = build_smart_clo_suggestions(metrics, clos)
            
    if request.method == 'POST':
        mapping_data = {}
        # Pre-compute a dict of q_id to marks
        q_marks = {str(q['question_id']): q.get('marks', 1.0) for q in parsed_questions}
        
        for col in columns:
            col_clos = [clo for clo in request.form.getlist(f"clo_{col}") if clo and clo != "IGNORE"]
            max_score_str = request.form.get(f"max_{col}")
            if max_score_str:
                try:
                    max_score = float(max_score_str)
                except ValueError:
                    max_score = q_marks.get(str(col), 1.0)
            else:
                max_score = q_marks.get(str(col), 1.0)
                
            mapping_data[col] = {"clos": col_clos, "max_score": max_score}
            
        session['ab_mapping'] = mapping_data
        session.modified = True
        return redirect(url_for('assessment_balance_report'))
        
    detected_clo_mappings = metrics.get('detected_clo_mappings') or {}
    max_scores = metrics.get('max_scores') or {}
    existing_mapping = session.get('ab_mapping') or {}
    
    if not existing_mapping and detected_clo_mappings:
        existing_mapping = {
            column: {
                'clos': resolve_detected_clos_to_course_list(detected_clo_mappings.get(column, []), clos),
                'max_score': max_scores.get(column, 1.0)
            }
            for column in columns
            if resolve_detected_clos_to_course_list(detected_clo_mappings.get(column, []), clos)
        }
        
    return render_template(
        'assessment_balance_mapping.html',
        columns=columns,
        clos=clos,
        course_name=course_name,
        existing_mapping=existing_mapping,
        max_scores=max_scores,
        metrics=metrics,
        parsed_questions=parsed_questions
    )

@app.route('/assessment-balance-report')
def assessment_balance_report():
    mapping_data = session.get('ab_mapping') or {}
    course_name = session.get('ab_course_name')
    assessment_name = session.get('ab_assessment_name')
    
    if not mapping_data or not course_name:
        return redirect(url_for('assessment_balance_check_service'))
        
    clos = get_course_clos(course_name)
    
    clo_scores = {clo: 0.0 for clo in clos}
    clo_question_count = {clo: 0 for clo in clos}
    total_score = 0.0
    
    for col, data in mapping_data.items():
        col_clos = data.get('clos', [])
        max_score = data.get('max_score', 0.0)
        
        if col_clos:
            points_per_clo = max_score / len(col_clos)
            for clo in col_clos:
                if clo in clo_scores:
                    clo_scores[clo] += points_per_clo
                    clo_question_count[clo] += 1
            total_score += max_score
            
    clo_distribution = []
    for clo in clos:
        score = clo_scores[clo]
        pct = (score / total_score * 100) if total_score > 0 else 0.0
        clo_distribution.append({
            'clo': clo,
            'score': score,
            'pct': pct,
            'count': clo_question_count[clo]
        })
    total_questions = len(mapping_data)
        
    return render_template(
        'assessment_balance_report.html',
        course_name=course_name,
        assessment_name=assessment_name,
        total_score=total_score,
        clo_distribution=clo_distribution,
        total_questions=total_questions
    )

@app.route('/question-clo-mapping', methods=['GET', 'POST'])
def question_clo_mapping_service():
    courses = get_available_courses()
    if request.method == 'POST':
        course_name = (request.form.get('course_name') or '').strip()
        if course_name:
            clos = get_course_clos(course_name)
        else:
            flash(translate('index.error_course'), "error")
            return redirect(request.url)

        if not clos:
            flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
            return redirect(request.url)

        paper_file = request.files.get('exam_paper')
        if not paper_file or not paper_file.filename:
            flash("Please upload an exam paper.", "error")
            return redirect(request.url)
        file_ext = os.path.splitext(paper_file.filename)[1].lower()
        if file_ext not in {'.pdf', '.docx', '.txt'}:
            flash("Exam paper must be a PDF, DOCX, or TXT file.", "error")
            return redirect(request.url)

        stored_name = f"{uuid.uuid4()}{file_ext}"
        filepath = get_upload_path(stored_name)
        paper_file.save(filepath)
        try:
            metrics = parse_exam_paper_metrics(filepath)
            draft_id = save_question_mapping_draft({
                'course_name': course_name,
                'filename': paper_file.filename,
                'metrics': metrics,
            })
        except Exception as e:
            flash(f"Error reading exam paper: {e}", "error")
            return redirect(request.url)

        return redirect(url_for('question_clo_mapping_review_get', draft_id=draft_id))

    return render_template(
        'question_clo_mapping.html',
        courses=courses
    )

@app.route('/question-clo-mapping/review/<draft_id>', methods=['GET'])
def question_clo_mapping_review_get(draft_id):
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    metrics = draft.get('metrics') or {}
    return render_template(
        'question_clo_review.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        review_summary=build_question_review_summary(metrics, clos),
        filename=draft.get('filename') or '',
        draft_id=draft_id
    )

@app.route('/question-clo-mapping/map', methods=['POST'])
def question_clo_mapping_map():
    draft_id = (request.form.get('draft_id') or '').strip()
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    metrics, paper_detected_mappings = build_question_review_metrics_from_form(clos)
    if not metrics.get('questions'):
        flash(translate('question_mapping.no_questions'), "error")
        return redirect(url_for('question_clo_mapping_review_get', draft_id=draft_id))

    previous_ai_selections = (draft.get('metrics') or {}).get('ai_draft_clo_selections') or {}
    if previous_ai_selections:
        current_questions = set(metrics.get('questions') or [])
        metrics['ai_draft_clo_selections'] = {
            question: values
            for question, values in previous_ai_selections.items()
            if question in current_questions and values
        }

    if paper_detected_mappings:
        detected = dict(metrics.get('detected_clo_mappings') or {})
        detected.update(paper_detected_mappings)
        metrics['detected_clo_mappings'] = detected
    review_summary = build_question_review_summary(metrics, clos)
    draft['metrics'] = metrics
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(draft, f, ensure_ascii=False)

    if review_summary['all_mapped']:
        return redirect(url_for('question_clo_mapping_link_get', draft_id=draft_id))

    metrics = build_ai_suggestions_for_unmapped(metrics, clos, review_summary)
    draft['metrics'] = metrics
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(draft, f, ensure_ascii=False)

    return redirect(url_for('question_clo_mapping_ai_get', draft_id=draft_id))

@app.route('/question-clo-mapping/ai/<draft_id>', methods=['GET'])
def question_clo_mapping_ai_get(draft_id):
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    metrics = draft.get('metrics') or {}
    review_summary = build_question_review_summary(metrics, clos)

    return render_template(
        'question_clo_ai.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        review_summary=review_summary,
        filename=draft.get('filename') or '',
        draft_id=draft_id
    )

@app.route('/question-clo-mapping/ai-back', methods=['POST'])
def question_clo_mapping_ai_back():
    draft_id = (request.form.get('draft_id') or '').strip()
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    form_metrics = merge_ai_default_clo_selections_from_form(
        build_question_final_metrics_from_form(clos),
        clos
    )
    draft_metrics = dict(draft.get('metrics') or {})
    selections = {}
    for question in form_metrics.get('questions') or []:
        selected = form_metrics.get('detected_clo_mappings', {}).get(question, [])
        if selected:
            selections[question] = selected
    draft_metrics['ai_draft_clo_selections'] = selections
    if form_metrics.get('ai_suggested_clos'):
        draft_metrics['ai_suggested_clos'] = form_metrics.get('ai_suggested_clos')
    if form_metrics.get('ai_removed_clos'):
        draft_metrics['ai_removed_clos'] = form_metrics.get('ai_removed_clos')
    else:
        draft_metrics.pop('ai_removed_clos', None)
    draft['metrics'] = draft_metrics
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(draft, f, ensure_ascii=False)
    return redirect(url_for('question_clo_mapping_review_get', draft_id=draft_id))


@app.route('/question-clo-mapping/ai-review', methods=['POST'])
def question_clo_mapping_ai_review_post():
    draft_id = (request.form.get('draft_id') or '').strip()
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    form_metrics = merge_ai_default_clo_selections_from_form(
        build_question_final_metrics_from_form(clos),
        clos
    )
    if not form_metrics.get('questions'):
        flash(translate('question_mapping.no_questions'), "error")
        return redirect(url_for('question_clo_mapping_ai_get', draft_id=draft_id))

    draft_metrics = dict(draft.get('metrics') or {})
    form_metrics = ensure_final_review_clo_selections(form_metrics, draft_metrics, clos)
    draft_metrics['final_review_metrics'] = form_metrics
    draft_metrics['ai_draft_clo_selections'] = {
        question: selected
        for question, selected in (form_metrics.get('detected_clo_mappings') or {}).items()
        if selected
    }
    if form_metrics.get('ai_suggested_clos'):
        draft_metrics['ai_suggested_clos'] = form_metrics.get('ai_suggested_clos')
    if form_metrics.get('ai_removed_clos'):
        draft_metrics['ai_removed_clos'] = form_metrics.get('ai_removed_clos')
    else:
        draft_metrics.pop('ai_removed_clos', None)
    draft['metrics'] = draft_metrics
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(draft, f, ensure_ascii=False)
    return redirect(url_for('question_clo_mapping_final_review_get', draft_id=draft_id))


@app.route('/question-clo-mapping/final-review/<draft_id>', methods=['GET'])
def question_clo_mapping_final_review_get(draft_id):
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    draft_metrics = draft.get('metrics') or {}
    metrics = draft_metrics.get('final_review_metrics') or {}
    if not metrics.get('questions'):
        return redirect(url_for('question_clo_mapping_ai_get', draft_id=draft_id))
    metrics = ensure_final_review_clo_selections(metrics, draft_metrics, clos)
    draft_metrics['final_review_metrics'] = metrics
    draft['metrics'] = draft_metrics
    with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
        json.dump(draft, f, ensure_ascii=False)

    # Prepare matrix_data
    matrix_payload = {'questions': []}
    for q_id in metrics.get('questions') or []:
        matrix_payload['questions'].append({'clos': metrics.get('detected_clo_mappings', {}).get(q_id, [])})
    matrix_data = compute_exam_alignment_matrix(matrix_payload, clos)
    language = get_export_report_language() if has_request_context() else 'en'
    coverage_summary = generate_assessment_coverage_summary(matrix_data, clos, language)

    return render_template(
        'question_clo_final_review.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        filename=draft.get('filename') or '',
        draft_id=draft_id,
        matrix_data=matrix_data,
        coverage_summary=coverage_summary
    )



@app.route('/question-clo-mapping/save-review', methods=['POST'])
def question_clo_mapping_save_review():
    draft_id = (request.form.get('draft_id') or '').strip()
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    metrics, _paper_detected_mappings = build_question_review_metrics_from_form(clos)
    if not metrics.get('questions'):
        flash(translate('question_mapping.no_questions'), "error")
    else:
        draft['metrics'] = metrics
        with open(question_mapping_draft_path(draft_id), 'w', encoding='utf-8') as f:
            json.dump(draft, f, ensure_ascii=False)
        flash(translate('question_mapping.review_saved'))

    return redirect(url_for('question_clo_mapping_review_get', draft_id=draft_id))

@app.route('/question-clo-mapping/final', methods=['POST'])
def question_clo_mapping_final():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    draft_id = (request.form.get('draft_id') or '').strip()
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    draft_metrics = draft.get('metrics') or {}
    metrics = merge_ai_default_clo_selections_from_form(
        build_question_final_metrics_from_form(clos),
        clos
    )
    if not metrics.get('questions'):
        flash(translate('question_mapping.no_questions'), "error")
        return redirect(url_for('question_clo_mapping_service'))
    metrics = ensure_final_review_clo_selections(metrics, draft_metrics, clos)

    filename = draft.get('filename') or ''
    question_mapping_metadata = draft_metrics.get('question_clo_mapping_metadata') or {}
    question_suggestions = draft_metrics.get('smart_clo_suggestions') or {}
    extraction_metadata = {
        'source': draft_metrics.get('question_extraction_source') or '',
        'model': draft_metrics.get('question_extraction_model') or '',
        'duration_seconds': draft_metrics.get('question_extraction_duration_seconds'),
    }
    mapping_source = draft_metrics.get('question_clo_suggestion_source') or ''
    mapping_model = {
        'gemini': GEMINI_MODEL,
        'qwen': GROQ_MODEL,
        'local': 'local-semantic',
    }.get(mapping_source, '')
    mapping_durations = [
        item.get('duration_seconds')
        for item in question_mapping_metadata.values()
        if isinstance(item, dict) and item.get('duration_seconds') is not None
    ]
    ai_mapping_summary = {
        'source': mapping_source,
        'model': mapping_model,
        'duration_seconds': round(sum(float(value or 0) for value in mapping_durations), 3) if mapping_durations else None,
        'diagnostics': draft_metrics.get('question_clo_diagnostics') or [],
    }
    cleaned_questions = []
    seen_questions = set()
    for question in metrics.get('questions', []):
        question = str(question or '').strip()
        if not question or question in seen_questions:
            continue
        seen_questions.add(question)
        
        q_clos = metrics.get('detected_clo_mappings', {}).get(question, [])
        q_text = metrics.get('question_texts', {}).get(question, '')
        q_type = metrics.get('question_types', {}).get(question, '')
        q_ai_suggested = metrics.get('ai_suggested_clos', {}).get(question)
        mapping_meta = dict(question_mapping_metadata.get(question) or {})
        if q_ai_suggested and not mapping_meta:
            suggested_item = next(
                (
                    item for item in question_suggestions.get(question, [])
                    if item.get('clo') == q_ai_suggested
                ),
                {}
            )
            mapping_meta = {
                'source': draft_metrics.get('question_clo_suggestion_source') or '',
                'model': '',
                'duration_seconds': None,
                'confidence': suggested_item.get('score'),
            }
        
        cleaned_questions.append({
            'question': question,
            'text': q_text,
            'type': q_type,
            'clos': q_clos,
            'ai_suggested_clo': q_ai_suggested,
            'mapping_source': mapping_meta.get('source') or ('manual' if q_clos else ''),
            'mapping_model': mapping_meta.get('model') or '',
            'mapping_duration_seconds': mapping_meta.get('duration_seconds'),
            'mapping_confidence': mapping_meta.get('confidence'),
            'mapping_metadata': mapping_meta,
        })

    created_at = datetime.now().strftime('%Y-%m-%d %H:%M')
    base_title = os.path.splitext(os.path.basename(filename or ''))[0].strip() or course_name or 'Question CLO Mapping'
    payload = {
        'course_name': course_name,
        'filename': filename,
        'questions': cleaned_questions,
        'question_extraction': extraction_metadata,
        'ai_mapping': ai_mapping_summary,
        'created_at': created_at,
    }

    with get_db() as conn:
        existing_titles = {
            row['title']
            for row in conn.execute(
                "SELECT title FROM saved_exams WHERE user_id = ?",
                (user['id'],)
            ).fetchall()
        }
        title = base_title
        counter = 2
        while title in existing_titles:
            title = f"{base_title} ({counter})"
            counter += 1
        cursor = conn.execute(
            """
            INSERT INTO saved_exams (user_id, title, course_name, filename, payload_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            RETURNING id
            """,
            (user['id'], title, course_name, filename, json.dumps(payload, ensure_ascii=False), created_at)
        )
        new_exam_id = cursor.fetchone()[0]

    if request.form.get('final_action') == 'export_word':
        return redirect(url_for('export_exam_docx', exam_id=new_exam_id))
    if request.form.get('final_action') == 'export_pdf':
        return redirect(url_for('export_exam_pdf', exam_id=new_exam_id))
    
    flash(translate('exams.saved'), 'success')
    return redirect(url_for('question_clo_mapping_final_review_get', draft_id=draft_id))

@app.route('/question-clo-mapping/link/<draft_id>', methods=['GET'])
def question_clo_mapping_link_get(draft_id):
    try:
        draft = load_question_mapping_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('question_clo_mapping_service'))

    course_name = (draft.get('course_name') or '').strip()
    clos = get_course_clos(course_name)
    if not clos:
        flash("No CLOs were found for the selected course. Add or update the course through My Courses.", "error")
        return redirect(url_for('question_clo_mapping_service'))

    metrics = draft.get('metrics') or {}
    return render_template(
        'question_clo_link.html',
        course_name=course_name,
        clos=clos,
        metrics=metrics,
        filename=draft.get('filename') or ''
    )

@app.route('/course-report-service', methods=['GET', 'POST'])
def course_report_service():
    if request.method == 'POST':
        user = current_user()
        course_name = (request.form.get('course_name') or '').strip()
        if not course_name:
            flash(translate('index.error_course'), "error")
            return redirect(url_for('course_report_service'))
        courses = safe_available_courses()
            
        if not user:
            session['selected_course_name'] = course_name
            return render_template(
                'course_report_select.html',
                courses=courses,
                error_no_report=True,
                selected_course_name=course_name
            )

        try:
            with get_db() as conn:
                report_rows = conn.execute(
                    "SELECT id, title, course_name, created_at, payload_json FROM saved_reports WHERE user_id = ? AND course_name = ? ORDER BY id DESC",
                    (user['id'], course_name)
                ).fetchall()
        except Exception:
            app.logger.exception("Failed to load associated course reports")
            flash("Could not load associated reports for this course. Please try again.", "error")
            return render_template(
                'course_report_select.html',
                courses=courses,
                selected_course_name=course_name,
                associated_reports=[]
            )
            
        associated_reports = []
        for row in report_rows:
            payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}
            report_type = payload.get('report_type') or 'clo_attainment'
            if report_type != 'clo_attainment':
                continue
            associated_reports.append({
                'id': row_get(row, 'id'),
                'display_title': display_saved_report_title(row),
                'created_at': row_get(row, 'created_at')
            })

        if not associated_reports:
            session['selected_course_name'] = course_name
            return render_template(
                'course_report_select.html',
                courses=courses,
                error_no_report=True,
                selected_course_name=course_name,
                associated_reports=[]
            )
            
        return render_template(
            'course_report_select.html',
            courses=courses,
            selected_course_name=course_name,
            associated_reports=associated_reports
        )

    return render_template(
        'course_report_select.html',
        courses=safe_available_courses(),
        associated_reports=None
    )

@app.route('/course-report-service/reports', methods=['GET', 'POST'])
def course_report_service_inputs_multi():
    user = current_user()
    if not user:
        flash("Please login to create a course report.", "error")
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        report_ids = request.form.getlist('report_ids')
        course_name = request.form.get('course_name')
        session['last_report_ids'] = report_ids
        session['last_course_name'] = course_name
        
        final_grades_file = request.files.get('final_grades_file')
        if final_grades_file and final_grades_file.filename:
            final_grades_ext = os.path.splitext(final_grades_file.filename)[1].lower()
            if final_grades_ext in {'.csv', '.xlsx', '.xls', '.pdf'}:
                grade_filepath = get_upload_path(f"{uuid.uuid4()}{final_grades_ext}")
                final_grades_file.save(grade_filepath)
                try:
                    grade_distribution = parse_final_grade_distribution(grade_filepath, final_grades_ext)
                    if grade_distribution and grade_distribution.get('total'):
                        session['temp_grade_distribution'] = grade_distribution
                    else:
                        session.pop('temp_grade_distribution', None)
                except Exception as e:
                    app.logger.warning("Could not read final grades file in step 1: %s", e)
                    flash("Warning: Could not parse the provided final grades file.", "error")
                    session.pop('temp_grade_distribution', None)
                finally:
                    try:
                        os.remove(grade_filepath)
                    except OSError:
                        pass
            else:
                flash("Final grades file must be CSV, Excel, or PDF.", "error")
                session.pop('temp_grade_distribution', None)
        else:
            session.pop('temp_grade_distribution', None)
            
        return redirect(url_for('course_report_service_inputs_multi'))
    else:
        report_ids = session.get('last_report_ids', [])
        course_name = session.get('last_course_name', '')
        
    if not report_ids:
        flash(translate('course_report.select_one_report'), "error")
        return redirect(url_for('course_report_service'))
    try:
        records = load_course_report_records(report_ids, user['id'])
        if not records:
            flash("CLO attainment report not found.", "error")
            return redirect(url_for('course_report_service'))
        return render_course_report_inputs_from_records(
            records,
            url_for('export_selected_course_report_docx'),
            course_name or '',
            grade_distribution_provided=('temp_grade_distribution' in session)
        )
    except Exception as exc:
        app.logger.exception(
            "Failed to open course report inputs for user_id=%s report_ids=%s course=%s",
            user['id'],
            report_ids,
            course_name or ''
        )
        flash(course_report_input_error_message(exc, report_ids, course_name or ''), "error")
        return redirect(url_for('course_report_service'))

@app.route('/course-report-service/report/<int:report_id>')
def course_report_service_inputs(report_id):
    user = current_user()
    if not user:
        flash("Please login to create a course report.", "error")
        return redirect(url_for('login'))
    try:
        records = load_course_report_records([report_id], user['id'])
        if not records:
            flash("CLO attainment report not found.", "error")
            return redirect(url_for('course_report_service'))
        return render_course_report_inputs_from_records(
            records,
            url_for('export_saved_course_report_docx', report_id=report_id),
            records[0].get('course_name') or '',
            grade_distribution_provided=('temp_grade_distribution' in session)
        )
    except Exception as exc:
        app.logger.exception(
            "Failed to open single course report input for user_id=%s report_id=%s",
            user['id'],
            report_id
        )
        flash(course_report_input_error_message(exc, [report_id], records[0].get('course_name') if 'records' in locals() and records else ''), "error")
        return redirect(url_for('course_report_service'))

@app.route('/course-report-service/report/<int:report_id>/export', methods=['POST'])
def export_saved_course_report_docx(report_id):
    redirect_url = url_for('course_report_service_inputs', report_id=report_id)
    if not require_export_profile():
        return redirect(redirect_url)

    user = current_user()
    row, payload = load_saved_report_payload(report_id, user['id'])
    if not row:
        flash("CLO attainment report not found.", "error")
        return redirect(url_for('course_report_service'))

    enriched_course_info = enrich_course_info_from_course(payload.get('course_info') or {}, row['course_name'])
    course_report_inputs, error_response = read_course_report_export_inputs(
        redirect_url,
        payload.get('stats') or {},
        enriched_course_info,
        payload.get('total_students') or None
    )
    if error_response:
        return error_response

    if request.form.get('course_report_action') == 'export_docx':
        try:
            docx_bytes = build_course_report_docx(
                payload.get('stats') or {},
                course_report_inputs,
                enriched_course_info,
                payload.get('total_students') or None
            )
        except ValueError as e:
            flash(str(e), "error")
            return redirect(redirect_url)
        return course_report_docx_response(docx_bytes)

    payload = {
        'stats': payload.get('stats') or {},
        'course_report_inputs': course_report_inputs,
        'course_info': enriched_course_info,
        'total_students': payload.get('total_students') or None,
        'source_report_ids': [report_id],
        'report_type': 'course_report'
    }
    draft_id = save_course_report_draft(payload)
    return redirect(url_for('course_report_preview_draft_get', draft_id=draft_id))

@app.route('/course-report-service/reports/export', methods=['POST'])
def export_selected_course_report_docx():
    redirect_url = url_for('course_report_service')
    if not require_export_profile():
        return redirect(redirect_url)

    user = current_user()
    report_payloads = load_selected_report_payloads(request.form.getlist('report_ids'), user['id'])
    if not report_payloads:
        flash(translate('course_report.select_one_report'), "error")
        return redirect(redirect_url)

    combined_stats, course_info, total_students, selected_reports = aggregate_course_report_payloads(
        report_payloads,
        request.form.get('course_name') or ''
    )
    course_report_inputs, error_response = read_course_report_export_inputs(
        redirect_url,
        combined_stats,
        course_info,
        total_students
    )
    if error_response:
        return error_response

    if request.form.get('course_report_action') == 'export_docx':
        try:
            docx_bytes = build_course_report_docx(combined_stats, course_report_inputs, course_info, total_students)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(redirect_url)
        return course_report_docx_response(docx_bytes)

    source_report_ids = []
    for value in request.form.getlist('report_ids'):
        try:
            source_report_ids.append(int(value))
        except (TypeError, ValueError):
            pass
            
    payload = {
        'stats': combined_stats,
        'course_report_inputs': course_report_inputs,
        'course_info': course_info,
        'total_students': total_students,
        'source_report_ids': source_report_ids,
        'report_type': 'course_report'
    }
    draft_id = save_course_report_draft(payload)
    return redirect(url_for('course_report_preview_draft_get', draft_id=draft_id))

@app.route('/course-report-service/preview/draft/<draft_id>', methods=['GET'])
def course_report_preview_draft_get(draft_id):
    user = current_user()
    if not user:
        return redirect(url_for('login'))
    
    try:
        draft = load_course_report_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('course_report_service'))
        
    return render_course_report_preview(draft_id, draft, is_draft=True)

@app.route('/course-report-service/preview/draft/<draft_id>/export/pdf')
def export_course_report_draft_pdf(draft_id):
    user = current_user()
    if not user:
        return redirect(url_for('login'))

    try:
        draft = load_course_report_draft(draft_id)
        pdf_bytes = build_course_report_pdf(
            draft.get('stats') or {},
            draft.get('course_report_inputs') or {},
            draft.get('course_info') or {},
            draft.get('total_students') or None
        )
    except Exception as exc:
        app.logger.exception("Failed to export draft course report PDF: %s", exc)
        flash(f"Failed to generate PDF: {exc}", "error")
        return redirect(url_for('course_report_preview_draft_get', draft_id=draft_id))

    return course_report_pdf_response(pdf_bytes)

@app.route('/course-report-service/preview/draft/<draft_id>/export/docx')
def export_course_report_draft_docx(draft_id):
    user = current_user()
    if not user:
        return redirect(url_for('login'))

    try:
        draft = load_course_report_draft(draft_id)
        docx_bytes = build_course_report_docx(
            draft.get('stats') or {},
            draft.get('course_report_inputs') or {},
            draft.get('course_info') or {},
            draft.get('total_students') or None
        )
    except Exception as exc:
        app.logger.exception("Failed to export draft course report Word: %s", exc)
        flash(f"Failed to generate Word: {exc}", "error")
        return redirect(url_for('course_report_preview_draft_get', draft_id=draft_id))

    return course_report_docx_response(docx_bytes)

@app.route('/course-report-service/preview/draft/<draft_id>/save', methods=['POST'])
def save_course_report_draft_action(draft_id):
    user = current_user()
    if not user:
        return redirect(url_for('login'))
        
    try:
        draft = load_course_report_draft(draft_id)
    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for('course_report_service'))
        
    combined_stats = draft.get('stats') or {}
    course_report_inputs = draft.get('course_report_inputs') or {}
    course_info = draft.get('course_info') or {}
    total_students = draft.get('total_students') or 0
    source_report_ids = draft.get('source_report_ids') or []
    
    if request.form.get('student_results_comment') is not None:
        course_report_inputs['student_results_comment'] = request.form.get('student_results_comment').strip()
        
    if course_report_inputs.get('course_improvement_plan'):
        for i, item in enumerate(course_report_inputs['course_improvement_plan']):
            rec_text = request.form.get(f'rec_text_{i}')
            if rec_text is not None:
                item['recommendation'] = rec_text.strip()
            rec_action = request.form.get(f'rec_action_{i}')
            if rec_action is not None:
                item['actions_needed'] = rec_action.strip()
            rec_support = request.form.get(f'rec_support_{i}')
            if rec_support is not None:
                item['support'] = rec_support.strip()
    
    save_result = save_course_report_snapshot(combined_stats, course_report_inputs, course_info, total_students, source_report_ids)
    if not save_result.get('allowed'):
        flash(translate('billing.limit_message') if translate('billing.limit_message') != 'billing.limit_message' else "Please upgrade or add report credits to save more reports.", "error")
        return redirect(url_for('course_report_preview_draft_get', draft_id=draft_id))
        
    saved_id = save_result.get('id')
    
    final_action = request.form.get('final_action')
    if final_action == 'export_word':
        return redirect(url_for('export_saved_course_report_word', report_id=saved_id))
    elif final_action == 'export_pdf':
        return redirect(url_for('export_saved_course_report_pdf', report_id=saved_id))
        
    return redirect(url_for('report_detail', report_id=saved_id))

@app.route('/clo-attainment', methods=['GET', 'POST'])
def clo_attainment():
    if request.method == 'GET' and request.args.get('fresh') == '1':
        reset_course_workflow_session()
    courses = get_available_courses()
    if request.method == 'POST':
        course_name = (request.form.get('course_name') or '').strip()

        if not course_name:
            flash(translate('index.error_course'), "error")
            return redirect(request.url)
        
        # Extract target percentages and edited CLO text
        target_percentages = {}
        custom_clos = []
        
        # Get the number of CLOs by finding indices in the form data
        indices = set()
        for key in request.form.keys():
            if key.startswith('clo_text_'):
                indices.add(key.replace('clo_text_', ''))
                
        for idx in sorted(list(indices), key=lambda x: int(x)):
            clo_text = request.form.get(f'clo_text_{idx}', '').strip()
            if clo_text:
                custom_clos.append(clo_text)
                try:
                    target_percentages[clo_text] = float(request.form.get(f'target_{idx}', 60.0))
                except ValueError:
                    target_percentages[clo_text] = 60.0 # Default

        if not target_percentages:
            # Fallback if no specific CLO targets are provided
            global_target = request.form.get('target_percentage', type=float, default=60.0)
            target_percentages = {"_global": global_target}

        assessment_files = []

        allowed_assessment_types = {'Final', 'Midterm', 'Quiz', 'Project', 'Assignment', 'Other'}
        pending_uploads = []
        uploaded_files = request.files.getlist('assessment_files')
        uploaded_types = request.form.getlist('assessment_types')

        for index, file in enumerate(uploaded_files):
            if not file or not file.filename:
                continue

            assessment_type = uploaded_types[index].strip() if index < len(uploaded_types) else 'Quiz'
            if assessment_type not in allowed_assessment_types:
                assessment_type = 'Quiz'
            pending_uploads.append({'file': file, 'type': assessment_type})

        if not pending_uploads:
            row_ids = request.form.getlist('assessment_row_ids')
            for row_id in row_ids:
                file = request.files.get(f'assessment_file_{row_id}')
                if not file or not file.filename:
                    continue

                assessment_type = request.form.get(f'assessment_type_{row_id}', 'Quiz').strip()
                if assessment_type not in allowed_assessment_types:
                    assessment_type = 'Quiz'
                pending_uploads.append({'file': file, 'type': assessment_type})

        type_totals = {}
        for upload in pending_uploads:
            type_totals[upload['type']] = type_totals.get(upload['type'], 0) + 1

        type_counts = {}
        for upload in pending_uploads:
            file = upload['file']
            base_label = upload['type']
            type_counts[base_label] = type_counts.get(base_label, 0) + 1
            if type_totals[base_label] > 1 or base_label in {'Quiz', 'Assignment'}:
                label = f"{base_label} {type_counts[base_label]}"
            else:
                label = base_label

            file_ext = os.path.splitext(file.filename)[1].lower()
            if file_ext not in {'.csv', '.xlsx', '.xls', '.pdf'}:
                flash(f"Invalid {label} file format. Please upload PDF, CSV, or Excel.")
                return redirect(request.url)

            file_id = str(uuid.uuid4())
            stored_name = f"{file_id}{file_ext}"
            filepath = get_upload_path(stored_name)
            file.save(filepath)

            try:
                metrics = infer_spreadsheet_metrics(filepath, file_ext)
            except Exception:
                metrics = {
                    'questions': [],
                    'total_questions': 0,
                    'total_students': 0,
                    'confidence': 'Low',
                    'text_sample': '',
                    'max_scores': {}
                }

            assessment_files.append({
                'label': label,
                'stored_name': stored_name,
                'ext': file_ext,
                'original_name': file.filename,
                'metrics': metrics
            })

        if not assessment_files:
            flash("Please upload at least one Quiz, Assignment, Midterm, Final, or Project file.", "error")
            return redirect(request.url)

        report_metrics = combine_assessment_metrics(assessment_files)
        session.pop('file_id', None)
        session.pop('file_ext', None)
        session['assessment_files'] = assessment_files
        session['course_name'] = course_name
        session['target_percentages'] = target_percentages
        session['custom_clos'] = custom_clos
        session['report_metrics'] = report_metrics
        session.pop('mapping', None)
        session.pop('mapping_method', None)

        return redirect(url_for('mapping_method'))
            
    selected_course_name = session.pop('selected_course_name', None) or session.get('course_name', '')
    return render_template(
        'report_index.html',
        courses=courses,
        selected_course_name=selected_course_name,
        saved_target_percentages=session.get('target_percentages', {}),
        saved_custom_clos=session.get('custom_clos', [])
    )

@app.route('/mapping-method', methods=['GET', 'POST'])
def mapping_method():
    user = current_user()
    if not user:
        flash(translate('courses.login_required'), "error")
        return redirect(url_for('login'))

    assessment_files = session.get('assessment_files') or []
    file_id = session.get('file_id')
    file_ext = session.get('file_ext')
    course_name = session.get('course_name')

    if not (assessment_files or file_id):
        return redirect(url_for('clo_attainment'))

    if request.method == 'POST':
        method = (request.form.get('mapping_method') or 'manual').strip().lower()
        if method == 'manual':
            session['mapping_method'] = 'manual'
            session.modified = True
            return redirect(url_for('mapping'))

        if method != 'ai':
            flash("Please choose a valid mapping method.", "error")
            return redirect(request.url)

        report_ids = request.form.getlist('report_ids')
        if not report_ids:
            flash("Please select at least one mapping report.", "error")
            return redirect(request.url)

        exam_metrics = {
            'detected_clo_mappings': {},
            'question_texts': {},
            'question_types': {}
        }
        
        try:
            with get_db() as conn:
                placeholders = ','.join('?' * len(report_ids))
                query = f"SELECT payload_json FROM saved_exams WHERE user_id = ? AND id IN ({placeholders})"
                rows = conn.execute(query, [user['id']] + report_ids).fetchall()
                
                for row in rows:
                    payload = safe_json_loads(row_get(row, 'payload_json'), {}) or {}
                    for q in payload.get('questions') or []:
                        q_id = q.get('question') or ''
                        if not q_id: continue
                        
                        mapped_clos = q.get('clos') or []
                        if mapped_clos:
                            exam_metrics['detected_clo_mappings'][q_id] = mapped_clos
                        elif q.get('ai_suggested_clo'):
                            exam_metrics['detected_clo_mappings'][q_id] = [q.get('ai_suggested_clo')]
                            
                        if q.get('text'):
                            exam_metrics['question_texts'][q_id] = q.get('text')
                        if q.get('type'):
                            exam_metrics['question_types'][q_id] = q.get('type')
        except Exception as exc:
            app.logger.exception("Failed to load mapping reports: %s", exc)
            flash(f"Could not load the mapping reports: {exc}", "error")
            return redirect(request.url)

        course_clos = session.get('custom_clos') or get_course_clos(course_name)
        course_clos = list(course_clos or [])
        report_metrics = dict(session.get('report_metrics') or {})
        report_metrics = apply_exam_paper_mappings(report_metrics, exam_metrics)
        if course_clos:
            report_metrics = build_gemini_question_clo_suggestions(report_metrics, course_clos)
            report_metrics = build_smart_clo_suggestions(
                report_metrics,
                course_clos,
                only_unmapped=report_metrics.get('question_clo_suggestion_source') == 'gemini'
            )

        paper_name = "Imported Mapping Reports"
        updated_assessment_files = []
        for item in assessment_files:
            item = dict(item)
            item['paper_original_name'] = paper_name
            updated_assessment_files.append(item)

        session['assessment_files'] = updated_assessment_files
        session['report_metrics'] = report_metrics
        session['mapping_method'] = 'ai'
        session.modified = True
        return redirect(url_for('mapping'))

    try:
        with get_db() as conn:
            report_rows = conn.execute(
                "SELECT id, title, created_at FROM saved_exams WHERE user_id = ? AND course_name = ? ORDER BY id DESC",
                (user['id'], course_name or '')
            ).fetchall()
    except Exception:
        app.logger.exception("Failed to load associated exam reports")
        report_rows = []

    associated_reports = []
    for row in report_rows:
        associated_reports.append({
            'id': row_get(row, 'id'),
            'display_title': row_get(row, 'title') or 'Mapping Report',
            'created_at': row_get(row, 'created_at')
        })

    return render_template(
        'mapping_method.html',
        course_name=course_name,
        associated_reports=associated_reports
    )

@app.route('/mapping', methods=['GET', 'POST'])
def mapping():
    assessment_files = session.get('assessment_files') or []
    file_id = session.get('file_id')
    file_ext = session.get('file_ext')
    course_name = session.get('course_name')
    
    if not (assessment_files or file_id):
        return redirect(url_for('clo_attainment'))

    if request.method == 'GET' and not session.get('mapping_method'):
        return redirect(url_for('mapping_method'))

    report_metrics = session.get('report_metrics') or {}

    numeric_cols = []
    fallback_student_count = 0
    if not assessment_files:
        filepath = get_upload_path(f"{file_id}{file_ext}")
        try:
            if file_ext == '.csv':
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
        except Exception as e:
            flash(f"Error reading file: {e}")
            return redirect(url_for('clo_attainment'))
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        fallback_student_count = len(df)

    detected_questions = report_metrics.get('questions') or []
    columns = detected_questions if detected_questions else numeric_cols
    mapping_groups = build_mapping_groups(columns, assessment_files)
    total_students = report_metrics.get('total_students') or fallback_student_count
    total_questions = report_metrics.get('total_questions') or len(columns)
    max_scores = report_metrics.get('max_scores') or {}
    detected_clo_mappings = report_metrics.get('detected_clo_mappings') or {}
    
    # Get CLOs for the selected course
    # Use custom edited CLOs if available, otherwise load from config
    course_clos = session.get('custom_clos')
    if not course_clos:
        course_clos = get_course_clos(course_name)
    course_clos = list(course_clos or [])
    if course_clos:
        report_metrics = build_smart_clo_suggestions(report_metrics, course_clos)
        session['report_metrics'] = report_metrics
        session.modified = True

    if request.method == 'POST':
        mapping_data = {}
        missing_questions = []
        for col in columns:
            clos = []
            for clo in request.form.getlist(f"clo_{col}"):
                if not clo or clo == "IGNORE":
                    continue
                compact = compact_clo_value(clo)
                if compact and compact not in clos:
                    clos.append(compact)
            max_score_str = request.form.get(f"max_{col}")
            
            if not clos:
                missing_questions.append(str(col))
                continue

            try:
                max_score = float(max_score_str) if max_score_str else 1.0
            except ValueError:
                max_score = 1.0
            mapping_data[col] = {"clos": clos, "max_score": max_score}

        if missing_questions:
            missing_text = format_missing_mapping_questions(missing_questions, assessment_files)
            flash(f"{translate('question_mapping.select_at_least_one')} {missing_text}", "error")
            session['mapping'] = mapping_data
            session.modified = True
            return redirect(url_for('mapping'))

        session['mapping'] = mapping_data
        session.modified = True
        return redirect(url_for('results'))

    existing_mapping = session.get('mapping', {})
    if not existing_mapping and detected_clo_mappings:
        existing_mapping = {
            column: {
                'clos': resolve_detected_clos_to_course_list(detected_clo_mappings.get(column, []), course_clos),
                'max_score': max_scores.get(column, 1.0)
            }
            for column in columns
            if resolve_detected_clos_to_course_list(detected_clo_mappings.get(column, []), course_clos)
        }

    return render_template(
        'report_mapping.html',
        columns=columns,
        mapping_groups=mapping_groups,
        clos=course_clos,
        course_name=course_name,
        total_students=total_students,
        total_questions=total_questions,
        detection_confidence=report_metrics.get('confidence', 'Low'),
        detection_note=report_metrics.get('text_sample', ''),
        max_scores=max_scores,
        existing_mapping=existing_mapping,
        student_count_warning=report_metrics.get('student_count_warning', ''),
        compact_clo_value=compact_clo_value
    )

@app.route('/results')
def results():
    stats, total_students, student_achievement_rows, error = calculate_clo_results()
    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))
    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    course_info = get_course_report_info()
    
    user = current_user()
    if user:
        with get_db() as conn:
            saved_count = conn.execute("SELECT COUNT(*) FROM saved_reports WHERE user_id = ?", (user['id'],)).fetchone()[0]
            if not report_creation_entitlement(user, saved_count):
                flash(translate('billing.limit_message'), "error")
                return redirect(url_for('billing'))
                
    is_saved = is_report_saved(stats, total_students, student_achievement_matrix, course_info)

    return render_template('report_results.html',
                           stats=stats,
                           stats_items=sorted_clo_items(stats),
                           total_students=total_students,
                           course_info=course_info,
                           student_achievement_rows=student_achievement_rows,
                           student_achievement_matrix=student_achievement_matrix,
                           clo_definitions=build_clo_definitions(stats.keys()),
                           clo_number=clo_number,
                           display_student_id=display_student_id,
                           course_topics=get_course_topics(session.get('course_name') or ''),
                           show_exports=True,
                           is_saved=is_saved,
                           format_question_label=format_question_label,
                           format_mapped_questions_for_report=format_mapped_questions_for_report,
                           student_count_warning=(session.get('report_metrics') or {}).get('student_count_warning', ''))

@app.route('/save-report', methods=['POST'])
def save_report():
    user = current_user()
    if not user:
        flash(translate('results.login_export') if translate('results.login_export') != 'results.login_export' else "Please login to save reports.", "error")
        return redirect(url_for('login'))
    stats, total_students, student_achievement_rows, error = calculate_clo_results()
    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))
    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    course_info = get_course_report_info()
    save_result = save_report_snapshot(stats, total_students, student_achievement_matrix, course_info)
    
    if save_result and not save_result.get('allowed', True):
        flash(translate('billing.limit_message'), "error")
        return redirect(url_for('billing'))
        
    flash(translate('exams.saved') if translate('exams.saved') != 'exams.saved' else "Report saved successfully.")
    return redirect(url_for('results'))

@app.route('/update-branding', methods=['POST'])
def update_branding():
    if not session.get('mapping'):
        return redirect(url_for('clo_attainment'))
    try:
        update_report_branding_from_request()
    except ValueError as e:
        flash(str(e), "error")
    else:
        flash("Report visual identity updated.")
    return redirect(url_for('results'))

@app.route('/export-results/csv')
def export_results_csv():
    if not require_export_profile():
        return redirect(url_for('results'))
    try:
        stats, total_students, student_achievement_rows, error = calculate_clo_results()
    except Exception as e:
        flash(f"Error exporting CSV: {e}")
        return redirect(url_for('clo_attainment'))

    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))

    course_info = get_course_report_info()
    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    return build_clo_csv_response(stats, total_students, course_info, student_achievement_matrix, get_report_branding())

@app.route('/export-results/pdf')
def export_results_pdf():
    if not require_export_profile():
        return redirect(url_for('results'))
    try:
        stats, total_students, student_achievement_rows, error = calculate_clo_results()
    except Exception as e:
        flash(f"Error exporting PDF: {e}")
        return redirect(url_for('clo_attainment'))

    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))

    pdf_bytes = build_results_pdf(stats, total_students, get_course_report_info(), student_achievement_rows, get_report_branding())
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'attachment; filename="clo_achievement_report.pdf"'
    return response

@app.route('/export-results/docx')
def export_results_docx():
    if not require_export_profile():
        return redirect(url_for('results'))
    try:
        stats, total_students, student_achievement_rows, error = calculate_clo_results()
    except Exception as e:
        flash(f"Error exporting Word: {e}")
        return redirect(url_for('clo_attainment'))

    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))

    student_achievement_matrix = build_student_achievement_matrix(student_achievement_rows, stats.keys())
    docx_bytes = build_clo_results_docx(stats, total_students, get_course_report_info(), student_achievement_matrix)
    return docx_response(docx_bytes, "clo_attainment_report.docx")

@app.route('/export-course-report/docx', methods=['POST'])
def export_course_report_docx():
    redirect_url = url_for('results')
    if not require_export_profile():
        return redirect(redirect_url)

    try:
        stats, total_students, student_achievement_rows, error = calculate_clo_results()
    except Exception as e:
        flash(f"Error exporting course report: {e}")
        return redirect(url_for('clo_attainment'))

    if error:
        flash(error)
        return redirect(url_for('clo_attainment'))

    course_info_for_report = get_course_report_info()
    course_report_inputs, error_response = read_course_report_export_inputs(
        redirect_url,
        stats,
        course_info_for_report,
        total_students
    )
    if error_response:
        return error_response

    try:
        docx_bytes = build_course_report_docx(stats, course_report_inputs, course_info_for_report, total_students)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(redirect_url)

    return course_report_docx_response(docx_bytes)

if __name__ == '__main__':
    port = int(os.environ.get('PORT') or 8093)
    print(f"Starting ETQAN v2 on http://127.0.0.1:{port}")
    app.run(port=port, debug=True)
