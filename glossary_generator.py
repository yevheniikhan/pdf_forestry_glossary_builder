#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
ГЕНЕРАТОР ГЛОСАРІЮ ЛІСІВНИЧИХ ТЕРМІНІВ
================================================================================

Скрипт для автоматичної генерації PDF-глосарію з Excel-таблиць.
Створює професійно оформлений документ із термінами, визначеннями та перекладами.

Автор: Євгеній Хань
Версія: 1.0
Ліцензія: CC BY-SA 4.0

Google Colab: https://colab.research.google.com/drive/1D9BGjpHWCxUOBsUBukOLxYZlbp_efrla?usp=sharing

Вхідні дані:
    - Excel-файл з визначеннями (колонки: TERM, UKRAINIAN EQUIVALENT,
      ORIGINAL DEFINITION, ORIGINAL DEFINITION (TRANSLATION), SOURCE)
    - Excel-файл з термінами без визначень (колонки: Term, Ukrainian Equivalent)

Вихідні дані:
    - glossary.pdf — готовий глосарій
    - glossary.tex — LaTeX-файл для подальшого редагування

Використання:
    Запустіть у Google Colab, послідовно виконуючи комірки.
    Скрипт запитає необхідні файли та параметри.

================================================================================
"""

# ==============================================================================
# КРОК 1: ВСТАНОВЛЕННЯ ЗАЛЕЖНОСТЕЙ
# ==============================================================================
# Встановлення LaTeX-дистрибутива та Python-бібліотек для роботи з Excel

!apt-get update -qq
!apt-get install -y texlive-latex-base texlive-fonts-recommended \
    texlive-fonts-extra texlive-latex-extra texlive-lang-cyrillic \
    texlive-xetex -qq

!pip install pandas openpyxl -q

print("✓ Залежності встановлено")

# ==============================================================================
# КРОК 2: ІМПОРТ БІБЛІОТЕК
# ==============================================================================

import pandas as pd
import subprocess
import os
import re
from datetime import datetime
from google.colab import files

print("✓ Бібліотеки імпортовано")

# ==============================================================================
# КРОК 3: НАЛАШТУВАННЯ ПАРАМЕТРІВ ДОКУМЕНТА
# ==============================================================================
# Введення дати та тексту застереження

# --- Дата документа ---
print("=" * 60)
print("НАЛАШТУВАННЯ ДАТИ ДОКУМЕНТА")
print("=" * 60)

# Отримання поточної дати українською
MONTHS_UA = {
    1: "січня", 2: "лютого", 3: "березня", 4: "квітня",
    5: "травня", 6: "червня", 7: "липня", 8: "серпня",
    9: "вересня", 10: "жовтня", 11: "листопада", 12: "грудня"
}

today = datetime.now()
current_date_ua = f"{today.day} {MONTHS_UA[today.month]} {today.year} року"

print(f"\nПоточна дата: {current_date_ua}")
change_date = input("Бажаєте змінити дату? (так/ні): ").strip().lower()

if change_date in ['так', 'yes', 'y', 'т']:
    DOCUMENT_DATE = input("Введіть нову дату (наприклад, '1 жовтня 2025 року'): ").strip()
    if not DOCUMENT_DATE:
        DOCUMENT_DATE = current_date_ua
else:
    DOCUMENT_DATE = current_date_ua

print(f"✓ Встановлено дату: {DOCUMENT_DATE}")

# --- Текст застереження ---
print("\n" + "=" * 60)
print("НАЛАШТУВАННЯ ТЕКСТУ ЗАСТЕРЕЖЕННЯ")
print("=" * 60)

DEFAULT_DISCLAIMER = """Цей глосарій є автоматично згенерованим довідковим матеріалом на основі наданих вхідних даних. Глосарій призначений для кращого розуміння змісту документів, дотичних до лісового сектору, мовою оригіналу, а також для підтримки інтеграції машинного перекладу та контекстного перекладу при використанні рішень на основі великих мовних моделей (LLM). Перелік термінів не є вичерпним і може оновлюватися та змінюватися. Визначення термінів можуть походити з кількох джерел для забезпечення повнішого розуміння контексту. Визначення зазвичай перекладались з мови оригіналу українською, однак залежно від контексту могли бути спрощені або адаптовані. Глосарій має виключно інформаційний характер і не є офіційним перекладом чи тлумаченням термінології. При виявленні розбіжностей із першоджерелами пріоритет мають оригінальні документи. Укладачі не несуть відповідальності за можливі неточності або наслідки використання цього матеріалу."""

print("\nПоточний текст застереження:")
print("-" * 40)
print(DEFAULT_DISCLAIMER[:200] + "...")
print("-" * 40)

change_disclaimer = input("\nБажаєте змінити текст застереження? (так/ні): ").strip().lower()

if change_disclaimer in ['так', 'yes', 'y', 'т']:
    print("\nВведіть новий текст застереження (для завершення введіть порожній рядок):")
    lines = []
    while True:
        line = input()
        if line == "":
            break
        lines.append(line)

    if lines:
        DISCLAIMER_TEXT = " ".join(lines)
    else:
        DISCLAIMER_TEXT = DEFAULT_DISCLAIMER
else:
    DISCLAIMER_TEXT = DEFAULT_DISCLAIMER

print(f"✓ Текст застереження {'оновлено' if change_disclaimer in ['так', 'yes', 'y', 'т'] else 'залишено без змін'}")

# ==============================================================================
# КРОК 4: ЗАВАНТАЖЕННЯ ВХІДНИХ ФАЙЛІВ
# ==============================================================================

# --- Файл з визначеннями ---
print("\n" + "=" * 60)
print("ЗАВАНТАЖЕННЯ ФАЙЛУ З ВИЗНАЧЕННЯМИ")
print("=" * 60)
print("Очікувані колонки: TERM, UKRAINIAN EQUIVALENT, ORIGINAL DEFINITION,")
print("                   ORIGINAL DEFINITION (TRANSLATION), SOURCE")
print("-" * 60)

uploaded1 = files.upload()
excel_file_definitions = list(uploaded1.keys())[0]
print(f"✓ Завантажено: {excel_file_definitions}")

# --- Файл з термінами ---
print("\n" + "=" * 60)
print("ЗАВАНТАЖЕННЯ ФАЙЛУ З ТЕРМІНАМИ (без визначень)")
print("=" * 60)
print("Очікувані колонки: Term, Ukrainian Equivalent")
print("-" * 60)

uploaded2 = files.upload()
excel_file_terms = list(uploaded2.keys())[0]
print(f"✓ Завантажено: {excel_file_terms}")

# ==============================================================================
# КРОК 5: ЧИТАННЯ ТА ОБРОБКА ДАНИХ
# ==============================================================================

df_definitions = pd.read_excel(excel_file_definitions)
df_terms = pd.read_excel(excel_file_terms)

print(f"\n{'=' * 60}")
print("СТАТИСТИКА ВХІДНИХ ДАНИХ")
print("=" * 60)
print(f"Файл з визначеннями:")
print(f"  • Колонки: {list(df_definitions.columns)}")
print(f"  • Кількість термінів: {len(df_definitions)}")
print(f"\nФайл з термінами:")
print(f"  • Колонки: {list(df_terms.columns)}")
print(f"  • Кількість термінів: {len(df_terms)}")

# ==============================================================================
# КРОК 6: ДОПОМІЖНІ ФУНКЦІЇ
# ==============================================================================

def normalize_columns(df):
    """
    Нормалізує назви колонок DataFrame до стандартного формату.

    Підтримує різні варіанти написання назв колонок (з великої/малої літери,
    з пробілами/підкресленнями) та приводить їх до єдиного формату.

    Args:
        df: pandas DataFrame з даними

    Returns:
        DataFrame з нормалізованими назвами колонок
    """
    column_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'term':
            column_mapping[col] = 'TERM'
        elif col_lower in ['ukrainian equivalent', 'ukrainian_equivalent']:
            column_mapping[col] = 'UKRAINIAN EQUIVALENT'
        elif col_lower in ['original definition', 'original_definition']:
            column_mapping[col] = 'ORIGINAL DEFINITION'
        elif 'translation' in col_lower:
            column_mapping[col] = 'ORIGINAL DEFINITION (TRANSLATION)'
        elif col_lower == 'source':
            column_mapping[col] = 'SOURCE'

    return df.rename(columns=column_mapping)


def escape_latex_safe(text):
    """
    Екранує спеціальні символи LaTeX та замінює проблемні Unicode-символи.

    LaTeX використовує ряд символів для форматування (&, %, $, #, _, {, }, тощо).
    Ця функція замінює їх на безпечні еквіваленти для коректного відображення.

    Args:
        text: вхідний текст для обробки

    Returns:
        Екранований текст, безпечний для використання в LaTeX
    """
    if pd.isna(text) or text is None:
        return ""

    text = str(text).strip()

    # Видалення символів нового рядка
    text = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

    # Заміна фігурних дужок на круглі (запобігає помилкам парсингу)
    text = text.replace('{', '(')
    text = text.replace('}', ')')

    # Екранування спеціальних символів LaTeX
    text = text.replace('\\', '')
    text = text.replace('&', r'\&')
    text = text.replace('%', r'\%')
    text = text.replace('$', r'\$')
    text = text.replace('#', r'\#')
    text = text.replace('_', r'\_')
    text = text.replace('~', r'\textasciitilde{}')
    text = text.replace('^', r'\textasciicircum{}')
    text = text.replace('<', r'\textless{}')
    text = text.replace('>', r'\textgreater{}')
    text = text.replace('|', r'\textbar{}')

    # Заміна типографських символів Unicode на ASCII-еквіваленти
    text = text.replace('–', '--')      # en-dash
    text = text.replace('—', '---')     # em-dash
    text = text.replace('"', "''")      # закриваюча лапка
    text = text.replace('"', "''")      # відкриваюча лапка
    text = text.replace('„', ",,")      # нижня лапка
    text = text.replace(''', "'")       # апостроф
    text = text.replace(''', "'")       # апостроф
    text = text.replace('…', '...')     # три крапки
    text = text.replace('•', '-')       # маркер списку
    text = text.replace('·', '-')       # середня крапка
    text = text.replace('°', ' deg ')   # градус
    text = text.replace('±', '+/-')     # плюс-мінус
    text = text.replace('×', 'x')       # множення
    text = text.replace('÷', '/')       # ділення
    text = text.replace('≤', '<=')      # менше або дорівнює
    text = text.replace('≥', '>=')      # більше або дорівнює
    text = text.replace('≈', '~')       # приблизно
    text = text.replace('²', '2')       # степінь 2
    text = text.replace('³', '3')       # степінь 3
    text = text.replace('¹', '1')       # степінь 1
    text = text.replace('½', '1/2')     # одна друга
    text = text.replace('¼', '1/4')     # одна четверта
    text = text.replace('¾', '3/4')     # три четвертих

    # Видалення подвійних пробілів
    while '  ' in text:
        text = text.replace('  ', ' ')

    return text.strip()


def format_source_with_url(text):
    """
    Форматує текст джерела з виділенням URL на окремому рядку.

    Знаходить URL у тексті та розміщує його на новому рядку
    для кращої читабельності.

    Args:
        text: текст джерела, що може містити URL

    Returns:
        Відформатований текст з URL на окремому рядку
    """
    if pd.isna(text) or text is None:
        return ""

    text = str(text).strip()

    # Пошук URL у тексті
    url_pattern = r'(https?://[^\s]+)'
    match = re.search(url_pattern, text)

    if match:
        url = match.group(1)
        # Видалення зайвих символів пунктуації в кінці URL
        url = re.sub(r'[.,;:\)\]]+$', '', url)

        # Текст до URL
        text_part = text[:match.start()].strip()
        text_part = escape_latex_safe(text_part)

        # Формування результату: текст + URL на новому рядку
        if text_part:
            return text_part + r' \newline \url{' + url + '}'
        else:
            return r'\url{' + url + '}'

    return escape_latex_safe(text)


# Нормалізація назв колонок
df_definitions = normalize_columns(df_definitions)
df_terms = normalize_columns(df_terms)

print(f"\n✓ Колонки нормалізовано")

# ==============================================================================
# КРОК 7: ГЕНЕРАЦІЯ LATEX-ДОКУМЕНТА
# ==============================================================================

def generate_latex(df_definitions, df_terms, document_date, disclaimer_text):
    """
    Генерує повний LaTeX-документ глосарію.

    Створює структурований документ із титульною сторінкою, застереженням,
    основною частиною (терміни з визначеннями) та додатком (список термінів).

    Args:
        df_definitions: DataFrame з термінами та визначеннями
        df_terms: DataFrame з термінами без визначень (для таблиці)
        document_date: дата документа для відображення
        disclaimer_text: текст застереження

    Returns:
        Рядок із повним LaTeX-кодом документа
    """

    # Підготовка даних
    df_def = df_definitions.dropna(subset=['TERM'])
    df_def_sorted = df_def.sort_values('TERM', key=lambda x: x.str.lower())

    df_trm = df_terms.dropna(subset=['TERM'])
    df_trm_sorted = df_trm.sort_values('TERM', key=lambda x: x.str.lower())

    # -------------------------------------------------------------------------
    # ПРЕАМБУЛА ДОКУМЕНТА
    # -------------------------------------------------------------------------
    latex = r"""\documentclass[a4paper, 11pt]{article}

% -----------------------------------------------------------------------------
% НАЛАШТУВАННЯ КОДУВАННЯ ТА ШРИФТІВ
% -----------------------------------------------------------------------------
\usepackage{fontspec}
\usepackage[ukrainian,english]{babel}
\usepackage[margin=2.5cm, top=2cm, bottom=2cm]{geometry}

\setmainfont{DejaVu Serif}
\setsansfont{DejaVu Sans}

% -----------------------------------------------------------------------------
% КОЛЬОРОВА СХЕМА (темно-зелені відтінки)
% -----------------------------------------------------------------------------
\usepackage{xcolor}
\definecolor{termcolor}{RGB}{0, 80, 40}
\definecolor{headercolor}{RGB}{0, 100, 50}
\definecolor{linkcolor}{RGB}{0, 80, 40}
\definecolor{boxcolor}{RGB}{245, 250, 247}
\definecolor{linecolor}{RGB}{180, 210, 190}

% -----------------------------------------------------------------------------
% ДОДАТКОВІ ПАКЕТИ
% -----------------------------------------------------------------------------
\usepackage{parskip}
\usepackage{fancyhdr}
\usepackage{tcolorbox}
\usepackage{longtable}
\usepackage{array}

% Виправлення широких пробілів після крапок
\frenchspacing

% -----------------------------------------------------------------------------
% НАЛАШТУВАННЯ ГІПЕРПОСИЛАНЬ
% -----------------------------------------------------------------------------
\usepackage{hyperref}
\hypersetup{
    colorlinks=true,
    linkcolor=linkcolor,
    urlcolor=linkcolor,
    breaklinks=true,
    pdfstartview=FitH,
    pdftitle={Glossary of Forestry Terms},
    pdfauthor={-}
}

% -----------------------------------------------------------------------------
% КОЛОНТИТУЛИ
% -----------------------------------------------------------------------------
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\textit{Глосарій термінів}}
\fancyhead[R]{\small\thepage}
\fancyfoot[C]{\small\textit{Forestry Terminology Glossary}}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\footrulewidth}{0.4pt}

% =============================================================================
% ПОЧАТОК ДОКУМЕНТА
% =============================================================================
\begin{document}

% -----------------------------------------------------------------------------
% ТИТУЛЬНА СТОРІНКА
% -----------------------------------------------------------------------------
\thispagestyle{empty}
\begin{center}
    \vspace*{2cm}
    {\Huge\bfseries\color{headercolor} ГЛОСАРІЙ ТЕРМІНІВ \par}
    \vspace{0.8cm}
    {\Large Forestry Terminology Glossary \par}
    \vspace{2cm}
    \hrule height 2pt
    \vspace{0.5cm}

    {\small
    Документ згенеровано автоматично за допомогою скрипта для створення глосаріїв лісівничих термінів на основі наданих вхідних Excel-файлів.
    \par}

    \vspace{0.5cm}
    \hrule height 2pt
\end{center}

\vspace{1.5cm}

% -----------------------------------------------------------------------------
% БЛОК ЗАСТЕРЕЖЕННЯ
% -----------------------------------------------------------------------------
\begin{tcolorbox}[
    colback=boxcolor,
    colframe=linecolor,
    boxrule=0.5pt,
    arc=3pt,
    left=12pt,
    right=12pt,
    top=10pt,
    bottom=10pt
]
{\large\bfseries\color{headercolor} Застереження}

\vspace{8pt}

{\small
""" + escape_latex_safe(disclaimer_text) + r"""

\vspace{8pt}

\textbf{CC BY-SA 4.0}
}
\end{tcolorbox}

\vfill

\begin{center}
{\large """ + document_date + r""" \par}
\end{center}

\newpage

"""

    # -------------------------------------------------------------------------
    # ЧАСТИНА 1: ТЕРМІНИ З ВИЗНАЧЕННЯМИ
    # -------------------------------------------------------------------------
    current_letter = ""
    processed = 0

    for idx, row in df_def_sorted.iterrows():
        try:
            term_raw = row.get('TERM', '')
            if pd.isna(term_raw) or not str(term_raw).strip():
                continue

            term = escape_latex_safe(term_raw)
            first_letter = str(term_raw).strip()[0].upper()

            # Нова літера — нова сторінка
            if first_letter != current_letter and first_letter.isalpha():
                if current_letter != "":
                    latex += r"\newpage" + "\n"
                current_letter = first_letter
                latex += f"""
\\noindent{{\\Huge\\bfseries\\color{{headercolor}} {current_letter}}}
\\vspace{{5pt}}
\\hrule height 2pt
\\vspace{{10pt}}

"""

            # Отримання та екранування полів
            ukr = escape_latex_safe(row.get('UKRAINIAN EQUIVALENT', ''))
            orig_def = escape_latex_safe(row.get('ORIGINAL DEFINITION', ''))
            ukr_def = escape_latex_safe(row.get('ORIGINAL DEFINITION (TRANSLATION)', ''))
            source_formatted = format_source_with_url(row.get('SOURCE', ''))

            # Генерація блоку терміну
            latex += r"""
\begin{tcolorbox}[
    colback=boxcolor,
    colframe=linecolor,
    boxrule=0.5pt,
    arc=3pt,
    left=12pt,
    right=12pt,
    top=10pt,
    bottom=10pt,
    before skip=12pt,
    after skip=12pt
]
"""
            latex += f"{{\\Large\\bfseries\\color{{termcolor}} {term}}}\n"
            latex += r"\vspace{10pt}" + "\n\n"
            latex += f"\\textbf{{Український еквівалент:}} {ukr}\n\n"
            latex += r"\vspace{4pt}" + "\n"
            latex += f"\\textbf{{Оригінальне визначення (англ.):}} \\textit{{{orig_def}}}\n\n"
            latex += r"\vspace{4pt}" + "\n"
            latex += f"\\textbf{{Переклад:}} \\textit{{{ukr_def}}}\n\n"
            latex += r"\vspace{6pt}" + "\n"
            latex += f"{{\\footnotesize \\textbf{{Джерело:}} {source_formatted}}}\n"
            latex += r"\end{tcolorbox}" + "\n\n"

            processed += 1

        except Exception as e:
            print(f"⚠ Помилка обробки терміну: {e}")

    # -------------------------------------------------------------------------
    # ЧАСТИНА 2: ТАБЛИЦЯ ТЕРМІНІВ БЕЗ ВИЗНАЧЕНЬ
    # -------------------------------------------------------------------------
    latex += r"""
\newpage

% -----------------------------------------------------------------------------
% ДОДАТОК: ТАБЛИЦЯ ТЕРМІНІВ ТА ПЕРЕКЛАДІВ
% -----------------------------------------------------------------------------
\noindent{\Huge\bfseries\color{headercolor} Терміни та переклади}
\vspace{5pt}
\hrule height 2pt
\vspace{15pt}

\begin{longtable}{p{0.48\textwidth} >{\raggedright\arraybackslash}p{0.48\textwidth}}
\textbf{\color{termcolor}Термін} & \textbf{\color{termcolor}Український еквівалент} \\
\hline
\endhead
"""

    for idx, row in df_trm_sorted.iterrows():
        term_raw = row.get('TERM', '')
        if pd.isna(term_raw) or not str(term_raw).strip():
            continue

        term = escape_latex_safe(term_raw)
        ukr = escape_latex_safe(row.get('UKRAINIAN EQUIVALENT', ''))

        latex += f"{term} & {ukr} \\\\\n"

    latex += r"""
\end{longtable}

\end{document}
"""

    print(f"\n✓ Оброблено термінів з визначеннями: {processed}")
    print(f"✓ Термінів у таблиці: {len(df_trm_sorted)}")

    return latex


# Генерація LaTeX-коду
print("\n" + "=" * 60)
print("ГЕНЕРАЦІЯ LATEX-ДОКУМЕНТА")
print("=" * 60)

latex_content = generate_latex(df_definitions, df_terms, DOCUMENT_DATE, DISCLAIMER_TEXT)

# Збереження у файл
with open('glossary.tex', 'w', encoding='utf-8') as f:
    f.write(latex_content)

print(f"✓ LaTeX-файл створено: glossary.tex ({os.path.getsize('glossary.tex'):,} байт)")

# ==============================================================================
# КРОК 8: КОМПІЛЯЦІЯ PDF
# ==============================================================================

print("\n" + "=" * 60)
print("КОМПІЛЯЦІЯ PDF")
print("=" * 60)
print("Використовується XeLaTeX для підтримки Unicode...")

for i in range(2):
    result = subprocess.run(
        ['xelatex', '-interaction=nonstopmode', 'glossary.tex'],
        capture_output=True,
        text=True
    )
    print(f"  • Прохід {i+1}/2 завершено")

# Перевірка результату
if os.path.exists('glossary.pdf'):
    size_kb = os.path.getsize('glossary.pdf') / 1024
    print(f"\n✓ PDF успішно створено: glossary.pdf ({size_kb:.1f} KB)")

    if size_kb < 200:
        print("\n⚠ Увага: розмір файлу менший за очікуваний.")
        print("  Перевірка логу помилок...")

        if os.path.exists('glossary.log'):
            with open('glossary.log', 'r', errors='ignore') as f:
                log = f.read()
            errors = [l for l in log.split('\n') if l.startswith('!')]
            if errors:
                print("  Знайдені помилки LaTeX:")
                for e in errors[:5]:
                    print(f"    {e}")
    else:
        print("✓ Розмір файлу відповідає очікуваному")
else:
    print("\n✗ Помилка: PDF-файл не створено")
    print("  Перевірте лог glossary.log для деталей")

# ==============================================================================
# КРОК 9: ЗАВАНТАЖЕННЯ РЕЗУЛЬТАТІВ
# ==============================================================================

print("\n" + "=" * 60)
print("ЗАВАНТАЖЕННЯ ФАЙЛІВ")
print("=" * 60)

if os.path.exists('glossary.pdf'):
    files.download('glossary.pdf')
    print("✓ glossary.pdf — готовий глосарій")

files.download('glossary.tex')
print("✓ glossary.tex — вихідний код для редагування")

print("\n" + "=" * 60)
print("ГОТОВО!")
print("=" * 60)
