# Генератор глосарію лісівничих термінів

Скрипт для автоматичної генерації PDF-глосарію з Excel-таблиць. Створює професійно оформлений документ із термінами, визначеннями та перекладами.

## Запуск

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/drive/1D9BGjpHWCxUOBsUBukOLxYZlbp_efrla?usp=sharing)

## Структура проекту

```
pdf_forestry_glossary_builder/
├── glossary_generator.py   # Основний скрипт
├── input/                  # Папка для вхідних Excel-файлів
├── output/                 # Папка для згенерованих PDF та LaTeX
└── README.md
```

## Вхідні дані

Скрипт очікує два Excel-файли:

### 1. Файл з визначеннями
Колонки:
- `TERM` — термін англійською
- `UKRAINIAN EQUIVALENT` — український еквівалент
- `ORIGINAL DEFINITION` — оригінальне визначення (англ.)
- `ORIGINAL DEFINITION (TRANSLATION)` — переклад визначення
- `SOURCE` — джерело

### 2. Файл з термінами (без визначень)
Колонки:
- `Term` — термін англійською
- `Ukrainian Equivalent` — український еквівалент

## Вихідні дані

- `glossary.pdf` — готовий глосарій у форматі PDF
- `glossary.tex` — LaTeX-файл для подальшого редагування

## Використання

1. Відкрийте [Google Colab](https://colab.research.google.com/drive/1D9BGjpHWCxUOBsUBukOLxYZlbp_efrla?usp=sharing)
2. Послідовно виконуйте комірки
3. Завантажте вхідні Excel-файли коли буде запропоновано
4. Отримайте згенерований PDF-глосарій

## Ліцензія

CC BY-SA 4.0

## Автор

Євгеній Хань
