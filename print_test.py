import pandas as pd
import random
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# Load the Excel file
file_path = "test.xlsx"
df = pd.read_excel(file_path)

# Extract words, their meanings, level, and pronunciation
words = df[['单词', '中文', '等级', '发音']].drop_duplicates()

# Function to create a random vocabulary test and answers
def create_vocabulary_test(words, num_words=20):
    # Select random words
    selected_words = words.sample(num_words)
    
    # Create test and answer lists
    test_list = selected_words[['等级', '发音', '中文']].values.tolist()
    answer_words = selected_words[['单词']].values.flatten().tolist()
    
    return test_list, answer_words

# Paths to fonts
#chinese_font_path = '/Users/apple/Library/Fonts/NotoSansCJKscVF.ttf'
chinese_font_path = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
phonetic_font_path = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'

# Register the fonts
try:
    pdfmetrics.registerFont(TTFont('NotoSansCJK', chinese_font_path))
    pdfmetrics.registerFont(TTFont('ArialUnicode', phonetic_font_path))
    chinese_font_name = 'NotoSansCJK'
    phonetic_font_name = 'ArialUnicode'
except Exception as e:
    print(f"Error registering fonts: {e}")
    chinese_font_name = 'Helvetica'  # Fall back to Helvetica if the fonts are not found
    phonetic_font_name = 'Helvetica'

# Function to generate PDF test paper as a table
def generate_test_paper_as_table(filename, title, content_list):
    doc = SimpleDocTemplate(filename)
    elements = []

    # Create the table data with headers
    table_data = [['编号', '等级', '发音', '词义', '默写']]
    for idx, content in enumerate(content_list, 1):
        level, pronunciation, meaning = content
        table_data.append([str(idx), level, pronunciation, meaning, ""])

    # Create the table
    table = Table(table_data, colWidths=[40, 50, 100, 250, 100])
    table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), chinese_font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('FONTNAME', (2, 1), (2, -1), phonetic_font_name),  # Set phonetic font for pronunciation column
        ('FONTNAME', (0, 0), (-1, 0), chinese_font_name),  # Set font for the header row
        ('FONTWEIGHT', (0, 0), (-1, 0), 'bold'),  # Make header row text bold
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    # Add the table to the elements
    elements.append(table)

    # Build the PDF
    doc.build(elements)

# Function to generate answer sheet PDF
def generate_answer_sheet(filename, title, answer_list):
    c = canvas.Canvas(filename)
    c.setFont(chinese_font_name, 16)
    c.drawString(100, 800, title)

    # Set font for the content
    y_position = 780
    c.setFont(chinese_font_name, 12)
    for idx, answer in enumerate(answer_list, 1):
        line = f"{idx}. {answer}"
        c.drawString(50, y_position, line)

        y_position -= 20
        if y_position < 50:  # Add a new page if content exceeds the page
            c.showPage()
            y_position = 800
            c.setFont(chinese_font_name, 12)

    c.save()

# Function to generate multiple random test papers
def generate_multiple_tests(num_tests, num_words=20):
    for i in range(1, num_tests + 1):
        test_list, answer_words = create_vocabulary_test(words, num_words)
        test_filename = f"vocabulary_test_table_{i}.pdf"
        answer_filename = f"vocabulary_answers_{i}.pdf"
        generate_test_paper_as_table(test_filename, f"Vocabulary Test {i}", test_list)
        generate_answer_sheet(answer_filename, f"Vocabulary Test Answers {i}", answer_words)

# Generate 5 random test papers
generate_multiple_tests(num_tests=1, num_words=36)
