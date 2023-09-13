import json
from docx import Document

def load_formatting_standards(file_path):
    with open(file_path, 'r') as json_file:
        standards = json.load(json_file)
    return standards

def check_formatting(docx_file, standards):
    doc = Document(docx_file)
    errors = []

    for paragraph_number, paragraph in enumerate(doc.paragraphs, start=1):
        if not paragraph.runs:
            continue

        # Check if the paragraph contains Persian text
        contains_persian = contains_persian_text(paragraph.text)

        # Combine runs within a paragraph into a single text block
        paragraph_text = ''.join(run.text for run in paragraph.runs if run.text)

        # Check font name (case-insensitive)
        font_names = [run.font.name.lower() for run in paragraph.runs if run.font and run.font.name]
        if contains_persian and standards["persian_font"].lower() not in font_names:
            errors.append(f"Font in Persian paragraph {paragraph_number} is not {standards['persian_font']}")
        elif not contains_persian and standards["english_font"].lower() not in font_names:
            errors.append(f"Font in paragraph {paragraph_number} is not {standards['english_font']}")

        # Check text size if available
        text_sizes = [run.font.size.pt for run in paragraph.runs if run.font and run.font.size and run.font.size.pt]
        if text_sizes:
            avg_text_size = sum(text_sizes) / len(text_sizes)
            if abs(avg_text_size - standards["font_size"]) > 0.01:
                if contains_persian:
                    errors.append(f"Text size in Persian paragraph {paragraph_number} is not {standards['font_size']} pt")
                else:
                    errors.append(f"Text size in paragraph {paragraph_number} is not {standards['font_size']} pt")

        # Check line spacing (approximate check)
        if paragraph.alignment != 0 and paragraph.paragraph_format.line_spacing is not None:
            line_spacing = paragraph.paragraph_format.line_spacing
            if abs(line_spacing - standards["line_spacing"]) > 0.01:
                if contains_persian:
                    errors.append(f"Line spacing in Persian paragraph {paragraph_number} is not {standards['line_spacing']}")
                else:
                    errors.append(f"Line spacing in paragraph {paragraph_number} is not {standards['line_spacing']}")

    return errors



# Function to check if a string contains Persian characters
def contains_persian_text(text):
    persian_chars = set("ابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی")
    return any(char in persian_chars for char in text)

def main():

    # get sample.docx and settings.json
    docx_file = 'sample.docx'
    standards_file = 'settings.json'

    # read the json file
    standards = load_formatting_standards(standards_file)

    #run the function
    errors = check_formatting(docx_file, standards)

    if errors:
        print("Formatting errors:")
        for error in errors:
            print(error)
    else:
        print("Formatting is correct.")

if __name__ == "__main__":
    main()
