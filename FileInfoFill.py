import re
import pypdf
import pdfrw


def write_to_pdf(filename, work_order_num, address, description):
    # Read the PDF file
    pdf = pdfrw.PdfReader(filename)

    # Get the form fields
    form_fields = pdf.Root.AcroForm.Fields

    # Print the form field values
    for field in form_fields:
        if not field.V:
            name = field.T
            if name == "(Work Order #)":
                field.V = work_order_num
            elif name == "(Address)":
                field.V = address
            elif name == "(Description)":
                field.V = description

    pdfrw.PdfWriter().write(filename, pdf)


def get_pdf_info(filename):
    with open(filename, 'rb') as file:
        pdf_reader = pypdf.PdfReader(file)
        work_order_text = ""
        location_text = ""
        description_text = ""
        for page in range(len(pdf_reader.pages)):
            if work_order_text and location_text and description_text:
                break
            page_text = pdf_reader.pages[page].extract_text()
            work_order_match = re.search(r'Work Order #(.*?)Request By', page_text, re.DOTALL)
            if work_order_match:
                work_order_text = work_order_match.group(1).strip()
            location_match = re.search(r'Location.*?Location(.*?)Description', page_text, re.DOTALL)
            if location_match:
                location_text = location_match.group(1).strip().split(",")
            description_match = re.search(r'Description(?! of Work)(.*?)Closing Comments', page_text, re.DOTALL)
            if description_match:
                description_text = description_match.group(1).strip()

    return [location_text[0], description_text]


def read_compressed_pdf(filename):
    pdf_text = pypdf.PdfReader(filename)
    print(pdf_text.pages[2].extract_text())


def get_work_hours(filename):
    pdf_file = pypdf.PdfReader(filename)
    num_pages_to_check = min(3, len(pdf_file.pages))

    for page_num in range(num_pages_to_check):
        page_text = pdf_file.pages[page_num].extract_text()
        lines = page_text.strip().split('\n')

        if len(lines) >= 2:
            second_line = lines[1]
            date_format_pattern = r"\d{1,2}/\d{1,2}/\d{1,2}"
            if re.search(date_format_pattern, second_line):
                return second_line.split(' ')[1:]
    return None
