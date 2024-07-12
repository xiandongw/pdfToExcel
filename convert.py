import pandas as pd
import fitz  # PyMuPDF


# Define a function to extract text from the PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    pdf_document = fitz.open(pdf_path)
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text


def convert_to_excel(src, target, header=True, columns=2, beginText=None):
    # Path to the uploaded PDF
    pdf_path = src

    # Extract text from the PDF
    pdf_text = extract_text_from_pdf(pdf_path)

    # Split the text into lines
    lines = pdf_text.split('\n')
    if beginText and beginText in lines:
        lines = lines[lines.index(beginText):]
    df = []
    headers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    for idx in range(0, len(lines), columns):
        row = lines[idx:idx + columns]
        if idx == 0:
            if header:
                for i in range(0, columns):
                    df.append({'header': row[i], 'data': []})
            else:
                for i in range(0, columns):
                    df.append({'header': headers[i], 'data': [row[i]]})
        else:
            for i in range(0, columns):
                df[i]['data'].append(row[i] if i < len(row) else '')

    # Create a DataFrame from the extracted data
    df = pd.DataFrame({item['header']: item['data'] for item in df})

    # Save the DataFrame to an Excel file
    output_excel_path = target
    df.to_excel(output_excel_path, index=False)


if __name__ == "__main__":
    convert_to_excel(src='/Users/XiandongWang/Downloads/quince address update.pdf',
                     target='/Users/XiandongWang/Downloads/quince address update.xlsx',
                     header=True, columns=3, beginText='NAME')
