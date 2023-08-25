from docx import Document

def extract_table_from_docx(file_path):
    """Extracts table data from a Word document and returns it as a list of rows."""
    doc = Document(file_path)
    table = doc.tables[0]  # Assuming only one table in the document

    data = []
    for row in table.rows:
        rowData = [cell.text for cell in row.cells]
        data.append(rowData)
    return data

def save_as_csv(data, output_file):
    """Saves table data as a CSV file."""
    with open(output_file, 'w', encoding='utf-8') as file:
        for row in data:
            file.write(','.join(['"' + cell.replace('"', '""') + '"' for cell in row]) + '\n')

if __name__ == "__main__":
    file_path = 'vocab.docx'
    output_csv = 'output.csv'

    table_data = extract_table_from_docx(file_path)
    save_as_csv(table_data, output_csv)
    print(f"Data saved to {output_csv}")
