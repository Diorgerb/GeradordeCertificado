import docx
import csv
import smtplib
import subprocess
import os
import win32com.client

def make_certificate(filename, name):
    doc = docx.Document(filename)
    for p in doc.paragraphs:
        if 'name' in p.text:
            inline = p.runs

            for i in range(len(inline)):
                if 'name' in inline[i].text:
                    inline[i].text = inline[i].text.replace('name', '')
                    inline[1].text = name + ' '
                    inline[1].bold = True

    output_folder = 'certificados'  # Folder to save the certificates
    os.makedirs(output_folder, exist_ok=True)  # Create the folder if it doesn't exist

    output_filename = os.path.join(output_folder, '{}.docx'.format(name))

    doc.save(output_filename)

    # Convert Word document to PDF using pywin32
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    doc_path = os.path.abspath(output_filename)
    pdf_path = os.path.abspath(output_filename.replace('.docx', '.pdf'))
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

    # Delete the temporary Word document
    os.remove(doc_path)

def certificate(filename):
    with open(filename, 'r', encoding='utf-8-sig') as csv_file:
        attendants = csv.reader(csv_file, delimiter=',')
        header = next(attendants)  # Read the header row

        name_column_index = header.index('Aluno')  # Get the index of the 'Aluno' column

        for row in attendants:
            name = row[name_column_index]  # Retrieve the name from the 'Aluno' column
            make_certificate('certificadopadrao.docx', name)


if __name__ == '__main__':
    certificate('resultados.csv')
