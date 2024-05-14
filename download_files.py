from office365_api import SharePoint
import re
from io import BytesIO
import pandas as pd
from docx import Document

# 1 args = SharePoint folder name. May include subfolders YouTube/2022
# FOLDER_NAME = sys.argv[1]
# 2 args = SharePoint file name. This is used when only one file is being downloaded
# If all files will be downloaded, then set this value as "None"
# FILE_NAME = sys.argv[2]
# 3 args = SharePoint file name pattern
# If no pattern match files are required to be downloaded, then set this value as "None"
# FILE_NAME_PATTERN = sys.argv[3]


def get_file_list(folder_name):
    file_list = SharePoint()._get_files_list(folder_name)
    files = []
    for file in file_list:
        FILE_NAME = file.properties["Name"]
        extension = get_file_extension(FILE_NAME)
        if extension:
            if extension == "docx":
                files.append(FILE_NAME)
    if files:
        crate_table(files)


def crate_table(files):
    DOCUMENTS = []
    document = {}
    for FILE_NAME in files:
        file_obj = SharePoint().download_file(FILE_NAME, FOLDER_NAME)
        CONTENT = get_file_text(file_obj)
        document = {'Titulo': FILE_NAME, 'Conteudo': CONTENT}
        DOCUMENTS.append(document)
    df = pd.DataFrame(DOCUMENTS)
    df.columns = ["Titulo", "Conteudo"]
    print(df)
    print("OK?")


def get_file_extension(filename):
  parts = filename.split('.')
  if len(parts) > 1:
    return parts[-1].lower()  # Ensure lowercase extension
  else:
    return None


def get_file_text(file_obj):
    full_text = []
    try:
        data = Document(BytesIO(file_obj))
        for paragraph in data.paragraphs:
            full_text.append(paragraph.text)
    except ImportError:
        print(
            "Warning: python-docx library not installed. Text extraction for .docx might be incomplete.")
    context = '\n'.join(full_text)
    return context


def upload_file(file_name, folder_name, content):
    SharePoint().upload_file(file_name, folder_name, content)


def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    get_file_text(file_obj)


def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)


def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)


if __name__ == '__main__':
    FOLDER_NAME = "General/2. INSTRUÇÕES DE TRABALHO/1. INSTRUÇÃO DE TRABALHO"

    if FOLDER_NAME:
        get_file_list(FOLDER_NAME)

    # FILE_NAME = "IT-0001 REV.06 - PREPARAÇÃO DE MANGUEIRA COMPOSTA.docx"
    # FILE_NAME_PATTERN = None

    # if FILE_NAME != 'None':
    #     get_file(FILE_NAME, FOLDER_NAME)
    # elif FILE_NAME_PATTERN != 'None':
    #     get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
    # else:
    #     get_files(FOLDER_NAME)