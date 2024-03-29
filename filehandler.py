import os
import io
import re
import shutil
from global_vars import folders

# Install imports
try:
    import pdfminer
    from pdfminer.converter import TextConverter
    from pdfminer.pdfinterp import PDFPageInterpreter
    from pdfminer.pdfinterp import PDFResourceManager
    from pdfminer.pdfpage import PDFPage
except ModuleNotFoundError:
    print("pip install --user pdfminer")
    os._exit(0)
    
try:
    import tabula
except ModuleNotFoundError:
    print("pip install --user tabula-py")
    os._exit(0)

# Globals
regex_station = re.compile(r"(MSA|MSB)")
regex_line = re.compile("linje " + r"[0-9]{1}[0-1]?")
regex_type = re.compile(r"[0-9]{2}:[0-9]{2}" + "(14|1)")

def make_folders(settings):
    for f in folders:
        if not os.path.isdir(f"{settings['out_folder']}/{f}"):
            os.mkdir(f"{settings['out_folder']}/{f}")
    return folders
    

def find_pdffiles(settings):
    pdffiles = list()
    print(settings)
    for file in os.listdir(settings["data_folder"]):
        if file.endswith(".pdf"):
            pdffiles.append(file)
    return pdffiles

def rename_files(path, text, out_folder):
    try:
        station = re.findall(regex_station, text)[0]

    except IndexError:
        print(f"Fant ingen målestasjon i fil: {path}")
        return None

    try:
        line = re.findall(regex_line, text)[0].split(" ")[1]
    except IndexError:
        print(f"Fant ingen løp i fil: {path}")
        return None

    try:
        oil_type = re.findall(regex_type, text)[0]
        if oil_type == "0":
            return None

    except IndexError:
        print(f"Fant ingen oljetype, (oljetype 0) i fil {path}")
        return None

    new_path = f"{out_folder}/{station}_{oil_type}/{station}_{line}_{oil_type}.pdf"

    try:
        shutil.move(path, new_path)
    except shutil.SameFileError:
        print(f"filename already exists: {path}")
        return None


def extract_text(path):
    resource_manager = PDFResourceManager()
    fake_file_handler = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handler)
    page_interpreter = PDFPageInterpreter(resource_manager, converter)
    
    with open(path, "rb") as obj:
        for page in PDFPage.get_pages(obj, caching=True, check_extractable=True):
            page_interpreter.process_page(page)
            text = fake_file_handler.getvalue()
            break

    return text

def rename_and_move_files(settings):
    folders = make_folders(settings)
    pdffiles = find_pdffiles(settings)
    for file in pdffiles:
        try:
            text = extract_text(f"{settings['data_folder']}/{file}")
        except pdfminer.pdfparser.PDFSyntaxError:
            print(f"EOF error in file: {file}\nFile will be removed")
            os.remove(f"{settings['data_folder']}/{file}")
            continue
        rename_files(f"{settings['data_folder']}/{file}", text, settings["out_folder"])

def convert_pdffiles_to_csv(settings):
    print("Converting pdf files to csv...")
    for f in folders: 
        tabula.convert_into_by_batch(f"{settings['out_folder']}/{f}", output_format="csv", pages="all") # java_options=["java.awt.headless=true"])
    print("pdf files converted to csv")


# if __name__ == "__main__":
#        print("Yeet")
#        settings = {"data_folder": "./data"}
#        rename_and_move_files() 
#        convert_pdffiles_to_csv()
