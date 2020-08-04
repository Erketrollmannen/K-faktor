#
# Sponset av nettverk
#
"""

""" 
# -----[ import stuff ]-----


import os
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import io

import shutil
import re

f = open("rundir.txt", "r")

directory = f.readline()

print("directory: " + directory)

#print(directory)

os.chdir(directory)

# -----[ Global Variables ]-----
folders = ["MSA_1", "MSA_14", "MSB_1", "MSB_14", "0"]
målestasjoner = "MSA|MSB"           
oljetyper = ["14", "1"]
pdffiles = []


# -----[ Regex Stuff ]-----

regexMålestasjon = re.compile(r"(MSA|MSB)")

regexLinje = re.compile("linje " + r"[0-9]{1}[0-1]?")

regexOljetype =  re.compile(r"[0-9]{2}:[0-9]{2}" + "(14|1)")



#print("regex Oljetyper: " + str(regexOljetype))
#print("regex Målestasjon: " + str(regexMålestasjon))


# Make folders

for f in folders:
    if not os.path.isdir(f):
        os.mkdir(f)       

# Make list of pdf files
for filename in os.listdir("."):
    if filename.endswith(".pdf"):
        pdffiles.append(filename)

print(pdffiles)


for filename in pdffiles:
    newfilename = ""
    matchOljetype = ""
    matchLinje = ""
    

    fh = open(filename, 'rb')
    for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):

        resource_manager = PDFResourceManager()
        fake_file_handle = io.StringIO()
        converter = TextConverter(resource_manager, fake_file_handle)
        page_interpreter = PDFPageInterpreter(resource_manager, converter)
        page_interpreter.process_page(page)
            
        text = fake_file_handle.getvalue()
        #yield text + " "

            # close open handles
            
        converter.close()
        fake_file_handle.close()   
        fh.close()

        #print()
        #print("Old filename. " + filename)

        Målestasjon = re.search(regexMålestasjon, text)

        #print("Målestasjon_regex: " + str(Målestasjon))

        
        if Målestasjon == None:
            break
        else:
            matchMålestasjon = Målestasjon.group(0)

        #print("Målestasjon" + str(matchMålestasjon))

        Linje = re.search(regexLinje, text)

        #print("Linje Regex: " + str(Linje))
        if Linje == None:
            break
        
        else:
            matchLinje = Linje.group(0)

        oljetype = re.search(regexOljetype, text)

        #print("Oljetype_regex: " + str(oljetype))
        if oljetype == None:
            break
        
        else:
            matchOljetype_mau = oljetype.group(0)

        

        if matchOljetype_mau.endswith("4"):
            matchOljetype = "14"
        if matchOljetype_mau.endswith("1"):
            matchOljetype = "1"            
        if matchOljetype_mau.endswith("0"):
            break


           
     
        newfilename = str(matchMålestasjon) + "_" + str(matchLinje) + "_" + matchOljetype + ".pdf"

        location = str(matchMålestasjon) + "_" + str(matchOljetype)

        #print("Målestasjon: " + str(Målestasjon))

        #print("Oljetype: " + str(oljetype))

        #print("Linje: " + str(Linje))

        #print("location: " + location)

        #print("new Filename: " + newfilename)

 

        
        try:

            os.rename(filename, newfilename)
            shutil.move(newfilename, location)
                                    
            break
        except:
            print("failed:\n " + "Old filename: " + str(filename) + "\n New Filename: " + str(newfilename) + "\n")
            
            break 




    
    
    
