#bit of code to try and save corrupted docx files

import docx2txt #package to convert docx files to txt files
import glob #package to move through files
import os #packaged to get directory

#path to folder containing documents 
main_path = os.path.dirname(os.path.realpath(__file__))

#list containing all the file names to be extracted
all_files = glob.glob(main_path + "/input_docx" + "/*.docx")

for filename in all_files:

    OUTPUT_TXT = docx2txt.process("BLViva.docx") #parsed text

    new_filename = filename.replace('.docx','.txt') #filename for output

    new_filename = new_filename.replace('input_docx','output_txt')
    
    file = open(new_filename, "w") #open file and overwrite
    file.write(OUTPUT_TXT) #print txt
    file.close() #close file