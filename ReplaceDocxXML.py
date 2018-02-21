import os
import zipfile
import tempfile
import shutil
import sys

#Update Template.docx with generated XML and output to new report file
def update_docx(inFile, xmlFile, xmlLoc):
	tempdir = tempfile.mkdtemp()
	try:
		tempname = os.path.join(tempdir, 'temp.docx')
		templateDocx = zipfile.ZipFile(inFile)
		newDocx = zipfile.ZipFile(tempname, "w", zipfile.ZIP_DEFLATED, allowZip64=True)
		for file in templateDocx.filelist:
			if not file.filename == xmlLoc:
				newDocx.writestr(file.filename, templateDocx.read(file))	
		newDocx.write(xmlFile, xmlLoc)
		templateDocx.close()
		newDocx.close()
		shutil.move(tempname, inFile)
	finally:
		shutil.rmtree(tempdir)

if __name__ == '__main__':
    if len(sys.argv) != 4:
    	print("Usage: python3 ReplaceDocxXML.py {report file} {xml file} {custom/item#.xml}")
    	sys.exit()
    else:
    	update_docx(sys.argv[1], sys.argv[2], sys.argv[3])
