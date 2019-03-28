import docx2txt
test = docx2txt.process("test.docx")
with open("test.txt", "w") as textfile:
    textfile.write(test)
    textfile.close()
