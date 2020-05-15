from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import os


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=False):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

word_list = ''
word = 0
with open("A.txt", 'r') as data_dict:
    file_text = data_dict.read().splitlines()
    print(file_text)
    for file in os.listdir("./1")[0:]:
        try:
            print(file + " opened")
            pdfText = convert_pdf_to_txt("./1/" + file)
            pdfText = pdfText.split('.')
            file_name = file.replace('.pdf', '')
            line_index: int = 0
            while line_index < pdfText.__len__():
                pdfText[line_index] = pdfText[line_index].replace('\n', ' ')
                v = 0
                while v < file_text.__len__():
                    line = file_text[v].split('|')
                    for word_list in line:
                        test = True
                        if word_list == line[-1]:
                            break
                        word_list = word_list.split(',')
                        for word in word_list:
                            try:
                                pdfText[line_index].upper().index(word.upper())
                                string_test = pdfText[line_index]
                                test = False
                                try:
                                    line[-2].index(word)
                                    with open("./Community_" + file_name + ".txt", 'a+') as out_file:
                                        # print("Writing to " + word + ".txt has begun")
                                        out_file.write('\n\n')
                                        out_file.write(
                                            '--------------------------------------------------------------------------'
                                            '-------------------------------------------------- '
                                            + line[-1] +
                                            '--------------------------------------------------------------------------'
                                            '-------------------------------------------------- \n')
                                        # print((pdfText[line_index].encode('ascii', 'ignore')).decode('utf-8') + '\n')
                                        out_file.write(
                                            (pdfText[line_index].encode('ascii', 'ignore')).decode('utf-8') + '\n')
                                        # print("Writing to " + word + ".txt has executed")
                                        out_file.write('\n\n')
                                    break
                                except ValueError:
                                    pass
                            except ValueError:
                                pass
                                # print("string_test")
                        if test:
                            break
                        # except ValueError:
                        #     continue
                        #     # print("g")
                    v += 1
                line_index = line_index + 1
        except ValueError:
            print(file + "failed")
            continue
