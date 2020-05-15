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

i = ''
j = 0
with open("A.txt", 'r') as data_dict:
    file_text = data_dict.read().splitlines()
    print(file_text)
    for file in os.listdir("./1")[0:]:
        try:
            print(file + " opened")
            pdfText = convert_pdf_to_txt("./1/" + file)
            pdfText = pdfText.split('.')
            filename = file.replace('.pdf', '')
            r_count: int = 0
            while r_count < pdfText.__len__():
                pdfText[r_count] = pdfText[r_count].replace('\n', ' ')
                v = 0
                while v < file_text.__len__():
                    line = file_text[v].split('|')
                    for i in line:
                        test = True
                        if i == line[-1]:
                            break
                        i = i.split(',')
                        while j < i.__len__():
                            try:
                                pdfText[r_count].upper().index(i[j].upper())
                                f = pdfText[r_count]
                                j += 1
                                test = False
                                with open("./Community_" + filename + ".txt", 'a+') as out_file:
                                    # print("Writing to " + word + ".txt has begun")
                                    out_file.write('\n\n')
                                    out_file.write(
                                        '--------------------------------------------------------------------------'
                                        '-------------------------------------------------- '
                                        + line[-1] +
                                        '--------------------------------------------------------------------------'
                                        '-------------------------------------------------- \n')
                                    # print((pdfText[r_count].encode('ascii', 'ignore')).decode('utf-8') + '\n')
                                    out_file.write(
                                        (pdfText[r_count].encode('ascii', 'ignore')).decode('utf-8') + '\n')
                                    # print("Writing to " + word + ".txt has executed")
                                    out_file.write('\n\n')
                                break
                            except ValueError:
                                j += 1
                                pass
                                # print("f")
                        if test:
                            break
                        # except ValueError:
                        #     continue
                        #     # print("g")
                    v += 1
                r_count = r_count + 1
        except ValueError:
            print(file + "failed")
            continue
