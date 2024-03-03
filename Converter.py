from io import BytesIO
import io


import pandas as pd
from docx import Document
from docx.shared import RGBColor


from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)


class Txt_to_xlsx:

    def __init__(self, id):
        self.arr = []

        self.time = "900"

        self.count = 0
        self.separator = "|"
        self.store = f"Question Text{self.separator}Question Type{self.separator}Option 1{self.separator}Option 2{self.separator}Option 3{
            self.separator}Option 4{self.separator}Correct Answer{self.separator}Time in seconds{self.separator}Image Link\n"
        self.file1 = drive.CreateFile({'title': 'test.txt'})

        self.file_id = id
        self.file_obj = drive.CreateFile({'id': self.file_id})
        self.file_obj.FetchContent()
        self.content = io.BytesIO(self.file_obj.content.getvalue())

        self.read_docx()
        self.reformattingfile()

        self.correctanswer()
        self.stringbuilding()

        self.structurereplace()
        self.csv_to_excel()

    def read_docx(self):

        doc = Document(self.content)
        text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
        self.file1.SetContentString(text)
        self.file1.Upload()

    def correctanswer(self):
        for para in Document(self.content).paragraphs:
            for run in para.runs:
                if (run.font.color.rgb == RGBColor(255, 000, 000)) or (run.underline) or (run.bold):
                    if run.text.startswith("A."):
                        temp = "1"
                        self.arr.append(temp)

                    elif run.text.startswith("B."):
                        temp = "2"
                        self.arr.append(temp)

                    elif run.text.startswith("C."):
                        temp = "3"
                        self.arr.append(temp)

                    elif run.text.startswith("D."):
                        temp = "4"
                        self.arr.append(temp)

    def reformattingfile(self):
        temp = self.file1.GetContentString()

        for word in ["A.", "B.", "C.", "D."]:
            temp = temp.replace(word, f"\n{word}")

        self.file1.SetContentString(temp)
        self.file1.Upload()

    def stringbuilding(self):

        flag = False
        content = self.file1.GetContentString()
        for line in content.splitlines():

            if not line.startswith("A.") and flag:
                self.store += line
            else:
                if line.startswith("Câu"):
                    self.store += line
                    flag = True

                elif line.startswith("A."):
                    flag = False
                    self.store += (self.separator +
                                   "Multiple Choice" + self.separator + line)

                elif line.startswith("B.") or line.startswith("C."):
                    self.store += self.separator + line

                elif line.startswith("D."):
                    # self.store += self.separator + line + self.separator + self.time + "\n"
                    self.store += self.separator + line + self.separator + self.arr[
                        self.count] + self.separator + self.time + "\n"
                    self.count += 1
        # else:
        #     self.store += line

    def structurereplace(self):

        counter = 1
        while counter <= self.count:
            ques = f"Câu {counter}: "
            self.store = self.store.replace(ques, '')
            counter += 1

        counter = 1
        while counter <= self.count:
            ques = f"Câu {counter}. "
            self.store = self.store.replace(ques, '')

            counter += 1

        for char in ['A.', 'B.', 'C.', 'D.']:
            self.store = self.store.replace(char, '')

        self.file1.SetContentString(self.store)
        self.file1['title'] = 'test.csv'
        self.file1.Upload()
        print("helo")

    def csv_to_excel(self):
        # print("hellooo")
        # self.store.strip()
        # self.store_bytes = self.store.encode('utf-8')
        # parsed_data = [line.decode('utf-8').split('|')
        #                for line in self.store_bytes.strip().split(b'\n')]

        # # Create DataFrame from parsed data
        # df = pd.DataFrame(parsed_data, columns=['Question Text', 'Question Type', 'Option 1',
        #                                         'Option 2', 'Option3', 'Option 4', 'Correct Answer', 'Time in seconds', 'Image link'])

        # # Create a BytesIO object to hold Excel file in memory
        # excel_buffer = BytesIO()

        # # Write DataFrame to Excel file in memory
        # df.to_excel(excel_buffer, index=False)

        # # Upload Excel file to Google Drive
        # excel_buffer.seek(0)
        excel_file = drive.CreateFile({'title': 'converted_excel_file',
                                      'mimeType': 'application/vnd.google-apps.spreadsheet'})

        excel_file.SetContentString(self.store)
        excel_file.Upload()
        print("hello world")

        # Print the link to the uploaded file
        print("Uploaded Excel file link:", excel_file['alternateLink'])


id = '1BLPYIDxza864itrQ2FH7z_Ai9wyKsZMP'
c1 = Txt_to_xlsx(id)

# python3 converter.py


# file2 = drive.CreateFile({'title': self.excel})
#         csv = pd.read_csv(self.file1, sep=self.separator, on_bad_lines='skip')
#         csv.to_excel(self.excel, index=False)

#         file2.SetContentFile(self.file1)
#         file2.Upload()
