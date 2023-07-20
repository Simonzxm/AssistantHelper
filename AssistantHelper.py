# a py program for Fat Xu :)

import pandas as pd
import win32com.client as win32
import os


def process_document(name, document, mistakes):
    try:
        document.Application.Selection.WholeStory()
        document.Application.Selection.MoveRight()
        document.Application.Selection.TypeText('\n')
        for j in mistakes:
            document.Application.Selection.InsertFile(j)
    except Exception as exception:
        print(f'Error:{name}\n{exception}')


# read data.xlsx
sheet = pd.read_excel('data.xlsx')
student_data = sheet.to_dict(orient='records')
student_name = []
student_mistake = []
cwd = os.getcwd()
for i in student_data:
    student_name.append(list(i.values())[0])
    student_mistake.append(list(i.values())[2])
student_mistakes = []
for i in student_mistake:
    student_mistakes.append([os.path.join(cwd, j + '.docx') for j in list(str(i).split('，'))])
    # 错题之间分隔符为中文逗号 可以自己在上面改成其他的

# edit documents
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
student_index = 0
for i in student_name:
    if student_mistakes[student_index] != [os.path.join(cwd, '全对.docx')]:
        try:
            doc = word.Documents.Open(os.path.join(cwd, i + '.docx'))
            process_document(i, doc, student_mistakes[student_index])
            doc.Save()
            doc.Close()
        except Exception as e:
            try:
                doc = word.Documents.Add()
                process_document(i, doc, student_mistakes[student_index])
                doc.SaveAs(os.path.join(cwd, i + '.docx'))
                doc.Close()
                print(f'New File Created: {student_name}.docx')
            except Exception as f:
                print(f'Error:{i}\n{f}')
    student_index += 1
word.Quit()

print('Done!')
input('press enter to exit.')
