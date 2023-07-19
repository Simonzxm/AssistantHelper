# a py program for Fat Xu :)

import pandas as pd
import win32com.client as win32
import os

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
            doc.Application.Selection.WholeStory()
            doc.Application.Selection.MoveRight()
            doc.Application.Selection.TypeText('\n')
            for j in student_mistakes[student_index]:
                doc.Application.Selection.InsertFile(j)
            doc.Save()
            doc.Close()
        except Exception as e:
            print(f'Error:{i}\n{e}')

    student_index += 1
word.Quit()

print('Done!')
input('press enter to exit.')
