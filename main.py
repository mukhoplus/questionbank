import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import random
from excel_parser import parse_excel_file
from docx_generator import create_document, save_documents
from utils import handling_input_exception

IMAGES_OUTPUT_FOLDER = '.\images'
SHEET_NAME = '문제은행'
PADDIND_X = 40

def generate_documents():
  if not handling_input_exception(input_title.get(), input_number.get(), year.get()):
    return

  excel_path = f'.\하위권도우미_사업_기출문제_및_정답_{input_title.get()}.xlsx'  # 엑셀 파일 경로
  is_mid = True if exam_type.get() == '중간' else False

  try:
    question_bank = parse_excel_file(excel_path, SHEET_NAME, IMAGES_OUTPUT_FOLDER, is_mid)

    limit = len(question_bank)
    if limit == 0:
      messagebox.messagebox.showerror('오류', '저장된 문제가 없습니다.')
      return
    question_number = int(input_number.get()) if int(input_number.get()) <= limit else limit

    shuffled_list = random.sample(question_bank, limit)
    title_template = f'{year.get()}학년도 {semester.get()}학기 {exam_type.get()}고사 대비 모의고사 {input_title.get()} 과목 '

    doc_question, doc_empty, doc_answer = create_document(title_template, question_number, shuffled_list)
    save_documents(doc_question, doc_empty, doc_answer, input_title.get())
  except:
    return

window = ThemedTk(theme='breeze')
window.title('모의고사 생성기')
window.geometry('300x400')
window.minsize(300, 400)
window.maxsize(300, 400)

icon_image = Image.open(f'{IMAGES_OUTPUT_FOLDER}\khu_logo.png')
icon_photo = ImageTk.PhotoImage(icon_image)
window.iconphoto(False, icon_photo)

mukho_label = tk.Label(window, text=' Made By Mukho')
mukho_label.pack(padx=5, pady=5, side='top', anchor='w')

title_frame = ttk.Frame(window)
title_label = ttk.Label(title_frame, text='과목명')
input_title = ttk.Entry(title_frame)
title_label.pack(side='left')
input_title.pack(side='right')
title_frame.pack(padx=PADDIND_X, pady=10, fill='x')

number_frame = ttk.Frame(window)
number_label = ttk.Label(number_frame, text='문제 수')
input_number = ttk.Entry(number_frame)
number_label.pack(side='left')
input_number.pack(side='right')
number_frame.pack(padx=PADDIND_X, pady=10, fill='x')

year_frame = ttk.Frame(window,)
year_label = ttk.Label(year_frame, text='학년도')
year = ttk.Entry(year_frame)
year_label.pack(side='left')
year.pack(side='right')
year_frame.pack(padx=PADDIND_X, pady=10, fill='x')

semester_frame = ttk.LabelFrame(window, text='학기')
semester = tk.IntVar()
semester.set(1)
semester_radio_button1 = ttk.Radiobutton(semester_frame, text='1학기', variable=semester, value=1)
semester_radio_button2 = ttk.Radiobutton(semester_frame, text='2학기    ', variable=semester, value=2)
semester_radio_button1.pack(side='left', padx=5)
semester_radio_button2.pack(side='right', padx=5)
semester_frame.pack(padx=PADDIND_X, pady=10, fill='x')

exam_frame = ttk.LabelFrame(window, text='시험')
exam_type = tk.StringVar()
exam_type.set('중간')  # 기본값 설정
exam_radio_button1 = ttk.Radiobutton(exam_frame, text='중간고사', variable=exam_type, value='중간')
exam_radio_button2 = ttk.Radiobutton(exam_frame, text='기말고사', variable=exam_type, value='기말')
exam_radio_button1.pack(side='left', padx=5)
exam_radio_button2.pack(side='right', padx=5)
exam_frame.pack(padx=PADDIND_X, pady=10, fill='x')

generate_button = ttk.Button(window, width=25, text='문제지 생성', command=generate_documents)
generate_button.pack(padx=PADDIND_X, pady=20, fill='x')

window.mainloop()
