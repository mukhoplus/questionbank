import openpyxl, os
from tkinter import messagebox

def parse_excel_file(excel_path, sheet_name, images_output_folder, is_mid):
  try:
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    worksheet = workbook[sheet_name]
    
    if not worksheet:
      messagebox.showerror('오류', '"문제은행" 시트를 찾을 수 없습니다.')
      return

    # 이미지가 저장될 폴더가 없으면 생성
    if not os.path.exists(images_output_folder):
      os.makedirs(images_output_folder)

    question_bank = []
    for row_index, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):  # 행 번호를 시작 인덱스로 사용
      if not row[5]:
        break
      
      if is_mid and not row[1].startswith('중간'):
        continue

      question = {
        '검수용 번호': row[0],
        '시험': row[1],
        '유형': row[2], # '중간고사', '기말고사', '중간고사, 기말고사' 3가지의 입력이 있다고 합의
        '교수님': row[3],
        '<출제 년도>': row[4],
        '<문제>': row[5],
        '<답가지>': row[6] if row[6] else '',  # None 값 처리
        '<제시그림>': row[7], # '.\images\image_검수용 번호' (과목_검수용 번호 검토 중)
        '<정답>': row[8],
        '<족보 페이지 또는 해설>': row[9] if row[9] else '',
      }
      question_bank.append(question)

    return question_bank
  except FileNotFoundError:
    messagebox.showerror('오류', '엑셀 파일을 찾을 수 없습니다.')
    return
  except Exception:
    messagebox.showerror('오류', '엑셀 파일을 여는 중 오류가 발생했습니다.')
    return