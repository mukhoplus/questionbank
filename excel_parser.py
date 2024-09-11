import openpyxl
import os

def parse_excel_file(excel_path, sheet_name, images_output_folder):
  try:
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    worksheet = workbook[sheet_name]

    # 이미지가 저장될 폴더가 없으면 생성
    if not os.path.exists(images_output_folder):
      os.makedirs(images_output_folder)

    question_bank = []
    for row_index, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):  # 행 번호를 시작 인덱스로 사용
      if not row[5]:
        break

      question = {
        '검수용 번호': row[0],
        '시험': row[1],
        '유형': row[2],
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
    raise FileNotFoundError('해당 파일이 없습니다.')