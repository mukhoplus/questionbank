import pandas as pd
from docx import Document
from docx.shared import Inches

temp_input = '테스트과목' # input('(테스트용)파일명을 입력하세요: ')

# 엑셀 파일에서 데이터 불러오기
excel_path = f'./{temp_input}.xlsx'  # 엑셀 파일 경로
sheet_name = '문제은행'  # 시트 이름

# Word 문서 생성
doc_question = Document()
doc_question.add_heading('Excel Data to Word 1', level=1)

doc_answer = Document()
doc_answer.add_heading('Excel Data to Word 2', level=1)

question_bank = []

# 엑셀 파일 읽기 (첫 번째 행을 헤더로 간주)
try:
  df = pd.read_excel(excel_path, sheet_name=sheet_name)

  # 행마다 데이터를 파싱하여 워드 문서에 추가
  for index, row in df.iterrows():
    q_number = row['번호']
    q_type = row['유형']
    q_exam = row['시험']
    q_professor = row['교수님']
    q_year = row['<출제 년도>']
    q_question = row['<문제>']
    q_select = row['<답가지>']
    q_image = row['<제시그림>']
    q_description = row['<족보 페이지 또는 해설>']
    q_inspection = row['검수용\n번호']

    question = {
      '번호': q_number,
      '유형': q_type,
      '시험': q_exam,
      '교수님': q_professor,
      '<출제 년도>': q_year,
      '<문제>': q_question,
      '<답가지>': q_select if not pd.isna(q_select) else '',
      '<제시그림>': q_image if not pd.isna(q_image) else '',
      '<족보 페이지 또는 해설>': q_description if not pd.isna(q_description)else '',
      '검수용 번호': q_inspection if not pd.isna(q_inspection) else ''
    }

    question_bank.append(question)
except FileNotFoundError:
  print('해당 파일이 없습니다.')

for question in question_bank:
  print(question)

'''
# 문제 유형 추가
doc_question.add_heading(f"문제 유형: {row['문제유형']}", level=2)

# 문제 추가
doc_question.add_paragraph(f"문제: {row['문제']}")

# 이미지 추가 (이미지가 있을 경우)
if pd.notna(row['이미지']):
  try:
    doc_question.add_picture(row['이미지'], width=Inches(2.0))  # 이미지 경로에서 이미지를 추가
  except Exception as e:
    doc_question.add_paragraph(f"이미지 로드 실패: {e}")

# 객관식 선지 추가 (선지가 존재할 경우)
for i in range(1, 6):
  option_col = f'선지{i}'
  if option_col in df.columns and pd.notna(row[option_col]):
    doc_question.add_paragraph(f"선지{i}: {row[option_col]}")

# 정답 추가
doc_question.add_paragraph(f"정답: {row['정답']}")

# 해설 추가
doc_question.add_paragraph(f"해설: {row['해설']}")

# 행 구분을 위한 빈 줄 추가
doc_question.add_paragraph("\n")

# 워드 파일로 저장
question_path = f'[{temp_input}] 문제지.docx'  # 워드 파일 경로
doc_question.save(question_path)
answer_path = f'[{temp_input}] 정답지.docx'  # 워드 파일 경로
doc_answer.save(answer_path)

print(f"워드 파일이 {question_path}에 저장되었습니다.")
print(f"워드 파일이 {answer_path}에 저장되었습니다.")
'''