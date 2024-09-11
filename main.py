import random
from excel_parser import parse_excel_file
from docx_generator import create_document, save_documents

images_output_folder = '.\images'
input_title = '심계내과학2' # input('(테스트용)파일명을 입력하세요: ')
input_number = 5 # input('(테스트용)문제 수를 입력하세요: ')

excel_path = f'.\{input_title}.xlsx'  # 엑셀 파일 경로
sheet_name = '문제은행' # 시트 이름

question_bank = parse_excel_file(excel_path, sheet_name, images_output_folder)

limit = len(question_bank)
if limit == 0:
    raise ValueError('저장된 문제가 없습니다.')
if input_number > limit:
    input_number = limit

shuffled_list = random.sample(question_bank, limit)

year = 2024 # int(input('년도를 입력하세요: '))
semester = 2 # input('학기를 입력하세요: ')
exam_type = '중간' # input('시험 종류를 입력하세요(중간/기말): ')

title_template = f'{year}학년도 {semester}학기 {exam_type}고사 대비 모의고사 {input_title} 과목 '

doc_question, doc_empty, doc_answer = create_document(title_template, input_number, shuffled_list)
save_documents(doc_question, doc_empty, doc_answer, input_title, images_output_folder)
