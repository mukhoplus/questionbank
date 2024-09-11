import pandas as pd

def parse_excel_file(excel_path, sheet_name):
  try:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    question_bank = []

    for index, row in df.iterrows():
      question = {
        '번호': row['번호'],
        '유형': row['유형'],
        '시험': row['시험'],
        '교수님': row['교수님'],
        '<출제 년도>': row['<출제 년도>'],
        '<문제>': row['<문제>'],
        '<답가지>': row['<답가지>'] if not pd.isna(row['<답가지>']) else '',
        '<제시그림>': row['<제시그림>'] if not pd.isna(row['<제시그림>']) else '',
        '<정답>': row['<정답>'],
        '<족보 페이지 또는 해설>': row['<족보 페이지 또는 해설>'] if not pd.isna(row['<족보 페이지 또는 해설>']) else '',
        '검수용 번호': row['검수용\n번호'] if not pd.isna(row['검수용\n번호']) else ''
      }
      question_bank.append(question)

    return question_bank
  except FileNotFoundError:
      raise FileNotFoundError('해당 파일이 없습니다.')