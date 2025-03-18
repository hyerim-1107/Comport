import os
import pandas as pd

def find_file(filename, search_path):
    """
    주어진 폴더와 그 하위 폴더 전체에서 filename을 검색하여,
    처음 발견한 파일의 전체 경로를 반환합니다.
    """
    for root, dirs, files in os.walk(search_path):
        if filename in files:
            return os.path.join(root, filename)
    return None

# 엑셀 파일 입력
input_file = input("엑셀 파일 이름을 적어주세요 (Ex. ABC.xlsx) : ")
output_file = "cleaned_data.xlsx"

# 현재 폴더에 파일이 있는지 확인
if not os.path.isfile(input_file):
    print(f"'{input_file}' 파일을 현재 폴더에서 찾을 수 없습니다.")
    decision = input("전체 경로 검색을 진행하시겠습니까? (y/n): ")
    if decision.lower() == 'y':
        # 현재 작업 디렉토리부터 하위 모든 폴더를 검색
        found_file = find_file(input_file, os.getcwd())
        if found_file:
            print(f"파일을 '{found_file}' 경로에서 찾았습니다.")
            input_file = found_file
        else:
            print("전체 경로 검색 결과 파일을 찾을 수 없습니다.")
            exit()
    else:
        exit()

# 파일 읽기
try:
    df = pd.read_excel(input_file)
except Exception as e:
    print(f"파일을 읽는 중 오류가 발생했습니다: {e}")
    exit()

# "이름"과 "전화번호" 키워드가 포함된 컬럼 찾기
name_cols = [col for col in df.columns if "성함" in col]
phone_cols = [col for col in df.columns if "전화번호" in col]

# 키워드가 포함된 컬럼 존재 여부 확인
if not name_cols:
    print("파일에 '성함' 키워드가 포함된 컬럼이 없습니다.")
    exit()

if not phone_cols:
    print("파일에 '전화번호' 키워드가 포함된 컬럼이 없습니다.")
    exit()

# 결측값이 있는 행 제거 (찾은 모든 이름 및 전화번호 관련 컬럼 사용)
df_cleaned = df.dropna(subset=name_cols + phone_cols)

# 중복 제거: 찾은 첫 번째 '전화번호' 관련 컬럼을 기준으로 마지막 데이터만 유지
df_cleaned = df_cleaned.drop_duplicates(subset=[phone_cols[0]], keep="last")

# 결과를 파일로 저장
try:
    df_cleaned.to_excel(output_file, index=False)
    print(f"중복이 정리된 파일이 '{output_file}'로 저장되었습니다!")
except Exception as e:
    print(f"파일 저장 중 오류가 발생했습니다: {e}")
