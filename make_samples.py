import pandas as pd
import numpy as np
import os

# 1. 파일들이 저장될 폴더 만들기
folder_name = "sample_excels"
os.makedirs(folder_name, exist_ok=True)
print(f"📂 '{folder_name}' 폴더를 생성했습니다.")

# ---------------------------------------------------------
# 데이터 생성 도우미 함수
# ---------------------------------------------------------
def create_dummy_df(start_line, count, start_date):
    """
    start_line부터 시작하는 가짜 데이터를 만듭니다.
    컬럼: Line(필수), Date(날짜), Value_A(랜덤값), Value_B(랜덤값)
    """
    lines = range(start_line, start_line + count)
    dates = pd.date_range(start=start_date, periods=count, freq='D') # 하루씩 증가

    data = {
        "Line": lines,
        "Date": dates,
        "Temperature": np.random.randint(20, 35, size=count), # 20~35도 랜덤
        "Pressure": np.random.rand(count) * 100 # 0~100 압력 랜덤
    }
    return pd.DataFrame(data)

# ---------------------------------------------------------
# 파일 1: 아주 정상적인 데이터 (Line 1 ~ 20)
# ---------------------------------------------------------
df1 = create_dummy_df(start_line=1, count=20, start_date="2024-01-01")
file1_path = os.path.join(folder_name, "test_data_1.xlsx")
df1.to_excel(file1_path, index=False)
print(f"✅ 생성 완료: {file1_path} (Line 1~20)")

# ---------------------------------------------------------
# 파일 2: 여러 시트 + 일부 겹치는 데이터 (Line 15 ~ 35)
# ---------------------------------------------------------
# Sheet1: 정상 데이터 (Line 15~35, 앞 파일과 15~20이 겹침 -> 중복 처리 테스트용)
df2_sheet1 = create_dummy_df(start_line=15, count=21, start_date="2024-01-15")

# Sheet2: 'Line' 컬럼이 없는 엉뚱한 데이터 (제외 처리 테스트용)
df2_sheet2 = pd.DataFrame({
    "Memo": ["이 시트는", "Line 컬럼이", "없어서", "무시되어야 합니다."],
    "Author": ["Kim", "Lee", "Park", "Choi"]
})

file2_path = os.path.join(folder_name, "test_data_2.xlsx")
with pd.ExcelWriter(file2_path) as writer:
    df2_sheet1.to_excel(writer, sheet_name="Valid_Data", index=False)
    df2_sheet2.to_excel(writer, sheet_name="Invalid_Sheet", index=False)
print(f"✅ 생성 완료: {file2_path} (Sheet1: Line 15~35, Sheet2: 무효 데이터)")

# ---------------------------------------------------------
# 파일 3: 날짜가 비어있는 데이터 (Line 36 ~ 50)
# ---------------------------------------------------------
df3 = create_dummy_df(start_line=36, count=15, start_date="2024-02-05")

# 일부러 날짜를 지워서 결측치(NaN) 만들기 -> ffill/bfill 테스트용
df3.loc[2:5, "Date"] = np.nan  # 2~5번째 행 날짜 삭제
df3.loc[10, "Date"] = np.nan   # 10번째 행 날짜 삭제

file3_path = os.path.join(folder_name, "test_data_3_missing_date.xlsx")
df3.to_excel(file3_path, index=False)
print(f"✅ 생성 완료: {file3_path} (Line 36~50, 날짜 일부 누락)")

print("\n🎉 모든 테스트 파일 준비 완료! 메인 프로그램을 실행해서 이 파일들을 선택해보세요.")
