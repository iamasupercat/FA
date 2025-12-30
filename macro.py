import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# ==========================================
# 1. 사용자 설정 영역
# ==========================================
EXCEL_FILE_PATH = '차량목록.xlsx' # (파일 이름이 다르다면 꼭 수정해주세요!)
URL = 'https://gaos.glovis.net'

# [수정 1] 아까 화면에서 찾아낸 ID의 핵심 단어 'CARNO' 적용
# 대소문자가 중요하므로 화면에 보였던 대로 대문자로 적었습니다.
INPUT_BOX_SELECTOR = "input[id*='CARNO']"  

# 조회 버튼 (만약 안 눌리면 버튼의 텍스트나 class 확인 필요)
BUTTON_SELECTOR = ".search-btn" 

# 결과 텍스트 (조회 후 나오는 날짜 등)
RESULT_TEXT_SELECTOR = '#repairHistory .date' 

# 엑셀 컬럼명
COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    try:
        # [수정 2] header=1 추가 (엑셀의 2번째 줄을 제목으로 읽으라는 뜻)
        df = pd.read_excel(EXCEL_FILE_PATH, header=1)
        print(f"엑셀 로드 성공: {len(df)}개 행을 읽었습니다.")
        print(f"읽어온 컬럼 목록: {df.columns.tolist()}") # 확인용 출력
    except Exception as e:
        print(f"엑셀 파일 오류: {e}")
        return

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized') 
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)

    try:
        # =======================================================
        # [Step 1] 사용자 수동 준비
        # =======================================================
        print("\n" + "="*60)
        print("🚨 [사용자 준비 단계] 🚨")
        print("1. 브라우저에서 직접 [로그인]을 해주세요.")
        print("2. [원부카히스토리] -> [카히스토리관리] 메뉴까지 직접 이동해주세요.")
        print("3. 화면에 '차량번호 입력창'이 보이면 준비 끝!")
        print("-" * 60)
        input("👉 준비되셨으면 엔터(Enter)를 누르세요!")
        print("="*60 + "\n")
        
        # =======================================================
        # [Step 2] 입력창 찾기
        # =======================================================
        print("🤖 입력창을 찾는 중...")

        try:
            # 설정한 CARNO가 포함된 입력창 찾기
            input_box = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, INPUT_BOX_SELECTOR)))
            print(f" -> 입력창 찾기 성공! (ID: {input_box.get_attribute('id')})")
            
        except:
            print("❌ 입력창을 못 찾았습니다.")
            print("팁: F12를 눌러 ID에 'CARNO'가 포함되어 있는지 다시 확인해주세요.")
            return

    except Exception as e:
        print(f"설정 오류: {e}")
        driver.quit()
        return

    # =======================================================
    # [Step 3] 데이터 반복 조회
    # =======================================================
    print("--- 데이터 조회 시작 ---")
    
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            # 1. 입력창 잡기
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            
            # 2. 내용 지우고 입력
            input_box.clear()
            input_box.click() # 넥사크로 활성화 클릭
            input_box.send_keys(str(car_num))

            # 3. 조회 버튼 클릭
            confirm_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            driver.execute_script("arguments[0].click();", confirm_btn)

            # 4. 결과 대기 및 추출
            result_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)))
            extracted_text = result_element.text
            
            df.at[index, COL_REG_DATE] = extracted_text
            print(f"[{car_num}] : {extracted_text}")
            time.sleep(0.5) 

        except Exception as e:
            print(f"[{car_num}] 실패 혹은 데이터 없음: {e}")
            df.at[index, COL_REG_DATE] = "조회실패"

    # 결과 저장 (파일명 앞에 '결과포함_' 붙여서 저장)
    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 작업 완료! '{save_name}' 파일을 확인하세요.")
    driver.quit()

if __name__ == "__main__":
    run_macro()