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
EXCEL_FILE_PATH = '차량목록.xlsx'
URL = 'https://gaos.glovis.net'

# [핵심 수정] 넥사크로의 긴 ID를 처리하는 방법
# ID가 정확히 일치하는 게 아니라, "특정 단어를 포함하는" 요소를 찾습니다.
# 예: input[id*='carNum'] -> ID 중간에 'carNum'이 들어가는 input 태그
INPUT_BOX_SELECTOR = "input[id*='carNum']"  

# 버튼도 마찬가지로 class나 ID의 일부분으로 찾습니다.
# 만약 버튼이 안 눌리면 F12에서 버튼의 텍스트(예: '조회')를 확인해주세요.
BUTTON_SELECTOR = ".search-btn" 

RESULT_TEXT_SELECTOR = '#repairHistory .date' 

# 엑셀 컬럼명
COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
        print(f"엑셀 로드 성공: {len(df)}개")
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
        # [Step 1] 사용자 수동 준비 (로그인 & 메뉴 이동)
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
        # [Step 2] 입력창 찾기 (iframe 없이 바로 찾기)
        # =======================================================
        print("🤖 입력창을 찾는 중...")

        try:
            # 1. ID에 'carNum'이 포함된 input 태그 찾기
            # (wait.until을 써서 넥사크로가 요소를 그릴 때까지 기다림)
            input_box = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, INPUT_BOX_SELECTOR)))
            print(f" -> 입력창 찾기 성공! (ID: {input_box.get_attribute('id')})")
            
        except:
            print("❌ 입력창을 못 찾았습니다.")
            print("팁: F12를 눌러 입력창 태그를 확인해보세요.")
            print("만약 <input id='...'> 가 아니라 <div id='...'> 라면, 클릭을 먼저 해야 input이 생기는 구조일 수 있습니다.")
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
            # 1. 입력창 다시 잡기 (넥사크로는 페이지 갱신 시 요소가 바뀔 수 있음)
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            
            # 2. 내용 지우고 입력
            input_box.clear()
            # 넥사크로 입력창은 click을 한번 해줘야 활성화되는 경우가 많음
            input_box.click() 
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
            print(f"[{car_num}] 실패: {e}")
            df.at[index, COL_REG_DATE] = "실패"

    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print("완료!")
    driver.quit()

if __name__ == "__main__":
    run_macro()