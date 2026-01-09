import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains # [추가] 마우스 이동용
import time
import random

# ==========================================
# 1. 사용자 설정 영역
# ==========================================
EXCEL_FILE_PATH = '카히스토리관리_20251230112324.xlsx' 
URL = 'https://gaos.glovis.net'

# 입력창
INPUT_BOX_SELECTOR = "input[id*='CARNO']"

# [핵심] 조회 버튼 (찾아내신 ID)
BUTTON_SELECTOR = "div[id*='searchBtn']"

# 결과 텍스트
RESULT_TEXT_SELECTOR = "div[id*='Grid01'][id*='_5:text']"

COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, header=1)
        print(f"✅ 엑셀 로드 성공: {len(df)}개")
    except Exception as e:
        print(f"❌ 엑셀 파일 오류: {e}")
        return

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized') 
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)
    action = ActionChains(driver) # 마우스 조작 도구 준비

    try:
        # =======================================================
        # [Step 1] 사용자 수동 준비
        # =======================================================
        print("\n" + "="*60)
        print("🚨 [사용자 준비 단계] 🚨")
        print("1. 로그인 후 [카히스토리관리] 메뉴로 이동하세요.")
        print("2. 입력창이 보이면 터미널 클릭 후 Enter!")
        print("-" * 60)
        input("👉 준비되셨으면 엔터(Enter)를 누르세요!")
        print("="*60 + "\n")
        
        # 입력창 찾기 확인
        try:
            input_box = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, INPUT_BOX_SELECTOR)))
            print(f" -> 입력창 찾기 성공!")
        except:
            print("❌ 입력창을 못 찾았습니다.")
            return

    except Exception as e:
        print(f"설정 오류: {e}")
        driver.quit()
        return

    # =======================================================
    # [Step 2] 데이터 반복 조회
    # =======================================================
    print("--- 데이터 조회 시작 ---")
    
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            # 1. 입력창에 값 넣기
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.click()
            input_box.clear()
            input_box.send_keys(str(car_num))
            time.sleep(random.uniform(0.1, 0.5))  # 랜덤 딜레이: 0.1~0.5초

            # [중요] 입력 확정을 위해 빈 공간(body) 한번 클릭
            # (커서가 입력창에 남아있으면 조회가 안 되는 경우가 있음)
            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(random.uniform(0.2, 0.8))  # 랜덤 딜레이: 0.2~0.8초

            # 2. 버튼 찾기 및 클릭 시도 (JS 강제 클릭)
            try:
                # 2-1. 버튼 요소 찾기
                search_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
                
                # 마우스 이동이고 뭐고 다 생략하고, 브라우저 엔진에 직접 클릭 명령 1회 전송
                driver.execute_script("arguments[0].click();", search_btn)
                print(" -> JS 강제 클릭 완료")
                
            except Exception as e:
                # 실패 시 예외 처리
                print(" -> Selector 실패, 텍스트로 조회 시도")
                xpath_btn = driver.find_element(By.XPATH, "//*[contains(text(), '조회')]")
                driver.execute_script("arguments[0].click();", xpath_btn)
                
            # 3. 결과 수집
            time.sleep(random.uniform(1.5, 3.0))  # 랜덤 딜레이: 1.5~3.0초 (조회 로딩 대기)
            
            results = driver.find_elements(By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)
            
            if len(results) > 0:
                # 텍스트 가져오기
                text_list = [r.text for r in results if r.text.strip() != ""]
                full_text = "\n".join(text_list)
                
                df.at[index, COL_REG_DATE] = full_text
                print(f"[{car_num}] 성공! ({len(text_list)}행)")
            else:
                print(f"[{car_num}] 조회 결과 없음 (혹은 버튼 안 눌림)")
                df.at[index, COL_REG_DATE] = "내역없음"
            
            # 다음 조회 전 랜덤 딜레이 (매크로 탐지 방지)
            time.sleep(random.uniform(0.5, 1.5))

        except Exception as e:
            print(f"[{car_num}] 에러: {e}")
            df.at[index, COL_REG_DATE] = "에러"

    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 끝! '{save_name}' 저장 완료.")
    driver.quit()

if __name__ == "__main__":
    run_macro()