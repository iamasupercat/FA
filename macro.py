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

# 메인 페이지 진입 후 설정
SYSTEM_BTN_SELECTOR = "img[src*='btn_TF_MenuS.png']" # 시스템 버튼
MENU_BOX_SELECTOR = ".nexaedge"                      # 메뉴 박스

# 조회 화면 설정
INPUT_BOX_SELECTOR = '#carNumInput'           # 차량번호 입력창
BUTTON_SELECTOR = '.search-btn'               # 검색 버튼
RESULT_TEXT_SELECTOR = '#repairHistory .date' # 결과 텍스트

# 엑셀 컬럼명
COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    # 1. 엑셀 로드
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
        print(f"엑셀 로드 성공: {len(df)}개")
    except Exception as e:
        print(f"엑셀 파일 오류: {e}")
        return

    # 2. 브라우저 열기 (최대화)
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized') 
    
    # [중요] 자동화 탐지 피하기 옵션 (보안 프로그램이 조금 덜 민감하게 반응함)
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)

    try:
        # =======================================================
        # [Step 1] 수동 로그인 대기 (여기가 핵심!)
        # =======================================================
        print("\n" + "="*50)
        print("🚨 [사용자 개입 필요] 🚨")
        print("1. 열린 브라우저에서 직접 로그인을 완료해주세요.")
        print("2. 로그인이 끝나고 '메인 화면'이 보이면...")
        input("👉 이 검은색 창(터미널)을 클릭하고 [Enter] 키를 누르세요! (엔터 누르면 시작됨)")
        print("="*50 + "\n")
        
        # =======================================================
        # [Step 2] 여기서부터 로봇이 이어받음
        # =======================================================
        print("🤖 매크로 작동 시작! 시스템 메뉴를 찾습니다...")

        # 1. 시스템 버튼 클릭
        system_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SYSTEM_BTN_SELECTOR)))
        system_btn.click()
        print(" -> 시스템 버튼 클릭")

        # 2. 메뉴 박스 대기 및 클릭
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, MENU_BOX_SELECTOR)))
        
        # 메뉴 클릭
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '원부카히스토리')]"))).click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '카히스토리관리')]"))).click()
        print(" -> 메뉴 이동 완료")

        # 입력 화면 대기
        time.sleep(3)

        # iframe 체크 (혹시 입력창이 iframe에 있는지)
        if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) == 0:
            print(" -> 메인에 입력창 없음. iframe 탐색...")
            for frame in driver.find_elements(By.TAG_NAME, 'iframe'):
                driver.switch_to.default_content()
                try:
                    driver.switch_to.frame(frame)
                    if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) > 0:
                        print(" -> iframe 진입 성공")
                        break
                except: pass

    except Exception as e:
        print(f"초기 설정 실패: {e}")
        driver.quit()
        return

    # [Step 3] 반복 조회 시작
    print("--- 데이터 조회 시작 ---")
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.clear()
            input_box.send_keys(str(car_num))

            confirm_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            driver.execute_script("arguments[0].click();", confirm_btn)

            result_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)))
            extracted_text = result_element.text
            
            df.at[index, COL_REG_DATE] = extracted_text
            print(f"[{car_num}] : {extracted_text}")
            time.sleep(0.5) 

        except Exception as e:
            print(f"[{car_num}] 실패")
            df.at[index, COL_REG_DATE] = "실패"

    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print("작업 끝!")
    driver.quit()

if __name__ == "__main__":
    run_macro()