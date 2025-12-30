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

# 조회 화면 설정 (F12로 찾은 값들)
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
    
    # 자동화 탐지 회피 옵션
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)

    try:
        # =======================================================
        # [Step 1] 완전 수동 준비 구간 (로그인 + 메뉴 이동)
        # =======================================================
        print("\n" + "="*60)
        print("🚨 [사용자 준비 단계] 🚨")
        print("1. 브라우저에서 직접 [로그인]을 해주세요.")
        print("2. [시스템] -> [원부카히스토리] -> [카히스토리관리] 메뉴로 이동해주세요.")
        print("3. '차량번호 입력창'이 화면에 보이면 준비 끝!")
        print("-" * 60)
        input("👉 준비가 다 되셨으면, 여기(터미널)를 클릭하고 [Enter] 키를 누르세요!")
        print("="*60 + "\n")
        
        # =======================================================
        # [Step 2] 입력창 위치 찾기 (자동 감지)
        # =======================================================
        print("🤖 매크로 작동 시작! 입력창을 찾는 중입니다...")

        # 입력창이 바로 보이는지, 아니면 iframe 안에 숨어있는지 확인
        found_input = False
        
        # 1. 메인 화면에서 바로 찾기 시도
        if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) > 0:
            print(" -> 메인 화면에서 입력창 발견!")
            found_input = True
        else:
            # 2. 메인에 없으면 iframe 뒤지기
            print(" -> 메인에 입력창 없음. iframe 내부 탐색 시작...")
            iframes = driver.find_elements(By.TAG_NAME, 'iframe')
            
            for i, frame in enumerate(iframes):
                try:
                    driver.switch_to.default_content() # 초기화
                    driver.switch_to.frame(frame)      # i번째 프레임 진입
                    
                    if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) > 0:
                        print(f" -> {i}번째 iframe 안에서 입력창 발견! (진입 성공)")
                        found_input = True
                        break # 찾았으면 이 상태(iframe 안) 유지하고 반복 종료
                except:
                    pass
        
        if not found_input:
            print(f"❌ 오류: '{INPUT_BOX_SELECTOR}' 입력창을 찾을 수 없습니다.")
            print(" -> F12를 눌러 ID나 Class가 맞는지 다시 확인해주세요.")
            print(" -> 혹시 팝업창으로 떴다면 driver.switch_to.window가 필요할 수 있습니다.")
            return

    except Exception as e:
        print(f"초기 설정 실패: {e}")
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
            # 입력
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.clear()
            input_box.send_keys(str(car_num))

            # 조회 버튼 클릭
            confirm_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            driver.execute_script("arguments[0].click();", confirm_btn)

            # 결과 텍스트 대기 및 추출
            result_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)))
            extracted_text = result_element.text
            
            df.at[index, COL_REG_DATE] = extracted_text
            print(f"[{car_num}] : {extracted_text}")
            
            # 너무 빠르면 오류날 수 있으니 약간 대기
            time.sleep(0.5) 

        except Exception as e:
            print(f"[{car_num}] 조회 실패: {e}")
            df.at[index, COL_REG_DATE] = "실패"

    # 저장 및 종료
    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 모든 작업 완료! '{save_name}' 파일에 저장되었습니다.")
    driver.quit()

if __name__ == "__main__":
    run_macro()