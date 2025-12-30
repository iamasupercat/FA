import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# ==========================================
# 1. 사용자 설정 영역 (이 부분만 본인 정보로 채우세요)
# ==========================================
EXCEL_FILE_PATH = '차량목록.xlsx'
URL = 'https://gaos.glovis.net'

# [로그인 정보 입력]
MY_ID = '여기에_본인_아이디'
MY_PW = '여기에_본인_비밀번호'

# [로그인 페이지 Selector] (F12 -> Copy Selector로 찾아오세요)
LOGIN_ID_SELECTOR = '#userId'        # 예시: 아이디 입력창
LOGIN_PW_SELECTOR = '#userPw'        # 예시: 비번 입력창
LOGIN_BTN_SELECTOR = '#btnLogin'     # 예시: 로그인 버튼

# [메인 페이지 Selector]
SYSTEM_BTN_SELECTOR = "img[src*='btn_TF_MenuS.png']" # 시스템 버튼
MENU_BOX_SELECTOR = ".nexaedge"                      # 메뉴 박스

# [조회 화면 Selector]
INPUT_BOX_SELECTOR = '#carNumInput'           # 차량번호 입력창
BUTTON_SELECTOR = '.search-btn'               # 검색 버튼
RESULT_TEXT_SELECTOR = '#repairHistory .date' # 결과 텍스트

# 엑셀 컬럼명
COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except Exception as e:
        print(f"엑셀 파일 오류: {e}")
        return

    options = webdriver.ChromeOptions()
    # options.add_argument('--start-maximized') # 창을 최대화해서 보고 싶으면 주석 해제
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    driver.get(URL)
    wait = WebDriverWait(driver, 20) # 로그인 및 로딩 시간 고려해서 넉넉히 20초

    try:
        # -------------------------------------------------------
        # [Step 1] 로그인 프로세스
        # -------------------------------------------------------
        print("1. 로그인 시도 중...")
        
        # 1. 아이디 입력
        id_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, LOGIN_ID_SELECTOR)))
        id_box.clear()
        id_box.send_keys(MY_ID)
        
        # 2. 비밀번호 입력
        pw_box = driver.find_element(By.CSS_SELECTOR, LOGIN_PW_SELECTOR)
        pw_box.clear()
        pw_box.send_keys(MY_PW)
        
        # 3. 로그인 버튼 클릭
        login_btn = driver.find_element(By.CSS_SELECTOR, LOGIN_BTN_SELECTOR)
        driver.execute_script("arguments[0].click();", login_btn) # 확실한 클릭
        
        print(" -> 로그인 버튼 클릭. 메인 페이지 로딩 대기...")

        # -------------------------------------------------------
        # [Step 2] 메인 페이지 진입 및 메뉴 이동
        # -------------------------------------------------------
        
        # [중요] 로그인이 성공해서 '시스템 버튼'이 보일 때까지 기다림
        # 이 부분이 없으면 로그인이 다 되기도 전에 버튼을 찾으려다 에러가 납니다.
        system_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SYSTEM_BTN_SELECTOR)))
        print("2. 메인 페이지 진입 성공")
        
        # 시스템 버튼 클릭
        system_btn.click()
        print(" -> 시스템 버튼 클릭 완료")

        # 메뉴 박스(nexaedge) 대기
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, MENU_BOX_SELECTOR)))
        
        # 텍스트로 메뉴 찾기
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '원부카히스토리')]"))).click()
        print(" -> '원부카히스토리' 클릭")
        
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '카히스토리관리')]"))).click()
        print(" -> '카히스토리관리' 클릭")

        # 입력 화면 전환 대기
        time.sleep(3)
        
        # iframe 체크 (혹시 입력창이 iframe에 있으면 진입)
        if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) == 0:
            print(" -> 메인 프레임에 입력창 없음. iframe 탐색...")
            for frame in driver.find_elements(By.TAG_NAME, 'iframe'):
                driver.switch_to.default_content()
                try:
                    driver.switch_to.frame(frame)
                    if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) > 0:
                        print(" -> 입력창 iframe 진입 성공")
                        break
                except: pass

    except Exception as e:
        print(f"[오류 발생] 초기 설정 실패: {e}")
        driver.quit()
        return

    # -------------------------------------------------------
    # [Step 3] 데이터 조회 반복 (기존과 동일)
    # -------------------------------------------------------
    print("--- 엑셀 데이터 조회 시작 ---")
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            # 입력
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.clear()
            input_box.send_keys(str(car_num))

            # 조회 버튼
            confirm_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            driver.execute_script("arguments[0].click();", confirm_btn)

            # 결과 대기 및 추출
            result_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)))
            extracted_text = result_element.text
            
            df.at[index, COL_REG_DATE] = extracted_text
            print(f"[{car_num}] : {extracted_text}")
            time.sleep(0.5) # 서버 부하 방지용 딜레이

        except Exception as e:
            print(f"[{car_num}] 조회 실패")
            df.at[index, COL_REG_DATE] = "실패"

    # 결과 저장
    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print("모든 작업 완료!")
    driver.quit()

if __name__ == "__main__":
    run_macro()