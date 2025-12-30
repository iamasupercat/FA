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
URL = 'https://gaos.glovis.net' # 접속할 주소

COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'

# [시스템 버튼]
SYSTEM_BTN_SELECTOR = "img[src*='btn_TF_MenuS.png']"

# [메뉴 박스] 방금 찾으신 class 이름
MENU_BOX_SELECTOR = ".nexaedge" 

# [나머지 설정]
INPUT_BOX_SELECTOR = '#carNumInput'
BUTTON_SELECTOR = '.search-btn'
RESULT_TEXT_SELECTOR = '#repairHistory .date'
# ==========================================

def run_macro():
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except Exception as e:
        print(f"엑셀 오류: {e}")
        return

    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)

    try:
        print("1. 메인 페이지 진입")
        
        # 1. 시스템 버튼 클릭
        system_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SYSTEM_BTN_SELECTOR)))
        system_btn.click()
        print(" -> 시스템 버튼 클릭 완료")

        # 2. 'nexaedge' 박스가 뜰 때까지 대기
        # (iframe으로 들어가는 게 아니라, 그냥 화면에 이 박스가 생길 때까지 기다림)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, MENU_BOX_SELECTOR)))
        print(" -> 메뉴 박스(nexaedge) 열림 확인")

        # 3. 텍스트로 메뉴 찾아서 클릭 ('원부카히스토리')
        # nexaedge 박스 안이든 밖이든, 화면에 이 글자가 보이면 클릭하게 설정
        menu_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '원부카히스토리')]")))
        menu_1.click()
        print(" -> '원부카히스토리' 클릭 성공")
        
        # 4. 이어서 '카히스토리관리' 클릭
        menu_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '카히스토리관리')]")))
        menu_2.click()
        print(" -> '카히스토리관리' 클릭 성공")

        # 5. 입력창이 뜰 때까지 대기
        # (메뉴 누르고 화면이 바뀔 때까지 잠시 대기)
        print(" -> 화면 전환 대기 중...")
        time.sleep(3) 

        # 혹시 입력창이 iframe 안에 있을 수도 있으니 확인
        # (대부분 넥사크로는 iframe을 잘 안 쓰지만, 혹시 입력화면만 iframe일 수 있음)
        if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) == 0:
            print(" -> 메인화면에 입력창이 안보임. 내부 iframe(workFrame) 탐색 시도...")
            iframes = driver.find_elements(By.TAG_NAME, 'iframe')
            for frame in iframes:
                driver.switch_to.default_content()
                try:
                    driver.switch_to.frame(frame)
                    if len(driver.find_elements(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)) > 0:
                        print(" -> 입력창이 있는 iframe 발견 및 진입!")
                        break
                except:
                    pass
        else:
            print(" -> 입력창 발견! (iframe 아님)")

    except Exception as e:
        print(f"초기 메뉴 진입 실패: {e}")
        driver.quit()
        return

    # --- 반복 조회 로직 ---
    print("--- 데이터 조회 시작 ---")
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.clear()
            input_box.send_keys(str(car_num))

            confirm_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            # 넥사크로/JS 버튼은 click()이 잘 안 먹힐 때가 있어 execute_script 사용 권장
            driver.execute_script("arguments[0].click();", confirm_btn)

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
    print("완료")
    driver.quit()

if __name__ == "__main__":
    run_macro()