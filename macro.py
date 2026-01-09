import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import random

# ==========================================
# 1. 사용자 설정 영역
# ==========================================
EXCEL_FILE_PATH = '카히스토리관리_20251230112324.xlsx' 
URL = 'https://gaos.glovis.net'

# 입력창 Selector
INPUT_BOX_SELECTOR = "input[id*='CARNO']"

# 결과 텍스트 Selector
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
    
    # ActionChains는 마우스 이동용으로만 준비 (클릭용 아님)
    action = ActionChains(driver)

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
            time.sleep(random.uniform(0.3, 0.7)) # 입력 후 잠시 대기

            # [중요] 입력 확정을 위해 빈 공간(body) 클릭
            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(random.uniform(0.2, 0.5)) 

            # 2. 버튼 클릭 (부모 요소 타겟팅 + JS 1회 클릭)
            try:
                # (1) '검색' 글자를 가진 요소를 먼저 찾습니다.
                text_element = driver.find_element(By.XPATH, "//*[text()='검색']")
                
                # (2) 그 텍스트의 '바로 위 부모(버튼 상자)'를 찾습니다.
                # XPath에서 '/..' 는 '내 부모'를 뜻합니다.
                parent_btn = text_element.find_element(By.XPATH, "./..")
                
                # (3) 부모 버튼에 JS로 클릭 명령 1회 전송 (서버 부하 최소화)
                driver.execute_script("arguments[0].click();", parent_btn)
                print(" -> 검색 버튼(부모 요소) 클릭 완료")
                
                # (4) [필수] 클릭 직후 서버가 반응할 시간을 충분히 줍니다.
                time.sleep(1.0) 
                
            except Exception as e:
                print(f" -> 버튼 클릭 실패: {e}")
                
            # 3. 결과 수집 (로딩 대기 포함)
            # 서버 응답 시간에 따라 이 시간을 조절하세요 (기본 1.5~3.0초)
            time.sleep(random.uniform(1.5, 3.0))
            
            results = driver.find_elements(By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)
            
            if len(results) > 0:
                text_list = [r.text for r in results if r.text.strip() != ""]
                full_text = "\n".join(text_list)
                df.at[index, COL_REG_DATE] = full_text
                print(f"[{car_num}] 성공! ({len(text_list)}행)")
            else:
                print(f"[{car_num}] 결과 없음 (혹은 로딩 지연)")
                df.at[index, COL_REG_DATE] = "내역없음"
            
            # 다음 차례 넘어가기 전 안전 딜레이 (서버 보호)
            time.sleep(random.uniform(0.5, 1.0))

        except Exception as e:
            print(f"[{car_num}] 에러 발생: {e}")
            df.at[index, COL_REG_DATE] = "에러"

    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 작업 종료. '{save_name}' 파일을 확인하세요.")
    driver.quit()

if __name__ == "__main__":
    run_macro()