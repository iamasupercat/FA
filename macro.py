import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time

# ==========================================
# 1. 사용자 설정 영역
# ==========================================
EXCEL_FILE_PATH = '카히스토리관리_20251230112324.xlsx' # 파일명 확인 필수!
URL = 'https://gaos.glovis.net'

# 입력창 설정
INPUT_BOX_SELECTOR = "input[id*='CARNO']"

# [수정됨] 조회 버튼 설정 (찾아내신 ID 적용!)
BUTTON_SELECTOR = "div[id*='searchBtn']"

# 결과 텍스트 (모든 행의 수리내역)
# Grid01 표 안에 있는 '5번째 열(_5:text)'을 모두 찾습니다.
RESULT_TEXT_SELECTOR = "div[id*='Grid01'][id*='_5:text']"

COL_CAR_NUM = '차량번호'
COL_REG_DATE = '최초등록일'
# ==========================================

def run_macro():
    try:
        # header=1 적용 (2번째 줄을 제목으로 인식)
        df = pd.read_excel(EXCEL_FILE_PATH, header=1)
        print(f"✅ 엑셀 로드 성공: {len(df)}개 데이터")
        print(f"   (읽어온 제목: {df.columns.tolist()})") 
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

    try:
        # =======================================================
        # [Step 1] 사용자 수동 준비
        # =======================================================
        print("\n" + "="*60)
        print("🚨 [사용자 준비 단계] 🚨")
        print("1. 로그인 후 [카히스토리관리] 메뉴로 이동하세요.")
        print("2. 입력창이 보이면 터미널을 클릭하고 Enter를 누르세요.")
        print("-" * 60)
        input("👉 준비되셨으면 엔터(Enter)를 누르세요!")
        print("="*60 + "\n")
        
        # 입력창 찾기 확인
        print("🤖 입력창 찾는 중...")
        try:
            input_box = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, INPUT_BOX_SELECTOR)))
            print(f" -> 입력창 찾기 성공! (ID: {input_box.get_attribute('id')})")
        except:
            print("❌ 입력창을 못 찾았습니다. 페이지가 맞는지 확인해주세요.")
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
            # 1. 입력창 찾고 -> 클릭 -> 지우기 -> 입력
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            input_box.click() 
            input_box.clear()
            input_box.send_keys(str(car_num))
            time.sleep(0.2) # 입력 안정화 대기

            # 2. [수정됨] 찾아낸 ID로 버튼 클릭! 🚀
            search_btn = driver.find_element(By.CSS_SELECTOR, BUTTON_SELECTOR)
            # 넥사크로 버튼은 일반 click()보다 자바스크립트 클릭이 훨씬 확실합니다.
            driver.execute_script("arguments[0].click();", search_btn)
            
            # 3. 결과 수집 (시간을 조금 넉넉히 줌)
            time.sleep(2) 
            
            # [핵심] find_elements(복수형)로 화면에 있는 모든 수리내역 긁어오기
            results = driver.find_elements(By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)
            
            if len(results) > 0:
                # 리스트에 담긴 텍스트들을 줄바꿈(\n)으로 연결해서 하나로 합침
                # 내용이 비어있지 않은 것만 가져옴
                text_list = [r.text for r in results if r.text.strip() != ""]
                full_text = "\n".join(text_list)
                
                df.at[index, COL_REG_DATE] = full_text
                print(f"[{car_num}] 성공! ({len(text_list)}건 발견)")
            else:
                print(f"[{car_num}] 내역 없음 (화면에 표시된 결과가 0개)")
                df.at[index, COL_REG_DATE] = "내역없음"

        except Exception as e:
            print(f"[{car_num}] 에러 발생: {e}")
            df.at[index, COL_REG_DATE] = "에러"

    # 저장
    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 작업 완료! '{save_name}' 파일을 확인하세요.")
    driver.quit()

if __name__ == "__main__":
    run_macro()