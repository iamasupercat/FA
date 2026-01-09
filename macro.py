import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
import random

# ==========================================
# 1. 사용자 설정 영역
# ==========================================
EXCEL_FILE_PATH = '카히스토리관리_20251230112324.xlsx' 
URL = 'https://gaos.glovis.net'

INPUT_BOX_SELECTOR = "input[id*='CARNO']"
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

    # ---------------------------------------------------------
    # 탭 크래시 방지 옵션
    # ---------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized') 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)
    action = ActionChains(driver)

    try:
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

    print("--- 데이터 조회 시작 ---")
    
    for index, row in df.iterrows():
        car_num = row[COL_CAR_NUM]
        if pd.isna(car_num): continue

        try:
            # 1. 입력창에 값 넣기
            input_box = driver.find_element(By.CSS_SELECTOR, INPUT_BOX_SELECTOR)
            
            # JS로 깔끔하게 값 비우기
            driver.execute_script("arguments[0].value = '';", input_box)
            input_box.click()
            input_box.send_keys(str(car_num))
            time.sleep(random.uniform(0.3, 0.5))

            # 빈 공간 클릭 (입력 확정)
            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(random.uniform(0.2, 0.5)) 

            # 2. 버튼 클릭 및 상태 확인 (2회 반복)
            try:
                # '검색' 텍스트를 가진 요소의 부모(버튼 본체) 찾기
                text_element = driver.find_element(By.XPATH, "//*[text()='검색']")
                parent_btn = text_element.find_element(By.XPATH, "./..")
                
                print(f" -> [클릭 시도] 버튼 발견")

                for i in range(2):
                    # (A) 클릭 전 상태 확인
                    before_status = parent_btn.get_attribute("userstatus")
                    
                    # (B) ActionChains로 강력 클릭 (마우스 이동 -> 클릭)
                    # 넥사크로는 move_to_element를 해야 mouseover 상태가 되어 클릭이 잘 먹힘
                    action.move_to_element(parent_btn).click().perform()
                    
                    # (C) 클릭 직후 상태 확인 (매우 빠르게 지나가서 null일 수도 있음)
                    # 약간의 딜레이 후 확인 (pushed 상태인지, 혹은 반응이 있었는지)
                    time.sleep(0.1) 
                    after_status = parent_btn.get_attribute("userstatus")
                    
                    print(f"    ({i+1}/2회차) 상태변화: {before_status} -> {after_status}")
                    
                    # (D) 만약 ActionChains가 안 먹혔을 경우 대비용 JS 이벤트 발송 (속성 변경 아님)
                    # 실제 마우스 이벤트를 시뮬레이션
                    driver.execute_script("""
                        var btn = arguments[0];
                        var event = new MouseEvent('click', {
                            'view': window,
                            'bubbles': true,
                            'cancelable': true
                        });
                        btn.dispatchEvent(event);
                    """, parent_btn)
                    
                    time.sleep(0.5) # 클릭 간 간격

                print(f" -> 검색 명령 전달 완료")
                
                # 로딩 대기
                time.sleep(1.5) 
                
            except Exception as e:
                print(f" -> 버튼 조작 실패: {e}")
                
            # 3. 결과 수집
            time.sleep(random.uniform(1.0, 2.0))
            
            results = driver.find_elements(By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)
            
            if len(results) > 0:
                text_list = [r.text for r in results if r.text.strip() != ""]
                full_text = "\n".join(text_list)
                
                if full_text:
                    df.at[index, COL_REG_DATE] = full_text
                    print(f"[{car_num}] 성공!")
                else:
                    df.at[index, COL_REG_DATE] = "값 없음"
                    print(f"[{car_num}] 값 없음")
            else:
                print(f"[{car_num}] 결과 없음")
                df.at[index, COL_REG_DATE] = "내역없음"
            
            time.sleep(random.uniform(0.5, 1.0))

        except Exception as e:
            print(f"[{car_num}] 에러: {e}")
            df.at[index, COL_REG_DATE] = "에러"

    save_name = '결과포함_' + EXCEL_FILE_PATH
    df.to_excel(save_name, index=False)
    print(f"\n✅ 끝! '{save_name}' 저장 완료.")
    driver.quit()

if __name__ == "__main__":
    run_macro()