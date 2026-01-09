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
# 결과 수집은 이제 ID 패턴으로 루프를 돌리므로 이 SELECTOR는 쓰지 않습니다.

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
    # [수정됨] 화면 미표시 문제 해결을 위해 옵션 최소화
    # ---------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized')  # 창 최대화
    
    # 봇 탐지 방지 (로그인 차단 막기 위해 이건 유지하는 게 좋습니다)
    chrome_options.add_argument("disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # ※ no-sandbox, disable-gpu 등은 모두 삭제했습니다.

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
            driver.execute_script("arguments[0].value = '';", input_box)
            input_box.click()
            input_box.send_keys(str(car_num))
            time.sleep(random.uniform(0.3, 0.5))

            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(random.uniform(0.2, 0.5)) 

            # 2. 버튼 클릭 (2회 반복)
            try:
                text_element = driver.find_element(By.XPATH, "//*[text()='검색']")
                parent_btn = text_element.find_element(By.XPATH, "./..")
                
                print(f" -> [클릭 시도] 버튼 발견")

                for i in range(2):
                    # ActionChains로 강력 클릭
                    action.move_to_element(parent_btn).click().perform()
                    
                    # JS 이벤트도 같이 발송 (보험용)
                    driver.execute_script("""
                        var btn = arguments[0];
                        var event = new MouseEvent('click', {
                            'view': window, 'bubbles': true, 'cancelable': true
                        });
                        btn.dispatchEvent(event);
                    """, parent_btn)
                    
                    time.sleep(0.5)

                print(f" -> 검색 명령 전달 완료")
                time.sleep(1.5) # 로딩 대기
                
            except Exception as e:
                print(f" -> 버튼 조작 실패: {e}")
                
            # =================================================================
            # 3. 결과 수집 (ID 패턴 반복문 적용)
            # =================================================================
            time.sleep(random.uniform(1.0, 2.0))
            
            collected_texts = []
            idx = 0  # 0번부터 시작
            
            while True:
                # 넥사크로 ID 패턴 (gridrow_0...cell_0_5)
                target_id = f"Grid01_00_01_00.body.gridrow_{idx}.cell_{idx}_5"
                
                rows = driver.find_elements(By.ID, target_id)
                
                if not rows:
                    break # 더 이상 없으면 종료
                
                text = rows[0].text.strip()
                if text:
                    collected_texts.append(text)
                
                idx += 1
            
            # 수집된 결과 저장
            if collected_texts:
                full_text = "\n".join(collected_texts)
                df.at[index, COL_REG_DATE] = full_text
                print(f"[{car_num}] 성공! (총 {idx}건 발견)")
            else:
                print(f"[{car_num}] 결과 없음 (행 발견 못함)")
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