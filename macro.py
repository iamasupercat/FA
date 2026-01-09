import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options # Options 모듈 명시
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
    # [중요] 탭 크래시 방지 및 성능 최적화 옵션
    # ---------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized') 
    chrome_options.add_argument("--no-sandbox")            # 탭 다운 방지 필수
    chrome_options.add_argument("--disable-dev-shm-usage") # 메모리 부족 방지
    chrome_options.add_argument("--disable-gpu")           # GPU 가속 끄기
    chrome_options.add_argument("disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get(URL)
    wait = WebDriverWait(driver, 15)

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
            
            # 입력 전 초기화 로직 강화 (JS로 비우기)
            driver.execute_script("arguments[0].value = '';", input_box)
            input_box.click()
            input_box.send_keys(str(car_num))
            time.sleep(random.uniform(0.3, 0.5))

            # 빈 공간 클릭 (Focus Out 효과)
            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(random.uniform(0.2, 0.5)) 

            # 2. [수정됨] 넥사크로 이벤트 강제 주입 (Attribute 변경)
            try:
                # '검색' 텍스트를 가진 요소의 부모(버튼 본체) 찾기
                text_element = driver.find_element(By.XPATH, "//*[text()='검색']")
                parent_btn = text_element.find_element(By.XPATH, "./..")
                
                # 자바스크립트로 status, userstatus 속성 강제 변경
                # (단순 click() 대신 넥사크로 엔진이 반응하도록 상태값 조작)
                driver.execute_script("""
                    var btn = arguments[0];
                    
                    // 1. 마우스 오버 상태로 변경
                    btn.setAttribute('status', 'mouseover');
                    
                    // 2. 잠시 후 클릭(Pushed) 상태로 변경하여 이벤트 트리거 유도
                    setTimeout(function() {
                        btn.setAttribute('userstatus', 'pushed');
                        // 상태 변경과 함께 클릭 이벤트도 dispatch하여 확실하게 처리
                        btn.click(); 
                    }, 100);
                    
                    // 3. (옵션) 잠시 후 상태 복구 (다음 루프를 위해)
                    setTimeout(function() {
                        btn.removeAttribute('userstatus');
                        btn.setAttribute('status', 'enabled');
                    }, 500);
                """, parent_btn)

                print(f" -> 검색 실행 (Attribute Injection)")
                
                # 로딩 대기 (서버 응답 시간)
                time.sleep(1.5) 
                
            except Exception as e:
                print(f" -> 버튼 조작 실패: {e}")
                
            # 3. 결과 수집
            time.sleep(random.uniform(1.0, 2.0))
            
            results = driver.find_elements(By.CSS_SELECTOR, RESULT_TEXT_SELECTOR)
            
            if len(results) > 0:
                text_list = [r.text for r in results if r.text.strip() != ""]
                full_text = "\n".join(text_list)
                
                # 결과값이 기존 값과 같다면(로딩 지연 등) 재시도 로직이 필요할 수 있음
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