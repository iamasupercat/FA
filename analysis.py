from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# URL 설정
URL = 'https://gaos.glovis.net'

def run_scanner():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    # 보안 관련 옵션 유지
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)

    try:
        # 1. 수동 준비 단계
        print("\n" + "="*60)
        print("🚨 [진단 모드] 🚨")
        print("1. 직접 로그인하고 '차량번호 입력창'이 있는 화면까지 이동하세요.")
        print("2. 입력창이 눈에 보이면 아래 터미널을 클릭하고 [Enter]를 누르세요.")
        print("-" * 60)
        input("👉 준비 완료되면 엔터(Enter)를 누르세요!")
        print("="*60 + "\n")

        print("🔍 화면 스캔 시작...")
        
        # 2. 모든 iframe을 다 뒤져서 input 태그 찾기
        # (1) 메인 프레임 스캔
        print(f"--- [1] 메인 프레임(Main) 스캔 결과 ---")
        scan_inputs(driver)

        # (2) iframe 내부 스캔
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        print(f"\n--- [2] iframe 스캔 결과 (총 {len(iframes)}개 발견) ---")
        
        for i, frame in enumerate(iframes):
            print(f"\n>> {i}번째 iframe 내부 진입 시도...")
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                scan_inputs(driver)
            except Exception as e:
                print(f"   (접근 불가: {e})")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        print("\n✅ 스캔 종료. 이 결과를 복사해서 알려주세요.")
        # driver.quit() # 확인을 위해 창 안 닫음

def scan_inputs(driver):
    # 화면에 보이는 input 태그만 찾음
    try:
        inputs = driver.find_elements(By.TAG_NAME, 'input')
        visible_count = 0
        
        for inp in inputs:
            try:
                # 눈에 보이거나 크기가 0보다 큰 경우만 출력
                if inp.is_displayed() or inp.size['width'] > 0:
                    visible_count += 1
                    input_id = inp.get_attribute('id')
                    input_class = inp.get_attribute('class')
                    input_name = inp.get_attribute('name')
                    print(f"   Found! [Type: Input] | ID: {input_id} | Class: {input_class} | Name: {input_name}")
            except:
                pass
        
        # input이 없으면 textarea도 찾아봄
        textareas = driver.find_elements(By.TAG_NAME, 'textarea')
        for ta in textareas:
             if ta.is_displayed():
                visible_count += 1
                print(f"   Found! [Type: TextArea] | ID: {ta.get_attribute('id')}")

        if visible_count == 0:
            print("   (이 영역에는 눈에 보이는 입력창이 없습니다)")
            
    except Exception as e:
        print(f"   스캔 중 에러: {e}")

if __name__ == "__main__":
    run_scanner()