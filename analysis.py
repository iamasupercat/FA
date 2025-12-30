import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# 설정 영역
# ==========================================
URL = 'https://gaos.glovis.net'

def run_button_detective():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(URL)

    try:
        # 1. 수동 준비
        print("\n" + "="*60)
        print("🕵️ [조회 버튼 범인 찾기] 🕵️")
        print("1. 브라우저에서 직접 로그인하고 메뉴를 이동해주세요.")
        print("2. '조회' 버튼이 눈에 보이면...")
        input("👉 여기(터미널)를 클릭하고 엔터(Enter)를 누르세요! 수사를 시작합니다.")
        print("="*60 + "\n")

        print("🔍 화면에서 '조회'나 'Search'와 관련된 요소를 싹 긁어모으는 중...")

        # 후보군 수집 전략
        candidates = []
        
        # 전략 1: "조회"라는 텍스트를 가진 모든 요소
        try:
            candidates.extend(driver.find_elements(By.XPATH, "//*[contains(text(), '조회')]"))
        except: pass
        
        # 전략 2: ID나 Class에 'btn', 'search'가 들어간 요소 (Nexacro 버튼 패턴)
        try:
            candidates.extend(driver.find_elements(By.CSS_SELECTOR, "[id*='btn'], [class*='btn']"))
            candidates.extend(driver.find_elements(By.CSS_SELECTOR, "[id*='Search'], [id*='search']"))
        except: pass

        # 중복 제거 및 눈에 보이는 것만 필터링
        visible_candidates = []
        seen_ids = set()
        
        for elem in candidates:
            try:
                if elem.is_displayed() and elem.size['width'] > 0:
                    eid = elem.get_attribute('id')
                    if eid not in seen_ids:
                        visible_candidates.append(elem)
                        seen_ids.add(eid)
            except: pass

        print(f"👉 총 {len(visible_candidates)}개의 용의자를 확보했습니다. 하나씩 확인합니다.\n")

        # 2. 하나씩 빨간 박스 치면서 물어보기
        for i, elem in enumerate(visible_candidates):
            try:
                elem_id = elem.get_attribute('id')
                elem_txt = elem.text.strip()
                elem_tag = elem.tag_name
                
                # 시각적 강조 (빨간 테두리 + 노란 배경)
                driver.execute_script("arguments[0].style.border='5px solid red'", elem)
                driver.execute_script("arguments[0].style.backgroundColor='yellow'", elem)
                
                print(f"[{i+1}/{len(visible_candidates)}] 화면을 보세요! 빨간 박스가 쳐졌나요?")
                print(f"   정보: Tag={elem_tag} | Text='{elem_txt}'")
                print(f"   ID: {elem_id}")
                
                answer = input("👉 이게 '조회 버튼'이 맞으면 'y', 아니면 엔터: ").strip().lower()

                # 강조 해제
                driver.execute_script("arguments[0].style.border=''", elem)
                driver.execute_script("arguments[0].style.backgroundColor=''", elem)

                if answer == 'y':
                    print("\n🎉 범인 검거 완료!")
                    print("="*50)
                    print("코드의 BUTTON_SELECTOR 변수를 아래 내용으로 바꾸세요:")
                    
                    # 꿀팁: 가장 확실한 Selector 생성해주기
                    if elem_id:
                        # ID가 너무 길면 뒤에 짤라서 키워드만 추출
                        parts = elem_id.split('.')
                        keyword = parts[-1] if len(parts) > 0 else elem_id
                        print(f'\nBUTTON_SELECTOR = "div[id*=\'{keyword}\']"')
                        print(f"# (참고: 원본 ID는 {elem_id})")
                    else:
                        print(f'\nBUTTON_SELECTOR = "//*[contains(text(), \'{elem_txt}\')]"')
                    
                    print("="*50)
                    break
            
            except Exception as e:
                print(f"   (확인 중 에러 발생, 다음으로 넘어갑니다)")
                continue

    except Exception as e:
        print(f"오류 발생: {e}")
    
    print("\n수사 종료. 창을 닫아도 됩니다.")
    # driver.quit()

if __name__ == "__main__":
    run_button_detective()