# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템
- API 통신 및 데이터 수집 (Network Layer)
"""
import os
import sys
try:
    import requests

    REQUESTS_AVAILABLE = True
except ImportError:
    requests = None
    REQUESTS_AVAILABLE = False
    print("경고: requests 라이브러리가 설치되지 않았습니다.")

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    SELENIUM_AVAILABLE = True
except ImportError:
    webdriver = None
    Options = None
    SELENIUM_AVAILABLE = False
    print("경고: selenium 라이브러리가 설치되지 않았습니다.")


class ApiManager:
    def __init__(self):
        self.session = None
        self.driver = None

    def setup_manual_session(self, cookie_val: str) -> None:
        """수동 쿠키를 이용해 Requests 세션을 생성합니다."""
        if not REQUESTS_AVAILABLE:
            raise RuntimeError("requests 라이브러리가 필요합니다.")
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Cookie": cookie_val
        })

    def start_selenium_login(self) -> None:
        """Selenium을 이용해 크롬 브라우저를 띄웁니다."""
        if not SELENIUM_AVAILABLE:
            raise RuntimeError("Selenium 라이브러리가 필요합니다.")
        
        options = Options()
        service = None
        
        # 1. 로컬 경로에서 chromedriver.exe 확인 (프로그램 폴더 또는 실행 파일 위치)
        base_dir = os.getcwd()
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
            
        local_driver = os.path.join(base_dir, "chromedriver.exe")
        
        if os.path.exists(local_driver):
            try:
                service = Service(local_driver)
            except Exception:
                service = None

        # 2. 로컬에 없거나 실패했다면 ChromeDriverManager 시도
        if service is None:
            try:
                # [수정] ChromeDriverManager를 사용하여 브라우저 버전에 맞는 드라이버를 자동 설치/사용합니다.
                # 네트워크 연결이 불안정할 경우 여기서 ConnectionError 등이 발생할 수 있습니다.
                driver_path = ChromeDriverManager().install()
                service = Service(driver_path)
            except Exception as e:
                # 3. 모든 자동 설치가 실패했다면 상세한 안내와 함께 에러 발생
                error_msg = (
                    f"크롬 드라이버를 준비하지 못했습니다: {e}\n\n"
                    "오프라인 환경이거나 네트워크 연결이 불안정합니다.\n"
                    "해결 방법:\n"
                    "1. 인터넷 연결 상태를 확인하고 다시 시도하세요.\n"
                    "2. 또는 https://googlechromelabs.github.io/chrome-for-testing/ 에서\n"
                    "   본인의 크롬 버전에 맞는 'chromedriver.exe'를 다운로드하여\n"
                    "   프로그램 설치 폴더에 직접 넣어주세요."
                )
                raise RuntimeError(error_msg)

        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.get("https://ls.coupang.net")

    def finalize_selenium_login(self) -> None:
        """로그인된 크롬 브라우저에서 쿠키를 추출하여 Requests 세션으로 전환합니다."""
        if not self.driver:
            raise RuntimeError("실행된 브라우저가 없습니다.")
        if not REQUESTS_AVAILABLE:
            raise RuntimeError("requests 라이브러리가 필요합니다.")

        cookies = self.driver.get_cookies()
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        })
        for cookie in cookies:
            self.session.cookies.set(cookie['name'], cookie['value'])

        self.driver.quit()
        self.driver = None

    def patch_data(self, url: str, payload: dict) -> dict:
        """주어진 URL에 PATCH 요청으로 데이터를 업데이트합니다. (타임아웃 30초)"""
        if not self.session:
            raise RuntimeError("연결된 세션(쿠키)이 없습니다. 먼저 로그인을 진행하세요.")

        headers = {"Content-Type": "application/json"}
        try:
            response = self.session.patch(url, json=payload, headers=headers, timeout=30)

            if response.status_code in (401, 403):
                raise PermissionError(f"인증 에러({response.status_code})가 발생했습니다. 쿠키 세션이 만료되었을 수 있습니다.\n브라우저 로그인을 다시 진행해주세요.")

            response.raise_for_status()

            try:
                return response.json()
            except ValueError:
                return {"status": "ok", "text": response.text[:200]}

        except Exception as e:
            raise e

    def put_data(self, url: str, payload: dict) -> dict:
        """주어진 URL에 PUT 요청으로 데이터를 완전히 업데이트합니다. (타임아웃 30초)"""
        if not self.session:
            raise RuntimeError("연결된 세션(쿠키)이 없습니다. 먼저 로그인을 진행하세요.")

        headers = {"Content-Type": "application/json"}
        try:
            response = self.session.put(url, json=payload, headers=headers, timeout=30)

            if response.status_code in (401, 403):
                raise PermissionError(f"인증 에러({response.status_code})가 발생했습니다. 쿠키 세션이 만료되었을 수 있습니다.\n브라우저 로그인을 다시 진행해주세요.")

            response.raise_for_status()

            try:
                return response.json()
            except ValueError:
                return {"status": "ok", "text": response.text[:200]}

        except Exception as e:
            raise e

    def fetch_data(self, url: str) -> dict:
        """주어진 URL에서 JSON 데이터를 수집합니다. (타임아웃 30초, 최대 3회 재시도)"""
        if not self.session:
            raise RuntimeError("연결된 세션(쿠키)이 없습니다. 먼저 로그인을 진행하세요.")

        last_err = None
        for attempt in range(3):
            try:
                response = self.session.get(url, timeout=30)
                
                # 인증 에러 (세션 만료 등)
                if response.status_code in (401, 403):
                    raise PermissionError(f"인증 에러({response.status_code})가 발생했습니다. 쿠키 세션이 만료되었을 수 있습니다.\n브라우저 로그인을 다시 진행해주세요.")

                response.raise_for_status()

                try:
                    return response.json()
                except ValueError:
                    preview_text = response.text[:200].replace('\n', ' ')
                    raise ValueError(f"서버에서 올바른 데이터를 받지 못했습니다. (세션 만료 또는 접속 차단 의심)\n응답내용 미리보기: {preview_text}")

            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                last_err = e
                print(f"[네트워크] 데이터 수집 재시도 중... ({attempt + 1}/3): {e}")
                if attempt < 2:
                    import time
                    time.sleep(1) # 잠깐 대기 후 재시도
                continue
            except Exception as e:
                raise e

        # 모든 시도 실패 시
        raise last_err if last_err else RuntimeError("데이터 수집에 실패했습니다.")

    def fetch_truck_order_tracking(self, start_date: str, end_date: str, location_start: int = 5) -> dict:
        """Truck Order Tracking 리스트를 조회합니다.
        URL 예시:
        https://ls.coupang.net/data/truckOrderTracking?statuses=CONFIRMED,BACK&orderStartDate={start_date}&orderEndDate={end_date}&usageTypes=&locationStart={location_start}&locale=ko_KR&page=0&pageSize=300
        """
        if not self.session:
            raise RuntimeError("API 세션이 없습니다. 로그인 후 사용하세요.")
        base_url = "https://ls.coupang.net/data/truckOrderTracking"
        params = (
            f"?statuses=CONFIRMED,BACK"
            f"&orderStartDate={start_date}"
            f"&orderEndDate={end_date}"
            f"&usageTypes="
            f"&locationStart={location_start}"
            f"&locale=ko_KR"
            f"&page=0&pageSize=300"
        )
        url = base_url + params
        try:
            response = self.session.get(url, timeout=30)
            if response.status_code in (401, 403):
                raise PermissionError(f"인증 에러({response.status_code})가 발생했습니다. 쿠키 세션이 만료되었을 수 있습니다.\n브라우저 로그인을 다시 진행해주세요.")
            response.raise_for_status()
            return response.json()
        except Exception as e:
            raise e
