import os
import random
import time
import pandas as pd
import schedule
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from PIL import Image
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from joonggonara_auto_upload import myIdPW # 네이버 ID PW 가 있는 파일


# #TODO 중고나라 올린 양식대로 게시글을 올리고 설정한 시간마다 다시 글써주는
class NaverAutoPoster:
    def __init__(self, id, pw):
        self.id = id
        self.pw = pw
        self.self.driver = self.init_self.driver()
        self.self.df = pd.read_excel('/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/joonggonara_macro.xlsx') #TODO 엑셀파일이 있는 경로
        self.image_directory = "/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/TrashImage"    #TODO 원본 이미지 저장 경로
        self.self.save_directory = "/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/joonggo_img" #TODO 이미지 합치기 저장 경로
        self.no_up = 0

    def init_driver(self): #mac m1 os 에서 셀레니움 4.10.0 을 사용중이다.
        options = Options()
        options.add_argument("start-maximized")
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("user-data-dir=/Users/kangjoon/Desktop/WORKSPACE/PYTHON/Macro")
        options.add_argument(
            'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36')

        return self.webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)#크롬드라이버 자동 설치 셀레니움 4.10.0 을 사용중

    def login(self):
        self.self.driver.get('http://cafe.naver.com/joonggonara')
        login_btn = self.self.driver.find_element(By.ID, 'gnb_login_button')
        login_btn.click()
        self.self.driver.execute_script(f"document.getElementsByName('id')[0].value='{self.id}'") #캡챠 방지를 위해 execute_script
        time.sleep(1)
        self.self.driver.execute_script(f"document.getElementsByName('pw')[0].value='{self.pw}'")
        time.sleep(1)
        self.self.driver.find_element(By.ID, 'log.login').click() # 로그인 버튼을 찾아 클릭

    def merge_image(self,images_paths, title_name):  # 이미지 합치기
        images = [Image.open(os.path.join(self.image_directory, x.strip())) for x in images_paths]

        if self.self.img_width > -1:
            image_sizes = [(int(self.self.img_width), int(self.self.img_width * x.size[1] / x.size[0])) for x in images]
        else:
            image_sizes = [(x.size[0], x.size[1]) for x in images]

        widths, heights = zip(*(image_sizes))
        max_width, total_height = max(widths), sum(heights)
        if self.img_space > 0:
            total_height += (self.img_space * (len(images) - 1))

        result_img = Image.new("RGB", (max_width, total_height), (255, 255, 255))
        y_offset = 0

        for idx, img in enumerate(images):
            if self.img_width > -1:
                img = img.resize(image_sizes[idx])

            result_img.paste(img, (0, y_offset))
            y_offset += (img.size[1] + self.img_space)

        file_name = title_name + '.' + self.img_format  # 파일 이름은 본문 제목 + 이미지 포멧
        dest_path = os.path.join(self.save_directory, file_name)
        result_img.save(dest_path)
        print(f"Operation completed. Saved merged image as {file_name}.")
        return dest_path

    def find_files(self, filename, search_path):  # 파일찾기 mac 버전
        result = []
        for root, dir, files in os.walk(search_path):
            if filename in files:
                result.append(os.path.join(root, filename))
        return result

    def post_article(self):
        for i in range(len(self.df.index)):
            time.sleep(0.5)
            if self.df.loc[i, '카테고리'] in ('공연', '연극', '영화'):
                self.driver.find_element(By.ID, 'menuLink1285').click()
            elif self.df.loc[i, '카테고리'] in ('스포츠'):
                self.driver.find_element(By.ID, 'menuLink1286').click()

            self.driver.switch_to.frame('cafe_main')  # 중요 Iframe 을 바꿔야해
            self.driver.find_element(By.ID, 'writeFormBtn').click()  # 글쓰기
            self.driver.switch_to.window(self.driver.window_handles[i + 1 - self.no_up])

            self.driver.find_element(By.CLASS_NAME, 'textarea_input').send_keys(self.df.loc[i, '상품명'])
            time.sleep(0.5)
            self.driver.find_element(By.CLASS_NAME, 'input_text').send_keys(int(self.df.loc[i, '판매가격']))

            # 상품 상태 입력
            if self.df.loc[i, '상품상태'] == '미개봉':
                quality_status = 1
            elif self.df.loc[i, '상품상태'] == '거의새것':
                quality_status = 2
            elif self.df.loc[i, '상품상태'] == '사용감있음':
                quality_status = 3
            self.driver.find_element(By.ID, f'quality{quality_status}').click()  # 상품 상태 클릭

            # 직접결제
            payment_type = self.driver.find_elements(By.XPATH, '//div[@class="deal_item"]')
            payment_type[1].click()
            naver_pay_type = self.driver.find_element(By.XPATH, '//label[@class="switch_slider"]')
            naver_pay_type.click()

            # 배송방법
            if self.df.loc[i, '배송방법'] == '직거래':
                delivery_status = 0
            elif self.df.loc[i, '배송방법'] == '택배거래':
                delivery_status = 1
            elif self.df.loc[i, '배송방법'] == '온라인전송':
                delivery_status = 2
            self.driver.find_element(By.XPATH,
                                f'//label[@for="delivery{delivery_status}"]').click()  # input 을 클릭할 수 없어서 label 을 대신

            # 셀러 번호 공개
            self.driver.find_element(By.XPATH, '//label[@for="agree1"]').click()  # 휴대전화번호 노출 동의
            self.driver.find_element(By.XPATH, '//label[@for="agree2"]').click()  # 안심번호 이용

            self.driver.find_element(By.XPATH, '//button[@data-name="image"]').click()  # 이미지 첨부 클릭
            time.sleep(1)
            os.system(
                '''osascript -e 'tell application "System Events" to keystroke "w" using command down' ''')  # 첨부파일 창 닫기
            time.sleep(1)

            # 이미지 첨부
            images_paths = self.df.loc[i, '이미지'].split(" ")
            title_names = self.df.loc[i, '상품명'].split(" ")
            title_name = title_names[0] + title_names[1]

            # 설정한 경로에 사진 통합 이미지가 있으면 그냥 사용하고, 없으면 merge_image() 함수로 여러장의 사진을 한장으로 통합합니다.

            search_path = '/Users/kangjoon/Desktop/seller/Resale/joonggo_img'
            file_with_000 = [f for f in os.listdir(search_path) if
                             title_name in f and f.endswith((".png", ".jpg", ".jpeg"))]
            file_found = False
            for file in file_with_000:
                full_file_path = self.find_files(file, search_path)
                for f in full_file_path:
                    print(f)
                    file_found = True

            if file_found:
                image = os.path.join(search_path, title_name + '.' + self.img_format)
                file_found = False
            else:
                image = self.merge_image(images_paths, title_name)

            input_file_element = self.driver.find_element(By.ID, 'hidden-file')
            input_file_element.send_keys(image)

            time.sleep(random.uniform(8, 10))

            # 본문
            webdriver.ActionChains(self.driver).send_keys(self.df.loc[i, '본문']).perform()
            time.sleep(random.uniform(4, 8))

            # 작성 완료
            self.driver.find_element(By.XPATH, '//a[@role="button"][contains(@class, "BaseButton")]').click()  # 작성완료
            time.sleep(random.uniform(3, 5))
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(random.uniform(3, 5))

    def run(self):
        self.login()
        for i in range(len(self.self.df.index)):
            self.post_article()
            time.sleep(random.uniform(20, 50))


def main():
    naver_id = myIdPW.naver_id
    naver_pw = myIdPW.naver_pw
    poster = NaverAutoPoster(naver_id, naver_pw)
    poster.run()

    schedule.every(6).hours.do(poster.run)

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    main()
