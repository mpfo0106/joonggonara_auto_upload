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
import myIdPW

EXCEL_PATH = "/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/joonggonara_macro.xlsx"
IMAGES_DIR = "/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/images"
COMPLETE_JOONGGO_IMG_DIR = "/Users/kangjoon/WORKSPACE/PYTHON/Macro/joonggonara_auto_upload/joonggo_img"


class NaverAutoPoster:
    def __init__(self, user_id, user_pw):
        self.id = user_id
        self.pw = user_pw
        self.driver = self.init_driver()
        self.df = pd.read_excel(EXCEL_PATH)
        self.image_directory = IMAGES_DIR
        self.save_directory = COMPLETE_JOONGGO_IMG_DIR
        self.img_format = "jpeg"
        self.img_width = -1
        self.img_space = 0

    def init_driver(self):
        options = Options()
        options.add_argument("start-maximized")
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("user-data-dir=/Users/kangjoon/Desktop/WORKSPACE/PYTHON/Macro")
        options.add_argument(
            'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36')

        return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    def login(self):
        self.driver.get('http://cafe.naver.com/joonggonara')
        login_btn = self.driver.find_element(By.ID, 'gnb_login_button')
        login_btn.click()
        time.sleep(0.5)
        self.driver.execute_script(f"document.getElementsByName('id')[0].value='{self.id}'")
        time.sleep(1)
        self.driver.execute_script(f"document.getElementsByName('pw')[0].value='{self.pw}'")
        time.sleep(1)
        self.driver.find_element(By.ID, 'log.login').click()

    def get_image_sizes(self, images):
        if self.img_width > -1:
            return [(int(self.img_width), int(self.img_width * x.size[1] / x.size[0])) for x in images]
        else:
            return [(x.size[0], x.size[1]) for x in images]

    def create_blank_canvas(self, image_sizes):
        widths, heights = zip(*(image_sizes))
        max_width, total_height = max(widths), sum(heights)
        if self.img_space > 0:
            total_height += (self.img_space * (len(image_sizes) - 1))

        return Image.new("RGB", (max_width, total_height), (255, 255, 255))

    def paste_images_onto_canvas(self, images, canvas, image_sizes):
        y_offset = 0
        for idx, img in enumerate(images):
            if self.img_width > -1:
                img = img.resize(image_sizes[idx])

            canvas.paste(img, (0, y_offset))
            y_offset += (img.size[1] + self.img_space)

    def merge_image(self, images_paths, title_name):
        images = [Image.open(os.path.join(self.image_directory, x.strip())) for x in images_paths]
        image_sizes = self.get_image_sizes(images)
        result_img = self.create_blank_canvas(image_sizes)
        self.paste_images_onto_canvas(images, result_img, image_sizes)

        file_name = f"{title_name}.{self.img_format}"  # File name is body title + image format
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

    def get_category(self, category):
        #TODO 원하는 카테고리 추가해주면됨
        categories = {
            '공연': 'menuLink1285',
            '연극': 'menuLink1285',
            '영화': 'menuLink1285',
            '스포츠': 'menuLink1286',
            '남성패션': 'menuLink358',
            '남성잡화': 'menuLink358'
        }
        return categories.get(category)

    def switch_to_frame_and_write(self):
        self.driver.switch_to.frame('cafe_main')  # 중요 Iframe 을 바꿔야해
        self.driver.find_element(By.ID, 'writeFormBtn').click()  # 글쓰기

    def input_product_name_and_price(self, product_name, sale_price):
        self.driver.switch_to.window(self.driver.window_handles[-1])
        time.sleep(1)
        self.driver.find_element(By.CLASS_NAME, 'textarea_input').send_keys(product_name)
        time.sleep(0.5)
        self.driver.find_element(By.CLASS_NAME, 'input_text').send_keys(int(sale_price))

    def input_quality_status(self, product_condition):
        quality_status_mapping = {
            '미개봉' : 1,
            '거의새것': 2,
            '사용감있음': 3,
        }
        quality_status = quality_status_mapping.get(product_condition)
        self.driver.find_element(By.ID, f'quality{quality_status}').click()  # 상품 상태 클릭

    def input_payment_and_delivery_info(self, delivery_method):
        # 직접결제
        payment_type = self.driver.find_elements(By.XPATH, '//div[@class="deal_item"]')[1].click()
        naver_pay_type = self.driver.find_element(By.XPATH, '//label[@class="switch_slider"]').click()

        # 배송방법
        delivery_status_mapping={
            '직거래':0,
            '택배거래': 1,
            '온라인전송': 2,
        }
        delivery_status = delivery_status_mapping.get(delivery_method)
        self.driver.find_element(By.XPATH,
                                 f'//label[@for="delivery{delivery_status}"]').click()  # input 을 클릭할 수 없어서 label 을 대신

        # 셀러 번호 공개
        self.driver.find_element(By.XPATH, '//label[@for="agree1"]').click()  # 휴대전화번호 노출 동의
        self.driver.find_element(By.XPATH, '//label[@for="agree2"]').click()  # 안심번호 이용

    def attach_image(self,image_paths, product_name):
        self.driver.find_element(By.XPATH, '//button[@data-name="image"]').click()  # 이미지 첨부 클릭
        time.sleep(2)
        os.system(
            '''osascript -e 'tell application "System Events" to keystroke "w" using command down' ''')  # 첨부파일 창 닫기
        time.sleep(1)

        # 저장된 중고 완성 사진 폴더에서 사진 확장자들 모두 불러오기
        file_with_000 = [f for f in os.listdir(self.save_directory) if
                         product_name in f and f.endswith((".png", ".jpg", ".jpeg"))]
        file_found = False
        for file in file_with_000:
            full_file_path = self.find_files(file, self.save_directory)
            for f in full_file_path:
                print(f)
                file_found = True

        if file_found:  # 만들어진 파일이 있으면
            image = os.path.join(self.save_directory, product_name + '.' + self.img_format)
            file_found = False
        else:  # 없으면
            image = self.merge_image(image_paths, product_name)

        input_file_element = self.driver.find_element(By.ID, 'hidden-file')
        input_file_element.send_keys(image)

        time.sleep(random.uniform(8, 10))

    def post_article(self):
        no_up = 0
        for i, row in self.df.iterrows():   #행을 돌게
            category_element_id = self.get_category(row['카테고리'])
            self.driver.find_element(By.ID,category_element_id).click()

            self.switch_to_frame_and_write()
            self.input_product_name_and_price(row['상품명'], row['판매가격'])
            self.input_quality_status(row['상품상태'])
            self.input_payment_and_delivery_info(row['배송방법'])
            self.attach_image(row['이미지'].split(" "), row['상품명'])

            # 본문
            webdriver.ActionChains(self.driver).send_keys(self.df.loc[i, '본문']).perform()
            time.sleep(random.uniform(4, 8))
            # 작성 완료
            self.driver.find_element(By.XPATH, '//a[@role="button"][contains(@class, "BaseButton")]').click()
            time.sleep(random.uniform(3, 5))
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(random.uniform(3, 5))

    def run(self):
        self.login()
        for _ in range(len(self.df.index)):
            self.post_article()
            time.sleep(random.uniform(20, 50))


def main():
    naver_id = myIdPW.naver_id
    naver_pw = myIdPW.naver_pw
    poster = NaverAutoPoster(naver_id, naver_pw)
    poster.run()

    schedule.every(20).minutes.do(poster.run)   #20분 마다 반복

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    main()