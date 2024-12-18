import csv
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def init_csv():
    """
    Tạo file CSV cùng cấp với file main.py với tiêu đề các cột.
    """
    # Đảm bảo thư mục 'report' tồn tại
    report_folder = 'report'
    if not os.path.exists(report_folder):
        os.makedirs(report_folder)
        print(f"Đã tạo thư mục: {report_folder}")

    # Đặt tên file với ngày và tháng
    now = datetime.now()
    csv_file = os.path.join(report_folder, f"Report_{now.strftime('%d-%m')}.csv")

    # Tạo file CSV với tiêu đề cột
    with open(csv_file, mode='w', encoding='utf-8-sig', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Tên TGDD', 'Giá TGDD', 'URL TGDD', '', 'Tên CellphoneS', 'Giá CellphoneS', 'URL CellphoneS'])
    
    print(f"Đã tạo file CSV: {csv_file}")
    return csv_file

def save_to_csv(csv_file, data, start_column):
    """
    Ghi dữ liệu vào file CSV tại các cột bắt đầu từ start_column.
    """
    with open(csv_file, mode='a', encoding='utf-8-sig', newline='') as file:
        writer = csv.writer(file)
        for row in data:
            row_data = [''] * start_column + row
            writer.writerow(row_data)

def open_tgdd_page():
    """
    Mở trang Thegioididong và đặt trình duyệt ở nửa màn hình bên trái.
    """
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)

    screen_width = driver.execute_script("return window.screen.width")
    screen_height = driver.execute_script("return window.screen.height")

    driver.get("https://www.thegioididong.com")
    driver.set_window_position(0, 0)
    driver.set_window_size(screen_width // 2, screen_height)

    return driver

def get_data_tgdd(driver, csv_file):
    """
    Lấy dữ liệu từ Thegioididong và ghi vào file CSV (cột A, B, C).
    """
    driver.find_element(By.ID, 'skw').send_keys('iphone 16 promax')
    driver.find_element(By.XPATH, "//button[i[@class='icon-search']]").click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, "//li[contains(@class, 'item cat42')]"))
    )

    products = driver.find_elements(By.XPATH, "//li[contains(@class, 'item cat42')]")
    data = []

    for product in products:
        try:
            product_name = product.find_element(By.XPATH, ".//h3").text.strip()
            product_price = product.find_element(By.XPATH, ".//strong[@class='price']").text.strip()
            product_url = product.find_element(By.XPATH, ".//a").get_attribute("href")
            data.append([product_name, product_price, product_url])
            print(f"Đã ghi TGDD: {product_name}, {product_price}, {product_url}")
        except Exception as e:
            print(f"Lỗi khi xử lý sản phẩm TGDD: {e}")

    save_to_csv(csv_file, data, start_column=0)

def open_cellphone_page():
    """
    Mở trang CellphoneS và đặt trình duyệt ở nửa màn hình bên phải.
    """
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)

    screen_width = driver.execute_script("return window.screen.width")
    screen_height = driver.execute_script("return window.screen.height")

    driver.get("https://cellphones.com.vn")
    driver.set_window_position(screen_width // 2, 0)
    driver.set_window_size(screen_width // 2, screen_height)

    return driver


def get_data_cellphone(driver, csv_file):
    """
    Lấy dữ liệu từ CellphoneS và ghi vào file CSV (cột E, F, G).
    """
    driver.find_element(By.ID, 'inp$earch').send_keys('iphone 16 promax')
    driver.find_element(By.XPATH, "//div[@class='input-group-btn']").click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//div[@class="product-info-container product-item"]'))
    )

    products = driver.find_elements(By.XPATH, '//div[@class="product-info-container product-item"]')
    data = []

    for product in products:
        try:
            product_name = product.find_element(By.XPATH, ".//div[@class='product__name']/h3").text.strip()
            product_price = product.find_element(By.XPATH, './/p[@class="product__price--show"]').text.strip()
            product_url = product.find_element(By.XPATH, ".//a[@class='product__link button__link']").get_attribute("href")
            data.append([product_name, product_price, product_url])
            print(f"Đã ghi CellphoneS: {product_name}, {product_price}, {product_url}")
        except Exception as e:
            print(f"Lỗi khi xử lý sản phẩm CellphoneS: {e}")

    save_to_csv(csv_file, data, start_column=4)