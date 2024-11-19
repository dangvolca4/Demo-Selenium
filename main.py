from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unit

# Đường dẫn file Excel template
template_file = r'D:\Selenium\Rp So sanh.xlsx'
template_sheet_name = 'template'

# Khởi tạo Excel
workbook, new_sheet = unit.setup_excel(template_file, template_sheet_name)

# Khởi tạo trình duyệt
driver1 = unit.setup_browser()
driver2 = unit.setup_browser()

# Chia đôi màn hình
unit.split_screen(driver1, driver2)

# Mở trang web
driver1.get("https://www.thegioididong.com")
driver2.get("https://cellphones.com.vn")

# Xử lý trang Thegioididong
driver1.find_element(By.ID, 'skw').send_keys('iphone 16 promax')
driver1.find_element(By.XPATH, "//button[i[@class='icon-search']]").click()

WebDriverWait(driver1, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, "//li[contains(@class, 'item cat42')]"))
)

products1 = driver1.find_elements(By.XPATH, "//li[contains(@class, 'item cat42')]")
data1 = []
for product in products1:
    try:
        name = product.find_element(By.XPATH, ".//h3").text.strip()
        price = product.find_element(By.XPATH, ".//strong[@class='price']").text.strip()
        url = product.find_element(By.XPATH, ".//a").get_attribute("href")
        data1.append((name, price, url))
    except Exception as e:
        print(f"Lỗi khi xử lý sản phẩm Thegioididong: {e}")

# Ghi dữ liệu Thegioididong vào Excel
unit.write_to_excel(new_sheet, start_row=3, start_column=1, data=data1)

# Xử lý trang CellphoneS
driver2.find_element(By.ID, 'inp$earch').send_keys('iphone 16 promax')
driver2.find_element(By.XPATH, "//div[@class='input-group-btn']").click()

WebDriverWait(driver2, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, '//div[@class="product-info-container product-item"]'))
)

products2 = driver2.find_elements(By.XPATH, '//div[@class="product-info-container product-item"]')
data2 = []
for product in products2:
    try:
        name = product.find_element(By.XPATH, ".//div[@class='product__name']/h3").text.strip()
        price = product.find_element(By.XPATH, './/p[@class="product__price--show"]').text.strip()
        url = product.find_element(By.XPATH, ".//a[@class='product__link button__link']").get_attribute("href")
        data2.append((name, price, url))
    except Exception as e:
        print(f"Lỗi khi xử lý sản phẩm CellphoneS: {e}")

# Ghi dữ liệu CellphoneS vào Excel
unit.write_to_excel(new_sheet, start_row=3, start_column=6, data=data2)

# Đóng trình duyệt
driver1.quit()
driver2.quit()

# Lưu workbook vào file Excel
workbook.save(template_file)
print(f"Đã lưu dữ liệu vào file Excel: {template_file}")
