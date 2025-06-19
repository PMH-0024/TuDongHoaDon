from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
# Cấu hình
download_dir = os.path.abspath("hoadon")
os.makedirs(download_dir, exist_ok=True)
def get_driver():
    # Tự động tải và cài đặt ChromeDriver
    service = Service(ChromeDriverManager().install())
    # Cấu hình Chrome
    chrome_options = webdriver.ChromeOptions()  
    # Cài đặt thư mục tải về
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # Khởi tạo trình duyệt
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver
def doc_ma_tra_cuu():
    # Đơn giản: Trả về danh sách mã tra cứu mẫu
    # Bạn có thể thay đổi thành đọc từ file Excel sau
    return [
        {"ten": "Mẫu 1", "mst": "123456789", "ma_tra_cuu": "ABC123"},
        {"ten": "Mẫu 2", "mst": "987654321", "ma_tra_cuu": "XYZ789"}
    ]
def main():
    print("Đang khởi động trình duyệt...")
    driver = get_driver()

    try:
        # Lấy danh sách mã tra cứu
        danh_sach = doc_ma_tra_cuu()
        
        for ncc in danh_sach:
            print(f"\nĐang xử lý: {ncc['ten']} - {ncc['ma_tra_cuu']}")
            
            # Mở trang tra cứu
            driver.get("https://meinvoice.vn/tra-cuu")
            print("Đã mở trang tra cứu")
            
            # Nhập mã tra cứu
            try:
                input_tc = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='mã tra cứu']"))
                )
                input_tc.clear()
                input_tc.send_keys(ncc['ma_tra_cuu'])
                print(f"Đã nhập mã: {ncc['ma_tra_cuu']}")   
                # Tìm và nhấn nút tìm kiếm
                btn_tim = driver.find_element(
                    By.XPATH, 
                    "//button[contains(., 'Tìm') or contains(., 'Tra cứu')]"
                )
                btn_tim.click()
                print("Đã nhấn tìm kiếm") 
                # Đợi 5 giây để xem kết quả
                time.sleep(5) 
                # Kiểm tra xem có thông báo lỗi không
                try:
                    popup = driver.find_element(By.CLASS_NAME, "modal-content")
                    if "không tìm thấy" in popup.text.lower():
                        print("Không tìm thấy hóa đơn")
                        continue
                except:
                    pass
                    
                print("Tìm thấy hóa đơn")
                
                # Thử tải file nếu tìm thấy nút tải
                try:
                    btn_tai = driver.find_element(
                        By.XPATH, 
                        "//button[contains(., 'Tải') or contains(., 'Download')]"
                    )
                    btn_tai.click()
                    print("Đã nhấn tải hóa đơn")
                    time.sleep(3)  # Đợi tải file
                except:
                    print("Không tìm thấy nút tải")
                
            except Exception as e:
                print(f"Lỗi khi xử lý: {e}")
            
            # Đợi giữa các lần truy vấn
            time.sleep(2)        
    except Exception as e:
        print(f"Có lỗi xảy ra: {e}")
    finally:
        input("\nNhấn Enter để đóng trình duyệt...")
        driver.quit()

if __name__ == "__main__":
    print("Bắt đầu chương trình tra cứu hóa đơn")
    main()
    print("\nKết thúc chương trình")