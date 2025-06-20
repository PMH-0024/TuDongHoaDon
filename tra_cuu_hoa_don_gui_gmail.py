# import các thư viện cần thiết
from selenium import webdriver                    
from selenium.webdriver.common.by import By     
from selenium.webdriver.common.keys import Keys   
from selenium.webdriver.support.ui import WebDriverWait  
from selenium.webdriver.support import expected_conditions as EC  

import time       
import os     
# Gửi email
import smtplib   
 # Mã hóa kết nối Gmail
import ssl        
from email.message import EmailMessage 
MA_TRA_CUU_LIST = [
    # "PZH_FWQ4BN3",
    "B1HEIRR8N0WP",
    # "VBHKSL682918"
]
# Thông tin cấu hình
# MA_TRA_CUU = "B1HEIRR8N0WP"               
EMAIL_NHAN = "hp859869@gmail.com"     
EMAIL_GUI = "hp859869@gmail.com"        
MAT_KHAU_UNG_DUNG = "wvqo nsxe oels rqmz" 
FOLDER_DOWNLOAD = "hoadon" 
# Hàm tự động mở web và tải hóa đơn
def tra_cuu_va_tai_hoa_don(ma_tra_cuu):
    # Tạo thư mục nếu chưa có
    os.makedirs(FOLDER_DOWNLOAD, exist_ok=True)
    # Cài đặt trình duyệt Chrome
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # Cài đặt thư mục tải về
    prefs = {
        "download.default_directory": os.path.abspath(FOLDER_DOWNLOAD),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    # Mở trình duyệt
    driver = webdriver.Chrome(options=options)
    driver.get("https://www.meinvoice.vn/tra-cuu")
    try:
        # Đợi đến khi ô nhập mã xuất hiện
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "txtCode"))
        )
        # Nhập mã tra cứu và search
        input_box = driver.find_element(By.NAME, "txtCode")
        input_box.send_keys(ma_tra_cuu)
        input_box.clear  # Nhấn Enter để tìm kiếm
        # input_box.send_keys(Keys.ENTER)
        print("Đang tra cứu mã:", ma_tra_cuu)
        search_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "btnSearchInvoice"))
        )
        search_button.click()
        print("✅ Đã nhấn nút 'Tra cứu'")
        print("Đã nhấn nút tìm kiếm")
        # Đợi nút "Tải hóa đơn" hiện ra và nhấn vào
        download_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Tải hóa đơn')]"))
        )
        time.sleep(10)  # Chờ thêm 1 giây rồi nhấn nút tải
        download_button.click()
        print("Đã nhấn nút tải hóa đơn")
    except Exception as e:
        print("Lỗi khi tải hóa đơn:", e)
    finally:
        # Tắt trình duyệt
        driver.quit()
# Hàm tìm file vừa mới tải xong
def tim_file_moi_nhat(folder):
    print(" Đang chờ file tải xong...")
    timeout = 30    
    start_time = time.time()
    file_path = None
    while time.time() - start_time < timeout:
        # Lấy danh sách file (loại bỏ file đang tải .crdownload)
        files = os.listdir(folder)
        full_paths = [os.path.join(folder, f) for f in files if not f.endswith(".crdownload")]
        # Nếu có file thì lấy file mới nhất
        if full_paths:
            file_path = max(full_paths, key=os.path.getctime)
            break
        time.sleep(1)  # Chờ 1 giây rồi kiểm tra lại
    return file_path
# Hàm gửi email có đính kèm file
def gui_file_ve_gmail(file_path, email_gui, mat_khau_app, email_nhan):
    subject = "Hóa đơn điện tử từ meInvoice"
    body = "Đây là file hóa đơn bạn vừa tra cứu tự động."
    # Tạo nội dung email
    msg = EmailMessage()
    msg["From"] = email_gui
    msg["To"] = email_nhan
    msg["Subject"] = subject
    msg.set_content(body)
    # Đọc nội dung file đính kèm
    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)
    # Gắn file vào email
    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
    # Gửi email qua Gmail bằng kết nối SSL
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(email_gui, mat_khau_app)
        server.send_message(msg)
        print(f"Đã gửi file '{file_name}' về Gmail:", email_nhan)
if __name__ == "__main__":
    print("Bắt đầu tra cứu và gửi hóa đơn...")
    for ma in MA_TRA_CUU_LIST:
        tra_cuu_va_tai_hoa_don(ma)
        file_moi = tim_file_moi_nhat(FOLDER_DOWNLOAD)
        if file_moi:
            print("✅ File đã tải:", file_moi)
            gui_file_ve_gmail(file_moi, EMAIL_GUI, MAT_KHAU_UNG_DUNG, EMAIL_NHAN)
        else:
            print("⚠️ Không tìm thấy file nào để gửi.")
        print("--------------------------------------------------")
