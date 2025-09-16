import time
import pandas as pd
import json
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
import ctypes
# جلوگیری از sleep سیستم
ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

# بررسی اتصال اینترنت
def is_connected():
    try:
        requests.get('https://www.google.com', timeout=5)
        return True
    except requests.ConnectionError:
        return False

# ذخیره وضعیت پردازش در فایل
def save_progress(category_idx):
    with open('progress.json', 'w') as f:
        json.dump({'category_idx': category_idx}, f)

# بارگذاری وضعیت پردازش از فایل
def load_progress():
    try:
        with open('progress.json', 'r') as f:
            progress = json.load(f)
            return progress.get('category_idx', 0)  # بازگشت به دسته‌ای که از آن شروع شده
    except FileNotFoundError:
        return 0  # اگر هیچ پیشرفتی ذخیره نشده باشد، از ابتدا شروع می‌کنیم

# راه‌اندازی مرورگر
def setup_driver(chromedriver_path):
    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service)
    return driver

# اسکرول صفحه به وسط
def simple_scroll_to_middle(driver):
    total_height = driver.execute_script("return document.body.scrollHeight")
    target_pos = total_height / 2
    driver.execute_script(f"window.scrollTo(0, {target_pos});")
    time.sleep(1.5)

# استخراج نام زیر دسته
def extract_subcategory_name(driver, wait):
    attempts = 0
    max_attempts = 3
    while attempts < max_attempts:
        try:
            wrapper = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.callery_calculate_wraper")))
            h4 = wrapper.find_element(By.TAG_NAME, "h4")
            name = h4.text.strip()
            if name:
                print(f"نام زیر دسته استخراج شد: {name}")
                return name
            else:
                print("نام زیر دسته خالی است، تلاش مجدد...")
                attempts += 1
                time.sleep(0.5)
        except Exception as e:
            print(f"خطا در استخراج نام زیر دسته: {e}")
            attempts += 1
            time.sleep(0.5)
    print("نام زیر دسته پس از چند تلاش خالی ماند، مقدار 'نامشخص' قرار داده شد.")
    return "نامشخص"

# استخراج داده‌های مودال
def extract_modal_data(driver, wait):
    results = []
    try:
        subcat_name = extract_subcategory_name(driver, wait)

        callery_items = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "callery_item")))

        for item in callery_items:
            try:
                item_name = item.find_element(By.XPATH, ".//span[not(@class)]").text
            except:
                item_name = "نامشخص"
            try:
                item_value_str = item.find_element(By.CLASS_NAME, "callery_item_value").text
                item_value_float = float(item_value_str)
            except:
                item_value_float = 0.0

            print(f"ویژگی : {item_name}  , مقدار : {item_value_float}")
            results.append((item_name, item_value_float))
            time.sleep(0.3)

        close_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "close_btn")))
        close_btn.click()
        print("مودال بسته شد")
        time.sleep(1)

    except Exception as e:
        print(f"خطا در استخراج یا بستن مودال: {e}")

    return subcat_name, results

# کلیک کردن روی زیر دسته‌ها و استخراج داده‌ها
def click_all_subcategories_and_extract(driver, wait, category_box, category_idx, category_name):
    all_results = []

    subcategories = category_box.find_elements(By.CLASS_NAME, "callery_product")
    print(f"تعداد زیر دسته‌ها برای دسته {category_idx + 1}: {len(subcategories)}")

    idx = 0
    max_retry = 5

    while idx < len(subcategories):
        retry = 0
        success = False
        while retry < max_retry and not success:
            try:
                subcategories = category_box.find_elements(By.CLASS_NAME, "callery_product")
                subcat = subcategories[idx]

                driver.execute_script("""
                    arguments[0].scrollIntoView({behavior: 'smooth', block: 'center', inline: 'nearest'});
                """, subcat)
                time.sleep(0.5)

                wait.until(EC.element_to_be_clickable(subcat))

                try:
                    subcat.click()
                    time.sleep(3)
                except Exception:
                    print("کلیک مستقیم موفق نبود، استفاده از جاوااسکریپت برای کلیک")
                    driver.execute_script("arguments[0].click();", subcat)
                    time.sleep(3)

                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.callery_calculate_wraper")))

                subcat_name, modal_data = extract_modal_data(driver, wait)

                all_results.append({
                    "category_index": category_idx + 1,
                    "category_name": category_name,
                    "subcategory_index": idx + 1,
                    "subcategory_name": subcat_name,
                    "data": modal_data
                })

                idx += 1
                success = True

            except (StaleElementReferenceException, ElementClickInterceptedException) as e:
                retry += 1
                print(f"خطا {e} در زیر دسته {idx + 1}، تلاش مجدد {retry}/{max_retry}")
                time.sleep(1)
            except Exception as e:
                import traceback
                print(f"خطای غیرمنتظره در زیر دسته {idx + 1}: {repr(e)}")
                traceback.print_exc()
                idx += 1
                success = True

        if not success:
            print(f"تلاش برای کلیک روی زیر دسته {idx + 1} ناموفق بود، ادامه می‌دهیم.")
            idx += 1

    # بستن دسته بعد از پایان زیر دسته‌ها (کلیک روی عنوان دسته)
    try:
        category_title_element = category_box.find_element(By.CLASS_NAME, "callery_posts_box_title")
        driver.execute_script("""
            arguments[0].scrollIntoView({behavior: 'smooth', block: 'center', inline: 'nearest'});
        """, category_title_element)
        time.sleep(0.3)
        category_title_element.click()
        print(f"دسته {category_idx + 1} بسته شد (زیر دسته‌ها جمع شدند).")
        time.sleep(0.5)
    except Exception as e:
        print(f"خطا در بستن دسته {category_idx + 1}: {e}")

    return all_results

# ذخیره داده‌ها در فایل اکسل
def save_category_data_to_excel(all_results, category_name, category_index):
    safe_name = "".join(c for c in category_name if c.isalnum() or c in (' ', '_')).rstrip()
    filename = f"calories_data_category_{category_index+1}_{safe_name}.xlsx"
    print(f"ذخیره داده‌های دسته {category_index + 1} با نام '{category_name}' در فایل {filename}")

    flat_data = []
    for entry in all_results:
        subcat_idx = entry["subcategory_index"]
        subcat_name = entry["subcategory_name"]
        for name, value in entry["data"]:
            flat_data.append({
                "category_index": category_index + 1,
                "category_name": category_name,
                "subcategory_index": subcat_idx,
                "subcategory_name": subcat_name,
                "feature_name": name,
                "feature_value": value
            })

    if not flat_data:
        flat_data.append({
            "category_index": category_index + 1,
            "category_name": category_name,
            "subcategory_index": None,
            "subcategory_name": None,
            "feature_name": "هیچ داده‌ای یافت نشد",
            "feature_value": None
        })

    df = pd.DataFrame(flat_data)
    df.to_excel(filename, index=False)
    print(f"داده‌ها در فایل {filename} ذخیره شدند.")

# پردازش تمام دسته‌ها
def process_all_categories(driver, wait):
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    simple_scroll_to_middle(driver)
    print("اسکرول به وسط صفحه انجام شد.")

    categories = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "callery_posts_box")))
    idx = load_progress()  # بارگذاری وضعیت پردازش از فایل

    max_retry_category = 3

    while idx < len(categories):
        retry_category = 0
        success_category = False

        while retry_category < max_retry_category and not success_category:
            try:
                categories = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "callery_posts_box")))
                if idx >= len(categories):
                    print("دسته‌ای باقی نمانده یا به انتها رسیدیم. پایان برنامه.")
                    return

                category_box = categories[idx]
                category_name = category_box.find_element(By.CLASS_NAME, "callery_posts_box_title").text.strip()

                driver.execute_script("""
                    arguments[0].scrollIntoView({behavior: 'smooth', block: 'center', inline: 'nearest'});
                """, category_box)
                time.sleep(0.5)

                try:
                    category_box.click()
                except ElementClickInterceptedException:
                    driver.execute_script("window.scrollTo(0, 0);")
                    time.sleep(0.5)
                    category_box.click()

                print(f"کلیک روی دسته {idx + 1} انجام شد: {category_name}")

                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "callery_product")))

                all_results = click_all_subcategories_and_extract(driver, wait, category_box, idx, category_name)

                save_category_data_to_excel(all_results, category_name, idx)

                save_progress(idx)  # ذخیره وضعیت پردازش

                idx += 1
                success_category = True

            except (StaleElementReferenceException, TimeoutException) as e:
                retry_category += 1
                print(f"خطا {e} در دسته {idx + 1}، تلاش مجدد {retry_category}/{max_retry_category}")
                time.sleep(2)
            except Exception as e:
                print(f"خطا در پردازش دسته {idx + 1}: {e}")
                idx += 1
                success_category = True

        if not success_category:
            print(f"عدم موفقیت در پردازش دسته {idx + 1} پس از چند تلاش، ادامه به دسته بعدی.")
            idx += 1

# اجرای کد
def main():
    chromedriver_path = "F:\\project_Calori\\final_webscript\\chromedriver.exe"  # مسیر کروم‌درایور را اینجا تنظیم کن
    driver = setup_driver(chromedriver_path)
    wait = WebDriverWait(driver, 20)

    url = "https://badankhooba.com/calculating-calories/"
    driver.get(url)

    process_all_categories(driver, wait)

    driver.quit()

if __name__ == "__main__":
    main()
