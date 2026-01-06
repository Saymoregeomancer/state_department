from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager




def get_last_cell_text(url: str):
    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # розкоментуй, якщо не хочеш бачити браузер
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        driver.get(url)
        try:
            # Чекаємо до 15 секунд, поки таблиця завантажиться
            WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.ecl-table__row"))
            )

            js_code = """
            let data = document.querySelectorAll('tr.ecl-table__row');
            let arr = [];
            data?.forEach(el => {
                let elem = el.querySelectorAll('.ecl-table__cell')[1]?.innerText;
                arr.push(elem);
            });
            return arr.join('|') || null;
            """
            result = driver.execute_script(js_code)
        except:
            result = None  # якщо таблиця не завантажилась
        return result
    finally:
        driver.quit()
