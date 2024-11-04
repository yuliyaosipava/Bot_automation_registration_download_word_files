import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Установите ваш логин и пароль здесь
username = 'kemp@kemp.by'
password = '*****'

# Создание экземпляра ChromeDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Путь для сохранения отчетов
output_dir = 'c:\\Users\\user\\Documents\\reports\\'
os.makedirs(output_dir, exist_ok=True)

try:
    # Открытие сайта
    driver.get('https://ilex.by/login/')
    print("Открыт сайт.")

    # Вход на сайт
    login_field = driver.find_element(By.ID, 'username-input')
    login_field.send_keys(username)
    password_field = driver.find_element(By.ID, 'password-input')
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)
    print("Введены логин и пароль.")

    # Ожидание и проверка на наличие и видимость кнопки "Продолжить"
# Ожидание и проверка на наличие и видимость кнопки "Продолжить"
    try:
        time.sleep(5)  # Добавление задержки
        wait = WebDriverWait(driver, 45)
        continue_button = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[text()="Продолжить"]')))
        continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Продолжить"]')))
        continue_button.click()
        print("Нажата кнопка 'Продолжить'.")
    except TimeoutException:
        print("Кнопка 'Продолжить' не найдена или не кликабельна.")
        driver.save_screenshot("screenshot.png")  # Создание скриншота для отладки


    # Ожидание перенаправления на новую страницу
    wait.until(EC.url_contains("https://ilex-private.ilex.by/home"))
    print("Перенаправление на главную страницу.")

    # Ожидание загрузки строки поиска на новой странице
    search_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchAutocompleteInput"]')))
    print("Найдена строка поиска.")

    # Ввод текста в строку поиска
    search_field.send_keys('об исполнении областного бюджета за 2023')
    search_field.send_keys(Keys.RETURN)
    print("Запрос введен в строку поиска.")

    time.sleep(10)  # Подождите некоторое время для завершения поиска

    # Переход на вторую страницу с результатами
    next_button = driver.find_element(By.CSS_SELECTOR, 'button.mat-paginator-navigation-next')
    next_button.click()
    time.sleep(5)  # Подождите некоторое время для загрузки второй страницы

    while True:
        # Получение ссылок на отчеты
        report_links = driver.find_elements(By.XPATH, '//a[contains(@href, "view-document/BEMLAW")]')

        # Скачивание отчетов
        for link in report_links:
            report_url = link.get_attribute('href')
            report_name = link.text.strip() + '.docx'
            print(f"Обнаружена ссылка: {report_url}")

            # Клик по ссылке и переход в новое окно
            link.click()
            time.sleep(5)

            # Переключение на новое окно
            driver.switch_to.window(driver.window_handles[-1])

            # Нажатие кнопки для экспорта в Word
            export_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.icon > i.export-word')))
            export_button.click()
            time.sleep(10)  # Подождите некоторое время для завершения экспорта

            # Перемещение скачанного файла в нужную папку
            downloaded_file = os.path.join('c:\\Users\\user\\Downloads', 'filename.docx')  # Убедитесь, что заменили 'filename.docx' на реальное имя скачанного файла
            if os.path.exists(downloaded_file):
                os.rename(downloaded_file, os.path.join(output_dir, report_name))
                print(f"Скачанный файл перемещен: {report_name}")
            else:
                print(f"Не удалось найти скачанный файл: {report_name}")

            # Закрытие текущего окна и переключение обратно
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Переход к следующей странице
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'button.mat-paginator-navigation-next')
            next_button.click()
            time.sleep(5)  # Подождите некоторое время для загрузки следующей страницы
        except:
            print("Больше страниц нет.")
            break

finally:
    # Закрытие браузера
    driver.quit()
    print("Браузер закрыт.")
