import time
import subprocess
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver

from logs_to import print_logger
from manifest_menu import mac, win

MAC_PRNT_BTN = "div[aria-label='Друк (⌘P)']"


def print_button_finder(system):
    if system in mac:
        print_logger.info(f"Обрано операційну систему {system}.")
        return "/Users/itsamac/Library/Application Support/Google/Chrome/Profile 1"
    elif system in win:
        print_logger.info(f"Обрано операційну систему {system}.")
        return "C:\\Users\\БожеЯкеКончене\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1"
    else:
        raise ValueError(f"Ти вказав не ту систему - '{system}'!")


def system_choice(driver, system):
    if system in mac:
        # Натискаємо кнопку друку для macOS
        print_button = driver.find_element(By.CSS_SELECTOR,
                                           MAC_PRNT_BTN)
        time.sleep(2)
        print_button.click()

        wait = WebDriverWait(driver, 10)  # Чекаємо до 10 секунд
        next_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR,
             "div.docs-material-button-raised-primary.docs-material-"
             "button.waffle-printing-print-button")))
        next_button.click()
    elif system in win:
        pass
    else:
        # Чекаю на інформацію від віндовс-системи
        print_logger.info("Ця функція ще не налаштована для інших ОС.")


def set_options(system):
    options = webdriver.ChromeOptions()

    # Обираємо систему
    user_data_directory = print_button_finder(system)
    options.add_argument(f"user-data-dir={user_data_directory}")
    if system in mac:
        options.add_argument("profile-directory=Profile 1")
    elif system in win:
        options.add_argument("profile-directory=Default")

    # Вмикає JavaScript
    options.add_argument("enable-javascript")

    # Відкриває порт для налагодження
    options.add_argument("--remote-debugging-port=9222")

    options.add_argument("--allow-running-insecure-content")
    options.add_experimental_option(
        "excludeSwitches", ["enable-automation"])
    options.add_experimental_option(
        'useAutomationExtension', False)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # Автоматичне друкування
    options.add_argument("--kiosk-printing")

    # Маскує автоматизацію
    options.add_argument("--disable-blink-features=AutomationControlled")
    # Додаткові налаштування
    options.add_argument("--disable-web-security")

    return options


def wait_for_authorization(driver, system):
    """
    Функція для очікування авторизації без обмеження часу.
    Використовує цикл, щоб перевіряти наявність певного елемента.
    """
    print_logger.info("Очікування на авторизацію...")

    while True:
        try:
            # Перевіряємо, чи присутній певний елемент на сторінці
            # Це може бути якийсь елемент, що з'являється після авторизації
            if system in mac:
                driver.find_element(
                    By.CSS_SELECTOR, "div[aria-label='Друк (⌘P)']"
                )
                print_logger.info("Авторизація пройдена.")
                break
            elif system in win:
                pass
        except NoSuchElementException:
            # Якщо елемент не знайдений, чекаємо ще кілька секунд і повторюємо
            time.sleep(5)  # Інтервал між перевірками


def wait_for_url_change(driver, expected_url, timeout=30):
    """
    Функція для очікування зміни URL.

    :param driver: WebDriver
    :param expected_url: Очікуваний URL після авторизації або іншої дії
    :param timeout: Час очікування (за замовчуванням 30 секунд)
    """
    try:
        WebDriverWait(driver, timeout).until(EC.url_contains(expected_url))
        print_logger.info(f"Успішно досягнуто URL: {expected_url}")
    except NoSuchElementException:
        # Якщо елемент не знайдений, чекаємо ще кілька секунд і повторюємо перевірку
        time.sleep(5)  # Інтервал між перевірками


def is_chrome_running():
    """Перевіряє, чи запущений Chrome."""
    script = """
        tell application "System Events"
            set appList to (name of every process)
        end tell
        return appList contains "Google Chrome"
    """
    result = subprocess.run(["osascript", "-e", script], capture_output=True,
                            text=True)
    return "true" in result.stdout.lower()


def activate_chrome():
    """Робить вікно Chrome активним."""
    script = """
        tell application "Google Chrome" to activate
    """
    subprocess.run(["osascript", "-e", script])


def close_chrome() -> None:
    """
    Закриває chrome, якщо він відкритий.
    :return:
    """
    try:
        # Використовуємо pkill для завершення процесу
        subprocess.run(["pkill", "-f", "Google Chrome"], check=True)
    except subprocess.CalledProcessError:
        print("Не вдалося закрити chrome. Переконайтеся, що він запущений.")


def print_talon(system):
    # Налаштування існуючого профілю Chrome
    print_logger.info("Розпочато процес друку талонів.")
    options = set_options(system)

    # Створення драйвера з обраними опціями
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    # Відкриваємо сторінку
    driver.get(
        "https://docs.google.com/spreadsheets/d"
        "/АДРЕСА СТОРІНКИ/edit?usp=sharing"
    )

    # Очікування авторизації
    wait_for_authorization(driver, system)

    # Очікування зміни URL
    expected_url = "/АДРЕСА СТОРІНКИ/"
    wait_for_url_change(driver, expected_url)

    # Вибір дій на основі ОС
    system_choice(driver, system)

    # Пауза для завершення друку
    time.sleep(5)

    # Закриття Chrome після завершення роботи
    # close_chrome()


