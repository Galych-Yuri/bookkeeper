import pyautogui
import time
import subprocess
import pyperclip


class ViberSender:
    """
    Class representing a Viber connection
    """

    @staticmethod
    def open_viber():
        """Відкриває Viber."""
        subprocess.run(["open", "/Applications/Viber.app"])  # macOS
        time.sleep(10)  # Чекаємо, поки додаток завантажиться

    @staticmethod
    def is_running(app_name="Viber"):
        """Перевіряє, чи запущений Viber."""
        script = f"""
            tell application "System Events"
                set appList to (name of every process)
            end tell
            return appList contains "{app_name}"
            """
        result = subprocess.run(["osascript", "-e", script],
                                capture_output=True, text=True)
        print(result.stdout)

        return "true" in result.stdout.lower()

    @staticmethod
    def activate(app_name="Viber"):
        """Активує Viber, якщо він запущений."""
        script = f"""
            if application "{app_name}" is running then
                tell application "{app_name}" to activate
            end if
            """
        subprocess.run(["osascript", "-e", script])
        time.sleep(2.0)

    @staticmethod
    def close_viber() -> None:
        """
        Закриває Viber, якщо він відкритий.
        :return:
        """
        try:
            # Використовуємо pkill для завершення процесу Viber
            subprocess.run(["pkill", "-f", "Viber"], check=True)
        except subprocess.CalledProcessError:
            print("Не вдалося закрити Viber. Переконайтеся, що він запущений.")

    @staticmethod
    def send_file(contact_name):
        """Надсилає файл контакту у Viber."""

        # Словник координат з описами
        clicks = {
            "Відкрити пошук абонента": (384, 91),
            "Вставити контакт": contact_name,
            "Вибір контакту": (383, 242),
            "Відкрити вибір файлів": (541, 731),
            "Обрати папку 'Завантаження'": (351, 456),
            "dropdown_list": (525, 201),
            "choice_list": (541, 316),
            "dropdown_group": (579, 202),
            "date_added": (602, 368)
        }

        # Виконати скорочену кількість дій
        if ViberSender.is_running():
            ViberSender.activate()
            pyautogui.click(*clicks["Відкрити вибір файлів"])
            time.sleep(1.5)
        else:
            # Виконати повний список дій
            ViberSender.open_viber()
            for action, value in clicks.items():
                if isinstance(value, tuple):  # Якщо це координати
                    pyautogui.click(*value)
                    time.sleep(1.5)  # Час очікування між кліками
                elif action == "Вставити контакт":  # Особливий випадок
                    pyperclip.copy(contact_name)
                    time.sleep(0.5)
                    pyautogui.hotkey('command', 'v')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(1.5)

        # Інші дії
        pyautogui.press('down')  # Натиснути стрілку вниз
        time.sleep(1.5)
        pyautogui.press('enter')  # Надіслати файл
        time.sleep(9)

        print(f"Надіслав файл контакту: {contact_name}.")
        # ViberSender.close_viber()


class NoFileOpen:
    pass
