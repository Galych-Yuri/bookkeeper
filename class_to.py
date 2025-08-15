import os
from logs_to import class_logger
from manifest_menu import out
from print_to import print_talon, close_chrome
from viber_starter import ViberSender
from email_send import send_email_with_pdf

import gspread
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheetsManager:
    """Клас для роботи з Google Sheets"""

    def __init__(self, sheet_name):
        # Логування інформаційного повідомлення
        class_logger.info(f"Підключення до Google Sheets: {sheet_name}.")
        self.sheet = self.connect_to_sheet(sheet_name)

    @staticmethod
    def connect_to_sheet(sheet_name):
        try:
            scope = ["https://spreadsheets.google.com/feeds",
                     "https://www.googleapis.com/auth/spreadsheets"]
            # Отримати шлях до виконуваного файлу
            # Тепер використовуємо json_path для відкриття файлу
            current_dir = os.path.dirname(os.path.abspath(__file__))
            json_path = os.path.join(current_dir,
                                     'ВАШ КЛЮЧ УПРАВЛІННЯ АККАУНТОМ.json')

            creds = ServiceAccountCredentials.from_json_keyfile_name(
                json_path, scope)

            client = gspread.authorize(creds)

            spreadsheet_id = "АДРЕСА ТАБЛИЦІ"

            spreadsheet = client.open_by_key(spreadsheet_id)
            class_logger.info(
                f"Успішне підключення до таблиці '{sheet_name}'.")
            return spreadsheet.worksheet(sheet_name)
        except Exception as e:
            class_logger.error(
                f"Помилка підключення до '{sheet_name}' Google Sheets: {e}")
            print(f"Помилка підключення до '{sheet_name}' Google Sheets: {e}")

    def get_all_records(self):
        try:
            class_logger.info(f"Отримую дані з таблиці: {self.sheet}")
            return self.sheet.get_all_records()
        except Exception as e:
            class_logger.error(f"Помилка при отриманні даних з таблиці: {e}")
            print(f"Помилка при отриманні даних з таблиці: {e}")

    def batch_update(self, updates):
        self.sheet.batch_update(updates)

    def update_cell(self, values, cell_range):
        self.sheet.update(values, cell_range)

    def clear_ranges(self, ranges):
        self.sheet.batch_clear(ranges)

    def insert_row_sheet(self, values, index):
        self.sheet.insert_row(values, index)

    def col_values(self, col):
        self.sheet.col_values(col)

    def get_last_filled_row(self):
        """
        Метод для знаходження останнього заповненого рядка.
        :return: Індекс останнього заповненого рядка.
        """
        data = self.sheet.get_all_values()  # Отримуємо всі значення з таблиці
        return len(data)  # Кількість рядків і є останнім заповненим рядком

    def append_row_below_last(self, values):
        """
        Метод для вставки нових рядків під останнім заповненим.
        :param values: Список значень, які будуть вставлені.
        """
        last_row = self.get_last_filled_row()  # останній заповнений рядок
        next_row_index = last_row + 1

        # Вставка кожного рядка з values
        for row in values:
            self.sheet.insert_row(row, next_row_index)
            next_row_index += 1  # Переходимо до наступного рядка


class SubabonentsManager:
    """Клас для управління субспоживачами"""

    def __init__(self, worksheet, today_date):
        self.worksheet = worksheet
        self.today_date = today_date

    @staticmethod
    def prepare_for_printing(row):
        b9_b10 = [
            [row["№ квитанції"]],
            [row["ПІБ"]]
        ]
        a14_b14 = [
            [row["Період з"], row["Період по"]]
        ]
        a16_b16 = [
            [row["Попередні покази"], row["Поточні покази"]]
        ]
        e14 = [[row["Борг"]]] if row["Борг"] else [[""]]
        c16 = [[5.616]] if row["ПІБ"] == "ФО Вегнер Макс" else [[4.968]]
        e16 = [[row["Сума"]]]

        for_printing = [
            {"range": "B9:B10", "values": b9_b10},
            {"range": "A14:B14", "values": a14_b14},
            {"range": "A16:B16", "values": a16_b16},
            {"range": "C16", "values": c16},
            {"range": "E16", "values": e16}
        ]

        if row["Борг"]:
            for_printing.append({"range": "E14", "values": e14})

        return for_printing

    def process_subabonents(self, data, electro_talon):
        # choice_system = input(
        #     "Вкажи яка система використовується Mac чи Win? ").strip().lower()
        choice_system = 'mac'

        # можливо цей цикл буде краще за enumerate
        for row_idx, row in zip(range(len(data) + 1, 1, -1), reversed(data)):
            if row["Друкована?"] == "":
                updates = self.prepare_for_printing(row)
                class_logger.info(f"Створено квитанцію: {row} \n{updates}.")
                self.worksheet.batch_update(updates)
                print_talon(choice_system)

                new_name = [row['ПІБ'],
                            row['Період з'],
                            row['Період по'],
                            row['Сума']
                            ]
                new_name = '_'.join(map(str, new_name))

                self.find_and_rename_latest_file(new_name=new_name)
                pdf_file = self.find_latest_file()
                send_email_with_pdf(pdf_file, new_name)

                # ViberSender.send_file('КОМУ ВІДСИЛАЄМО ФАЙЛ')
                
                user_input = input(
                    f"Ти надрукував квитанцію {row_idx - 1}:{row['ПІБ']}?\n"
                    f"Натисни RETURN щоб продовжити друк або 'q' щоб вийти: "
                )
                
                user_input.strip().lower()

                if user_input in out:
                    break
                else:
                    electro_talon.update_cell([["так"]], f"N{row_idx}")
                    electro_talon.update_cell([[self.today_date]],
                                              f"M{row_idx}")
                    class_logger.info(
                        f"Успішно роздрукована: {row}. Статус квитанції "
                        f"'друкована?' змінено на ТАК.")
                    self.worksheet.clear_ranges(
                        ["B9:B10", "A14:B14", "A16:B16", "C16", "E14", "E16"]
                    )
        else:
            print("\nНема шо дуркувати!")
            close_chrome()

    def _get_last_talon(self, talon_data):
        """Метод для повернення номера останньої квитанції."""
        last_num_talon = None
        for iteration in reversed(talon_data):
            if iteration["№ квитанції"]:
                last_num_talon = iteration["№ квитанції"]
                break
        return last_num_talon

    def _get_last_abonent_talon_debt_overnt(self, row, talon_data):
        """
        Шукаємо останню сплачену (виставлений рахунок) квитанцію абонента.
        :param row: Dictionary
        :param talon_data: GoogleSheetsManager("Квитанції за електрохарчування")
        :return: Повертаємо борг та переплату до функції create_talon
        """
        debt, overpayment = 0, 0
        for last_talon in reversed(talon_data):
            if row["Споживач"] == last_talon["ПІБ"]:
                if last_talon["Борг"]:
                    debt += int(last_talon["Борг"])
                if last_talon["Переплата"]:
                    overpayment += int(last_talon["Переплата"])
                break
        return debt, overpayment

    def _calculate_overpayment_and_amount(self, overpayment, amount):
        """Метод для обробки переплати і суми."""
        overpayment = int(overpayment)
        amount = int(amount)

        if overpayment >= amount:
            overpayment -= amount
            amount = 0
        else:
            amount -= overpayment
            overpayment = 0  # Вичерпали переплату
        return overpayment, amount

    def create_talon(self, data_abonents, talon_data):
        """
        Метод для формування платіжного талона. Це інформація яка міститься
        в таблиці "Квитанції за електрохарчування"
        :param data_abonents: GoogleSheetsManager("Субспоживачі_1")
        :param talon_data: GoogleSheetsManager("Квитанції за електрохарчування")
        :return: [ПІБ,НОМЕР,ДАТА З,ДАТА ПО,1003,0,0,0,17598...]
        """
        result = []

        # Створюємо лічильник номера талона
        number_talon = self._get_last_talon(talon_data)

        for row in data_abonents:

            if row["Споживач"] == "Ворота":
                break  # Якщо споживач "Ворота", зупиняємо формування квитанцій

            # Отримуємо останній і передостанній показники лічильника
            last_key, last_value = list(row.items())[-1]
            penultimate_key, penultimate_value = list(row.items())[-2]

            if (last_value not in [' ', '0', 0] and
                    penultimate_value not in [' ', '0', 0]):
                # Перевірка на талон що існує з тим самим показником
                for last_talon in reversed(talon_data):
                    if (row["Споживач"] == last_talon["ПІБ"] and
                            last_value == last_talon["Поточні покази"]):
                        break  # Пропускаємо, якщо вже є талон з цими показниками
                else:
                    # Створюємо новий талон, якщо показники нові
                    number_talon += 1
                    abonent_name = row["Споживач"]

                    debt, overpayment = (
                        self._get_last_abonent_talon_debt_overnt(row,
                                                                 talon_data)
                    )  # Вираховуємо чи є борг або переплата

                    difference = int(last_value) - int(penultimate_value)

                    # Розрахунок суми залежно від абонента
                    amount = difference * 5.616 \
                        if abonent_name == "ФО Вегнер Макс" \
                        else difference * 4.968

                    # Оновлення переплати та суми
                    overpayment, amount = (
                        self._calculate_overpayment_and_amount(overpayment,
                                                               amount))
                    overpayment = f"{round(overpayment)}"
                    amount = f"{round(amount)}"

                    talon_entry = [
                        abonent_name, str(number_talon), penultimate_key,
                        last_key, amount, "0", str(debt), overpayment,
                        penultimate_value, last_value, difference
                    ]

                    result.append(talon_entry)
                    class_logger.info(f"Створено талон: {talon_entry}")

        if result:
            print("\nЧекай, я вже переношу квитанції через дорогу за ручку.")
        return result

    def find_latest_file(self, folder_path=os.path.expanduser("~/Downloads")):
        """
        Повертає повний шлях до останнього файлу в теці.
        """
        try:
            files = [os.path.join(folder_path, f)
                     for f in os.listdir(folder_path)
                     if os.path.isfile(os.path.join(folder_path, f))]

            if not files:
                print("У папці немає файлів.")
                return None

            latest_file = max(files, key=os.path.getctime)
            print(f"Останній доданий файл: \n{latest_file}")
            return latest_file

        except Exception as e:
            print(f"Помилка при пошуку файлів: {e}")
            return None

    def find_and_rename_latest_file(self, new_name,
                                    folder_path=os.path.expanduser(
                                        "~/Downloads")):
        """
        Потрібно скоротити та вирізати код який шукає останній файл тому що вже
        є окремий для цього метод вище
        :param new_name:
        :param folder_path:
        :return:
        """
        try:
            # Отримати список файлів у теці
            files = [os.path.join(folder_path, f)
                     for f in os.listdir(folder_path)
                     if os.path.isfile(os.path.join(folder_path, f))
                     ]

            if not files:
                print("У папці немає потрібних файлів.")
                return

            # Знайти останній доданий файл
            latest_file = max(files, key=os.path.getctime)
            print(f"Останній доданий файл: \n{latest_file}")

            # Витягнути розширення файлу
            file_extension = os.path.splitext(latest_file)[1]

            # Сформувати новий шлях для файлу
            new_file_path = os.path.join(folder_path,
                                         new_name + file_extension)

            # Перейменувати файл
            os.rename(latest_file, new_file_path)
            print(f"Файл перейменовано на: \n{new_file_path}")

        except Exception as e:
            print(f"Виникла помилка: {e}")
