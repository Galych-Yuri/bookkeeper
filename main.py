from datetime import datetime

from manifest_menu import MENU, GREETENGS, out
from class_to import SubabonentsManager, GoogleSheetsManager
from logs_to import main_logger


def main():
    main_logger.info("Початок роботи програми.")

    print("Підключаюся до таблиці...")
    today_date = datetime.today().strftime('%d.%m.%Y')
    electro_talon = GoogleSheetsManager("Квитанції за електрохарчування")
    electro_abonents = GoogleSheetsManager("Субспоживачі_1")
    print_page = GoogleSheetsManager("Матриця на друк")

    print(MENU)
    user_input = input("\nТопаз дай команду: ")

    while user_input not in out:

        if user_input == "1":
            main_logger.info("Початок формування квитанцій.")
            data_talon = electro_talon.get_all_records()
            data_abonents = electro_abonents.get_all_records()
            electro_proceser = SubabonentsManager(electro_abonents, today_date)
            talon_data = electro_proceser.create_talon(data_abonents,
                                                       data_talon)
            if talon_data:
                electro_talon.append_row_below_last(talon_data)
            else:
                print("Немає чого додавати, всі квитанції вже внесені.")

        elif user_input == "2":
            main_logger.info("Обрано опцію друку квитанцій.")
            data_to_print = electro_talon.get_all_records()
            subabonents_proceser = SubabonentsManager(print_page, today_date)
            subabonents_proceser.process_subabonents(data_to_print,
                                                     electro_talon)

        else:
            main_logger.warning(f"Невідома команда: {user_input}")
            print(f"\n'{user_input}' - Невідома команда.")
        
        user_input = input("\nТопаз дай команду: ")

    if user_input in out:
        main_logger.info("Вихід з програми.")


if __name__ == '__main__':
    print(GREETENGS)
    main()
