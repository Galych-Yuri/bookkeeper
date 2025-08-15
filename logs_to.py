import os
import logging

# Отримати шлях до каталогу програми
base_dir = os.path.dirname(os.path.abspath(__file__))
log_file_path = os.path.join(base_dir, 'payment_logs.log')


def check_write_access(file_path):
    directory = os.path.dirname(file_path)
    return os.access(directory, os.W_OK)


# Перевірка доступу для запису в каталог
if not check_write_access(log_file_path):
    raise PermissionError(
        f"Немає доступу для запису в каталог: {log_file_path}")


def setup_logging():
    # Ініціалізуємо логування
    logging.basicConfig(
        filename=log_file_path,
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s "
               "[%(filename)s:%(lineno)d] "
               "%(message)s",
        datefmt="%d-%m-%Y %H:%M:%S"
    )
    logging.debug("Логування ініціалізовано.")


# Викликаємо функцію налаштування логування
setup_logging()

# Основний логер програми
main_logger = logging.getLogger('main')
main_logger.setLevel(logging.INFO)

# Логер для принтера (дочірній логер).
print_logger = logging.getLogger('main.print_to')
print_logger.setLevel(logging.INFO)

# Логер для бази даних (дочірній логер).
class_logger = logging.getLogger('main.class_to')
class_logger.setLevel(logging.INFO)
