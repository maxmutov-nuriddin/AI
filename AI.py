import speech_recognition as sr
import pygetwindow as gw
import ctypes
import pyautogui
import webbrowser
import os

# Инициализация распознавателя речи
recognizer = sr.Recognizer()

# Задайте путь к папке "Фото"
photos_directory = os.path.expanduser("~/Pictures")  # Для Windows, macOS и Linux

# Убедитесь, что путь существует
if not os.path.exists(photos_directory):
    os.makedirs(photos_directory)  # Создает папку, если она не существует


def listen_command():
    with sr.Microphone() as source:
        print("Скажите команду...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio, language="ru-RU")
            command = command.lower()  # Привести к нижнему регистру
            print(f"Распознанная команда: {command}")  # Отладочный вывод
            return command
        except sr.UnknownValueError:
            print("Команда не распознана.")
            return ""
        except sr.RequestError:
            print("Ошибка соединения с сервисом распознавания.")
            return ""


def execute_command(command):
    if "открой сайт youtube" in command:
        webbrowser.open("https://www.youtube.com")
    elif "открой браузер" in command:
        os.system("start chrome")  # Открыть Google Chrome
    elif "выключи компьютер" in command:
        os.system("shutdown /s /t 1")  # Выключение
    elif "перезагрузи компьютер" in command:
        os.system("shutdown /r /t 1")  # Перезагрузка
    elif "сделай скриншот" in command:
        screenshot = pyautogui.screenshot()  # Делает скриншот
        screenshot_path = os.path.join(
            photos_directory, "screenshot.png"
        )  # Путь для сохранения
        screenshot.save(screenshot_path)  # Сохраняет скриншот как файл
        print(f"Скриншот сделан и сохранен как {screenshot_path}")
    elif "сверни все окна" in command:
        pyautogui.hotkey("win", "d")  # Сворачивает все окна
        print("Все окна свернуты.")
    elif "верни все окна" in command:
        pyautogui.hotkey("win", "d")  # Возвращает все окна
        print("Все окна вернуты.")
    elif "открой мой пк" in command:
        os.startfile("explorer.exe")  # Открытия мой компьютер
        print("Открываю 'Мой компьютер'.")
    elif "закрой мой компьютер" in command or "закрой мой пк" in command:
        try:
            # Получаем список всех открытых окон
            windows = gw.getWindowsWithTitle("Этот компьютер")
            if windows:
                for window in windows:
                    window.close()  # Закрывает найденные окна
                print("Закрываю 'Мой компьютер'.")
            else:
                print("Окно 'Мой компьютер' не открыто.")
        except Exception as e:
            print(f"Ошибка при закрытии окна: {e}")
    elif "открой telegram" in command:
        os.startfile(
            r"C:\Program Files\WindowsApps\TelegramMessengerLLP.TelegramDesktop_5.5.5.0_x64__t4vj0pshhgkwm\Telegram.exe"
        )
        print("Телеграмм открыт.")
    elif "очистить корзину" in command:
        # Очищаем корзину без подтверждения
        try:
            # Используем флаг SHERB_NOCONFIRMATION (0x00000001)
            SHERB_NOCONFIRMATION = 0x00000001
            ctypes.windll.shell32.SHEmptyRecycleBinW(0, None, SHERB_NOCONFIRMATION)
            print("Корзина очищена.")
        except Exception as e:
            print(f"Ошибка при очистке корзины: {e}")
    elif "конец" in command:
        pyautogui.hotkey("ctrl", "c")  # Заканчивается
    else:
        print("Неизвестная команда.")


def listen_for_commands():
    while True:
        command = listen_command()
        if command:
            execute_command(command)


if __name__ == "__main__":
    try:
        listen_for_commands()
    except KeyboardInterrupt:
        print("Программа остановлена пользователем.")
        exit()
