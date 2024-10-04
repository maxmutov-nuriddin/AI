import speech_recognition as sr
import pygetwindow as gw
import ctypes
import pyautogui
import webbrowser
import os
import tkinter as tk

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


def show_message(message, duration=3000):
    # Создаем графическое окно для отображения сообщения
    message_window = tk.Tk()
    message_window.title("Сообщение")

    # Настраиваем размеры окна и позицию
    message_window.geometry(
        "300x100+100+100"
    )  # 300x100 пикселей, смещение на 100 пикселей по обеим осям
    message_window.attributes("-topmost", True)  # Делаем окно всегда на первом плане

    # Создаем метку с сообщением
    label = tk.Label(message_window, text=message)
    label.pack(pady=20)

    # Закрываем окно через определенное время
    message_window.after(duration, message_window.destroy)

    # Запускаем цикл обработки событий
    message_window.mainloop()


def execute_command(command):
    if "открой youtube" in command:
        webbrowser.open("https://www.youtube.com")
        show_message("Открываю YouTube.")
    elif "закрой youtube" in command:
        pyautogui.hotkey('ctrl', 'w')
        show_message("Закрываю YouTube.")
    elif "открой искусственный интеллект" in command:
        webbrowser.open("https://chatgpt.com/")
        show_message("Открываю искусственный интеллект.")
    elif "закрой искусственный интеллект" in command:
        pyautogui.hotkey('ctrl', 'w')
        show_message("Закрываю искусственный интеллект.")
    elif "открой браузер" in command:
        os.system("start chrome")  # Открыть Google Chrome
        show_message("Браузер открыт.")
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
        show_message(f"Скриншот сделан и сохранен как {screenshot_path}")
    elif "сверни все окна" in command:
        pyautogui.hotkey("win", "d")  # Сворачивает все окна
        show_message("Все окна свернуты.")
    elif "верни все окна" in command:
        pyautogui.hotkey("win", "d")  # Возвращает все окна
        show_message("Все окна вернуты.")
    elif "открой мой пк" in command:
        os.startfile("explorer.exe")  # Открывает 'Мой компьютер'
        show_message("Открываю 'Мой компьютер'.")
    elif "закрой мой компьютер" in command or "закрой мой пк" in command:
        try:
            windows = gw.getWindowsWithTitle("Этот компьютер")
            if windows:
                for window in windows:
                    window.close()  # Закрывает найденные окна
                show_message("Закрываю 'Мой компьютер'.")
            else:
                show_message("Окно 'Мой компьютер' не открыто.")
        except Exception as e:
            show_message(f"Ошибка при закрытии окна: {e}")
    elif "открой telegram" in command:
        os.startfile(
            r"C:\Program Files\WindowsApps\TelegramMessengerLLP.TelegramDesktop_5.5.5.0_x64__t4vj0pshhgkwm\Telegram.exe"
        )
        show_message("Телеграмм открыт.")
    elif "очистить корзину" in command:
        try:
            SHERB_NOCONFIRMATION = 0x00000001
            ctypes.windll.shell32.SHEmptyRecycleBinW(0, None, SHERB_NOCONFIRMATION)
            show_message("Корзина очищена.")
        except Exception as e:
            show_message(f"Ошибка при очистке корзины: {e}")
    elif "конец" in command:
        pyautogui.hotkey("ctrl", "c")  # Завершение
        show_message("Завершение программы.")
    else:
        show_message("Неизвестная команда.")


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
