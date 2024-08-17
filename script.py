import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def send_email(smtp_server, port, sender_email, sender_password, receiver_email, subject, body, attachments):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Прикрепляем файлы
    for file in attachments:
        try:
            ctype, encoding = mimetypes.guess_type(file)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)

            with open(file, 'rb') as attachment:
                part = MIMEBase(maintype, subtype)
                part.set_payload(attachment.read())
                encoders.encode_base64(part)

                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{file.split("/")[-1]}"'
                )
                msg.attach(part)
        except Exception as e:
            print(f"Не удалось прикрепить файл {file}. Ошибка: {e}")

    try:
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        print(f"Письмо отправлено успешно на {receiver_email}!")
    except Exception as e:
        print(f"Не удалось отправить письмо на {receiver_email}. Ошибка: {e}")

def browse_file(entry, multiple=False):
    filetypes = [("Все файлы", "*.*")]
    if multiple:
        filenames = filedialog.askopenfilenames(filetypes=filetypes)
        if filenames:
            entry.delete(0, tk.END)
            entry.insert(0, "; ".join(filenames))
            attachment_files.clear()  # Очистка текущего списка вложений
            attachment_files.extend(filenames)  # Добавление новых файлов
    else:
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)

def send_emails():
    smtp_server = "smtp.gmail.com"  # Значение по умолчанию
    port = 587  # Значение по умолчанию
    sender_email = "youremail@gmail.com"  # Значение по умолчанию
    sender_password = "your apps pass word"  # Значение по умолчанию
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END).strip()  # Получаем текст из текстового поля

    excel_file = excel_entry.get()

    if not body:
        messagebox.showerror("Ошибка", "Текст письма не может быть пустым.")
        return

    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл Excel не найден. Убедитесь, что путь указан правильно.")
        return
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл Excel. Ошибка: {e}")
        return

    if 'Email' in df.columns:
        for email in df['Email']:
            send_email(smtp_server, port, sender_email, sender_password, email, subject, body, attachment_files)
        messagebox.showinfo("Успех", "Письма отправлены успешно!")
    else:
        messagebox.showerror("Ошибка", "В файле Excel нет колонки с именем 'Email'")

# Создание GUI
root = tk.Tk()
root.title("Рассыльщик писем")

# Изменение цвета фона окна
root.configure(bg="#f0f0f0")

# Функция для создания подписанных элементов с метками и входными полями
def create_labeled_entry(root, label_text, row, column, **kwargs):
    tk.Label(root, text=label_text, bg="#f0f0f0").grid(row=row, column=column, padx=10, pady=5, sticky="e")
    entry = tk.Entry(root, width=50, **kwargs)
    entry.grid(row=row, column=column + 1, padx=10, pady=5)
    return entry

# Создание элементов интерфейса
subject_entry = create_labeled_entry(root, "Тема письма:", 0, 0)

tk.Label(root, text="Текст письма:", bg="#f0f0f0").grid(row=1, column=0, padx=10, pady=5, sticky="ne")
body_text = tk.Text(root, width=50, height=10)
body_text.grid(row=1, column=1, padx=10, pady=5)

excel_entry = create_labeled_entry(root, "Файл Excel:", 2, 0)
tk.Button(root, text="Выбрать", command=lambda: browse_file(excel_entry)).grid(row=2, column=2, padx=10, pady=5)

attachment_entry = create_labeled_entry(root, "Вложения:", 3, 0)
tk.Button(root, text="Выбрать", command=lambda: browse_file(attachment_entry, multiple=True)).grid(row=3, column=2, padx=10, pady=5)

tk.Button(root, text="Отправить письма", command=send_emails).grid(row=4, column=1, padx=10, pady=10)

# Глобальная переменная для хранения путей к прикрепляемым файлам
attachment_files = []

root.mainloop()
