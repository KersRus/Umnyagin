import os
import smtplib
import requests
import pandas as pd
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.header import Header
from dateutil.relativedelta import relativedelta

# --- Настройки (берутся из переменных окружения GitHub Secrets) ---
YANDEX_DISK_URL = os.environ.get("YANDEX_DISK_URL")
SMTP_USER = os.environ.get("SMTP_USER")          # ваш логин на Яндексе
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")  # пароль приложения
RECIPIENTS = os.environ.get("RECIPIENTS", "").split(",") # список email через запятую

def download_yandex_disk_file(public_url):
    """Получает прямую ссылку на скачивание файла с Яндекс.Диска и скачивает его."""
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    params = {"public_key": public_url}
    response = requests.get(api_url, params=params)
    response.raise_for_status()
    download_url = response.json()["href"]
    
    file_response = requests.get(download_url)
    file_response.raise_for_status()
    
    local_filename = "insurance_contracts.xlsx"
    with open(local_filename, "wb") as f:
        f.write(file_response.content)
    return local_filename

def check_contracts(file_path):
    """Читает таблицу и отбирает договоры, до окончания которых осталось <= 2 месяцев."""
    # Предполагаем, что данные на первом листе. Если файл .csv, замените на pd.read_csv
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
    
    # Приводим название колонок к единому виду (убираем пробелы, приводим к нижнему регистру)
    df.columns = df.columns.str.strip().str.lower()
    
    # Ожидаемые колонки (на случай, если в таблице они называются чуть иначе)
    col_insured = "страхователь"
    col_type = "вид страхования"
    col_end_date = "период (до)"
    col_manager = "ответственный"
    
    # Проверяем, что нужные колонки есть
    for col in [col_insured, col_type, col_end_date, col_manager]:
        if col not in df.columns:
            raise KeyError(f"В таблице не найдена колонка: '{col}'")
    
    # Преобразуем даты, игнорируя ошибки
    df[col_end_date] = pd.to_datetime(df[col_end_date], errors="coerce", dayfirst=True)
    
    # Вычисляем дату "через 2 месяца" от сегодня
    today = datetime.now().date()
    two_months_later = today + relativedelta(months=2)
    
    # Фильтруем: дата окончания не пустая, >= сегодня и <= через 2 месяца
    mask = (
        df[col_end_date].notna() &
        (df[col_end_date].dt.date >= today) &
        (df[col_end_date].dt.date <= two_months_later)
    )
    filtered = df[mask]
    
    return filtered

def build_email_body(df):
    """Формирует текст письма на основе отфильтрованных данных."""
    if df.empty:
        return None  # Если договоров нет, письмо не отправляем
    
    lines = ["Добрый день!\n", "Напоминаем о договорах, требующих продления в ближайшие 2 месяца:\n"]
    for _, row in df.iterrows():
        lines.append(
            f"• Страхователь: {row['страхователь']}\n"
            f"  Вид страхования: {row['вид страхования']}\n"
            f"  Ответственный: {row['ответственный']}\n"
            f"  Период до: {row['период (до)'].strftime('%d.%m.%Y')}\n"
        )
    lines.append("\nПожалуйста, свяжитесь с ответственными лицами для продления.\n")
    return "\n".join(lines)

def send_email(subject, body, recipients):
    """Отправляет письмо через SMTP Яндекс.Почты."""
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(recipients)
    
    with smtplib.SMTP_SSL("smtp.yandex.ru", 465) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, recipients, msg.as_string())

def main():
    print("Скачивание файла с Яндекс.Диска...")
    file_path = download_yandex_disk_file(YANDEX_DISK_URL)
    
    print("Анализ договоров...")
    filtered_df = check_contracts(file_path)
    print(f"Найдено договоров для продления: {len(filtered_df)}")
    
    if filtered_df.empty:
        print("Договоров, требующих продления, нет. Письмо не отправляется.")
        return
    
    body = build_email_body(filtered_df)
    subject = f"Напоминание о продлении договоров ({datetime.now().strftime('%d.%m.%Y')})"
    
    print(f"Отправка письма на адреса: {RECIPIENTS}")
    send_email(subject, body, RECIPIENTS)
    print("Письмо отправлено.")

if __name__ == "__main__":
    main()
