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
RECIPIENTS = os.environ.get("RECIPIENTS", "").split(",")  # список email через запятую

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
    """Читает таблицу и отбирает договоры, у которых дата окончания ровно через 2 месяца."""
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")

    # Приводим название колонок к единому виду
    df.columns = df.columns.str.strip().str.lower()

    col_insured = "страхователь"
    col_type = "вид страхования"
    col_end_date = "период (до)"
    col_manager = "ответственный"

    for col in [col_insured, col_type, col_end_date, col_manager]:
        if col not in df.columns:
            raise KeyError(f"В таблице не найдена колонка: '{col}'")

    df[col_end_date] = pd.to_datetime(df[col_end_date], errors="coerce", dayfirst=True)

    # Дата "ровно через 2 месяца"
    today = datetime.now().date()
    target_date = today + relativedelta(months=2)

    # Фильтруем: дата окончания точно равна target_date
    mask = df[col_end_date].dt.date == target_date
    filtered = df[mask]

    print(f"Сегодня: {today} | Ищем договоры с датой окончания: {target_date}")
    print(f"Найдено совпадений: {len(filtered)}")

    return filtered

def build_email_body(df):
    """Формирует текст письма на основе отфильтрованных данных."""
    if df.empty:
        return None

    lines = [
        "Добрый день!\n",
        "Напоминаем о договорах, срок действия которых истекает ровно через 2 месяца:\n",
    ]
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

    if filtered_df.empty:
        print("Договоров с датой окончания ровно через 2 месяца нет. Письмо не отправляется.")
        return

    body = build_email_body(filtered_df)
    subject = (
        f"Напоминание о продлении договоров на "
        f"{(datetime.now().date() + relativedelta(months=2)).strftime('%d.%m.%Y')}"
    )

    print(f"Отправка письма на адреса: {RECIPIENTS}")
    send_email(subject, body, RECIPIENTS)
    print("Письмо отправлено.")

if __name__ == "__main__":
    main()
