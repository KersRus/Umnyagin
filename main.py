import os
import smtplib
import html
import requests
import pandas as pd
from datetime import datetime
from email.mime.text import MIMEText
from email.header import Header
from dateutil.relativedelta import relativedelta

# --- Настройки (берутся из переменных окружения GitHub Secrets) ---
YANDEX_DISK_URL = os.environ.get("YANDEX_DISK_URL")
SMTP_USER = os.environ.get("SMTP_USER")          # ваш логин на Яндексе
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")  # пароль приложения
RECIPIENTS = os.environ.get("RECIPIENTS", "").split(",")  # список email через запятую

# --- Разрешённые ответственные (Даня исключён) ---
ALLOWED_MANAGERS = {"вика", "иван", "маша"}


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
    """Отбирает договоры с окончанием ровно через 2 месяца и ровно через 2 недели,
    оставляя только ответственных Вика / Иван / Маша."""
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")

    # Приводим название колонок к единому виду
    df.columns = df.columns.str.strip().str.lower()

    col_insured = "страхователь"
    col_type = "вид страхования"
    col_end_date = "период (до)"
    col_manager = "ответственный"
    col_broker = "брокер/бывшее название"

    for col in [col_insured, col_type, col_end_date, col_manager, col_broker]:
        if col not in df.columns:
            raise KeyError(f"В таблице не найдена колонка: '{col}'")

    df[col_end_date] = pd.to_datetime(df[col_end_date], errors="coerce", dayfirst=True)

    today = datetime.now().date()
    target_2m = today + relativedelta(months=2)
    target_2w = today + relativedelta(weeks=2)

    # Фильтр по ответственным (без учёта регистра и лишних пробелов)
    manager_ok = df[col_manager].astype(str).str.strip().str.lower().isin(ALLOWED_MANAGERS)

    # Фильтр по датам: ровно через 2 месяца ИЛИ ровно через 2 недели
    is_2m = df[col_end_date].dt.date == target_2m
    is_2w = df[col_end_date].dt.date == target_2w

    filtered = df[manager_ok & (is_2m | is_2w)].copy()

    # Категория для сортировки и оформления
    filtered["_urgent"] = filtered[col_end_date].dt.date == target_2w
    filtered = filtered.sort_values(by=["_urgent", col_end_date], ascending=[False, True])

    print(f"Сегодня: {today} | 2 недели: {target_2w} | 2 месяца: {target_2m}")
    print(f"Найдено совпадений: {len(filtered)}")

    return filtered, target_2w, target_2m


def build_email_body(df, target_2w, target_2m):
    """Формирует HTML-письмо: договоры через 2 недели — красным жирным."""
    if df.empty:
        return None

    parts = [
        "<p>Добрый день!</p>",
        "<p>Напоминаем о договорах, требующих продления:</p>",
        "<ul>",
    ]

    for _, row in df.iterrows():
        urgent = bool(row["_urgent"])
        style = "color:#c00000;font-weight:bold;" if urgent else ""
        broker = row.get("брокер/бывшее название", "")
        broker = "" if pd.isna(broker) else str(broker)

        parts.append(
            f'<li style="{style}">'
            f"<b>Страхователь:</b> {html.escape(str(row['страхователь']))}<br>"
            f"<b>Вид страхования:</b> {html.escape(str(row['вид страхования']))}<br>"
            f"<b>Брокер/Бывшее название:</b> {html.escape(broker)}<br>"
            f"<b>Ответственный:</b> {html.escape(str(row['ответственный']))}<br>"
            f"<b>Период до:</b> {row['период (до)'].strftime('%d.%m.%Y')}"
            f"</li><br>"
        )

    parts.append("</ul>")
    parts.append(
        f"<p><span style='color:#c00000;font-weight:bold;'>"
        f"Красным жирным</span> выделены договоры, истекающие через 2 недели "
        f"(до {target_2w.strftime('%d.%m.%Y')}).<br>"
        f"Остальные истекают через 2 месяца (до {target_2m.strftime('%d.%m.%Y')}).</p>"
    )
    parts.append("<p>Пожалуйста, свяжитесь с ответственными лицами для продления.</p>")

    return "".join(parts)


def send_email(subject, body, recipients):
    """Отправляет HTML-письмо через SMTP Яндекс.Почты."""
    msg = MIMEText(body, "html", "utf-8")
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
    filtered_df, target_2w, target_2m = check_contracts(file_path)

    if filtered_df.empty:
        print("Подходящих договоров нет. Письмо не отправляется.")
        return

    body = build_email_body(filtered_df, target_2w, target_2m)
    subject = (
        f"Продление договоров: 2 недели (до {target_2w.strftime('%d.%m.%Y')}) "
        f"и 2 месяца (до {target_2m.strftime('%d.%m.%Y')})"
    )

    print(f"Отправка письма на адреса: {RECIPIENTS}")
    send_email(subject, body, RECIPIENTS)
    print("Письмо отправлено.")


if __name__ == "__main__":
    main()
