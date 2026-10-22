import os
import smtplib
import html
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
    """Отбирает договоры трёх категорий:
    - 1w: срок уже истёк ИЛИ истекает в течение ближайшей недели (по сегодня+7);
    - 2w: дата окончания ровно через 2 недели;
    - 2m: дата окончания ровно через 2 месяца.
    Оставляет только ответственных Вика / Иван / Маша.
    """
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

    df[col_end_date] = pd.to_datetime(df[col_end_date], errors="coerce", dayfirst=True).dt.date

    today = datetime.now().date()
    week_end = today + timedelta(days=7)          # правая граница «в течение 1 недели»
    target_2w = today + relativedelta(weeks=2)    # ровно через 2 недели
    target_2m = today + relativedelta(months=2)   # ровно через 2 месяца

    # Фильтр по ответственным (без учёта регистра и лишних пробелов)
    manager_ok = df[col_manager].astype(str).str.strip().str.lower().isin(ALLOWED_MANAGERS)

    def category(d):
        if pd.isna(d):
            return None
        if d <= week_end:          # включает просроченные (в прошлом) и попадающие в ближайшие 7 дней
            return "1w"
        if d == target_2w:
            return "2w"
        if d == target_2m:
            return "2m"
        return None

    df["_category"] = df[col_end_date].apply(category)
    filtered = df[manager_ok & df["_category"].notna()].copy()

    # Сортировка: сначала 1 неделя/просрочено, затем 2 недели, затем 2 месяца; внутри — по дате
    order = {"1w": 0, "2w": 1, "2m": 2}
    filtered["_order"] = filtered["_category"].map(order)
    filtered = filtered.sort_values(by=["_order", col_end_date], ascending=[True, True])

    print(
        f"Сегодня: {today} | 1 неделя/просрочено: до {week_end} | "
        f"2 недели: {target_2w} | 2 месяца: {target_2m}"
    )
    print(f"Найдено совпадений: {len(filtered)}")

    return filtered, today, week_end, target_2w, target_2m


def build_email_body(df, today, week_end, target_2w, target_2m):
    """Формирует HTML-письмо.
    - 1w: дата «Период до» выделена красным жирным, остальной текст обычный;
    - 2w: весь блок выделен чёрным жирным (как было);
    - 2m: обычный шрифт.
    Между договорами — пустая строка.
    """
    if df.empty:
        return None

    parts = [
        "<p>Добрый день!</p>",
        "<p>Напоминаем о договорах, требующих продления:</p>",
    ]

    labels = {
        "1w": f"срок истёк или истекает в течение 1 недели (по {week_end.strftime('%d.%m.%Y')})",
        "2w": f"ровно через 2 недели ({target_2w.strftime('%d.%m.%Y')})",
        "2m": f"ровно через 2 месяца ({target_2m.strftime('%d.%m.%Y')})",
    }

    for _, row in df.iterrows():
        cat = row["_category"]
        date_str = row["период (до)"].strftime("%d.%m.%Y")

        # Выделение периода (до) для категории «1 неделя / просрочено»
        if cat == "1w":
            date_html = f'<span style="color:#c00000;font-weight:bold;">{date_str}</span>'
        else:
            date_html = date_str

        broker = row.get("брокер/бывшее название", "")
        broker = "" if pd.isna(broker) else str(broker)

        item = (
            f"<b>Страхователь:</b> {html.escape(str(row['страхователь']))}<br>"
            f"<b>Вид страхования:</b> {html.escape(str(row['вид страхования']))}<br>"
            f"<b>Брокер/Бывшее название:</b> {html.escape(broker)}<br>"
            f"<b>Ответственный:</b> {html.escape(str(row['ответственный']))}<br>"
            f"<b>Период до:</b> {date_html}<br>"
            f"<b>Срок:</b> {labels[cat]}"
        )

        # Категория «2 недели» — весь блок жирным чёрным (как раньше)
        if cat == "2w":
            item = f'<span style="color:#000000;font-weight:bold;">{item}</span>'

        # Двойной <br> даёт пустую строку между договорами
        parts.append(item + "<br><br>")

    parts.append(
        "<hr>"
        f"<p><span style='color:#c00000;font-weight:bold;'>Красным</span> выделена дата «Период до» "
        f"у договоров, срок которых истёк или истекает в течение 1 недели "
        f"(по {week_end.strftime('%d.%m.%Y')}).<br>"
        f"Чёрным жирным выделены договоры, заканчивающиеся ровно через 2 недели "
        f"({target_2w.strftime('%d.%m.%Y')}).<br>"
        f"Обычным шрифтом — договоры, заканчивающиеся ровно через 2 месяца "
        f"({target_2m.strftime('%d.%m.%Y')}).</p>"
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
    filtered_df, today, week_end, target_2w, target_2m = check_contracts(file_path)

    if filtered_df.empty:
        print("Подходящих договоров нет. Письмо не отправляется.")
        return

    body = build_email_body(filtered_df, today, week_end, target_2w, target_2m)
    subject = (
        f"Продление договоров: 1 неделя/просрочено (до {week_end.strftime('%d.%m.%Y')}), "
        f"2 недели ({target_2w.strftime('%d.%m.%Y')}), "
        f"2 месяца ({target_2m.strftime('%d.%m.%Y')})"
    )

    print(f"Отправка письма на адреса: {RECIPIENTS}")
    send_email(subject, body, RECIPIENTS)
    print("Письмо отправлено.")


if __name__ == "__main__":
    main()
