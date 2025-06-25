import pandas as pd
import xml.etree.ElementTree as ET
from lxml import etree
import os
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import traceback
import re

# --- Словари ---
acct_type_dict = { "1": "Пополнение оборотных средств", "2.1": "Приобретение земельного участка", "2.2": "Приобретение жилого здания", "2.3": "Приобретение нежилого здания", "2.4": "Приобретение жилого помещения на первичном рынке", "2.5": "Приобретение жилого помещения на вторичном рынке", "2.6": "Приобретение нежилого помещения", "2.7": "Приобретение иной недвижимости", "3": "Приобретение основных средств, за исключением недвижимости", "4.1": "Строительство жилого здания", "4.2": "Реконструкция жилого здания", "4.3": "Строительство нежилого здания", "4.4": "Реконструкция нежилого здания", "4.5": "Имущественные права по ДДУ (жильё)", "4.6": "Имущественные права по ДДУ (нежильё)", "4.7": "Инвест. проект (жильё и нежильё)", "4.8": "Инвест. проект (только нежильё)", "4.9": "Инвест. проект (инфраструктура)", "5": "Приобретение ценных бумаг", "6": "Участие в торгах/аукционе", "7": "Вклад в уставной капитал", "8": "Рефинансирование в своей организации", "9": "Рефинансирование в другой организации", "10": "Погашение долга третьего лица", "11": "Финансирование лизинга", "12": "Приобретение прав по займам", "13": "Займ другому лицу", "14": "POS-заем", "15": "Бытовые нужды", "16.1": "Образовательный кредит с господдержкой", "16.2": "Без господдержки", "16.3": "Иное на образование", "17": "Авто с пробегом до 1000 км", "18": "Авто с пробегом от 1000 км", "19": "Цель не определена", "20": "Компенсация по закону о защите инвалидов", "99": "Иная цель" }

owner_indic_dict = { "1": "Заемщик", "2": "Поручитель", "3": "Принципал по гарантии", "4": "Лизингополучатель", "5": "Финансирование/обеспечение", "99": "Иной вид участия" }

def convert_types_credit_report(df):
    date_fields = [
        "Дата открытия <openedDt>",
        "Дата обновления информации по займу <lastUpdatedDt>",
        "Дата обновления информации по платежу <lastUpdatedDt>",
        "Дата последнего платежа <lastPaymtDt>",
        "Плановая дата закрытия <closedDt>",
        "Плановая дата закрытия RUTDF <closeDt>",
        "Дата статуса договора <accountRatingDate>",
        "Дата создания записи <fileSinceDt>",
        "Дата передачи финансирования <fundDate>",
        "Дата формирования кредитной информации <headerReportingDt>",
        "Дата ближайшего платежа по основному долгу <principalTermsAmtDt>",
        "Дата возникновения обязательства субъекта <commitDate>",
        "Дата расчета <amtDate>",
        "Дата расчета <calcDate>",
        "Дата возникновения срочной задолженности <startDt>",
        "Дата платежа <paymtDate>",
        "Дата расчета -dueArrear <calcDate>",
        "Дата расчета -pastdueArrear <calcDate>",
        "Плановая дата закрытия trade <closeDt>",
        "Дата возникновения обязательства субъекта trade <commitDate>",
        "Дата открытия trade <openedDt>",
        "Дата расчета -accountAmt <amtDate>",
        "Дата ближайшего следующего платежа по основному долгу -paymtCondition <principalTermsAmtDt>",
    ]

    numeric_fields = [
        "Кредитный лимит <creditLimit>",
        "Сумма задолженности <amtOutstanding>",
        "Сумма платежа <paymtAmt>",
        "Сумма просроченной задолженности <amtPastDue>",
        "Сумма просроченной задолжности -dueArrear <amtPastDue>",
        "Сумма просроченной задолжности -pastdueArrear <amtPastDue>",
        "Текущий баланс <curBalanceAmt>",
        "Остаток суммы по договору <principalOutstanding>",
        "Просрочка <paymtPat>",
        "Ставка <creditTotalAmt>",
        "Сумма внесенных платежей по процентам <intTotalAmt>",
        "Сумма среднемесячного платежа <averPaymtAmt>",
        "Дни просрочки <daysPastDue>"
    ]

    for col in date_fields:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in numeric_fields:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')

    return df

ogrn_to_bki = { "1057746710713": "НБКИ", "1247700058319": "БКИ СБ", "1047796788819": "ОКБ", "1057747734934": "БКИ КредитИнфо" }

combined_fields = {
    "creditLimit": "Кредитный лимит <creditLimit>",
    "openedDt": "Дата открытия <openedDt>",
    "amtOutstanding": "Сумма задолженности <amtOutstanding>",
    "paymtAmt": "Сумма платежа <paymtAmt>",
    "lastUpdatedDt": "Дата обновления информации по займу <lastUpdatedDt>",
    "amtPastDue": "Сумма просроченной задолженности <amtPastDue>",
    "currencyCode": "Валюта <currencyCode>",
    "accountRating": "Статус договора <accountRating>",
    "lastPaymtDt": "Последний платёж <lastPaymtDt>",
    "closedDt": "Плановая дата закрытия <closedDt>",
    "closeDt": "Плановая дата закрытия RUTDF <closeDt>",
    "accountRatingDate": "Дата статуса договора <accountRatingDate>",
    "curBalanceAmt": "Текущий баланс <curBalanceAmt>",
    "fileSinceDt": "Дата создания записи <fileSinceDt>",
    "fundDate": "Дата передачи финансирования <fundDate>",
    "headerReportingDt": "Дата формирования кредитной информации <headerReportingDt>",
    "loanIndicator": "Признак прекращения обязательства <loanIndicator>",
    "principalOutstanding": "Остаток суммы по договору <principalOutstanding>",
    "paymtPat": "Просрочка <paymtPat>",
    "accountRatingText": "Статус договора <accountRatingText>",
    "acctType": "Тип займа <acctType>",
    "acctTypeText": "Тип займа <acctTypeText>",
    "ownerIndic": "Отношение к кредиту <ownerIndic>",
    "ownerIndicText": "Отношение к кредиту <ownerIndicText>",
    "creditTotalAmt": "Ставка <creditTotalAmt>",
    "paymtFreqText": "График платежей <paymtFreqText>",
    "businessCategory": "Вид организации <businessCategory>",
    "principalTermsAmtDt": "Дата ближайшего платежа по основному долгу <principalTermsAmtDt>",
    "averPaymtAmt": "Сумма среднемесячного платежа <averPaymtAmt>",
    "loanKindCode": "Код вида займа (кредита) <loanKindCode>",
    "commitDate": "Дата возникновения обязательства субъекта <commitDate>",
    "amtDate": "Дата расчета <amtDate>",
    "calcDate": "Дата расчета <calcDate>",
    "startDt": "Дата возникновения срочной задолженности <startDt>"
}

combined_payment_fields = {
    "paymtDate": "Дата платежа <paymtDate>",
    "paymtAmt": "Сумма платежа <paymtAmt>",
    "principalPaymtAmt": "Платёж по основному долгу <principalPaymtAmt>",
    "intPaymtAmt": "Платёж по процентам <intPaymtAmt>",
    "otherPaymtAmt": "Прочие платежи <otherPaymtAmt>",
    "totalAmt": "Общая сумма платежа <totalAmt>",
    "daysPastDue": "Дни просрочки <daysPastDue>",
    "currencyCode": "Валюта <currencyCode>",
    "lastUpdatedDt": "Дата обновления информации по платежу <lastUpdatedDt>",
    "intTotalAmt": "Сумма внесенных платежей по процентам <intTotalAmt>",
    "principalTotalAmt": "Общая сумма платежей по основному долгу <principalTotalAmt>"
}

# Функция отбора ССП
def evaluate_row_conditions(row, preply_df):
    comments_simple = set()
    comments_rutdf = set()
    marker_simple = "Идет в расчет"
    marker_rutdf = "Идет в расчет"
    criteria_simple = set()
    criteria_rutdf = set()

    # Базовая проверка разницы дней
    diff_days = row.get("Разница дней", 0)
    try:
        diff_days = int(diff_days)
    except Exception:
        diff_days = 0

    if pd.isna(row.get("Разница дней")) or diff_days >= 90:
        comments_simple.add("Более 90 дней с даты заявки")
        comments_rutdf.add("Более 90 дней с даты заявки")
        marker_simple = "Не идет в расчет"
        marker_rutdf = "Не идет в расчет"

    # Проверка на дубликат
    if row.get("Маркер дубликатов") == "Дубликат":
        comments_simple.add("Дубликат")
        comments_rutdf.add("Дубликат")
        marker_simple = "Не идет в расчет"
        marker_rutdf = "Не идет в расчет"


    # Только если БКИ = НБКИ
    if row.get("БКИ") == "НБКИ":
        contract_id = row.get("UUID договора")
        contract_rows = preply_df[preply_df["UUID договора"] == contract_id]

        if contract_rows.empty:
            comments_simple.add("Отсутствуют данные по договору")
            marker_simple = "Не идет в расчет"
            comments_rutdf.add("Отсутствуют данные по договору")
            marker_rutdf = "Не идет в расчет"
        else:
            def aggregate_rows(rows):
                if "Дата обновления информации по займу <lastUpdatedDt>" in rows.columns and not rows.empty:
                    rows = rows.copy()
                    rows["Дата обновления информации по займу <lastUpdatedDt>"] = pd.to_datetime(
                        rows["Дата обновления информации по займу <lastUpdatedDt>"], errors="coerce"
                    )
                    idx = rows["Дата обновления информации по займу <lastUpdatedDt>"].idxmax()
                    rows = rows.loc[[idx]]
                else:
                    rows = rows.head(1)

                aggregated = {}
                for col in rows.columns:
                    values = rows[col].dropna().values
                    aggregated[col] = values[0] if len(values) > 0 else None
                return aggregated
            aggregated_preply2 = {}

            # "Договор"
            contract_rows_simple = contract_rows[contract_rows["Тип"] == "Договор"]
            if not contract_rows_simple.empty:
                aggregated_preply = aggregate_rows(contract_rows_simple)

                date_request = row.get("Дата заявки")
                lastupdateDt = aggregated_preply.get("Дата обновления информации по займу <lastUpdatedDt>")
                closedDt = aggregated_preply.get("Плановая дата закрытия <closedDt>")
                openedDt = aggregated_preply.get("Дата открытия <openedDt>")
                acctType = aggregated_preply.get("Тип займа <acctType>")
                principal_outstanding = aggregated_preply.get("Остаток суммы по договору <principalOutstanding>")
                account_rating = aggregated_preply.get("Статус договора <accountRating>")
                account_rating_text = aggregated_preply.get("Статус договора <accountRatingText>")
                ownerIndic = aggregated_preply.get("Отношение к кредиту <ownerIndic>")
                amtPastDue = aggregated_preply.get("Сумма просроченной задолженности <amtPastDue>")

                field_map = aggregated_preply

                field_map = {
                    "Дата открытия <openedDt>": openedDt,
                    "Тип займа <acctType>": acctType,
                    "Статус договора <accountRating>": account_rating,
                    "Отношение к кредиту <ownerIndic>": ownerIndic,
                    "Плановая дата закрытия <closedDt>": closedDt,
                }

                missing_fields = []
                for name, val in field_map.items():
                    if pd.isna(val):
                        missing_fields.append(name)

                if missing_fields:
                    comments_simple.add("Отсутствуют данные в полях: " + ", ".join(missing_fields))
                    marker_simple = "Не идет в расчет"
                    criteria_simple.add("5.1")

                # Условие 1: Дата последнего обновления более 31 дня назад
                if pd.notna(lastupdateDt) and pd.notna(date_request):
                    if (date_request - lastupdateDt).days > 31:
                        comments_simple.add("Последнее обновление информации более 31 дня назад")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("1.1")
                
                closedDt_raw = row.get("closedDt")
                try:
                    closedDt = pd.to_datetime(closedDt_raw)
                except Exception:
                    closedDt = pd.NaT

                # Условие 2: Активный договор, но дата закрытия была более чем за 31 дней до подачи заявки
                if pd.notna(closedDt) and pd.notna(date_request):
                    delta_days = (closedDt - date_request).days

                    if account_rating == "0" and closedDt < (date_request - pd.Timedelta(days=31)):
                        comments_simple.add("Активный договор, но дата закрытия более чем за 31 дней до заявки")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("2.1")

                    elif closedDt < date_request:
                        comments_simple.add(f"Договор уже закрыт, прошло {abs(delta_days)} дней с даты закрытия")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("2.1")

                    elif delta_days < 31:
                        comments_simple.add(f"До плановой даты закрытия менее 31 дня: осталось {delta_days} дней")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("2.1")

                # Условие 3: Активный договор и остаток задолженности = 0 или отсутствует
                try:
                    if account_rating == "0" and (principal_outstanding is None or float(str(principal_outstanding).replace(",", ".")) <= 0):
                        comments_simple.add("Активный договор, но остаток задолженности равен нулю или отсутствует")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("3.1")
                except:
                    comments_simple.add("Некорректное значение остатка задолженности")
                    marker_simple = "Не идет в расчет"

                # Условие 4: Просрочен, но просроченной задолженности нет
                try:
                    if account_rating == "52" and (amtPastDue is None or amtPastDue == 0):
                        comments_simple.add("Договор с просрочкой, но сумма просроченной задолженности отсутствует или равна 0")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("4.1")
                except:
                    comments_simple.add("Некорректное значение просроченной задолженности")
                    marker_simple = "Не идет в расчет"

                # Условие 6: Договор закрыт (account_rating == "13")
                try:
                    if str(account_rating) == "13":
                        comments_simple.add("Статус договора — закрыт")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("6.1")
                except:
                    pass

                # Условие 7: Счет закрыт и передан в другую организацию (account_rating == "14") и это единственная запись по UUID
                try:
                    if account_rating == "14" and account_rating_text == "Счет закрыт - переведен на обслуживание в другую организацию":
                        comments_simple.add("Счет закрыт - переведен на обслуживание в другую организацию")
                        marker_simple = "Не идет в расчет"
                        criteria_simple.add("7.1")
                except:
                    pass

            # "Договор RUTDF"
            contract_rows_rutdf = contract_rows[contract_rows["Тип"] == "Договор RUTDF"]
            if not contract_rows_rutdf.empty:
                aggregated_preply2 = aggregate_rows(contract_rows_rutdf)
                field_map = aggregated_preply2

                required_fields = [
                    "Тип займа trade <acctType>",
                    "Отношение к кредиту trade <ownerIndic>",
                    "Дата открытия trade <openedDt>",
                    "Плановая дата закрытия trade <closeDt>",
                    "Код вида займа (кредита) trade <loanKindCode>"
                ]

                def set_marker_rutdf(new_marker):
                    nonlocal marker_rutdf
                    if new_marker == "Не идет в расчет":
                        marker_rutdf = new_marker

                # Условие 3: Проверка отсутствия обязательных полей
                missing_fields = []
                for field in required_fields:
                    val = field_map.get(field)
                    if pd.isna(val):
                        missing_fields.append(field)

                if missing_fields:
                    comments_rutdf.add("Отсутствуют данные в полях: " + ", ".join(missing_fields))
                    set_marker_rutdf("Не идет в расчет")
                    criteria_rutdf.add("3.2")

                # Условие 1–2: Проверка просрочки и задолженности
                loan_indicator = field_map.get("loanIndicator")
                pastdue_amtPastDue = field_map.get("Сумма просроченной задолжности -pastdueArrear <amtPastDue>")
                due_amtOutstanding = field_map.get("Сумма задолжности -dueArrear <amtOutstanding>")

                def is_zero_or_empty(val):
                    try:
                        return pd.isna(val) or float(val) == 0.0
                    except:
                        return True

                if pd.isna(loan_indicator):  # Активный договор
                    if not is_zero_or_empty(pastdue_amtPastDue):
                        comments_rutdf.add("Есть просроченная задолженность, идет в расчет")
                        set_marker_rutdf("Идет в расчет")
                        criteria_rutdf.add("1.2")
                    elif not is_zero_or_empty(due_amtOutstanding):
                        comments_rutdf.add("Нет просрочки, но есть задолженность, идет в расчет")
                        set_marker_rutdf("Идет в расчет")
                        criteria_rutdf.add("1.2")
                    else:
                        comments_rutdf.add("Нет просрочки и задолженности")
                        set_marker_rutdf("Не идет в расчет")
                        criteria_rutdf.add("2.2")
                else:
                    comments_rutdf.add("Договор не активен (loanIndicator заполнен)")
                    set_marker_rutdf("Не идет в расчет")
                    criteria_rutdf.add("1.2")

                    # Условие 4: loanIndicator ≠ 2
                    try:
                        if int(loan_indicator) != 2:
                            comments_rutdf.add("loanIndicator присутствует, но не равен 2 — договор закрыт без признака принудительного исполнения")
                            set_marker_rutdf("Не идет в расчет")
                            criteria_rutdf.add("4.2")
                    except:
                        comments_rutdf.add("Ошибка при обработке loanIndicator")
                        set_marker_rutdf("Не идет в расчет")

    return pd.Series([
        "; ".join(sorted(comments_simple)), marker_simple, ", ".join(sorted(criteria_simple)),
        "; ".join(sorted(comments_rutdf)), marker_rutdf, ", ".join(sorted(criteria_rutdf))
    ])

def ask_date_request():
    while True:
        date_str = simpledialog.askstring("Дата заявки", "Введите дату заявки (в формате ДД.ММ.ГГГГ):")
        if date_str is None:
            raise Exception("Дата заявки не указана.")
        try:
            return pd.to_datetime(date_str, format="%d.%m.%Y", errors="raise")
        except Exception:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ.")

# Функция парсинга ССП и удаления дубликатов
def parse_monthly_payment(xml_path, date_request, preply_df):
    contract_mkk = os.path.splitext(os.path.basename(xml_path))[0]
    first_word = contract_mkk.split()[0]
    tree = ET.parse(xml_path)
    root = tree.getroot()
    if root.tag == "ОтветНаЗапросСведений":
        parents = root.findall("Сведения")
    elif root.tag == "СведенияОПлатежах":
        parents = [root]
    else:
        parents = [root] 

    data = []
    for parent in parents:
        for kbki in parent.findall("КБКИ"):
            ogrn = kbki.attrib.get("ОГРН")
            bki_name = ogrn_to_bki.get(ogrn, ogrn)

            obligations = kbki.find("Обязательства")
            if obligations is None:
                continue

            bki = obligations.find("БКИ")
            if bki is None:
                continue

            for dogovor in bki.findall("Договор"):
                uid = dogovor.attrib.get("УИД")
                payment = dogovor.find("СреднемесячныйПлатеж")
                if payment is None:
                    continue

                date_calc = payment.attrib.get("ДатаРасчета")
                amount = payment.text.strip() if payment.text else None
                currency = payment.attrib.get("Валюта")

                data.append({
                    "БКИ": bki_name,
                    "UUID договора": uid,
                    "ДатаРасчета": date_calc,
                    "Сумма": amount,
                    "Валюта": currency,
                    "Дата заявки": date_request,
                    "Договор в МКК": first_word
                })

    df = pd.DataFrame(data)

    for col in ["ДатаРасчета", "Дата заявки"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Разница дней"] = (df["Дата заявки"] - df["ДатаРасчета"]).dt.days
    df["Сумма"] = pd.to_numeric(df["Сумма"].astype(str).str.replace(",", "."), errors="coerce")

    grouped = df.groupby("UUID договора")
    result = []

    for uid, group in grouped:
        group = group.copy()
        group["Маркер дубликатов"] = "Дубликат"

        # 1. Отбор по максимальной дате расчёта
        max_date = group["ДатаРасчета"].max()
        candidates = group[group["ДатаРасчета"] == max_date]

        # 2. Если таких несколько — отбор по максимальной сумме
        if len(candidates) > 1:
            max_amount = candidates["Сумма"].max()
            candidates = candidates[candidates["Сумма"] == max_amount]

        # 3. Если таких несколько — отбор по НБКИ
        if len(candidates) > 1 and (candidates["БКИ"] == "НБКИ").any():
            candidates = candidates[candidates["БКИ"] == "НБКИ"]

        # 4. Финальный выбор — первая подходящая
        idx_final = candidates.index[0]
        group.loc[idx_final, "Маркер дубликатов"] = "Оригинал"

        result.append(group)

    if not result:
        return pd.DataFrame(), pd.DataFrame()

    df_final = pd.concat(result)

    # Применяем комментарии и маркеры и критерии
    df_final[
        [
            "Комментарии простого договора",
            "Маркер простого договора",
            "Критерий простого договора",
            "Комментарии RUTDF",
            "Маркер RUTDF",
            "Критерий RUTDF",
        ]
    ] = df_final.apply(lambda row: evaluate_row_conditions(row, preply_df), axis=1)

    return df_final

# Функция парсинга кредитного отчета
def parse_credit_report(xml_path):
    tree = etree.parse(xml_path)
    root = tree.getroot()
    rows = []

    for parent_block in root:
        parent_tag = parent_block.tag

        for node_type in ["AccountReply", "AccountReplyRUTDF"]:
            for acc in parent_block.findall(f".//{node_type}"):
                serial = acc.findtext("serialNum")
                uuid = acc.findtext("uuid")
                contract = {
                    "Родительский тег": parent_tag,
                    "Тип": f"Договор{' RUTDF' if node_type == 'AccountReplyRUTDF' else ''}",
                    "Тип договора": node_type,
                    "Номер договора": serial,
                    "UUID договора": uuid
                }

                for tag, label in combined_fields.items():
                    val = acc.findtext(tag)
                    if tag == "closeDt" and not val:
                        val = acc.findtext("closedDt")
                    contract[label] = val

                if node_type == "AccountReplyRUTDF":
                    code = contract.get("Тип займа <acctType>")
                    contract["Тип займа <acctTypeText>"] = acct_type_dict.get(code)
                    owner_code = contract.get("Отношение к кредиту <ownerIndic>")
                    contract["Отношение к кредиту <ownerIndicText>"] = owner_indic_dict.get(owner_code)

                    paymtCondition_block = acc.find("paymtCondition")
                    if paymtCondition_block is not None:
                        contract["Дата ближайшего следующего платежа по основному долгу -paymtCondition <principalTermsAmtDt>"] = paymtCondition_block.findtext("principalTermsAmtDt")

                    monthAverPaymt_block = acc.find("monthAverPaymt")
                    if monthAverPaymt_block is not None:
                        contract["Величина среднемесячного платежа -monthAverPaymt <averPaymtAmt>"] = monthAverPaymt_block.findtext("averPaymtAmt")

                    trade_block = acc.find("trade")
                    if trade_block is not None:
                        contract["Код вида займа (кредита) trade <loanKindCode>"] = trade_block.findtext("loanKindCode")
                        contract["Дата возникновения обязательства субъекта trade <commitDate>"] = trade_block.findtext("commitDate")
                        contract["Тип займа trade <acctType>"] = trade_block.findtext("acctType")
                        contract["Отношение к кредиту trade <ownerIndic>"] = trade_block.findtext("ownerIndic")
                        contract["Плановая дата закрытия trade <closeDt>"] = trade_block.findtext("closeDt")
                        contract["Дата открытия trade <openedDt>"] = trade_block.findtext("openedDt")

                    accountAmt_block = acc.find("accountAmt")
                    if accountAmt_block is not None:
                        contract["Дата расчета -accountAmt <amtDate>"] = accountAmt_block.findtext("amtDate")

                    def parse_date(date_str):
                        try:
                            return datetime.strptime(date_str, "%Y-%m-%d")
                        except:
                            return None

                    # Самый свежий pastdueArrear по calcDate
                    pastdue_latest = None
                    pastdue_amt = None

                    for past in acc.findall("pastdueArrear"):
                        calc_date_str = past.findtext("calcDate")
                        calc_date = parse_date(calc_date_str)
                        if calc_date and (pastdue_latest is None or calc_date > pastdue_latest):
                            pastdue_latest = calc_date
                            pastdue_amt = past.findtext("amtPastDue")
                    
                    contract["Дата расчета -pastdueArrear <calcDate>"] = pastdue_latest.strftime("%Y-%m-%d") if pastdue_latest else None
                    contract["Сумма просроченной задолжности -pastdueArrear <amtPastDue>"] = pastdue_amt


                    # Самый свежий dueArrear по calcDate
                    due_latest = None
                    due_outstanding = None

                    for due in acc.findall("dueArrear"):
                        calc_date_str = due.findtext("calcDate")
                        calc_date = parse_date(calc_date_str)
                        if calc_date and (due_latest is None or calc_date > due_latest):
                            due_latest = calc_date
                            due_outstanding = due.findtext("amtOutstanding")

                    contract["Дата расчета -dueArrear <calcDate>"] = due_latest.strftime("%Y-%m-%d") if due_latest else None
                    contract["Сумма задолжности -dueArrear <amtOutstanding>"] = due_outstanding

                rows.append(contract)

                payments = []
                for p in acc.findall(".//payment"):
                    row = {
                        "Родительский тег": parent_tag,
                        "Тип": "Платёж",
                        "Тип договора": node_type,
                        "Номер договора": serial,
                        "UUID договора": uuid,
                        "Маркер дубликатов": ""
                    }
                    for tag, label in combined_payment_fields.items():
                        row[label] = p.findtext(tag)
                    payments.append(row)
                    rows.append(row)

                if payments:
                        total = sum(
                            float(str(p.get("Сумма платежа <paymtAmt>")).replace(",", "."))
                            for p in payments if p.get("Сумма платежа <paymtAmt>")
                        )
                        rows.append({
                            "Родительский тег": parent_tag,
                            "Тип": "Итого",
                            "Тип договора": node_type,
                            "Сумма платежа <paymtAmt>": total,
                            "Номер договора": serial,
                            "UUID договора": uuid
                        })

    df = pd.DataFrame(rows)

    payments_df = df[df["Тип"] == "Платёж"].copy()

    if any(payments_df["Родительский тег"].str.contains("preply2")):
        payments_df = mark_duplicates_preply2(payments_df)
    else:
        payments_df = mark_duplicates_preply(payments_df)

    df.loc[payments_df.index, "Маркер дубликатов"] = payments_df["Маркер дубликатов"]
    
    return df


# Дубликаты платежей в кредитном отчете (preply)
def mark_duplicates_preply(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_original_index"] = df.index  # Сохраняем исходный порядок строк

    def mark_group(group):
        if len(group) == 1:
            group["Маркер дубликатов"] = "Оригинал"
            return group

        group = group.sort_values("Дата обновления информации по платежу <lastUpdatedDt>", ascending=False)
        group["Маркер дубликатов"] = "Дубликат"
        group.iloc[0, group.columns.get_loc("Маркер дубликатов")] = "Оригинал"
        return group

    df_payments = df[df["Тип"] == "Платёж"].copy()
    df_others = df[df["Тип"] != "Платёж"].copy()

    df_payments = df_payments.groupby(
        ["UUID договора", "Дата платежа <paymtDate>", "Сумма платежа <paymtAmt>"],
        group_keys=False
    ).apply(mark_group)

    df = pd.concat([df_payments, df_others], ignore_index=True)

    # Восстанавливаем исходный порядок
    df = df.sort_values("_original_index").reset_index(drop=True)

    df = df.drop(columns=["_original_index"])
    return df
    
# Дубликаты платежей в кредитном отчете (preply2)
def mark_duplicates_preply2(df):
    if 'Маркер дубликатов' not in df.columns:
        df['Маркер дубликатов'] = 'Оригинал'
    else:
        df['Маркер дубликатов'] = 'Оригинал'

    mask = (df['Родительский тег'] == 'preply2') & (df['Тип'] == 'Платёж')
    df_payments = df[mask]

    group_cols = ['UUID договора', 'Сумма платежа <paymtAmt>', 'Дата платежа <paymtDate>']

    for key, group in df_payments.groupby(group_cols):
        # Проверяем, одинаковы ли totalAmt в группе
        if group['Общая сумма платежа <totalAmt>'].nunique() == 1:
            # Если да, выделяем только самый свежий по дате обновления как оригинал
            idx_latest = group['Дата обновления информации по платежу <lastUpdatedDt>'].idxmax()
            df.loc[group.index, 'Маркер дубликатов'] = 'Дубликат'
            df.loc[idx_latest, 'Маркер дубликатов'] = 'Оригинал'
        else:
            # Если разные totalAmt — все оригиналы
            df.loc[group.index, 'Маркер дубликатов'] = 'Оригинал'
    return df

def get_desktop_processed_path(ko_path):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    processed_folder = os.path.join(desktop, "Обработанные")
    os.makedirs(processed_folder, exist_ok=True)

    # Извлекаем числовую часть из названия файла (например, 5608421 из "5608421 КО.xml")
    base_name = os.path.splitext(os.path.basename(ko_path))[0]
    match = re.search(r"\d+", base_name)
    if match:
        core_name = match.group()
    else:
        core_name = "результат"

    # Базовое имя файла
    result_filename = f"{core_name}.xlsx"
    full_path = os.path.join(processed_folder, result_filename)

    # Проверка на уникальность
    counter = 1
    unique_path = full_path
    while os.path.exists(unique_path):
        unique_filename = f"{core_name} ({counter}).xlsx"
        unique_path = os.path.join(processed_folder, unique_filename)
        counter += 1

    return unique_path

def select_file(title):
    path = filedialog.askopenfilename(title=title, filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
    if not path:
        raise Exception(f"{title} не был выбран.")
    return path

# Сводка по месяцам + СМД по КИ
def make_monthly_summary_split(df: pd.DataFrame, writer: pd.ExcelWriter, df_simple_all: pd.DataFrame):
    required_cols = {
        "Тип", "Маркер дубликатов", "Дата платежа <paymtDate>",
        "Сумма платежа <paymtAmt>", "Родительский тег", "Дни просрочки <daysPastDue>"
    }
    if not required_cols.issubset(df.columns):
        print("В датафрейме отсутствуют необходимые колонки для сводки.")
        return

    # Только оригинальные платежи
    df_payments = df[
        (df["Тип"] == "Платёж") & 
        (df["Маркер дубликатов"] == "Оригинал")
    ].copy()

    if df_payments.empty:
        print("Нет оригинальных платежей для сводки.")
        return

    # Преобразования
    df_payments["Дата платежа <paymtDate>"] = pd.to_datetime(df_payments["Дата платежа <paymtDate>"], errors="coerce")
    df_payments["Сумма платежа <paymtAmt>"] = pd.to_numeric(
        df_payments["Сумма платежа <paymtAmt>"].astype(str).str.replace(",", "."),
        errors="coerce"
    )
    df_payments["Дни просрочки <daysPastDue>"] = pd.to_numeric(
        df_payments["Дни просрочки <daysPastDue>"].astype(str).str.replace(",", "."),
        errors="coerce"
    )

    # Назначаем Месяц
    df_payments["Месяц"] = df_payments["Дата платежа <paymtDate>"].dt.to_period("M")

    # Группировка preply — всё без фильтра
    df_preply = df_payments[df_payments["Родительский тег"] == "preply"]
    sum_preply = df_preply.groupby("Месяц")["Сумма платежа <paymtAmt>"].sum()

    # preply2 — фильтрация по просрочке < 30 ИЛИ NaN
    df_preply2 = df_payments[
        (df_payments["Родительский тег"] == "preply2") &
        (df_payments["Маркер дубликатов"] == "Оригинал") &
        ((df_payments["Дни просрочки <daysPastDue>"].isna()) | (df_payments["Дни просрочки <daysPastDue>"] < 30))
    ]
    sum_preply2 = df_preply2.groupby("Месяц")["Сумма платежа <paymtAmt>"].sum()

    # Сбор диапазона месяцев
    combined = pd.concat([sum_preply, sum_preply2])
    if combined.empty:
        print("Нет данных для формирования сводки.")
        return

    start = combined.index.min()
    if start.year < 2020:
        start = combined.index[combined.index.to_timestamp() >= pd.Timestamp("2020-01-01")].min()
    end = combined.index.max()

    all_months = pd.period_range(start=start, end=end, freq="M")

    # Финальная таблица
    df_summary = pd.DataFrame({"Месяц": all_months})
    df_summary["Сумма платежей в месяц preply"] = df_summary["Месяц"].map(sum_preply).fillna(0)
    df_summary["Сумма платежей в месяц preply2"] = df_summary["Месяц"].map(sum_preply2).fillna(0)
    df_summary["Разница"] = df_summary["Сумма платежей в месяц preply"] - df_summary["Сумма платежей в месяц preply2"]

    # Создаем новый столбец с комментарием
    df_summary["Комментарий"] = ""
    mask_added = (df_summary["Сумма платежей в месяц preply"] == 0) & (df_summary["Сумма платежей в месяц preply2"] == 0)
    df_summary.loc[mask_added, "Комментарий"] = "Месяц добавлен автоматически"

    # ---- Вставляем блок вычисления СМД ----

    # Получаем дату запроса из df_simple_all (предполагается, что есть столбец 'date_request')
    date_request = pd.to_datetime(df_simple_all["Дата заявки"].iloc[0], format="%d.%m.%Y", errors="coerce")

    if pd.isna(date_request):
        print("Ошибка: дата запроса отсутствует или некорректна.")
    else:
        for col in ["Сумма платежей в месяц preply", "Сумма платежей в месяц preply2", "Разница"]:
            df_summary[col] = df_summary[col].astype(str).str.replace(",", ".").astype(float)

        df_summary["Месяц_дата"] = df_summary["Месяц"].dt.to_timestamp()

        # Вычисляем стартовый месяц — первый день предыдущего месяца от date_request
        start_month = (date_request - pd.DateOffset(months=1)).replace(day=1)
        print(f"date_request: {date_request}, стартовый месяц: {start_month}")

        if start_month not in df_summary["Месяц_дата"].values:
            print(f"Стартовый месяц {start_month} не найден в данных, используем минимальный месяц.")
            start_month = df_summary["Месяц_дата"].min()
        print(f"Финальный стартовый месяц: {start_month}")

        def find_actual_start(start_date, series):
            # Получаем индекс по дате
            try:
                idx = series.index.get_loc(start_date)
            except KeyError:
                print(f"⚠️ Дата {start_date} не найдена, используем первую позицию")
                idx = 0

            # Ищем до 6 месяцев вперед первый с платежом != 0
            for i in range(idx, min(idx + 6, len(series))):
                if series.iat[i] != 0:
                    print(f"Найден платеж != 0 в месяце {series.index[i]} на позиции {i}")
                    return i

            # Если все нули, возвращаем индекс + 6 (7-й месяц)
            fallback_idx = idx + 6 if idx + 6 < len(series) else idx
            print(f"Все 6 месяцев платежи = 0, берём месяц под индексом {fallback_idx}")
            return fallback_idx

        # Индексы по дате
        df_summary = df_summary.sort_values("Месяц_дата", ascending=True).reset_index(drop=True)  # обязательно отсортировать по дате по возрастанию

        preply_series = df_summary.set_index("Месяц_дата")["Сумма платежей в месяц preply"]
        preply2_series = df_summary.set_index("Месяц_дата")["Сумма платежей в месяц preply2"]

        actual_start_idx_preply = find_actual_start(start_month, preply_series)
        actual_start_idx_preply2 = find_actual_start(start_month, preply2_series)

        print(f"actual_start_idx_preply: {actual_start_idx_preply}")
        print(f"actual_start_idx_preply2: {actual_start_idx_preply2}")

        # Для preply — берём 24 месяца вниз (в сторону уменьшения индекса)
        start_idx = actual_start_idx_preply
        end_idx = max(start_idx - 23, 0)
        slice_preply = preply_series.iloc[end_idx:start_idx + 1]

        # Для preply2 — то же самое
        start_idx2 = actual_start_idx_preply2
        end_idx2 = max(start_idx2 - 23, 0)
        slice_preply2 = preply2_series.iloc[end_idx2:start_idx2 + 1]

        def count_months_with_payment(slice_):
            count = (slice_ != 0).sum()
            return max(count, 18)

        count_preply = count_months_with_payment(slice_preply)
        count_preply2 = count_months_with_payment(slice_preply2)

        print(f"count_preply: {count_preply}, count_preply2: {count_preply2}")
        print(f"slice_preply.sum(): {slice_preply.sum()}, slice_preply2.sum(): {slice_preply2.sum()}")

        smd_preply = slice_preply.sum() / count_preply * 1.3 if count_preply > 0 else 0
        smd_preply2 = slice_preply2.sum() / count_preply2 * 1.3 if count_preply2 > 0 else 0

        # Добавляем итоговую строку с результатом
        new_row = {
            "Месяц": "СМД по КИ",
            "Сумма платежей в месяц preply": smd_preply,
            "Сумма платежей в месяц preply2": smd_preply2,
            "Разница": smd_preply - smd_preply2,
            "Комментарий": ""
        }
        df_summary = pd.concat([df_summary, pd.DataFrame([new_row])], ignore_index=True)
        print(f"Количество месяцев с платежом preply: {(slice_preply != 0).sum()}")
        print(f"Количество месяцев с платежом preply2: {(slice_preply2 != 0).sum()}")
        print("Месяцы с платежами preply:")
        print(slice_preply[slice_preply != 0])

        print("Месяцы с платежами preply2:")
        print(slice_preply2[slice_preply2 != 0])
        
        # Сортируем по дате, кроме итоговой строки
        df_data = df_summary[df_summary["Месяц"] != "СМД по КИ"].copy()
        df_data = df_data.sort_values("Месяц_дата", ascending=False)
        df_data["Месяц"] = df_data["Месяц_дата"].dt.strftime("%d.%m.%Y")

        # Добавляем итоговую строку обратно
        df_summary = pd.concat([df_data, df_summary[df_summary["Месяц"] == "СМД по КИ"]], ignore_index=True)

        # Сохраняем
        df_summary.to_excel(writer, sheet_name="Сводка платежей", index=False)

# Окошки
def main():
    root = tk.Tk()
    root.withdraw()

    try:
        # === 1. Получение входных данных ===
        date_request = ask_date_request()
        ssp_path = select_file("Выберите файл ССП")
        ko_path = select_file("Выберите файл КО")
        output_path = get_desktop_processed_path(ko_path)

        # === 2. Парсинг КО и обработка дубликатов ===
        credit_df = parse_credit_report(ko_path)
        credit_df_full = credit_df.copy()  # Сохраняем копию до фильтрации

        preply_df = credit_df[credit_df["Тип"] == "Договор"].copy()
        preply2_df = credit_df[credit_df["Тип"] == "Договор RUTDF"].copy()
        preply_df = mark_duplicates_preply(preply_df)
        preply2_df = mark_duplicates_preply2(preply2_df)
        credit_df = pd.concat([preply_df, preply2_df], ignore_index=True)

        # Приведение типов
        credit_df = convert_types_credit_report(credit_df)
        
        # === 3. Парсинг и фильтрация ССП ===
        # Передаем весь credit_df (включая платежи!)
        df_full = parse_monthly_payment(ssp_path, date_request, credit_df)
        # Определим колонки для простого договора и RUTDF
        cols_simple = [
            "БКИ", "UUID договора", "ДатаРасчета", "Сумма", "Валюта", "Дата заявки",
            "Договор в МКК", "Разница дней", "Маркер дубликатов",
            "Комментарии простого договора", "Маркер простого договора", "Критерий простого договора"
        ]
        cols_rutdf = [
            "БКИ", "UUID договора", "ДатаРасчета", "Сумма", "Валюта", "Дата заявки",
            "Договор в МКК", "Разница дней", "Маркер дубликатов",
            "Комментарии простого договора", "Комментарии RUTDF",
            "Маркер RUTDF", "Критерий RUTDF"
        ]

        # === 4. Подготовка таблиц ===
        df_simple_all = df_full[cols_simple].copy()
        df_rutdf_all = df_full[cols_rutdf].copy()

        # === 5. Сохранение в Excel ===
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            credit_df_full.to_excel(writer, sheet_name="Кредитный отчёт", index=False)
            df_simple_all.to_excel(writer, sheet_name="Среднемесячные платежи preply", index=False)
            df_rutdf_all.to_excel(writer, sheet_name="Среднемесячные платежи preply2", index=False)

            make_monthly_summary_split(credit_df_full, writer, df_simple_all)

    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")


if __name__ == "__main__":
    main()
