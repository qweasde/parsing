import pandas as pd
import xml.etree.ElementTree as ET
from lxml import etree
import os
from collections import defaultdict
from datetime import datetime

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
        "Дата платежа <paymtDate>"
    ]

    numeric_fields = [
        "Кредитный лимит <creditLimit>",
        "Сумма задолженности <amtOutstanding>",
        "Сумма платежа <paymtAmt>",
        "Сумма просроченной задолженности <amtPastDue>",
        "Текущий баланс <curBalanceAmt>",
        "Остаток суммы по договору <principalOutstanding>",
        "Просрочка <paymtPat>",
        "Ставка <creditTotalAmt>",
        "Сумма среднемесячного платежа <averPaymtAmt>"
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
def evaluate_preply_conditions(row, preply_df):
    comments = set()
    marker = "Идет в расчет"
    criteria = set()
    result_rows = []

    if row.get("БКИ") == "НБКИ":
        contract_id = row.get("UUID договора")
        preply_rows = preply_df[preply_df["UUID договора"] == contract_id]

        if preply_rows.empty:
            comments.add("Отсутствуют данные по договору")
            marker = "Не идет в расчет"
            criteria.add("0")
        else:
            for _, preply_row in preply_rows.iterrows():
                # Извлечение нужных полей
                date_request = row.get("Дата заявки")
                last_update = preply_row.get("Дата обновления информации по займу <lastUpdatedDt>")
                closed_dt = preply_row.get("Плановая дата закрытия <closedDt>")
                opened_dt = preply_row.get("Дата открытия <openedDt>")
                account_rating = preply_row.get("Статус договора <accountRating>")
                acct_type = preply_row.get("Тип займа <acctType>")
                owner_indic = preply_row.get("Отношение к кредиту <ownerIndic>")
                amt_outstanding = preply_row.get("Остаток суммы по договору <amtOutstanding>")
                amt_past_due = preply_row.get("Просроченная задолженность <amtPastDue>")

                # Правило 1
                if pd.notna(last_update) and pd.notna(date_request):
                    if (date_request - last_update).days > 30:
                        comments.add("Последнее обновление свыше 30 дней")
                        marker = "Не идет в расчет"
                        criteria.add("1")

                # Правило 2
                if account_rating == "0" and pd.notna(closed_dt) and pd.notna(date_request):
                    if (closed_dt - date_request).days < -30:
                        comments.add("Плановая дата закрытия раньше даты заявки более чем на 30 дней")
                        marker = "Не идет в расчет"
                        criteria.add("2")

                # Правило 3
                if account_rating == "0" and (pd.isna(amt_outstanding) or amt_outstanding == 0):
                    comments.add("Остаток по активному договору нулевой")
                    marker = "Не идет в расчет"
                    criteria.add("3")

                # Правило 4
                if account_rating == "52" and (pd.isna(amt_past_due) or amt_past_due == 0):
                    comments.add("Просрочен, но нет просроченной задолженности")
                    marker = "Не идет в расчет"
                    criteria.add("4")

                # Правило 5 — обязательные поля
                field_map = {
                    "Статус договора <accountRating>": account_rating,
                    "Тип займа <acctType>": acct_type,
                    "Отношение к кредиту <ownerIndic>": owner_indic,
                    "Дата открытия <openedDt>": opened_dt,
                    "Плановая дата закрытия <closedDt>": closed_dt,
                }
                missing_fields = [name for name, val in field_map.items() if pd.isna(val)]
                if missing_fields:
                    comments.add("Отсутствуют данные в полях: " + ", ".join(missing_fields))
                    marker = "Не идет в расчет"
                    criteria.add("5")

                # Правило 6
                if account_rating == "13":
                    comments.add("Договор закрыт (accountRating=13)")
                    marker = "Не идет в расчет"
                    criteria.add("6")

                # Правило 7 — условие по закрытому договору и единственной записи
                if account_rating == "14":
                    same_uuid = preply_df[preply_df["UUID договора"] == contract_id]
                    if len(same_uuid) == 1:
                        comments.add("Счет закрыт и переведен в другую организацию — единственная запись")
                        marker = "Не идет в расчет"
                        criteria.add("7")

                diff_days = row.get("Разница дней", 0)
                try:
                    diff_days = int(diff_days)
                except Exception:
                    diff_days = 0

                if pd.isna(row.get("Разница дней")) or diff_days >= 90:
                    comments.add("Более 90 дней с даты заявки")
                    marker = "Не идет в расчет"
                    criteria.add("дополнительный код для 90 дней")

                # Проверка дубликата
                if row.get("Маркер дубликатов") == "Дубликат":
                    comments.add("Дубликат")
                    marker = "Не идет в расчет"
                    criteria.add("дополнительный код для дубликата")
                    
                # Сбор результата
                enriched_row = row.copy()
                enriched_row["Комментарий"] = "; ".join(comments)
                enriched_row["Маркер учета"] = marker
                enriched_row["Причины исключения"] = ", ".join(criteria)
                result_rows.append(enriched_row)

    return result_rows


def evaluate_preply2_conditions(row, preply2_df):
    comments = set()
    marker = "Идет в расчет"
    criteria = set()
    result_rows = []

    if row.get("БКИ") == "НБКИ":
        contract_id = row.get("UUID договора")
        rutdf_rows = preply2_df[preply2_df["UUID договора"] == contract_id]

        if rutdf_rows.empty:
            comments.add("Отсутствуют данные по договору (preply2)")
            marker = "Не идет в расчет"
            criteria.add("0")
        else:
            for _, r in rutdf_rows.iterrows():
                loan_indicator = r.get("loanIndicator")
                pastdue_amt = r.get("Просроченная задолженность <amtPastDue>")
                due_amt = r.get("Остаток задолженности <amtOutstanding>")
                pastdue_calc_date = r.get("Дата расчета просрочки <calcDate_pastdue>")
                due_calc_date = r.get("Дата расчета задолженности <calcDate_due>")

                acct_type = r.get("Тип займа trade <acctType>")
                owner_indic = r.get("Отношение к кредиту trade <ownerIndic>")
                opened_dt = r.get("Дата открытия trade <openedDt>")
                close_dt = r.get("Плановая дата закрытия trade <closeDt>")
                loan_kind_code = r.get("Код вида займа (кредита) trade <loanKindCode>")

                # Правило 1: Активный договор, нет просрочки и остатка
                if pd.isna(loan_indicator):
                    if pd.isna(pastdue_amt) or pastdue_amt == 0:
                        if pd.isna(due_amt) or due_amt == 0:
                            pass  # Всё ок
                        else:
                            comments.add("Остаток задолженности > 0")
                            marker = "Не идет в расчет"
                            criteria.add("1.2")
                    else:
                        comments.add("Просроченная задолженность > 0")
                        marker = "Не идет в расчет"
                        criteria.add("1.1")

                # Правило 2: отсутствует хотя бы один важный параметр
                important_fields = {
                    "Тип займа": acct_type,
                    "Отношение к кредиту": owner_indic,
                    "Дата открытия": opened_dt,
                    "Плановая дата закрытия": close_dt,
                    "Код вида займа": loan_kind_code
                }
                missing_fields = [name for name, val in important_fields.items() if pd.isna(val)]
                if missing_fields:
                    comments.add("Отсутствуют параметры: " + ", ".join(missing_fields))
                    marker = "Не идет в расчет"
                    criteria.add("2")

                # Правило 3: договор закрыт, но без признака принудительного исполнения
                if pd.notna(loan_indicator) and str(loan_indicator) != "2":
                    comments.add("Закрыт, нет признака принудительного исполнения")
                    marker = "Не идет в расчет"
                    criteria.add("3")

                diff_days = row.get("Разница дней", 0)
                try:
                    diff_days = int(diff_days)
                except Exception:
                    diff_days = 0

                if pd.isna(row.get("Разница дней")) or diff_days >= 90:
                    comments.add("Более 90 дней с даты заявки")
                    marker = "Не идет в расчет"
                    criteria.add("дополнительный код для 90 дней")

                # Проверка дубликата
                if row.get("Маркер дубликатов") == "Дубликат":
                    comments.add("Дубликат")
                    marker = "Не идет в расчет"
                    criteria.add("дополнительный код для дубликата")

                enriched_row = row.copy()
                enriched_row["Комментарий"] = "; ".join(comments)
                enriched_row["Маркер учета"] = marker
                enriched_row["Причины исключения"] = ", ".join(criteria)
                result_rows.append(enriched_row)

    return result_rows

# Функция парсинга ССП и удаления дубликатов
def parse_monthly_payment(xml_path, date_request, preply_df):
    contract_mkk = os.path.splitext(os.path.basename(xml_path))[0][:7]
    tree = ET.parse(xml_path)
    root = tree.getroot()

    data = []
    for kbki in root.findall("КБКИ"):
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
                "Договор в МКК": contract_mkk
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
        has_nbki = (group["БКИ"] == "НБКИ").any()
        if has_nbki:
            nbki_rows = group[group["БКИ"] == "НБКИ"]
            idx_max = nbki_rows["ДатаРасчета"].idxmax()
            group["Маркер дубликатов"] = "Дубликат"
            group.loc[idx_max, "Маркер дубликатов"] = "Оригинал"
        else:
            idx_max = group["ДатаРасчета"].idxmax()
            group["Маркер дубликатов"] = "Дубликат"
            group.loc[idx_max, "Маркер дубликатов"] = "Оригинал"
        result.append(group)

    if not result:
        return pd.DataFrame(), pd.DataFrame()

    df_final = pd.concat(result)

    # Применяем комментарии и маркеры и критерии
    df_final[["Комментарии", "Маркер учета", "Критерий отбора"]] = df_final.apply(lambda row: evaluate_row_conditions(row, preply_df), axis=1)

    # Отобранные — только те, которые идут в расчет
    df_selected = df_final[df_final["Маркер учета"] == "Идет в расчет"]
    
    # Строка Итого для отобранных
    total_sum = df_selected["Сумма"].sum()
    total_row = pd.Series({col: "" for col in df_selected.columns}, name="Итого")
    total_row["Сумма"] = total_sum
    df_selected_with_total = pd.concat([df_selected, pd.DataFrame([total_row])])
    return df_final, df_selected_with_total

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

                    pastdueArrear_block = acc.find("pastdueArrear")
                    if pastdueArrear_block is not None:
                        contract["Дата расчета -pastdueArrear <calcDate>"] = pastdueArrear_block.findtext("calcDate")
                        contract["Сумма просроченной задолжности -pastdueArrear <amtPastDue>"] = pastdueArrear_block.findtext("amtPastDue")

                    dueArrear_block = acc.find("dueArrear")
                    if dueArrear_block is not None:
                        contract["Дата расчета -dueArrear <calcDate>"] = dueArrear_block.findtext("calcDate")
                        contract["Сумма просроченной задолжности -dueArrear <amtPastDue>"] = dueArrear_block.findtext("amtPastDue")

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


# Дубликаты в кредитном отчете (preply)
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
    
# Дубликаты в кредитном отчете (preply2)
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

def get_unique_filename(path):
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path

    while os.path.exists(new_path):
        new_path = f"{base} ({counter}){ext}"
        counter += 1

    return new_path

# Окошки
def main():
    # Жестко задаём дату заявки
    date_request_str = "02.01.2025"
    try:
        date_request = pd.to_datetime(date_request_str, format="%d.%m.%Y", errors="raise")
    except Exception as e:
        print(f"Ошибка преобразования даты: {e}")
        return

    # Жестко задаём пути к файлам
    ssp_path = "C:/Users/islam/Desktop/Договоры/5608421 ССП.xml"
    ko_path = "C:/Users/islam/Desktop/Договоры/5608421 КО.xml"
    output_path = "C:/Users/islam/Desktop/Договоры/result.xlsx"
    safe_path = get_unique_filename(output_path)

    try:
        # Парсим кредитный отчёт
        credit_df = parse_credit_report(ko_path)

        # Теперь вызываем mark_duplicates_preply именно для credit_df
        credit_df = mark_duplicates_preply(credit_df)
        
        # Приведение типов
        credit_df = convert_types_credit_report(credit_df)

        # Отбираем только договоры и договоры RUTDF
        preply_df = credit_df[credit_df["Тип"].isin(["Договор", "Договор RUTDF"])].copy()

        # Парсим среднемесячные платежи
        df_full, df_selected = parse_monthly_payment(ssp_path, date_request, preply_df)

        # Сохраняем результат в Excel
        with pd.ExcelWriter(safe_path, engine="openpyxl") as writer:
            credit_df.to_excel(writer, sheet_name="Кредитный отчёт", index=False)
            df_full.to_excel(writer, sheet_name="Среднемесячные платежи", index=False)
            df_selected.to_excel(writer, sheet_name="Отобранные", index=False)

        print(f"Результаты сохранены в файл: {output_path}")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Произошла ошибка: {e}")

if __name__ == "__main__":
    main()
