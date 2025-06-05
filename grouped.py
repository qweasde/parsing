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
def evaluate_row_conditions(row, preply_df):

    comments = set()
    marker = "Идет в расчет"
    criteria = set()

    # Проверка: разница дней >= 90
    diff_days = row.get("Разница дней", 0)
    try:
        diff_days = int(diff_days)
    except Exception:
        diff_days = 0

    if pd.isna(row.get("Разница дней")) or diff_days >= 90:
        comments.add("Более 90 дней с даты заявки")
        marker = "Не идет в расчет"

    # Проверка: дубликат
    if row.get("Маркер дубликатов") == "Дубликат":
        comments.add("Дубликат")
        marker = "Не идет в расчет"

    # Проверка условий по НБКИ (если БКИ == "НБКИ")
    if row.get("БКИ") == "НБКИ":
        contract_id = row.get("UUID договора")
        preply_rows = preply_df[preply_df["UUID договора"] == contract_id]

        if preply_rows.empty:
            comments.add("Отсутствуют данные по договору")
            marker = "Не идет в расчет"
        else:
            for _, preply_row in preply_rows.iterrows():
                date_request = row["Дата заявки"]
                lastupdateDt = preply_row.get("Дата обновления информации по займу <lastUpdatedDt>")
                closedDt = preply_row.get("Плановая дата закрытия <closedDt>")
                openedDt = preply_row.get("Дата открытия <openedDt>")
                acctType = preply_row.get("Тип займа <acctType>")
                principal_outstanding = preply_row.get("Остаток суммы по договору <principalOutstanding>")
                account_rating = preply_row.get("Статус договора <accountRating>")
                ownerIndic = preply_row.get("Отношение к кредиту <ownerIndic>")

                field_map = {
                    "Дата открытия <openedDt>": openedDt,
                    "Тип займа <acctType>": acctType,
                    "Статус договора <accountRating>": account_rating,
                    "Отношение к кредиту <ownerIndic>": ownerIndic,
                    "Плановая дата закрытия <closedDt>": closedDt,
                }

                missing_fields = [name for name, val in field_map.items() if pd.isna(val)]
                if missing_fields:
                    comments.add("Отсутствуют данные в полях: " + ", ".join(missing_fields))
                    marker = "Не идет в расчет"
                    criteria.add("4")

                if pd.notna(lastupdateDt) and pd.notna(date_request):
                    if (date_request - lastupdateDt).days > 31:
                        comments.add("Последнее обновление свыше 31 дня")
                        marker = "Не идет в расчет"
                        criteria.add("1")

                if pd.notna(closedDt) and pd.notna(date_request):
                    if (closedDt - date_request).days < 31:
                        comments.add("До плановой даты закрытия менее 31 дня")
                        marker = "Не идет в расчет"
                        criteria.add("2")

                try:
                    if principal_outstanding is None or float(str(principal_outstanding).replace(",", ".")) <= 0:
                        comments.add("Остаток задолженности равен нулю или отсутствует")
                        marker = "Не идет в расчет"
                        criteria.add("3")
                except:
                    comments.add("Некорректное значение остатка задолженности")
                    marker = "Не идет в расчет"

                try:
                    if int(account_rating) == 13:
                        comments.add("Статус кредитного договора закрыт")
                        marker = "Не идет в расчет"
                        criteria.add("5")
                except:
                    pass

    return pd.Series(["; ".join(sorted(comments)), marker, ", ".join(sorted(criteria))])

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


def mark_duplicates_preply(df):
    # Убедимся, что даты в нужном формате
    df['Дата платежа <paymtDate>'] = pd.to_datetime(df['Дата платежа <paymtDate>'])
    df['Дата обновления информации по платежу <lastUpdatedDt>'] = pd.to_datetime(df['Дата обновления информации по платежу <lastUpdatedDt>'])
    
    # Сортируем по UUID, дате платежа, сумме и по дате обновления (по убыванию, чтобы сверху была самая свежая)
    df = df.sort_values(by=['UUID договора', 'Дата платежа <paymtDate>', 'Сумма платежа <paymtAmt>', 'Дата обновления информации по платежу <lastUpdatedDt>'], ascending=[True, True, True, False])
    
    # Функция для маркировки внутри каждой группы
    def mark_group(group):
        group = group.copy()
        # Первая (самая свежая) — Оригинал, остальные — Дубликат
        group['Маркер дубликатов'] = ['Оригинал'] + ['Дубликат'] * (len(group) - 1)
        return group
    
    # Группируем по UUID + дата платежа + сумма платежа и применяем маркировку
    df = df.groupby(['UUID договора', 'Дата платежа <paymtDate>', 'Сумма платежа <paymtAmt>'], group_keys=False).apply(mark_group)
    
    return df

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
                        contract["Дата ближайшего следующего платежа по основному долгу <principalTermsAmtDt>"] = paymtCondition_block.findtext("principalTermsAmtDt")

                    monthAverPaymt_block = acc.find("monthAverPaymt")
                    if monthAverPaymt_block is not None:
                        contract["Величина среднемесячного платежа <averPaymtAmt>"] = monthAverPaymt_block.findtext("averPaymtAmt")

                    trade_block = acc.find("trade")
                    if trade_block is not None:
                        contract["Код вида займа (кредита) <loanKindCode>"] = trade_block.findtext("loanKindCode")
                        contract["Дата возникновения обязательства субъекта <commitDate>"] = trade_block.findtext("commitDate")

                    accountAmt_block = acc.find("accountAmt")
                    if accountAmt_block is not None:
                        contract["Дата расчета <amtDate>"] = accountAmt_block.findtext("amtDate")

                    pastdueArrear_block = acc.find("pastdueArrear")
                    if pastdueArrear_block is not None:
                        contract["Дата расчета <calcDate>"] = pastdueArrear_block.findtext("calcDate")

                    dueArrear_block = acc.find("dueArrear")
                    if dueArrear_block is not None:
                        contract["Дата возникновения срочной задолженности <startDt>"] = dueArrear_block.findtext("calcDate")

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
    ko_path = "C:/Users/islam/Desktop/Договоры/5484455 КО.xml"
    output_path = "C:/Users/islam/Desktop/Договоры/result.xlsx"
    safe_path = get_unique_filename(output_path)

    try:
        # Парсим кредитный отчёт
        credit_df = parse_credit_report(ko_path)
        credit_df = parse_credit_report(ko_path)
        
        # Приведение типов
        credit_df = convert_types_credit_report(credit_df)

        # Отбираем только договоры
        preply_df = credit_df[credit_df["Тип"] == "Договор"].copy()

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
