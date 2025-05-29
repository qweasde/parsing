import pandas as pd
import xml.etree.ElementTree as ET
from lxml import etree
from tkinter import Tk, filedialog, simpledialog, messagebox
import os
from datetime import datetime

# --- Словари ---
acct_type_dict = { "1": "Пополнение оборотных средств", "2.1": "Приобретение земельного участка", "2.2": "Приобретение жилого здания", "2.3": "Приобретение нежилого здания", "2.4": "Приобретение жилого помещения на первичном рынке", "2.5": "Приобретение жилого помещения на вторичном рынке", "2.6": "Приобретение нежилого помещения", "2.7": "Приобретение иной недвижимости", "3": "Приобретение основных средств, за исключением недвижимости", "4.1": "Строительство жилого здания", "4.2": "Реконструкция жилого здания", "4.3": "Строительство нежилого здания", "4.4": "Реконструкция нежилого здания", "4.5": "Имущественные права по ДДУ (жильё)", "4.6": "Имущественные права по ДДУ (нежильё)", "4.7": "Инвест. проект (жильё и нежильё)", "4.8": "Инвест. проект (только нежильё)", "4.9": "Инвест. проект (инфраструктура)", "5": "Приобретение ценных бумаг", "6": "Участие в торгах/аукционе", "7": "Вклад в уставной капитал", "8": "Рефинансирование в своей организации", "9": "Рефинансирование в другой организации", "10": "Погашение долга третьего лица", "11": "Финансирование лизинга", "12": "Приобретение прав по займам", "13": "Займ другому лицу", "14": "POS-заем", "15": "Бытовые нужды", "16.1": "Образовательный кредит с господдержкой", "16.2": "Без господдержки", "16.3": "Иное на образование", "17": "Авто с пробегом до 1000 км", "18": "Авто с пробегом от 1000 км", "19": "Цель не определена", "20": "Компенсация по закону о защите инвалидов", "99": "Иная цель" }

owner_indic_dict = { "1": "Заемщик", "2": "Поручитель", "3": "Принципал по гарантии", "4": "Лизингополучатель", "5": "Финансирование/обеспечение", "99": "Иной вид участия" }

def convert_types_credit_report(df):
    date_fields = [
        "Дата открытия <openedDt>",
        "Дата обновления информации по займу <lastUpdatedDt>",
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
        "Дата возникновения срочной задолженности <startDt>"
    ]

        # Список полей с числами (пример по combined_fields)
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

# --- Среднемесячные платежи ---
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
                "УИД_Договора": uid,
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
    

    grouped = df.groupby("УИД_Договора")
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

    df_final = pd.concat(result)

    def check_nbki_conditions(row, preply_df):


        if row.get("ParentTag", "") != "preply":
            return pd.Series(["", "Идет в расчет"])

        contract_id = str(row.get("УИД_Договора")).strip()
        preply_df["УИД_Договора"] = preply_df["УИД_Договора"].astype(str).str.strip()

        if preply_row.empty:
            print(f"Не найдено совпадение в preply_df для УИД {contract_id}")

        preply_row = preply_row.iloc[0]

        date_request = row["Дата заявки"]
        lastupdateDt = pd.to_datetime(preply_row.get("Дата обновления информации по займу <lastUpdatedDt>"), errors="coerce")
        closedDt = pd.to_datetime(preply_row.get("Плановая дата закрытия <closedDt>"))
        openedDt = pd.to_datetime(preply_row.get("Дата открытия <openedDt>"))
        acctType = preply_row.get("Тип займа <acctType>")
        principal_outstanding = preply_row.get("Остаток суммы по договору <principalOutstanding>")
        account_rating = preply_row.get("Статус договора <accountRating>")
        ownerIndic = preply_row.get("Отношение к кредиту <ownerIndic>")

        # Отсутствие данных в полях
        necessary_fields = [closedDt, openedDt, acctType, account_rating, ownerIndic]
        for field in necessary_fields:
            if pd.isna(field):
                return pd.Series(["Отсутствуют данные в полях", "Не идет в расчет"])

        # Последнее обновление > 31 день
        if pd.notna(lastupdateDt) and pd.notna(date_request):
            delta1 = (date_request - lastupdateDt).days
            if delta1 > 31:
                return pd.Series(["Последнее обновление свыше 31 дня", "Не идет в расчет"])
        else:
            return pd.Series(["Отсутствуют данные в полях", "Не идет в расчет"])
    
        # До планового закрытия менее 31 дня
        if pd.notna(closedDt) and pd.notna(date_request):
            delta2 = (closedDt - date_request).days
            if delta2 > 31:
                return pd.Series(["До плановой даты закрытия прошло более 31 дня", "Не идет в расчет"])
        else:
            return pd.Series(["Отсутствуют данные в полях", "Не идет в расчет"])

        # Договор выплачен
        try:
            print(f"type: {type(preply_row['Остаток суммы по договору <principalOutstanding>'])}")
            if principal_outstanding is None or float(principal_outstanding) <= 0:
                return pd.Series(["Остаток задолженности равен нулю или отсутствует", "Не идет в расчет"])
        except Exception:
            return pd.Series(["Некорректное значение principalOutstanding", "Не идет в расчет"])

        # Договор закрыт
        if str(account_rating).strip() == "13":
            return pd.Series(["Статус кредитного договора закрыт", "Не идет в расчет"])

        print("NBKI CHECK:")
        print(f"  УИД: {contract_id}")
        print(f"  lastupdateDt: {lastupdateDt}")
        print(f"  closedDt: {closedDt}")
        print(f"  openedDt: {openedDt}")
        print(f"  acctType: {acctType}")
        print(f"  account_rating: {account_rating}")
        print(f"  ownerIndic: {ownerIndic}")

        return pd.Series(["", "Идет в расчет"])
        

    def comment_and_marker(row, nbki_comment, nbki_marker):
        comments = []
        marker = "Идет в расчет"

        # Добавим комментарий от NBKI, если есть
        if nbki_comment:
            comments.append(nbki_comment)
        if nbki_marker == "Не идет в расчет":
            marker = "Не идет в расчет"

        # Проверка дубликатов
        if row.get("Маркер дубликатов") == "Дубликат":
            comments.append("Дубликат")
            marker = "Не идет в расчет"

        # Проверка по дате
        diff_days = row.get("Разница дней", 0)
        try:
            diff_days = int(diff_days)
        except Exception:
            diff_days = 0

        if pd.isna(row.get("Разница дней")) or diff_days >= 90:
            comment_final = "Более 90 дней с даты заявки"
            marker = "Не идет в расчет"
        else:
            comment_final = "; ".join(comments) if comments else ""

        return pd.Series([comment_final, marker])
        
    def combined_conditions(row, preply_df):
    # код как я показал выше
        nbki_comment, nbki_marker = check_nbki_conditions(row, preply_df)
        return comment_and_marker(row, nbki_comment, nbki_marker)

    # Применяем к df_final
    df_final[["Комментарии", "Маркер учета"]] = df_final.apply(lambda row: combined_conditions(row, preply_df), axis=1)



    df["Сумма"] = pd.to_numeric(df["Сумма"].astype(str).str.replace(",", "."), errors="coerce")
    df["ДатаРасчета"] = pd.to_datetime(df["ДатаРасчета"], format="%Y-%m-%d", errors="coerce")
    df["Дата заявки"] = pd.to_datetime(df["Дата заявки"], format="%Y-%m-%d", errors="coerce")

    # Отобранные — только те, которые идут в расчет
    filtered = df_final[df_final["Маркер учета"] == "Идет в расчет"]

    return df_final, filtered

# --- Кредитный отчёт ---
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

                # Расшифровка кодов
                if node_type == "AccountReplyRUTDF":
                    code = contract.get("Тип займа <acctType>")
                    contract["Тип займа <acctTypeText>"] = acct_type_dict.get(code)
                    owner_code = contract.get("Отношение к кредиту <ownerIndic>")
                    contract["Отношение к кредиту <ownerIndicText>"] = owner_indic_dict.get(owner_code)

                    # Вложенные блоки
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
                        "Родительский тег": parent_tag,  # ✅ теперь тег у каждого платежа
                        "Тип": "Платёж",
                        "Тип договора": node_type,
                        "Номер договора": serial,
                        "UUID договора": uuid
                    }
                    for tag, label in combined_payment_fields.items():
                        row[label] = p.findtext(tag)
                    payments.append(row)
                    rows.append(row)

                if payments:
                    total = sum(
                        float(str(p.get("Сумма платежа <paymtAmt>", "0")).replace(",", ".")) for p in payments if p.get("Сумма платежа <paymtAmt>")
                    )
                    rows.append({
                        "Родительский тег": parent_tag,
                        "Тип": "Итого",
                        "Тип договора": node_type,
                        "Сумма платежа <paymtAmt>": total,
                        "Номер договора": serial,
                        "UUID договора": uuid
                    })

    return pd.DataFrame(rows)
    df = convert_types_credit_report(df)
    return(df)

# Окошки
def main():
    root = Tk()
    root.withdraw()
    date_request = simpledialog.askstring("Дата заявки", "Введите дату заявки (ДД.ММ.ГГГГ):")
    if not date_request:
        messagebox.showerror("Ошибка", "Дата заявки не указана.")
        return

    messagebox.showinfo("Выбор файла", "Сначала выберите XML файл со среднемесячными платежами")
    file1 = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not file1:
        return

    messagebox.showinfo("Выбор файла", "Теперь выберите XML файл с кредитным отчётом")
    file2 = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not file2:
        return

    try:
        # Парсим кредитный отчёт
        credit_df = parse_credit_report(file2)

        # Парсим среднемесячные платежи, передавая кредитный DF для проверки
        monthly_payments_full, monthly_payments_filtered = parse_monthly_payment(file1, date_request, credit_df)

        # Сохраняем результат в Excel с двумя листами
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx")],
                                                   title="Сохранить результат как")
        if not output_path:
            return

        with pd.ExcelWriter(output_path) as writer:
            credit_df.to_excel(writer, sheet_name="Кредитный отчёт", index=False)
            monthly_payments_full.to_excel(writer, sheet_name="Все среднемесячные платежи", index=False)
            monthly_payments_filtered.to_excel(writer, sheet_name="Отобранные платежи", index=False)

        messagebox.showinfo("Готово", f"Результаты успешно сохранены в файл:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
if __name__ == "__main__":
    df = pd.DataFrame({
        "Дата открытия <openedDt>": ["01.01.2020", "15.02.2021"],
        "Кредитный лимит <creditLimit>": ["10000,50", "20000,00"]
    })
    df = convert_types_credit_report(df)
    print(df)
    main()
