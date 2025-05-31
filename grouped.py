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

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Преобразуем типы
    df["ДатаРасчета"] = pd.to_datetime(df["ДатаРасчета"], errors="coerce")
    df["Дата заявки"] = pd.to_datetime(df["Дата заявки"], errors="coerce")
    df["Разница дней"] = (df["Дата заявки"] - df["ДатаРасчета"]).dt.days
    df["Сумма"] = pd.to_numeric(df["Сумма"].astype(str).str.replace(",", "."), errors="coerce")

    # Группировка по УИД
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

    df_final = pd.concat(result)

    # 🔍 Проверка условий
    def evaluate_row_conditions(row):
        comments = []
        marker = "Идет в расчет"

        # Условие 1: >90 дней
        try:
            diff_days = int(row.get("Разница дней", 0))
            if pd.isna(row.get("Разница дней")) or diff_days >= 90:
                comments.append("Более 90 дней с даты заявки")
                marker = "Не идет в расчет"
        except:
            pass

        # Условие 2: Дубликат
        if row.get("Маркер дубликатов") == "Дубликат":
            comments.append("Дубликат")
            marker = "Не идет в расчет"

        # Условие 3+: только для НБКИ
        if row.get("БКИ") == "НБКИ":
            contract_id = row.get("UUID договора")
            preply_rows = preply_df[preply_df["UUID договора"] == contract_id]

            if preply_rows.empty:
                comments.append("Отсутствуют данные по договору")
                marker = "Не идет в расчет"
            else:
                matched = False
                for _, preply_row in preply_rows.iterrows():
                    date_request = row["Дата заявки"]
                    lastupdateDt = preply_row.get("Дата обновления информации по займу <lastUpdatedDt>")
                    closedDt = preply_row.get("Плановая дата закрытия <closedDt>")
                    openedDt = preply_row.get("Дата открытия <openedDt>")
                    acctType = preply_row.get("Тип займа <acctType>")
                    principal_outstanding = preply_row.get("Остаток суммы по договору <principalOutstanding>")
                    account_rating = preply_row.get("Статус договора <accountRating>")
                    ownerIndic = preply_row.get("Отношение к кредиту <ownerIndic>")

                    # Условие 3: отсутствие данных
                    necessary_fields = [closedDt, openedDt, acctType, account_rating, ownerIndic]
                    if any(pd.isna(field) for field in necessary_fields):
                        continue

                    # Условие 4: последнее обновление > 31 день
                    if pd.notna(lastupdateDt) and pd.notna(date_request):
                        if (date_request - lastupdateDt).days > 31:
                            comments.append("Последнее обновление свыше 31 дня")
                            marker = "Не идет в расчет"

                    # Условие 5: плановая дата закрытия < 31 дня
                    if pd.notna(closedDt) and pd.notna(date_request):
                        if (date_request - closedDt).days < 31:
                            comments.append("До плановой даты закрытия менее 31 дня")
                            marker = "Не идет в расчет"
                        else:
                            comments.append("До плановой даты закрытия менее 31 дня")
                            marker = "Не идет в расчет"

                    # Условие 6: остаток задолженности
                    try:
                        if principal_outstanding is None or float(str(principal_outstanding).replace(",", ".")) <= 0:
                            comments.append("Остаток задолженности равен нулю или отсутствует")
                            marker = "Не идет в расчет"
                    except:
                        comments.append("Некорректное значение остатка задолженности")
                        marker = "Не идет в расчет"

                    # Условие 7: договор закрыт
                    try:
                        if int(account_rating) == 13:
                            comments.append("Статус кредитного договора закрыт")
                            marker = "Не идет в расчет"
                    except:
                        pass

                    matched = True
                    break
                
                if not matched:
                    comments.append("Отсутствуют данные в полях")
                    marker = "Не идет в расчет"

        return pd.Series(["; ".join(comments), marker])

    # Применяем
    df_final[["Комментарии", "Маркер учета"]] = df_final.apply(evaluate_row_conditions, axis=1)

    # Отобранные
    df_selected = df_final[df_final["Маркер учета"] == "Идет в расчет"]

    return df_final, df_selected

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

# Окошки
def main():
    root = Tk()
    root.withdraw()
    date_request_str = simpledialog.askstring("Дата заявки", "Введите дату заявки (ДД.ММ.ГГГГ):")
    if not date_request_str:
        messagebox.showerror("Ошибка", "Дата заявки не указана.")
        return

    try:
        date_request = pd.to_datetime(date_request_str, format="%d.%m.%Y", errors="raise")
    except Exception:
        messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ.")
        return

    messagebox.showinfo("Выбор файла", "Сначала выберите XML файл со среднемесячными платежами")
    ssp_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not ssp_path:
        return

    messagebox.showinfo("Выбор файла", "Теперь выберите XML файл с кредитным отчётом")
    ko_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not ko_path:
        return

    try:
        # Парсим кредитный отчёт
        credit_df = parse_credit_report(ko_path)

        # 🔧 Приведение типов
        credit_df = convert_types_credit_report(credit_df)

        # Отбираем только договоры
        preply_df = credit_df[credit_df["Тип"] == "Договор"].copy()

        # Парсим среднемесячные платежи
        df_full, df_selected = parse_monthly_payment(ssp_path, date_request, preply_df)

        # Сохраняем
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx")],
                                                   title="Сохранить результат как")
        if not output_path:
            return

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            credit_df.to_excel(writer, sheet_name="Кредитный отчёт", index=False)
            df_full.to_excel(writer, sheet_name="Среднемесячные платежи", index=False)
            df_selected.to_excel(writer, sheet_name="Отобранные", index=False)

        messagebox.showinfo("Готово", f"Результаты сохранены в файл:\n{output_path}")

    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
if __name__ == "__main__":
    main()
