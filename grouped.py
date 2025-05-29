import pandas as pd
import xml.etree.ElementTree as ET
from lxml import etree
from tkinter import Tk, filedialog, simpledialog, messagebox
import os
from datetime import datetime

# --- Словари ---
acct_type_dict = { 
    "1": "Пополнение оборотных средств", 
    "2.1": "Приобретение земельного участка", 
    "2.2": "Приобретение жилого здания", 
    "2.3": "Приобретение нежилого здания", 
    "2.4": "Приобретение жилого помещения на первичном рынке", 
    "2.5": "Приобретение жилого помещения на вторичном рынке", 
    "2.6": "Приобретение нежилого помещения", 
    "2.7": "Приобретение иной недвижимости", 
    "3": "Приобретение основных средств, за исключением недвижимости", 
    "4.1": "Строительство жилого здания", 
    "4.2": "Реконструкция жилого здания", 
    "4.3": "Строительство нежилого здания", 
    "4.4": "Реконструкция нежилого здания", 
    "4.5": "Имущественные права по ДДУ (жильё)", 
    "4.6": "Имущественные права по ДДУ (нежильё)", 
    "4.7": "Инвест. проект (жильё и нежильё)", 
    "4.8": "Инвест. проект (только нежильё)", 
    "4.9": "Инвест. проект (инфраструктура)", 
    "5": "Приобретение ценных бумаг", 
    "6": "Участие в торгах/аукционе", 
    "7": "Вклад в уставной капитал", 
    "8": "Рефинансирование в своей организации", 
    "9": "Рефинансирование в другой организации", 
    "10": "Погашение долга третьего лица", 
    "11": "Финансирование лизинга", 
    "12": "Приобретение прав по займам", 
    "13": "Займ другому лицу", 
    "14": "POS-заем", 
    "15": "Бытовые нужды", 
    "16.1": "Образовательный кредит с господдержкой", 
    "16.2": "Без господдержки", 
    "16.3": "Иное на образование", 
    "17": "Авто с пробегом до 1000 км", 
    "18": "Авто с пробегом от 1000 км", 
    "19": "Цель не определена", 
    "20": "Компенсация по закону о защите инвалидов", 
    "99": "Иная цель" 
}

owner_indic_dict = { 
    "1": "Заемщик", 
    "2": "Поручитель", 
    "3": "Принципал по гарантии", 
    "4": "Лизингополучатель", 
    "5": "Финансирование/обеспечение", 
    "99": "Иной вид участия" 
}

ogrn_to_bki = { 
    "1057746710713": "НБКИ", 
    "1247700058319": "БКИ СБ", 
    "1047796788819": "ОКБ", 
    "1057747734934": "БКИ КредитИнфо" 
}

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
def parse_monthly_payment(xml_path, date_request):
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
    df["ДатаРасчета"] = pd.to_datetime(df["ДатаРасчета"], errors="coerce")
    df["Дата заявки"] = pd.to_datetime(df["Дата заявки"], dayfirst=True, errors="coerce")
    df["Разница дней"] = (df["Дата заявки"] - df["ДатаРасчета"]).dt.days
    df["Сумма"] = pd.to_numeric(df["Сумма"].str.replace(",", "."), errors="coerce")

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
    
    df_final["Комментарии"] = ""
    df_final["Маркер учета"] = "Идет в расчет"

    mask = df_final["Разница дней"] >= 90
    df_final.loc[mask, "Комментарии"] = "Более 90 дней с даты заявки"
    df_final.loc[mask, "Маркер учета"] = "Не идет в расчет"

    df_selected = df_final[df_final["Маркер учета"] == "Идет в расчет"].drop_duplicates()

    df_final["ДатаРасчета"] = df_final["ДатаРасчета"].dt.strftime("%d.%m.%Y")
    df_final["Дата заявки"] = df_final["Дата заявки"].dt.strftime("%d.%m.%Y")
    df_selected["ДатаРасчета"] = df_selected["ДатаРасчета"].dt.strftime("%d.%m.%Y")
    df_selected["Дата заявки"] = df_selected["Дата заявки"].dt.strftime("%d.%m.%Y")
    return df_final, df_selected

# --- Функция проверки по 5 условиям для договоров preply ---
def process_nbki_preply_check(df_avg_pay, df_credit):
    # Оставляем в df_credit только договоры с parent_tag == "preply"
    if "parent_tag" in df_credit.columns:
        col = df_credit["parent_tag"]
    elif "node_type" in df_credit.columns:
        col = df_credit["node_type"]
    else:
        col = pd.Series("", index=df_credit.index)

    df_credit_preply = df_credit[col == "preply"]

    def check_conditions(row):
    # 0. Проверка на "Разница дней >= 90"
        days_diff = row["Разница дней"] if "Разница дней" in row and pd.notna(row["Разница дней"]) else None
        if days_diff is not None and int(days_diff) >= 90:
            return pd.Series(["Более 90 дней с даты заявки", "Не идет в расчет"])

        # 1. Не НБКИ — пропускаем
        if row.get("БКИ") != "НБКИ":
            return pd.Series(["", "Идет в расчет"])

        uuid = row.get("УИД_Договора")
        if not uuid or "Дата заявки" not in row:
            return pd.Series(["Данные неполные", "Требует проверки"])

        date_zayavki = pd.to_datetime(row["Дата заявки"], errors="coerce")
        if pd.isna(date_zayavki):
            return pd.Series(["Дата заявки некорректна", "Требует проверки"])

        credit_row = df_credit_preply[df_credit_preply["UUID договора"] == uuid]
        if credit_row.empty:
            return pd.Series(["Нет данных по договору (preply)", "Требует проверки"])

        credit_row = credit_row.iloc[0]

        if "Дата обновления информации по займу <lastUpdatedDt>" not in credit_row or pd.isna(credit_row["Дата обновления информации по займу <lastUpdatedDt>"]):
            return pd.Series(["Отсутствует дата обновления", "Не идет в расчет"])

        last_update_str = credit_row["Дата обновления информации по займу <lastUpdatedDt>"]
        last_update = pd.to_datetime(last_update_str, errors="coerce")

        if pd.isna(last_update):
            return pd.Series(["Дата обновления некорректна", "Не идет в расчет"])


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
                        "UUID договора": uuid
                    }
                    for tag, label in combined_payment_fields.items():
                        row[label] = p.findtext(tag)
                    payments.append(row)
                    rows.append(row)
                if payments:
                    total = sum(
                        float(str(p.get("Сумма платежа <paymtAmt>", "0")).replace(",", ".")) 
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
    return pd.DataFrame(rows)

# --- Основная часть (Окошки) ---
def main():
    # Инициализация Tkinter
    root = Tk()
    root.withdraw()
    try:
        # Выбор файлов
        credit_xml_path = filedialog.askopenfilename(title="Выберите XML файл с кредитным отчетом (preply)")
        if not credit_xml_path:
            messagebox.showwarning("Отмена", "Не выбран файл кредитного отчёта")
            return
        payment_xml_path = filedialog.askopenfilename(title="Выберите XML файл со среднемесячными платежами (preply2)")
        if not payment_xml_path:
            messagebox.showwarning("Отмена", "Не выбран файл среднемесячных платежей")
            return
        
        # Запрос даты заявки (можно взять из имени файла, или запросить)
        date_request_str = simpledialog.askstring("Дата заявки", "Введите дату заявки (ДД.ММ.ГГГГ)")
        if not date_request_str:
            messagebox.showwarning("Отмена", "Дата заявки не введена")
            return
        
        df_all_payments, df_selected_payments = parse_monthly_payment(payment_xml_path, date_request_str)
        
        df_credit = parse_credit_report(credit_xml_path)
        
        df_credit["node_type"] = df_credit["Тип договора"].apply(lambda x: "preply" if x == "AccountReply" else "preply2")
        
        df_selected_payments_checked = process_nbki_preply_check(df_selected_payments.copy(), df_credit)
        
        df_all_payments = df_all_payments.merge(
            df_selected_payments_checked[["УИД_Договора", "Комментарии", "Маркер учета"]],
            on="УИД_Договора",
            how="left",
            suffixes=("", "_отобранные")
        )
        df_all_payments["Комментарии"] = df_all_payments["Комментарии_отобранные"].fillna(df_all_payments["Комментарии"])
        df_all_payments["Маркер учета"] = df_all_payments["Маркер учета_отобранные"].fillna(df_all_payments["Маркер учета"])
        df_all_payments.drop(columns=["Комментарии_отобранные", "Маркер учета_отобранные"], inplace=True)
        
        # Сохраняем в Excel
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")], title="Сохранить Excel отчет")
        if not save_path:
            messagebox.showwarning("Отмена", "Не выбран файл для сохранения")
            return
        
        with pd.ExcelWriter(save_path) as writer:
            df_all_payments.to_excel(writer, sheet_name="Среднемесячные платежи", index=False)
            df_credit.to_excel(writer, sheet_name="Кредитный отчет", index=False)
            df_selected_payments_checked.to_excel(writer, sheet_name="Отобранные", index=False)
            
        
        messagebox.showinfo("Готово", f"Данные успешно сохранены в {save_path}")
    
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
    finally:
        root.destroy()

if __name__ == "__main__":
    main()
