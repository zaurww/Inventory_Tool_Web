"""Readable Excel labels; canonical keys keep old workbooks compatible."""
SHEETS = {
    'Settings': 'Настройки', 'Products': 'Товары', 'Purchases': 'Покупки',
    'Expenses': 'Расходы', 'Sales': 'Продажи', 'Returns': 'Возвраты',
    'Counterparties': 'Контрагенты', 'Guide': 'Инструкция',
}
FIELDS = {
    'Setting': 'Параметр', 'Value': 'Значение', 'SKU': 'Код товара (SKU) *',
    'Product': 'Наименование товара *', 'Unit': 'Единица *',
    'Line ID': 'Код строки', 'Date': 'Дата *',
    'Shipment ID': 'Номер декларации / e-qaimə *', 'Document': 'Документ *',
    'Quantity': 'Количество *', 'Currency': 'Валюта *',
    'Amount FCY': 'Сумма в валюте *', 'FX to AZN': 'Курс к AZN *',
    'Expense ID': 'Код расхода', 'Document Date': 'Дата документа *',
    'Expense Type': 'Вид расхода *', 'Customer': 'Покупатель', 'Type': 'Операция *',
    'Revenue AZN': 'Выручка AZN', 'Return ID': 'Код возврата',
    'Original Line ID': 'Код исходной строки *',
    'Revenue Reversal AZN': 'Уменьшение выручки AZN',
    'Supplier': 'Поставщик *', 'Purchase Type': 'Вид покупки *',
    'Counterparty': 'Контрагент *', 'Name': 'Наименование *',
}
OPTIONAL_HEADERS = {'Purchases': ['Supplier', 'Purchase Type'], 'Expenses': ['Counterparty', 'Document']}
INPUT_COLUMNS = {
    'Settings': ['Setting', 'Value'],
    'Products': ['SKU', 'Product', 'Unit'],
    'Counterparties': ['Name'],
    'Purchases': ['Date', 'Purchase Type', 'Shipment ID', 'Supplier', 'Document', 'SKU', 'Quantity', 'Currency', 'Amount FCY', 'FX to AZN', 'Line ID'],
    'Expenses': ['Document Date', 'Shipment ID', 'Expense Type', 'Counterparty', 'Document', 'Currency', 'Amount FCY', 'FX to AZN', 'Expense ID'],
    'Sales': ['Date', 'Document', 'Customer', 'SKU', 'Type', 'Quantity', 'Revenue AZN', 'Line ID'],
    'Returns': ['Date', 'Type', 'Original Line ID', 'Quantity', 'Revenue Reversal AZN', 'Return ID'],
}
SETTINGS = {'Company': 'Компания', 'Start Month': 'Начало истории', 'Report Through': 'Отчёт по месяц включительно'}
TYPES = {'Продажа': 'Sale', 'Списание': 'Write-off', 'Покупатель': 'Customer', 'Поставщик': 'Supplier', 'Импорт': 'Import', 'Местная': 'Local'}


def canonical_columns(values):
    reverse = {v: k for labels in (FIELDS, AZ_FIELDS) for k, v in labels.items()}
    return [reverse.get(str(v).strip(), str(v).strip()) if v is not None else '' for v in values]


AZ_SHEETS = {
    'Settings': 'Parametrlər', 'Products': 'Mallar', 'Purchases': 'Alışlar',
    'Expenses': 'Xərclər', 'Sales': 'Satışlar', 'Returns': 'Qaytarmalar',
    'Counterparties': 'Kontragentlər', 'Guide': 'Təlimat',
}
AZ_FIELDS = {
    'Setting': 'Parametr', 'Value': 'Dəyər', 'SKU': 'Mal kodu (SKU) *',
    'Product': 'Malın adı *', 'Unit': 'Ölçü vahidi *', 'Line ID': 'Sətir kodu',
    'Date': 'Tarix *', 'Shipment ID': 'Bəyannamə / e-qaimə nömrəsi *',
    'Document': 'Sənəd *', 'Quantity': 'Miqdar *', 'Currency': 'Valyuta *',
    'Amount FCY': 'Valyuta ilə məbləğ *', 'FX to AZN': 'AZN məzənnəsi *',
    'Expense ID': 'Xərc kodu', 'Document Date': 'Sənədin tarixi *',
    'Expense Type': 'Xərc növü *', 'Customer': 'Alıcı', 'Type': 'Əməliyyat *',
    'Revenue AZN': 'Satış məbləği AZN', 'Return ID': 'Qaytarma kodu',
    'Original Line ID': 'İlkin sətirin kodu *',
    'Revenue Reversal AZN': 'Satış məbləğinin azalması AZN',
    'Supplier': 'Təchizatçı *', 'Purchase Type': 'Alış növü *',
    'Counterparty': 'Kontragent *', 'Name': 'Adı *',
}
AZ_SETTINGS = {'Company': 'Şirkət', 'Start Month': 'Uçotun başlanğıc ayı',
               'Report Through': 'Hesabatın son ayı'}
AZ_TYPES = {'Sale': 'Satış', 'Write-off': 'Silinmə', 'Customer': 'Alıcı',
            'Supplier': 'Təchizatçı', 'Import': 'İdxal', 'Local': 'Yerli'}
META_SHEET = '_InventoryMeta'
SCHEMA_VERSION = 2


def canonical_type(value):
    if not isinstance(value, str):
        return value
    value = value.strip()
    return {**TYPES, **{v: k for k, v in AZ_TYPES.items()}}.get(value, value)


def canonical_setting(value):
    return {v: k for labels in (SETTINGS, AZ_SETTINGS) for k, v in labels.items()}.get(value, value)


def matching_sheets(wb, name):
    return [title for title in (name, SHEETS[name], AZ_SHEETS[name]) if title in wb.sheetnames]


def resolve_sheet(wb, name):
    from inventory import InvalidInput, issue
    found = matching_sheets(wb, name)
    if len(found) > 1:
        raise InvalidInput([issue(name, 6, 'Eyni jurnal bir neçə dildə mövcuddur. Yalnız birini saxlayın.')])
    return found[0] if found else None


def schema_metadata(wb):
    from inventory import InvalidInput, issue
    if META_SHEET not in wb.sheetnames:
        return {}
    rows = list(wb[META_SHEET].iter_rows(values_only=True))
    keys = [r[0] for r in rows if r and r[0] is not None]
    meta = {r[0]: r[1] if len(r) > 1 else None for r in rows if r and r[0] is not None}
    if (len(keys) != len(set(keys)) or type(meta.get('schema_version')) is not int
            or meta['schema_version'] != SCHEMA_VERSION or meta.get('language') != 'az'
            or meta.get('validation_profile') not in ('modern', 'legacy')
            or type(meta.get('validate_parties')) is not bool):
        raise InvalidInput([issue(META_SHEET, 0, 'Kitabın format versiyası və ya metadatası dəstəklənmir. Proqramı yeniləyin.')])
    return meta
