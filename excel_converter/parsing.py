"""Conservative locale normalization without guessing ambiguous numeric values."""
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation


MONTHS = {}
for number, words in enumerate([
    'январь января янв january jan', 'февраль февраля фев february feb',
    'март марта мар march mar', 'апрель апреля апр april apr',
    'май мая may', 'июнь июня июн june jun', 'июль июля июл july jul',
    'август августа авг august aug', 'сентябрь сентября сен сент september sep sept',
    'октябрь октября окт october oct', 'ноябрь ноября ноя november nov',
    'декабрь декабря дек december dec',
], 1):
    MONTHS.update({word: number for word in words.split()})


def flexible_date(value, spec):
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            raise ValueError('Дата содержит часовой пояс. Уточните календарную дату вручную.')
        if value.time().isoformat() != '00:00:00' and not spec.get('drop_time', False):
            raise ValueError('Дата содержит время. Включите «Отбрасывать время» или укажите дату вручную.')
        return value.date()
    if isinstance(value, date):
        return value
    text = ' '.join(str(value).strip().split()).lower()
    # Parse a complete time suffix, including zero seconds/fractions, before dropping it.
    stamp = re.fullmatch(r'(.+?)[t ](\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,6}))?)?', text)
    if stamp:
        day, hour, minute, second, fraction = stamp.groups()
        if int(hour) > 23 or int(minute) > 59 or int(second or 0) > 59:
            raise ValueError('Некорректное время в дате.')
        if any(int(x or 0) for x in (hour, minute, second, fraction)) and not spec.get('drop_time', False):
            raise ValueError('Дата содержит время. Включите «Отбрасывать время» или укажите дату вручную.')
        text = day
    text = re.sub(r'\s*г(?:ода|\.)?$', '', text).strip()
    try:
        named = re.fullmatch(r'(\d{1,2})[ .-]+([а-яёa-z]+)\.?[ ,-]+(\d{4})', text)
        english = re.fullmatch(r'([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})', text)
        if named and named[2] in MONTHS:
            return date(int(named[3]), MONTHS[named[2]], int(named[1]))
        if english and english[1] in MONTHS:
            return date(int(english[3]), MONTHS[english[1]], int(english[2]))
        year_first = re.fullmatch(r'(\d{4})([./-])(\d{1,2})\2(\d{1,2})', text)
        if year_first:
            return date(int(year_first[1]), int(year_first[3]), int(year_first[4]))
        local = re.fullmatch(r'(\d{1,2})([./-])(\d{1,2})\2(\d{4})', text)
        if local:
            a, separator, b, year = local.groups()
            a, b = int(a), int(b)
            order = spec.get('date_order', 'auto')
            if order == 'dmy' or (order == 'auto' and separator == '.'):
                day, month = a, b
            elif order == 'mdy':
                month, day = a, b
            elif a > 12 or a == b:
                day, month = a, b
            elif b > 12:
                month, day = a, b
            else:
                raise ValueError('Неоднозначный порядок дня и месяца. Выберите ДМГ/МДГ в настройках или исправьте дату.')
            return date(int(year), month, day)
    except ValueError as exc:
        if 'Неоднозначный' in str(exc):
            raise
        raise ValueError('Такой календарной даты не существует.') from None
    raise ValueError('Дата не распознана. Нужен четырёхзначный год; например 31/12/2026, 2026-12-31 или 31 декабря 2026.')


CURRENCY_TOKENS = {
    'kzt': 'KZT', 'тенге': 'KZT', 'теңге': 'KZT', 'тг.': 'KZT', 'тг': 'KZT', '₸': 'KZT',
    'usd': 'USD', 'us$': 'USD', 'eur': 'EUR', 'евро': 'EUR', '€': 'EUR',
    'rub': 'RUB', 'руб.': 'RUB', 'руб': 'RUB', '₽': 'RUB',
    'gbp': 'GBP', '£': 'GBP', 'cny': 'CNY',
}
TOKEN_PATTERN = '|'.join(re.escape(s) for s in sorted(CURRENCY_TOKENS, key=len, reverse=True))


def amount_parts(value):
    text = ' '.join(str(value).strip().split()).replace('−', '-').replace('’', "'")
    currencies = []
    # Recognized currency labels may surround the number or its parentheses.
    for _ in range(3):
        found = re.fullmatch(rf'({TOKEN_PATTERN})\s*(.+)', text, re.I)
        tail = re.fullmatch(rf'(.+?)\s*({TOKEN_PATTERN})', text, re.I)
        if found:
            currencies.append(CURRENCY_TOKENS[found[1].lower()])
            text = found[2]
        elif tail:
            currencies.append(CURRENCY_TOKENS[tail[2].lower()])
            text = tail[1]
        else:
            break
    if len(set(currencies)) > 1:
        raise ValueError('В сумме указаны разные валюты.')
    return text.strip(), currencies[0] if currencies else None


def flexible_amount(value, spec):
    text, _ = amount_parts(value)
    negative = text.startswith('(') and text.endswith(')')
    if negative:
        text = text[1:-1].strip()
        if text.startswith(('+', '-')):
            raise ValueError('В сумме одновременно скобки и знак. Уточните знак вручную.')
    sign = ''
    if text.startswith(('+', '-')):
        sign, text = text[0], text[1:].strip()
    if re.fullmatch(r'[0-9]+(?:\.[0-9]+)?[eE][+-]?[0-9]{1,3}', text):
        canonical = text
    else:
        decimal = '.' if isinstance(value, (int, float, Decimal)) else spec.get('decimal_separator', 'auto')
        if decimal == 'auto':
            if '.' in text and ',' in text:
                decimal = '.' if text.rfind('.') > text.rfind(',') else ','
            elif '.' in text or ',' in text:
                separator = '.' if '.' in text else ','
                if text.count(separator) > 1:
                    decimal = None
                else:
                    suffix = text.split(separator)[1]
                    if len(suffix) == 3:
                        raise ValueError('Неоднозначный разделитель: это дробь или тысячи. Выберите десятичный разделитель в настройках.')
                    decimal = separator
            else:
                decimal = None
        if decimal and decimal in text:
            if text.count(decimal) != 1:
                raise ValueError('Несколько десятичных разделителей.')
            integer, fraction = text.split(decimal)
            if not re.fullmatch(r'[0-9]+', fraction):
                raise ValueError('Некорректная дробная часть суммы.')
        else:
            integer, fraction = text, ''
        if not integer and fraction:
            integer = '0'
        separators = set(re.findall(r"[^0-9]", integer))
        if separators:
            if len(separators) != 1 or not separators <= {' ', "'", '.', ','}:
                raise ValueError('Некорректные или смешанные разделители тысяч.')
            group = next(iter(separators))
            if not re.fullmatch(r'[0-9]{1,3}(?:' + re.escape(group) + r'[0-9]{3})+', integer):
                raise ValueError('Некорректная группировка тысяч: ожидается, например, 1 234 567.')
            integer = integer.replace(group, '')
        if not re.fullmatch(r'[0-9]+', integer):
            raise ValueError('Сумма не распознана. Пример: 1 234,56; 1,234.56; (250,00).')
        canonical = integer + ('.' + fraction if fraction else '')
    try:
        amount = Decimal(('-' if negative else sign) + canonical)
    except InvalidOperation:
        raise ValueError('Сумма не распознана.') from None
    if not amount.is_finite():
        raise ValueError('Сумма должна быть конечным числом.')
    parts = amount.as_tuple()
    digits, exponent = list(parts.digits), parts.exponent
    while len(digits) > 1 and digits[-1] == 0:
        digits.pop()
        exponent += 1
    if amount == 0:
        exponent = 0
    if exponent < -2:
        raise ValueError('Больше двух значащих десятичных знаков. Округление требует исправления вручную.')
    if len(digits) + max(0, exponent) > 15:
        raise ValueError('Сумма превышает безопасную точность Excel (15 цифр).')
    return amount
