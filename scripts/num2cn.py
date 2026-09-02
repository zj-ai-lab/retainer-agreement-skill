#!/usr/bin/env python3
"""人民币金额数字 → 中文大写。

输入：int 或 float（人民币元，可含两位小数）
输出：中文大写金额字符串（不带"人民币"/"元"前后缀；由调用方决定）

约定：
- 整数部分支持 0..99,999,999（亿位）
- 不输出"圆"（用"元"由调用方拼接）
- 小数支持两位（角/分）；若 ==0 整数则不输出角分部分
- 0 → "零"
- 不支持负数（律师费场景用不到）

例：
  30000      → "叁万"
  15000      → "壹万伍仟"
  25800      → "贰万伍仟捌佰"
  100000     → "壹拾万"
  1000000    → "壹佰万"
  108        → "壹佰零捌"
  10080      → "壹万零捌拾"
  3000.50    → "叁仟元伍角"  ← 由调用方在元字面间插
  小数处理仅用于 fractional ；调用方常常只用整数版本。
"""
from __future__ import annotations

DIGITS = '零壹贰叁肆伍陆柒捌玖'
UNITS_SMALL = ['', '拾', '佰', '仟']
UNITS_BIG = ['', '万', '亿', '万亿']


def _section_to_cn(n: int) -> str:
    """0..9999 段（不含万/亿单位）→ 中文。"""
    if n == 0:
        return ''
    s = ''
    str_n = str(n).zfill(4)
    started = False
    zero_pending = False
    for i, ch in enumerate(str_n):
        digit = int(ch)
        unit = UNITS_SMALL[3 - i]
        if digit == 0:
            if started:
                zero_pending = True
            continue
        if zero_pending:
            s += '零'
            zero_pending = False
        s += DIGITS[digit] + unit
        started = True
    return s


def int_to_cn(n: int) -> str:
    """整数 → 中文大写。"""
    if n == 0:
        return '零'
    if n < 0:
        raise ValueError('negative not supported')

    sections = []  # 低位在前
    while n > 0:
        sections.append(n % 10000)
        n //= 10000

    result = ''
    for i in range(len(sections) - 1, -1, -1):
        sec = sections[i]
        if sec == 0:
            if result and not result.endswith('零'):
                result += '零'
            continue
        sec_cn = _section_to_cn(sec)
        # 小段内 < 1000 且不是最高段时，前面补零
        if i < len(sections) - 1 and sec < 1000:
            if result and not result.endswith('零'):
                result += '零'
        result += sec_cn + UNITS_BIG[i]

    # 清理结尾多余零
    while result.endswith('零'):
        result = result[:-1]
    return result


def amount_to_cn(amount) -> str:
    """金额（int / float / str）→ 中文大写（不带'元'后缀）。"""
    if isinstance(amount, str):
        amount = float(amount.replace(',', '').replace('，', ''))
    if isinstance(amount, int):
        return int_to_cn(amount)
    # float: split int + 2-digit fraction
    yuan = int(amount)
    cents = round((amount - yuan) * 100)
    if cents == 0:
        return int_to_cn(yuan)
    jiao = cents // 10
    fen = cents % 10
    s = int_to_cn(yuan) + '元'
    if jiao:
        s += DIGITS[jiao] + '角'
    if fen:
        s += DIGITS[fen] + '分'
    elif jiao:
        s += '整'  # 有角无分加"整"清晰
    return s


if __name__ == '__main__':
    tests = [
        (30000, '叁万'),
        (15000, '壹万伍仟'),
        (25800, '贰万伍仟捌佰'),
        (100000, '壹拾万'),
        (108, '壹佰零捌'),
        (10080, '壹万零捌拾'),
        (1000000, '壹佰万'),
        (50000, '伍万'),
        (12345678, '壹仟贰佰叁拾肆万伍仟陆佰柒拾捌'),
        (0, '零'),
    ]
    ok = True
    for inp, expected in tests:
        got = int_to_cn(inp)
        mark = '✓' if got == expected else '✗'
        if got != expected:
            ok = False
        print(f'{mark} {inp} -> {got!r} (expect {expected!r})')
    print('OK' if ok else 'FAIL')
