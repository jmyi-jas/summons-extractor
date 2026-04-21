# ============================================================
#  传票信息提取 & 日历同步工具
#  依赖：pdfplumber  icalendar  pytz
#  pip install pdfplumber icalendar pytz
# ============================================================

import os
import re
import uuid
import pytz
import pdfplumber
from datetime import datetime, timedelta
from collections import defaultdict
from icalendar import Calendar, Event, Alarm
import sys

def get_target_folder() -> str:
    """优先读命令行参数，其次交互输入"""
    # 方式1：python 传票助手.py "C:\我的传票文件夹"
    if len(sys.argv) > 1:
        folder = sys.argv[1]
        if os.path.isdir(folder):
            return folder
        print(f' 命令行参数路径不存在：{folder}')

    # 方式2：运行时手动输入
    while True:
        folder = input('请输入传票文件夹路径：').strip().strip('"')
        if os.path.isdir(folder):
            return folder
        print(f'  路径不存在，请重新输入。')

# ===== 配置区（按需修改）=====
TARGET_FOLDER = r"C:\传票试运行" 
TIMEZONE      = 'Asia/Shanghai'
OUTPUT_TXT    = os.path.join(TARGET_FOLDER, '传票信息汇总.txt')
OUTPUT_HTML   = os.path.join(TARGET_FOLDER, '传票索引.html')
OUTPUT_ICS    = os.path.join(TARGET_FOLDER, '传票日历.ics')
# ============================


# ──────────────────────────────────────────────
# 字段提取层
# ──────────────────────────────────────────────

def extract_case_number(text: str) -> str:
    """
    提取案号，格式如：（2026）沪0101民初101号
    """
    pattern = r'[（(]\d{4}[）)]\s*[^（(）)\s]*?\d+\s*号'
    match = re.search(pattern, text)
    if match:
        return re.sub(r'\s+', '', match.group(0))
    return ''


def extract_cause(text: str) -> str:
    """提取案由"""
    patterns = [
        r'案由[\s　]*[：:]\s*([^\n]+)',
        r'案由\s{1,4}([^\n]+)',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return m.group(1).strip()
    return ''


def extract_respondent(text: str) -> str:
    """
    提取被传唤人。
    传票示例：被传唤人  xxxx公司（带证据……）
    策略：取内容，去掉括号内的注意事项说明。
    """
    patterns = [
        r'被传唤人[\s　]*[：:]?\s*([^\n]+)',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            val = m.group(1).strip()
            # 去掉长括号说明（10字以上视为注意事项）
            val = re.sub(r'（[^）]{8,}）', '', val)
            val = re.sub(r'\([^)]{8,}\)', '', val)
            return val.strip()
    return ''


def extract_time(text: str) -> str:
    """
    提取应到时间，清理括号内注意事项。
    示例：2026年01月01日13:00（因进法院要安检，请提前30分钟到法院）
    """
    patterns = [
        # 带冒号时间
        r'应到时间[\s　]*[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*\d{1,2}:\d{2})',
        # 带"时分"的中文时间
        r'应到时间[\s　]*[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*\d{1,2}\s*时\s*\d{2}\s*分)',
        # 纯日期
        r'应到时间[\s　]*[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)',
        # 无字段名，直接匹配时间串
        r'(\d{4}年\d{1,2}月\d{1,2}日\d{1,2}:\d{2})',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            val = m.group(1).strip()
            # 去掉括号内注意事项
            val = re.sub(r'（[^）]*）', '', val)
            val = re.sub(r'\([^)]*\)', '', val)
            return val.strip()
    return ''


def extract_location(text: str) -> str:
    """
    提取应到处所。
    注意：传票中该行常与"注意事项"合并，需截断。
    示例：xx路1xx号第1x法庭（1xx）
    """
    patterns = [
        r'应到处所[\s　]*[：:]?\s*([^\n]+)',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            val = m.group(1).strip()
            # 如果"注意事项"混入，截断
            val = re.split(r'注意事项|被传唤人必须', val)[0].strip()
            return val
    return ''


def extract_court(text: str) -> str:
    """从文首提取法院名称"""
    m = re.search(r'[\u4e00-\u9fa5]+人民法院', text)
    if m:
        return m.group(0)
    m = re.search(r'[\u4e00-\u9fa5]+法院', text)
    return m.group(0) if m else ''


# ──────────────────────────────────────────────
# PDF 核心解析（表格优先，正则兜底）
# ──────────────────────────────────────────────

# 表格 key → 内部字段名映射（处理各地法院表头差异）
TABLE_KEY_MAP = {
    '案号':         '案号',
    '案由':         '案由',
    '被传唤人':     '被传唤人',
    '应到时间':     '应到时间',
    '应到处所':     '应到处所',
    '应到处所注意事项': '应到处所',   # 部分法院合并列
}

def extract_info_from_pdf(pdf_path: str, debug: bool = False) -> dict:
    """
    解析传票PDF，返回包含6个字段的字典。
    优先走表格提取，失败则正则全文兜底。
    """
    result = {
        '法院':    '',
        '案号':    '',
        '案由':    '',
        '被传唤人': '',
        '应到时间': '',
        '应到处所': '',
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ''
            table_dict: dict[str, str] = {}

            for page in pdf.pages:
                # ── 表格提取 ──
                for table in (page.extract_tables() or []):
                    for row in table:
                        if not row or len(row) < 2:
                            continue
                        raw_key = str(row[0] or '').replace('\n', '').replace(' ', '').strip()
                        raw_val = ' '.join(
                            str(c or '') for c in row[1:]
                        ).replace('\n', ' ').strip()
                        if raw_key:
                            table_dict[raw_key] = raw_val

                # ── 全文拼接 ──
                page_text = page.extract_text() or ''
                full_text += page_text + '\n'

            if debug:
                print('\n===== 表格原始数据 =====')
                for k, v in table_dict.items():
                    print(f'  [{k}] → {v}')
                print('\n===== 全文前600字 =====')
                print(full_text[:600])
                print('========================\n')

            # ── 用映射表填充结果 ──
            for raw_key, field in TABLE_KEY_MAP.items():
                if raw_key in table_dict and table_dict[raw_key] and not result[field]:
                    result[field] = table_dict[raw_key]

            # ── 清洗时间括号注意事项 ──
            if result['应到时间']:
                result['应到时间'] = re.sub(r'（[^）]*）', '', result['应到时间']).strip()
                result['应到时间'] = re.sub(r'\([^)]*\)', '', result['应到时间']).strip()

            # ── 清洗处所中混入的注意事项文字 ──
            if result['应到处所']:
                result['应到处所'] = re.split(r'注意事项|被传唤人必须', result['应到处所'])[0].strip()

            # ── 法院：从全文头部提取 ──
            result['法院'] = extract_court(full_text)

            # ── 正则兜底（表格提取失败时） ──
            fallbacks = {
                '案号':     lambda t: extract_case_number(t),
                '案由':     lambda t: extract_cause(t),
                '被传唤人': lambda t: extract_respondent(t),
                '应到时间': lambda t: extract_time(t),
                '应到处所': lambda t: extract_location(t),
            }
            for field, fn in fallbacks.items():
                if not result[field]:
                    result[field] = fn(full_text)

    except Exception as e:
        print(f'  ⚠️  PDF解析出错：{e}')

    return result


# ──────────────────────────────────────────────
# 人工确认交互
# ──────────────────────────────────────────────

FIELD_DISPLAY = {
    '法院':    '法　　院',
    '案号':    '案　　号',
    '案由':    '案　　由',
    '被传唤人': '被传唤人',
    '应到时间': '应到时间',
    '应到处所': '应到处所',
}
FIELDS_ORDER = ['法院', '案号', '案由', '被传唤人', '应到时间', '应到处所']


def confirm_info(info: dict) -> tuple[dict, bool]:
    """
    逐字段展示提取结果，供用户确认或手动修改。
    返回 (confirmed_dict, 是否生成日历)
    """
    print('\n' + '=' * 62)
    print('  📋  请逐项确认提取信息（直接 Enter = 确认，输入新值 = 修改）')
    print('=' * 62)

    confirmed = {}
    for field in FIELDS_ORDER:
        value   = info.get(field, '')
        label   = FIELD_DISPLAY[field]
        display = value if value else '⚠️  未识别'
        hint    = ' ← 未识别，建议手动填写' if not value else ''

        print(f'\n  {label}：{display}{hint}')
        user_input = input('    → ').strip()
        confirmed[field] = user_input if user_input else value

    print('\n' + '─' * 62)
    print('  ✅  最终确认信息：')
    for field in FIELDS_ORDER:
        print(f'      {FIELD_DISPLAY[field]}：{confirmed[field] or "（空）"}')
    print('─' * 62)

    ans = input('\n  确认无误，生成日历提醒？[Y/n]：').strip().upper()
    generate = (ans != 'N')
    if not generate:
        print('  ⏭️   已跳过日历生成。')
    return confirmed, generate


# ──────────────────────────────────────────────
# 时间解析
# ──────────────────────────────────────────────

def parse_datetime(time_str: str) -> datetime | None:
    """将中文时间字符串解析为带时区的 datetime，兼容数字与汉字之间有空格的情况"""
    # 去除括号注意事项
    time_str = re.sub(r'（[^）]*）|\([^)]*\)', '', time_str).strip()
    tz = pytz.timezone(TIMEZONE)

    candidates = [
        # ✅ 全面加 \s* 兼容：2026 年 05 月 20 日 14:30
        (r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*(\d{1,2})\s*[:\uff1a]\s*(\d{2})',
         True),
        # 中文时分：2026年2月6日13时45分
        (r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*(\d{1,2})\s*时\s*(\d{2})\s*分',
         True),
        # 纯日期（默认09:00）
        (r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日',
         False),
    ]

    for pat, has_time in candidates:
        m = re.search(pat, time_str)
        if not m:
            continue
        g = [int(x) for x in m.groups()]
        try:
            dt = datetime(g[0], g[1], g[2], g[3], g[4]) if has_time \
                 else datetime(g[0], g[1], g[2], 9, 0)
            return tz.localize(dt)
        except ValueError:
            continue

    # 最终兜底：打印调试信息
    print(f'  ⚠️  时间格式无法解析，原始字符串：「{time_str}」')
    return None



# ──────────────────────────────────────────────
# ICS 日历生成
# ──────────────────────────────────────────────

REMINDERS = [
    (timedelta(hours=-2),  '【2小时后开庭】'),
    (timedelta(days=-1),   '【明天开庭】'),
    (timedelta(weeks=-1),  '【一周后开庭】'),
]


def _make_alarm(delta: timedelta, prefix: str, case_number: str) -> Alarm:
    alarm = Alarm()
    alarm.add('ACTION',      'DISPLAY')
    alarm.add('DESCRIPTION', f'{prefix}{case_number}')
    alarm.add('TRIGGER',     delta)
    return alarm


def generate_ics(records: list, output_path: str) -> None:
    """为所有勾选生成日历的记录输出单个 .ics 文件"""
    cal = Calendar()
    cal.add('PRODID',          '-//传票日历助手//CN')
    cal.add('VERSION',         '2.0')
    cal.add('CALSCALE',        'GREGORIAN')
    cal.add('X-WR-CALNAME',   '开庭提醒')
    cal.add('X-WR-TIMEZONE',   TIMEZONE)

    count = 0
    for rec in records:
        if not rec.get('_gen_ics'):
            continue

        dt = parse_datetime(rec['应到时间'])
        if not dt:
            print(f'  ⚠️  [{rec["案号"]}] 时间解析失败，跳过。')
            continue

        event = Event()
        event.add('UID',      str(uuid.uuid4()))
        event.add('SUMMARY',  f'【开庭】{rec["被传唤人"]} | {rec["案号"]}')
        event.add('DTSTART',  dt)
        event.add('DTEND',    dt + timedelta(hours=2))
        event.add('LOCATION', rec['应到处所'])
        event.add('DESCRIPTION',
            f'案  号：{rec["案号"]}\n'
            f'案  由：{rec["案由"]}\n'
            f'被传唤人：{rec["被传唤人"]}\n'
            f'法  院：{rec["法院"]}\n'
            f'应到处所：{rec["应到处所"]}\n'
            f'传票文件：{rec.get("_pdf_path", "")}'
        )

        # ── 三档提醒 ──
        for delta, prefix in REMINDERS:
            event.add_component(_make_alarm(delta, prefix, rec['案号']))

        cal.add_component(event)
        count += 1

    try:
        with open(output_path, 'wb') as f:
            f.write(cal.to_ical())
        print(f'\n✅ 日历文件已生成：{os.path.abspath(output_path)}')
        print(f'   共 {count} 个事件，每个含 提前2h / 提前1天 / 提前1周 三档提醒')
        print('   ics文件适用范围很广，谷歌outlook以及安卓苹果都能导入 → 点击导入日历即可')
    except Exception as e:
        print(f'  ⚠️  日历写入失败：{e}')


# ──────────────────────────────────────────────
# 输出：TXT 汇总
# ──────────────────────────────────────────────

def write_to_txt(records: list, filename: str) -> None:
    lines = ['传票信息汇总', '=' * 50, '']
    for i, rec in enumerate(records, 1):
        lines += [
            f'【{i}】',
            f'  法　　院：{rec["法院"]}',
            f'  案　　号：{rec["案号"]}',
            f'  案　　由：{rec["案由"]}',
            f'  被传唤人：{rec["被传唤人"]}',
            f'  应到时间：{rec["应到时间"]}',
            f'  应到处所：{rec["应到处所"]}',
            '',
        ]
    try:
        with open(filename, 'w', encoding='utf-8-sig') as f:
            f.write('\n'.join(lines))
        print(f'✅ TXT 已保存：{os.path.abspath(filename)}')
    except Exception as e:
        print(f'  ⚠️  TXT写入失败：{e}')


# ──────────────────────────────────────────────
# 输出：HTML 可点击索引
# ──────────────────────────────────────────────

_HTML_STYLE = """
<style>
  * { box-sizing: border-box; }
  body { font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
         background: #f0f2f5; padding: 24px; color: #333; }
  h1 { font-size: 22px; margin-bottom: 16px; }
  table { border-collapse: collapse; width: 100%; background: #fff;
          border-radius: 8px; overflow: hidden;
          box-shadow: 0 2px 8px rgba(0,0,0,.12); }
  th { background: #1a3c5e; color: #fff; padding: 11px 14px;
       text-align: left; font-weight: 500; white-space: nowrap; }
  td { padding: 10px 14px; border-bottom: 1px solid #eef0f3;
       vertical-align: top; font-size: 14px; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #eaf4ff; }
  a.btn { display: inline-block; padding: 3px 10px; background: #1a3c5e;
          color: #fff; border-radius: 4px; text-decoration: none;
          font-size: 13px; white-space: nowrap; }
  a.btn:hover { background: #2e6da4; }
  .warn { color: #c0392b; font-style: italic; }
</style>
"""

def write_to_html(records: list, filename: str) -> None:
    sorted_recs = sorted(records, key=lambda r: r['应到时间'])
    rows_html = ''
    for rec in sorted_recs:
        pdf_path = rec.get('_pdf_path', '')
        # Windows 路径转 file:// URI
        uri  = 'file:///' + pdf_path.replace('\\', '/') if pdf_path else ''
        link = f'<a class="btn" href="{uri}" target="_blank">📄 传票</a>' if uri else ''

        def td(val):
            return f'<td>{val if val else "<span class=warn>未填</span>"}</td>'

        rows_html += (
            '<tr>'
            + td(rec['应到时间'])
            + td(rec['被传唤人'])
            + td(rec['案由'])
            + td(rec['案号'])
            + td(rec['法院'])
            + td(rec['应到处所'])
            + f'<td>{link}</td>'
            + '</tr>\n'
        )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8"><title>传票索引</title>{_HTML_STYLE}</head>
<body>
<h1>📋 传票索引</h1>
<table>
  <thead>
    <tr>
      <th>应到时间</th><th>被传唤人</th><th>案由</th>
      <th>案号</th><th>法院</th><th>应到处所</th><th>传票</th>
    </tr>
  </thead>
  <tbody>
{rows_html}
  </tbody>
</table>
</body>
</html>"""

    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f'✅ HTML索引已保存：{os.path.abspath(filename)}')
    except Exception as e:
        print(f'  ⚠️  HTML写入失败：{e}')


# ──────────────────────────────────────────────
# PDF 重命名
# ──────────────────────────────────────────────

def rename_pdf(pdf_path: str, case_number: str, respondent: str) -> str:
    """
    重命名为：传票_案号_被传唤人（前15字）.pdf
    便于在日历 DESCRIPTION 路径或 HTML 里一眼识别。
    """
    dir_name = os.path.dirname(pdf_path)
    _, ext   = os.path.splitext(pdf_path)

    safe = lambda s: re.sub(r'[\\/:*?"<>|]', '', s)
    new_name = f'传票_{safe(case_number)}_{safe(respondent)[:15]}{ext}'
    new_path = os.path.join(dir_name, new_name)

    if new_path == pdf_path:
        return pdf_path
    try:
        os.rename(pdf_path, new_path)
        print(f'  📝 已重命名：{new_name}')
        return new_path
    except Exception as e:
        print(f'  ⚠️  重命名失败：{e}')
        return pdf_path


# ──────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────

def main():
    target_folder = get_target_folder()
    output_txt    = os.path.join(target_folder, '传票信息汇总.txt')
    output_html   = os.path.join(target_folder, '传票索引.html')
    output_ics    = os.path.join(target_folder, '传票日历.ics')
    if not os.path.isdir(TARGET_FOLDER):
        print(f'错误：目标文件夹不存在：{TARGET_FOLDER}')
        return

    # ── 收集传票PDF ──
    pdf_files = [
        os.path.join(root, f)
        for root, _, files in os.walk(TARGET_FOLDER)
        for f in files
        if f.lower().endswith('.pdf') and '传票' in f
    ]

    if not pdf_files:
        print('未找到文件名含"传票"的 PDF 文件。')
        return

    print(f'\n🔍 共找到 {len(pdf_files)} 个传票PDF\n{"=" * 62}')
    results = []

    for idx, pdf_path in enumerate(pdf_files, 1):
        print(f'\n【{idx} / {len(pdf_files)}】{os.path.basename(pdf_path)}')
        print(f'  路径：{pdf_path}')

        # ── Step 1：提取信息 ──
        raw_info = extract_info_from_pdf(pdf_path, debug=False)

        # ── Step 2：先提供查阅快捷键，方便核对原件 ──
        print('\n  📂 可在确认信息前查阅原始传票：')
        while True:
            op = input(
                '  按 O 打开所在文件夹 / F 打开传票PDF / Enter 直接确认信息：'
            ).strip().upper()
            if op == 'O':
                os.startfile(os.path.dirname(pdf_path))
            elif op == 'F':
                os.startfile(pdf_path)
            elif op == '':
                break
            else:
                print('  无效输入，请重试。')

        # ── Step 3：人工逐字段确认 ──
        confirmed, gen_ics = confirm_info(raw_info)
        confirmed['_gen_ics'] = gen_ics

        # ── Step 4：重命名 ──
        new_path = rename_pdf(pdf_path, confirmed['案号'], confirmed['被传唤人'])
        confirmed['_pdf_path'] = new_path

        results.append(confirmed)

        # ── Step 5：处理完毕后的退出选项 ──
        quit_flag = False
        ans = input('\n  按 Enter 处理下一个 / Q 提前结束：').strip().upper()
        if ans == 'Q':
            quit_flag = True
            print('\n提前退出，正在保存已处理内容……')

        print('\n' + '─' * 62)

        if quit_flag:
            break

    # ── 统一输出 ──
    print('\n' + '=' * 62)
    write_to_txt(results,  OUTPUT_TXT)
    write_to_html(results, OUTPUT_HTML)
    generate_ics(results,  OUTPUT_ICS)

    print('\n🎉 全部处理完毕。')
    if input('\n是否立即打开传票索引页面？[y/N]：').strip().upper() == 'Y':
        os.startfile(OUTPUT_HTML)


if __name__ == '__main__':
    main()
