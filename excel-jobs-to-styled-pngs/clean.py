#!/usr/bin/env python3
"""Stage 0 — Clean a recruitment Excel into a tidy format that render.py can consume.

Input : xlsx path (CLI arg or default)
Output: <stem>_cleaned.xlsx       — one row per (company, job), no section/policy noise
        <stem>_clean_report.txt   — list of dropped rows and the reason for each

Edit CONFIGS at the top to point at your file / sheet(s) and column layout.
"""
import openpyxl
import sys
from pathlib import Path

# (sheet_name, min_data_row, company_col, job_col, salary_col)
CONFIGS = [
    ('公众号最新版', 2, 2, 3, 6),
    ('招聘详情',    3, 2, 3, 5),
]

# 公司列里出现这些关键字的行 → 当政策/资料说明丢掉
NOISE_KEYWORDS = ('政策', '贷款', '详见')

# 序号列以这些前缀开头的行 → section header
SECTION_PREFIXES_SKIP  = ('一、',)   # 单条 skip
SECTION_PREFIXES_BREAK = ('二、',)   # break 整个 sheet 的处理


def iter_clean(ws, min_row, company_col, job_col, salary_col):
    """Yield (excel_row, drop_reason | None, record | None) for every input row.

    `drop_reason` is a string explaining why the row was dropped (or None if kept).
    `record` is a dict {company, job, salary} when the row is kept (else None).
    Stops iterating when a "二、" section header is hit.
    """
    current_company = None
    for excel_row, row in enumerate(
        ws.iter_rows(min_row=min_row, values_only=True), start=min_row
    ):
        seq = row[1] if len(row) > 1 and row[1] else ''

        if isinstance(seq, str):
            if any(seq.startswith(p) for p in SECTION_PREFIXES_BREAK):
                return
            if any(seq.startswith(p) for p in SECTION_PREFIXES_SKIP):
                yield excel_row, 'section_header', None
                continue
            if seq == '序号':
                yield excel_row, 'header', None
                continue

        company = (row[company_col] or '').strip() if len(row) > company_col else ''
        job     = (row[job_col]     or '').strip() if len(row) > job_col     else ''
        salary  = (row[salary_col]  or '').strip() if len(row) > salary_col  else ''

        if not company and not job and not salary:
            yield excel_row, 'empty', None
            continue
        if not job or not salary:
            yield excel_row, 'missing_data', None
            continue
        if any(kw in company for kw in NOISE_KEYWORDS):
            yield excel_row, f'noise({company})', None
            continue

        if company:
            current_company = company
        if not current_company:
            yield excel_row, 'no_company_yet', None
            continue

        yield excel_row, None, {
            'company': current_company.replace('\n', ' '),
            'job':     job,
            'salary':  salary,
        }


def clean(xlsx_path):
    xlsx_path = Path(xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    cleaned_wb = openpyxl.Workbook()
    cleaned_wb.remove(cleaned_wb.active)
    report = [f'Clean report for {xlsx_path.name}', '']

    for sheet_name, min_row, comp_c, job_c, sal_c in CONFIGS:
        report.append(f'--- [{sheet_name}] ---')
        if sheet_name not in wb.sheetnames:
            report.append('  NOT FOUND, skipped')
            continue
        ws = wb[sheet_name]

        out_ws = cleaned_wb.create_sheet(sheet_name)
        out_ws.append(['序号', '公司', '岗位', '薪酬'])

        kept, dropped = 0, 0
        idx = 0
        for excel_row, drop_reason, record in iter_clean(ws, min_row, comp_c, job_c, sal_c):
            if drop_reason:
                report.append(f'  drop A{excel_row}: {drop_reason}')
                dropped += 1
            else:
                idx += 1
                out_ws.append([idx, record['company'], record['job'], record['salary']])
                kept += 1
        report.append(f'  kept={kept}  dropped={dropped}')

    cleaned_path = xlsx_path.with_name(f'{xlsx_path.stem}_cleaned.xlsx')
    cleaned_wb.save(cleaned_path)
    report_path = xlsx_path.with_name(f'{xlsx_path.stem}_clean_report.txt')
    report_path.write_text('\n'.join(report) + '\n', encoding='utf-8')
    print(f'cleaned → {cleaned_path}')
    print(f'report  → {report_path}')
    return cleaned_path


if __name__ == '__main__':
    clean(sys.argv[1] if len(sys.argv) > 1 else '岗位总表-20260812.xlsx')
