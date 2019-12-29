import os
from glob import glob
import pandas as pd

import common as C
from process_woori_bank import woori_bank
from process_card import bank_salad
from aggregator import aggregate_data


def merge_data(srcpattern, dstfile):
    """
    여러 파일의 데이터를 중복 없이 합쳐서 저장
    """
    srcpattern = os.path.join(C.DATA_ROOT, 'result', srcpattern)
    dstfile = os.path.join(C.DATA_ROOT, 'result', dstfile)
    srcfiles = glob(srcpattern)
    merged = []
    for file in srcfiles:
        print("[merge_data] filename:", file)
        data = pd.read_excel(file, sheet_name=C.SHEET_NAME, header=0, converters={'시간': pd.to_datetime})
        merged.append(data)
    merged = pd.concat(merged, axis=0)
    merged = merged.drop_duplicates(subset=['시간', '내용'], keep='first')
    merged = merged.sort_values(['자동분류', '시간'], ascending=[True, False])
    merged.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)


def read_files(srcfile, dstfile):
    srcfile = os.path.join(C.DATA_ROOT, 'result', srcfile)
    dstfile = os.path.join(C.DATA_ROOT, 'result', dstfile)
    srcdata = pd.read_excel(srcfile, sheet_name=C.SHEET_NAME, header=0, converters={'시간': pd.to_datetime})
    if os.path.isfile(dstfile):
        dstdata = pd.read_excel(dstfile, sheet_name=C.SHEET_NAME, header=0, converters={'시간': pd.to_datetime})
    else:
        dstdata = None
    return srcdata, dstdata


def update_auto_label(srcfile, dstfile):
    """
    자동분류가 바뀌게 되면 수동분류한 내용은 그대로 두고 나머지만 업데이트
    """
    srcdata, dstdata = read_files(srcfile, dstfile)
    dstfile = os.path.join(C.DATA_ROOT, 'result', dstfile)
    if dstdata is None:
        srcdata = srcdata.loc[:, C.CORRECT_COLUMNS]
        srcdata.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)
        return

    dstdata = dstdata.rename(columns={'비고': '기존비고'})
    dstdata = srcdata.merge(dstdata.loc[:, ['시간', '내용', '수동분류', '기존비고']], how='left', on=['시간', '내용'])
    dstdata.loc[:, '분류'] = dstdata['자동분류']
    dstdata.loc[~dstdata['수동분류'].isna(), '분류'] = dstdata.loc[~dstdata['수동분류'].isna(), '수동분류']
    dstdata.loc[~dstdata['기존비고'].isna(), '비고'] = dstdata.loc[~dstdata['기존비고'].isna(), '기존비고']
    dstdata = dstdata.loc[:, C.CORRECT_COLUMNS]
    dstdata = dstdata.sort_values(['분류', '시간'], ascending=[True, False])
    dstdata.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)


def main():
    pd.set_option('max_columns', 10)
    pd.set_option('display.width', 200)
    pd.set_option('display.float_format', lambda x: '%.1f' % x)
    start_date = pd.to_datetime("2019-01-01 00:00:00")
    end_date = pd.to_datetime("2020-01-01 00:00:00")

    woori_bank('2019_우리은행.xlsx', '2019_월급통장', 'label_2019_우리은행.xlsx', start_date, end_date)
    bank_salad('2019_뱅크샐러드.xlsx', '가계부 내역', 'label_2019_카드.xlsx', start_date, end_date)
    merge_data('label_*은행.xlsx', 'merge_은행.xlsx')
    merge_data('label_*카드.xlsx', 'merge_카드.xlsx')
    update_auto_label('merge_은행.xlsx', 'correct_은행.xlsx')
    update_auto_label('merge_카드.xlsx', 'correct_카드.xlsx')
    aggregate_data('correct_*.xlsx')


if __name__ == '__main__':
    main()
