import os
import common as C
from glob import glob
import pandas as pd
import matplotlib.pyplot as plt

from matplotlib import font_manager, rc
font_name = font_manager.FontProperties(fname="c:/Windows/Fonts/malgun.ttf").get_name()
rc('font', family=font_name)


def aggregate_data(file_pattern):
    data = load_data(file_pattern)
    data.loc[:, 'YearMonth'] = data['시간'].map(lambda x: x.strftime('%Y%m'))
    aggdata = data.loc[:, ['YearMonth', '분류', '입금', '출금']].groupby(by=['YearMonth', '분류']).sum()
    aggdata = aggdata.reset_index()
    aggdata = fill_missing_rows(aggdata)
    aggdata = integrate_inout(aggdata)
    aggdata = aggdata.sort_values(['분류', 'YearMonth'])
    dstfile = os.path.join(C.DATA_ROOT, 'result', 'aggregation.xlsx')
    aggdata.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)
    sum_io = sum_ins_outs(aggdata)
    plot_graphs(aggdata, sum_io)


def load_data(file_pattern):
    file_pattern = os.path.join(C.DATA_ROOT, 'result', file_pattern)
    files = glob(file_pattern)
    total = []
    for file in files:
        print("[aggregate_data] filename:", file)
        data = pd.read_excel(file, sheet_name=C.SHEET_NAME, header=0, converters={'시간': pd.to_datetime})
        total.append(data)
    total = pd.concat(total, axis=0, ignore_index=True)
    return total


def fill_missing_rows(aggdata):
    yearmonth = aggdata['YearMonth'].unique()
    categories = aggdata['분류'].unique()
    ymfull = []
    ctfull = []
    for ym in yearmonth:
        ymfull += [ym]*len(categories)
        ctfull.extend(categories)
    full_rows = [[ym, ct] for ym, ct in zip(ymfull, ctfull)]
    full_rows = pd.DataFrame(data=full_rows, columns=['YearMonth', '분류'])
    aggdata = full_rows.merge(aggdata, how='left', on=['YearMonth', '분류'])
    aggdata = aggdata.fillna(0)
    aggdata.loc[:, ['입금', '출금']] = aggdata.loc[:, ['입금', '출금']].astype(int)
    return aggdata


def integrate_inout(aggdata):
    """
    분류에 따라 '입출합' 계산: 입금이 중심인 분류는 '입금-출금'으로, 출금이 중심인 항목은 '출금-입금'으로
    """
    in_positive = ['월급', '기타수입', '이체']
    out_positive = ['일상지출', '정기지출', '특별지출', '카드', '대출', '투자', '저축']
    aggdata.loc[:, '입출합'] = 0
    for categ in in_positive:
        mask = aggdata['분류'] == categ
        aggdata.loc[mask, '입출합'] = aggdata.loc[mask, '입금'] - aggdata.loc[mask, '출금']

    for categ in out_positive:
        mask = aggdata['분류'] == categ
        aggdata.loc[mask, '입출합'] = aggdata.loc[mask, '출금'] - aggdata.loc[mask, '입금']
    return aggdata


def sum_ins_outs(aggdata):
    in_categs = ['월급', '기타수입', '이체']
    out_categs = ['일상지출', '정기지출', '특별지출']
    sum_in = aggdata.loc[aggdata['분류'].isin(in_categs), ['YearMonth', '입출합']].groupby('YearMonth').sum()
    sum_out = aggdata.loc[aggdata['분류'].isin(out_categs), ['YearMonth', '입출합']].groupby('YearMonth').sum()
    sum_sv = aggdata.loc[aggdata['분류'] == '저축', ['YearMonth', '입출합']].groupby('YearMonth').sum()
    sum_in = sum_in.rename(columns={'입출합': '수입합'})
    sum_out = sum_out.rename(columns={'입출합': '지출합'})
    sum_sv = sum_sv.rename(columns={'입출합': '저축합'})
    sum_io = sum_in.merge(sum_out, how='left', on='YearMonth')
    sum_io = sum_io.merge(sum_sv, how='left', on='YearMonth')
    total_io = sum_io.sum()
    print("-"*50, f"\n전체 입출금\n{total_io}\n", "-"*50)
    total_io.to_csv(os.path.join(C.DATA_ROOT, 'result', 'total_io.csv'), encoding='utf-8')
    return sum_io


def plot_graphs(aggdata, sum_io):
    categories = ['월급', '기타수입', '이체', '일상지출', '정기지출', '특별지출', '카드', '저축']
    fig, axes = plt.subplots(2)
    for categ in categories:
        data = aggdata.loc[aggdata['분류'] == categ, ['YearMonth', '입출합']]
        axes[0].plot(data['YearMonth'], data['입출합'], label=categ)
    axes[0].legend()
    axes[0].grid(True)
    axes[1].plot(sum_io.index, sum_io['수입합'], label='수입합')
    axes[1].plot(sum_io.index, sum_io['지출합'], label='지출합')
    axes[1].plot(sum_io.index, sum_io['저축합'], label='저축합')
    axes[1].legend()
    axes[1].grid(True)
    plt.tight_layout()
    plt.show()

