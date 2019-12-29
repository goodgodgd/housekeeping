import os
import pandas as pd
from tabulate import tabulate
import common as C

SALAD_RENAMER = {"금액": "출금", "결제수단": "취급점"}


class CardLabeler:
    def a_set_unknown(self, data):
        data['자동분류'] = '미정'
        print("a_set_unknown")
        return data

    def b_special_expense(self, data):
        mask = (data['출금'] > 100000) & (data['자동분류'] == '미정')
        data.loc[mask, '자동분류'] = '특별지출'
        print("b_special_expense")
        return data

    def z_set_rest(self, data):
        data.loc[data['자동분류'] == '미정', '자동분류'] = '일상지출'
        print("z_set_rest")
        return data


def preprocess(data, start_date, end_date):
    data = rearrange_columns(data)
    data = data.loc[(data['시간'] >= start_date) & (data['시간'] < end_date), :]
    data = filter_only_card(data)
    data = rearrange_inout(data)
    data['종류'] = '카드'
    print("preprocessed salad data:\n", tabulate(data.head(10), headers='keys'))
    return data


def rearrange_columns(data):
    data['시간'] = data['날짜'].astype(str) + ' ' + data['시간'].astype(str)
    data['시간'] = pd.to_datetime(data['시간'])
    data = data.rename(columns=SALAD_RENAMER)
    data = data.loc[:, C.LABEL_COLUMNS]
    return data


def filter_only_card(data):
    mask = data['취급점'].str.contains('삼성카드') | data['취급점'].str.startswith('LG U+') | \
           data['취급점'].str.startswith('뉴LG전자')
    data = data.loc[mask, :].reset_index(drop=True)
    data.loc[data['취급점'].str.startswith('삼성카드'), '취급점'] = '삼성카드'
    data.loc[data['취급점'].str.startswith('뉴LG전자'), '취급점'] = '베스트샵플러스'
    data.loc[data['취급점'].str.startswith('LG U+'), '취급점'] = 'LGU+하이Light'
    return data


def rearrange_inout(data):
    data.loc[data['출금'] > 0, '입금'] = data.loc[data['출금'] > 0, '출금']
    # 뱅크샐러드에서 결제취소한 내역이 입금으로 뜨는데 먼저 결재한 내용은 나오지 않으므로 입금내역만 삭제하면 됨
    data = data.loc[data['출금'] < 0, :]
    data.loc[:, '출금'] = -data['출금']
    assert data.loc[data['출금'] < 0, :].empty, 'expenditure must be positive'
    return data


def bank_salad(srcfile, srcsheet, dstfile, start_date, end_date):
    srcfile = os.path.join(C.DATA_ROOT, 'src', srcfile)
    dstfile = os.path.join(C.DATA_ROOT, 'result', dstfile)
    if os.path.isfile(dstfile):
        print("bank_salad file is already processed:", dstfile)
        return
    data = pd.read_excel(srcfile, sheet_name=srcsheet, header=0, converters={'날짜': str, '시간': str})

    data = preprocess(data, start_date, end_date)
    data = C.label_data(data, CardLabeler)
    print(tabulate(data.head(10), headers='keys'))
    data = data.sort_values(['자동분류', '시간'], ascending=[True, False])

    data.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)
    return data
