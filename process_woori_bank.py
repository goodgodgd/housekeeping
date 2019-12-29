import os
import pandas as pd
from tabulate import tabulate
import common as C

BANK_RENAMER = {"기재내용": "내용", "찾으신금액(원)": "출금", "맡기신금액(원)": "입금", "거래후잔액(원)": "잔액", "메모": "비고"}
CATEGORIES = "월급, 기타수입, 일상지출, 정기지출, 특별지출, 이체, 대출, 투자, 저축, 카드, 예외"


class BankLabeler:
    def a_set_unknown(self, data):
        data['분류'] = '미정'
        print("a_set_unknown")
        return data

    def b_salary(self, data):
        mask = (data['입금'] > 0) & (data['분류'] == '미정') & (data['내용'] == '순천향급여')
        data.loc[mask, '분류'] = '월급'
        print("b_salary")
        return data

    def c_other_income(self, data):
        mask = (data['입금'] > 0) & (data['분류'] == '미정') & \
               (data['내용'].str.contains('엘지전자') | data['내용'].str.contains('상담수당') |
                data['내용'].str.startswith('순천향') | data['내용'].str.startswith('SCH'))
        data.loc[mask, '분류'] = '기타수입'
        print("c_other_income")
        return data

    def d_card_payment(self, data):
        mask = (data['출금'] > 0) & (data['분류'] == '미정') & (data['내용'].isin(['신한카드', '신한카드(주)', '삼성카드']))
        data.loc[mask, '분류'] = '카드'
        print("d_card_payment")
        return data

    def e_saving(self, data):
        mask = (data['출금'] > 0) & (data['분류'] == '미정') & (data['내용'].str.contains('청약통'))
        data.loc[mask, '분류'] = '저축'
        print("e_saving")
        return data

    def f_regular_expense(self, data):
        mask = (data['출금'] > 0) & (data['분류'] == '미정') & \
               (data['내용'].isin(['월고정지출', '탭+부모님', '카카-최혁두', '교통미용']) | data['내용'].str.startswith('교보'))
        data.loc[mask, '분류'] = '정기지출'
        print("f_regular_expense")
        return data

    def g_special_expense(self, data):
        mask = (data['출금'] > 100000) & (data['분류'] == '미정')
        data.loc[mask, '분류'] = '특별지출'
        print("g_special_expense")
        return data

    def z_set_rest(self, data):
        data.loc[(data['출금'] > 0) & (data['분류'] == '미정'), '분류'] = '일상지출'
        data.loc[(data['입금'] > 0) & (data['분류'] == '미정'), '분류'] = '이체'
        print("z_set_rest")
        return data


def preprocess(data, start_date, end_date):
    data.loc[:, '시간'] = pd.to_datetime(data['거래일시'])
    data = data.rename(columns=BANK_RENAMER)
    data = data.loc[:, C.LABEL_COLUMNS]
    mask = (data['시간'] >= start_date) & (data['시간'] < end_date)
    data.loc[:, '종류'] = '은행'
    data = data.loc[mask, :]
    print(tabulate(data.head(10), headers='keys'))
    return data


def woori_bank(srcfile, srcsheet, dstfile, start_date, end_date):
    srcfile = os.path.join(C.DATA_ROOT, 'src', srcfile)
    dstfile = os.path.join(C.DATA_ROOT, 'result', dstfile)
    data = pd.read_excel(srcfile, sheet_name=srcsheet, header=8)

    data = preprocess(data, start_date, end_date)
    data = C.label_data(data, BankLabeler)
    print(tabulate(data.head(10), headers='keys'))
    data = data.sort_values(['분류', '시간'], ascending=[True, False])

    data.to_excel(dstfile, sheet_name=C.SHEET_NAME, index=False)
    return data
