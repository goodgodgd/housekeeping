import inspect

LABEL_COLUMNS = ['시간', '종류', '내용', '취급점', '입금', '출금', '잔액', '분류', '비고']
CORRECT_COLUMNS = ['시간', '종류', '내용', '취급점', '입금', '출금', '잔액', '분류', '수동분류', '최종분류', '비고']
DATA_ROOT = "D:/Work/housekeeping/data"
SHEET_NAME = 'record'


def label_data(data, method_class):
    obj = method_class()
    methods = [getattr(obj, f[0]) for f in inspect.getmembers(obj, predicate=inspect.ismethod)]
    print("methods", methods)
    for f in methods:
        data = f(data)
    return data
