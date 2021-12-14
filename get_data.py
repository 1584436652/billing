from study import OperatingData
import re


class MyDemo(OperatingData):

    @staticmethod
    def remove(text):
        pattern_symbol = re.compile('[’!"#$%&\'()*+,-/:;<=>?@[\\]^_`{|}~，。, ：]')
        pattern_character = re.compile('[\u4e00-\u9fa5]')
        result = re.sub(pattern_character, '', text)
        res = re.sub(pattern_symbol, '', result)
        return res

    def get(self, su):
        for ke, va in su.items():
            if ke == "备注":
                for i, j in enumerate(va):
                    va[i] = self.remove(j)


def delete(obj):
    data = []
    sh = obj.load_data('test.xlsx', "Sheet1")
    for k in sh:
        k[-1] = obj.remove(k[-1])
        if len(k[-1]) < 25:
            data.append(k)
    print(data)


if __name__ == '__main__':
    xl = MyDemo()
    delete(xl)
