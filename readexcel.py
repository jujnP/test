import openpyxl


class TestData:
    pass


class ReadExcel:
    def __init__(self, filename, sheet_name):
        """
        输入文件名(带路径)，文件内的表单名
        :param filename:  文件名
        :param sheet_name: 表单名
        """
        self.filename = filename
        self.sheet_name = sheet_name

    def open(self):
        self.file = openpyxl.load_workbook(self.filename)
        self.raw_data = self.file[self.sheet_name]

    def read_data(self):
        self.open()
        raw_data = list(self.raw_data.rows)

        data_keys = []
        for i in raw_data[0]:
            data_keys.append(i.value)
        # print(data_keys)
        test_data = []
        for i in raw_data[1:]:
            data_value = []
            for y in i:
                data_value.append(y.value)
            test_data.append(dict(zip(data_keys, data_value)))

        return test_data

    def close(self):
        self.file.close()

    def write_date(self, row, column, value):
        self.open()
        self.raw_data.cell(row=row, column=column, value=value)
        self.file.save(self.filename)
        self.close()


if __name__ == '__main__':
    case = ReadExcel(r"t.xlsx", "Sheet1").read_data()
    print(case)
