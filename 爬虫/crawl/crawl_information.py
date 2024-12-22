import pandas as pd      # 文件操作
def acquire_code_property_information(path1:str,path2:str)->list:
    """
    need package: pandas && openpyxl
    传入两个表的地址返回list:[表头信息，重复列]
    :param path1: 要处理的表一路径
    :param path2: 要处理的表二路径
    :return: [heading1,heading2,property_columns,repeat_columns]
             表一的表头，表二的表头，表头信息，表一的行数，表二的行数，重复列
    """
    print("正在读取文件基本信息....")

    # 表一的信息
    sheet1 = pd.read_excel(path1)
    # heading为表头
    heading1 = sheet1.columns
    # rows为行数，表头默认不算
    # columns为列数
    rows1 = sheet1.shape[0]
    columns1 = sheet1.shape[1]

    # 表二的信息
    sheet2 = pd.read_excel(path2)
    # heading为表头
    heading2 = sheet2.columns
    # rows为行数，表头默认不算
    # columns为列数
    rows2 = sheet2.shape[0]
    columns2 = sheet2.shape[1]

    # 直接把表头：heading1/heading2 都先转为列表
    heading1 = list(heading1)
    heading2 = list(heading2)

    # 获取每个记录的属性
    # a.先把表一的表头给property_columns
    # b.遍历表二
    property_columns = []
    property_columns.extend(heading1)
    for count in range(len(heading2)):
        # 跳出标记
        break_flag = 0
        for i in range(len(property_columns)):
            if property_columns[i] == heading2[count]:
                break_flag = 1
                break
        if break_flag == 0:
            # break_flag=0表示遍历表1一圈，没找到相同的属性
            property_columns.append(heading2[count])

    # 获得重复列
    repeat_columns = []
    for x in range(columns1):
        for y in range(columns2):
            if heading1[x]==heading2[y]:
                # 这里因为属性是按顺序存储的所以相差columns1
                x_y = [x, y + columns1]
                # 如果记录属性相同则进行添加
                repeat_columns.append(x_y)
    print("文件基本信息读取完毕")
    return [heading1,heading2,rows1,rows2,property_columns,repeat_columns]

# 修正数据
def fix_the_data(data:list,repeat_columns:list)->list:
    """

    :param data: 游标对象读取全部数据
    :param repeat_columns: 重复列
    :return: 处理重复数据后的记录
    """
    # print(data)
    # print("游标对象读取全部数据的得到的数据类型为%s" % type(data))
    # 遍历游标对象读取全部数据
    for count in range(len(data)):
        # print(data[count])
        # print(type(data[count]))
        # 修改每行记录的类型为列表：tuples->list
        data[count] = list(data[count])
    print("游标对象读取全部数据的所有行为%d" % len(data))
    print("正在修正数据....")

    # 处理重复的数据
    for repeat_columns_index in range(len(repeat_columns)):
        for count in range(len(data)):
            if data[count][repeat_columns[repeat_columns_index][0]] is None and data[count][
                repeat_columns[repeat_columns_index][1]] is None:
                # 如果两个值都为空
                continue
            elif data[count][repeat_columns[repeat_columns_index][0]] is not None and data[count][
                repeat_columns[repeat_columns_index][1]] is None:
                # 第一个数值不为空,第二个数值为空
                continue
            elif data[count][repeat_columns[repeat_columns_index][0]] is None and data[count][
                repeat_columns[repeat_columns_index][1]] is not None:
                # 第一个数值为空,第二个数值为不为空,就交换数据
                data[count][repeat_columns[repeat_columns_index][0]] = data[count][repeat_columns[repeat_columns_index][1]]
                data[count][repeat_columns[repeat_columns_index][1]] = None
            else:
                # 这里主要处理主键数据
                # 都不为空,将第二个数据置空
                # (若出现此情况除了主键外数据错误)
                data[count][repeat_columns[repeat_columns_index][1]] = None
    print("数据修正成功")
    print("_______________________________________________________")

    return data


