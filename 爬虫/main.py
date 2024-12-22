"""
将要提取的文件命名为爬取信息1.xlsx,爬取信息2.xlsx....爬取信息n.xlsx，
并时将主键属性放在第一列
"""
from crawl import *

def model_function(server_list:list)->None:
    """
    文件读写复用接口
    :param server_list:  服务器信息
    :return: None
    """
    # 写入数据
    path1 = '爬取信息1.xlsx'
    index1 = 1
    sql_server_operate.write_excel_to_sql(path1, index1, server_list)
    path2 = '爬取信息2.xlsx'
    index2 = 2
    sql_server_operate.write_excel_to_sql(path2, index2, server_list)
    # 获取两个表格的基本信息
    comprehensive_information = crawl_information.acquire_code_property_information(path1, path2)
    heading1 = comprehensive_information[0]          # 表一的表头
    heading2 = comprehensive_information[1]          # 表二的表头
    rows1    = comprehensive_information[2]          # 表一的行数，
    rows2    = comprehensive_information[3]          # 表二的行数
    property_columns = comprehensive_information[4]  # 表头信息
    repeat_columns   = comprehensive_information[5]  # 重复列

    # 获得第一手sql处理后的数据
    data = sql_server_operate.acquire_sql_output_data(server_list, heading1, heading2, rows1, rows2)

    # 修正数据
    data = crawl_information.fix_the_data(data, repeat_columns)

    try:
        # 写入文件
        file_operate.output_file(property_columns, repeat_columns, data)
        # print("表格自动填充完毕，请到当前所在文件夹查看")
    except Exception as error:
        print(f"{error}")

    sql_server_operate.drop_table(server_list,"爬取信息1")
    sql_server_operate.drop_table(server_list,"爬取信息2")

    return

# server=input("server=")
# username=input("username=")
# password=input("password=")
# quantity=int(input("quantity="))

server='.'        # 服务器名称
username='Test'   # 用户名
password='123456' # 登录密码
quantity=4        # 总的Excel文件数量

# 创建crawl_database
sql_server_operate.create_crawl_database(server,username,password)

# 服务器信息
server_list_=[server,username,password,'CrawlDB']

model_function(server_list_)

start=3      # 初值
while start<=quantity:
    # 删除用过的表格时刻只保留表一，表二
    file_operate.remove_excel_file("爬取信息1.xlsx")
    file_operate.remove_excel_file("爬取信息2.xlsx")

    # 生成表一
    old_file_path_1= "爬取信息汇总Output.xlsx"
    new_file_path_1= "爬取信息1.xlsx"
    file_operate.rename_file(old_file_path_1, new_file_path_1)

    # 生成表二
    old_file_path_2= f"爬取信息{start}.xlsx"
    new_file_path_2= "爬取信息2.xlsx"
    file_operate.rename_file(old_file_path_2, new_file_path_2)

    model_function(server_list_)
    start+=1

file_operate.remove_excel_file("爬取信息1.xlsx")
file_operate.remove_excel_file("爬取信息2.xlsx")

print("_______________________________________________________")
print("表格自动填充完毕，请到当前所在文件夹查看")
print("Press any key to continue ",end='')
input()



