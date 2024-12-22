import pymssql     #SQL Server
def create_crawl_database(server:str,username:str,password:str)->None:
    """
     need packages: pymssql
    :param server:   服务器名称
    :param username: 登录名
    :param password: 登录密码
    :return:  创建CrawlDB数据库
    """
    try:
        # 连接到数据库
        conn = pymssql.connect(server=server, user=username, password=password, autocommit=True)
        print(f"您好！{username},登录成功")
        print("开始创建CrawlDB数据库")
    except Exception as error:
        print(error)
        exit(-1)

    sql = "CREATE DATABASE CrawlDB"
    cursor = conn.cursor()
    cursor.execute(sql)

    # 关闭数据库连接
    cursor.close()
    conn.close()
    print("CrawlDB数据库创建成功")

import pandas as pd
from sqlalchemy import create_engine
def write_excel_to_sql(path:str,index:int,server_list:list)->None:
    """
    need package: pandas && pyodbc
    use method: create_engine
    :param path:         文件路径
    :param index:        文件下标索引
    :param server_list:  进入SQL_Server服务器的基本信息
    :return:  None 将文件导入数据库
    """

    # 读取 Excel 文件
    file_path = path
    try:
        excel_data = pd.read_excel(file_path)
    except Exception as error:
        print(error)
        print("文件名无效")
        exit(-1)

    # 连接到 SQL Server 数据库
    server = server_list[0]             # 替换为你的 SQL Server 名称
    username = server_list[1]           # 替换为你的用户名
    password = server_list[2]           # 替换为你的密码
    database = server_list[3]           # 替换为你的数据库名称
    driver = '{ODBC Driver 17 for SQL Server}'  # 确保安装了 ODBC Driver 17 for SQL Server

    connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    # 导入引擎
    engine = create_engine(f'mssql+pyodbc:///?odbc_connect={connection_string}')

    # 将数据插入到数据库中
    table_name = '爬取信息%d' % index  # 替换 'your_table_name' 为你的表名
    excel_data.to_sql(table_name, con=engine, if_exists='replace', index=False)

    print(f"Data has been successfully imported to the table '{table_name}' in the database.")

def acquire_sql_output_data(server_list:list,heading1:list,heading2:list,rows1:int,rows2:int)->list:
    """
    :param server_list:   服务器信息
    :param heading1:      表一的表头
    :param heading2:      表二的表头
    :param rows1:         表一的行数
    :param rows2:         表二的行数
    :return:              sql全连接后的数据
    """
    # 构建连接对象
    connect = pymssql.connect(
        server=server_list[0],     # SQL Server 名称
        user=server_list[1],       # 用户名
        password=server_list[2],   # 登录密码
        database=server_list[3]    # 连接的数据库
    )
    # 生成sql语句
    part_sql=''
    for count in range(len(heading1)):
        if heading1[count][-1]==')':
            # 修正带单位的数据
            heading1[count]=f"[{heading1[count]}]"
        temp_sql =f'convert(nvarchar(50),爬取信息1.{heading1[count]}),'
        part_sql=part_sql+temp_sql

    for count in range(len(heading2)):
        if heading2[count][-1]==')':
            # 修正带单位的数据
            heading2[count]=f"[{heading2[count]}]"
        temp_sql =f'convert(nvarchar(50),爬取信息2.{heading2[count]}),'
        if count==len(heading2)-1:
            # 最后一次不再加','
            temp_sql=f'convert(nvarchar(50),爬取信息2.{heading2[count]})'
            part_sql = part_sql + temp_sql
        else:
            part_sql=part_sql+temp_sql
    # print(part_sql)
    # 获取行数大的表的索引以及其第一属性
    which_max_code= 1 if rows1>rows2 else 2
    max_code_first_property= heading1[0] if rows1>rows2 else heading2[0]

    # 获得游标对象
    cursor = connect.cursor()

    # 编写SQL查询语句
    # 有n前缀的，n表示Unicode字符，即所有字符都占两个字节,nchar,nvarchar
    # 例如一个一位数字也按两个字符储存，nvarchar(1)只能储存一个一位数字，或一个汉字
    # sql= "select convert(nvarchar(4),SNO),convert(nvarchar(4),name),convert(nvarchar(4),sex),age,convert(nvarchar(4),special)  from student"
    # 表的连接

    sql =f"""
        USE CrawlDB
        SELECT  {part_sql}
        FROM 爬取信息1 
        FUll JOIN 爬取信息2 
        ON (爬取信息1.{heading1[0]}=爬取信息2.{heading2[0]})
        ORDER BY 爬取信息{which_max_code}.{max_code_first_property} ASC
        """
    # 执行SQL语句
    print(sql)
    print("sql语句生成完毕")

    cursor.execute(sql)
    # 读取数据
    data = cursor.fetchall()

    # 关闭数据库连接
    cursor.close()
    connect.close()
    print("表格FUll JOIN 完成")
    return data

def drop_table(server_list:list,table_name:str)->None:
    """
    删除数据表
    :param server_list: 服务器信息
    :param table_name:  要删除的数据表
    :return: None
    """

    # 构建连接对象
    connect = pymssql.connect(
        server=server_list[0],    # SQL Server 名称
        user=server_list[1],      # 用户名
        password=server_list[2],  # 登录密码
        database=server_list[3]   # 连接的数据库
    )

    # 获得游标对象
    cursor = connect.cursor()
    sql=f"drop table {table_name}"

    # 执行SQL语句
    cursor.execute(sql)

    # 提交事务并关闭数据库连接
    connect.commit()
    cursor.close()
    connect.close()
    print(f"SQL_server.dbo{table_name}删除成功")
    return

