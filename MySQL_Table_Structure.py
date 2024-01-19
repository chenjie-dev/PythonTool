import pandas as pd
import pymysql
from openpyxl.workbook import Workbook


def remove_digits(string):
    result = ''
    for char in string:
        if not char.isdigit():  # 检查字符是否为数字
            result += char
    return result


def write_query_result_to_excel(host, username, password, database, querySQL, excel_filename):
    # 连接到数据库
    connection = pymysql.connect(
        host=host,
        user=username,
        password=password,
        database=database
    )

    try:
        query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = {};".format(
            '\'' + database + '\'')
        cursor = connection.cursor()
        cursor.execute(query)
        # 获取所有结果
        results = cursor.fetchall()

        # 创建一个新的Excel文件
        workbook = Workbook()

        tableNameWithoutDigits = set()

        for result in results:
            for table in result:
                # print(table)
                tableNameWithoutDigit = remove_digits(table)
                if tableNameWithoutDigit in tableNameWithoutDigits:
                    continue  # 如果已存在，则跳过当前循环
                else:
                    print(table)
                    tableNameWithoutDigits.add(tableNameWithoutDigit)  # 添加到集合中

                sql = querySQL
                sql = sql.replace('TABLENAME', table)

                # 执行查询语句并获取结果
                df = pd.read_sql_query(sql, connection)

                # 将查询结果写入Excel文件的不同工作表中
                # 创建新的工作表
                sheet = workbook.create_sheet(title=table)
                for index, row in df.iterrows():
                    data = list(row)
                    sheet.append(data)

        # 保存Excel文件
        workbook.save(excel_filename)
        print(f"查询结果已成功写入Excel文件：{excel_filename}")

    except Exception as e:
        print(f"写入Excel文件时出现错误：{e}")

    finally:
        # 关闭数据库连接
        connection.close()


# 使用示例
host = '10.31.63.16'
username = 'root'
password = 'iztNN49c5B'
database = 'backend_offlinecalc_db'
querySQL = '''
                SELECT
                    @row_number := @row_number + 1 AS 行号,
                    t.*
                FROM
                    (SELECT
                        COLUMN_NAME AS 字段名,
                        COLUMN_COMMENT AS 中文名,
                        DATA_TYPE AS 类型,
                        SUBSTRING_INDEX(SUBSTRING_INDEX(COLUMN_TYPE, '(', -1), ')', 1) AS 长度,
                        COLUMN_COMMENT AS 说明
                    FROM
                        INFORMATION_SCHEMA.COLUMNS
                    WHERE
                        TABLE_SCHEMA = {}
                        AND TABLE_NAME = 'TABLENAME'
                    ORDER BY
                        COLUMN_NAME) AS t,
                    (SELECT @row_number := 0) AS r;
            '''.format('\'' + database + '\'')

excel_filename = database + '.xlsx'

if __name__ == '__main__':
    write_query_result_to_excel(host, username, password, database, querySQL, excel_filename)
