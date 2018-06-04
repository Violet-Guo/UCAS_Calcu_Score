#coding:utf-8
import xlrd
import xlwt

classes = ['硕士学位英语', '足球普修', '女子自由泳', '硕士学位英语（免修）',
           '太极拳（男女混合）', '健美操（男女混班）', '吴氏太极拳', '英语B免修考试',
           '乒乓球（男女混班）', '排舞（男女混班）', '软式排球普修（男女混班）',
           '中国马克思主义与当代', '北欧行走与拓展（男女混班）', '男子自由泳', '男子蛙泳',
           '男子篮球普修', '自然辩证法概论', '英语B', '瑜伽（男女混班）', '男子健身',
           '网球', '女子自游泳', '健身气功《八段锦、易筋经》（男女混班）', '太极拳（男女混班）',
           '中国特色社会主义理论与实践研究',	'排舞[Linedance]（男女混班）', '女子蛙泳',
           '羽毛球（男女混班）', '男子健身健美']
score = ['优秀', '良好', '合格']

"""
:param file: 文件路径
:param sheet_index: 读取的工作表索引
:return: 二维数组
"""
def read_file(file, sheet_index = 0):
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(sheet_index)
    print("工作表名称:", sheet.name)
    print("行数:", sheet.nrows)
    print("列数:", sheet.ncols)

    data = []
    for i in range(0, sheet.nrows):
        tmp = [sheet.row_values(i)[0], sheet.row_values(i)[1],
               sheet.row_values(i)[7], sheet.row_values(i)[10],
               sheet.row_values(i)[12]]
        data.append(tmp)

    return data

def calcu(data):
    length = len(data)

    result = []
    result.append(['学号', '姓名', '学分总分', '总', '平均', '能否评优'])
    print(result[0])

    student_id = data[1][0]
    student_name = data[1][1]
    student_xuefen = 0
    sum = 0
    cnt = 1
    flag = 1

    for i in range(1, length):
        if data[i][0] == student_id:
            if data[i][2] in classes or data[i][4] in score:
                continue
            if float(data[i][4]) < 70.0:
                flag = 0
            sum = sum + float(data[i][3]) * float(data[i][4])
            student_xuefen += float(data[i][3])
        else:
            cnt += 1
            avg = sum / float(student_xuefen)
            if flag == 1:
                tmp = [student_id, student_name, student_xuefen, sum, avg, 'Yes']
            else:
                tmp = [student_id, student_name, student_xuefen, sum, avg, 'No']
            print(tmp)
            result.append(tmp)
            flag = 1
            if i < length:
                sum = 0
                student_id = data[i+1][0]
                student_name = data[i+1][1]
                sum = float(data[i][3]) * float(data[i][4])
                student_xuefen = float(data[i][3])
                if float(data[i][4]) <= 70.0:
                    flag = 0

    avg = sum / float(student_xuefen)
    if flag == 1:
        tmp = [student_id, student_name, student_xuefen, sum, avg, 'Yes']
    else:
        tmp = [student_id, student_name, student_xuefen, sum, avg, 'No']

    print(tmp)
    result.append(tmp)

    return result

def write_file(result):
    workbook = xlwt.Workbook(encoding='utf-8')
    table = workbook.add_sheet("706 score")
    length = len(result)
    print(length)

    for i in range(0, length):
        print(result[i])
        for j in range(0, 6):
            table.write(i, j, result[i][j])

    workbook.save('result.xls')


if __name__ == '__main__':
    data = read_file("../data/成绩汇总表706.xls")
    print(data)
    result = calcu(data)
    write_file(result)