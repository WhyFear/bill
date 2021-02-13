# -*- coding: utf-8 -*-
"""
@author: Zeng LH
@contact: 893843891@qq.com
@software: pycharm
@file: AddTag.py
@time: 2021/2/13 0013 11:56
@desc: 作者的版本只能一个一个手动添加，还是希望在python里面输入，稍微方便一点
"""
import sys
import openpyxl
import pandas as pd
import tkinter.filedialog


def find_tag(input_tag, tags):
    alpha = input_tag[:1]
    digit = input_tag[1:]
    row = 0  # 写死，只能是第二行
    column = ord(alpha) - ord('a') + 1
    # print(tags.iloc[row, column])
    # print(tags.iloc[row + int(digit) + 1, column])
    return tags.iloc[row, column], tags.iloc[row + int(digit) + 1, column]


def read_data(path, auto=False):
    data = pd.read_excel(path, sheet_name="明细", engine='openpyxl', keep_default_na=False)
    tags = pd.read_excel(path, sheet_name="消费类型2.0", engine='openpyxl', keep_default_na=False)
    if auto:  # 自动处理
        for row in data.index.values:
            if data.iloc[row, 12] == '' or data.iloc[row, 13] == '':
                a = data.iloc[row, 6]
                # b = data.iloc[row, 7]
                # price = data.iloc[row, 8]
                if "哈啰" in a or "滴滴出行" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("h2", tags)
                elif "西南交通大学" in a or "饿了么" in a or "西南科技大学" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("a1", tags)
                elif "龙泉山鲜果园" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("a4", tags)
                elif "红旗连锁" in a or "友宝" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("a5", tags)
                elif "成都金控数据服务有限公司" in a or "绵州通" in a or "地铁" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("h1", tags)
                elif "中国铁路" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("h3", tags)
                elif "舞东风" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("a5", tags)
                elif "天然气" in a or "中国联通" in a or "中国移动" in a or "话费" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("d2", tags)
                elif "dengyueyue" in a:
                    data.iloc[row, 12], data.iloc[row, 13] = find_tag("g8", tags)
        print("正在保存！")
        pd.DataFrame(data).to_excel("output.xlsx", sheet_name='明细', index=False, header=True)
        print("处理完毕！")
        sys.exit(0)

    else:
        try:
            with open("progress", "r", encoding="utf-8") as f:
                start = int(f.read())
            start_with_progress = input("是否从之前进度 %s 开始？(Y/N)" % start).lower()
            if start_with_progress != 'y':
                start = 0
        except:
            start = 0
        temp = dict({})
        # 先遍历之前数据获取tag
        for row in range(0, start):
            a = data.iloc[row, 6]
            b = data.iloc[row, 7]
            if a != '' and b != '':
                if a not in temp:
                    temp[a] = dict({})
                    temp[a][b] = [data.iloc[row, 12], data.iloc[row, 13]]
                elif b not in temp[a]:
                    temp[a][b] = [data.iloc[row, 12], data.iloc[row, 13]]
        # 再从上次的地方继续
        for row in range(start, data.shape[0]):
            a = data.iloc[row, 6]
            b = data.iloc[row, 7]
            if data.iloc[row, 12] == '' or data.iloc[row, 13] == '':
                print("\n第%s行数据，平台：%s，商品：%s，金额：%s元" % (row, a, b, data.iloc[row, 11]))
                try:
                    while True:
                        text = "请输入标签,输入exit或者按下CTRL+C退出: "
                        skip = True
                        tag = ""
                        if a in temp and b in temp[a]:
                            tag = temp[a][b]
                            text = "曾经有过相同数据，是否写入?回车写入%s：" % tag
                            skip = False

                        input_tag = input(text).lower()  # 只支持第一个为字母，后面跟着数字的形式。
                        if input_tag == '':
                            if skip:
                                print("跳过数据")
                            else:
                                data.iloc[row, 12], data.iloc[row, 13] = tag
                            break
                        elif input_tag == 'exit':
                            save_and_close(data, row)
                        elif input_tag[:1].isalpha() and input_tag[1:].isdigit():
                            first, second = find_tag(input_tag, tags)
                            if first == '' or second == '':
                                print("输入有误，请重新输入")
                                continue
                            data.iloc[row, 12], data.iloc[row, 13] = first, second
                            temp[a] = dict({})
                            temp[a][b] = [first, second]
                            print(first, second + "已写入")
                            break
                        else:
                            print("输入错误，请重新输入")
                except KeyboardInterrupt:
                    save_and_close(data, row)
            else:
                if a not in temp:
                    temp[a] = dict({})
                    temp[a][b] = [data.iloc[row, 12], data.iloc[row, 13]]
                elif b not in temp[a]:
                    temp[a][b] = [data.iloc[row, 12], data.iloc[row, 13]]


def save_and_close(data, row):
    print("正在保存文件，请稍候")
    pd.DataFrame(data).to_excel("output.xlsx", sheet_name='明细', index=False, header=True)
    with open("progress", "w", encoding="utf-8") as f:
        f.write(str(row))
    print("再见！")
    sys.exit(0)


if __name__ == '__main__':
    # 路径设置
    print('提示：请在弹窗中选择要标记的源文件\n')
    origin_file_path = tkinter.filedialog.askopenfilename(title='选择要处理的源数据文件：',
                                                          filetypes=[('所有文件', '.*'), ('Excel表格', '.xlsx')])
    if origin_file_path != '':
        read_data(origin_file_path, False)
