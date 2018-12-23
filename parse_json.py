# encoding=utf-8

import json
import re
import matplotlib.pyplot as plt
import xlrd


def read_json_info(excel_name):
    # 打开文件
    workbook = xlrd.open_workbook(excel_name)
    sheets_name = workbook.sheet_names()
    dushus = []  # 储存度数
    count = 0  # 统计评论总数
    appand_comment_count = 0  # 统计追评总数
    comment_time_list = []  # 储存评论日期
    dushu_dict = {}  # 统计相应度数的人数
    dushu_map = {"['400', '650']": "400度-650度", "['650', '850']": "650度-850度", "['400']": "400度以下", "['850']": "850度以上"}
    with open("appendComment.txt", 'w') as f:
        for item in sheets_name:
            print(item)
            item = workbook.sheet_by_name(item)
            jsons = item.col_values(0)  # json字符串
            for index in range(1, len(jsons)):
                json_data = jsons[index]
                try:
                    dict_data = json.loads(json_data, strict=False)
                    for json_item in dict_data['rateDetail']['rateList']:
                        if json_item['appendComment'] is not None:
                            appand_comment_count = appand_comment_count + 1
                            f.write(json_item['appendComment']['content'] + "\n")
                        comment_time = json_item['rateDate'].split()[0]  # 评论日期
                        dushu = json_item['auctionSku'][json_item['auctionSku'].index('镜片适合度数') + 7:]
                        dushu = re.findall(r"\d+\.?\d*", dushu)  # 提取度数
                        dushus.append(dushu)
                        comment_time_list.append(comment_time)
                        count = count + 1
                except Exception:
                    continue
        print('评论总数', str(count), str(appand_comment_count))

        for key in dushus:
            dushu_dict[dushu_map[str(key)]] = dushu_dict.get(dushu_map[str(key)], 0) + 1  # 数据清洗
        comment_time_dict = {}
        for key in comment_time_list:
            comment_time_dict[str(key)] = comment_time_dict.get(str(key), 0) + 1
        print(dushu_dict)  # 度数分布源数据
        print(comment_time_dict)  # 评论时间源数据
        # 饼状图
        img1 = plt.figure(1)
        source_data = sorted(dushu_dict.items(), key=lambda x: x[1], reverse=True)  # 以人数排序，正序
        print(source_data)
        labels = [source_data[i][0] for i in range(len(source_data))]  # 设置标签
        fracs = [source_data[i][1] for i in range(len(source_data))]
        explode = [x * 0.01 for x in range(len(source_data))]  # 与labels一一对应，数值越大离中心区越远
        plt.axes(aspect=1)  # 设置X轴 Y轴比例
        # labeldistance标签离中心距离  pctdistance百分百数据离中心区距离 autopct 百分比的格式 shadow阴影
        plt.pie(x=fracs, labels=labels, explode=explode, autopct='%3.1f %%',
                shadow=False, labeldistance=1.1, startangle=0, pctdistance=0.8, center=(-1, 0))
        # 控制位置：bbox_to_anchor数组中，前者控制左右移动，后者控制上下。ncol控制 图例所列的列数。默认值为1。fancybox 圆边
        plt.legend(loc=7, bbox_to_anchor=(1.2, 1.1), ncol=3, fancybox=True, shadow=True, fontsize=10)
        plt.savefig("cat_spider_img\dushu.png")
        img1.show()

        # 柱状图
        img2 = plt.figure(2)
        comment_time_dict = {k: v for k, v in comment_time_dict.items() if v > 80}  # 过滤购买人数低于25人的日期
        for a, b in comment_time_dict.items():
            plt.text(a, b + 0.05, '%.0f' % b, ha='center', va='bottom',
                     fontsize=11)  # ha 文字指定在柱体中间， va指定文字位置 fontsize指定文字体大小
        # 设置X轴Y轴数据，两者都可以是list或者tuple
        x_axis = tuple(comment_time_dict.keys())
        y_axis = tuple(comment_time_dict.values())
        plt.bar(x_axis, y_axis, color='rgb')  # 如果不指定color，所有的柱体都会是一个颜色
        plt.xlabel(u"日期")  # 指定x轴描述信息
        plt.ylabel(u"购买人数")  # 指定y轴描述信息
        plt.xticks(rotation=90)
        plt.savefig("cat_spider_img\comment_time.png")
        img2.show()


read_json_info('cat_goods_comment.xls')
