from snownlp import SnowNLP
import matplotlib.pyplot as plt

file = "appendComment.txt"  # 评论数据文件
bad = "bad.txt"  # 差评保存文件
good = "good.txt"
comment_dict = {}
z = 0  # 总数
with open(file, "r", encoding="gbk", errors='ignore') as text:
    bad = open(bad, "w", encoding="utf-8")
    good = open(good, 'w', encoding='utf-8')
    for comment in text:
        z += 1
        s = SnowNLP(comment)  # 文本分析
        s = s.sentiments  # 情感系数
        if s >= 0.66:
            good.write(comment)
            comment_dict['好评数'] = comment_dict.get('好评数', 0) + 1
        elif 0.66 > s > 0.33:
            comment_dict['中评数'] = comment_dict.get('中评数', 0) + 1
        else:
            bad.write(comment)  # 写入差评数
            comment_dict['差评数'] = comment_dict.get('差评数', 0) + 1
    bad.close()
    good.close()
for a, b in comment_dict.items():
    pctb = b/z*100
    # ha 文字指定在柱体中间， va指定文字位置 fontsize指定文字体大小
    plt.text(a, b + 0.05, '%.0f%%' % pctb, ha='center', va='bottom', fontsize=11)
# 设置X轴Y轴数据，两者都可以是list或者tuple
x_axis = tuple(comment_dict.keys())
y_axis = tuple(comment_dict.values())
plt.bar(x_axis, y_axis, color='rgb')  # 如果不指定color，所有的柱体都会是一个颜色

plt.savefig("cat_spider_img\\append_comment_analyze.png")  # 保存为图片
plt.show()
print(comment_dict)
