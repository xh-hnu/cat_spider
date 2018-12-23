# encoding=utf-8
"""
商品url //detail.tmall.com/item.htm?id=39310588779&skuId=3642495193234&areaId=430100&user_id=775323974&cat_id=50041298&is_b=1&rn=1f4a2c261ac23e5987794fcc03e7a9c7
天猫评论 url  https://rate.tmall.com/list_detail_rate.htm?itemId=39310588779&sellerId=775323974&currentPage=1&callback=jsonp1210
itemId：商品id   商品url 的 id
sellerId：卖家id  商品url中的user_id
page：页码
callback：
作为回调函数的一部分，这部分不加其实无伤大雅，但是对取数据会有一定的影响（就是你不能取json数据，取起来比较麻烦，数据不会很整齐），
这部分经过我的测试发现，“callback=jsonp”这部分是固定不变的，后面的数字利用random函数生成一个随机数拼接上去就可以了，
当然这个随机数尽可能给个大的范围（我就给了100到1800之间随机生成）。

评论数大于100 才爬
"""
import json
from xlutils.copy import copy
from random import randint
from urllib import parse
import xlrd, xlwt
from time import sleep
from lxml import etree
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import msvcrt


def read_goods_info(excel_name):
    # 打开文件
    workbook = xlrd.open_workbook(excel_name)
    sheets_name = workbook.sheet_names()
    new_list = []
    for item in sheets_name:
        item = workbook.sheet_by_name(item)
        col_1 = item.col_values(0)  # 店名
        col_4 = item.col_values(3)  # 评论数
        col_5 = item.col_values(4)  # 商品地址
        # 将评论数 店名 url 拼接成一个字符串
        new_list.append([col_1[index] + col_4[index] + col_5[index]
                         for index in range(1, len(col_4)) if int(col_4[index]) > 200])
        result = [n for a in new_list for n in a]  # 二维list转换成一维list
    return result


class CatCommentSpider:

    def __init__(self, url):
        self.url = url
        self.browser = webdriver.Edge()
        self.browser.maximize_window()
        self.wait = WebDriverWait(self.browser, 20)
        self.browser.get(self.url)

    def get_source(self):
        return self.browser.page_source

    def refresh_tab(self):
        self.browser.get(self.url)

    def get_json(self):
        html = etree.HTML(self.browser.page_source)
        json_str_list = html.xpath('//body/text()')  # json 数据
        json_str = ''.join(json_str_list).replace(" ", "")
        json_result = json_str[json_str.index('(') + 1: len(json_str) - 1]
        return json_result

    def close_browser(self):
        self.browser.quit()

    def open_new_tab(self, url):
        js = 'window.open("' + url + '");'
        self.browser.execute_script(js)
        handles = self.browser.window_handles  # 获取当前窗口句柄集合（列表类型）
        self.browser.switch_to.window(handles[1])
        self.browser.refresh()


if __name__ == "__main__":
    goods_info = read_goods_info('cat_goods_info.xls')
    # 得到评论url   得到商品评论保存的文件名
    save_names = []
    urls_str = []
    for good_info in goods_info:
        url_param = good_info.split('?')
        save_name = good_info.split('/')
        save_names.append(save_name[0])
        res = parse.parse_qs(url_param[1])
        url_str = 'https://rate.tmall.com/list_detail_rate.htm?itemId=' + ''.join(res['id']) + '&sellerId=' \
                  + ''.join(res['user_id']) + '&currentPage=1&callback=jsonp' + str(randint(1000, 20000))
        urls_str.append(url_str)
    # print(urls_str)
    # print(save_names)
    wb = xlrd.open_workbook(filename=u'cat_goods_comment.xls')
    newb = copy(wb)
    for url_item_index in range(1, len(urls_str)):
        # 评论数据持久化
        print(urls_str[url_item_index], save_names[url_item_index])
        sheet01 = newb.add_sheet(save_names[url_item_index])
        url_item_str = urls_str[url_item_index]
        cat_comments_spider = CatCommentSpider(url_item_str)
        json_str = cat_comments_spider.get_json()
        # dict_data = json.loads(json_str, strict=False)
        # last_page = dict_data['rateDetail']['paginator']['lastPage']  # int type
        if u'"rgv587_flag":"sm"' in json_str:
            check_url = json_str[json_str.index(u'"url":"') + 7:json_str.index(u'"}')]
            check_url = check_url.replace('amp;', '')
            print(check_url)
            cat_comments_spider.open_new_tab(check_url)
            sleep(1)
            while True:
                anything = input('输入任意字符继续')
                cat_comments_spider.close_browser()
                cat_comments_spider = CatCommentSpider(url_item_str)
                json_str = cat_comments_spider.get_json()
                break

        cat_comments_spider.close_browser()
        last_page = int(json_str[json_str.index('"lastPage":') + 11:json_str.index('"page":') - 1])
        print(url_item_index, last_page, json_str)
        sheet01.write(0, 0, str(json_str))
        sleep(1 + randint(1, 2))  # 停止几秒，防止反爬虫
        for page in range(2, last_page + 1):
            url_item_str = urls_str[url_item_index].replace('currentPage=1', 'currentPage=' +
                                                            str(page))  # 更新url
            cat_comments_spider = CatCommentSpider(url_item_str)
            json_str = cat_comments_spider.get_json()
            if u'"rgv587_flag":"sm"' in json_str:
                check_url = json_str[json_str.index(u'"url":"') + 7:json_str.index(u'"}')]
                check_url = check_url.replace('amp;', '')
                cat_comments_spider.open_new_tab(check_url)
                sleep(1)
                while True:
                    anything = input('输入任意字符继续')
                    cat_comments_spider.close_browser()
                    cat_comments_spider = CatCommentSpider(url_item_str)
                    json_str = cat_comments_spider.get_json()
                    break

            cat_comments_spider.close_browser()
            # dict_data = json.loads(json_str, strict=False)
            print(page, url_item_str, json_str)
            sheet01.write(page - 1, 0, json_str)
            sleep(1 + randint(1, 2))
        newb.save(u'cat_goods_comment.xls')

