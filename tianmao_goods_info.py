from random import randint
from time import sleep
import xlrd, xlwt
from lxml import etree
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from xlutils.copy import copy

'''
天猫url https://list.tmall.com/search_product.htm
?cat=50041298&s=120&q=%BD%FC%CA%D3%D1%DB%BE%B5%C6%AC&sort=s&style=g&from=mallfp..pc_1_searchbutton&active=2&industryCatId=50041298&spm=a220m.1000858.0.0.2f4e58981G3iTJ&type=pc#J_Filter
其中 s 代表页面 (page_num - 1) * 60 = s
 '''


class CatSpider:

    def __init__(self, url):
        self.url = url
        self.browser = webdriver.Edge()
        self.browser.maximize_window()
        self.wait = WebDriverWait(self.browser, 20)
        self.browser.get(self.url)
        self.source = self.browser.page_source

    def get_source(self):
        return self.source

    def get_comment(self):
        self.browser.get("")

    def get_goods_info(self, page):
        html = etree.HTML(self.source)
        productShopTemp = []
        productShop = html.xpath('//div[@class="product-iWrap"]/div[@class="productShop"]/a//text()')
        # 拼接得到正确店名
        for index in range(len(productShop)):
            if productShop[index][:1] is '\n' and productShop[index][len(productShop[index]) - 1:] is '\n':
                productShopTemp.append(productShop[index][1:len(productShop[index]) - 1])
            elif productShop[index][:1] is '\n'and productShop[index][len(productShop[index]) - 1:] is not '\n':
                productShopTemp.append(productShop[index][1:len(productShop[index])] + productShop[index+1] +
                                       productShop[index+2][:len(productShop[index+2]) - 1])

        productPrice = html.xpath('//div[@class="product-iWrap"]/p[@class="productPrice"]/em/@title')
        productTitle = html.xpath('//div[@class="product-iWrap"]/p[@class="productTitle"]/a/@title')
        productUrl = html.xpath('//div[@class="product-iWrap"]/p[@class="productTitle"]/a/@href')
        payNumPerMonth = html.xpath('//div[@class="product-iWrap"]/p[@class="productStatus"]/span/em/text()')
        commentNum = html.xpath('//div[@class="product-iWrap"]/p[@class="productStatus"]/span/a/text()')
        print(productShopTemp, productPrice, productTitle, payNumPerMonth, commentNum)
        print(len(productTitle), len(productShopTemp), len(productPrice), len(payNumPerMonth), len(commentNum))
        # 持久化
        wb = xlrd.open_workbook(filename=u'cat_goods_info.xls')
        newb = copy(wb)
        sheet_name = str(page)
        sheet01 = newb.add_sheet(sheet_name)
        # 写标题
        sheet01.write(0, 0, '店名')
        sheet01.write(0, 1, '付款人数')
        sheet01.write(0, 2, '标价')
        sheet01.write(0, 3, '评论数')
        sheet01.write(0, 4, '商品地址')

        for i in range(len(productShopTemp)):
            sheet01.write(i + 1, 0, productShopTemp[i])
            sheet01.write(i + 1, 1, payNumPerMonth[i])
            sheet01.write(i + 1, 2, productPrice[i])
            sheet01.write(i + 1, 3, commentNum[i])
            sheet01.write(i + 1, 4, productUrl[i])
        newb.save(u'cat_goods_info.xls')
        self.browser.quit()


if __name__ == "__main__":
    for page in range(1, 78):
        s = (page - 1)*60
        search_url = "https://list.tmall.com/search_product.htm" + "?cat=50041298&s=" + str(s)\
                     + "&q=%BD%FC%CA%D3%D1%DB%BE%B5%C6%AC&sort=s&style=g&from=mallfp..pc_1_searchbutton&active=2" \
                       "&industryCatId=50041298&spm=a220m.1000858.0.0.2f4e58981G3iTJ&type=pc#J_Filter "
        cat_spider = CatSpider(url=search_url)
        cat_spider.get_goods_info(page=page)
        print(search_url)
        sleep(2 + randint(1, 4))

'''
天猫url https://list.tmall.com/search_product.htm
?cat=50041298&s=120&q=%BD%FC%CA%D3%D1%DB%BE%B5%C6%AC&sort=s&style=g&from=mallfp..pc_1_searchbutton&active=2&industryCatId=50041298&spm=a220m.1000858.0.0.2f4e58981G3iTJ&type=pc#J_Filter
其中 s 代表页面 (page_num - 1) * 60 = s
 '''