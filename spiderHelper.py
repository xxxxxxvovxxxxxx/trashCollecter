import xlwt
import requests
from bs4 import BeautifulSoup


def getPage(url):
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36"}
    res = requests.get(url, headers=header)
    res.encoding = 'utf-8'
    return res.text


def getInfo(res):
    fTitle, fUrl, rate, payable = [], [], [], []
    soup = BeautifulSoup(res, "html.parser")
    for info in soup.find_all("div", "info"):
        fTitle.append(info.find_next("span", "title").text)
        fUrl.append(info.find_next('a')['href'])
        rate.append(info.find_next("span", "rating_num").text)
        try:
            payable.append(info.find_next('span', 'playable').text)
        except AttributeError:
            payable.append('未知')
    return list(zip(fTitle, fUrl, rate, payable))


def writeExcel(infoList, sheetName, path):
    # need to optimize
    head = ['电影', '地址', '评分', '状态']
    if len(infoList) == 0:
        return "无写入信息"
    row, col = len(infoList), len(infoList[0])
    worksbook = xlwt.Workbook(encoding='utf-8')
    Worksheet = worksbook.add_sheet(sheetName)
    Worksheet.col(0).width, Worksheet.col(1).width, Worksheet.col(
        2).width, Worksheet.col(3).width = 256*20, 256*60, 256*10, 256*10
    for item in range(col):
        Worksheet.write(0, item, head[item])
    try:
        for i in range(0, row):
            for j in range(0, col):
                if 1 == j:
                    # to hyperLink
                    Worksheet.write(
                        i+1, j, xlwt.Formula('HYPERLINK("{0}")'.format(infoList[i][j])))
                else:
                    Worksheet.write(i+1, j, infoList[i][j])
        worksbook.save(path)
    except Exception as e:
        return "出现意外："+str(e.args)
    return '完美完成'


def sysLog():
    # TODO
    return


def test():
    # for unit test
    return


if __name__ == "__main__":
    infoList = []
    for page in range(10):
        url = 'https://movie.douban.com/top250?start={0}&filter='.format(
            page*25)
        html = getPage(url)
        infoList.extend(getInfo(html))
    # print(infoList)
    res = writeExcel(infoList, '高分电影榜', 'C:\\Users\\JYX\\Desktop\\挨到看.xls')
    print(res)
