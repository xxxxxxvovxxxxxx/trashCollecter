import xlwt
import requests
from bs4 import BeautifulSoup


class excelConfig:
    sheetName = '未命名'
    path = ''
    head = []
    rowHeight = []
    colWidth = []
    style = None

    def __init__(self, sheetName, path):
        self.sheetName = sheetName
        self.path = path
    
    def setStyle(self,style):
        self.style = style

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


def writeExcel(infoList, config):
    # need to optimize
    if infoList is None:
        return "无写入信息"
    row, col = len(infoList), len(infoList[0])
    worksbook = xlwt.Workbook(encoding='utf-8')
    Worksheet = worksbook.add_sheet(config.sheetName)

    if len(config.colWidth) != 0:
        for item in range(col):
            Worksheet.col(item).width = 256*int(config.colWidth[item])
    if len(config.head) != 0:
        for item in range(col):
            Worksheet.write(0, item, config.head[item])
    
    try:
        for i in range(0, row):
            for j in range(0, col):
                # if 1 == j:
                #     # to hyperLink
                #     Worksheet.write(
                #         i+1, j, xlwt.Formula('HYPERLINK("{0}")'.format(infoList[i][j])))
                # else:
                #     Worksheet.write(i+1, j, infoList[i][j])
                Worksheet.write(i+1, j, infoList[i][j],config.style)
        worksbook.save(config.path)
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
    config = excelConfig('高分电影榜','C:\\Users\\JYX\\Desktop\\挨到看.xls')
    config.colWidth = [20, 60, 10, 10]
    config.head = ['电影', '地址', '评分', '状态']
    print(config)
    for page in range(10):
        url = 'https://movie.douban.com/top250?start={0}&filter='.format(
            page*25)
        html = getPage(url)
        infoList.extend(getInfo(html))
    # print(infoList)
    res = writeExcel(infoList, config)
    print(res)
