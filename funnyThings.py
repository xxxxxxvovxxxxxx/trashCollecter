import trashHelper

def getInfo(html):
    author, content, vote = [], [], []
    soup = trashHelper.BeautifulSoup(html, "html.parser")
    for info in soup.find_all("div", "author clearfix"):
        author.append(info.find_next("h2").text)
    for info in soup.find_all("div", "content"):
        content.append(info.find_next("span").text)
    for info in soup.find_all("div", "stats"):
        vote.append(info.find_next("i", "number").text)
    # print(content[0])
    return list(zip(author, content, vote))


if __name__ == "__main__":
    infoList = []
    config = trashHelper.excelConfig('段子', 'C:\\Users\\JYX\\Desktop\\还行吧.xls')
    config.head = ['作者', '内容', '点赞数']
    config.colWidth = [20, 60, 10]
    style = trashHelper.xlwt.XFStyle()
    style.alignment.wrap = 1
    config.setStyle(style)
    #config.style.alignment.wrap = 1
    for page in range(1,14):
        res = trashHelper.getPage('https://www.qiushibaike.com/text/page/{0}/'.format(page))
        infoList.extend(getInfo(res))
    res = trashHelper.writeExcel(infoList, config)
    print(res)
