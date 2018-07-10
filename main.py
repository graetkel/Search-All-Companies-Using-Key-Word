
from openpyxl import load_workbook
import re
import requests
from bs4 import BeautifulSoup


companyList = []
ignoreText1 = "FiscalYear 2018"
ignoreText2 = "Supplier Name (supplier number)"


wb = load_workbook(filename='DSLB.xlsx', read_only=True)
ws = wb['Sheet1']

for row in ws.rows:
    for cell in row:
        cv = cell.value
        if (cv is not None and cv != ignoreText1 and cv != ignoreText2):
            cv = re.sub("\([^>]+\)", "", cv)
            # print(cv)
            companyList.append(cv)




def companyTotalResults(theCompanyName):
    # companyName = "Dicks Sporting Goods"
    companyName = theCompanyName

    companyNamePlus = re.sub(" ", "+", companyName)
    companyNamePlus = re.sub("\n", "", companyNamePlus)
    companyNamePlus = companyNamePlus.lower()

    # print(companyNamePlus)

    myLink = "https://www.google.com/search?q=" + "\"" + companyNamePlus + "\"" + "+\"prison+labor\""
    page = requests.get(myLink)

    soup = BeautifulSoup(page.content, 'html.parser')

    container = soup.find(id="resultStats")
    tempCount = "" + str(container)

    # print(tempCount)

    tempCount = tempCount.replace("<div class=\"sd\" id=\"resultStats\">About ", "")
    tempCount = tempCount.replace(" results</div>", "")
    tempCount = tempCount.replace(",", "")

    result = 0
    try:
        result = int(tempCount)
    except:
        result = 0

    return result



companyCountList = []

for index in range(len(companyList)):
    # print(companyList[index] + " " + str(companyTotalResults(companyList[index])))
    companyCountList.append(companyTotalResults(companyList[index]))




for i in range(len(companyList)):
    biggestCount = 0
    for k in range(len(companyList)):
        if (int(companyCountList[k]) > biggestCount):
            biggestCount = companyCountList[k]

    tempIndex = companyCountList.index(biggestCount)
    print(companyList[tempIndex] + " " + str(companyCountList[tempIndex]))

    companyList.remove(companyList[tempIndex])
    companyCountList.remove(companyCountList[tempIndex])


























#
