from selenium import webdriver
from bs4 import BeautifulSoup
import xlsxwriter, datetime, argparse

# parser = argparse.ArgumentParser(description='Path to files')
# parser.add_argument('--outdir', type=str, help='Output dir for files')
# args = parser.parse_args()

url = "https://old.bankrot.fedresurs.ru/Messages.aspx"
driver = webdriver.Chrome()
driver.get(url)
innerHTML = driver.execute_script("return document.body.innerHTML")
soup = BeautifulSoup(innerHTML, "html.parser")
alllinks = soup.find_all('a', href=True)

def getinfo(url):

    finaldata = []

    driver = webdriver.Chrome()
    driver.get(url)
    innerHTML = driver.execute_script("return document.body.innerHTML")
    soup = BeautifulSoup(innerHTML, "html.parser")

    tables = soup.find_all('table', class_='personInfo')
    tableperson = tables[0]
    tabledata = tables[1]

    for child in tableperson.children:
        for child2 in child.children:
            stringdata = []
            try:
                for mystring in child2.children:
                    stringdata.append(mystring.text)
                finaldata.append(stringdata)
            except:
                print('Не та строка')

    finaldata.append('')

    for child in tabledata.children:
        for child2 in child.children:
            stringdata = []
            try:
                for mystring in child2.children:
                    stringdata.append(mystring.text)
                finaldata.append(stringdata)
            except:
                print('Не та строка')

    return finaldata

def saveinfo(info):
    date = datetime.datetime.now().strftime("%Y-%m-%d %H-%M")
    # savepath = args.outdir
    # if savepath == None:
    savepath = 'C:\Temp\\'

    with xlsxwriter.Workbook(f'{savepath}{date}.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        for row_num, data in enumerate(info):
            worksheet.write_row(row_num, 0, data)
        print('Данные успешно сохранены!')

# Тестовый пример
# url = 'https://old.bankrot.fedresurs.ru/MessageWindow.aspx?ID=EE784CF94C5E4A0EB34815F017D778D4'
# info = getinfo(url)
# saveinfo(info)

for link in alllinks:
    if link.text.strip() == 'Отчет оценщика об оценке имущества должника':
        title = link['title']
        weblink = link['href']
        pageirl = 'https://old.bankrot.fedresurs.ru' + weblink
        info = getinfo(pageirl)
        saveinfo(info)






