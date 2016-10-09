#! python3
import time
import openpyxl
from selenium import webdriver
from bs4 import BeautifulSoup
import csv
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary

#driver = webdriver.Chrome('C:/Users/Leonid/Downloads/chromedriver_win32/chromedriver.exe')
driver = webdriver.PhantomJS()
#driver = webdriver.PhantomJS('C:/Users/Leonid/Downloads/phantomjs-2.1.1-windows/phantomjs-2.1.1-windows/bin/phantomjs.exe')
driver.set_window_size(1120, 550)
url = 'http://ati.su/Tables/Default.aspx?EntityType=Truck&FromGeoList=255ccce5-4b8b-e611-a612-002590e45781_0&ToGeoList=255ccce5-4b8b-e611-a612-002590e45781_0&FromGeoRadius=-1&ToGeoRadius=-1&qdsv=-1'
#url = 'http://ati.su/Tables/Default.aspx?EntityType=Truck&FromGeo=2_71&ToGeo=2_71&FromGeoRadius=-1&ToGeoRadius=-1&qdsv=-1'
#finals = []


def parse (html, projects):
    soup = BeautifulSoup(html, "html.parser")
    #print(soup)
    #вычленяем таблицу, содержащую необходимую информацию
    table_body = soup.find('table', class_='atiTables')
    table = table_body.tbody

    rec1 = 0  # кол-во рекомендаций
    pret1 = 0 # кол-во претензий
    np1 = 0   # кол-во недобросовестных партнеров
    tel = None # телефон
    weight = 0
    firm = None
    base_cty = None
    car_type_1 = None
    ld_city = None
    ud_city = None
    ball = None

    #projects = []

    for row in table.select('tr[class]'):
        #print(row)
        col = row.find_all('td')
        firm = row.find('span', id = 'ShortenedLbl')
        #print(firm)
        base_cty = row.find('td', style = 'white-space: nowrap; padding-left: 5px;')

        car_types = col[2].find('div', class_='gridCell')
        #print(car_types)
        if car_types:
            car_type_1 = car_types.find('span').text
            #print(car_type_1)

        car_input = col[3].find('div', id = 'divCity')
        if car_input:
            ld_city = (car_input.b.text)

        car_cty_output = col[4].find('div', style="white-space: nowrap")
        car_reg_output = col[4].find('div', style="white-space: nowrap;")
        car_country_output = col[4].find('span', style="white-space: nowrap")
        #print(col[4])
        if car_cty_output:
            #print('1')
            ud_city = (car_cty_output.b.text.strip())
        elif car_reg_output and car_reg_output.b.text.strip():
            #print('2')
            ud_city = (car_reg_output.b.text.strip())
        elif car_country_output:
            #print('3')
            ud_city = ('Россия')

        tel = row.find_all('a', class_='PhoneNumberRef')

        rate = row.find('span', class_='forumtopictitle')
        if rate:
            ball = rate['ratedescription'][20:24]
            ball = float(ball.replace(",", "."))
        else:
            ball = 0

        rec = row.find('span', style="color: #00AA00;")
        #print(rec)
        if rec:
            rec1 = int(rec.text[1:])
        else:
            rec1 = 0

        pret = row.find('span', style="color: #FF0000;")
        if pret:
            pret1 = int(pret.text[1:])
        else:
            pret1 = 0

        np = row.find('span', style="color: #666666;")
        if np:
            np1 = int(np.text[2:])
        else:
            np1 = 0

        weight_tmp = col[2].find('div', class_='gridCell')
        if weight_tmp:
            weight = (float(weight_tmp.div.span.next_sibling[12:15].replace('/','').replace(',', '.').strip()))
            #print(weight)

        #if firm:
        if ball >= 1.0 and rec1 >= 3 and np1 <= 2 and weight >= 1.5:
            #print(weight)
            projects.append({
                'name': firm.text,
                'base': base_cty.text,
                'type': car_type_1,
                'load_city': ld_city,
                'unload_city': ud_city,
                'phone_numbers': [phone_number.text for phone_number in tel],
                'stars': ball,
                'recommends': rec1,
                'claims': pret1,
                'unscrupulous partners': np1,
                'capacity': weight
        })
        #finals += projects
        rec1 = 0
        pret1 = 0
        np1 = 0

    #for project in projects:
    #    print(project)
    bool_next_page = soup.find(title = 'Перейти на следующую страницу')
    if bool_next_page:
        url_next_page ='http://ati.su' + bool_next_page.get('href')
        print('Retrieving from ' + url_next_page)
        driver.get(url_next_page)
        time.sleep(3)
        parse(driver.page_source, projects)
        time.sleep(1)


def save(projects, path):

    with open(path, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(('Название', 'Город базирования', 'Тип кузова', 'Грузоподъемность', 'Пункт загрузки', 'Пункт выгрузки', 'Тел. номер', 'Звезды', 'Рекомендации', 'Жалобы', 'Недобросовестные партнеры'))
        for project in projects:
            writer.writerow((project['name'], project['base'], project['type'], project['capacity'], project['load_city'], project['unload_city'], ', '.join(project['phone_numbers']), project['stars'], project['recommends'], project['claims'], project['unscrupulous partners']))

def main():
    my_login = "WeberGreen"
    my_password = "Onyxia187"
    projects = []

    #driver = webdriver.Chrome()
    print('Accessing main site...')
    driver.get('http://ati.su/login/login.aspx')
    print('Logging in...')
    username = driver.find_element_by_id("ctl00_ctl00_main_PlaceHolderMain_extLogin_ucLoginFormPage_tbLogin")
    password = driver.find_element_by_id("ctl00_ctl00_main_PlaceHolderMain_extLogin_ucLoginFormPage_tbPassword")

    username.send_keys(my_login)
    password.send_keys(my_password)

    driver.find_element_by_name("ctl00$ctl00$main$PlaceHolderMain$extLogin$ucLoginFormPage$btnPageLogin").click()

    driver.get(url)
    print('Retrieving from ' + url)
    html_source = driver.page_source
    parse(html_source, projects)
    print('Saving...')
    save(projects, 'project.csv')
    print('Parsing has completed successfully.')
    input('Press any key to Exit...')


if __name__ == '__main__':
    main()
