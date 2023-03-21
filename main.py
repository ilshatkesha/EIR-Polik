import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options


dr = webdriver.Firefox() #эта хуйня управляет файрфоксом

dr.get('http://10.2.213.165:90/User/Login') #заходим

dr.find_element("name", "Login").send_keys('ilshatkesha')
dr.find_element("name", "Password").send_keys('0020275')
dr.find_element("xpath", "//button[@type='submit']").click()
dr.find_element("xpath", "//a[contains(text(),'Подсистемы')]").click()
dr.find_element("xpath", "//a[contains(text(),'Направления на госпитализацию')]").click()
dr.find_element("xpath", "//a[contains(text(),'Направления на госпитализацию (Поликлиники)')]").click() #тут крч входит там все дела
dr.find_element("xpath", "//button[@id='polyclinic-find']").click()

time.sleep(25)

columns_list = ['ФИО пациента', 'Страховой полис', 'Электронное направление', 'Вид госпитализации', '№ направления', 'Дата направления', 'Отделение', 'Основной диагноз'] #а тут инициализирует колонки для считывания

df = pd.read_excel(r"d:\\0103-1003.xlsx", sheet_name=0, dtype=str)[columns_list] #считывает эксельку
df['Дата направления'] = pd.to_datetime(df['Дата направления']) #меняет формат даты для ввода
k = 0
#kost = df[['Основной диагноз']].iloc[0]

for i in range(0,df.shape[0]): #основной цикл
        otd = df['Отделение'][i]
        hosp_vid = df['Вид госпитализации'][i]
        vid_napr = df['Электронное направление'][i]

        if (vid_napr.lower() == 'нет') and (hosp_vid.lower() == 'экстренно'): #проверка направления и вида госпитализации
                print('экстренно без направления', df['ФИО пациента'][i])
                #k = k+1
                continue

        #if otd.isna().values:
                #print(i + 1, " Нет отделения  ", df['ФИО пациента'][i])
                #continue

        time.sleep(2)
        if '22198. Пульмонологическое отделение' in otd:  #сопоставление отделения
                dr.find_element("xpath", "//td[contains(.,'Пульмонологическое (РДКБ) стационар')]").click()
        elif '22201. Кардиологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Кардиологическое (РДКБ) стационар')]").click()
        elif '22200. Аллергологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Аллергологическое (РДКБ) стационар')]").click()
        elif '22202. Отделение челюстно-лицевой хирургии' in otd:
                dr.find_element("xpath", "//td[contains(.,'Челюстно-лицевая хирургия (РДКБ) стационар')]").click()
        elif '22203. Гематологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Гематологическое (РДКБ) стационар')]").click()
        elif '22204. Эндокринологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Эндокринологическое (РДКБ) стационар')]").click()
        elif '22205. Хирургическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (общая хирургия) (РДКБ) стационар')]").click()
        elif '22206. Урологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Урологическое (РДКБ) стационар')]").click()
        elif '22208. Отоларингологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Отоларингологическое (РДКБ) стационар')]").click()
        elif '22210. Гастроэнтерологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Гастроэнтерологическое (РДКБ) стационар')]").click()
        elif '22211. Нефрологическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Нефрологическое (РДКБ) стационар')]").click()
        elif '22213. Нейрохирургическое отделение' in otd:
                dr.find_element("xpath", "//td[contains(.,'Нейрохирургическое (РДКБ) стационар')]").click()
        elif '22231. Онкологическое' in otd:
                dr.find_element("xpath", "//td[contains(.,'Онкологическое (РДКБ) стационар')]").click()
        elif '22232. Ревматологическое' in otd:
                dr.find_element("xpath", "//td[contains(.,'Ревматологическое (РДКБ) стационар')]").click()
        elif '22233. Гинекологическое (хирургия)' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (гинекология) (РДКБ) стационар')]").click()
        elif '22235. Ортопедическое' in otd:
                dr.find_element("xpath", "//td[contains(.,'Ортопедическое (РДКБ) стационар')]").click()
        elif '22330. Хирургическое отделение № 2' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургия новорожденных (РДКБ) стационар')]").click()
        elif '22331. Отделение для недоношенных и патологии новорожденных' in otd:
                dr.find_element("xpath", "//td[contains(.,'Патология недоношенных (РДКБ) стационар')]").click()
        elif '22341. Психоневрологическое отделение №1' in otd:
                dr.find_element("xpath", "//td[contains(.,'Неврологическое №1 (РДКБ) стационар')]").click()
        elif '22342. Психоневрологическое отделение №2' in otd:
                dr.find_element("xpath", "//td[contains(.,'Неврологическое №2 (РДКБ) стационар')]").click()
        elif '22343. Психоневрологическое отделение №3' in otd:
                dr.find_element("xpath", "//td[contains(.,'Неврологическое №3 (РДКБ) стационар')]").click()
        elif '22352. Хирургия детей раннего возраста' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургия детей раннего возраста (РДКБ) стационар')]").click()
        elif ('22438. Дневной стационар гинекологический' in otd) :
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22438. Дневной стационар ЛОР' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22438. Дневной стационар нейрохирургический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22438. Дневной стационар ортопедический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22438. Дневной стационар хирургический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Хирургическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар ревматологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар нефрологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар аллергологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар урологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар эндокринологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22439. Дневной стационар гастроэнтерологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Педиатрическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22440. Дневной стационар онкологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Онкологическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22440. Дневной стационар гематологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Онкологическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22441. Детская реабилитация (дн.ст.при ДЦПНиЭ)' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская реабилитация в МО (ДЦПНиЭ) (РДКБ) дн.стац. при стационаре')]").click()
        elif '22446. Дневной стационар неврологический' in otd:
                dr.find_element("xpath", "//td[contains(.,'Неврологическое (РДКБ) дн.стац. при стационаре')]").click()
        elif '22609. Медреабилитация ПНО 1' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская реабилитация в МО (неврология №1) (РДКБ) стационар')]").click()
        elif '22610. Медреабилитация ПНО 2' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская реабилитация в МО (неврология №2) (РДКБ) стационар')]").click()
        elif '22613. Неврологическое отделение (дневной стационар при КДП)' in otd:
                dr.find_element("xpath", "//td[contains(.,'Неврологическое (РДКБ) дн.стац. при поликлинике')]").click()
        elif '22614. Детская реабилитация дн.ст. при КДП' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская реабилитация в МО (РДКБ) дн.стац. при поликлинике')]").click()
        elif '22648. Детская паллиативная помощь РДКБ (стационар)' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская паллиативная помощь (РДКБ) стационар')]").click()
        elif '22606. Медреабилитация ОПН' in otd:
                dr.find_element("xpath", "//td[contains(.,'Детская реабилитация в МО (патология новорожденных) (РДКБ) стационар')]").click()
        elif '22650. Отделение клинической иммунологии' in otd:
                continue
        #elif '22651. Педиатрическое отделение (нефро)' in otd:
                #dr.find_element("xpath", "//td[contains(.,'Кардиологическое (РДКБ) стационар')]").click()
        else:
                print('нет отделения', otd)
                continue  #сопоставление отделения

        dr.find_element("xpath", "//td[@id='polyclinic-pager_left']/table/tbody/tr/td/div").click() #направить пациента
        time.sleep(10)
        try:
                polis = dr.find_element("name", "polyclinic-pacient-refferal-add-polis")
        except Exception:
                time.sleep(3)
                polis = dr.find_element("name", "polyclinic-pacient-refferal-add-polis") #добавить полис
        p = df['Страховой полис'][i]
        polis.send_keys(p)
        dr.find_element("xpath", "//button[@id='polyclinic-pacient-refferal-add-find']").click() #проверить полис
        time.sleep(4)


        if hosp_vid.lower() == 'планово': #выбор вида госпитализации
                #dr.find_element("xpath", "//select[@id='polyclinic-pacient-refferal-add-help-form']/option[2]").click()
                pass

        else:
                try:
                        dr.find_element("xpath", "// select[ @ id = 'polyclinic-pacient-refferal-add-help-form'] / option").click()
                except Exception:  # пациент не найден по полису
                        dr.find_element("xpath", "(// button[@ type='button'])[9]").click()
                        dr.find_element("xpath", "(//button[@type='button'])[7]").click()
                        print(i + 1, " Полис не действителен  ", df['ФИО пациента'][i])
                        continue
        try: #добавление направления
                naprav = dr.find_element("name", "polyclinic-pacient-refferal-add-direction-number")
                n = df['№ направления'][i]
                naprav.send_keys(n)
        except Exception: #я хз что это но это обход какой-то ошибки
                dr.find_element("xpath", "(// button[@ type='button'])[9]").click()
                dr.find_element("xpath", "(//button[@type='button'])[7]").click()
                # continue

        date = f"{df['Дата направления'][i]}".split()[0] #получение даты в виде даты а не текста
        time.sleep(2)
        dr.find_element("xpath", "//input[@id='polyclinic-pacient-refferal-add-direction-date']").send_keys(date)
        dr.find_element("xpath", "//input[@id='polyclinic-pacient-refferal-add-plan-date']").send_keys(date) #ввод даты направления
        time.sleep(8)
        try:
                dr.find_element("xpath", "//button[@id='polyclinic-pacient-refferal-add-mkb']/img").click() #добавить мкб
        except Exception: #пациент не найден по полису
                dr.find_element("xpath", "(// button[@ type='button'])[9]").click()
                dr.find_element("xpath", "(//button[@type='button'])[7]").click()
                print(i + 1, " Полис не действителен  ", df['ФИО пациента'][i])
                continue
        if i == 0:
                time.sleep(35)
        else:
                time.sleep(15)
        dr.find_element("xpath", "//input[@id='gs_Id']").click()
        mkb = dr.find_element("xpath", "//input[@id='gs_Id']")
        m = df[['Основной диагноз']].iloc[i]

        if m.isna().values: #проверка на наличие мкб
                print(i + 1, " Нет МКБ  ", df['ФИО пациента'][i])
                dr.find_element("xpath", "(//button[@type='button'])[10]").click()
                dr.find_element("xpath", "(//button[@type='button'])[7]").click()
                continue
        else:
                m = m.values
                m = df['Основной диагноз'][i][:5] #ввод мкб
                mkb.send_keys(m)
                time.sleep(2)
                #print(kost, m)
                try:
                        dr.find_element("css selector", "td:nth-child(3) > input").click()
                        time.sleep(2)
                        #if (kost == m):
                                #dr.find_element("css selector", "td:nth-child(3) > input").click()
                                #time.sleep(2)
                except Exception:
                        mkb.clear()
                        m = df['Основной диагноз'][i][:3] #ввод мкб с 3 символами если нет 5
                        mkb.send_keys(m)
                        time.sleep(3)
                        dr.find_element("css selector", "td:nth-child(3) > input").click()
                        time.sleep(2)
                        #if (kost == m):
                                #dr.find_element("css selector", "td:nth-child(3) > input").click()
                #n = df[['Основной диагноз']].iloc[i]
                #n = n.values
                #n = df['Основной диагноз'][i][:5]
                #kost = m
        time.sleep(3)
        dr.find_element("xpath", "(// button[@ type='button'])[9]").click()
        time.sleep(8)
        # dr.find_element("xpath", "//input[@id='polyclinic-pacient-refferal-add-direction-date']").click()
        dr.find_element("xpath", "(//button[@type='button'])[6]").click()
        time.sleep(12)
        try:
                dr.find_element("xpath", "(//button[@type='button'])[4]").click()
        except Exception:
                dr.find_element("xpath", "(//button[@type='button'])[10]").click()
                dr.find_element("xpath", "(//button[@type='button'])[7]").click()
                print(i + 1, " Пациент уже был направлен ", df['ФИО пациента'][i])
                continue
        time.sleep(5)
        print(i + 1, " ", df['ФИО пациента'][i])
        print(n, m)