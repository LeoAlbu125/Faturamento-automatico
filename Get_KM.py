from select import select
import time
import os
import pandas as pd
from datetime import datetime as dt
from datetime import timedelta as td
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import openpyxl.drawing.image as imag
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import win32com.client
import win32process

meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

n_mes = [li for li in range(1, 13)]

dict_nm_m = dict(zip(n_mes, meses))


def to_pdf(fname):
    save_pdf = os.path.splitext(fname)[0] + '.pdf'
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(Filename=fname)
    book.ExportAsFixedFormat(0, save_pdf)
    sheet = None
    book = None
    excel.Quit()

    excel = None


def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True


def check_timer_click(xpath, timer, tes="", click=0, send=0):

    for i in range(0, (timer+1)*2):
        if check_exists_by_xpath(xpath=xpath):
            if click == 1:
                driver.find_element_by_xpath(xpath=xpath).click()
            if send == 1:
                driver.find_element_by_xpath(xpath=xpath).send_keys(tes)
            break
        time.sleep(0.5)


dset = pd.read_excel("Lista_Cliente.xlsx")

# Filtro da tabela mestra quando o dia atual for segunda
if dt.today().weekday() == 0:

    td_dset = dset[(dset["Dt Renov"] <= dt.today()) & (
        dset["Dt Renov"] >= (dt.today() - td(days=3)))]

# Filtro da tabela mestra para a data atual
else:
    td_dset = dset[dset["Dt Renov"] == dt.today().strftime(r"%Y-%m-%d")]

link = "link"
print(dt.today())
print(dset.head(1)["Dt Renov"])
# Looping com a tebela filtrada de clientes com data renovação = hoje
for x in range(0, len(td_dset)):

    # essa variavel recria o nome da pasta do cliente
    # (6 primeiros digitos CPF) + (primeiro nome)
    var = (r"\{}_{}".format(str(td_dset.iloc[x].loc["CPF"])[
        :6], str(td_dset.iloc[x].loc["Name"]).split()[0]))

    # código para pegar o nome da pasta (CPF + NOME)
    # monta o path da pasta do cliente
    direc = (os.path.dirname(os.path.abspath(__file__))+var)

    chromeOptions = webdriver.ChromeOptions()
    # muda o diretório padrão do dowload para a pasta do cliente
    # assim quando baixar o arquivo ele ja estara guardado na pasta do cliente
    prefs = {"download.default_directory": direc}
    chromeOptions.add_experimental_option("prefs", prefs)
    chromeOptions.add_argument("--headless")

    driver = webdriver.Chrome(
        ChromeDriverManager().install(), options=chromeOptions)

    driver.get(link)

    # Garante que a pagina já carregou
    for y in range(0, 11):
        try:
            assert "Acessar | Site de administração do Django" in driver.title
            break
        except:
            time.sleep(1)
            pass

    # Escreve o email no campo login
    check_timer_click('//*[@id="id_username"]', 10,
                      "", 0, 1)
    # Escreve a senha
    check_timer_click('//*[@id="id_password"]', 10, "", 0, 1)
    # envia
    check_timer_click('//*[@id="login-form"]/div[3]/input', 10, click=1)

    # Monta um dicionario com o nome cliente e o path HTML da check box dele
    listaxpath = driver.find_elements_by_xpath('//table/tbody/tr/td/input')
    listanames = [kron.text for kron in driver.find_elements_by_xpath(
        '//table/tbody/tr/th/a')]

    dirNaPa = dict(zip(listanames, listaxpath))

    dirNaPa[td_dset.iloc[x].loc["Name"]].click()

    check_timer_click('//*[@id="btn-export"]', 10, click=1)

    for c in range(0, 11):
        try:
            assert "Exportar | Site de administração do Django" in driver.title
            break
        except:
            time.sleep(1)
            pass

    select = Select(driver.find_element_by_xpath('//*[@id="id_file_format"]'))

    select.select_by_value("2")

    #====================Selecionar a data max na caixinha do datapicker=================================#
    check_timer_click(
        '//*[@id="end-date"]', 10, Keys.ENTER, 1, 1)

    dtmax = dt.date(td_dset.iloc[x].loc["Dt Renov"]) - td(days=1)
    dtmin = dt.date(td_dset.iloc[x].loc["Dt Renov"]) - td(days=31)

    # Variavel com texto "month year"
    mon_yea = str(dict_nm_m[dtmax.month])+" "+str(dtmax.year)

    k = 0
    # compara o mes ano do colandario com a variavel mon_yea
    # caso seja diferente muda o calendario para o mes anterior até a var mon_yea bater com mes ano
    for k in range(0, 51):
        time.sleep(0.5)
        if (driver.find_element_by_xpath('/html/body/div[2]/div[3]/ul[1]/li[2]').text != mon_yea):
            check_timer_click(
                '/html/body/div[2]/div[3]/ul[1]/li[1]', 10, click=1)
        else:
            break
    # dicionario com nome da var data-view no html e o path
    _lista = driver.find_elements_by_xpath('/html/body/div[2]/div[3]/ul[3]/li')
    _lista_hum = [gran.get_attribute("data-view")+gran.text for gran in _lista]
    dict_data = dict(zip(_lista_hum, _lista))
    new_dict_data = {o: p for o, p in dict_data.items() if not "day " in o}
    new_dict_data["day"+str(dtmax.day)].click()

    #====================================================================================================#

    #====================Selecionar a data min na caixinha do datapicker=================================#
    check_timer_click(
        '//*[@id="start-date"]', 10, Keys.ENTER, 1, 1)

    mon_yea_2 = str(dict_nm_m[dtmin.month])+" "+str(dtmin.year)

    k = 0

    for k in range(0, 51):
        time.sleep(0.5)
        if (driver.find_element_by_xpath('/html/body/div[3]/div[3]/ul[1]/li[2]').text != mon_yea_2):
            check_timer_click(
                '/html/body/div[3]/div[3]/ul[1]/li[1]', 10, click=1)
        else:
            break

    _lista = driver.find_elements_by_xpath('/html/body/div[3]/div[3]/ul[3]/li')
    _lista_hum = [gran.get_attribute("data-view")+gran.text for gran in _lista]
    dict_data = dict(zip(_lista_hum, _lista))
    new_dict_data = {o: p for o, p in dict_data.items() if not "day " in o}
    new_dict_data["day"+str(dtmin.day)].click()

#====================================================================================================#

    check_timer_click('//*[@id="form_data"]/div/input', 10, click=1)
    t_start = time.time()

    for zys in range(0, 101):
        if os.path.isfile(direc+"\Customer-{}.xlsx".format(dt.today().strftime(r"%Y-%m-%d"))):
            break
        else:
            time.sleep(0.1)

    kmlo = direc+r"\Customer-{}.xlsx".format(dt.today().strftime(r"%Y-%m-%d"))
    td_km_dt = pd.read_excel(kmlo)

    dset["KM"][(dset["CPF"] == td_dset.iloc[x].loc["CPF"]) & (dset["Dt Renov"] == dt.today(
    ).strftime(r"%Y-%m-%d"))] = td_km_dt["meter_run_day"].sum()/1000

    dset.to_excel("Lista_Cliente.xlsx", index=False)

    driver.close()

    time.sleep(1)

#=====================excel_writer=================================================================#

    KM_value = dset["KM"][(dset["CPF"] == td_dset.iloc[x].loc["CPF"]) & (dset["Dt Renov"] == dt.today(
    ).strftime(r"%Y-%m-%d"))]
    print(KM_value)
    wb = load_workbook(filename=("Templete.xlsx"))

    dt_cobr = dt.date(td_dset.iloc[x].loc["Dt Renov"]) + td(days=2)
    dt_cobr = dt_cobr.strftime(r"%m/%d/%Y")

    sheet_ranges = wb['Output']

    sheet_ranges["G2"] = td_dset.iloc[x].loc["Name"]
    sheet_ranges["D9"] = td_dset.iloc[x].loc["PRICE"]/24
    sheet_ranges["B11"] = str(int(KM_value))+" KM"
    sheet_ranges["C11"] = td_dset.iloc[x].loc["PKM"]

    if float(KM_value) <= 1200:
        print("OK")
        sheet_ranges["D11"] = float(KM_value)*float(td_dset.iloc[x].loc["PKM"])
        sheet_ranges["D14"] = float(
            td_dset.iloc[x].loc["PRICE"]/24) + float(KM_value)*float(td_dset.iloc[x].loc["PKM"])

    else:
        print("esquisito")
        sheet_ranges["D11"] = 1200*float(td_dset.iloc[x].loc["PKM"])
        sheet_ranges["D14"] = float(
            td_dset.iloc[x].loc["PRICE"]/24) + 1200*float(td_dset.iloc[x].loc["PKM"])

    sheet_ranges["G10"] = "Nome: " + td_dset.iloc[x].loc["Name"]
    sheet_ranges["G13"] = "Modelo: " + td_dset.iloc[x].loc["Veicle model"]
    sheet_ranges["G14"] = "Marca: " + td_dset.iloc[x].loc["Veicle mark"]
    sheet_ranges["G15"] = "Placa: " + td_dset.iloc[x].loc["Plate"]
    sheet_ranges["G20"] = "Lançamento: " + str(dt_cobr)
    sheet_ranges["C20"] = td_dset.iloc[x].loc["W2P"]
    sheet_ranges["H30"] = str(int(KM_value)) + " km"
    print(str(int(KM_value)))
    sheet_ranges["H32"] = str(int(int(KM_value)/30)) + " km/dia"

    x_ax = td_km_dt["date"]
    y_ax = td_km_dt["meter_run_day"]/1000

    plt.figure(num=1, figsize=(10, 5), dpi=100)
    plt.bar(x_ax, y_ax, width=0.8, color="GOLD")

    date_form = DateFormatter(r"%d")
    plt.gca().xaxis.set_major_formatter(date_form)
    plt.figure(num=1, figsize=(10, 5), dpi=1080)
    plt.xticks(x_ax)
    plt.grid(axis="y")

    plt.savefig(direc+r"\temp.jpeg", bbox_inches='tight')

    img = imag.Image(direc+r"\temp.jpeg")
    img.width = 420
    img.height = 300
    sheet_ranges.add_image(img, "B30")

    data_max_km = td_km_dt.iloc[td_km_dt["meter_run_day"].idxmax()].loc["date"]

    sheet_ranges["H34"] = dt.date(data_max_km).strftime(r"%d/%m/%Y")
    sheet_ranges["H36"] = td_dset.iloc[x].loc["Dt start"]

    n_dtmin = dtmin.strftime(r"%d/%m/%Y")
    n_dtmax = dtmax.strftime(r"%d/%m/%Y")

    sheet_ranges["G5"] = "De {} até {}".format(n_dtmin, n_dtmax)
    print("aqui")
    wb.save(direc+r"\fatura_{}_{}.xlsx".format(
        td_dset.iloc[x].loc["Name"].split()[0], dt.today().strftime(r"%d-%m-%Y")))

    to_pdf(direc+r"\fatura_{}_{}.xlsx".format(
        td_dset.iloc[x].loc["Name"].split()[0], dt.today().strftime(r"%d-%m-%Y")))

    wb.close()
