# -------------------------------- #
# Autor: Thiago Jacob Lannes
# Versão: 0.8
# Data: 09/07/2022
# -------------------------------- #

from json.encoder import ESCAPE
import time
from datetime import date, timedelta
from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from tkinter import *
from tkcalendar import DateEntry
import tkinter.font as font
import json

root = Tk()
root.title('FastBO')
#root.iconbitmap('C:/')
root.geometry("260x300")
# ------------------------------------------------ FILEIRA ESQUERDA ------------------------------------------------

LabelServer = Label(root, text="Servidor:")
LabelServer.place(x=10, y=5)

SERVER_Entry = Entry(root)
SERVER_Entry.place(x=70, y=5)

LabelUser = Label(root, text="Usuário:")
LabelUser.place(x=10, y=30)

User_Entry = Entry(root)
User_Entry.place(x=70, y=30)

LabelSenha = Label(root, text="Senha:")
LabelSenha.place(x=10, y=55)

def toggle():
    global Pass_Entry, Passcheckbutton
    if Passcheckbutton.var.get():
        Pass_Entry.config(show = "*")
    else:
        Pass_Entry.config(show = "")

Pass_Entry = Entry(root)
Pass_Entry.default_show_val = Pass_Entry.config(show = "*")
Pass_Entry.config(show="*")
Pass_Entry.place(x=70, y=55)

Passcheckbutton = Checkbutton(root, text="Mostrar", onvalue=False, offvalue=True, command=toggle)
Passcheckbutton.var = BooleanVar(value=True)
Passcheckbutton['variable'] = Passcheckbutton.var
Passcheckbutton.place(x=150, y=52)

LabelCal = Label(root, text="Data:")
LabelCal.place(x=10, y=80)

cal_Entry=DateEntry(root, selectmode='day', date_pattern="yyyy/mm/dd")
cal_Entry.place(x=50, y=80)

Script_Dep_e_Sec_Entry = IntVar()
Script_DS = Checkbutton(root, text="Departamento e Secçao", variable=Script_Dep_e_Sec_Entry)
Script_DS.place(x=10, y=105)
Script_Forma_de_Pag_Entry = IntVar()
Script_FP = Checkbutton(root, text="Forma de Pagamento", variable=Script_Forma_de_Pag_Entry)
Script_FP.place(x=10, y=130)
Script_Venda_e_Fecho_TPA_Entry = IntVar()
Script_VFTPA = Checkbutton(root, text="Venda e Fecho TPA - Detalhe", variable=Script_Venda_e_Fecho_TPA_Entry)
Script_VFTPA.place(x=10, y=155)
Script_RCE_Entry = IntVar()
Script_R = Checkbutton(root, text="RCE", variable=Script_RCE_Entry)
Script_R.place(x=10, y=180)


# ------------------------------------------------ FILEIRA DIREITA ------------------------------------------------
"""
LabelFATU = Label(root, text="DOWNLOAD DE FATURA")
LabelFATU.place(x=280, y=5)

LabelServer_FATU = Label(root, text="Servidor:")
LabelServer_FATU.place(x=230, y=30)

SERVER_Entry_FATU = Entry(root, width=15)
SERVER_Entry_FATU.place(x=284, y=30)
"""

# ------------------------------------------------ RUN ------------------------------------------------
def RunPrep():    
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    driver = webdriver.Chrome(PATH)
    SERVER_URL = SERVER_Entry.get()
    driver.get(SERVER_URL)

    delay = 10 # segundos
    hoje = date.today()

    OperatorID = driver.find_element(by=By.ID, value="ctl0_main_txtOperatorID")
    OperatorPass = driver.find_element(by=By.ID, value="ctl0_main_txtAccessPassword")

    User = User_Entry.get()
    OperatorID.send_keys(User)
    OperatorID.send_keys(Keys.TAB)

    time.sleep(1)
    Pass = Pass_Entry.get()
    OperatorPass.send_keys(Pass)
    OperatorLoginBt = driver.find_element(by=By.ID, value="ctl0_main_buttonLogin") 
    OperatorLoginBt.click()

    try:
        PageLoad = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'nav')))

    except TimeoutException:
        print ("Perfomance não está OK")
        
    time.sleep(.5)
    Arquivos = driver.find_element(by=By.ID, value="m14").click() #ARQUIVOS
    time.sleep(.5)
    Touch = driver.find_element(by=By.ID, value="m36").click() #TOUCH
    time.sleep(.5)

    # CRIAÇÃO TECLA

    Touch13 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[1]/div/div/div/ul/li[13]")
    Touch13.click()   
    Touch13Categ = driver.find_element(by=By.ID, value="ctl0_Main_rptFastSalesCategory_ctl12_lnbCategoryID").click()
    time.sleep(.5)
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("PROMOCOES")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("PROMOCOES TIPO 2, 4 E 12 - RETEK")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()
    
    time.sleep(2)

    Touch13 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[1]/div/div/div/ul/li[13]")
    Touch13.click()
    Touch13Itens = driver.find_element(by=By.ID, value="ctl0_Main_rptFastSalesCategory_ctl12_lnbCategoryName").click()

    # 1 ARTIGO TIPO 2 PERCENTAGEM

    time.sleep(2)  

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[13]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl12_ctl0").click()
    
    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("5601217135155")
    time.sleep(.5)
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("DET PO MAQ 100D - 2 PERC.")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("TIPO 2 - PERCENTAGEM")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()
    
    time.sleep(2)   

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN  

    # 2 ARTIGO TIPO 2 PERCENTAGEM

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[14]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl13_ctl0").click()

    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("5601217134165")
    time.sleep(.5)
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("ANTI CALCARIO - 2 PERC.")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("TIPO 2 - PERCENTAGEM")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()
    
    time.sleep(2)  

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN

    # 3 ARTIGO TIPO 2 VALOR

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[15]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl14_ctl0").click()

    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("2100128386418")
    time.sleep(.5)
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("FACA VEGETAIS - 2 VALOR 30.")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("TIPO 2 - VALOR")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()

    time.sleep(2)
    
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN

    # 4 ARTIGO TIPO 4 L3P2 

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[10]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl9_ctl0").click()

    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("5601493450553")
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("SACO - L3P2")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("TIPO 4 - 100 PERC.")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()

    time.sleep(2)  

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN

    # 4 ARTIGO TIPO 12 DISP

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[11]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl10_ctl0").click()

    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("2100128184854")
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("DISP - TIPO 12 - 0,01 DESC TIPO 2")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("DISP - TIPO 12 - 0,01 DESC TIPO 2")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()

    time.sleep(2)    

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # SCROLL DOWN

    # 5 ARTIGO TIPO 12 OFERTA

    Touch13Item1 = driver.find_element(By.XPATH, value="/html/body/form/div[2]/div[2]/div[3]/div[3]/div/div[2]/div/div[2]/div/ul/li[12]")
    Touch13Item1.click()
    Touch13Item1 = driver.find_element(by=By.ID, value="ctl0_Main_rptFastItem_ctl11_ctl0").click()

    time.sleep(.5)

    Touch13ItemEAN = driver.find_element(by=By.ID, value="ctl0_Main_txtItemPOSIdentityID")
    Touch13ItemEAN.clear()
    Touch13ItemEAN.send_keys("5601493489126")
    Touch13Categ_Name = driver.find_element(by=By.ID, value="ctl0_Main_txtName")
    Touch13Categ_Name.clear()
    Touch13Categ_Name.send_keys("BIKE OFERTA TIPO 12")
    Touch13Categ_Name2 = driver.find_element(by=By.ID, value="ctl0_Main_txtDescription")
    Touch13Categ_Name2.clear()
    Touch13Categ_Name2.send_keys("BIKE OFERTA TIPO 12")
    Touch13Categ_NameBt = driver.find_element(by=By.ID, value="ctl0_Main_btnUpdate").click()

    time.sleep(4)

"""
def RunFATU():
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    driver = webdriver.Chrome(PATH)
    driver.get("http://10.126.96.185/sccm_unifo/")

    User = "THIAGOJACOB"
    UserFATU = driver.find_element(by=By.ID, value="edit-name")
    UserFATU.send_keys(User)
    
    time.sleep(1)
    
    Pass = Pass_Entry.get()
    PassFATU = driver.find_element(by=By.ID, value="edit-pass")
    PassFATU.send_keys(Pass)

    BtLoginFATU = driver.find_element(by=By.ID, value="edit-submit").click()

    ServidorFATU = SERVER_Entry_FATU.get()
    
    time.sleep(3)
    
    #LabTest_Nav = driver.find_element(by=By.CLASS_NAME, value="menu-310").click()
    
    #time.sleep(1)

    driver.get("/sccm_unifo/?q=ExtractUniFOFiles")
"""

def RunRet():
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    Options = webdriver.ChromeOptions()
    settings = {
        "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings), 'savefile.default_directory': 'C:\\Users\\thiago.j.lannes\\Desktop\\Relatórios'}
    Options.add_experimental_option('prefs', prefs)
    Options.add_argument('--kiosk-printing')

    driver = webdriver.Chrome(options=Options, executable_path=PATH)
    SERVER_URL = SERVER_Entry.get()
    driver.get(SERVER_URL)

    """
    # ---- SHADOW HOST GET ----

    def getShadowRoot(host):
        shadowRoot = driver.execute_script("return arguments[0].shadowRoot", host)
        return shadowRoot

    """
    original_window = driver.current_window_handle # window handle da página principal
    delay = 10 # segundos
    hoje = date.today() # dia atual

    OperatorID = driver.find_element(by=By.ID, value="ctl0_main_txtOperatorID")
    OperatorPass = driver.find_element(by=By.ID, value="ctl0_main_txtAccessPassword")

    User = User_Entry.get()
    OperatorID.send_keys(User)
    OperatorID.send_keys(Keys.TAB)

    time.sleep(1)
    Pass = Pass_Entry.get()
    OperatorPass.send_keys(Pass)
    OperatorLoginBt = driver.find_element(by=By.ID, value="ctl0_main_buttonLogin") 
    OperatorLoginBt.click()

    try:
        PageLoad = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'nav')))

    except TimeoutException:
        print ("Perfomance não está OK")

    # Data
    DataEscolhida = cal_Entry.get()
    # Aba de Relatórios
    time.sleep(1)
    RelatorioNav = driver.find_element(by=By.ID, value="m66")
    RelatorioNav.click()
    # Obtenção de opções de Relatórios
    Bool_Script_Dep_e_Sec = Script_Dep_e_Sec_Entry.get()
    Bool_Script_Forma_de_Pag = Script_Forma_de_Pag_Entry.get()
    Bool_Script_Venda_e_Fecho_TPA = Script_Venda_e_Fecho_TPA_Entry.get()
    Bool_Script_RCE = Script_RCE_Entry.get()
    # Opções de Relatórios
    time.sleep(0.5)

    if Bool_Script_Dep_e_Sec == 1:
        Dep_e_Sec = driver.find_element(by=By.ID, value="m112")
        Dep_e_Sec.click()
        Data_Dep_e_Sec = driver.find_element(by=By.ID, value="ctl0_Main_tslReportSearch_dateBegin_text_dateBegin")
        if DataEscolhida != hoje:
            Data_Dep_e_Sec.clear()
            for char in DataEscolhida:
                Data_Dep_e_Sec.send_keys(char)
                time.sleep(.1)
        time.sleep(0.5)
        BtDep_e_Sec = driver.find_element(by=By.ID, value="ctl0_Main_tslReportSearch_btnLabelButtonGenerate")
        BtDep_e_Sec.click()
        try:
            PageLoad = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'ctl0_Main_pnlContentReport')))
        except TimeoutException:
            print ("Perfomance não está OK")
        BtPrint = driver.find_element(by=By.ID, value="ctl0_btnToolbarReportExport2PDF")
        BtPrint.click()
        
        WaitTab = WebDriverWait(driver, delay).until(EC.number_of_windows_to_be(2))
        
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                break
        driver.execute_script('window.print();')
        
        # NECESSÁRIO SLEEP
        driver.close(window_handle[1])
        driver.switch_to.window(original_window)
        
    if Bool_Script_Forma_de_Pag == 1:
        RelatorioNav = driver.find_element(by=By.ID, value="m66").click()
        time.sleep(.5)
        Forma_de_Pag = driver.find_element(by=By.ID, value="m113").click()
        time.sleep(2)
        DataBegin_Forma_de_Pag = driver.find_element(by=By.ID, value="ctl0_Main_tslReportSearch_dateBegin_text_dateBegin").send_keys(DataEscolhida)
        time.sleep(.5)
        DataEnd_Forma_de_Pag = driver.find_element(by=By.ID, value="ctl0_Main_tslReportSearch_dateEnd_text_dateEnd").send_keys(DataEscolhida)
        time.sleep(.5)
        BtForma_de_Pag = driver.find_element(by=By.ID, value="ctl0_Main_tslReportSearch_btnLabelButtonGenerate").click()

#    if Bool_Script_Venda_e_Fecho_TPA == 1:


    time.sleep(300)
# ------------------------------------------------ BT ESQUERDA ------------------------------------------------
LeftButton = Button(root, text="COMEÇAR SCRIPT RETAGUARDA", command=RunRet)
LeftButton.place(x=35, y=230)

# ------------------------------------------------ BT DIREITA ------------------------------------------------

RightButton = Button(root, text="CONFIGURAR TECLAS PROMO", command=RunPrep)
FontSize = font.Font(size=8)
RightButton['font'] = FontSize
RightButton.place(x=48, y=265)


root.mainloop()



"""
------------------------- AGENDAMENTO DE FUNDO ------------------------
time.sleep(1)

FundoTab = driver.find_element(by=By.ID, value="m167") 
FundoTab.click()
time.sleep(0.5)
FundoScheduleBt = driver.find_element(by=By.ID, value="m173")
FundoScheduleBt.click()
time.sleep(0.5)
NovoFundoBt = driver.find_element(by=By.ID, value="ctl0_Main_btnInsert")
NovoFundoBt.click()
time.sleep(0.5)
FundoReforcoBt = driver.find_element(by=By.ID, value="ctl0_Main_txsCashFund_txtCode")
CodFundo = "00002"
FundoReforcoBt.send_keys(CodFundo)
FundoReforcoBt.send_keys(Keys.TAB)

wait=WebDriverWait(driver, 10)

try:
    wait.until(EC.element_to_be_clickable((By.ID, "ctl0_Main_txsOperatorSchedule_txtCode"))).send_keys(User)
except:
    print('Timeout')

OperatorFundo = driver.find_element(by=By.ID, value="ctl0_Main_txsOperatorSchedule_txtCode")

OperatorFundo.send_keys(Keys.TAB)

try:
    wait.until(EC.element_to_be_clickable((By.NAME, "checkbox[{}]".format(User)))).click()
except:
    print('Timeout')

try:
    wait.until(EC.element_to_be_clickable((By.ID, "ctl0_Main_btnScheduling"))).click()
except:
    print('Timeout')

time.sleep(8)

driver.quit()
"""