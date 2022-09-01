from pywinauto.application import Application
from pywinauto import mouse
from pywinauto.keyboard import send_keys
from time import sleep
from string import ascii_uppercase
import os
from os import path
from pywinauto import Desktop
import re
# import psutil

panel = Application(backend='uia')
panel2 = Desktop(backend='uia')


def taskbar_control(menu_items, unpin=False):
    button_text = "Przypnij do paska zadań" if unpin is False else "Odepnij od paska zadań"
    for application in menu_items:
        send_keys('{RWIN}'+application)
        send_keys('{VK_APPS}')
        panel.connect(title='Wyszukiwarka')
        wyszukiwarka = panel.window(title='Wyszukiwarka')
        wyszukiwarka.wait('enabled')
        menu = wyszukiwarka.child_window(title="Menu kontekstowe", control_type="List").items()
        for item in menu:
            send_keys('{DOWN}')
            if item.texts()[0] == button_text:
                send_keys('{ENTER}')
                break


def default_apps(setting_title, app_title, ustawienia):
    ustawienia.child_window(title_re=setting_title,  control_type="Button").click()
    send_keys('{ENTER}')
    ustawienia.child_window(title=app_title, control_type="Button").click()
    send_keys('{ENTER}')
    try:
        ustawienia.child_window(title="Zanim przełączysz").wait('enabled')
        send_keys('{TAB} {ENTER}')
    except:
        print('Nie znaleziono okna :)')


def set_default_apps(settings):
    send_keys('{LWIN} aplikacje {SPACE} domy {ENTER}')
    panel.connect(title='Ustawienia', timeout=100)
    ustawienia = panel.window(title='Ustawienia')
    ustawienia.wait('enabled')
    for setting in settings:
        default_apps(setting['setting_title'], setting['app_title'], ustawienia)
    ustawienia.child_window(title="Zamknij aplikację Ustawienia", control_type="Button").click()


def firefox_config():
    send_keys('{RWIN} firefox {ENTER}')
    try:
        panel.connect(title_re='Witamy w przegl')
        przegladarka=panel.window(title_re='Witamy w przegl')
        przegladarka.wait('enabled')
    except:
        panel.connect(title='Mozilla Firefox')
        przegladarka2=panel.window(title='Mozilla Firefox')
        przegladarka2.wait('enabled')
    send_keys('^t')
    send_keys('about+:preferences {ENTER}')
    firefox = panel.window(title_re='Ustawienia')
    firefox.wait('enabled')
    send_keys('aplikacje')
    zrodlo = firefox.window(title_re="Dokument PDF", control_type="ListItem").wrapper_object().texts()
    if zrodlo[0].find('Firefox') != -1:
        firefox.window(title="Dokument PDF Otwórz w programie Firefox", control_type="ListItem").select()
        send_keys('{TAB} {DOWN 3}')
    try:
        firefox.child_window(title="Zamknij", control_type="Button").click()
    except:
        firefox.child_window(title="Zamknij", control_type="Button", found_index=0).click()
        


def reader_config():
    send_keys('{RWIN} acrobat {ENTER}')
    panel.connect(path='C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe')
    # panel.start("C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe")
    # panel.connect(path="C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe")
    sleep(2)
    reader = panel.window(title_re='Adobe Acrobat Reader ')
    sleep(1)
    umowa = reader.child_window(title_re='Adobe Acrobat Reader DC - ')
    umowa.wait('enabled')
    umowa.child_window(title="Akceptuj", control_type="Button").click()
    panel.kill()
    send_keys('{LWIN down} r {LWIN up} +'+(os.path.splitdrive(__file__)[0][0])+'+;\\nowy+_user {ENTER}')
    panel.connect(path='C:\\Windows\\explorer.exe')
    plik_testowy = panel.window(title_re='nowy_')
    plik_testowy.child_window(title="tester", control_type="ListItem").select()
    send_keys('{VK_APPS} {UP} {ENTER}')
    sleep(1)
    panel.connect(title_re = 'Właściwości:')
    wlasciwosci = panel.window(title_re = 'Właściwości:')
    wlasciwosci.wait('enabled')
    wlasciwosci.child_window(title="Zmień...", control_type="Button").click()
    #windows = Desktop(backend="uia").windows()
    #print([w for w in windows])
    okno_latane = panel2.window(best_match='Jak chcesz od teraz otwierać pliki')
    okno_latane.wait('enabled')
    okno_latane.child_window(title_re="Adobe Acrobat", control_type="ListItem").select()
    send_keys('{ENTER}')
    sleep(1)
    wlasciwosci.wait('enabled')
    wlasciwosci.child_window(title="Zastosuj", control_type="Button").click()
    wlasciwosci.child_window(title="OK", control_type="Button").click()
    panel.kill()
    sleep(1)


def outlook_config():
    send_keys('{RWIN} outlook {ENTER}')
    panel.connect(path='C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE')
    otwieranie = panel.window(title_re='Otwieranie')
    otwieranie.wait('enabled')
    panel.connect(title='Outlook — Zapraszamy!')
    office = panel.window(title='Outlook — Zapraszamy!')
    office.wait('visible')
    office_tekst =office.child_window(title_re="Zainstalowane są następujące aplikacje", control_type="Text").texts()
    print(office_tekst[0][-4:])
    wersja = office_tekst[0][-4:]
    office.child_window(title_re="Zaakceptuj i uruchom", control_type="Button").click()
    sleep(1)
    panel.connect(title_re='Outlook')
    outlook = panel.window(title_re='Outlook')
    outlook.wait('enabled')
    send_keys('^c')
    outlook.child_window(title="Połącz", control_type="Button").click()
    outlook.child_window(title="Ukończono konfigurację konta", control_type="Text").wait('enabled')
    outlook.child_window(title_re="Skonfiguruj też aplikację ", control_type="CheckBox").click()
    outlook.child_window(title="OK", control_type="Button").click()
    sleep(9)
    #----- TO DO TRY EXCEPTA------#
    #for proc in psutil.process_iter():
        #if proc.name() == 'Microsoft Edge':
            #proc.kill()

    panel.connect(title_re='Inbox ')
    outlook_wlasciwy = panel.window(title_re='Inbox')
    outlook_wlasciwy.wait('enabled')
    outlook_wlasciwy.child_window(title="Zamknij", control_type="Button", found_index=0).click()
    sleep(1)


def cloud_config():
    ##windows = Desktop(backend="uia").windows()
    ##print([w for w in windows])
    panel.connect(class_name='Shell_TrayWnd')
    pasek = panel.window(class_name='Shell_TrayWnd')
    badziew = pasek.child_window(title="Przycisk powiadomień", auto_id="1502", control_type="Button")
    badziew.click()
    panel.connect(class_name='NotifyIconOverflowWindow', timeout=100)
    rozwijane = panel.window(class_name='NotifyIconOverflowWindow')
    rozwijane.wait('enabled')
    rozwijane.child_window(title_re="OneDrive", control_type="Button").click()
    sleep(1)
    panel.connect(title='Microsoft OneDrive')
    drive = panel.window(title='Microsoft OneDrive')
    drive.wait('enabled')
    drive.child_window(title="Zaloguj się", control_type="Button").click()
    panel.connect(title='Microsoft OneDrive')
    onedrive = panel.window(title='Microsoft OneDrive')
    onedrive.wait('enabled')
    send_keys('^v')
    onedrive.child_window(title="Zaloguj się", control_type="Button").click()
    sleep(3)
    panel.connect(title='Microsoft OneDrive')
    konf = panel.window(title='Microsoft OneDrive')
    konf.wait('enabled') 
    konf.child_window(title="Twój folder usługi OneDrive", control_type="Text").wait('visible')
    konf.child_window(title="Twój folder usługi OneDrive", control_type="Text").wait('enabled')
    konf.child_window(title="Dalej", control_type="Button").click()
    konf.child_window(title="Utwórz kopię zapasową swoich folderów", control_type="Text").wait('visible')
    konf.child_window(title="Utwórz kopię zapasową swoich folderów", control_type="Text").wait('enabled')
    konf.child_window(title="Kontynuuj", control_type="Button").click()
    sleep(2)
    konf.child_window(title="Poznaj usługę OneDrive", control_type="Text").wait('visible')
    konf.child_window(title="Poznaj usługę OneDrive", control_type="Text").wait('enabled')
    konf.child_window(title="Dalej", control_type="Button").click()
    konf.child_window(title="Udostępnianie plików i folderów", control_type="Text").wait('visible')
    konf.child_window(title="Udostępnianie plików i folderów", control_type="Text").wait('enabled')
    konf.child_window(title="Dalej", control_type="Button").click()
    konf.child_window(title="Wszystkie pliki są gotowe do użycia i dostępne na żądanie", control_type="Text").wait('visible')
    konf.child_window(title="Wszystkie pliki są gotowe do użycia i dostępne na żądanie", control_type="Text").wait('enabled')
    konf.child_window(title="Dalej", control_type="Button").click()
    konf.child_window(title="Pobieranie aplikacji mobilnej ", control_type="Text").wait('visible')
    konf.child_window(title="Pobieranie aplikacji mobilnej ", control_type="Text").wait('enabled')
    konf.child_window(title="Później", control_type="Button").click()
    konf.child_window(title="Usługa OneDrive jest gotowa", control_type="Text").wait('visible')
    konf.child_window(title="Usługa OneDrive jest gotowa", control_type="Text").wait('enabled')
    konf.child_window(title="Otwórz mój folder usługi OneDrive", control_type="Button").click()
    sleep(1)
    panel.connect(title='OneDrive - Gdańskie Centrum Informatyczne')
    folderek = panel.window(title='OneDrive - Gdańskie Centrum Informatyczne')
    folderek.wait('enabled')
    folderek.child_window(title="Zamknij", control_type="Button").click()


#set_default_apps([
    #{'setting_title': 'Poczta e-mail,', 'app_title': 'Outlook'},
    #{'setting_title': 'Przeglądarka sieci Web,', 'app_title': 'Firefox'}
#])

#outlook_config()
#firefox_config()
reader_config()
cloud_config()

taskbar_control(['outlook', 'firefox', ])
taskbar_control(['poczta', 'store', 'vantage'], True)


