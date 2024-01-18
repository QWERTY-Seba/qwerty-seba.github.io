# -*- coding: utf-8 -*-
"""
Created on Mon Jan 23 14:28:38 2023

@author: Seba
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import date
import win32gui
import win32con
import win32api
import ctypes
from contextlib import suppress


# agregar validacion de version del driver con chrome
ruta_driver = r"C:\Users\Seba\Documents\chromedriver.exe"
url = "https://onedrive.live.com/..."
email = "..."
contrasena = "..."
ruta_almacenar_csv = r"..."
nombre_archivo_csv = "valores_carga_datos_{fecha}.csv"
formato_fecha = "%d%m%Y"
nombre_addin = "Office Add-in Cargador de Datos"
maximo_timeout_element = 15

today = date.today().strftime(formato_fecha)
nombre_archivo_csv = nombre_archivo_csv.format(fecha = today )
new_text = r"..."
chrome_options =  Options() 
chrome_options.binary_location = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

chrome_options.add_argument(r"--user-data-dir=C:\SELENIUM_TEST_PROFILES")
chrome_options.add_argument("--profile-directory=Profile 4")

#Posible error aca, el navegador ya esta abierto
driver = webdriver.Chrome(options = chrome_options                          )#executable_path=ruta_driver) 
driver.get("https://www.google.com")
#driver.set_network_conditions(offline=False,latency=100, throughput=500000)


def automatizacion_login():
    #Controlar posibles errores aca, si es que el archivo no existe, no se tiene acceso, etc
    #Controlar tambien si no se redirecciona al archvivo
    driver.get(url)
    if(driver.current_url.startswith("https://login.live.com/")):
        WebDriverWait(driver,maximo_timeout_element).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(email)
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']"))).click()
                
        WebDriverWait(driver,maximo_timeout_element).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='password']"))).send_keys(contrasena)
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']"))).click()
        
                
    if(driver.current_url.startswith("https://login.live.com/ppsecure")):
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.ID, "KmsiCheckboxField"))).click()
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        
    
def automatizacion_carga_extension():
    WebDriverWait(driver, timeout = maximo_timeout_element).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe#WacFrame_Excel_0")))
    
        
    try:
        WebDriverWait(driver, timeout = maximo_timeout_element).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, f"iframe[title='{nombre_addin}']")))
    except TimeoutException:
        print("El maximo de timeout se ha exedido, el addin no se encuentra cargado")
        #Pasar a funcion esta parte de la excepcion
        #Faltan los wait aca
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button#Insert"))).click()
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button#InsertAppsForOffice"))).click()
        
        WebDriverWait(driver,maximo_timeout_element).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe#InsertDialog")))
        
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span#UploadMyAddin"))).click()
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#BrowseButton"))).click()
        
        
        save_as_hwnd = 0
        tiempo_comienzo = time.time()
        while True:
            if(time.time() - tiempo_comienzo > maximo_timeout_element):
                break
            
            save_as_hwnd = buscar_ventana_AbrirUbicacion()
            
            if(save_as_hwnd != 0):
                print_dialog_children(save_as_hwnd)
                break
            time.sleep(0.1)
        
        assert save_as_hwnd != 0, "Timeout completado, No se ha encontrado la ventana de windows"
        
        
        WebDriverWait(driver,maximo_timeout_element).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#DialogInstall"))).click()
        WebDriverWait(driver, timeout = maximo_timeout_element).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe#WacFrame_Excel_0")))
        WebDriverWait(driver, timeout = maximo_timeout_element).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, f"iframe[title='{nombre_addin}']")))
      
def extraer_datos():
    ultimo_rango = None
    with open(ruta_almacenar_csv + "ultimo_rango_extraccion.txt", 'r') as f:
        ultimo_rango = f.read()
     
    try:
        ultimo_rango = int(ultimo_rango)
    except ValueError:
        print("Valor de ultimo rango invalido, revisar archivo")
        raise
        
    
    automatizacion_login()
    automatizacion_carga_extension()
    datos_excel = driver.execute_script(f"return traer_casos_no_procesados_ultimoRango({ultimo_rango}).then(e => resultado_a_csv(e))")
    with open(ruta_almacenar_csv + nombre_archivo_csv, 'w') as f:
        f.write(datos_excel)

    print("ejecucion finalizada")
    driver.quit()        


def cargar_datos():
    automatizacion_login()
    automatizacion_carga_extension()
    datos_csv = open(ruta_almacenar_csv + nombre_archivo_csv, 'r').read()
    print(f"`{datos_csv}`")
    driver.execute_script(f"await pegar_csv(`{datos_csv}`)")
    print("ejecucion finalizada")
    driver.quit()        

def buscar_ventana_AbrirUbicacion():
    hwnd_resp = 0
    
    def buscar_ventana_AbrirUbicacion_callback(hwnd, _):
        nonlocal hwnd_resp
        if win32gui.IsWindowVisible(hwnd):
            window_class = win32gui.GetClassName(hwnd)
            if(window_class == "#32770"): #Codigo de Clase de Ventanas Guardar como, Abrir Ubicacion, ETC
                window_parent = win32gui.GetParent(hwnd)
                window_parent_title = win32gui.GetWindowText(window_parent)
                if("Google Chrome" in window_parent_title):
                    print(hwnd, window_parent_title)
                    hwnd_resp = hwnd
                    return False 
        
    try:
        win32gui.EnumWindows(buscar_ventana_AbrirUbicacion_callback, None)
    except:
        pass
    return hwnd_resp
    # return dialogs

def set_dialog_text(hwnd, text):
    win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)

def print_dialog_children(hwnd):
    def enum_child_windows_callback(child_hwnd, _):
        window_class = win32gui.GetClassName(child_hwnd)
        
        if window_class == "Edit":
            # Modify the text in the path input box
            set_dialog_text(child_hwnd, new_text)
            print("Text set successfully")
            
            # Find the "Open" button control in the dialog
            open_button_hwnd = win32gui.FindWindowEx(hwnd, 0, "Button", "&Abrir")

            if open_button_hwnd:
                # Click the "Open" button
                win32api.SendMessage(open_button_hwnd, win32con.BM_CLICK, 0, 0)
                
                print("Button clicked successfully")
            else:
                print("Failed to find the 'Open' button")
            
            return False #EnumChildWindows continues until the last child window is enumerated or the callback function returns FALSE. https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumchildwindows
            

    
    win32gui.EnumChildWindows(hwnd, enum_child_windows_callback, None)
   

# Get the list of dialogs with window IDs and titles
#extraer_datos()

