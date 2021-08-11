from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
import datetime , time , sys , os
from pathlib import Path

#QUEDA PENDIENTE
#MOVER EL CALENDARIO ADELANTE CUANDO SEA FEBRERO
#CALIBAR LOS CALENDARIOS CUANDO EL MES SEA MENOR AL DE HOY
#CALIBAR LOS CALENDARIOS CUANDO EL MES SEA MAYOR AL DE HOY

class Flex_scrapping:

    month_dict = [  #diccionario datos mes en ingles e español con su valor numerico
        { 'month':'January' , 'mes':'enero' , 'valor':1 },
        { 'month':'February' , 'mes':'febrero' , 'valor':2},
        { 'month':'March' , 'mes':'marzo' , 'valor':3},
        { 'month':'April' , 'mes':'abril' , 'valor':4},
        { 'month':'May' , 'mes':'mayo' , 'valor':5},
        { 'month':'June' , 'mes':'junio' , 'valor':6},
        { 'month':'July' , 'mes':'julio' , 'valor':7},
        { 'month':'August' , 'mes':'agosto' , 'valor':8},
        { 'month':'September' , 'mes':'septiembre' , 'valor':9},
        { 'month':'October' , 'mes':'octubre' , 'valor':10},
        { 'month':'November' , 'mes':'noviembre' , 'valor':11},
        { 'month':'December' , 'mes':'diciembre' , 'valor':12}
    ]

    #################################################################################

    def valide_export_file(self):

        my_file = Path(r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents\export.xlsx")

        if my_file.is_file():

            os.remove(my_file) #borramos el archivo si existe antes de ejecutar el scrapping

    #################################################################################

    def scrap_conf(self):
        
        try:

            url = "URL"

            chromeOptions = webdriver.ChromeOptions()

            prefs = {
                "download.default_directory" : r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents",
                "download.promt_for_download" : False,
                "download.directory_upgrade" : True,
                "safebrowsing.enabled": True
            }

            chromeOptions.add_experimental_option("prefs",prefs)

            driver = webdriver.Chrome(executable_path="../Flex/SW/chromedriver.exe" , options = chromeOptions)

            driver.implicitly_wait(5)

            driver.get(url)

            time.sleep(10) # mantener esta parte o mejorar

            main = driver.find_element_by_class_name('main')

            side_bar = main.find_element_by_class_name('style_sideBarContainer_1nT7a') #buscar barra izquierda

            side_bar.find_element_by_class_name('style_link-reports_2UvKM').click() # seleccionar Report

            return driver

        except:

            driver.quit()
            
            sys.exit(0)

    #################################################################################   

    def get_date(self,fecha=None):

        #obtener fecha hoy
        today_y = time.strftime("%Y")
        today_m = time.strftime("%m")
        today_d = time.strftime("%d")


        if fecha!=None:

            #obtener fecha personalizado
            today_y = fecha[0]
            today_m = fecha[1]
            today_d = fecha[2]

        #le restamos un dia
        yesterday = datetime.datetime.strptime(str(today_y) + '-' + str(today_m) + '-' + str(today_d), '%Y-%m-%d') + datetime.timedelta(days=0)
        yesterday = datetime.datetime.strptime(str(yesterday.date()),  '%Y-%m-%d') - datetime.timedelta(days=1)

        #obtener el dia, mes , nombre del mes , año de ayer
        today_month = yesterday.strftime("%B")
        year = yesterday.strftime("%Y")
        month = yesterday.strftime("%m")
        day = yesterday.strftime("%d")

        for x in f.month_dict: #recorrer diccionario
            
            if x['month'] == today_month: #validar que el mes de ayer sea igual a uno del diccionario

                today_month = x['mes'] #lo pasamos a español para usarlo en el filtro del calendario

        actual = today_month + " " + str(today_y) #generamos el titulo dle calendario

        yesterday = year + "-" + month + "-" + day #obtenemos la fecha de ayer formateada como texto

        return yesterday , day , actual #retorna la fecha de ayer , el dia de ayer y el titulo del calendario

    #################################################################################

    def select_languague(self,driver):

        driver.find_element_by_class_name('style_name_3zBVE').click() #selecciona opcion de idioma

        t = driver.find_element_by_css_selector('div[class="Select Select--single is-searchable has-value"]') #selecciona comboBox

        t.click() # clic al combobox

        o = t.find_element_by_class_name('Select-menu-outer') #obtener opciones comboBox

        ops = o.find_elements_by_class_name('Select-option') #acceso a hacer click a los elementos

        for x in ops:

            if x.text == 'Español':

                x.click() #seleccionamos español

        return driver

    #################################################################################   

    def fill_check_box(self,driver):

        # Esta futura funcion tiene que ejecutarse 2 veces , total 4

        contador = 0

        check_box = driver.find_element_by_class_name('style_container_5oqce') #buscar opciones

        cb = check_box.find_elements_by_class_name('style_label_d2BbE') #obtener todos los checks

        check_box_list = ['LayeredAudit' , 'CriticalControlObservation' , 'PlannedTaskObservation' , 'TakeTimeTalk']

        for x in cb:

            status = x.find_element_by_name(check_box_list[contador]).is_selected() #obtenemos el estado del check True: seleciconado | False: no selecciconado

            if contador == 1 or contador == 3: #validar las opciones que queremos marcar

                if status:

                    pass
                
                else:

                    x.click() #seleccionamos el check

            elif contador == 0 or contador == 2: #validar las opciones que no queremos marcar

                if status:

                    x.click()

                else:

                    pass

            contador = contador + 1

            return driver

    #################################################################################    

    def select_calendar(self,driver,day):

        calendars = driver.find_element_by_class_name('style_group_3O8Gs') #codigo html con los datepicker

        calendars = calendars.find_elements_by_class_name('SingleDatePicker') #codigo html datepicker inicio y fin

        for c in calendars: #primer loop inicio, segundo fin

            c.click() #abrir calendario

            #?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

            #c.find_element_by_css_selector('button[class="DayPickerNavigation__prev DayPickerNavigation__prev--default"]').click() #mover calendario atras
            #c.click() #tenemos que volver a abrir el calendario par actualizar el codigo html
            # DayPickerNavigation__next DayPickerNavigation__next--default

            #?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

            table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

            for t in table:

                table_title = t.find_element_by_tag_name('strong')

                if len(str(table_title.text)) > 0: #validar que la tabla sea el mes actual

                    print(str(table_title.text))
                    
                    if table_title.text == actual: #validar que el mes y año del calendario sean igual a los del dia de hoy

                        #print(t.get_attribute('innerHTML'))

                        days = t.find_elements_by_class_name('CalendarDay__button') #obtener todos los botones dia del mes

                        for d in days: #recorrer los dias

                            if int(day) == int(d.text): #validar que sea el dia de hoy
                                
                                print(d.text + " <- este es") 

                                d.click()

                                break

                            else:

                                pass

                        break

                    else:

                        pass

                else:

                    pass

        return driver

    #################################################################################

    def pre_filtro(self,driver):

        try:

            driver.find_element_by_class_name('style_remove_3swMA').click()

        except:

            pass

        filtro_base = driver.find_element_by_css_selector('div[class="style_form_28sNE style_root_grXVc"]') #buscamos el html del formulario

        filtro = filtro_base.find_element_by_class_name('Select-control') #buscar combobox del formulario

        filtro.click() #seleccionamos combobox de filtro

        options = driver.find_element_by_class_name('Select-menu-outer') #obtenemos las opciones del filtro

        list_option = options.find_elements_by_class_name('Select-option') #guardamos las opciones en una lista

        for o in list_option: #recorremos las opciones

            if o.text == "Región / asset / operación / lugar visitado" : #validmaos que sea la opcion que queremos

                o.click() #seleccionamos la opcion

                break

        return driver

    #################################################################################  

    def filtro(self,driver,opcion):

        filtros = ['Minerals Americas' , 'Minerals Australia'] #lista con las opciones a filtrar

        driver.find_element_by_class_name('style_advancedSearch_6eDP0').click() #seleccionar icono filtro

        filtro = driver.find_elements_by_css_selector('div[class="style_subLevel_PWXJU"]') #obtener todas las opciones de filtro sin seleccionar

        filtro_seleccionado = None

        try:

            filtro_seleccionado = driver.find_element_by_css_selector('div[class="style_subLevel_PWXJU style_selected_1dOcY"]') #obtener filtro seleccionado por defecto

        except NoSuchElementException:

            pass

        if filtros[opcion] == filtro_seleccionado: #validar que la seleccionada sea la que quiero

            pass #no hacer nada, ya que esta seleccionado

        else: #buscar elemento en los no seleccionados

            for x in filtro:
                print(x.text)
                if x.text == filtros[opcion]:

                    x.click() #seleccionar

                    break

        modal_opcion = driver.find_element_by_class_name('style_buttons_ISpXW') #obtener boton ok y cancelar

        o =  modal_opcion.find_elements_by_tag_name('button') #buscar botones

        for x in o: #recorrer lista

            if x.text == 'OK':

                x.click() #cerramos el modal

        return driver , opcion

    #################################################################################

    def download_file(self,driver,opcion):

        if opcion == 1:

            dir_name = "M_Aus"

            file_name = "M_Aus_" + str(yesterday)

        elif opcion == 0:

            dir_name = "M_Ame"

            file_name = "M_Ame_" + str(yesterday)

        my_file = Path(r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents\export.xlsx")

        file_exist = Path(r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents\\"+dir_name+"\\"+file_name+r".xlsx")

        if file_exist.is_file():

            os.remove(file_exist)

        time.sleep(1)

        driver.find_element_by_class_name('style_yes_2yfPo').click() #incluir sub niveles

        time.sleep(1) # se cayo en la linea siguiente, si vuelve a ocurrir , mejorar esta parte

        driver.find_element_by_class_name('style_export_3XB1c').click() #exportar data

        while True: #esperamos a que finalice la descarga

            if my_file.is_file(): # valida que el archivo exista en el ordenador

                break

        time.sleep(5)

        os.rename(r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents\export.xlsx",
                r"C:\Users\caliev\Documents\Scrapping_bhp\Flex\Documents\\"+dir_name+"\\"+file_name+r".xlsx")

        return driver

    #################################################################################

    def run_functions(self,driver,opcion):

        driver , opcion = self.filtro(driver,opcion)

        driver = self.download_file(driver,opcion)

        return driver

    #################################################################################

############## Comienza ejecucion de clase ######################################

f = Flex_scrapping()

#################################################################################

f.valide_export_file()

#################################################################################

yesterday , day , actual = f.get_date([2019,1,2]) #MANUAL YYYY-MM-DD
#yesterday , day , actual = f.get_date() #AUTOMATICO

#################################################################################

driver = f.scrap_conf()

#################################################################################

driver = f.select_languague(driver)

#################################################################################

driver = f.fill_check_box(driver)
driver = f.fill_check_box(driver)

################################################################################

calendars = driver.find_element_by_class_name('style_group_3O8Gs') #codigo html con los datepicker

calendars = calendars.find_elements_by_class_name('SingleDatePicker') #codigo html datepicker inicio y fin

for c in calendars: #primer loop inicio, segundo fin

    c.click() #abrir calendario

    #c.find_element_by_css_selector('button[class="DayPickerNavigation__prev DayPickerNavigation__prev--default"]').click() #mover calendario atras
    #c.click() #tenemos que volver a abrir el calendario par actualizar el codigo html

    table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

    for t in table:

        table_title = t.find_element_by_tag_name('strong').text

        if len(table_title) > 0: #validar que la tabla sea el mes actual

            actual_m , actual_y = actual.split()

            table_title_m , table_title_y = table_title.split()

            for x in f.month_dict:

                if x['mes'] == table_title_m: #validar que el mes de ayer sea igual a uno del diccionario

                    table_title_m = x['valor'] #lo pasamos a español para usarlo en el filtro del calendario

                if x['mes'] == actual_m:

                    actual_m = x['valor']

            print("table_title: " + str(table_title))
            print("actual: " + str(actual))

            if table_title_y < actual_y or table_title_m < actual_m:

                while table_title_y < actual_y: #validamos que sea el mismo año

                    c.find_element_by_css_selector('button[class="DayPickerNavigation__next DayPickerNavigation__next--default"]').click()
                    c.click()

                    table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

                    for t in table:

                        try:

                            table_title = t.find_element_by_tag_name('strong').text

                        except:

                            table_title = ""

                        if len(table_title) > 0:

                            table_title_m , table_title_y = table_title.split()

                            print(table_title_m)
                            print(table_title_y)

                            for x in f.month_dict: #volver a convertir el mes e texto a entero

                                if x['mes'] == table_title_m: #validar que el mes de ayer sea igual a uno del diccionario

                                    table_title_m = x['valor'] #lo pasamos a español para usarlo en el filtro del calendario

                        else:

                            pass

                print("table_title_m : " + str(table_title_m))
                print("today_m : " + str(actual_m))
                print(type(table_title_m))
                print(type(actual_m))
                while table_title_m  < actual_m :#validamos que sea el mismo mes   
                    print("bucle infinito")
                    c.find_element_by_css_selector('button[class="DayPickerNavigation__next DayPickerNavigation__next--default"]').click()
                    c.click()
                    table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

                    for t in table:

                        try:

                            table_title = t.find_element_by_tag_name('strong').text

                        except:

                            table_title = ""

                        print(table_title)

                        if len(table_title) > 0:

                            table_title_m , table_title_y = table_title.split()

                            print(table_title_m)
                            print(table_title_y)

                            for x in f.month_dict: #volver a convertir el mes e texto a entero

                                if x['mes'] == table_title_m: #validar que el mes de ayer sea igual a uno del diccionario

                                    table_title_m = x['valor'] #lo pasamos a español para usarlo en el filtro del calendario
                                
                            break

                days = t.find_elements_by_class_name('CalendarDay__button') #obtener todos los botones dia del mes

                for d in days: #recorrer los dias

                    if int(day) == int(d.text): #validar que sea el dia de hoy
                        
                        print(d.text + " <- este es") 

                        d.click()

                        break

                print("****************************************++")

                break

            elif table_title_y > actual_y or table_title_m > actual_m:

                print("paso en el futuro")
                while table_title_y > actual_y: #validamos que sea el mismo año

                    c.find_element_by_css_selector('button[class="DayPickerNavigation__prev DayPickerNavigation__prev--default"]').click()
                    c.click()

                    table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

                    for t in table:

                        try:

                            table_title = t.find_element_by_tag_name('strong').text

                        except:

                            table_title = ""

                        if len(table_title) > 0:

                            table_title_m , table_title_y = table_title.split()

                            print(table_title_m)
                            print(table_title_y)

                            for x in f.month_dict: #volver a convertir el mes e texto a entero

                                if x['mes'] == table_title_m: #validar que el mes de ayer sea igual a uno del diccionario

                                    table_title_m = x['valor'] #lo pasamos a español para usarlo en el filtro del calendario

                        else:

                            pass

                print("table_title_m : " + str(table_title_m))
                print("today_m : " + str(actual_m))
                print(type(table_title_m))
                print(type(actual_m))
                while table_title_m  > actual_m :#validamos que sea el mismo mes   
                    print("bucle infinito")
                    c.find_element_by_css_selector('button[class="DayPickerNavigation__prev DayPickerNavigation__prev--default"]').click()
                    c.click()
                    table = c.find_elements_by_css_selector('div[class="CalendarMonth CalendarMonth--horizontal"]') 

                    for t in table:

                        try:

                            table_title = t.find_element_by_tag_name('strong').text

                        except:

                            table_title = ""

                        print(table_title)

                        if len(table_title) > 0:

                            table_title_m , table_title_y = table_title.split()

                            print(table_title_m)
                            print(table_title_y)

              *              for x in f.month_dict: #volver a convertir el mes e texto a entero

                                if x['mes'] == table_title_m: #validar que el mes de ayer sea igual a uno del diccionario

                                    table_title_m = x['valor'] #lo pasamos a español para usarlo en el filtro del calendario
                                
                            break

                days = t.find_elements_by_class_name('CalendarDay__button') #obtener todos los botones dia del mes

                for d in days: #recorrer los dias

                    if int(day) == int(d.text): #validar que sea el dia de hoy
                        
                        print(d.text + " <- este es") 

                        d.click()

                        break

                break   

            elif table_title_y == actual_y:

                pass #el calendario esta calibrado

#sys.exit(0)
################################################################################

driver = f.select_calendar(driver,day)

#################################################################################

driver = f.pre_filtro(driver)

#################################################################################

driver = f.run_functions(driver,0) #americans

time.sleep(2)

f.run_functions(driver,1) #australia

#driver.quit()