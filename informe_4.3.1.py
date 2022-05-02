from ast import Index
from cProfile import label
from numpy import append
import pandas as pd
from colorama import Back, Fore, init

# * LLEGENDA:
# TODO Coses a fer
# ? Dubtes i aclaracions
# ! Coses a corregir
# () Explicacio
# ¿ Control de versio


#¿ V_4.4. Pretenc eliminar: '=' i cometes del excel final.
#¿ V_4.3. Afegit AppCon


#() ##### V_4.3 #####

##Con esta funcion leo el excel y lo convierto en una lista con TODOS los datos que necesito.
def lectura_webfilter_pishing_malicius():


    lista_webfilter_pishing = []
    lista_webfilter_malicius = []

    #print (Fore.LIGHTGREEN_EX+"Leyendo y guardando parametros de la fx webfilter_pishing"+Fore.RESET)

    

    WF_Pishing_all = pd.read_excel('WF_Pishing_ALL.xlsx')
    WF_Malicius_all = pd.read_excel('WF_Malicius_ALL.xlsx')

    #print (data)
    data_selected_WF_Pishing = (WF_Pishing_all[['date','time','sede','srcip','url']])
    data_selected_WF_Malicius = (WF_Malicius_all[['date','time','sede','srcip','url']])
    #print (data_selected)

        

    for fila in data_selected_WF_Pishing.index:
        #print (data_selected['date'][fila]+' '+data_selected['time'][fila]+' '+data_selected['srcip'][fila]+' '+data_selected['url'][fila])
        fila_datos = (data_selected_WF_Pishing['date'][fila]+' '+data_selected_WF_Pishing['time'][fila]+' '+data_selected_WF_Pishing['sede'][fila]+' '+data_selected_WF_Pishing['srcip'][fila]+' '+data_selected_WF_Pishing['url'][fila])
        lista_webfilter_pishing.append(fila_datos)


    for fila in data_selected_WF_Malicius.index:
        fila_datos = (data_selected_WF_Malicius['date'][fila]+' '+data_selected_WF_Malicius['time'][fila]+' '+data_selected_WF_Malicius['sede'][fila]+' '+data_selected_WF_Malicius['srcip'][fila]+' '+data_selected_WF_Malicius['url'][fila])
        lista_webfilter_malicius.append(fila_datos)
        
    #print (lista_webfilter_pishing)
    #print (type(lista_webfilter_pishing))

    return lista_webfilter_pishing, lista_webfilter_malicius

def lectura_AppCon_Proxy_P2P():
    

    lista_AppCon_proxy = []
    lista_AppCon_p2p = []

    AppCon_Proxy_all = pd.read_excel('AppCon_Proxy_ALL.xlsx')
    AppCon_P2P_all = pd.read_excel('AppCon_P2P_ALL.xlsx')

    #print (data)
    data_selected_AppCon_Proxy = (AppCon_Proxy_all[['date','time','sede','srcip','app']])
    data_selected_AppCon_P2P = (AppCon_P2P_all[['date','time','sede','srcip','app']])
    #print (data_selected)

        

    for fila in data_selected_AppCon_Proxy.index:
        #print (data_selected['date'][fila]+' '+data_selected['time'][fila]+' '+data_selected['srcip'][fila]+' '+data_selected['url'][fila])
        fila_datos = (data_selected_AppCon_Proxy['date'][fila]+' '+data_selected_AppCon_Proxy['time'][fila]+' '+data_selected_AppCon_Proxy['sede'][fila]+' '+data_selected_AppCon_Proxy['srcip'][fila]+' '+data_selected_AppCon_Proxy['app'][fila])
        lista_AppCon_proxy.append(fila_datos)


    for fila in data_selected_AppCon_P2P.index:
        fila_datos = (data_selected_AppCon_P2P['date'][fila]+' '+data_selected_AppCon_P2P['time'][fila]+' '+data_selected_AppCon_P2P['sede'][fila]+' '+data_selected_AppCon_P2P['srcip'][fila]+' '+data_selected_AppCon_P2P['app'][fila])
        lista_AppCon_p2p.append(fila_datos)
        
    #print (lista_webfilter_pishing)
    #print (type(lista_webfilter_pishing))

    return lista_AppCon_proxy, lista_AppCon_p2p

def lectura_Intrusion_Prevention():
    

    lista_intrusion_prevention = []


    instrusion_prevention_all = pd.read_excel('IPS_ALL.xlsx')

    #print (data)
    data_selected_intrusion_prevention = (instrusion_prevention_all[['date','time','sede','srcip','attack']])

    #print (data_selected)

        

    for fila in data_selected_intrusion_prevention.index:
        #print (data_selected['date'][fila]+' '+data_selected['time'][fila]+' '+data_selected['srcip'][fila]+' '+data_selected['url'][fila])
        fila_datos = (data_selected_intrusion_prevention['date'][fila]+' '+data_selected_intrusion_prevention['time'][fila]+' '+data_selected_intrusion_prevention['sede'][fila]+' '+data_selected_intrusion_prevention['srcip'][fila]+' '+data_selected_intrusion_prevention['attack'][fila])
        lista_intrusion_prevention.append(fila_datos)


        
    #print (lista_webfilter_pishing)
    #print (type(lista_webfilter_pishing))

    return lista_intrusion_prevention

def lista_por_sede(logs):

    Lucta_Colombia = []
    Lucta_Queretaro = []
    Lucta_Central = []
    Lucta_Coll = []
    Lucta_Eureka = []
    Lucta_Granja = []
    Lucta_Guangzhou = []

    #print (type(lista_sedes))
    #print(logs)

    #print (logs)

    for i in logs:
        elemento = i.split(' ')
        #print (elemento)
        if 'LUCTA_Colombia' in elemento[2]:
            Lucta_Colombia.append(elemento)

        elif 'Lucta_Queretaro' in elemento[2]:
            Lucta_Queretaro.append(elemento)

        elif 'Lucta_Central' in elemento[2]:
            Lucta_Central.append(elemento)

        elif 'HA-Coll' in elemento[2]:
            Lucta_Coll.append(elemento)

        elif 'FW-EUREKA' in elemento[2]:
            Lucta_Eureka.append(elemento)
        
        elif 'GRANJA' in elemento[2]:
            Lucta_Granja.append(elemento)

        elif 'Lucta_Guangzhou' in elemento[2]:
            Lucta_Guangzhou.append(elemento)

    
    #print (Fore.LIGHTMAGENTA_EX+str(Lucta_Colombia)+Fore.RESET)
    #print (Fore.LIGHTMAGENTA_EX+str(Lucta_Queretaro)+Fore.RESET)
    #print (Fore.LIGHTMAGENTA_EX+str(Lucta_Central)+Fore.RESET)
    #print (Fore.LIGHTMAGENTA_EX+str(Lucta_Coll)+Fore.RESET)
    #print (Fore.LIGHTMAGENTA_EX+str(Lucta_Eureka)+Fore.RESET)


    return Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou


#() Las funciones Filtraje de datos, eliminan la informacion que NO ira al informe.
#() Como parametro, se pasa la lista de cada una de las sedes obtenida anteriormente.

# ! ES POSIBLE QUE PUGUI ELIMINAR FILTRAJE_WEBFILTER I FILTRAJE_APPCON
# ! PQE FILTRAJE FA LA TASCA DE TOTES DUES.

def filtraje_webfilter_pishing(lista_de_logs):

    lista_IP_unicas = []
    lista_primera_IP = []   
    nuevos_elementos = []
    apoyo = []
    
    

    print ("hola")
    for elemento in lista_de_logs:
        if elemento[3] not in lista_IP_unicas:
            lista_IP_unicas.append(elemento[3])
            lista_primera_IP.append(elemento)
    

#() En lista lista_primera_IP tenemos la fila completa que ha salido por primera vez

        
    print ("hola2")
    for elemento in lista_primera_IP:
        #print (elemento)
        for elemento2 in lista_de_logs:
            #print (elemento2)

            if (elemento[3]) == (elemento2[3]):
                #print ("posible match")

                if (elemento[4]) != (elemento2[4]):
                    print ("match absoluto")
                    new_match = (elemento[3])+' '+(elemento2[4])
                    print (Fore.GREEN+new_match+Fore.RESET)

                    if new_match not in apoyo:
                        apoyo.append(new_match)
                        print (apoyo)
                        nuevos_elementos.append(elemento2)

                    elif new_match in apoyo:
                        pass

                    
    

    lista_primera_IP.extend(nuevos_elementos)

    '''

    for elemento in lista_primera_IP:
        print (Fore.RED+str(elemento)+Fore.RESET)

    '''

    return lista_primera_IP

def filtraje_AppCon_Proxy_P2P(lista_de_logs):
    
    lista_IP_unicas = []
    lista_primera_IP = []   
    nuevos_elementos = []
    apoyo = []
    
    

    print ("hola")
    for elemento in lista_de_logs:
        if elemento[3] not in lista_IP_unicas:
            lista_IP_unicas.append(elemento[3])
            lista_primera_IP.append(elemento)
    

#() En lista lista_primera_IP tenemos la fila completa que ha salido por primera vez

        
    print ("hola2")
    for elemento in lista_primera_IP:
        #print (elemento)
        for elemento2 in lista_de_logs:
            #print (elemento2)

            if (elemento[3]) == (elemento2[3]):
                #print ("posible match")

                if (elemento[4]) != (elemento2[4]):
                    print ("match absoluto")
                    new_match = (elemento[3])+' '+(elemento2[4])
                    print (Fore.GREEN+new_match+Fore.RESET)

                    if new_match not in apoyo:
                        apoyo.append(new_match)
                        print (apoyo)
                        nuevos_elementos.append(elemento2)

                    elif new_match in apoyo:
                        pass

                    
    

    lista_primera_IP.extend(nuevos_elementos)

    

    for elemento in lista_primera_IP:
        print (Fore.RED+str(elemento)+Fore.RESET)

    

    return lista_primera_IP

def filtraje(lista_de_logs):
    
    lista_IP_unicas = []
    lista_primera_IP = []   
    nuevos_elementos = []
    apoyo = []
    
    

    print ("hola")
    for elemento in lista_de_logs:
        if elemento[3] not in lista_IP_unicas:
            lista_IP_unicas.append(elemento[3])
            lista_primera_IP.append(elemento)
    

#() En lista lista_primera_IP tenemos la fila completa que ha salido por primera vez

        
    print ("hola2")
    for elemento in lista_primera_IP:
        #print (elemento)
        for elemento2 in lista_de_logs:
            #print (elemento2)

            if (elemento[3]) == (elemento2[3]):
                #print ("posible match")

                if (elemento[4]) != (elemento2[4]):
                    print ("match absoluto")
                    new_match = (elemento[3])+' '+(elemento2[4])
                    print (Fore.GREEN+new_match+Fore.RESET)

                    if new_match not in apoyo:
                        apoyo.append(new_match)
                        print (apoyo)
                        nuevos_elementos.append(elemento2)

                    elif new_match in apoyo:
                        pass

                    
    

    lista_primera_IP.extend(nuevos_elementos)

    '''

    for elemento in lista_primera_IP:
        print (Fore.RED+str(elemento)+Fore.RESET)

    '''

    return lista_primera_IP

#() les funciones anomenades parsing_excel... converteixen columnes a DataFrame pq pandas pugui llegir
def parsing_excel_webfilter_pishing(a):

    col_names = ['Fecha', 'Hora', 'Sede','IP Origen', 'URL Destino']
    print ('AQUI2')
    print (len(Fore.CYAN+(str(a))+Fore.RESET))
    df = pd.DataFrame(list(a), columns = col_names)
    print (Fore.RED+df+Fore.RESET)

    return df

def parsing_excel_AppCon_Proxy_P2P(a):
    
    col_names = ['Fecha', 'Hora', 'Sede','IP Origen', 'App']
    #print ('AQUI2')
    #print (len(Fore.CYAN+(str(a))+Fore.RESET))
    df = pd.DataFrame(list(a), columns = col_names)
    print (Fore.RED+df+Fore.RESET)

    return df

def parsing_excel_IPS(a):
    
    col_names = ['Fecha', 'Hora', 'Sede','IP Origen', 'Attack']
    #print ('AQUI2')
    #print (len(Fore.CYAN+(str(a))+Fore.RESET))
    df = pd.DataFrame(list(a), columns = col_names)
    print (Fore.RED+df+Fore.RESET)

    return df

#() el parametro que se pasa es la lista unica por cada sede 
def webfilter_pishing_to_excel(a,b,c,d,e,f,g):

    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-webfilter-pishing.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    c.to_excel(writer,sheet_name='Central', index=False)
    d.to_excel(writer,sheet_name='Coll', index=False)
    e.to_excel(writer,sheet_name='Eureka', index=False)
    f.to_excel(writer,sheet_name='Granja', index=False)
    g.to_excel(writer,sheet_name='Guanzhou', index=False)



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")


    return

def webfilter_malicius_to_excel(a,b,c,d,e,f,g):
    
    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-webfilter-malicius.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    c.to_excel(writer,sheet_name='Central', index=False)
    d.to_excel(writer,sheet_name='Coll', index=False)
    e.to_excel(writer,sheet_name='Eureka', index=False)
    f.to_excel(writer,sheet_name='Granja', index=False)
    g.to_excel(writer,sheet_name='Guangzhou', index=False)



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")


    return

def AppCon_Proxy_to_excel(a,b,c,d,e,f,g):
    
    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-AppCon-Proxy.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    c.to_excel(writer,sheet_name='Central', index=False)
    d.to_excel(writer,sheet_name='Coll', index=False)
    e.to_excel(writer,sheet_name='Eureka', index=False)
    f.to_excel(writer,sheet_name='Granja', index=False)
    g.to_excel(writer,sheet_name='Guangzhou', index=False)



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")


    return

def AppCon_P2P_to_excel(a,b,c,d,e,f,g):
    
    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-AppCon-P2P.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    c.to_excel(writer,sheet_name='Central', index=False)
    d.to_excel(writer,sheet_name='Coll', index=False)
    e.to_excel(writer,sheet_name='Eureka', index=False)
    f.to_excel(writer,sheet_name='Granja', index=False)
    g.to_excel(writer,sheet_name='Guangzhou', index=False)



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")


    return

def IPS_to_excel(a,b,c,d,e,f,g):
    
    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-IPS.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    c.to_excel(writer,sheet_name='Central', index=False)
    d.to_excel(writer,sheet_name='Coll', index=False)
    e.to_excel(writer,sheet_name='Eureka', index=False)
    f.to_excel(writer,sheet_name='Granja', index=False)
    g.to_excel(writer,sheet_name='Guangzhou', index=False)



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")


    return


### MAIN ### 


print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Webfilter Pishing'+Fore.RESET)
lista_webfilter_pishing, lista_webfilter_malicius = lectura_webfilter_pishing_malicius()
Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_webfilter_pishing)
lista_final_colombia = filtraje_webfilter_pishing(Lucta_Colombia)
lista_final_queretaro = filtraje_webfilter_pishing(Lucta_Queretaro)
lista_final_central = filtraje_webfilter_pishing(Lucta_Central)
lista_final_coll = filtraje_webfilter_pishing(Lucta_Coll)
lista_final_eureka = filtraje_webfilter_pishing(Lucta_Eureka)
lista_final_granja = filtraje_webfilter_pishing(Lucta_Granja)
lista_final_guangzhou = filtraje_webfilter_pishing(Lucta_Guangzhou)
df_final_colombia = parsing_excel_webfilter_pishing(lista_final_colombia)
df_final_queretaro = parsing_excel_webfilter_pishing(lista_final_queretaro)
df_final_central = parsing_excel_webfilter_pishing(lista_final_central)
df_final_coll = parsing_excel_webfilter_pishing(lista_final_coll)
df_final_eureka = parsing_excel_webfilter_pishing(lista_final_eureka)
df_final_granja = parsing_excel_webfilter_pishing(lista_final_granja)
df_final_guanzhou = parsing_excel_webfilter_pishing(lista_final_guangzhou)
webfilter_pishing_to_excel(df_final_colombia,df_final_queretaro,df_final_central,df_final_coll,df_final_eureka,df_final_granja,df_final_eureka)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Webfilter Pishing'+Fore.RESET)

print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Webfilter Malicius'+Fore.RESET)

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_webfilter_malicius)
lista_final_colombia = filtraje_webfilter_pishing(Lucta_Colombia)
lista_final_queretaro = filtraje_webfilter_pishing(Lucta_Queretaro)
lista_final_central = filtraje_webfilter_pishing(Lucta_Central)
lista_final_coll = filtraje_webfilter_pishing(Lucta_Coll)
lista_final_eureka = filtraje_webfilter_pishing(Lucta_Eureka)
lista_final_granja = filtraje_webfilter_pishing(Lucta_Granja)
lista_final_guangzhou = filtraje_webfilter_pishing(Lucta_Guangzhou)
df_final_colombia = parsing_excel_webfilter_pishing(lista_final_colombia)
df_final_queretaro = parsing_excel_webfilter_pishing(lista_final_queretaro)
df_final_central = parsing_excel_webfilter_pishing(lista_final_central)
df_final_coll = parsing_excel_webfilter_pishing(lista_final_coll)
df_final_eureka = parsing_excel_webfilter_pishing(lista_final_eureka)
df_final_granja = parsing_excel_webfilter_pishing(lista_final_granja)
df_final_guanzhou = parsing_excel_webfilter_pishing(lista_final_guangzhou)
webfilter_malicius_to_excel(df_final_colombia,df_final_queretaro,df_final_central,df_final_coll,df_final_eureka,df_final_granja,df_final_guanzhou)


print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Webfilter Malicius'+Fore.RESET)

print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)



print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Application Control Proxy'+Fore.RESET)

lista_AppCon_Proxy, lista_AppCon_P2P = lectura_AppCon_Proxy_P2P()
Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_AppCon_Proxy)
lista_final_AppCon_colombia = filtraje_AppCon_Proxy_P2P(Lucta_Colombia)
lista_final_AppCon_queretaro = filtraje_AppCon_Proxy_P2P(Lucta_Queretaro)
lista_final_AppCon_central = filtraje_AppCon_Proxy_P2P(Lucta_Central)
lista_final_AppCon_coll = filtraje_AppCon_Proxy_P2P(Lucta_Coll)
lista_final_AppCon_eureka = filtraje_AppCon_Proxy_P2P(Lucta_Eureka)
lista_final_AppCon_granja = filtraje_AppCon_Proxy_P2P(Lucta_Granja)
lista_final_AppCon_Guangzhou = filtraje_AppCon_Proxy_P2P(Lucta_Guangzhou)
df_final_colombia = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_colombia)
df_final_queretaro = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_queretaro)
df_final_central = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_central)
df_final_coll = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_coll)
df_final_eureka = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_eureka)
df_final_granja = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_granja)
df_final_Guangzhou = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_Guangzhou)
AppCon_Proxy_to_excel(df_final_colombia,df_final_queretaro,df_final_central,df_final_coll,df_final_eureka,df_final_granja,df_final_Guangzhou)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Application Control Proxy'+Fore.RESET)

print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Application Control P2P'+Fore.RESET)

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_AppCon_P2P)
lista_final_AppCon_colombia = filtraje_AppCon_Proxy_P2P(Lucta_Colombia)
lista_final_AppCon_queretaro = filtraje_AppCon_Proxy_P2P(Lucta_Queretaro)
lista_final_AppCon_central = filtraje_AppCon_Proxy_P2P(Lucta_Central)
lista_final_AppCon_coll = filtraje_AppCon_Proxy_P2P(Lucta_Coll)
lista_final_AppCon_eureka = filtraje_AppCon_Proxy_P2P(Lucta_Eureka)
lista_final_AppCon_granja = filtraje_AppCon_Proxy_P2P(Lucta_Granja)
lista_final_AppCon_Guangzhou = filtraje_AppCon_Proxy_P2P(Lucta_Guangzhou)
df_final_colombia = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_colombia)
df_final_queretaro = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_queretaro)
df_final_central = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_central)
df_final_coll = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_coll)
df_final_eureka = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_eureka)
df_final_granja = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_granja)
df_final_Guangzhou = parsing_excel_AppCon_Proxy_P2P(lista_final_AppCon_Guangzhou)
AppCon_P2P_to_excel(df_final_colombia,df_final_queretaro,df_final_central,df_final_coll,df_final_eureka,df_final_granja,df_final_Guangzhou)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Application Control P2P'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe IPS'+Fore.RESET)

lista_IPS = lectura_Intrusion_Prevention()
Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_IPS)
lista_final_IPS_colombia = filtraje(Lucta_Colombia)
lista_final_IPS_queretaro = filtraje(Lucta_Queretaro)
lista_final_IPS_central = filtraje(Lucta_Central)
lista_final_IPS_coll = filtraje(Lucta_Coll)
lista_final_IPS_eureka = filtraje(Lucta_Eureka)
lista_final_IPS_granja = filtraje(Lucta_Granja)
lista_final_IPS_guangzhou = filtraje(Lucta_Guangzhou)
df_final_colombia = parsing_excel_IPS(lista_final_IPS_colombia)
df_final_queretaro = parsing_excel_IPS(lista_final_IPS_queretaro)
df_final_central = parsing_excel_IPS(lista_final_IPS_central)
df_final_coll = parsing_excel_IPS(lista_final_IPS_coll)
df_final_eureka = parsing_excel_IPS(lista_final_IPS_eureka)
df_final_granja = parsing_excel_IPS(lista_final_IPS_granja)
df_final_Guangzhou = parsing_excel_IPS(lista_final_IPS_guangzhou)
IPS_to_excel(df_final_colombia,df_final_queretaro,df_final_central,df_final_coll,df_final_eureka,df_final_granja,df_final_Guangzhou)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe IPS'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'######FINAL######'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)




























