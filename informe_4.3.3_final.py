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


#¿ V_4.3.3 Eliminar 'date' i '=' del excel. Done!
#¿ V_4.3.2 Afegit Concatenar parametres a tots els excels de sortida. Columna G i F.

#() ##### V_4.3.3 #####

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


    return Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou

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

#() les funciones anomenades parsing_to_dataframe... converteixen columnes a DataFrame pq pandas pugui llegir
def parsing_to_dataframe(a):
    
    col_names = ['Fecha', 'Hora', 'Sede','IP Origen', 'Destino']
    #print ('AQUI2')
    #print (len(Fore.CYAN+(str(a))+Fore.RESET))
    df = pd.DataFrame(list(a), columns = col_names)
    print (Fore.RED+df+Fore.RESET)

    return df

def delete_caracters(df):

    print (df)

    columnas = ['Fecha', 'Hora', 'Sede', 'IP Origen', 'Destino']

    df['Fecha'] = df['Fecha'].replace({'date=':''},regex=True)
    df['Fecha'] = df['Fecha'].replace({'"':''},regex=True)

    print (df)
    df['Hora'] = df['Hora'].replace({'time=':''},regex=True)
    df['Hora'] = df['Hora'].replace({'"':''},regex=True)
    df['Sede'] = df['Sede'].replace({'devname=':''},regex=True)
    df['Sede'] = df['Sede'].replace({'"':''},regex=True)
    df['IP Origen'] = df['IP Origen'].replace({'srcip=':''},regex=True)
    df['IP Origen'] = df['IP Origen'].replace({'"':''},regex=True)
    df['Destino'] = df['Destino'].replace({'url=':''},regex=True)
    df['Destino'] = df['Destino'].replace({'app=':''},regex=True)
    df['Destino'] = df['Destino'].replace({'"':''},regex=True)
    
    print (df)


    print (df)

    return df

def webfilter_pishing_to_excel(a,b,c,d,e,f,g):
    
    #! Caldria realizat la concatenacio a una funcio diferent.

    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-webfilter-pishing.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    a.to_excel(writer,sheet_name='Colombia', index=False)
    Colombia = writer.sheets['Colombia']
    Colombia['F2'] = '=A2&" "&B2'
    Colombia['G2'] = '=D2&" "&E2'

    b.to_excel(writer,sheet_name='Queretaro', index=False)
    Queretaro = writer.sheets['Queretaro']
    Queretaro['F2'] = '=A2&" "&B2'
    Queretaro['G2'] = '=D2&" "&E2'

    c.to_excel(writer,sheet_name='Central', index=False)
    Central = writer.sheets['Central']
    Central['F2'] = '=A2&" "&B2'
    Central['G2'] = '=D2&" "&E2'

    d.to_excel(writer,sheet_name='Coll', index=False)
    Coll = writer.sheets['Coll']
    Coll['F2'] = '=A2&" "&B2'
    Coll['G2'] = '=D2&" "&E2'

    e.to_excel(writer,sheet_name='Eureka', index=False)
    Eureka = writer.sheets['Eureka']
    Eureka['F2'] = '=A2&" "&B2'
    Eureka['G2'] = '=D2&" "&E2'

    f.to_excel(writer,sheet_name='Granja', index=False)
    Granja = writer.sheets['Granja']
    Granja['F2'] = '=A2&" "&B2'
    Granja['G2'] = '=D2&" "&E2'

    g.to_excel(writer,sheet_name='Guanzhou', index=False)
    Guanzhou = writer.sheets['Guanzhou']
    Guanzhou['F2'] = '=A2&" "&B2'
    Guanzhou['G2'] = '=D2&" "&E2'



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")

def webfilter_malicius_to_excel(a,b,c,d,e,f,g):
    
    
    print (type(a))

    print ("Abriendo Excel para guardar DataFrames...")
    writer = pd.ExcelWriter('C:\\Users\\marc.robles\\Desktop\\Informe\\informe-webfilter-malicius.xlsx')
    #df_checklist = pd.DataFrame(lista)
    #print (df_checklist)


    
    a.to_excel(writer,sheet_name='Colombia', index=False)
    Colombia = writer.sheets['Colombia']
    Colombia['F2'] = '=A2&" "&B2'
    Colombia['G2'] = '=D2&" "&E2'

    b.to_excel(writer,sheet_name='Queretaro', index=False)
    Queretaro = writer.sheets['Queretaro']
    Queretaro['F2'] = '=A2&" "&B2'
    Queretaro['G2'] = '=D2&" "&E2'

    c.to_excel(writer,sheet_name='Central', index=False)
    Central = writer.sheets['Central']
    Central['F2'] = '=A2&" "&B2'
    Central['G2'] = '=D2&" "&E2'

    d.to_excel(writer,sheet_name='Coll', index=False)
    Coll = writer.sheets['Coll']
    Coll['F2'] = '=A2&" "&B2'
    Coll['G2'] = '=D2&" "&E2'

    e.to_excel(writer,sheet_name='Eureka', index=False)
    Eureka = writer.sheets['Eureka']
    Eureka['F2'] = '=A2&" "&B2'
    Eureka['G2'] = '=D2&" "&E2'

    f.to_excel(writer,sheet_name='Granja', index=False)
    Granja = writer.sheets['Granja']
    Granja['F2'] = '=A2&" "&B2'
    Granja['G2'] = '=D2&" "&E2'

    g.to_excel(writer,sheet_name='Guanzhou', index=False)
    Guanzhou = writer.sheets['Guanzhou']
    Guanzhou['F2'] = '=A2&" "&B2'
    Guanzhou['G2'] = '=D2&" "&E2'



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
    Colombia = writer.sheets['Colombia']
    Colombia['F2'] = '=A2&" "&B2'
    Colombia['G2'] = '=D2&" "&E2'
    
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    Queretaro = writer.sheets['Queretaro']
    Queretaro['F2'] = '=A2&" "&B2'
    Queretaro['G2'] = '=D2&" "&E2'

    c.to_excel(writer,sheet_name='Central', index=False)
    Central = writer.sheets['Central']
    Central['F2'] = '=A2&" "&B2'
    Central['G2'] = '=D2&" "&E2'

    d.to_excel(writer,sheet_name='Coll', index=False)
    Coll = writer.sheets['Coll']
    Coll['F2'] = '=A2&" "&B2'
    Coll['G2'] = '=D2&" "&E2'

    e.to_excel(writer,sheet_name='Eureka', index=False)
    Eureka = writer.sheets['Eureka']
    Eureka['F2'] = '=A2&" "&B2'
    Eureka['G2'] = '=D2&" "&E2'

    f.to_excel(writer,sheet_name='Granja', index=False)
    Granja = writer.sheets['Granja']
    Granja['F2'] = '=A2&" "&B2'
    Granja['G2'] = '=D2&" "&E2'

    g.to_excel(writer,sheet_name='Guanzhou', index=False)
    Guanzhou = writer.sheets['Guanzhou']
    Guanzhou['F2'] = '=A2&" "&B2'
    Guanzhou['G2'] = '=D2&" "&E2'

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
    Colombia = writer.sheets['Colombia']
    Colombia['F2'] = '=A2&" "&B2'
    Colombia['G2'] = '=D2&" "&E2'
    
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    Queretaro = writer.sheets['Queretaro']
    Queretaro['F2'] = '=A2&" "&B2'
    Queretaro['G2'] = '=D2&" "&E2'

    c.to_excel(writer,sheet_name='Central', index=False)
    Central = writer.sheets['Central']
    Central['F2'] = '=A2&" "&B2'
    Central['G2'] = '=D2&" "&E2'

    d.to_excel(writer,sheet_name='Coll', index=False)
    Coll = writer.sheets['Coll']
    Coll['F2'] = '=A2&" "&B2'
    Coll['G2'] = '=D2&" "&E2'

    e.to_excel(writer,sheet_name='Eureka', index=False)
    Eureka = writer.sheets['Eureka']
    Eureka['F2'] = '=A2&" "&B2'
    Eureka['G2'] = '=D2&" "&E2'

    f.to_excel(writer,sheet_name='Granja', index=False)
    Granja = writer.sheets['Granja']
    Granja['F2'] = '=A2&" "&B2'
    Granja['G2'] = '=D2&" "&E2'

    g.to_excel(writer,sheet_name='Guanzhou', index=False)
    Guanzhou = writer.sheets['Guanzhou']
    Guanzhou['F2'] = '=A2&" "&B2'
    Guanzhou['G2'] = '=D2&" "&E2'

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
    Colombia = writer.sheets['Colombia']
    Colombia['F2'] = '=A2&" "&B2'
    Colombia['G2'] = '=D2&" "&E2'
    
    b.to_excel(writer,sheet_name='Queretaro', index=False)
    Queretaro = writer.sheets['Queretaro']
    Queretaro['F2'] = '=A2&" "&B2'
    Queretaro['G2'] = '=D2&" "&E2'

    c.to_excel(writer,sheet_name='Central', index=False)
    Central = writer.sheets['Central']
    Central['F2'] = '=A2&" "&B2'
    Central['G2'] = '=D2&" "&E2'

    d.to_excel(writer,sheet_name='Coll', index=False)
    Coll = writer.sheets['Coll']
    Coll['F2'] = '=A2&" "&B2'
    Coll['G2'] = '=D2&" "&E2'

    e.to_excel(writer,sheet_name='Eureka', index=False)
    Eureka = writer.sheets['Eureka']
    Eureka['F2'] = '=A2&" "&B2'
    Eureka['G2'] = '=D2&" "&E2'

    f.to_excel(writer,sheet_name='Granja', index=False)
    Granja = writer.sheets['Granja']
    Granja['F2'] = '=A2&" "&B2'
    Granja['G2'] = '=D2&" "&E2'

    g.to_excel(writer,sheet_name='Guanzhou', index=False)
    Guanzhou = writer.sheets['Guanzhou']
    Guanzhou['F2'] = '=A2&" "&B2'
    Guanzhou['G2'] = '=D2&" "&E2'



    writer.save()
    writer.close()

    print ("Fichero Excel Cerrado")





    return writer



## MAIN ##



print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Webfilter Pishing'+Fore.RESET)

lista_webfilter_pishing, lista_webfilter_malicius = lectura_webfilter_pishing_malicius()

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_webfilter_pishing)

lista_final_colombia = filtraje(Lucta_Colombia)
lista_final_queretaro = filtraje(Lucta_Queretaro)
lista_final_central = filtraje(Lucta_Central)
lista_final_coll = filtraje(Lucta_Coll)
lista_final_eureka = filtraje(Lucta_Eureka)
lista_final_granja = filtraje(Lucta_Granja)
lista_final_guangzhou = filtraje(Lucta_Guangzhou)

df_final_colombia = parsing_to_dataframe(lista_final_colombia)
df_final_queretaro = parsing_to_dataframe(lista_final_queretaro)
df_final_central = parsing_to_dataframe(lista_final_central)
df_final_coll = parsing_to_dataframe(lista_final_coll)
df_final_eureka = parsing_to_dataframe(lista_final_eureka)
df_final_granja = parsing_to_dataframe(lista_final_granja)
df_final_guanzhou = parsing_to_dataframe(lista_final_guangzhou)

df_final_colombia_cleaned = delete_caracters(df_final_colombia)
df_final_queretaro_cleaned = delete_caracters(df_final_queretaro)
df_final_central_cleaned = delete_caracters(df_final_central)
df_final_coll_cleaned = delete_caracters(df_final_coll)
df_final_eureka_cleaned = delete_caracters(df_final_eureka)
df_final_granja_cleaned = delete_caracters(df_final_granja)
df_final_guanzhou_cleaned = delete_caracters(df_final_guanzhou)

webfilter_pishing_to_excel(df_final_colombia_cleaned,df_final_queretaro_cleaned,df_final_central_cleaned,df_final_coll_cleaned,df_final_eureka_cleaned,df_final_granja_cleaned,df_final_guanzhou_cleaned)
    

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Webfilter Pishing'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Webfilter Malicius'+Fore.RESET)

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_webfilter_malicius)

lista_final_colombia = filtraje(Lucta_Colombia)
lista_final_queretaro = filtraje(Lucta_Queretaro)
lista_final_central = filtraje(Lucta_Central)
lista_final_coll = filtraje(Lucta_Coll)
lista_final_eureka = filtraje(Lucta_Eureka)
lista_final_granja = filtraje(Lucta_Granja)
lista_final_guangzhou = filtraje(Lucta_Guangzhou)

df_final_colombia = parsing_to_dataframe(lista_final_colombia)
df_final_queretaro = parsing_to_dataframe(lista_final_queretaro)
df_final_central = parsing_to_dataframe(lista_final_central)
df_final_coll = parsing_to_dataframe(lista_final_coll)
df_final_eureka = parsing_to_dataframe(lista_final_eureka)
df_final_granja = parsing_to_dataframe(lista_final_granja)
df_final_guanzhou = parsing_to_dataframe(lista_final_guangzhou)

df_final_colombia_cleaned = delete_caracters(df_final_colombia)
df_final_queretaro_cleaned = delete_caracters(df_final_queretaro)
df_final_central_cleaned = delete_caracters(df_final_central)
df_final_coll_cleaned = delete_caracters(df_final_coll)
df_final_eureka_cleaned = delete_caracters(df_final_eureka)
df_final_granja_cleaned = delete_caracters(df_final_granja)
df_final_guanzhou_cleaned = delete_caracters(df_final_guanzhou)

webfilter_malicius_to_excel(df_final_colombia_cleaned,df_final_queretaro_cleaned,df_final_central_cleaned,df_final_coll_cleaned,df_final_eureka_cleaned,df_final_granja_cleaned,df_final_guanzhou_cleaned)


print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Webfilter Malicius'+Fore.RESET)

print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)



print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Application Control Proxy'+Fore.RESET)

lista_AppCon_Proxy, lista_AppCon_P2P = lectura_AppCon_Proxy_P2P()

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_AppCon_Proxy)

lista_final_colombia = filtraje(Lucta_Colombia)
lista_final_queretaro = filtraje(Lucta_Queretaro)
lista_final_central = filtraje(Lucta_Central)
lista_final_coll = filtraje(Lucta_Coll)
lista_final_eureka = filtraje(Lucta_Eureka)
lista_final_granja = filtraje(Lucta_Granja)
lista_final_guangzhou = filtraje(Lucta_Guangzhou)

df_final_colombia = parsing_to_dataframe(lista_final_colombia)
df_final_queretaro = parsing_to_dataframe(lista_final_queretaro)
df_final_central = parsing_to_dataframe(lista_final_central)
df_final_coll = parsing_to_dataframe(lista_final_coll)
df_final_eureka = parsing_to_dataframe(lista_final_eureka)
df_final_granja = parsing_to_dataframe(lista_final_granja)
df_final_guanzhou = parsing_to_dataframe(lista_final_guangzhou)

df_final_colombia_cleaned = delete_caracters(df_final_colombia)
df_final_queretaro_cleaned = delete_caracters(df_final_queretaro)
df_final_central_cleaned = delete_caracters(df_final_central)
df_final_coll_cleaned = delete_caracters(df_final_coll)
df_final_eureka_cleaned = delete_caracters(df_final_eureka)
df_final_granja_cleaned = delete_caracters(df_final_granja)
df_final_guanzhou_cleaned = delete_caracters(df_final_guanzhou)

AppCon_Proxy_to_excel(df_final_colombia_cleaned,df_final_queretaro_cleaned,df_final_central_cleaned,df_final_coll_cleaned,df_final_eureka_cleaned,df_final_granja_cleaned,df_final_guanzhou_cleaned)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Application Control Proxy'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe Application Control P2P'+Fore.RESET)

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_AppCon_P2P)

lista_final_colombia = filtraje(Lucta_Colombia)
lista_final_queretaro = filtraje(Lucta_Queretaro)
lista_final_central = filtraje(Lucta_Central)
lista_final_coll = filtraje(Lucta_Coll)
lista_final_eureka = filtraje(Lucta_Eureka)
lista_final_granja = filtraje(Lucta_Granja)
lista_final_guangzhou = filtraje(Lucta_Guangzhou)

df_final_colombia = parsing_to_dataframe(lista_final_colombia)
df_final_queretaro = parsing_to_dataframe(lista_final_queretaro)
df_final_central = parsing_to_dataframe(lista_final_central)
df_final_coll = parsing_to_dataframe(lista_final_coll)
df_final_eureka = parsing_to_dataframe(lista_final_eureka)
df_final_granja = parsing_to_dataframe(lista_final_granja)
df_final_guanzhou = parsing_to_dataframe(lista_final_guangzhou)

df_final_colombia_cleaned = delete_caracters(df_final_colombia)
df_final_queretaro_cleaned = delete_caracters(df_final_queretaro)
df_final_central_cleaned = delete_caracters(df_final_central)
df_final_coll_cleaned = delete_caracters(df_final_coll)
df_final_eureka_cleaned = delete_caracters(df_final_eureka)
df_final_granja_cleaned = delete_caracters(df_final_granja)
df_final_guanzhou_cleaned = delete_caracters(df_final_guanzhou)

AppCon_P2P_to_excel(df_final_colombia_cleaned,df_final_queretaro_cleaned,df_final_central_cleaned,df_final_coll_cleaned,df_final_eureka_cleaned,df_final_granja_cleaned,df_final_guanzhou_cleaned)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe Application Control P2P'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)

print (Fore.LIGHTMAGENTA_EX+'Iniciando script para informe IPS'+Fore.RESET)

Lucta_Colombia, Lucta_Queretaro, Lucta_Central, Lucta_Coll, Lucta_Eureka, Lucta_Granja, Lucta_Guangzhou = lista_por_sede(lista_AppCon_P2P)

lista_final_colombia = filtraje(Lucta_Colombia)
lista_final_queretaro = filtraje(Lucta_Queretaro)
lista_final_central = filtraje(Lucta_Central)
lista_final_coll = filtraje(Lucta_Coll)
lista_final_eureka = filtraje(Lucta_Eureka)
lista_final_granja = filtraje(Lucta_Granja)
lista_final_guangzhou = filtraje(Lucta_Guangzhou)

df_final_colombia = parsing_to_dataframe(lista_final_colombia)
df_final_queretaro = parsing_to_dataframe(lista_final_queretaro)
df_final_central = parsing_to_dataframe(lista_final_central)
df_final_coll = parsing_to_dataframe(lista_final_coll)
df_final_eureka = parsing_to_dataframe(lista_final_eureka)
df_final_granja = parsing_to_dataframe(lista_final_granja)
df_final_guanzhou = parsing_to_dataframe(lista_final_guangzhou)

df_final_colombia_cleaned = delete_caracters(df_final_colombia)
df_final_queretaro_cleaned = delete_caracters(df_final_queretaro)
df_final_central_cleaned = delete_caracters(df_final_central)
df_final_coll_cleaned = delete_caracters(df_final_coll)
df_final_eureka_cleaned = delete_caracters(df_final_eureka)
df_final_granja_cleaned = delete_caracters(df_final_granja)
df_final_guanzhou_cleaned = delete_caracters(df_final_guanzhou)

IPS_to_excel(df_final_colombia_cleaned,df_final_queretaro_cleaned,df_final_central_cleaned,df_final_coll_cleaned,df_final_eureka_cleaned,df_final_granja_cleaned,df_final_guanzhou_cleaned)

print (Fore.LIGHTMAGENTA_EX+'Finalizando script para informe IPS'+Fore.RESET)


print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'######FINAL######'+Fore.RESET)
print (Fore.LIGHTGREEN_EX+'#################'+Fore.RESET)





















