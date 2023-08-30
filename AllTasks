# -*- coding: utf-8 -*-
#!/usr/bin/python
#! Python

import pyautogui
import sched #este se ocupa de los timepos
import time
import datetime
import subprocess
import re
import os
import os.path as path
import pathlib
import shutil
import errno
import sys
import copy
import pyperclip as clipboard
import getpass
import pyscreeze
from shutil import rmtree
from getpass import getuser
#from VarPassUserSA import *

############################################################################

#COMANDO Pre USER/PASS!!!
userDeWindows = getuser()

time.strftime("%H:%M:%S")
horaEnvio = time.strftime("%H:%M:%S")
DiaYHoraDeEnvio = time.asctime()
hora = time.strftime("%H")
minutos = time.strftime("%M")
segundos = time.strftime("%S")
dia = time.strftime("%d-%m-%y")

EJECUTAR = 'ejecutar'
FIREFOX = 'firefox'
webZabbix = 'https://zabbix/zabbix.php?action=dashboard.view&ddreset=1'
sa = str('c:\\SA\\sa.exe')
saDir = str('c:\\SA\\')
navegador = 'firefox'

compartida = r'\\10.10.100.6\dccpublico\LazyGv'
encompartida = r'\\10.10.100.6\dccpublico\LazyGv'

carpetaDescargasFirefox = str('c:\\Users\\' + userDeWindows + '\\Downloads\\')

#Escritorio de Usuario
escritorioUser = str('C:\\Users\\' + userDeWindows + '\\Desktop\\')

#Carpeta de Usuario
CptUsuario = str(escritorioUser + 'Lazy' + userDeWindows + '\\')

carpetaDevoluciones = str(CptUsuario + 'Devoluciones'+'\\')
carpetaCortes = str(CptUsuario + 'Cortes'+'\\')
passPath = str(CptUsuario + 'User+Pass'+'\\')
passSaTxt = str(passPath + userDeWindows + '-SaPass.txt')

carpetaEnvios = str(CptUsuario + 'Envios'+'\\')
carpetaConexiones = str(carpetaEnvios + 'Conexiones'+'\\')
carpetaDatos = str(carpetaEnvios  + 'Datos'+'\\')
carpetaReclamos = str(carpetaEnvios  + 'Reclamos'+'\\')

#Estructura Nombre de Archivo
nombreLogCC = hora + '-' + minutos + '-' + segundos + '-' + dia

#vars de hora de envio
ultimoEnvioConexion = str(0)
ultimoEnvioDatos = str(0)
ultimoEnvioReclamos = str(0)

############################################################################

#Crear Carpeta de Carga de Usuario
try:
    os.mkdir(CptUsuario)
    
except OSError as e:
    if e.errno != errno.EEXIST:
        raise

############################################################################

#Crear Carpeta de Pass de Usuario
try:
    os.mkdir(passPath)
    
except OSError as e:
    if e.errno != errno.EEXIST:
        raise

############################################################################

#Crear Carpeta de Cortes Diario
try:
    os.mkdir(carpetaCortes)
    
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################

#Carpeta de Devoluciones
try:
    os.mkdir(carpetaDevoluciones)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################

#Carpeta de Envios
try:
    os.mkdir(carpetaEnvios)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################
    
#Carpeta de Conexiones
try:
    os.mkdir(carpetaConexiones)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################

#Carpeta de Datos
try:
    os.mkdir(carpetaDatos)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################

#Carpeta de Reclamos
try:
    os.mkdir(carpetaReclamos)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    
############################################################################
    
#Levanta PASS de SA si esta en el Sistema sino la Pide

if path.exists(passSaTxt) == True:
    passRead = open(passSaTxt, 'r')
    largoPass = len(passRead.read())
    passRead.close()

    if largoPass >= 1:
        passRead = open(passSaTxt, 'r')
        global PASS
        PASS = passRead.read()
        passRead.close()

else:
    print("Ingrese Pass de SA")
    passSA = input()
    passWrite = open(passSaTxt, 'w')
    passWrite.write(passSA)
    passWrite.close()

#############################################################################

#Habilitar Scaneo de Pantalla
def hScan():
    print('Activar SCan de Pantalla? on/off     '+'\n'+
          '                                     '+'IMPORTANTE:'+'\n'+
          '                                     '+'Gracias a contar con excelentes'+'\n'+
          '                                     '+'equipos en la guardia es que la'+'\n'+
          '                                     '+'siguiente funcion puede requeri'+'\n'+    
          '                                     '+'mas recursos de los que se cuenta'+'\n'+
          '                                     '+'por lo que puede andar MUY lenta'+'\n')
    print('Eleccion:')
    opHScan = input()

    if opHScan != '1' and eleccion != '0' and eleccion != 'on' and eleccion != 'off' and eleccion != 'On' and eleccion != 'Off':
        print('\n' +
              '############################################################' + '\n' +
              "#  Opcion Inexistente Elegir (Off/On) o (0/1) "'             #' + '\n' +
              '############################################################' + '\n' )
        hScan()

    else:
        
        if opHScan == '1':
            prendeScan = '1'
            opciones()
            
        if opHScan == '0':
            prendeScan = '0'
            opciones()
              
#############################################################################

def scanScr():
    conEtiqueta = bool(pyautogui.locateOnScreen('C:\\Users\\gvolpi\\Desktop\\444\\etiqueta.png'))
    if conEtiqueta == True:
        pyautogui.hotkey('enter')
        print('Ya etiquetado')
        pyautogui.hotkey('enter')

def scanScrInv():
    numeroInvalido = bool(pyautogui.locateOnScreen('C:\\Users\\gvolpi\\Desktop\\444\\numeroInvalido.png'))
    if numeroInvalido == True:
        pyautogui.hotkey('esc')
        print('Se encontro')

############################################################################

def ejecutarEnvios():
    print('Veces? (1 to loop)' + "loop = ''L''")
    ingreso = input()
    veces = ingreso
    ingresoAlpha = ingreso.isalpha()
    if veces == '0':
        print('###################################################')
        print('# No puede contener Letras Ni ser un Campo Vacio! #')
        print('###################################################'+'\n')
        ejecutarEnvios()
        
    if ingreso == 'l':
        veces = 99999999999
        veces = int(veces)
        for a in range(0,veces):
                hacer()

    if  ingresoAlpha == True:
        print('###################################################')
        print('# No puede contener Letras Ni ser un Campo Vacio! #')
        print('###################################################'+'\n')
        ejecutarEnvios()

    if ingreso >= '1':
        veces = ingreso
        veces = int(veces)
        for a in range(0,veces):
                hacer()
    else:
        print('###################################################')
        print('# No puede contener Letras Ni ser un Campo Vacio! #')
        print('###################################################'+'\n')
        ejecutarEnvios()
    
    ingreso = str(ingreso)
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
    opciones()
    

############################################################################

def hacer():
            #Abriendo Sistema Administrativo SA
            pyautogui.hotkey('win', 'r')
            time.sleep(1)
            pyautogui.typewrite(sa, interval=0.5)
            pyautogui.hotkey('enter',interval=0.5)#Abre el SA
            pyautogui.hotkey('tab')
            pyautogui.typewrite(PASS, interval=0.5)#Cargamos la PASS
            pyautogui.hotkey('enter',interval=1)#Entramos al SA
            time.sleep(3)
            pyautogui.moveTo(452, 35, duration=1)#Abre Guardia De Conexiones
            pyautogui.click(452, 35)
            pyautogui.moveTo(493, 212, duration=1)#Adelantos firmados
            pyautogui.click(493, 212)
            time.sleep(10)

            #Envio de Conexiones
            pyautogui.moveTo(1203, 126, duration=1)#Actualizar
            pyautogui.click(1203, 126)
            time.sleep(1)
            pyautogui.moveTo(1203, 245, duration=1)#Seleccionar TODO
            pyautogui.click(1203, 245)
            time.sleep(1)
            pyautogui.moveTo(1203, 354, duration=1)#Enviar Notas
            pyautogui.click(1203, 354)
            time.sleep(6)
            pyautogui.click(1203, 354)
            pyautogui.hotkey('enter', interval=0.3)#ACEPTAR "Se realizaron los Envios Automaticamente"

            #Creando Archivo con Fecha Actual
            fecha = time.strftime("%d-%m-%y")
            conexionesPacks = str(carpetaConexiones +'\\'+ dia +'.txt')
        
            packC = open(conexionesPacks, 'a')
            
            global ultimoEnvioConexion
            ultimoEnvioConexion = str( DiaYHoraDeEnvio )
            packC.write(DiaYHoraDeEnvio+ '\n')
            time.sleep(1)
            
            
            #Envio de Datos
            pyautogui.moveTo(1247, 176, duration=1)#Menu Desplegable
            pyautogui.click(1247, 176)
            pyautogui.moveTo(1213, 204, duration=1)#Seleccion de Info de Datos
            pyautogui.click(1213, 204)
            pyautogui.moveTo(1203, 245, duration=1)#Seleccionar TODO
            pyautogui.click(1203, 245)
            pyautogui.moveTo(1203, 354, duration=1)#Enviar Notas
            pyautogui.click(1203, 354)
            time.sleep(6)
            pyautogui.click(1203, 354)
            pyautogui.hotkey('enter', interval=0.3)#ACEPTAR "Se realizaron los Envios Automaticamente"

            #Creando Archivo con Fecha Actual
            fecha = time.strftime("%d-%m-%y")
            datosPacks = str(carpetaDatos +'\\'+ dia +'.txt')
            
            packD = open(datosPacks, 'a')
            
            global ultimoEnvioDatos
            ultimoEnvioDatos = str(DiaYHoraDeEnvio)
            packD.write(DiaYHoraDeEnvio+ '\n')
            time.sleep(1)

            #Envio de Reclamos
            pyautogui.moveTo(1247, 176, duration=1)#Menu Desplegable
            pyautogui.click(1247, 176)
            pyautogui.moveTo(1203, 216, duration=1)#Seleccion de Info de Datos
            pyautogui.click(1203, 216)
            pyautogui.moveTo(1203, 245, duration=1)#Seleccionar TODO
            pyautogui.click(1203, 245)
            pyautogui.moveTo(1203, 354, duration=1)#Enviar Notas
            pyautogui.click(1203, 354)
            time.sleep(6)
            pyautogui.click(1203, 354)
            pyautogui.hotkey('enter', interval=0.3)#ACEPTAR "Se realizaron los Envios Automaticamente"

            #Creando Archivo con Fecha Actual
            fecha = time.strftime("%d-%m-%y")
            reclamosPacks = str(carpetaReclamos +'\\'+ dia +'.txt')
            
            packR = open(reclamosPacks, 'a')
            
            global ultimoEnvioReclamos
            ultimoEnvioReclamos = str(DiaYHoraDeEnvio)
            packR.write(DiaYHoraDeEnvio+ '\n')
            time.sleep(1)

            pyautogui.hotkey('alt', 'f4')
            packC.close()
            packD.close()
            packR.close()
            ultimoEnvioConexion = str( DiaYHoraDeEnvio )

            opciones()


        
############################################################################
    
def devolucion():
    #consulta de archivo a devolver
    print('Nombre Pack a Devolver (Sin Extension TXT.)')
    carpetaCortes = str(CptUsuario + 'Cortes'+'\\')
    dvuelv = input()
    dvuelv = dvuelv.lower()
    aDevolver = str(carpetaCortes + dvuelv + '.txt')
    print('Desea Hacer Nuevamente la Entrega? (S/N)')
    conEntrega = input()
    conEntrega = conEntrega.lower()
    
    #Abriendo Sistema Administrativo SA
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.typewrite(sa, interval=0.5)
    pyautogui.hotkey('enter',interval=1)#Abre el SA
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(PASS, interval=0.5)#Cargamos la PASS
    pyautogui.hotkey('enter',interval=1)#Entramos al SA
    time.sleep(3)
    
    #Va a Entrega Interna
    archivo1 = open(aDevolver, 'r')
    linea = archivo1.readline()

    pyautogui.moveTo(67, 56, duration=0.5)
    pyautogui.click(67, 56)

    pyautogui.moveTo(124, 99, duration=0.5)
    pyautogui.click(124, 99)
    
    #Los devuelve
    while linea != '':
        linea = archivo1.readline()
        clipboard.copy(linea)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)
        pyautogui.moveTo(670, 700, duration=0.5)
        pyautogui.click(670, 700)
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)        
        pyautogui.moveTo(1093, 798, duration=0.5)
        pyautogui.click(1093, 798)
        time.sleep(1)

    pyautogui.hotkey('esc')
    pyautogui.hotkey('esc')

    archivo1.close()
    time.sleep(3)
    archivo1 = open(aDevolver, 'r')
    linea = archivo1.readline()

    #hace La Entrega
    if conEntrega == 's':

        pyautogui.moveTo(67, 56, duration=0.5)
        pyautogui.click(67, 56)

        pyautogui.moveTo(124, 99, duration=0.5)
        pyautogui.click(124, 99)

        while linea != '':
            linea = archivo1.readline()
            clipboard.copy(linea)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)

        pyautogui.moveTo(595, 821, duration=1)
        pyautogui.click(595, 821)
        time.sleep(3)
              
        pyautogui.moveTo(77, 42, duration=1)
        pyautogui.click(77, 42)      
        time.sleep(3)

        #Envia a Imprimir y espera 2"            
        pyautogui.hotkey('enter')
        time.sleep(20)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('alt', 'f4')
                    
        time.sleep(1)

        archivo1.close()
        time.sleep(3)
        opciones()

    else:
        pyautogui.hotkey('alt','tab')
        print('Se Devolvio sin Hacer Entrega!')
        opciones()
        
#########################################################################

def soloEtiquetas():
    #consulta de archivo a devolver
    print('Nombre Pack a Etiquetar (Sin Extension TXT.)')
    carpetaCortes = str(CptUsuario + 'Cortes'+'\\')
    dvuelv = input()
    dvuelv = dvuelv.lower()
    
    aDevolver = str(carpetaCortes + dvuelv + '.txt')
    print('Desea Hacer las Etiquetas? (S/N)')
    soloEtiqueta = input()
    soloEtiqueta = soloEtiqueta.lower()
    
    #Abriendo Sistema Administrativo SA
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.typewrite(sa, interval=0.5)
    pyautogui.hotkey('enter',interval=1)#Abre el SA
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(PASS, interval=1)#Cargamos la PASS
    pyautogui.hotkey('enter',interval=1)#Entramos al SA
    time.sleep(3)
    
    archivo1 = open(aDevolver, 'r')
    linea = archivo1.readline()
    
    #Hace etiquetas
    if soloEtiqueta == 's':

        pyautogui.moveTo(65, 57, duration=1)
        pyautogui.click(65, 57)

        pyautogui.moveTo(109, 76, duration=0.5)
        pyautogui.click(109, 76)
        time.sleep(2)

        while linea != '':
            linea = archivo1.readline()
            clipboard.copy(linea)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(3)
        opciones()
    else:
        opciones()
    archivo1.close()
    
#########################################################################
def soloEntrega():
    #consulta de archivo a devolver
    print('Nombre Pack a Entregar (Sin Extension TXT.)')
    carpetaCortes = str(CptUsuario + 'Cortes'+'\\')
    dvuelv = input()
    dvuelv = dvuelv.lower()
    
    aDevolver = str(carpetaCortes + dvuelv + '.txt')
    
    print('Hacer la Entrega? (S/N)')
    hacerEntrega = input()
    hacerEntrega = hacerEntrega.lower()
    
    #Abriendo Sistema Administrativo SA
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.typewrite(sa, interval=0.5)
    pyautogui.hotkey('enter',interval=1)#Abre el SA
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(PASS, interval=1)#Cargamos la PASS
    pyautogui.hotkey('enter',interval=1)#Entramos al SA
    time.sleep(3)
    
    
    archivo1 = open(aDevolver, 'r')
    linea = archivo1.readline()
    
    #Hace la Entrega Entrega
    if hacerEntrega == 's':

        #Va a Entrega Interna
        pyautogui.moveTo(67, 56, duration=0.5)
        pyautogui.click(67, 56)

        pyautogui.moveTo(124, 99, duration=0.5)
        pyautogui.click(124, 99)

        while linea != '':
            linea = archivo1.readline()
            clipboard.copy(linea)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)

        pyautogui.moveTo(595, 821, duration=1)
        pyautogui.click(595, 821)
        time.sleep(3)
              
        pyautogui.moveTo(77, 42, duration=1)
        pyautogui.click(77, 42)      
        time.sleep(3)

        #Envia a Imprimir y espera 2"            
        pyautogui.hotkey('enter')
        time.sleep(20)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('alt', 'f4')
                    
        time.sleep(1)

        archivo1.close()
        time.sleep(3)
        opciones()

    else:
        pyautogui.hotkey('alt','tab')
        print('Se Devolvio sin Hacer Entrega!')
        opciones()
        
#########################################################################

def soloDevolucion():
    #consulta de archivo a devolver
    print('Nombre Pack a Devolver (Sin Extension TXT.)')
    carpetaCortes = str(CptUsuario + 'Cortes'+'\\')
    dvuelv = input()
    dvuelv = dvuelv.lower()
    aDevolver = str(carpetaCortes + dvuelv + '.txt')
    print('Desea Hacer la Devolucion? (S/N)')
    soloDevolucion = input()
    soloDevolucion = soloDevolucion.lower()
    
    #Abriendo Sistema Administrativo SA
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.typewrite(sa, interval=0.5)
    pyautogui.hotkey('enter',interval=1)#Abre el SA
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(PASS, interval=1)#Cargamos la PASS
    pyautogui.hotkey('enter',interval=1)#Entramos al SA
    time.sleep(3)
    
    #Va a Entrega Interna
    archivo1 = open(aDevolver, 'r')
    linea = archivo1.readline()

    pyautogui.moveTo(67, 56, duration=0.5)
    pyautogui.click(67, 56)

    pyautogui.moveTo(124, 99, duration=0.5)
    pyautogui.click(124, 99)

    #Los devuelve
    while linea != '':
        linea = archivo1.readline()
        clipboard.copy(linea)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)
        pyautogui.moveTo(670, 700, duration=0.5)
        pyautogui.click(670, 700)
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)        
        pyautogui.moveTo(1093, 798, duration=0.5)
        pyautogui.click(1093, 798)
        time.sleep(1)

    pyautogui.hotkey('esc')
    pyautogui.hotkey('esc')
    pyautogui.hotkey('alt', 'f4')

    archivo1.close()
 
    pyautogui.hotkey('alt','tab')
    print('Se Devolvio sin Hacer Entrega!')
    opciones()
     
########################################################################

nombreLogCC = hora + '-' + minutos + '-' +segundos + '-' + dia

numero = int(0)

def entrega():
    while 1==1:

        def inc():
            global numero
            numero += 1
            return numero
        
        nDeCarpeta = inc()
        
        #Creando Archivo con Fecha Actual
        fecha = time.strftime("%d-%m-%y")
        xc = str(numero)
        logPacks = str(carpetaCortes + '\\' + nombreLogCC + '- Pack Nº ' + xc + '.txt')
        pack = open(logPacks, 'a')
        base = str(00)
        pack.write(base + '\n')

        #carga de cds
        cds = []
        numd=0
                      
        while len(cds) <= 49:
            print("Ingrese Numero de CD:" + '    Total: ' + str(len(cds)))
            nCD = input()
            if nCD not in cds:
                if nCD != '' and nCD != ' ' and nCD != '0':
                    cds.insert(0,nCD)

                else:
                    print('\n' +
                           '############################################################'+'\n' +
                           "Numero de CD Invalido" + '\n' +
                           '############################################################'+'\n' )
            else:
                reverseCDs = cds.copy()
                reverseCDs.reverse()
                print('\n' +
                      '############################################################'+'\n' +
                      'Numero de CD ya Ingresado en Posicion '+ str(reverseCDs.index(str(nCD))+1)+ '\n'+
                      '############################################################'+'\n' )
            
        #carga la lista en el TXT de Log
        for a in range(1,50):
            numd
            pack.write(cds[numd]+'\n')
            numd = numd+1
        pack.close()
        

        print("50 CDs Listos para Etiquetar")
        print('Desea Imprimir las etiquetas S/N')
        deseaImprimir = input()
        deseaImprimir = deseaImprimir.lower()
        
        while deseaImprimir != "s":
            print("Ultima etiqueta " + cds[0] + "\n" + "ya alcanzo 50 CDs! desea Imprimir las etiquetas S/N" + "\n" )
            deseaImprimir = input()
            deseaImprimir = deseaImprimir.lower()
            if deseaImprimir == 'n':
                opciones()

        #Abriendo Sistema Administrativo SA

        pyautogui.hotkey('win', 'r')
        time.sleep(1)
        pyautogui.typewrite(sa, interval=0.5)
        pyautogui.hotkey('enter',interval=1)#Abre el SA
        time.sleep(1)
        pyautogui.hotkey('tab')
        pyautogui.typewrite(PASS, interval=1)#Cargamos la PASS
        pyautogui.hotkey('enter',interval=1)#Entramos al SA
        time.sleep(3)

        pyautogui.moveTo(65, 57, duration=1)
        pyautogui.click(65, 57)

        pyautogui.moveTo(109, 76, duration=0.5)
        pyautogui.click(109, 76)
        time.sleep(2)

        #Imprimiemdo Etiquetas
        archivo1 = open(logPacks, 'r')
        linea = archivo1.readline()

        while linea != '':
            linea = archivo1.readline()
            clipboard.copy(linea)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)
            if prendeScan == 1:
                scanScr()
            else:
                time.sleep(1)
        
        archivo1.close()
        time.sleep(3)
        
        pyautogui.hotkey('esc')    
        time.sleep(1)
        pyautogui.hotkey('esc')    
        time.sleep(1)

        pyautogui.moveTo(65, 57, duration=1)
        pyautogui.click(65, 57)

        pyautogui.moveTo(109, 103, duration=0.5)
        pyautogui.click(109, 103)
        time.sleep(3)

        #Entrega Interna
        archivo1 = open(logPacks, 'r')
        linea = archivo1.readline()
        
        while linea != '':
            linea = archivo1.readline()
            clipboard.copy(linea)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)

        archivo1.close()
        time.sleep(3)
        
        pyautogui.moveTo(595, 821, duration=1)
        pyautogui.click(595, 821)
        time.sleep(3)
        
        pyautogui.moveTo(77, 42, duration=1)
        pyautogui.click(77, 42)      
        time.sleep(3)
        
        pyautogui.hotkey('enter')
        time.sleep(20)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('esc')
        pyautogui.hotkey('alt', 'f4')
        
        time.sleep(1)
        opciones()

########################################################################

nombreLogCC = hora + '-' + minutos + '-' +segundos + '-' + dia

numero = int(0)

#OPCION 6
#SOLO PACKS

def soloPacks():
    while 1==1:

        def inc():
            global numero
            numero += 1
            return numero
        
        nDeCarpeta = inc()
        
        #Creando Archivo con Fecha Actual
        fecha = time.strftime("%d-%m-%y")
        xc = str(numero)
        logPacks = str( carpetaCortes +'\\'+ nombreLogCC + '- Pack Nº ' + xc + '.txt')

        pack = open(logPacks, 'a')
        base = str(00)
        pack.write(base + '\n')

        #carga de cds
        cds = []
        numd=0
        
        while len(cds) <= 49:
            print("Ingrese Numero de CD:" + '    Total: ' + str(len(cds)))
            nCD = input()
            if nCD not in cds:
                if nCD != '' and nCD != ' ' and nCD != '0':
                    cds.insert(0,nCD)

                else:
                    print('\n' +
                           '############################################################'+'\n' +
                           "Numero de CD Invalido" + '\n' +
                           '############################################################'+'\n' )
            else:
                reverseCDs = cds.copy()
                reverseCDs.reverse()
                print('\n' +
                      '############################################################'+'\n' +
                      'Numero de CD ya Ingresado en Posicion '+ str(reverseCDs.index(str(nCD))+1)+ '\n'+
                      '############################################################'+'\n' )
            
        #carga la lista en el TXT de Log
        for a in range(1,51):
            pack = open(logPacks, 'a')
            numd
            pack.write(cds[numd]+'\n')
            numd = numd+1
            pack.close()
        pack.close()

        print("50 CDs Listos para Etiquetar"+
              'Press Enter Para Terminar'+'\n' )
        deseaImprimir = input()
        deseaImprimir = deseaImprimir.lower()
        
        while deseaImprimir != "":
            print("Ultima etiqueta " + cds[0] + "\n" +
                  "ya alcanzo los 50 CDs!" + "\n" +
                  "desea Imprimir las etiquetas S/N" + "\n" +
                  'Press Enter Para Terminar'+'\n')
            deseaImprimir = input()
            deseaImprimir = deseaImprimir.lower()
            if deseaImprimir == 'n':
                opciones()

        time.sleep(1)
        opciones()

#############################################################################


eleccion = 0

def opciones():
    print("\n"+ "\n"+ "\n"+ "\n"+ "\n"+ "\n"+ "\n"+ "\n")
    print('Opciones: --->>> Selecciones(1,2,3, etc)')
    print('1- Entrega'+'\n'
          '2- Solo Etiquetas'+'                     Hora Ultimos Envios:'+'\n'
          '3- Devolucion + Entrega'+'                  Conexiones:'+'    '+ultimoEnvioConexion+'\n'
          '4- Solo Entregas'+'                          Datos:'+'        '+ultimoEnvioDatos+'\n'
          '5- Solo Devolucion'+'                         Reclmos:'+'     '+ultimoEnvioReclamos+'\n'
          '6- Solo ARMADO de Packs'+'\n'
          '7- Envio Automatico (Notas)'+'\n'
          '8- Scan (on/off)'+'\n'
          'x- Exit'+'\n')
    print('Eleccion:')

    eleccion = input()
    
    if eleccion != '1' and eleccion != '2' and eleccion != '3' and eleccion != '4' and eleccion != '5' and eleccion != '6'and eleccion != '7'and eleccion != '8' and eleccion != 'x':
        print('\n' +
              '############################################################' + '\n' +
              "Opcion Inexistente" + '\n' +
              '############################################################' + '\n' )
        opciones()

    if eleccion == '1':
        entrega()
        
    if eleccion == '2':
        soloEtiquetas()
        
    if eleccion == '3':
        devolucion()
        
    if eleccion == '4':
        soloEntrega()
        
    if eleccion == '5':
        soloDevolucion()

    if eleccion == '6':
        soloPacks()

    if eleccion == '7':
        ejecutarEnvios()

    if eleccion == '8':
        hScan()
         
    if eleccion == 'x':
        sys.exit()
    
    else:
        entregaCds = 1


    
########################################################################

opciones()

#########################################################################
