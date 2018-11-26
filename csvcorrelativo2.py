import xlrd
import Tkinter, tkFileDialog, re
import sys  #esta y las 2 lineas siguientes permiten que trabaje con n con gorrito y esas letras
reload(sys)
sys.setdefaultencoding('utf-8')

#Permite seleccionar el archivo
#_________________________________________________________________________
root = Tkinter.Tk() #esto se hace solo para eliminar la ventanita de Tkinter 
root.withdraw() #ahora se cierra 
file_path = tkFileDialog.askopenfilename() #abre el explorador de archivos y guarda la seleccion en la variable

#Almacena la informacion de excel
#___________________________________________________________________________
doc = xlrd.open_workbook(str(file_path))
H=doc.sheet_names()#entrega el nombre de las hojas del libro
#FPS=H.index("formato para subir")
hoja = doc.sheet_by_name("formato para subir")
#print hoja.nrows

cabecera=len(hoja.row_values(0))
print ""
print cabecera, "columnas analizadas"

listadefilas = [] #Una lista con todas las filas
for j in range ( 0, hoja.nrows):
      listadefilas.append(hoja.row(j))    
#print "largo lista de filas", len(listadefilas)

filasutiles=[]
for q in range(0,len(listadefilas)):
  #if type(listadefilas[q][0].value)==float:
  if len(listadefilas[q][0].value)>1:
    filasutiles.append(listadefilas[q])
#print "Cargos a analizar:", len(filasutiles)
#print "largo elementos", len(listadefilas[0][3].value)
#print filasutiles[0][-1] #ultimo valor primera fila

#############
#CAMBIA TODAS LOS ELEMENTOS DEL EXCEL EN TEXTO Y LUEGO PARA SABER LAS FILAS UTILES HAZLO DE ACUERDO A LA CANTIDAD DE ELEMENTOS, NO AL TIPO.
##############

if len(filasutiles[0])==29:
  print "Cantidad de columnas correctas"
else:
  print "La cantidad de columnas no es la correcta"

Correlativos=[]

for e in range(0,len(filasutiles)):#con todo esto se cambia el valor de cada elemento convirtiendolo en string,
                                   #luego se une cada elemento de las listas para ser un solo string por lista
      t=0
      while t<cabecera:
         filasutiles[e][t]=str(filasutiles[e][t].value)#convierte cada elemento en un string
         t+=1
         if t==cabecera:#se detiene cuando llega a la cantidad de columnas y ahora si se agrega la lista
            Correlativos.append(";".join(filasutiles[e]))#cada elemento de las listas las une como string (por lista)
print "Cargos a crear:", len(Correlativos)

     #
        #print str1
        #Correlativos.append(str1)

rep=[] #lista con los correlativos   #FALTA CODIGOS DE CARGO
for i in range(0, len(Correlativos)):
  rep.append(filasutiles[i][0])

if len(list(set(rep)))!=len(Correlativos): #compara la lista de correlativos con una lista que solo considera numeros unicos
  print "__________________________________________________________________________________"
  print ""
  print "HAY CORRELATIVOS REPETIDOS, REVISAR"
  print ""
  print "___________________________________________________________________________________"

for i in range(0, len(Correlativos)):
  if filasutiles[i][1]=="2" and filasutiles[i][9]=="0": #Si la ley es 18834 y el grado es 0, lanza una alerta
    print "__________________________________________________________________________________"
    print ""
    print "HAY GRADOS EN 0, REVISAR"
    print ""
    print "___________________________________________________________________________________"
    break
for i in range(0, len(Correlativos)):
  if filasutiles[i][1]=="2" and filasutiles[i][10]=="0": #Si la ley es 18834 y las horas son 0, lanza una alerta
    print "__________________________________________________________________________________"
    print ""
    print "FALTA ASIGNAR HORAS EN 18834, REVISAR"
    print ""
    print "___________________________________________________________________________________"
    break

#ADVERTENCIAS
#_____________________________________________________________      
fout = open ("Correlativos.csvmsdos","w")  #
for ii  in range (0, len (Correlativos)):
    fout.write (Correlativos[ii]+ "\n") # + "\r\n" 
fout.close()

fout = open ("Correlativos.txt","w")
for ii  in range (0, len (Correlativos)):
    fout.write (Correlativos[ii]+ "\n") # + "\r\n" 
fout.close()
print""
print "Archivo csvmsdos creado"
print""
print "Desarrollado en Python 2.7.11"

raw_input()