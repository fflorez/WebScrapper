from urllib import urlopen
from bs4 import BeautifulSoup
from xlrd import open_workbook
from xlutils.copy import copy
import re 

#Codigo Paises
paises_codigo = {'Argentina': 'AR', 'Bahamas': 'BH', 'Barbados':'BA','Belize':'BL','Bolivia': 'BO','Brazil':'Br','Chile':'CH', 'Colombia':'CO','Costa Rica': 'CR', 'Dominican Republic': 'DR', 'Ecuador':'EC', 'El Salvador':'ES', 'Guatemala':'GU', 'Guyana':'GY','Haiti':'HA', 'Honduras': 'HO','Jamaica':'JA', 'Mexico':'ME','Nicaragua':'NI', 'Panama':'PN', 'Paraguay':'PR', 'Peru':'PE','Suriname':'SU', 'Trinidad':'TT', 'Uruguay':'UR','Venezuela':'VE'}

#Imprime el pais y su codigo
for i in paises_codigo:
    print "Pais : " + i +' Codigo: '+ paises_codigo[i] 


#Pedir pais
codigo = raw_input('Escribir el codigo de Pais: ')


#Escribe la informacion en la hoja approved
def approved_excel(row, info, country):
    rb = open_workbook(country+'.xls')
    wb = copy(rb)
    ws = wb.get_sheet(1)
    ls = range(9)
    for i in ls:
        ws.write(row,i,info[i])
    wb.save(country+'.xls')
                  
def approved():
    row = 1     #row in the excel document
    counter = 1
    #visitar las 6 hojas de proyectos en preparacion, si es Brasil son 9 hojas
    repeats = 6
    if codigo == "BR" or codigo == "br":
        repeats = 9
    while counter <= repeats:
        n = str(counter)        #numero de pagina        
        preparation = "http://www.iadb.org/en/projects/advanced-project-search,1301.html?query=&adv=true&Country="+codigo+"&tab=2&pagePIP=1&pageAPP=1&order=asc&sort=country&page="+n
        counter = counter + 1
        #abrir la pagina y leer el html
        webpage = urlopen(preparation).read()
        #crear objeto de beautifulsoup
        soup = BeautifulSoup(webpage)
        tables = soup.find_all('table') #encontrar la tabla con informacion
        if len(tables) == 2:
            table = tables[1]
        else:
            table = tables[0]
        table = str(table)
        table = table.split("</tr>")
        table.pop(0)
        table.pop(-1)
        #tabla de todos los proyectos separada por filas. Para cada fila 
        for p in table:
            #convertir cada fila en string para mmanipular informacion
            project = str(p)
            #compilar regular expression para encontrar link al proyecto
            linkPattern = re.compile('<a href=\\"(.*)">(.*)</a>')
            result = re.findall(linkPattern, project)
            prelink = result[0][0]
            #link al proyecto actual
            projectlink = "http://www.iadb.org"+prelink
            #convertir el nombre del proyecto q se obtiene del link a unicode
            #por los caracteres de espanol y portuges
            name = unicode(result[0][1])
            name = name.decode('utf-8')
            name = unicode(name)
            #abrir el link q se encontro para sacar la informacion del proyecto
            projectwebpage = urlopen(projectlink).read()
            soup2 = BeautifulSoup(projectwebpage)
            #encontrar tabla que dice overview para extraer informacion
            overview= soup2.find("div", {"id": "projectListtabContent-1"})
            initial_info = overview.find_all("div") #crear array con info de overview
            #convertir numero de proyecto y estatus a string
            number = initial_info[5].text
            status = initial_info[11].text
            #encontrar toda la informacion del proyecto, crear array separando cada pedazo de informacion
            messy_info = soup2.find_all('td')
            #sacar solo el texto de cada pedazo de informacion
            clean_info = []
            for i in messy_info:
                clean_info.append(i.text)
            #encontrar el pais y el secor del proyecto en clean info
            country = clean_info[0]
            sector = clean_info[1]
            #encontrar la fecha cuando se aprobo el proyecto
            if "Approval Date" in clean_info:
                if clean_info.index("Approval Date") == 19:
                    date = clean_info[20]
                elif clean_info.index("Approval Date") == 10:
                    date = "n/a"
            #encontrar el costo historico del proyecto
            cost = "n/a"
            if "Total Cost - Historic" in clean_info:
                cost = clean_info[clean_info.index("Total Cost - Historic")+1]
            elif "Estimated Total Cost" in clean_info:
                cost = clean_info[clean_info.index("Estimated Total Cost")+1]
            #encontrar la cantidad de dinero undisbursed del proyecto
            undisbursed = "n/a"
            if "Undisbursed Amount - Historic" in clean_info:
                undisbursed = clean_info[clean_info.index("Undisbursed Amount - Historic")+1]
            #encontrar idb leader
            leader = "n/a"
            if "IDB Team Leader" in clean_info:
                leader = clean_info[clean_info.index("IDB Team Leader")+1]
           
            if status != "Completed":
                info=[]
                info.append(country)
                info.append(unicode(name))
                info.append("=HYPERLINK(\""+projectlink+"\", \""+number+"\")")
                info.append(sector)
                info.append(status)
                info.append(date)
                info.append(cost)
                info.append(undisbursed)
                info.append(leader)
                approved_excel(row, info, country)
                row = row+1

def preparation_excel(row, info, country):
    rb = open_workbook(country+'.xls')
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ls = range(6)
    for i in ls:
        ws.write(row,i,info[i])
    wb.save(country+'.xls')

def preparation():
    try:
        row = 1         #fila en documento de excel
        counter = 1     #pagina del proyecto 
        while counter <= 5:
            n = str(counter)        
            #pagina de proyectos en preparacion
            preparation = "http://www.iadb.org/en/projects/advanced-project-search,1301.html?query=&adv=true&Country="+codigo+"&tab=1&pagePIP=1&pageAPP=3&order=asc&sort=country&page=" + n
            counter = counter + 1
            #abrir pagina y crear beautifulsoup object
            webpage = urlopen(preparation).read()
            soup = BeautifulSoup(webpage)
            #encontrar todas las tablas y usar la primera tabla q se encuentra
            tables = soup.find_all('table')
            table = tables[0]
            table = str(table)
            #separar las filas de la tabla y limpiar la tabla
            table = table.split("</tr>") 
            table.pop(0)
            table.pop(-1)
            #para cada proyecto en la tabla
            for p in table:
                project = str(p)
                #compilar la regular expression para encontrar link
                linkPattern = re.compile('<a href=\\"(.*)">(.*)</a>')
                result = re.findall(linkPattern, project)
                prelink = result[0][0]      #link del proyecto
                projectlink = "http://www.iadb.org"+prelink
                #nombre del proyecto, encode a unicode por caracteres 
                #espanol y portuges
                name = unicode(result[0][1])
                name = name.decode('utf-8')
                name = unicode(name)
                #leer la informacion de la pagina del proyecto
                projectwebpage = urlopen(projectlink).read()
                soup2 = BeautifulSoup(projectwebpage)
                #encontrar tablas del proyecto
                table2= soup2.find_all('table')
                table2 = str(table2)
                table2 = table2.split("</tr>") #separar las filas de la tabla 
                table2.pop(0)
                table2.pop(-1)
                #compilar las regular expressions para cada pedazo de informacion
                countrypattern = re.compile('<tr valign="top"> <td colspan="2" width="270">Country</td> <td width="260">(.*)</td>')
                projectnumberpattern = re.compile('<tr valign="top"> <td colspan="2" width="270">Project Number</td> <td width="260">(.*)</td>')
                statuspattern = re.compile('<tr> <td colspan="2" height="12">Project Status</td> <td>(.*)</td>')
                costpattern = re.compile('<tr valign="top"> <td colspan="2" width="270">Estimated Total Cost</td> <td width="260">(.*)</td>')
                sectorpattern = re.compile('<tr valign="top"> <td colspan="2" width="270">Sector</td> <td width="260">(.*)</td> ')
                #pais
                country = re.findall(countrypattern,table2[1])
                country = country[0]
                #numero de proyecto
                number = re.findall(projectnumberpattern, table2[0])
                number = number[0]
                #sector
                sector = re.findall(sectorpattern, table2[2])
                sector = sector[0]
                #status
                status = re.findall(statuspattern, table2[7])
                status = status[0]
                #costo estimado
                cost = re.findall(costpattern, table2[9])
                cost = cost[0]
                info=[]
                info.append(country)
                info.append(unicode(name))
                info.append("=HYPERLINK(\""+projectlink+"\", \""+number+"\")")
                #info.append(number)
                info.append(status)
                info.append(cost)
                info.append(sector)
                preparation_excel(row,info,country)
                row = row +1
    except Exception, e:
        print e


preparation()
approved()
