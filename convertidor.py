import xlsxwriter
from os import scandir

def main():
    #valores iniciales
    file_output="output/reporte_estadisticas.xlsx";
    dir_input="input/"

    #creacion del archivo
    libro = xlsxwriter.Workbook(file_output,{'strings_to_urls': False})
    hoja = libro.add_worksheet()
    hoja.freeze_panes(1,0) #Freeze the first row
    
    #f=open("archivo_iica.log","r")
    cell_format = libro.add_format()

    row = 1
    crear_header_excel(hoja,cell_format)
    archivos=leer_input(dir_input)

    for archivo in archivos:
        try:
            ruta_archivo=dir_input+archivo
            row=insertarDatosReporte(ruta_archivo,hoja,row)

        except:
            print("An exception occurred inserted")

    ultima='I'+str(row)
    hoja.autofilter('A1:'+ultima)

    #Cerramos el libro
    libro.close()

def leer_input(path):
    return [obj.name for obj in scandir(path) if obj.is_file()]

def insertarDatosReporte(archivo,hoja,row):
    f=open(archivo,"r")
    for x in f:
        try:
            cols = x.split(" ")
            Ip=cols[0]
            Sesion=cols[2]

            fecha_completa=limpiar_fecha(cols[3])

            Metodo = limpiar_metodo(cols[5])
            Fecha_dia=fecha_completa[0]
            Fecha_hora=fecha_completa[1]
            Accion=cols[6]
            Url=limpiar_url_accion(Accion)
            Estatus_peticion=cols[8]
            Tamano=limpiar_tamanio(cols[9])

            hoja.write(row, 0,Ip)
            hoja.write(row, 1,Sesion)
            hoja.write(row, 2,Fecha_dia)
            hoja.write(row, 3,Fecha_hora)
            hoja.write(row, 4,Metodo)
            hoja.write(row, 5,Url)
            hoja.write(row, 6,Estatus_peticion)
            hoja.write(row, 7,Tamano)
            hoja.write(row, 8,Accion)
            
            row += 1
        except:
            print("An exception occurred")
    
    return row

def limpiar_fecha(fecha):
    fecha_completa=fecha.replace('[','').split(":",1)
    return fecha_completa

def limpiar_metodo(metodo):
    metodo_limpio=metodo.replace('"','').replace('&#34;','')
    return metodo_limpio

def limpiar_tamanio(tam):
    tam_limpio=tam.replace('\n','')
    return tam_limpio

def limpiar_url_accion(accion):
    url=accion.split("/",3);
    return url[2]

def crear_header_excel(hoja,estilo):
    estilo.set_bold(True)  # Also turns bold on.
    estilo.set_font_size(12)
    estilo.set_align('center')
    row=0
    hoja.write(row, 0,"IP",estilo)
    hoja.write(row, 1,"Sesion",estilo)
    hoja.write(row, 2,"Día",estilo)
    hoja.write(row, 3,"Hora",estilo)
    hoja.write(row, 4,"Metodo",estilo)
    hoja.write(row, 5,"Base",estilo)
    hoja.write(row, 6,"Estatus peticion",estilo)
    hoja.write(row, 7,"Tamaño",estilo)
    hoja.write(row, 8,"Url",estilo)

if __name__=="__main__":
    main()