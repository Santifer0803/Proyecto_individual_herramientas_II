# -*- coding: utf-8 -*-
"""
Created on Sun Jun  9 20:34:15 2024

@author: Santiago
"""

import xlsxwriter
import pandas as pd

class EditorExcel:
    
    # Constructor
    def __init__(self, df, nombre):
        self.__escritor = pd.ExcelWriter(nombre, engine = 'xlsxwriter')
        
        # Se convierte el DataFrame a un
        df.to_excel(self.__escritor, sheet_name = 'Hoja1', index = False)
        
        # Get the xlsxwriter objects from the dataframe writer object.
        self.__libro  = self.__escritor.book
        self.__hoja = self.__escritor.sheets['Hoja1']
    
    # Get y set
    @property
    def escritor(self):
        return self.__escritor
    
    @property
    def libro(self):
        return self.__libro
    
    @property
    def hoja(self):
        return self.__hoja
    
    @escritor.setter
    def escritor(self, nuevo_escritor):
        self.__escritor = nuevo_escritor
    
    @libro.setter
    def libro(self, nuevo_libro):
        self.__libro = nuevo_libro
    
    @hoja.setter
    def hoja(self, nuevo_hoja):
        self.__hoja = nuevo_hoja
        
    # str
    def __str__(self):
        ''' Texto explicando un resumen de la clase EditorExcel
            
        Parameters
        ---------
        None
        
        Returns
        -----
        Cadena : str
            Texto explicativo que resume la clase EditorExcel
        '''
        return print('Una clase que transforma un DataFrame de pandas a Excel y le hace las modificaciones que el usuario elija de los métodos disponibles')
    
    # Métodos
    def cerrar(self):
        ''' Método que cierra el editor, es necesario usarlo siempre que se termina de editar un archivo, de lo contrario el Excel no podrá ser abierto
            
        Parameters
        ---------
        None
        
        Returns
        -----
        cierre : NoneType
            Función que cierra el editor de archivos xlsx de python para que se pueda abrir el archivo como un Excel común
        '''
        return self.__escritor.close()
    
    def escribir(self, texto, formato_escritura = 'python', fila = 0, columna = 0, celda = 'A1'):
        ''' Método que permite escribir texto, números o fórmulas básica en una celda de excel
            
        Parameters
        ---------
        texto : str, int o float
            Lo que será escrito en la celda, acepta operaciones matemáticas como sumas, multiplicaciones y otras, siempre y cuando no lleven funciones propias
            de excel
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel (ej. C2), si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0 (ej. fila = 0, columna = 0 => 'A1')
        fila : int
            Fila en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        columna : int
            Columna en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        celda : str
            Celda de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1' por defecto
        
        Returns
        -----
        mensaje : str
            Devuelve un mensaje en caso de que el formato de escritura no sea válido
        '''
        if(formato_escritura == 'excel'):
            self.__hoja.write(celda, texto)
        elif(formato_escritura == 'python'):
            self.__hoja.write(fila, columna, texto)
        else:
            return print('El formato de escritura no es válido')
    
    def formula(self, texto, formato_escritura = 'python', fila_ini = 0, columna_ini = 0, fila_fin = 0, columna_fin = 0, celda = 'A1:A1'):
        ''' Método que permite escribir fórmulas de Excel en una o varias casillas
            
        Parameters
        ---------
        texto : str
            Fórmula de Excel iniciando en =, en inglés, con las casillas a evaluar
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel, si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0
        fila_ini : int
            Fila en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        columna_ini : int
            Columna en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        fila_fin : int
            Fila en la que termina la fórmula si se elige el formato de escritura 'python', 0 por defecto
        columna_fin : int
            Columna en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        celda : str
            Rango de celdas de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1:A1' por defecto
        
        Returns
        -----
        mensaje : str
            Devuelve un mensaje en caso de que el formato de escritura no sea válido
        '''
        if(formato_escritura == 'excel'):
            self.__hoja.write_dynamic_array_formula(celda, texto)
        elif(formato_escritura == 'python'):
            self.__hoja.write_dynamic_array_formula(fila_ini, columna_ini, fila_fin, columna_fin, texto)
        else:
            return print('El formato de escritura no es válido')
    
    def escribir_fila(self, lista, formato_escritura = 'python', fila = 0, columna = 0, celda = 'A1'):
        ''' Método que permite poner una lista de python, en una fila de Excel
            
        Parameters
        ---------
        lista : list
            Lista de datos que será escrita en cierta fila, se llena hasta que se acabe la lista
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel (ej. C2), si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0 (ej. fila = 0, columna = 0 => 'A1')
        fila : int
            Fila en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        columna : int
            Columna en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        celda : str
            Celda de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1' por defecto
        
        Returns
        -----
        mensaje : str
            Devuelve un mensaje en caso de que el formato de escritura no sea válido
        '''
        if(formato_escritura == 'excel'):
            self.__hoja.write_row(celda, lista)
        elif(formato_escritura == 'python'):
            self.__hoja.write_row(fila, columna, lista)
        else:
            return print('El formato de escritura no es válido')
    
    def escribir_columna(self, lista, formato_escritura = 'python', fila = 0, columna = 0, celda = 'A1'):
        ''' Método que permite poner una lista de python, en una columna de Excel
            
        Parameters
        ---------
        lista : list
            Lista de datos que será escrita en cierta columna, se llena hasta que se acabe la lista
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel (ej. C2), si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0 (ej. fila = 0, columna = 0 => 'A1')
        fila : int
            Fila en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        columna : int
            Columna en la que se escribe el texto si se elige el formato de escritura 'python', 0 por defecto
        celda : str
            Celda de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1' por defecto
        
        Returns
        -----
        mensaje : str
            Devuelve un mensaje en caso de que el formato de escritura no sea válido
        '''
        if(formato_escritura == 'excel'):
            self.__hoja.write_column(celda, lista)
        elif(formato_escritura == 'python'):
            self.__hoja.write_column(fila, columna, lista)
        else:
            return print('El formato de escritura no es válido')
    
    def autoajuste(self):
        ''' Método que permite autoajustar todas las columnas que tengan algo escrito en el documento
            
        Parameters
        ---------
        None
        
        Returns
        -----
        autoajuste : NoneType
            Función que autoajusta las columnas del archivo de Excel
        '''
        self.__hoja.autofit()
    
    def condicional(self, condicion, valor, formato_escritura = 'python', fila_ini = 0, columna_ini = 0, fila_fin = 0, columna_fin = 0, celdas = 'A1:A1'):
        ''' Método que permite poner formato condicional en un rango de casillas de Excel
            
        Parameters
        ---------
        condicion : str
            Condición a evaluar el condicional, puede ser '<', '>', '<=', '>=' y similares
        valor : float
            Valor con el cual se aplicará la condición
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel, si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0
        fila_ini : int
            Fila en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        columna_ini : int
            Columna en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        fila_fin : int
            Fila en la que termina la fórmula si se elige el formato de escritura 'python', 0 por defecto
        columna_fin : int
            Columna en la que inicia la fórmula si se elige el formato de escritura 'python', 0 por defecto
        celdas : str
            Rango de celdas de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1:A1' por defecto
        
        Returns
        -----
        mensaje : str
            Devuelve un mensaje en caso de que el formato de escritura no sea válido
        '''
        # Color de fondo rojo
        formato = self.__libro.add_format({'bg_color': '#FFC7CE'})
        
        # Casos para elejir entre Excel y python
        if(formato_escritura == 'excel'):
            self.__hoja.conditional_format(celdas, {'type': 'cell', 
                                                    'criteria': condicion, 
                                                    'value': valor, 
                                                    'format': formato})
        elif(formato_escritura == 'python'):
            self.__hoja.conditional_format(fila_ini, columna_ini, fila_fin, columna_fin, {'type': 'cell', 
                                                                                          'criteria': condicion, 
                                                                                          'value': valor,
                                                                                          'format' : formato})
        else:
            return print('El formato de escritura no es válido')
    
    def mini_barras(self, celda_grafico, celda_datos):
        ''' Método que permite crear un gráfico de barras pequeño (sparklines)
            
        Parameters
        ---------
        celda_grafico : str
            Celda de Excel en donde se pondrá el gráfico
        celda_datos : str
            Celda(s) de Excel de donde se toman los datos del gráfico
        
        Returns
        -----
        None
        '''
        self.__hoja.add_sparkline(celda_grafico, {"range": ("Hoja1!" + celda_datos), "type": "column", "style": 10})
    
    def filtro(self, formato_escritura = 'python', fila_ini = 0, columna_ini = 0, fila_fin = 0, columna_fin = 0, celdas = 'A1:B1'):
        ''' Método que agrega el autofiltro a una selección de celdas de Excel
            
        Parameters
        ---------
        formato_escritura : str
            Indica el formato en el que se desea usar la función, si es 'Excel' el parámetro celda se rellena con una celda de Excel, si es python
            (por defecto), se escribe la fila y la columna con numeración de python, es decir, solo con números iniciando en 0
        fila_ini : int
            Fila en la que inicia el filtro (header) si se elige el formato de escritura 'python', 0 por defecto
        columna_ini : int
            Columna en la que inicia el filtro (header) si se elige el formato de escritura 'python', 0 por defecto
        fila_fin : int
            Fila en la que termina el filtro si se elige el formato de escritura 'python', 0 por defecto
        columna_fin : int
            Columna en la que inicia el filtro si se elige el formato de escritura 'python', 0 por defecto
        celdas : str
            Rango de celdas de Excel en donde se escribe el texto en caso de elegir el formato de escritura 'Excel', 'A1:B1' por defecto
        
        Returns
        -----
        None
        '''
        if(formato_escritura == 'excel'):
            self.__hoja.autofilter(celdas)
        elif(formato_escritura == 'python'):
            self.__hoja.autofilter(fila_ini, columna_ini, fila_fin, columna_fin)
        else:
            return print('El formato de escritura no es válido')
    
    def proteger(self, contraseña, formato = True, eliminar_filas = True, eliminar_columnas = True):
        ''' Método que proteje la hoja, permite elegir si el usuario puede alterar ciertos objetos con una contraseña
            
        Parameters
        ---------
        contraseña : str
            Contraseña que el usuario debe poner para modificar las opciones protegidas
        formato : bool
            Decide si se protege el formato de las celdas
        eliminar_filas : bool
            Decide si se protege la eliminación de las filas del Excel
        eliminar_columnas : bool
            Decide si se protege la eliminación de las columnas del Excel
        
        Returns
        -----
        None
        '''
        self.__hoja.protect(contraseña, {'format_cells': formato,
                                         'delete_columns': eliminar_columnas,
                                         'delete_rows': eliminar_filas})