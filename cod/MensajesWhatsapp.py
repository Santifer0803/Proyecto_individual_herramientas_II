# -*- coding: utf-8 -*-
"""
Created on Sun Jun  9 17:07:03 2024

@author: Santiago
"""

# Módulo para mandar mensajes
import pywhatkit as kit

# Clase para mandar mensajes
class MensajesWhatsapp:
    """
    Clase que contiene métodos para mandar mensajes de Whatsapp

    Attributes
    ----------
    numero : str
        El número de celular del receptor, iniciando con el símbolo '+', su código de país y el número correspondiente, sin dejar espacios

    Methods
    -------
    programar_mensaje(self, mensaje, hora, minuto):
        Método que manda un mensaje de WhatsApp a la hora programada, es necesario tener WhatsApp web cargado, de lo contrario, se pueden presentar fallos
    def mensaje_instantaneo(self, mensaje):
        Método que manda un mensaje de WhatsApp instantáneamente, es necesario tener WhatsApp web cargado, de lo contrario, se pueden presentar fallos
     
    """
    
    # Constructor
    def __init__(self, numero):
        self.__numero = numero
    
    # Get y set
    @property
    def numero(self):
        return self.__numero
    
    @numero.setter
    def numero(self, nuevo_numero):
        self.__numero = nuevo_numero
        
    # str
    def __str__(self):
        ''' Texto explicando un resumen de la clase MensajesWhatsapp
            
        Parameters
        ---------
        None
        
        Returns
        -----
        Cadena : str
            Texto explicativo que resume la clase MensajesWhatsapp
        '''
        return f'''Un número de celular que corresponde al receptor de un mensaje de whatsapp, guardado para usarse en el módulo pywhatkey el cual es: \n
               Número: {self.__numero}'''
            
    def programar_mensaje(self, mensaje, hora, minuto):
        ''' Método que manda un mensaje de WhatsApp a la hora programada, es necesario tener WhatsApp web cargado, de lo contrario, se pueden presentar fallos
            
        Parameters
        ---------
        mensaje : str
            Mensaje que le será enviado al receptor
        hora : int
            Entero con valor entre 0 y 23, representa la hora a la cual se enviará el mensaje, está en formato de 24 horas
        minuto : int
            Entero con valor entre 0 y 59, representa el minuto en el que se enviará el mensaje
        
        Returns
        -----
        None
        '''
        kit.sendwhatmsg(self.__numero, mensaje, hora, minuto)
    
    def mensaje_instantaneo(self, mensaje):
        ''' Método que manda un mensaje de WhatsApp instantáneamente, es necesario tener WhatsApp web cargado, de lo contrario, se pueden presentar fallos
            
        Parameters
        ---------
        mensaje : str
            Mensaje que le será enviado al receptor
        
        Returns
        -----
        None
        '''
        kit.sendwhatmsg_instantly(self.__numero, mensaje)