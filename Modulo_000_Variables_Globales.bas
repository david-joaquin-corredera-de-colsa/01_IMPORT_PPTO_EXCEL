Attribute VB_Name = "Modulo_000_Variables_Globales"
Option Explicit

'******************************************************************************
' M�dulo: Global_Variables
' Fecha y Hora de Creaci�n: 2025-05-26 10:04:46 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripci�n:
' Este m�dulo contiene todas las variables globales utilizadas en el sistema
'******************************************************************************

' Constantes para valores por defecto
Private Const DEFAULT_COLUMN As String = "B"
Private Const DEFAULT_ROW As Long = 2
Private Const DEFAULT_DELIMITER As String = ";"

' Constantes para control de errores
Public Const ERROR_BASE_IMPORT As Long = vbObjectError + 1000

' Variables para hojas base del sistema
Public gstrHoja_EjecutarProcesos As String
Public gstrHoja_Inventario As String
Public gstrHoja_Log As String

' Variables para configuraci�n de importaci�n
Public gstrColumnaInicial_Importacion As String
Public glngFilaInicial_Importacion As Long
Public gstrDelimitador_Importacion As String
Public glngLineaInicial_HojaImportacion As Long
Public glngLineaFinal_HojaImportacion As Long

' Variables para nombres de hojas
Public gstrNuevaHojaImportacion As String
Public gstrNuevaHojaImportacion_Working As String
Public gstrNuevaHojaImportacion_Envio As String

' Variables para configuraci�n de importaci�n (adicional)
Public vColumnaInicial_Importacion As String
Public vFilaInicial_Importacion As Long
Public vDelimitador_Importacion As String

' Variables para detecci�n de rango
Public vLineaInicial_HojaImportacion As Long
Public vLineaFinal_HojaImportacion As Long

'------------------------------------------------------------------------------
' Variables para delimitadores del sistema
'------------------------------------------------------------------------------
Public vDelimitadorDecimales As String
Public vDelimitadorMiles As String


'------------------------------------------------------------------------------
' Procedimiento: InitializeGlobalVariables
' Descripci�n: Inicializa todas las variables globales con valores por defecto
'------------------------------------------------------------------------------
Public Sub InitializeGlobalVariables()
    ' Nombres de hojas base
    gstrHoja_EjecutarProcesos = "00_Ejecutar_Procesos"
    gstrHoja_Inventario = "01_Inventario"
    gstrHoja_Log = "02_Log"
    
    ' Configuraci�n de importaci�n
    gstrColumnaInicial_Importacion = DEFAULT_COLUMN
    glngFilaInicial_Importacion = DEFAULT_ROW
    gstrDelimitador_Importacion = DEFAULT_DELIMITER
    
    ' Inicializar variables de l�neas
    glngLineaInicial_HojaImportacion = 0
    glngLineaFinal_HojaImportacion = 0
    
    ' Inicializar nombres de hojas
    gstrNuevaHojaImportacion = ""
    gstrNuevaHojaImportacion_Working = ""
    gstrNuevaHojaImportacion_Envio = ""
    
    'Adicional
    ' Configuraci�n de importaci�n
    vColumnaInicial_Importacion = "B"        ' Columna B (2)
    vFilaInicial_Importacion = 2             ' Fila 2
    vDelimitador_Importacion = ";"           ' Delimitador por defecto
    
    ' Inicializar variables de rango
    vLineaInicial_HojaImportacion = 0
    vLineaFinal_HojaImportacion = 0
End Sub
