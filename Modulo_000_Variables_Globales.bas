Attribute VB_Name = "Modulo_000_Variables_Globales"
Option Explicit

'******************************************************************************
' Módulo: Global_Variables
' Fecha y Hora de Creación: 2025-05-26 10:04:46 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripción:
' Este módulo contiene todas las variables globales utilizadas en el sistema
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

' Variables para configuración de importación
Public gstrColumnaInicial_Importacion As String
Public glngFilaInicial_Importacion As Long
Public gstrDelimitador_Importacion As String
Public glngLineaInicial_HojaImportacion As Long
Public glngLineaFinal_HojaImportacion As Long

' Variables para nombres de hojas
Public gstrNuevaHojaImportacion As String
Public gstrNuevaHojaImportacion_Working As String
Public gstrNuevaHojaImportacion_Envio As String

' Variables para configuración de importación (adicional)
Public vColumnaInicial_Importacion As String
Public vFilaInicial_Importacion As Long
Public vDelimitador_Importacion As String

' Variables para detección de rango
Public vLineaInicial_HojaImportacion As Long
Public vLineaFinal_HojaImportacion As Long

' =============================================================================
' VARIABLES GLOBALES PARA DELIMITADORES DE EXCEL
' =============================================================================
' Fecha y hora de creación: 2025-05-26 17:43:59 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripción: Variables globales para el manejo de delimitadores de Excel
' =============================================================================

Public vHojaDelimitadoresExcelOriginales As String
Public vCelda_Header_Excel_UseSystemSeparators As String
Public vCelda_Header_Excel_DecimalSeparator As String
Public vCelda_Header_Excel_ThousandsSeparator As String
Public vCelda_Valor_Excel_UseSystemSeparators As String
Public vCelda_Valor_Excel_DecimalSeparator As String
Public vCelda_Valor_Excel_ThousandsSeparator As String
Public vExcel_UseSystemSeparators As String
Public vExcel_DecimalSeparator As String
Public vExcel_ThousandsSeparator As String


' =============================================================================
' VARIABLES GLOBALES ADICIONALES PARA RESTAURACIÓN DE DELIMITADORES
' =============================================================================
' Fecha y hora de creación: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripción: Variables globales adicionales para restaurar delimitadores originales
' =============================================================================

'Public vOcultarRepostiorioDelimitadores As Boolean
'vOcultarRepostiorioDelimitadores = True ' Cambiar a True si se desea ocultar la hoja
Public Const vOcultarRepostiorioDelimitadores As Boolean = True


' Variables para celdas que contienen valores originales
Public vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal As String
Public vCelda_Valor_Excel_DecimalSeparator_ValorOriginal As String
Public vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal As String

' Variables para almacenar valores originales leídos
Public vExcel_UseSystemSeparators_ValorOriginal As String
Public vExcel_DecimalSeparator_ValorOriginal As String
Public vExcel_ThousandsSeparator_ValorOriginal As String

' AUTOR: Sistema Automatizado
' VERSIÓN: 1.0
' COMPATIBILIDAD: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

' Variables globales para delimitadores
Public vDelimitadorDecimal_HFM As String
Public vDelimitadorMiles_HFM As String

Public Sub fun801_InicializarVariablesGlobales()

    ' =============================================================================
    ' FUNCIÓN: fun801_InicializarVariablesGlobales
    ' PROPÓSITO: Inicializa las variables globales con valores por defecto
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error GoTo ErrorHandler_fun801
    
    ' Inicializar delimitador decimal si está vacío
    If vDelimitadorDecimal_HFM = "" Or vDelimitadorDecimal_HFM = vbNullString Then
        vDelimitadorDecimal_HFM = "."
    End If
    
    ' Inicializar delimitador de miles si está vacío
    If vDelimitadorMiles_HFM = "" Or vDelimitadorMiles_HFM = vbNullString Then
        vDelimitadorMiles_HFM = ","
    End If
    
    Exit Sub

ErrorHandler_fun801:
    ' En caso de error, usar valores por defecto
    vDelimitadorDecimal_HFM = "."
    vDelimitadorMiles_HFM = ","
End Sub

'------------------------------------------------------------------------------
' Procedimiento: InitializeGlobalVariables
' Descripción: Inicializa todas las variables globales con valores por defecto
'------------------------------------------------------------------------------
Public Sub InitializeGlobalVariables()
    ' Nombres de hojas base
    gstrHoja_EjecutarProcesos = "00_Ejecutar_Procesos"
    gstrHoja_Inventario = "01_Inventario"
    gstrHoja_Log = "02_Log"
    
    ' Configuración de importación
    gstrColumnaInicial_Importacion = DEFAULT_COLUMN
    glngFilaInicial_Importacion = DEFAULT_ROW
    gstrDelimitador_Importacion = DEFAULT_DELIMITER
    
    ' Inicializar variables de líneas
    glngLineaInicial_HojaImportacion = 0
    glngLineaFinal_HojaImportacion = 0
    
    ' Inicializar nombres de hojas
    gstrNuevaHojaImportacion = ""
    gstrNuevaHojaImportacion_Working = ""
    gstrNuevaHojaImportacion_Envio = ""
    
    'Adicional
    ' Configuración de importación
    vColumnaInicial_Importacion = "B"        ' Columna B (2)
    vFilaInicial_Importacion = 2             ' Fila 2
    vDelimitador_Importacion = ";"           ' Delimitador por defecto
    
    ' Inicializar variables de rango
    vLineaInicial_HojaImportacion = 0
    vLineaFinal_HojaImportacion = 0
End Sub
