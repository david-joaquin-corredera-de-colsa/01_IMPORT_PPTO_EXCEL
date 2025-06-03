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

' Constantes para elegir cuantas hojas historicas dejar visibles
Public Const CONST_HOJAS_DE_ENVIO_VISIBLES As Integer = 4

' Constantes para mostrar o no mensajes durante la ejecución
Public Const CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS As Boolean = False

' Constantes para valores por defecto
Private Const DEFAULT_COLUMN As String = "B"
Private Const DEFAULT_ROW As Long = 2
Private Const DEFAULT_DELIMITER As String = ";"

' Constantes para control de errores
Public Const ERROR_BASE_IMPORT As Long = vbObjectError + 1000

' Constantes para nombres de hojas
Public Const CONST_HOJA_EJECUTAR_PROCESOS As String = "00_Ejecutar_Procesos"
Public Const CONST_HOJA_INVENTARIO As String = "01_Inventario"
Public Const CONST_HOJA_LOG As String = "02_Log"
Public Const CONST_HOJA_USERNAME As String = "05_Username"
Public Const CONST_HOJA_DELIMITADORES_ORIGINALES As String = "06_Delimitadores_Originales"


' Constantes para etiquetas de procesamiento de líneas (NUEVAS)
Public Const CONST_TAG_LINEA_TRATADA As String = "Linea_Tratada"
Public Const CONST_TAG_LINEA_SUMA As String = "Linea_Suma"
Public Const CONST_TAG_LINEA_REPETIDA As String = "Linea_Repetida"

' Variable para hoja de envío anterior
Public gstrPreviaHojaImportacion_Envio As String
' Variable para nombre de copia de hoja de envío anterior
Public gstrPrevDelHojaImportacion_Envio As String

' Variables para hojas base del sistema
Public gstrHoja_EjecutarProcesos As String
Public gstrHoja_Inventario As String
Public gstrHoja_Log As String
Public gstrHoja_DelimitadoresOriginales As String
Public gstrHoja_UserName As String

' Variables para configuración de importación
Public gstrColumnaInicial_Importacion As String
Public glngFilaInicial_Importacion As Long
Public gstrDelimitador_Importacion As String
Public glngLineaInicial_HojaImportacion As Long
Public glngLineaFinal_HojaImportacion As Long

' Constantes para nombres de hojas


' Variables para nombres de hojas
Public gstrNuevaHojaImportacion As String
Public gstrNuevaHojaImportacion_Working As String
Public gstrNuevaHojaImportacion_Envio As String
Public gstrNuevaHojaImportacion_Comprobacion As String

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

Public Const CONST_OCULTAR_REPOSITORIO_DELIMITADORES As Boolean = True 'Poner como True si se desea ocultar la hoja
Public Const CONST_OCULTAR_HOJA_USERNAME As Boolean = True 'Poner como True si se desea ocultar la hoja


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



' =============================================================================
' VARIABLES GLOBALES PARA DETECCIÓN DE RANGOS POR PALABRAS CLAVE
' =============================================================================
' Fecha y hora de creación: 2025-06-03 03:19:45 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripción: Variables para detectar rangos basados en palabras clave específicas
' =============================================================================

Public vPalabraClave_PrimeraFila As String
Public vPalabraClave_PrimeraColumna As String
Public vPalabraClave_UltimaFila As String
Public vPalabraClave_UltimaColumna As String


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
    
    ' Inicializar hoja de envio anterior
    gstrPreviaHojaImportacion_Envio = ""
    ' Inicializar nombre de copia de hoja anterior
    gstrPrevDelHojaImportacion_Envio = ""
    
    ' Nombres de hojas base
    gstrHoja_EjecutarProcesos = CONST_HOJA_EJECUTAR_PROCESOS
    gstrHoja_Inventario = CONST_HOJA_INVENTARIO
    gstrHoja_Log = CONST_HOJA_LOG
    gstrHoja_DelimitadoresOriginales = CONST_HOJA_DELIMITADORES_ORIGINALES
    gstrHoja_UserName = CONST_HOJA_USERNAME
    
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
    gstrNuevaHojaImportacion_Comprobacion = ""
    
    'Adicional
    ' Configuración de importación
    vColumnaInicial_Importacion = "B"        ' Columna B (2)
    vFilaInicial_Importacion = 2             ' Fila 2
    vDelimitador_Importacion = ";"           ' Delimitador por defecto
    
    ' Inicializar variables de rango
    vLineaInicial_HojaImportacion = 0
    vLineaFinal_HojaImportacion = 0
    
    ' Inicializar palabras clave para detección de rangos
    vPalabraClave_PrimeraFila = "BUDGET_OS"      ' Palabra clave para primera fila
    vPalabraClave_PrimeraColumna = "BUDGET_OS"   ' Palabra clave para primera columna
    vPalabraClave_UltimaFila = "BUDGET_OS"       ' Palabra clave para última fila
    vPalabraClave_UltimaColumna = "M12"          ' Palabra clave para última columna
    
    ' Configurar nombre de backup automático
    If gstrPreviaHojaImportacion_Envio <> "" Then
        gstrPrevDelHojaImportacion_Envio = "Backup_" & gstrPreviaHojaImportacion_Envio
    End If
End Sub
